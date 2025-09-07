"use strict";
const OPTIONS = ['RTSelect', 'RTShow', 'RTEdit'];
const RTDropDownTag = 'RTList';
const RTDuplicateTag = 'RTRepeat';
const RTSectionTag = 'RTSection';
const RTSelectTag = 'RTSelect';
const RTObsTag = 'RTObs';
const RTDescriptionTag = 'RTDesc';
const RTDescriptionStyle = 'RTDescription';
const RTSiTag = 'RTSi';
const RTSiStyles = ['RTSi0cm', 'RTSi1cm', 'RTSi2cm', 'RTSi3cm', 'RTSi4cm'];
let USERFORM, NOTIFICATION;
let RichText, RichTextInline, RichTextParag, ComboBox, CheckBox, dropDownList, Bounding, Hidden;
Office.onReady((info) => {
    // Check that we loaded into Word
    if (info.host !== Office.HostType.Word)
        return showNotification('This addin is designed to work on Word only');
    USERFORM = document.getElementById('userFormSection');
    NOTIFICATION = document.getElementById('notification');
    RichText = Word.ContentControlType.richText;
    RichTextInline = Word.ContentControlType.richTextInline;
    RichTextParag = Word.ContentControlType.richTextParagraphs;
    ComboBox = Word.ContentControlType.comboBox;
    CheckBox = Word.ContentControlType.checkBox;
    dropDownList = Word.ContentControlType.dropDownList;
    Bounding = Word.ContentControlAppearance.boundingBox;
    Hidden = Word.ContentControlAppearance.hidden;
    mainUI();
});
function showBtns(btns, append = true) {
    return btns.map(btn => insertBtn(btn, append));
}
;
function mainUI() {
    if (!USERFORM)
        return;
    USERFORM.innerHTML = '';
    const main = [[customizeContract, 'Customize Contract'], [prepareTemplate, 'Prepare Template']];
    const btns = showBtns(main);
    const back = [goBack, 'Go Back'];
    btns.forEach(btn => btn === null || btn === void 0 ? void 0 : btn.addEventListener('click', () => insertBtn(back, false)));
    function goBack() {
        USERFORM.innerHTML = '';
        showBtns(main);
    }
}
function prepareTemplate() {
    USERFORM.innerHTML = '';
    function wrap(title, tag, label, type, style, cannotEdit, cannotDelete) {
        return [
            () => wrapSelectionWithContentControl(title, tag, type, style, cannotEdit, cannotDelete),
            label
        ];
    }
    ;
    const insertDescription = () => findTextAndWrapItWithContentControl([RTDescriptionStyle], RTDescriptionTag, RTDescriptionTag, true, true);
    const btns = [
        wrap(RTSiTag, RTSiTag, 'Insert Single RT Si', RichText, RTSiStyles[0], true, true),
        wrap(RTDescriptionTag, RTDescriptionTag, 'Insert Single RT Description', RichText, RTDescriptionStyle, true, true),
        wrap(RTSelectTag, RTSelectTag, 'Insert Single RT Select', RichText, null, false, true),
        wrap(RTSectionTag, RTSectionTag, 'Insert Single RT Section', RichText, RTSectionTag, true, true),
        wrap(RTDuplicateTag, RTDuplicateTag, 'Insert Single RT Dublicate Block', RichText, null, false, true),
        [insertDropDownList, 'Insert a Dropdown List from selection'],
        wrap(RTObsTag, RTObsTag, 'Insert Single RT Obs', RichText, RTObsTag, true, true),
        [insertRTSiAll, 'Insert RT Si For All'],
        [insertDescription, 'Insert RT Description For All'],
    ];
    showBtns(btns);
}
function insertBtn([fun, label], append = true) {
    if (!USERFORM)
        return;
    const htmlBtn = document.createElement('button');
    append ? USERFORM.appendChild(htmlBtn) : USERFORM.prepend(htmlBtn);
    htmlBtn.innerText = label;
    htmlBtn.addEventListener('click', () => fun());
    return htmlBtn;
}
/**
 * Wraps every occurrence of text formatted with a specific character style in a rich text content control.
 *
 * @param style The name of the character style to find (e.g., "Emphasis", "Strong", "MyCustomStyle").
 * @returns A Promise that resolves when the operation is complete.
 */
async function findTextAndWrapItWithContentControl(styles, title, tag, cannotEdit, cannotDelete) {
    var _a, _b;
    const separator = '_&_';
    const search = (_a = (await promptForInput(`Provide the search string. You can provide more than one string to search by separated by ${separator} witohout space`, separator))) === null || _a === void 0 ? void 0 : _a.split(separator);
    if (!(search === null || search === void 0 ? void 0 : search.length))
        return showNotification('The provided search string is not valid');
    const matchWildcards = await promptConfirm('Match Wild Cards');
    if (!styles)
        styles = ((_b = (await promptForInput(`Provide the styles that that need to be matched separated by ","`))) === null || _b === void 0 ? void 0 : _b.split(',')) || [];
    if (!(styles === null || styles === void 0 ? void 0 : styles.length))
        return showNotification(`The styles[] has 0 length, no styles are included, the function will return`);
    await Word.run(async (context) => {
        for (const el of search) {
            const ranges = await searchString(el, context, matchWildcards);
            if (!ranges)
                continue;
            await wrapMatchingStyleRangesWithContentControls(ranges, styles, title, tag, cannotEdit, cannotDelete);
        }
        ;
    });
}
async function wrapMatchingStyleRangesWithContentControls(ranges, styles, title, tag, cannotEdit, cannotDelete) {
    ranges.load(['style', 'text', 'parentContentControlOrNullObject', 'parentContentControlOrNullObject.isNullObject', 'parentContentControlOrNullObject.tag']);
    await ranges.context.sync();
    if (!ranges.items.length) {
        showNotification(`No text matching the search string was found in the document.`);
        return;
    }
    showNotification(`Found ${ranges.items.length} ranges matching the search string. First range text = ${ranges.items[0].text}`);
    return ranges.items.map(async (range, index) => {
        var _a;
        if (!styles.includes(range.style))
            return;
        if (((_a = range.parentContentControlOrNullObject) === null || _a === void 0 ? void 0 : _a.tag) === tag)
            return;
        return await insertContentControl(range, title, tag, index, RichText, range.style, cannotEdit, cannotDelete);
    });
}
async function searchString(search, context, matchWildcards) {
    const searchResults = context.document.body.search(search, { matchWildcards: matchWildcards });
    searchResults.load(['style']);
    await context.sync();
    return searchResults;
}
async function addIDtoCtrlTitle(ctrls) {
    ctrls.load(['title', 'id']);
    await ctrls.context.sync();
    ctrls.items
        .filter(ctrl => !ctrl.title.endsWith(`-${ctrl.id}`))
        .forEach(ctrl => ctrl.title = `${ctrl.title}-${ctrl.id}`);
    await ctrls.context.sync();
}
async function insertRTSiAll() {
    await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load(['style', 'text', 'range', 'parentContentControlOrNullObject']);
        await context.sync();
        const parags = paragraphs.items
            .filter(p => RTSiStyles.includes(p.style));
        console.log(parags);
        for (const parag of parags) {
            parag.select();
            try {
                const parent = parag.parentContentControlOrNullObject;
                parent.load(['tag']);
                await parag.context.sync();
                if (parent.tag === RTSiTag)
                    continue;
                showNotification(`range style: ${parag.style} & text = ${parag.text}`);
                await insertContentControl(parag.getRange('Content'), RTSiTag, RTSiTag, parags.indexOf(parag), RichText, parag.style);
            }
            catch (error) {
                showNotification(`error: ${error}`);
                continue;
            }
        }
        await context.sync();
    });
}
async function insertDropDownList() {
    const range = await getSelectionRange();
    if (!range)
        return;
    range.load(["text"]);
    range.context.trackedObjects.add(range); //!This is important
    await range.context.sync();
    const options = range.text.split("/");
    if (!options.length)
        return showNotification("No options");
    showNotification(options.join());
    const ctrl = await insertContentControl(range, RTDropDownTag, RTDropDownTag, 0, dropDownList, null);
    if (!ctrl)
        return;
    ctrl.cannotEdit = false; //! If we do not set it to false, it will not be possible to select from the list
    ctrl.dropDownListContentControl.deleteAllListItems();
    options.forEach(option => ctrl.dropDownListContentControl.addListItem(option));
    await ctrl.context.sync();
}
async function insertContentControl(range, title, tag, index, type, style, cannotEdit = true, cannotDelete = true) {
    range.select();
    const styles = range.context.document.getStyles();
    styles.load(['nameLocal', 'type']);
    // Insert a rich text content control around the found range.
    //@ts-expect-error
    const ctrl = range.insertContentControl(type);
    ctrl.load(["id"]);
    range.context.trackedObjects.remove(range);
    range.context.trackedObjects.add(ctrl); //!This is very important otherwise we will not be able to call range.context.sync() after calling range.context.sync();
    await range.context.sync();
    // Set properties for the new content control.
    if (ctrl.id)
        showNotification(`the newly created ContentControl id = ${ctrl.id} `);
    try {
        ctrl.select();
        ctrl.title = `${title}&${ctrl.id}`;
        ctrl.tag = tag;
        ctrl.cannotDelete = cannotDelete;
        ctrl.cannotEdit = cannotEdit;
        ctrl.appearance = Word.ContentControlAppearance.boundingBox;
        const foundStyle = styles.items.find(s => s.nameLocal === style);
        if (style && (foundStyle === null || foundStyle === void 0 ? void 0 : foundStyle.type) === Word.StyleType.character)
            ctrl.style = style;
        await range.context.sync();
        showNotification(`Wrapped text in range ${index || 1} with a content control.`);
    }
    catch (error) {
        showNotification(`There was an error while setting the properties of the newly crated contentcontrol by insertContentControl(): ${error}.`);
    }
    return ctrl;
}
async function getSelectionRange() {
    return await Word.run(async (context) => {
        const range = context.document
            .getSelection()
            .getRange('Content');
        range.load(['style', 'isEmpty']);
        context.trackedObjects.add(range);
        await context.sync();
        if (range.isEmpty)
            return showNotification('The selection range is empty');
        return range;
    });
}
async function wrapSelectionWithContentControl(title, tag, type, style, cannotEdit, cannotDelete) {
    const range = await getSelectionRange();
    if (!range)
        return;
    if (RTSiStyles.includes(range.style))
        style = range.style;
    await insertContentControl(range, title, tag, 0, type, style, cannotEdit, cannotDelete);
}
;
async function promptConfirm(question, fun) {
    if (!question)
        question = 'No question was provided !!!';
    const container = createHTMLElement('div', 'promptContainer', '', USERFORM);
    const prompt = createHTMLElement('div', 'prompt', '', container);
    createHTMLElement('p', 'ask', question, prompt);
    const btns = createHTMLElement('div', 'btns', '', prompt);
    const btnOK = createHTMLElement('button', 'btnOK', 'OK', btns);
    const btnNo = createHTMLElement('button', 'btnCancel', 'NO', btns);
    return new Promise((resolve, reject) => {
        btnOK.onclick = () => resolve(confirm(true));
        btnNo.onclick = () => resolve(confirm(false));
    });
    function confirm(confirm) {
        container.remove();
        if (fun)
            fun(confirm);
        return confirm;
    }
}
;
async function customizeContract() {
    USERFORM.innerHTML = '';
    await selectCtrls();
    async function selectCtrls() {
        await Word.run(async (context) => {
            const allRT = context.document.contentControls;
            allRT.load(['title', 'tag', 'contentControls/items/title', 'contentControls/items/tag']);
            await context.sync();
            const tags = [...OPTIONS, RTDuplicateTag];
            const selectCtrls = allRT.items
                .filter(ctrl => tags.includes(ctrl.tag));
            const selected = [];
            for (const ctrl of selectCtrls)
                await promptForSelection(ctrl, selected);
            const keep = selected.filter(title => !title.startsWith('!'));
            showNotification(`keep = ${keep.join(', ')}`);
            try {
                await currentDoc();
                //await createNewDoc();
            }
            catch (error) {
                showNotification(`${error}`);
            }
            async function currentDoc() {
                for (const ctrl of selectCtrls) {
                    if (keep.includes(ctrl.title))
                        continue;
                    ctrl.select();
                    ctrl.cannotDelete = false;
                    showNotification(`Deleted Ctrl: ${ctrl.title}`);
                    ctrl.delete(false);
                }
                await context.sync();
            }
            ;
            async function createNewDoc() {
                return; //!Desactivating working with new document created from template until we find a solution to the context issue
                const template = await getTemplate();
                console.log(template);
                if (!template)
                    return showNotification('Failed to create the template');
                const newDoc = context.application.createDocument(template);
                const all = newDoc.contentControls;
                all.load(['title', 'tag']);
                await newDoc.context.sync();
                showNotification(`All ctrls from newDoc = : ${all.items.map(c => c.title).join(', ')}`);
                all.items.map(ctrl => {
                    if (keep.includes(ctrl.title))
                        return;
                    ctrl.cannotDelete = false;
                    ctrl.delete(false);
                });
                await newDoc.context.sync();
                newDoc.open();
            }
        });
    }
    async function promptForSelection(ctrl, selected) {
        if (ctrl.tag === RTDuplicateTag)
            return await duplicateBlock(ctrl);
        if (selected.find(t => t.includes(ctrl.title)))
            return; //!We need to exclude any ctrl that has already been passed to the function or has been excluded: when a ctrl is excluded, its children are added to the array as excluded ctrls ("![ctrl.title]"), they do not hence need to be treated again since we already know theyare to be  excluded. This also avoids the problem that happens sometimes, when a ctrl has its parent amongst its children list (this is an apparently known weird behavior if the ctrl range overlaps somehow with the range of another ctrl)
        ctrl.select();
        const [container, btnNext, checkBox] = await showUI();
        return new Promise((resolve, reject) => {
            btnNext.onclick = () => nextCtrl(ctrl, checkBox);
            async function nextCtrl(ctrl, checkBox) {
                const checked = checkBox.checked;
                container.remove();
                //ctrl.contentControls.load(['title', 'tag']);
                //await ctrl.context.sync();
                const subOptions = ctrl.contentControls.items
                    .filter(ctrl => OPTIONS.includes(ctrl.tag));
                if (checked)
                    await isSelected(ctrl, subOptions);
                else
                    isNotSelected(ctrl, subOptions);
                resolve(selected);
            }
            ;
        });
        async function isSelected(ctrl, subOptions) {
            selected.push(ctrl.title);
            for (const ctrl of subOptions)
                await promptForSelection(ctrl, selected);
            console.log(selected);
        }
        ;
        function isNotSelected(ctrl, subOptions) {
            const exclude = (title) => `!${title}`;
            selected.push(exclude(ctrl.title));
            subOptions
                .forEach(ctrl => selected.push(exclude(ctrl.title)));
            console.log(selected);
        }
        ;
        async function showUI() {
            const children = ctrl.contentControls;
            children.load(['title', 'tag']);
            await ctrl.context.sync();
            const RTSi = children.items.find(rt => rt.tag === RTSiTag);
            if (!RTSi)
                throw new Error('No RTSi');
            const ctrlRange = RTSi.getRange('Content');
            ctrlRange.load(['text', 'paragraphs']);
            await ctrl.context.sync();
            return UI(ctrlRange.text);
            function UI(text) {
                const container = createHTMLElement('div', 'promptContainer', '', USERFORM, ctrl.title);
                const prompt = createHTMLElement('div', 'selection', '', container);
                const checkBox = createHTMLElement('input', 'checkBox', '', prompt);
                createHTMLElement('label', 'label', text, prompt);
                checkBox.type = 'checkbox';
                const btns = createHTMLElement('div', 'btns', '', prompt);
                const btnNext = createHTMLElement('button', 'btnOK', 'Next', btns);
                return [container, btnNext, checkBox];
            }
        }
        async function duplicateBlock(ctrl) {
            try {
                await duplicate();
            }
            catch (error) {
                showNotification(`${error}`);
            }
            async function duplicate() {
                const label = ctrl.contentControls.items.find(c => c.tag === RTSectionTag);
                if (!label)
                    return showNotification(`No Section RT Within the Range of the Duplicate Ctrl. Ctrl id = ${ctrl.id}`);
                label === null || label === void 0 ? void 0 : label.load(['text']);
                const ctrlContent = ctrl.getOoxml();
                const range = ctrl.getRange();
                await ctrl.context.sync();
                if (!label.text)
                    return showNotification("No lable text");
                ctrl.select();
                const message = `How many ${label.text} parties are there?`;
                const answer = Number(await promptForInput(message));
                if (isNaN(answer))
                    return showNotification(`The provided text cannot be converted into a number: ${answer}`);
                for (let i = 1; i < answer; i++) {
                    range
                        .insertOoxml(ctrlContent.value, Word.InsertLocation.after);
                }
                await ctrl.context.sync();
            }
        }
    }
    function getFileURL() {
        let url;
        Office.context.document.getFilePropertiesAsync(undefined, (result) => {
            if (result.error)
                return;
            url = result.value.url;
        });
        return url;
    }
    async function getTemplate() {
        try {
            return await getDocumentBase64();
        }
        catch (error) {
            showNotification(`Failed to create new Doc: ${error}`);
        }
    }
}
;
async function promptForInput(question, deflt, fun) {
    if (!question)
        return '';
    const container = createHTMLElement('div', 'promptContainer', '', USERFORM);
    const prompt = createHTMLElement('div', 'prompt', '', container);
    const ask = createHTMLElement('p', 'ask', question, prompt);
    const input = createHTMLElement('input', 'answer', '', prompt);
    const btns = createHTMLElement('div', 'btns', '', prompt);
    const btnOK = createHTMLElement('button', 'btnOK', 'OK', btns);
    const btnCancel = createHTMLElement('button', 'btnCancel', 'Cancel', btns);
    if (deflt)
        input.value = deflt;
    return new Promise((resolve, reject) => {
        btnCancel.onclick = () => reject(container.remove());
        btnOK.onclick = () => {
            const answer = input.value;
            console.log('user answer = ', answer);
            container.remove();
            if (fun)
                fun(answer);
            resolve(answer);
        };
    });
}
;
/**
 * Asynchronously gets the entire document content as a Base64 string.
 * This function handles multi-slice documents by requesting each slice in parallel.
 * @returns A Promise that resolves with the Base64-encoded document content.
 */
async function getDocumentBase64() {
    const failed = (result) => result.status !== Office.AsyncResultStatus.Succeeded;
    const sliceSize = 16 * 1024; //!We need not to exceed the Maximum call stack limit when the slices will be passed to String.FromCharCode()
    return new Promise((resolve, reject) => {
        // Step 1: Request the document as a compressed file.
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: sliceSize }, (fileResult) => processFile(fileResult));
        function processFile(fileResult) {
            if (failed(fileResult))
                return reject(fileResult.error);
            const file = fileResult.value;
            const sliceCount = file.sliceCount;
            const slices = [];
            getSlice();
            function getSlice() {
                file.getSliceAsync(slices.length, (sliceResult) => processSlice(sliceResult));
            }
            function processSlice(sliceResult) {
                try {
                    if (failed(sliceResult))
                        return file.closeAsync(() => reject(sliceResult.error));
                    slices.push(sliceResult.value.data);
                    if (slices.length < sliceCount)
                        return getSlice();
                    const binaryString = slices.map(slice => String.fromCharCode(...slice)).join('');
                    const base64String = btoa(binaryString);
                    file.closeAsync(() => resolve(base64String));
                }
                catch (error) {
                    showNotification(`${error}, succeeded = ${sliceResult.status}, loaded = ${slices.length}`);
                }
            }
        }
    });
}
async function deleteAllNotSelected(selected, wdDoc) {
    const all = wdDoc.contentControls;
    all.load(['items', 'title', 'tag']);
    await wdDoc.context.sync();
    all.items
        .filter(ctrl => !selected.includes(ctrl.title))
        .forEach(ctrl => {
        ctrl.select();
        ctrl.cannotDelete = false;
        ctrl.delete(true);
    });
    await wdDoc.context.sync();
}
function createHTMLElement(tag, css, innerText, parent, id, append = true) {
    const el = document.createElement(tag);
    if (innerText)
        el.innerText = innerText;
    el.classList.add(css);
    if (id)
        el.id = id;
    append ? parent.appendChild(el) : parent.prepend(el);
    return el;
}
function showNotification(message, clear = false) {
    if (clear)
        NOTIFICATION.innerHTML = '';
    createHTMLElement('p', 'notification', message, NOTIFICATION, '', true);
}
async function setCanBeEditedForAllSelectCtrls(edit = true) {
    await Word.run(async (context) => {
        const ctrls = context.document
            .contentControls;
        ctrls.load(['title', 'tag']);
        await context.sync();
        ctrls.items.forEach(ctrl => {
            if (OPTIONS.indexOf(ctrl.tag) > -1)
                ctrl.cannotEdit = edit;
        });
        await context.sync();
    });
}
//# sourceMappingURL=app.js.map