"use strict";
const OPTIONS = ['RTSelect', 'RTShow', 'RTEdit'];
const RTDropDownTag = 'RTList';
const RTDropDownColor = '#991c63';
const RTDuplicateTag = 'RTRepeat';
const RTSectionTag = 'RTSection';
const RTSelectTag = 'RTSelect';
const RTOrTag = 'RTOr';
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
    function wrap(title, tag, type, style, cannotEdit, cannotDelete, label) {
        return [
            () => wrapSelectionWithContentControl(title, tag, type, style, cannotEdit, cannotDelete),
            label
        ];
    }
    ;
    const btns = [
        wrap(RTSiTag, RTSiTag, RichText, RTSiStyles[0], true, true, 'Insert Single RT Si'),
        [() => insertRTDescription(true), 'Insert Single RT Description'],
        wrap(RTSelectTag, RTSelectTag, RichText, null, false, true, 'Insert Single RT Select'),
        wrap(RTSectionTag, RTSectionTag, RichText, RTSectionTag, true, true, 'Insert Single RT Section'),
        wrap(RTOrTag, RTOrTag, RichText, null, false, true, 'Insert Single RT OR'),
        wrap(RTDuplicateTag, RTDuplicateTag, RichText, null, false, true, 'Insert Single RT Dublicate Block'),
        [insertDropDownList, 'Insert a Dropdown List from selection'],
        wrap(RTObsTag, RTObsTag, RichText, RTObsTag, true, true, 'Insert Single RT Obs'),
        [insertRTSiAll, 'Insert RT Si For All'],
        [insertRTDescription, 'Insert RT Description For All'],
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
    const all = [];
    return await Word.run(async (context) => {
        for (const el of search) {
            const ranges = await searchString(el, context, matchWildcards);
            if (!ranges)
                continue;
            const ctrls = await wrapMatchingStyleRangesWithContentControls(ranges, styles, title, tag, cannotEdit, cannotDelete);
            if (!ctrls)
                continue;
            all.push(ctrls);
        }
        ;
        return all.flat();
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
    const ctrls = ranges.items.map(async (range, index) => {
        var _a;
        if (!styles.includes(range.style))
            return;
        if (((_a = range.parentContentControlOrNullObject) === null || _a === void 0 ? void 0 : _a.tag) === tag)
            return;
        return await insertContentControl(range, title, tag, index, RichText, range.style, cannotEdit, cannotDelete);
    });
    return Promise.all(ctrls);
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
        .forEach(ctrl => ctrl.title = getCtrlTitle(ctrl.tag, ctrl.id));
    await ctrls.context.sync();
}
async function insertRTDescription(selection = false, style = 'Normal') {
    var _a;
    let ctrls;
    if (selection) {
        const range = await getSelectionRange();
        if (!range)
            return showNotification('No Text Was selected !');
        ctrls = [await insertContentControl(range, RTDescriptionTag, RTDescriptionTag, 0, RichText, RTDescriptionStyle, true, true)];
    }
    else
        ctrls = await findTextAndWrapItWithContentControl([RTDescriptionStyle], RTDescriptionTag, RTDescriptionTag, true, true);
    if (!(ctrls === null || ctrls === void 0 ? void 0 : ctrls.length))
        return;
    for (const ctrl of ctrls) {
        if (!ctrl)
            continue;
        const range = ctrl.getRange();
        const inserted = range.insertText('[*]\u00A0', Word.InsertLocation.before);
        inserted.style = style;
        inserted.font.bold = true;
    }
    await ((_a = ctrls[0]) === null || _a === void 0 ? void 0 : _a.context.sync());
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
            const style = RTSiStyles.includes(parag.style) ? parag.style : RTSiStyles[0];
            try {
                const parent = parag.parentContentControlOrNullObject;
                parent.load(['tag']);
                await parag.context.sync();
                if (parent.tag === RTSiTag)
                    continue;
                showNotification(`range style: ${parag.style} & text = ${parag.text}`);
                await insertContentControl(parag.getRange('Content'), RTSiTag, RTSiTag, parags.indexOf(parag), RichText, style);
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
    const ctrl = await insertContentControl(range, RTDropDownTag, RTDropDownTag, 0, dropDownList, null, false, true);
    if (!ctrl)
        return;
    ctrl.dropDownListContentControl.deleteAllListItems();
    options.forEach(option => ctrl.dropDownListContentControl.addListItem(option));
    setCtrlsFontColor([ctrl], RTDropDownColor);
    setCtrlsColor([ctrl], RTDropDownColor);
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
        ctrl.title = getCtrlTitle(title, ctrl.id);
        ctrl.tag = tag;
        ctrl.appearance = Word.ContentControlAppearance.boundingBox;
        const foundStyle = styles.items.find(s => s.nameLocal === style);
        if (style && (foundStyle === null || foundStyle === void 0 ? void 0 : foundStyle.type) === Word.StyleType.character)
            ctrl.style = style;
        if (style)
            ctrl.getRange().style = style;
        ctrl.cannotDelete = cannotDelete;
        ctrl.cannotEdit = cannotEdit; //!This must come at the end after the style has been set.
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
    const TAGS = [...OPTIONS, RTDuplicateTag];
    const getSelectCtrls = (ctrls) => ctrls.filter(ctrl => TAGS.includes(ctrl.tag));
    const selected = [];
    await selectCtrls();
    async function selectCtrls() {
        await Word.run(async (context) => {
            const allRT = context.document.getContentControls();
            allRT.load(['id', 'title', 'tag']);
            await context.sync();
            const selectCtrls = getSelectCtrls(allRT.items);
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
                const allRT = context.document.getContentControls();
                allRT.load(['id', 'title', 'tag']);
                await context.sync();
                const selectCtrls = getSelectCtrls(allRT.items); //!We need to retrieve all the selected items again because we may have added new ctrls by cloning the 'Duplicate' ctrls
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
        if (selected.find(t => t.includes(ctrl.title)))
            return; //!We need to exclude any ctrl that has already been passed to the function or has been excluded: when a ctrl is excluded, its children are added to the array as excluded ctrls ("![ctrl.title]"), they do not hence need to be treated again since we already know theyare to be  excluded. This also avoids the problem that happens sometimes, when a ctrl has its parent amongst its children list (this is an apparently known weird behavior if the ctrl range overlaps somehow with the range of another ctrl)
        ctrl.select();
        return showSelectPrompt([ctrl]);
    }
    async function showSelectPrompt(selectCtrls, labelTag = RTSiTag) {
        const blocks = [];
        for (const ctrl of selectCtrls) {
            if (ctrl.tag === RTDuplicateTag) {
                await duplicateBlock(ctrl);
                continue;
            }
            ;
            const addBtn = selectCtrls.indexOf(ctrl) + 1 === selectCtrls.length;
            blocks.push(await insertPromptBlock(ctrl, addBtn, labelTag) || undefined);
        }
        return btnPromise(blocks);
    }
    async function insertPromptBlock(ctrl, addBtn, labelTag) {
        try {
            return wordRun();
        }
        catch (error) {
            return showNotification(`${error}`);
        }
        async function wordRun() {
            return await Word.run(async (context) => {
                const ctrlSi = getFirstByTag(ctrl, labelTag);
                ctrlSi.select();
                ctrlSi.load(['id', 'title', 'tag']);
                ctrlSi.cannotEdit = false; //!We must unlock the text in order to be able to change the font.hidden property
                const rangeSi = ctrlSi.getRange();
                rangeSi.load(['text', 'font']);
                await context.sync();
                rangeSi.font.hidden = false; //!We must unhide the text, otherwise we will get an empty string
                await context.sync(); //!We mus sync after changing the font.hidden property
                const text = rangeSi.text;
                rangeSi.font.hidden = true;
                ctrlSi.cannotEdit = true;
                await context.sync();
                return { ctrl, ...appendHTMLElements(text, ctrl.title, addBtn) }; //The checkBox will have as id the title of the "select" contentcontrol}
            });
        }
    }
    function appendHTMLElements(text, id, addBtn = false) {
        const container = createHTMLElement('div', 'promptContainer', '', USERFORM);
        const checkBox = createHTMLElement('input', 'checkBox', '', container, id);
        createHTMLElement('label', 'label', text, container);
        checkBox.type = 'checkbox';
        if (!addBtn)
            return { container, checkBox };
        const btns = createHTMLElement('div', 'btns', '', container);
        const btnNext = createHTMLElement('button', 'btnOK', 'Next', btns);
        return { container, checkBox, btnNext };
    }
    function btnPromise(blocks) {
        return new Promise((resolve, reject) => {
            var _a;
            const btn = (_a = blocks.find(container => container === null || container === void 0 ? void 0 : container.btnNext)) === null || _a === void 0 ? void 0 : _a.btnNext;
            !btn ? resolve(selected) : btn.onclick = processBlocks;
            async function processBlocks() {
                const values = blocks
                    .filter(block => (block === null || block === void 0 ? void 0 : block.ctrl) && block.checkBox)
                    //@ts-ignore
                    .map(block => [block.ctrl, block.checkBox.checked]);
                blocks.forEach(block => block === null || block === void 0 ? void 0 : block.container.remove()); //We start by removing all the containers
                for (const [ctrl, checked] of values) {
                    if (!ctrl)
                        continue;
                    const subOptions = await getSubOptions(ctrl);
                    if (checked)
                        await isSelected(ctrl.title, subOptions);
                    else
                        isNotSelected(ctrl.title, subOptions);
                }
                resolve(selected);
            }
            ;
        });
    }
    async function isSelected(title, subOptions) {
        selected.push(title);
        if (subOptions)
            await showSelectPrompt(subOptions);
    }
    ;
    function isNotSelected(title, subOptions) {
        const exclude = (title) => `!${title}`;
        selected.push(exclude(title));
        subOptions === null || subOptions === void 0 ? void 0 : subOptions.forEach(ctrl => selected.push(exclude(ctrl.title)));
        console.log(selected);
    }
    ;
    async function getSubOptions(ctrl) {
        if (!ctrl)
            return;
        const children = ctrl.getContentControls();
        children.load(['id', 'tag', 'title', 'parentContentControl']);
        await ctrl.context.sync();
        return getSelectCtrls(children.items).filter(c => { var _a; return ((_a = c.parentContentControl) === null || _a === void 0 ? void 0 : _a.id) === ctrl.id; }); //!We need to make sure we get only the direct children of the ctrl and not all the nested ctrls
    }
    async function duplicateBlock(ctrl) {
        const replace = Word.InsertLocation.replace;
        const after = Word.InsertLocation.after;
        try {
            await duplicate();
        }
        catch (error) {
            showNotification(`${error}`);
        }
        async function duplicate() {
            const label = getFirstByTag(ctrl, RTSectionTag).getRange('Content');
            if (!label)
                return showNotification(`No Section RT Within the Range of the Duplicate Ctrl. Ctrl id = ${ctrl.id}`);
            label.load(['text']);
            await ctrl.context.sync();
            if (!label.text)
                return showNotification("No lable text");
            ctrl.select();
            const message = `How many ${label.text} parties are there?`;
            const answer = Number(await promptForInput(message));
            if (isNaN(answer))
                return showNotification(`The provided text cannot be converted into a number: ${answer}`);
            await insertClones(ctrl, answer);
        }
        async function insertClones(ctrl, answer) {
            const title = getCtrlTitle(ctrl.tag, ctrl.id);
            ctrl.title = title; //!We update the title in case it is no matching the id in the template.
            const ctrlContent = ctrl.getOoxml();
            await ctrl.context.sync();
            for (let i = 1; i < answer; i++)
                ctrl.getRange().insertOoxml(ctrlContent.value, after);
            const clones = ctrl.context.document.getContentControls().getByTitle(title);
            clones.load(['id', 'tag', 'title']);
            await ctrl.context.sync();
            const items = clones.items; //!clones.items.entries() caused the for loop to fail in scriptLab. The reason is unknown
            for (const clone of items)
                await processClone(clone, items.indexOf(clone) + 1);
            await ctrl.context.sync();
        }
        async function processClone(clone, i) {
            if (!clone)
                return;
            clone.title = getCtrlTitle(clone.tag, clone.id) + `-${i}`;
            const children = clone.getRange().getContentControls();
            children.load(['id', 'tag', 'title']);
            const label = children.getByTag(RTSectionTag).getFirst();
            label.load(['text']);
            await clone.context.sync();
            children.items
                .filter(ctrl => ctrl !== clone)
                .forEach(ctrl => ctrl.title = getCtrlTitle(ctrl.tag, ctrl.id)); //!We must update the title of the ctrls in order to udpated them with the new id
            const text = `${label.text} ${i}`;
            label.cannotEdit = false;
            label.select();
            label.getRange('Content').insertText(text, replace);
            const div = createHTMLElement('div', '', text, undefined, '', false);
            USERFORM.insertAdjacentElement('beforebegin', div);
            const selectCtrls = getSelectCtrls(children.items);
            await showSelectPrompt(selectCtrls);
            div.remove();
        }
        ;
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
function getCtrlTitle(tag, id) {
    return `${tag}&${id}`;
}
function getFirstByTag(range, tag) {
    return range.getContentControls().getByTag(tag).getFirst();
}
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
    if (!parent)
        return el;
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
function deleteCtrlById() {
    Word.run(async (context) => {
        let title = await promptForInput('Provide the title or the id of the control. If you provide the title, the id will be extracted from it');
        if (!title)
            return showNotification(`You did not provide a valid id or title: ${title}`);
        if (title.includes('&'))
            title = title.split('&')[1];
        const id = Number(title);
        if (isNaN(id))
            return showNotification(`The id could not be extracted from the title: ${title}`);
        const ctrl = context.document.contentControls.getById(id);
        const ctrls = ctrl.getContentControls();
        ctrls.load(['tag', 'id']);
        await context.sync();
        console.log('Ctrls = ', ctrls.items.map(c => c.tag));
        ctrl.cannotDelete = false;
        ctrls.items.forEach(ctrl => ctrl.cannotDelete = false);
        ctrl.delete(false);
        await context.sync();
    });
}
function updateAllCtrlsTitles() {
    Word.run(async (context) => {
        const ctrls = context.document.getContentControls();
        ctrls.load(['title', 'tag', 'id']);
        await context.sync();
        ctrls.items
            .filter(ctrl => ctrl.tag)
            .forEach(ctrl => ctrl.title = getCtrlTitle(ctrl.tag, ctrl.id));
        await context.sync();
    });
}
async function selectAllCtrlsByTag(tag, color) {
    Word.run(async (context) => {
        const ctrls = context.document.getContentControls();
        ctrls.load(['tag']);
        await context.sync();
        const sameTag = ctrls.items.filter(c => c.tag === tag);
        await setCtrlsColor(sameTag, color);
        await setCtrlsFontColor(sameTag, color);
        await context.sync();
    });
}
async function setCtrlsColor(ctrls, color) {
    ctrls.forEach(ctrl => ctrl.color = color);
}
function setCtrlsFontColor(ctrls, color) {
    ctrls.forEach(c => c.getRange().font.color = color);
}
function setRangeStyle(objs, style) {
    objs.forEach(o => o.getRange().style = style);
}
async function finalizeContract() {
    const tags = [RTSiTag, RTDescriptionTag, RTObsTag, RTDropDownTag];
    await removeRTs(tags);
}
async function removeRTs(tags) {
    Word.run(async (context) => {
        for (const tag of tags) {
            const ctrls = context.document.getContentControls().getByTag(tag);
            ctrls.load('tag');
            await context.sync();
            ctrls.items.forEach(ctrl => {
                ctrl.select();
                ctrl.cannotDelete = false;
                ctrl.delete(tag === RTDropDownTag);
            });
            await context.sync();
        }
    });
}
async function changeAllSameTagCtrlsCannEdit(tag, edit) {
    Word.run(async (context) => {
        const ctrls = context.document.getContentControls().getByTag(tag);
        ctrls.load('tag');
        await context.sync();
        ctrls.items.forEach(ctrl => ctrl.cannotEdit = edit);
        await context.sync();
    });
}
//# sourceMappingURL=app.js.map