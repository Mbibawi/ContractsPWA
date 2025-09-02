"use strict";
const OPTIONS = ['Select', 'Show', 'Edit'];
const RTSelectTag = 'Select';
const RTSelectTitle = 'RTSelect';
const RTObsTag = 'RTObs';
const RTDescriptionTag = 'RTDesc';
const RTDescriptionStyle = 'RTDescription';
const RTSiTag = 'RTSi';
const RTSiStyles = ['RTSi0cm', 'RTSi1cm', 'RTSi2cm', 'RTSi3cm', 'RTSi4cm'];
let USERFORM, NOTIFICATION;
let RichText, RichTextInline, RichTextParag, Bounding, Hidden;
Office.onReady((info) => {
    // Check that we loaded into Word
    if (info.host !== Office.HostType.Word)
        return showNotification('This addin is designed to work on Word only');
    USERFORM = document.getElementById('userFormSection');
    NOTIFICATION = document.getElementById('notification');
    RichText = Word.ContentControlType.richText;
    RichTextInline = Word.ContentControlType.richTextInline;
    RichTextParag = Word.ContentControlType.richTextParagraphs;
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
    function wrap(title, tag, label) {
        return [
            () => wrapSelectionWithContentControl(title, tag),
            label
        ];
    }
    ;
    const insertDescription = () => findTextAndWrapItWithContentControl([`"*"`, `«*»`], [RTDescriptionStyle], RTDescriptionTag, RTDescriptionTag, true);
    const btns = [
        wrap(RTSiTag, RTSiTag, 'Insert Single RT Si'),
        wrap(RTDescriptionTag, RTDescriptionTag, 'Insert Single RT Description'),
        wrap(RTSelectTitle, RTSelectTag, 'Insert Single RT Select'),
        wrap(RTObsTag, RTObsTag, 'Insert Single RT Obs'),
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
function openInputDialog(data) {
    let dialog;
    Office.context.ui.displayDialogAsync("https://mbibawi.github.io/ContractsPWA/dialog.html", {
        height: 60,
        width: 60,
        promptBeforeOpen: false,
        displayInIframe: false
    }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            showNotification(`failed to open an got ${asyncResult.error.message}`);
            return;
        }
        dialog = asyncResult.value;
        // Send initial payload to the dialog
        dialog.messageChild(JSON.stringify(data));
        // Optionally handle messages sent back from the dialog
        // Listen for messages from the dialog
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
    });
    async function onDialogMessage(arg) {
        //!args needs to be converted to an array of objects like {id: number, delete:boolean, text: string}
        const ctrls = JSON.parse(arg.message);
        const text = arg.message;
        dialog.close();
        await Word.run(async (context) => {
            //! Insert logic for looping the content controls and deleting those whose delete property is set to true, and updating the text of those with a text property.
            ctrls.forEach(async (ctrl) => {
                const cc = context.document.contentControls.getByIdOrNullObject(ctrl.id);
                cc.load("isNullObject");
                await context.sync();
                if (cc.isNullObject) {
                    showNotification(`ContentControl id=${ctrl.id} not found.`);
                    return;
                }
                if (ctrl.delete) {
                    cc.delete(true);
                }
                else if (ctrl.text) {
                    cc.insertText(ctrl.text, Word.InsertLocation.replace);
                }
                await context.sync();
            });
        });
    }
}
/**
 * Wraps every occurrence of text formatted with a specific character style in a rich text content control.
 *
 * @param style The name of the character style to find (e.g., "Emphasis", "Strong", "MyCustomStyle").
 * @returns A Promise that resolves when the operation is complete.
 */
async function findTextAndWrapItWithContentControl(search, styles, title, tag, matchWildcards) {
    await Word.run(async (context) => {
        for (const el of search) {
            const ranges = await searchString(el, context, matchWildcards);
            if (!ranges)
                continue;
            await wrapMatchingStyleRangesWithContentControls(ranges, styles, title, tag);
        }
        ;
    });
}
async function wrapMatchingStyleRangesWithContentControls(ranges, styles, title, tag) {
    ranges.load(['style', 'parentContentControlOrNullObject', 'parentContentControlOrNullObject.isNullObject', 'parentContentControlOrNullObject.tag']);
    await ranges.context.sync();
    return ranges.items.map(async (range, index) => {
        if (!styles.includes(range.style))
            return;
        const parent = range.parentContentControlOrNullObject;
        if (!parent.isNullObject && parent.tag === tag)
            return;
        return await insertContentControl(range, title, tag, index, range.style);
    });
}
async function searchString(search, context, matchWildcards) {
    const searchResults = context.document.body.search(search, { matchWildcards: matchWildcards });
    await context.sync();
    if (!searchResults.items.length) {
        showNotification(`No text matching the search string was found in the document.`);
        return;
    }
    showNotification(`Found ${searchResults.items.length} ranges matching the search string: ${search}.`);
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
                await insertContentControl(parag.getRange('Content'), RTSiTag, RTSiTag, parags.indexOf(parag), parag.style);
            }
            catch (error) {
                showNotification(`error: ${error}`);
                continue;
            }
        }
        await context.sync();
    });
}
async function insertContentControl(range, title, tag, index, style) {
    range.select();
    // Insert a rich text content control around the found range.
    const ctrl = range.insertContentControl();
    ctrl.load(['id']);
    if (!style)
        style;
    await range.context.sync();
    // Set properties for the new content control.
    ctrl.title = `${title}-${ctrl.id}`;
    ctrl.tag = tag;
    ctrl.cannotDelete = true;
    ctrl.cannotEdit = true;
    ctrl.appearance = Bounding;
    if (style)
        ctrl.style = style;
    showNotification(`Wrapped text in range ${index || 1} with a content control.`);
    return ctrl;
}
async function wrapAllSameStyleParagraphsWithContentControl(style, title, tag) {
    const range = await getSelectionRange();
    if (!range || range.style !== style)
        return;
    await insertContentControl(range, title, tag, 0, style);
}
;
async function getSelectionRange() {
    return await Word.run(async (context) => {
        const range = context.document
            .getSelection()
            .getRange('Content');
        range.load(['style', 'isEmpty']);
        await context.sync();
        if (range.isEmpty)
            return showNotification('The selection range is empty');
        return range;
    });
}
async function wrapSelectionWithContentControl(title, tag) {
    const range = await getSelectionRange();
    if (!range)
        return;
    await insertContentControl(range, title, tag, 0, range.style);
}
;
async function confirm(question, fun) {
    if (!question)
        return;
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
            allRT.load(['title', 'tag', 'contentControls']);
            await context.sync();
            const ctrls = allRT.items
                .filter(ctrl => OPTIONS.includes(ctrl.tag))
                .entries();
            const selected = [];
            for (const ctrl of ctrls)
                await promptForSelection(ctrl, selected);
            const keep = selected.filter(title => !title.startsWith('!'));
            showNotification(`keep = ${keep.join(', ')}`);
            try {
                await currentDoc();
                await createNewDoc();
            }
            catch (error) {
                showNotification(`${error}`);
            }
            async function currentDoc() {
                for (const [i, ctrl] of ctrls) {
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
    async function promptForSelection([index, ctrl], selected) {
        if (selected.find(t => t.includes(ctrl.title)))
            return; //!We need to exclude any ctrl that has already been passed to the function or has been excluded: when a ctrl is excluded, its children are added to the array as excluded ctrls ("![ctrl.title]"), they do not hence need to be treated again since we already know theyare to be  excluded. This also avoids the problem that happens sometimes, when a ctrl has its parent amongst its children list (this is an apparently known weird behavior if the ctrl range overlaps somehow with the range of another ctrl)
        ctrl.select();
        const [container, btnNext, checkBox] = await showUI();
        return new Promise((resolve, reject) => {
            btnNext.onclick = () => nextCtrl(ctrl, checkBox);
            async function nextCtrl(ctrl, checkBox) {
                const checked = checkBox.checked;
                container.remove();
                ctrl.contentControls.load(['title', 'tag']);
                await ctrl.context.sync();
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
            const entries = subOptions.entries();
            for (const entry of entries)
                await promptForSelection(entry, selected);
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
function promptForInput(question, fun) {
    if (!question)
        return;
    const container = createHTMLElement('div', 'promptContainer', '', USERFORM);
    const prompt = createHTMLElement('div', 'prompt', '', container);
    const ask = createHTMLElement('p', 'ask', question, prompt);
    const input = createHTMLElement('input', 'answer', '', prompt);
    const btns = createHTMLElement('div', 'btns', '', prompt);
    const btnOK = createHTMLElement('button', 'btnOK', 'OK', btns);
    const btnCancel = createHTMLElement('button', 'btnCancel', 'Cancel', btns);
    let answer = '';
    btnOK.onclick = () => {
        answer = input.value;
        console.log('user answer = ', answer);
        container.remove();
        if (fun)
            fun(answer);
    };
    btnCancel.onclick = () => container.remove();
    return answer;
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
//# sourceMappingURL=app.js.map