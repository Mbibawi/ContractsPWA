"use strict";
const USERFORM = document.getElementById('userFormSection');
const OPTIONS = ['Select', 'Show', 'Edit'];
const RTDescriptionTag = 'RTDesc';
const RTDescriptionStyle = 'RTDescription';
const RTSiTag = 'RTSi';
const RTSiStyle = 'RTSi';
Office.onReady((info) => {
    // Check that we loaded into Word
    if (info.host === Office.HostType.Word) {
        buildUI();
    }
});
function buildUI() {
    if (!USERFORM)
        return;
    (function insertBtns() {
        insertBtn(customizeContract, 'Customize Document');
        return;
        insertBtn(insertRichTextContentControlAroundSelection, 'Insert Rich Text Control');
        insertBtn(openInputDialog, 'Open Input Dialog');
        insertBtn(() => wrapTextWithContentControlsByStyle([`"*"`, `«*»`], RTDescriptionStyle, RTDescriptionTag, true), 'Insert RT Description All');
    })();
    function insertBtn(fun, text) {
        if (!USERFORM)
            return;
        const btn = document.createElement('button');
        USERFORM.appendChild(btn);
        btn.innerText = text;
        btn.onclick = () => fun();
    }
}
async function insertRichTextContentControlAroundSelection() {
    await Word.run(async (context) => {
        // get the current selection
        const selection = context.document.getSelection();
        selection.load('isEmpty');
        await context.sync();
        // abort if nothing is selected
        if (selection.isEmpty) {
            console.log('Please select some text first.');
            return;
        }
        // 3. Wrap the selection in a RichText content control
        const cc = selection.insertContentControl(Word.ContentControlType.richText);
        //cc.tag = window.prompt('Enter a tag for the new Rich Text control:', 'MyTag') || 'MyTag';
        //cc.title = window.prompt('Enter a title for the new Rich Text control:', 'My Title') || 'My Title';
        cc.appearance = Word.ContentControlAppearance.boundingBox;
        cc.color = "blue";
        // Log the content control properties
        console.log(`ContentControl created with ID: ${cc.id}, Tag: ${cc.tag}, Title: ${cc.title}`);
        await context.sync();
    });
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
            console.error("failed to open", asyncResult.error.message);
            return;
        }
        dialog = asyncResult.value;
        // Send initial payload to the dialog
        dialog.messageChild(JSON.stringify(data));
        // Optionally handle messages sent back from the dialog
        // Listen for messages from the dialog
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
    });
    function updateCtrls() {
        //This function will get the updated content controls from the document and return them as a string to be sent to the dialog;
        //for any select element in the dialog, each option will be converted into an object like {id: number, text: string, delete: boolean} where id is the id of the option, text = null. If the option is the selected option, the delete will be false, otherwise true.
        const message = [];
        document.querySelectorAll('.dropDown')
            .forEach(select => {
            const options = select.options;
            Array.from(options).forEach(opt => message.push({
                id: Number(opt.id),
                content: null,
                delete: !opt.selected
            }));
        });
        document.querySelectorAll('.checkBox')
            .forEach(chbx => {
            message.push({
                id: Number(chbx.id),
                content: null,
                delete: !chbx.checked
            });
        });
        return JSON.stringify(message);
    }
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
                    console.warn(`ContentControl id=${ctrl.id} not found.`);
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
async function getDocumentBase64() {
    return new Promise(async (resolve, reject) => {
        // 1. Request the document as a compressed file
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 64 * 1024 }, (file) => processFile(file) // 64KB per slice
        );
        function processFile(result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                reject(result.error);
                return;
            }
            const file = result.value;
            const sliceCount = file.sliceCount;
            const slices = [];
            let loaded = 0;
            // 2. Pull down each slice
            for (let i = 0; i < sliceCount; i++) {
                file.getSliceAsync(i, (sliceResult) => {
                    if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                        slices[sliceResult.value.index] = sliceResult.value.data;
                        loaded++;
                        // 3. Once all slices are in, close and resolve
                        if (loaded === sliceCount) {
                            file.closeAsync(() => {
                                resolve(slices.join(""));
                            });
                        }
                    }
                    else {
                        file.closeAsync(() => reject(sliceResult.error));
                    }
                });
            }
        }
    });
}
/**
 * Creates a new Word document based on the current document as a template,
 * then deletes a specified list of content controls by their ID from the new document.
 * This function handles the entire process asynchronously.
 *
 * @param toDelete An array of unique IDs for the content controls to be deleted.
 * @returns A Promise that resolves when the operation is complete.
 */
async function generateCustomizedContract(toDelete, toEdit) {
    // Wrap the entire function in a try/catch block for robust error handling.
    try {
        // Step 1: Get the current document's content as a Base64-encoded string.
        // The getBase64() method provides a file representation of the document.
        let templateContent = await getDocumentBase64();
        if (!templateContent) {
            console.error("Failed to get base64 string from the current document. Aborting.");
            return;
        }
        console.log("Successfully captured the current document as a template.");
        console.log("Creating new document...");
        // Step 2: Use context.application.createDocument() to create a new document.
        // This method returns a promise that resolves to a DocumentCreated object.
        const newDocument = await Word.run(async (context) => {
            const documentCreated = context.application.createDocument(templateContent);
            await context.sync();
            return documentCreated;
        });
        // Step 3: Run commands in the context of the new document.
        // We must call open() on the DocumentCreated object to switch the context.
        await newDocument.open();
        await processCtrls(newDocument, toDelete, deleteCtrl);
        await processCtrls(newDocument, toEdit, editCtrlText);
    }
    catch (error) {
        console.error("An error occurred during the process:", error);
        // Provide user-friendly error details if it's an Office Extension Error.
        if (error instanceof OfficeExtension.Error) {
            console.log(`Office Extension Error: ${error.code} - ${error.message}`);
        }
    }
}
async function processCtrls(wdDoc, ctrls, fun) {
    if (!wdDoc || !ctrls || !fun)
        return console.log('Either the document or the ctrls collection is/are missing');
    await Word.run(wdDoc, async (context) => {
        var _a;
        // Step 4: Iterate through the list of IDs and delete the corresponding content controls.
        for (const ctrl of ctrls) {
            if (!ctrl.title)
                continue;
            const contentControl = (_a = context.document.contentControls.getByTitle(ctrl.title)) === null || _a === void 0 ? void 0 : _a.items[0];
            if (!contentControl)
                continue;
            fun(contentControl, ctrl);
            // Load a property to check if the content control exists.
            //contentControl.load(['isNullObject', 'title']);
            // Synchronize the state with the new document.
            //await context.sync();
            //if (contentControl.isNullObject) continue; // Skip to the next ID if the control doesn't exist.
        }
        // Step 5: Execute all the delete commands at once on the new document.
        await context.sync();
        console.log("All specified content controls have been deleted from the new document.");
    });
}
// Delete the content control and all of its content.
function deleteCtrl(ctrl, data) {
    if (!ctrl)
        return;
    ctrl.delete(true);
    console.log(`Deleted content control with ID ${data.id}.`);
}
function editCtrlText(ctrl, data) {
    if (!data.content || !ctrl)
        return;
    // Edit the content control and all of its content.
    const range = ctrl.getRange();
    //range.clear();
    range.insertText(data.content, "Replace");
    console.log(`Edited content control with ID ${data.id}.`);
}
/**
 * Wraps every occurrence of text formatted with a specific character style in a rich text content control.
 *
 * @param style The name of the character style to find (e.g., "Emphasis", "Strong", "MyCustomStyle").
 * @returns A Promise that resolves when the operation is complete.
 */
async function wrapTextWithContentControlsByStyle(search, style, tag, matchWildcards) {
    await Word.run(async (context) => {
        for (const el of search) {
            if (!el)
                continue;
            await searchString(el);
        }
        ;
        const inserted = context.document.contentControls.getByTag(tag);
        await addIDtoCtrlTitle(inserted);
        console.log("Operation complete. All matching text ranges are now wrapped in content controls.");
        async function searchString(search) {
            const searchResults = context.document.body.search(search, { matchWildcards: matchWildcards });
            searchResults.load(['style', 'parentContentControlOrNullObject.tag', 'parentContentControlOrNullObject.isNullObject']);
            await context.sync();
            if (!searchResults.items.length) {
                console.log(`No text with the style "${style}" was found in the document.`);
                return;
            }
            console.log(`Found ${searchResults.items.length} ranges with the style "${style}".`);
            await context.sync();
            searchResults.items.map(async (range, index) => {
                if (!range.style || range.style !== style)
                    return;
                const parent = range.parentContentControlOrNullObject;
                if (!parent.isNullObject && parent.tag === tag)
                    return;
                return await insertContentControl(range, tag, tag, index);
            });
        }
    });
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
        const parags = paragraphs.items;
        console.log(parags);
        for (const parag of parags) {
            try {
                parag.select();
                if (!parag.style.startsWith(RTSiStyle))
                    continue;
                const parent = parag.parentContentControlOrNullObject;
                parent.load(['tag']);
                await parag.context.sync();
                if (parent.tag === RTSiTag)
                    continue;
                console.log(`range style: ${parag.style} & text = ${parag.text}`);
                await insertContentControl(parag.getRange('Content'), RTSiTag, RTSiTag, parags.indexOf(parag));
            }
            catch (error) {
                console.log(`error: ${error}`);
                continue;
            }
        }
        await context.sync();
    });
}
async function insertContentControl(range, title, tag, index) {
    // Insert a rich text content control around the found range.
    const contentControl = range.insertContentControl();
    contentControl.load(['id']);
    await range.context.sync();
    // Set properties for the new content control.
    contentControl.title = `${title}-${contentControl.id}`;
    contentControl.tag = tag;
    contentControl.cannotDelete = true;
    contentControl.cannotEdit = true;
    contentControl.appearance = Word.ContentControlAppearance.boundingBox;
    console.log(`Wrapped text in range ${index || 1} with a content control.`);
    return contentControl;
}
function wrapSelectedTextWithContentControl() {
}
function promptForInput(question) {
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
    };
    btnCancel.onclick = () => container.remove();
    return answer;
}
async function customizeContract() {
    return await Word.run(async (context) => {
        const allRT = context.document.contentControls;
        allRT.load(['title', 'tag']);
        await context.sync();
        const ctrls = allRT.items
            .filter(ctrl => OPTIONS.includes(ctrl.tag));
        const selected = [];
        for (const ctrl of ctrls) {
            await promptForSelection(ctrl, selected);
        }
        const keep = selected.filter(title => !title.startsWith('!'));
        const template = await getDocumentBase64();
        const newDoc = context.application.createDocument(template);
        await deleteAllNotSelected(keep, newDoc);
    });
}
async function deleteAllNotSelected(selected, document) {
    const all = document.contentControls;
    all.load(['title', 'tag']);
    await document.context.sync();
    all.items
        .filter(ctrl => !selected.includes(ctrl.title))
        .forEach(ctrl => {
        ctrl.select();
        ctrl.cannotDelete = false;
        ctrl.delete(true);
    });
    await document.context.sync();
}
async function _fixRTSelect() {
    await Word.run(async (context) => {
        const ctrls = context.document.contentControls;
        ctrls.load(['id', 'style', 'paragraphs', 'contentControls']);
        await context.sync();
        for (const ctrl of ctrls.items) {
            ctrl.contentControls.load(['title', 'tag', 'contentControls']);
            await context.sync();
            if (!ctrl.contentControls.items.length)
                continue;
            const first = ctrl.contentControls.getFirst();
            first.load('tag');
            await context.sync();
            if (first.tag === RTSiTag) {
                ctrl.tag = 'Select';
                ctrl.title = `RTSelect-${ctrl.id}`;
                await context.sync();
            }
        }
        await context.sync();
    });
}
async function _fixRTSi() {
    await Word.run(async (context) => {
        const ctrls = context.document.contentControls;
        ctrls.load(['id', 'style', 'tag']);
        await context.sync();
        for (const ctrl of ctrls.items) {
            if (ctrl.tag !== RTSiTag)
                continue;
            ctrl.select();
            ctrl.title = `${RTSiTag}-${ctrl.id}`;
        }
        await context.sync();
    });
}
async function _fixRTDesc() {
    await Word.run(async (context) => {
        const ctrls = context.document.contentControls;
        ctrls.load(['id', 'style', 'tag']);
        await context.sync();
        for (const ctrl of ctrls.items) {
            if (ctrl.tag !== RTSiTag)
                continue;
            ctrl.select();
            ctrl.title = `${RTDescriptionTag}-${ctrl.id}`;
        }
        await context.sync();
    });
}
async function _setCannotDelet() {
    await Word.run(async (context) => {
        const ctrls = context.document.contentControls;
        ctrls.load(['id', 'style', 'paragraphs', 'contentControls']);
        await context.sync();
        for (const ctrl of ctrls.items) {
            ctrl.contentControls.load(['title', 'tag', 'contentControls']);
            ctrl.cannotDelete = true;
            await context.sync();
        }
        await context.sync();
    });
}
async function setRTSiTag() {
    await Word.run(async (context) => {
        const ctrls = context.document.contentControls;
        ctrls.load(['id', 'style', 'paragraphs', 'contentControls']);
        await context.sync();
        for (const ctrl of ctrls.items) {
            ctrl.contentControls.load(['title', 'tag']);
            const first = ctrl.paragraphs.getFirst();
            first.load(['style']);
            await context.sync();
            if (first.style.startsWith(RTSiStyle))
                ctrl.tag = RTSiTag;
            ctrl.title = `${RTSiTag}-${ctrl.id}`;
        }
        await context.sync();
    });
}
async function promptForSelection(ctrl, selected) {
    const exclude = (title) => `!${title}`;
    if (selected.includes(exclude(ctrl.title)))
        return;
    const [container, btnNext, checkBox] = await showUI();
    return new Promise((resolve, reject) => {
        btnNext.onclick = nextCtrl;
        async function nextCtrl() {
            if (checkBox.checked)
                await isSelected(ctrl);
            else
                await isNotSelected(ctrl);
            container.remove();
            resolve(selected);
        }
        ;
    });
    async function isSelected(ctrl) {
        selected.push(ctrl.title);
        const subOptions = await getChildren(ctrl);
        for (const ctrl of subOptions) {
            await promptForSelection(ctrl, selected);
        }
        console.log(selected);
    }
    ;
    async function isNotSelected(ctrl) {
        selected.push(exclude(ctrl.title));
        const subOptions = await getChildren(ctrl);
        subOptions
            .forEach(ctrl => selected.push(exclude(ctrl.title)));
        console.log(selected);
    }
    ;
    async function getChildren(ctrl) {
        const children = ctrl.contentControls;
        children.load(['title', 'tag']);
        await ctrl.context.sync();
        return children
            .items
            .filter(ctrl => OPTIONS.includes(ctrl.tag));
    }
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
            const container = createHTMLElement('div', 'promptContainer', '', USERFORM);
            const prompt = createHTMLElement('div', 'selection', '', container);
            const label = createHTMLElement('label', 'label', text, prompt);
            const checkBox = createHTMLElement('input', 'checkBox', '', prompt);
            checkBox.type = 'checkbox';
            const btns = createHTMLElement('div', 'btns', '', prompt);
            const btnNext = createHTMLElement('button', 'btnOK', 'Next', btns);
            return [container, btnNext, checkBox];
        }
    }
}
function createHTMLElement(tag, css, innerText, parent, append = true) {
    const el = document.createElement(tag);
    if (innerText)
        el.innerText = innerText;
    el.classList.add(css);
    append ? parent.appendChild(el) : parent.prepend(el);
    return el;
}
async function _fixCTrlTitle(tag) {
    await Word.run(async (context) => {
        const ctrls = context.document.contentControls.getByTag(tag);
        ctrls.load('title');
        await context.sync();
        for (const ctrl of ctrls.items) {
            ctrl.title = ctrl.title.replace('}', '');
            console.log('ctrl title = ', ctrl.title);
        }
        await context.sync();
    });
}
//# sourceMappingURL=app.js.map