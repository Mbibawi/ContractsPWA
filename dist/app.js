"use strict";
Office.onReady((info) => {
    // Check that we loaded into Word
    if (info.host === Office.HostType.Word) {
        buildUI();
    }
});
function buildUI() {
    const userForm = document.getElementById("userFormSection");
    if (!userForm)
        return;
    (function addOnClick() {
        const btnEditWord = document.getElementById("edit");
        if (btnEditWord)
            btnEditWord.onclick = () => sayHello('Contracts App Works');
    })();
    (function insertBtns() {
        insertBtn(insertRichTextContentControlAroundSelection, 'Insert Rich Text Control');
        insertBtn(openInputDialog, 'Open Input Dialog');
        insertBtn(() => wrapTextWithContentControlsByStyle('RTDescription', 'RTDesc'), 'Insert RT Description');
    })();
    (function addElements() {
        getRichTextContentControlTitles()
            .then(ctrls => {
            console.log('RichText = ', ctrls);
            ctrls.forEach(ctrl => {
                if (!ctrl)
                    return;
                const p = document.createElement('p');
                p.textContent = ctrl.title || 'NoTitle';
                p.id = ctrl.id.toString();
                userForm.appendChild(p);
                p.onclick = () => deleteContentControl(ctrl.id);
            });
        });
    })();
    function insertBtn(fun, text) {
        if (!userForm)
            return;
        const btn = document.createElement('button');
        userForm.appendChild(btn);
        btn.innerText = text;
        btn.onclick = () => fun();
    }
}
function insertUIElements(cc) {
    if (cc.title.startsWith('List'))
        return dropDownList();
    else if (cc.title.startsWith('Opt'))
        return selectOption();
    else if (cc.title.startsWith('Cbx'))
        return checkBox();
    else
        return;
    function dropDownList() {
        const select = document.createElement('select');
        select.id = cc.id.toString();
        select.classList.add('dropDown');
    }
    function selectOption() {
        const option = document.createElement('option');
        option.id = cc.id.toString();
        option.classList.add('option');
        return option;
    }
    function checkBox() {
        const Cbx = document.createElement('input');
        Cbx.type = 'checkbox';
        Cbx.id = cc.id.toString();
        Cbx.classList.add('checkBox');
        return Cbx;
    }
}
function sayHello(sentence) {
    return Word.run((context) => {
        // insert a paragraph at the start of the document.
        const paragraph = context.document.body.insertParagraph(sentence, Word.InsertLocation.start);
        // sync the context to run the previous API call, and return.
        return context.sync();
    });
}
async function getRichTextContentControlTitles() {
    return Word.run(async (context) => {
        const getProps = (cc) => ({ title: cc.title || 'NoTitle', id: cc.id });
        // 1. Grab the collection of all content controls in the document
        const allControls = context.document.contentControls;
        // 2. Queue up a load for each control’s title and type
        allControls.load("items/title,id,type");
        // 3. Execute the queued commands
        await context.sync();
        // 4. Filter to only Rich Text controls and collect their titles
        return allControls.items
            .filter(cc => cc.type === Word.ContentControlType.richText)
            .map(cc => getProps(cc));
    });
}
/**
 * Hides the content control with the given ID by setting its appearance to "hidden".
 * @param ccId The unique ID (GUID as number) of the content control to hide.
 */
async function deleteContentControl(ccId) {
    await Word.run(async (context) => {
        // 1. Try to get the control by its ID (returns a null object if not found)
        const cc = context.document.contentControls.getByIdOrNullObject(ccId);
        cc.load("isNullObject");
        await context.sync();
        if (cc.isNullObject) {
            console.warn(`ContentControl id=${ccId} not found.`);
            return;
        }
        // 2. Set its appearance to hidden (no bounding box or tag marks)
        //cc.appearance = Word.ContentControlAppearance.hidden;
        //2. delete the contentControl and its content
        cc.delete(true);
        // 3. Push the change
        await context.sync();
        console.log(`ContentControl id=${ccId} is now hidden.`);
    });
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
        await Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 64 * 1024 }, (file) => processFile(file) // 64KB per slice
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
async function wrapTextWithContentControlsByStyle(style, tag) {
    await Word.run(async (context) => {
        // The Word.run context must be used for all operations.
        // Use a search option to search for all ranges with the specified character style.
        const searchResults = context.document.body.search("”*”", { matchWildcards: true });
        searchResults.load('style');
        await context.sync();
        if (!searchResults.items.length) {
            console.log(`No text with the style "${style}" was found in the document.`);
            return;
        }
        console.log(`Found ${searchResults.items.length} ranges with the style "${style}".`);
        // Iterate through the ranges in reverse order to avoid issues with the document changing.
        // When you insert a new content control, it can affect the ranges of other items in the collection.
        // By iterating in reverse, the ranges that haven't been processed yet remain valid.
        searchResults.items.map((range, index) => {
            if (!range.style || range.style !== style)
                return;
            // Insert a rich text content control around the found range.
            range.select("Select");
            const contentControl = range.insertContentControl();
            // Set properties for the new content control.
            contentControl.title = `${style}`;
            contentControl.tag = tag;
            contentControl.cannotDelete = true;
            contentControl.cannotEdit = true;
            contentControl.appearance = Word.ContentControlAppearance.boundingBox;
            console.log(`Wrapped text in range ${index} with a content control.`);
            return contentControl;
        });
        await context.sync();
        const inserted = context.document.contentControls.getByTag(tag);
        inserted.load('id');
        await context.sync();
        inserted.items.forEach(ctrl => ctrl.title = `${ctrl.title}-${ctrl.id}`);
        await context.sync();
        console.log("Operation complete. All matching text ranges are now wrapped in content controls.");
    });
}
//# sourceMappingURL=app.js.map