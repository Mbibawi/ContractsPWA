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
        insertBtn(insertRichTextContentControlAroundSelection);
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
    function insertBtn(fun) {
        if (!userForm)
            return;
        const btn = document.createElement('button');
        userForm.appendChild(btn);
        btn.innerText = 'Insert Rich Text';
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
    //@ts-ignore
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
        await context.sync();
        // abort if nothing is selected
        if (selection.isEmpty) {
            console.log('Please select some text first.');
            return;
        }
        // 3. Wrap the selection in a RichText content control
        const cc = selection.insertContentControl(Word.ContentControlType.richText);
        cc.tag = prompt('Enter a tag for the new Rich Text control:', 'MyTag') || 'MyTag';
        cc.title = prompt('Enter a title for the new Rich Text control:', 'My Title') || 'My Title';
        cc.appearance = Word.ContentControlAppearance.boundingBox;
        cc.color = "blue";
        // Log the content control properties
        console.log(`ContentControl created with ID: ${cc.id}, Tag: ${cc.tag}, Title: ${cc.title}`);
        await context.sync();
    });
}
//# sourceMappingURL=app.js.map