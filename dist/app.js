"use strict";
//@ts-nocheck
Office.onReady((info) => {
    var _a;
    // Check that we loaded into Word
    if (info.host === Office.HostType.Word) {
        (_a = document.getElementById("helloButton")) === null || _a === void 0 ? void 0 : _a.onclick = sayHello;
        const btnEditWord = document.getElementById("edit");
        btnEditWord === null || btnEditWord === void 0 ? void 0 : btnEditWord.onclick = () => altert("Edit Word Is Working");
    }
});
function sayHello() {
    return Word.run((context) => {
        // insert a paragraph at the start of the document.
        const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.start);
        // sync the context to run the previous API call, and return.
        return context.sync();
    });
}
//# sourceMappingURL=app.js.map