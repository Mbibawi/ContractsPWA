//@ts-nocheck

Office.onReady((info) => {
    // Check that we loaded into Word
    if (info.host === Office.HostType.Word) {
        //document.getElementById("helloButton")?.onclick = sayHello;
        const btnEditWord = document.getElementById("edit");
        btnEditWord?.onclick = () => alert("Edit Word Is Working");
        btnEditWord?.ondblclick = () => sayHello();
    }
});

function sayHello() {
    return Word.run((context) => {

        // insert a paragraph at the start of the document.
        const paragraph = context.document.body.insertParagraph("Contracts App Works", Word.InsertLocation.start);
        
        // sync the context to run the previous API call, and return.
        return context.sync();
    });
}