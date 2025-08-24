//@ts-ignore
Office.onReady((info) => {
    // Check that we loaded into Word
    //@ts-ignore
    if (info.host === Office.HostType.Word) {
        //document.getElementById("helloButton")?.onclick = sayHello;
        const btnEditWord = document.getElementById("edit");
        //btnEditWord?.onclick = () => alert("Edit Word Is Working");
        if (btnEditWord)
            btnEditWord.ondblclick = () => sayHello('Contracts App Works');
    }
});

function sayHello(sentence: string) {
    //@ts-ignore
    return Word.run((context) => {

        // insert a paragraph at the start of the document.
        //@ts-ignore
        const paragraph = context.document.body.insertParagraph(sentence, Word.InsertLocation.start);
        
        // sync the context to run the previous API call, and return.
        return context.sync();
    });
}