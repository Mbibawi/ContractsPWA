//@ts-ignore
import * as Word from "@microsoft/office-js/word";

//@ts-ignore
Office.onReady((info) => {
    // Check that we loaded into Word
 
    if (info.host === Office.HostType.Word) {
        const btnEditWord = document.getElementById("edit");

        if (btnEditWord)
            btnEditWord.ondblclick = () => sayHello('Contracts App Works');
        getRichTextContentControlTitles()
            .then(titels => console.log('Found titles = ', titels));
    }
});

function sayHello(sentence: string) {
    //@ts-ignore
    return Word.run((context) => {

        // insert a paragraph at the start of the document.
        
        const paragraph = context.document.body.insertParagraph(sentence, Word.InsertLocation.start);
        
        // sync the context to run the previous API call, and return.
        return context.sync();
    });
}

async function getRichTextContentControlTitles(): Promise<string[]> {
    return Word.run(async (context:any) => {
      // 1. Grab the collection of all content controls in the document
      const allControls = context.document.contentControls;
      
      // 2. Queue up a load for each control’s title and type
      allControls.load("items/title,type");
      
      // 3. Execute the queued commands
      await context.sync();
      
      // 4. Filter to only Rich Text controls and collect their titles
      const titles: string[] = (allControls.items as RichText[])
        .filter(cc => cc.type === Word.ContentControlType.richText)
        .map(cc => cc.title);
  
      // 5. (Optional) Log or return the titles
      console.log("Rich Text Content Control Titles:", titles);
      return titles;
    });
  }
  
  // Example invocation
  getRichTextContentControlTitles()
    .then(titles => {
      // Do something with the titles array
    })
    .catch(error => console.error(error));