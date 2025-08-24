Office.onReady((info) => {
    // Check that we loaded into Word
 
    if (info.host === Office.HostType.Word) {
        const btnEditWord = document.getElementById("edit");
        
        if (btnEditWord)
            btnEditWord.ondblclick = () => sayHello('Contracts App Works');
        getRichTextContentControlTitles()
        .then(ctrls => {
            const userForm = document.getElementById("userFormSection");
            if (!userForm || !ctrls) return;
      
                console.log('RichText = ', ctrls);
            ctrls.forEach(ctrl => {
                if (!ctrl) return;
                const p = document.createElement('p');
                p.textContent = ctrl.title || 'NoTitle';
                p.id = ctrl.id.toString();
                userForm.appendChild(p);
                p.onclick = () => hideContentControlById(ctrl.id)
            });
            });
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

async function getRichTextContentControlTitles(): Promise<RichTextProps[] | void[]> {
    return Word.run(async (context: any) => {
        const getProps = (cc: RichText): RichTextProps => ({title: cc.title || 'NoTitle', id: cc.id});
      // 1. Grab the collection of all content controls in the document
      const allControls = context.document.contentControls;
      
      // 2. Queue up a load for each control’s title and type
      allControls.load("items/title,id,type");
      
      // 3. Execute the queued commands
      await context.sync();
      
      // 4. Filter to only Rich Text controls and collect their titles
        return (allControls.items as RichText[])
            .filter(cc => cc.type === Word.ContentControlType.richText)
            .map(cc => getProps(cc));
    });
}

/**
 * Hides the content control with the given ID by setting its appearance to "hidden".
 * @param ccId The unique ID (GUID as number) of the content control to hide.
 */
async function hideContentControlById(ccId: number): Promise<void> {
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
      cc.appearance = Word.ContentControlAppearance.hidden;
      
      // 3. Push the change
      await context.sync();
      console.log(`ContentControl id=${ccId} is now hidden.`);
    });
  }
  

  