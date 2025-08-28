const OPTIONS = ['Select', 'Show', 'Edit'];
const RTSelectTag = 'Select';
const RTSelectTitle = 'RTSelect';
const RTObsTag = 'RTObs';
const RTDescriptionTag = 'RTDesc';
const RTDescriptionStyle = 'RTDescription';
const RTSiTag = 'RTSi';
const RTSiStyles = ['RTSi0cm', 'RTSi1cm', 'RTSi2cm', 'RTSi3cm', 'RTSi4cm'];
let USERFORM:HTMLDivElement;
Office.onReady((info) => {
    USERFORM = document.getElementById('userFormSection') as HTMLDivElement
    // Check that we loaded into Word
 
    if (info.host === Office.HostType.Word) {
        buildUI();
    }
});

function buildUI() {
    if (!USERFORM) return;
    
    (function insertBtns() {
        insertBtn(customizeContract, 'Customize Contract');
        insertBtn(prepareTemplate, 'Prepare Template');
        
        function prepareTemplate() {
            USERFORM.innerHTML = ''
            insertBtn(()=>wrapSelectionWithContentControl(RTSiTag, RTSiTag), 'Insert Single RT Si');
            insertBtn(()=>wrapSelectionWithContentControl(RTDescriptionTag, RTDescriptionTag), 'Insert Single RT Description');
            insertBtn(()=>wrapSelectionWithContentControl(RTSelectTitle, RTSelectTag), 'Insert Single RT Select');
            insertBtn(()=>wrapSelectionWithContentControl(RTObsTag, RTObsTag), 'Insert Single RT Obs');
            insertBtn(insertRTSiAll, 'Insert RT Si For All');
            insertBtn(() => findTextAndWrapItWithContentControl([`"*"`, `«*»`], [RTDescriptionStyle], RTDescriptionTag, RTDescriptionTag, true), 'Insert RT Description For All');
        }
    })();
    

    function insertBtn(fun:Function, text:string) {
        if (!USERFORM) return;
        const btn = document.createElement('button');
        USERFORM.appendChild(btn);
        btn.innerText = text;
        btn.onclick = () => fun();
    }
}
 

  async function insertRichTextContentControlAroundSelection(): Promise<void> {
    await Word.run(async context => {
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
  

function openInputDialog(data: object) {
    let dialog: Office.Dialog;
    Office.context.ui.displayDialogAsync(
        "https://mbibawi.github.io/ContractsPWA/dialog.html",
        {
            height: 60,
            width: 60,
            promptBeforeOpen: false,
            displayInIframe: false
        },
        (asyncResult) => {
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
        }
    );

    function updateCtrls(): string{
        //This function will get the updated content controls from the document and return them as a string to be sent to the dialog;
        //for any select element in the dialog, each option will be converted into an object like {id: number, text: string, delete: boolean} where id is the id of the option, text = null. If the option is the selected option, the delete will be false, otherwise true.
        const message:contentControl[] = [];

        (document.querySelectorAll('.dropDown') as NodeListOf<HTMLSelectElement>)
            .forEach(select => {
            const options = select.options;
                Array.from(options).forEach(opt =>
                message.push({
                id: Number(opt.id),
                content: null,
                delete: !opt.selected
            }))
            });
        
        (document.querySelectorAll('.checkBox') as NodeListOf<HTMLInputElement>)
            .forEach(chbx=> {
                message.push({
                    id: Number(chbx.id),
                    content: null,
                    delete: !chbx.checked
                })
            });
        
        return JSON.stringify(message)

    }

    async function onDialogMessage(arg: any) {
        //!args needs to be converted to an array of objects like {id: number, delete:boolean, text: string}
        const ctrls = JSON.parse(arg.message) as { id: number, delete: boolean, text: string }[];
        const text = arg.message;
        dialog.close();
      
        await Word.run(async (context) => {
            //! Insert logic for looping the content controls and deleting those whose delete property is set to true, and updating the text of those with a text property.
            ctrls.forEach(async ctrl => {
                const cc = context.document.contentControls.getByIdOrNullObject(ctrl.id);
                cc.load("isNullObject");
                await context.sync();
                if (cc.isNullObject) {
                    console.warn(`ContentControl id=${ctrl.id} not found.`);
                    return;
                }
                if (ctrl.delete) {
                    cc.delete(true);
                } else if (ctrl.text) {
                    cc.insertText(ctrl.text, Word.InsertLocation.replace);
                }
                await context.sync();
            });
        });
    }
}


async function processCtrls(wdDoc: Word.DocumentCreated | Word.Document, ctrls: contentControl[] | undefined, fun: Function) { 

    if (!wdDoc || !ctrls || !fun) return console.log('Either the document or the ctrls collection is/are missing');
    await Word.run(wdDoc, async (context) => {
        // Step 4: Iterate through the list of IDs and delete the corresponding content controls.
        for (const ctrl of ctrls) {
            if (!ctrl.title) continue;
            const contentControl = context.document.contentControls.getByTitle(ctrl.title)?.items[0];
            if (!contentControl) continue;
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
function deleteCtrl(ctrl: Word.ContentControl, data: contentControl) {
    if (!ctrl) return;
    ctrl.delete(true);
    console.log(`Deleted content control with ID ${data.id}.`);
}

function editCtrlText(ctrl: Word.ContentControl, data: contentControl) {
    if (!data.content || !ctrl) return;
    // Edit the content control and all of its content.
    const range = ctrl.getRange();
    //range.clear();
    range.insertText(data.content, "Replace")

    console.log(`Edited content control with ID ${data.id}.`);
    
}

/**
 * Wraps every occurrence of text formatted with a specific character style in a rich text content control.
 *
 * @param style The name of the character style to find (e.g., "Emphasis", "Strong", "MyCustomStyle").
 * @returns A Promise that resolves when the operation is complete.
 */
async function findTextAndWrapItWithContentControl(search:string[], styles: string[], title:string, tag:string, matchWildcards:boolean): Promise<void> {
    await Word.run(async (context) => {
        for (const el of search) {
            const ranges = await searchString(el, context, matchWildcards);
            if (!ranges) continue;
            await wrapMatchingStyleRangesWithContentControls(ranges, styles, title, tag);
        };
        
    });
}

async function wrapMatchingStyleRangesWithContentControls(ranges: Word.RangeCollection, styles: string[], title: string, tag: string) {
    ranges.load(['style', 'parentContentControlOrNullObject', 'parentContentControlOrNullObject.isNullObject', 'parentContentControlOrNullObject.tag']);

    await ranges.context.sync();
    
    return ranges.items.map(async (range, index) => {
            if (!styles.includes(range.style)) return;
            const parent = range.parentContentControlOrNullObject;
            if (!parent.isNullObject && parent.tag === tag) return;
            return await insertContentControl(range, title, tag, index)  
    });
}

async function searchString(search:string, context:Word.RequestContext, matchWildcards:boolean) {
    const searchResults = context.document.body.search(search, { matchWildcards: matchWildcards });
    await context.sync();
    if (!searchResults.items.length) {
        console.log(`No text matching the search string was found in the document.`);
        return;
    }
    
    console.log(`Found ${searchResults.items.length} ranges matching the search string: ${search}.`);
            
    return searchResults    
}

async function addIDtoCtrlTitle(ctrls: Word.ContentControlCollection) {
    ctrls.load(['title','id']);
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
      console.log(parags)
      for (const parag of parags) {
          parag.select();
          try{
            const parent = parag.parentContentControlOrNullObject;
           parent.load(['tag']);
            await parag.context.sync();
            if(parent.tag ===RTSiTag) continue;   
            console.log(`range style: ${parag.style} & text = ${parag.text}`);
              await insertContentControl(parag.getRange('Content'), RTSiTag, RTSiTag, parags.indexOf(parag));
            }catch(error){
              console.log(`error: ${error}`);
              continue
            }
      }
      await context.sync();
  
    })
  }
async function insertContentControl(range: Word.Range, title: string, tag: string, index: number) {
    range.select();
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
    return contentControl
  
}
 
async function wrapAllSameStyleParagraphsWithContentControl(style:string, title:string, tag:string) {
    await Word.run(async (context)=>{
        const selection = context.document.getSelection();
        const range = selection.getRange('Content');
        range.load(['style']);
        await range.context.sync();
        if (range.style !== style) return;
        await insertContentControl(range,title, tag, 0)
    })
};

async function wrapSelectionWithContentControl(title:string, tag:string) {
    await Word.run(async (context)=>{
        const selection = context.document.getSelection();
        const range = selection.getRange('Content');
        await insertContentControl(range, title, tag, 0)
    })
};

function promptForInput(question: string) {
    if (!question) return;
    const container = createHTMLElement('div', 'promptContainer', '', USERFORM);
    const prompt = createHTMLElement('div', 'prompt', '', container);
    const ask = createHTMLElement('p', 'ask', question, prompt);
    const input =createHTMLElement('input', 'answer', '', prompt) as HTMLInputElement;
    const btns =createHTMLElement('div', 'btns', '', prompt);
    const btnOK =createHTMLElement('button', 'btnOK', 'OK', btns);
    const btnCancel =createHTMLElement('button', 'btnCancel', 'Cancel', btns);
    
    let answer:string = '';

    btnOK.onclick = () => {
        answer = input.value;
        console.log('user answer = ', answer);
        container.remove();
    };
    btnCancel.onclick = () => container.remove();
    return answer ;
};

async function customizeContract() {
    USERFORM.innerHTML = '';
    createHTMLElement('button', 'button', 'Download Document', USERFORM, '', true);
    const template = await getTemplate() as Base64URLString;
    console.log(template);
    if (!template) return console.log('Failed to create the template');
    await selectCtrls();

    async function selectCtrls() {
        return await Word.run(async (context) => {
            const allRT = context.document.contentControls;
            allRT.load(['title', 'tag', 'contentControls']);
            await context.sync();
           const ctrls = allRT.items
               .filter(ctrl => OPTIONS.includes(ctrl.tag))
                .entries();
            
            const selected: string[] = [];
            for (const ctrl of ctrls) 
               await promptForSelection(ctrl, selected);
           
            const keep = selected.filter(title => !title.startsWith('!'));
            const newDoc = context.application.createDocument(template);
            newDoc.open();
            //context.document.close(Word.CloseBehavior.skipSave);
            await deleteAllNotSelected(keep, newDoc);
            await context.sync();
        });
    }

    async function getTemplate() {
        try {
            const template = await getDocumentBase64();
            return template;
     } catch (error) {
         console.log(`Failed to create new Doc: ${error}`)
     }
    }
};
/**
 * Asynchronously gets the entire document content as a Base64 string.
 * This function handles multi-slice documents by requesting each slice in parallel.
 * @returns A Promise that resolves with the Base64-encoded document content.
 */
async function getDocumentBase64(): Promise<Base64URLString> {
    const failed = (result: Office.AsyncResult<Office.File | Office.Slice>) => result.status !== Office.AsyncResultStatus.Succeeded;
    
    return new Promise((resolve, reject) => {
        // Step 1: Request the document as a compressed file.
        Office.context.document.getFileAsync(
            Office.FileType.Compressed,
            { sliceSize: 64 * 1024 },
            (fileResult)=>processFile(fileResult)
        );

        function processFile(fileResult: Office.AsyncResult<Office.File>) {
            if (failed(fileResult))
               return reject(fileResult.error);

            const file = fileResult.value;
            const sliceCount = file.sliceCount;
            const slices: number[] = new Array(sliceCount);
            let loadedSlices = 0;

            // Step 2: Use a loop to request each slice in parallel.
            for (let i = 0; i < sliceCount; i++) {
                if(isNaN(i)) break
                file.getSliceAsync(i, (sliceResult) =>processSlice(sliceResult));
            };
            
            function processSlice(sliceResult: Office.AsyncResult<Office.Slice>) {
                if(failed(sliceResult)) 
                    file.closeAsync(() => reject(sliceResult.error));
                else{
                    // Store the raw data of the slice in the correct index.
                    slices[sliceResult.value.index] = sliceResult.value.data;
                    loadedSlices++
   
                    // Step 3: Check if all slices have been received.
                    if (loadedSlices === sliceCount)
                        file.closeAsync(()=>resolve(slices.join('')));
                } 
                
            }
        }
    });
}



async function promptForSelection([index, ctrl]: [number, Word.ContentControl], selected: string[]) {
    const exclude = (title: string) => `!${title}`;
    if (selected.find(t=>t.includes(ctrl.title))) return;//!In some cases, ctrl.contentControl.items returns not only the child contentcontrols of ctrl, but includes also the parent contentcontrol of ctrl. Don't understand why this happens.
    
    ctrl.select();
    const [container, btnNext, checkBox] = await showUI();
    
    return new Promise((resolve, reject) => {
        btnNext.onclick = ()=>nextCtrl(ctrl, checkBox as HTMLInputElement);
        async function nextCtrl(ctrl:Word.ContentControl, checkBox:HTMLInputElement) {
            const checked = checkBox.checked;
            container.remove();
            ctrl.contentControls.load(['title', 'tag']);
            await ctrl.context.sync();
            const subOptions =
                ctrl.contentControls.items
                    .filter(ctrl => OPTIONS.includes(ctrl.tag));
            if (checked)
                await isSelected(ctrl, subOptions);
            else isNotSelected(ctrl, subOptions);
            resolve(selected);
         }; 
    });
     
    async function isSelected(ctrl: Word.ContentControl, subOptions:Word.ContentControl[]) {
        selected.push(ctrl.title);
        const entries = subOptions.entries()
        for (const entry of entries) {
            await promptForSelection(entry, selected);
        }
        console.log(selected);
    };


    function isNotSelected(ctrl: Word.ContentControl, subTitles:Word.ContentControl[]) {
        selected.push(exclude(ctrl.title));
        subTitles
            .forEach(ctrl=> selected.push(exclude(ctrl.title)));
        console.log(selected)
    };


    async function showUI() {
        const children = ctrl.contentControls;
        children.load(['title', 'tag']);
        await ctrl.context.sync();
        const RTSi = children.items.find(rt => rt.tag === RTSiTag);
        if (!RTSi) throw new Error('No RTSi');
        const ctrlRange = RTSi.getRange('Content');
        ctrlRange.load(['text', 'paragraphs']);
        await ctrl.context.sync();
        return UI(ctrlRange.text);
       
    
        function UI(text:string) {
            const container = createHTMLElement('div', 'promptContainer', '', USERFORM, ctrl.title);
                const prompt = createHTMLElement('div', 'selection', '', container);
                const checkBox = createHTMLElement('input', 'checkBox', '', prompt) as HTMLInputElement;
                createHTMLElement('label', 'label', text, prompt) as HTMLParagraphElement;
                checkBox.type = 'checkbox';
                const btns = createHTMLElement('div', 'btns', '', prompt);
                const btnNext = createHTMLElement('button', 'btnOK', 'Next', btns);
                return [container, btnNext, checkBox]
        }

    }
}

async function deleteAllNotSelected(selected: string[], document:Word.Document | Word.DocumentCreated) {
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

function createHTMLElement(tag: string, css: string, innerText:string, parent: HTMLElement | Document, id?:string, append:boolean = true) {
    const el = document.createElement(tag);
    if (innerText) el.innerText = innerText;
    el.classList.add(css);
    if (id) el.id = id;
    append ? parent.appendChild(el) : parent.prepend(el);
    return el
}
