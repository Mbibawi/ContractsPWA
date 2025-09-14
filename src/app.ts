const OPTIONS = ['RTSelect', 'RTShow', 'RTEdit'],
    RTDropDownTag = 'RTList',
    RTDropDownColor = '#991c63',
    RTDuplicateTag = 'RTRepeat',
    RTSectionTag = 'RTSection',
    RTSelectTag = 'RTSelect',
    RTOrTag = 'RTOr',
    RTObsTag = 'RTObs',
    RTDescriptionTag = 'RTDesc',
    RTDescriptionStyle = 'RTDescription',
    RTSiTag = 'RTSi',
    RTSiStyles = ['RTSi0cm', 'RTSi1cm', 'RTSi2cm', 'RTSi3cm', 'RTSi4cm'];

let USERFORM: HTMLDivElement, NOTIFICATION: HTMLDivElement;
let RichText: ContentControlType,
    RichTextInline: ContentControlType,
    RichTextParag: ContentControlType,
    ComboBox: ContentControlType,
    CheckBox: ContentControlType,
    dropDownList: ContentControlType,
    Bounding: Word.ContentControlAppearance,
    Hidden: Word.ContentControlAppearance


Office.onReady((info) => {
    // Check that we loaded into Word
    if (info.host !== Office.HostType.Word) return showNotification('This addin is designed to work on Word only');

    USERFORM = document.getElementById('userFormSection') as HTMLDivElement;
    NOTIFICATION = document.getElementById('notification') as HTMLDivElement;
    RichText = Word.ContentControlType.richText;
    RichTextInline = Word.ContentControlType.richTextInline;
    RichTextParag = Word.ContentControlType.richTextParagraphs;
    ComboBox = Word.ContentControlType.comboBox;
    CheckBox = Word.ContentControlType.checkBox;
    dropDownList = Word.ContentControlType.dropDownList;
    Bounding = Word.ContentControlAppearance.boundingBox;
    Hidden = Word.ContentControlAppearance.hidden;

    mainUI();
});

function showBtns(btns: Btn[], append = true) {
    return btns.map(btn => insertBtn(btn, append))
};

function mainUI() {
    if (!USERFORM) return;
    USERFORM.innerHTML = '';
    const main: Btn[] =
        [[customizeContract, 'Customize Contract'], [prepareTemplate, 'Prepare Template']];
    const btns = showBtns(main);
    const back = [goBack, 'Go Back'] as Btn;
    btns.forEach(btn => btn?.addEventListener('click', () => insertBtn(back, false)));
    function goBack() {
        USERFORM.innerHTML = '';
        showBtns(main)
    }
}


function prepareTemplate() {
    USERFORM.innerHTML = '';
    function wrap(title: string, tag: string, type: Word.ContentControlType, style: string | null, cannotEdit: boolean, cannotDelete: boolean, label: string) {
        return [
            () => wrapSelectionWithContentControl(title, tag, type, style, cannotEdit, cannotDelete),
            label
        ] as [Function, string]
    };


    const btns = [
        wrap(RTSiTag, RTSiTag, RichText, RTSiStyles[0], true, true, 'Insert Single RT Si'),
        [() => insertRTDescription(true), 'Insert Single RT Description'],
        wrap(RTSelectTag, RTSelectTag, RichText, null, false, true, 'Insert Single RT Select'),
        wrap(RTSectionTag, RTSectionTag, RichText, RTSectionTag, true, true, 'Insert Single RT Section'),
        wrap(RTOrTag, RTOrTag, RichText, null, false, true, 'Insert Single RT OR'),
        wrap(RTDuplicateTag, RTDuplicateTag, RichText, null, false, true, 'Insert Single RT Dublicate Block'),
        [insertDropDownList, 'Insert a Dropdown List from selection'],
        wrap(RTObsTag, RTObsTag, RichText, RTObsTag, true, true, 'Insert Single RT Obs'),
        [insertRTSiAll, 'Insert RT Si For All'],
        [insertRTDescription, 'Insert RT Description For All'],
    ] as [Function, string][];

    showBtns(btns);
}

function insertBtn([fun, label]: Btn, append: boolean = true) {
    if (!USERFORM) return;
    const htmlBtn = document.createElement('button');
    append ? USERFORM.appendChild(htmlBtn) : USERFORM.prepend(htmlBtn);
    htmlBtn.innerText = label;
    htmlBtn.addEventListener('click', () => fun());
    return htmlBtn
}

/**
 * Wraps every occurrence of text formatted with a specific character style in a rich text content control.
 *
 * @param style The name of the character style to find (e.g., "Emphasis", "Strong", "MyCustomStyle").
 * @returns A Promise that resolves when the operation is complete.
 */
async function findTextAndWrapItWithContentControl(styles: string[], title: string, tag: string, cannotEdit: boolean, cannotDelete: boolean) {
    const separator = '_&_'
    const search = (await promptForInput(`Provide the search string. You can provide more than one string to search by separated by ${separator} witohout space`, separator))?.split(separator) as string[];

    if (!search?.length) return showNotification('The provided search string is not valid');

    const matchWildcards = await promptConfirm('Match Wild Cards');

    if (!styles) styles = (await promptForInput(`Provide the styles that that need to be matched separated by ","`))?.split(',') || [];

    if (!styles?.length) return showNotification(`The styles[] has 0 length, no styles are included, the function will return`);

    const all: (Word.ContentControl | undefined)[][] = [];
    return await Word.run(async (context) => {
        for (const el of search) {
            const ranges = await searchString(el, context, matchWildcards);
            if (!ranges) continue;
            const ctrls = await wrapMatchingStyleRangesWithContentControls(ranges, styles, title, tag, cannotEdit, cannotDelete);
            if (!ctrls) continue
            all.push(ctrls)
        };
        return all.flat();
    });
}

async function wrapMatchingStyleRangesWithContentControls(ranges: Word.RangeCollection, styles: string[], title: string, tag: string, cannotEdit: boolean, cannotDelete: boolean) {

    ranges.load(['style', 'text', 'parentContentControlOrNullObject', 'parentContentControlOrNullObject.isNullObject', 'parentContentControlOrNullObject.tag']);

    await ranges.context.sync();
    if (!ranges.items.length) {
        showNotification(`No text matching the search string was found in the document.`);
        return;
    }

    showNotification(`Found ${ranges.items.length} ranges matching the search string. First range text = ${ranges.items[0].text}`);

    const ctrls = ranges.items.map(async (range, index) => {
        if (!styles.includes(range.style)) return;
        if (range.parentContentControlOrNullObject?.tag === tag) return;
        return await insertContentControl(range, title, tag, index, RichText, range.style, cannotEdit, cannotDelete)
    });
    return Promise.all(ctrls);
}

async function searchString(search: string, context: Word.RequestContext, matchWildcards: boolean) {
    const searchResults = context.document.body.search(search, { matchWildcards: matchWildcards });
    searchResults.load(['style']);
    await context.sync();
    return searchResults
}

async function addIDtoCtrlTitle(ctrls: Word.ContentControlCollection) {
    ctrls.load(['title', 'id']);
    await ctrls.context.sync();
    ctrls.items
        .filter(ctrl => !ctrl.title.endsWith(`-${ctrl.id}`))
        .forEach(ctrl => ctrl.title = getCtrlTitle(ctrl.tag, ctrl.id));

    await ctrls.context.sync();
}
async function insertRTDescription(selection: boolean = false, style: string = 'Normal') {
    let ctrls: (ContentControl | undefined)[] | void;
    if (selection) {
        const range = await getSelectionRange();
        if (!range) return showNotification('No Text Was selected !')
        ctrls = [await insertContentControl(range, RTDescriptionTag, RTDescriptionTag, 0, RichText, RTDescriptionStyle, true, true)];
    }
    else ctrls = await findTextAndWrapItWithContentControl([RTDescriptionStyle], RTDescriptionTag, RTDescriptionTag, true, true);

    if (!ctrls?.length) return;

    for (const ctrl of ctrls) {
        if (!ctrl) continue;
        const range = ctrl.getRange();
        const inserted = range.insertText('[*]\u00A0', Word.InsertLocation.before);
        inserted.style = style;
        inserted.font.bold = true;
    }
    await ctrls[0]?.context.sync();

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
            const style = RTSiStyles.includes(parag.style) ? parag.style : RTSiStyles[0];
            try {
                const parent = parag.parentContentControlOrNullObject;
                parent.load(['tag']);
                await parag.context.sync();
                if (parent.tag === RTSiTag) continue;
                showNotification(`range style: ${parag.style} & text = ${parag.text}`);
                await insertContentControl(parag.getRange('Content'),
                    RTSiTag, RTSiTag, parags.indexOf(parag), RichText, style);
            } catch (error) {
                showNotification(`error: ${error}`);
                continue
            }
        }
        await context.sync();

    })
}
async function insertDropDownList() {
    const range = await getSelectionRange();
    if (!range) return;
    range.load(["text"]);
    range.context.trackedObjects.add(range);//!This is important
    await range.context.sync();

    const options = range.text.split("/");
    if (!options.length) return showNotification("No options");
    showNotification(options.join());

    const ctrl = await insertContentControl(range, RTDropDownTag, RTDropDownTag, 0, dropDownList, null, false, true);
    if (!ctrl) return;
    ctrl.dropDownListContentControl.deleteAllListItems();
    options.forEach(option => ctrl.dropDownListContentControl.addListItem(option));
    setCtrlsFontColor([ctrl], RTDropDownColor);
    setCtrlsColor([ctrl], RTDropDownColor);
    await ctrl.context.sync();
}
async function insertContentControl(range: Word.Range, title: string, tag: string, index: number, type: Word.ContentControlType, style: string | null, cannotEdit: boolean = true, cannotDelete: boolean = true) {
    range.select();
    const styles = range.context.document.getStyles();
    styles.load(['nameLocal', 'type']);
    // Insert a rich text content control around the found range.
    //@ts-expect-error
    const ctrl = range.insertContentControl(type);
    ctrl.load(["id"]);
    range.context.trackedObjects.remove(range);
    range.context.trackedObjects.add(ctrl);//!This is very important otherwise we will not be able to call range.context.sync() after calling range.context.sync();
    await range.context.sync();
    // Set properties for the new content control.
    if (ctrl.id) showNotification(`the newly created ContentControl id = ${ctrl.id} `);
    try {
        ctrl.select();
        ctrl.title = getCtrlTitle(title, ctrl.id);
        ctrl.tag = tag;
        ctrl.appearance = Word.ContentControlAppearance.boundingBox;
        const foundStyle = styles.items.find(s => s.nameLocal === style);
        if (style && foundStyle?.type === Word.StyleType.character)
            ctrl.style = style;
        if (style) ctrl.getRange().style = style;
        ctrl.cannotDelete = cannotDelete;
        ctrl.cannotEdit = cannotEdit;//!This must come at the end after the style has been set.
        await range.context.sync();
        showNotification(`Wrapped text in range ${index || 1} with a content control.`);
    } catch (error) {
        showNotification(`There was an error while setting the properties of the newly crated contentcontrol by insertContentControl(): ${error}.`);
    }
    return ctrl;
}

async function getSelectionRange() {
    return await Word.run(async (context) => {
        const range = context.document
            .getSelection()
            .getRange('Content');
        range.load(['style', 'isEmpty']);
        context.trackedObjects.add(range);
        await context.sync();
        if (range.isEmpty) return showNotification('The selection range is empty');
        return range
    });
}

async function wrapSelectionWithContentControl(title: string, tag: string, type: Word.ContentControlType, style: string | null, cannotEdit: boolean, cannotDelete: boolean) {
    const range = await getSelectionRange();
    if (!range) return;
    if (RTSiStyles.includes(range.style)) style = range.style;
    await insertContentControl(range, title, tag, 0, type, style, cannotEdit, cannotDelete);
};

async function promptConfirm(question: string, fun?: Function): Promise<boolean> {
    if (!question) question = 'No question was provided !!!';
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

    function confirm(confirm: boolean) {
        container.remove();
        if (fun) fun(confirm);
        return confirm;
    }

};



async function customizeContract() {
    USERFORM.innerHTML = '';
    const processed = (id: number) => selected.find(t => t.includes(id.toString()));
    const TAGS = [...OPTIONS, RTDuplicateTag];
    const props = ['id', 'tag', 'title'];
    const getSelectCtrls = (ctrls: ContentControl[]) => ctrls.filter(ctrl => TAGS.includes(ctrl.tag));
    const selected: string[] = [];
    await loopSelectCtrls();

    async function loopSelectCtrls() {
        await Word.run(async (context) => {
            const allRT = context.document.getContentControls();
            allRT.load(props);
            await context.sync();
            const selectCtrls = getSelectCtrls(allRT.items);
            for (const ctrl of selectCtrls)
                await promptForSelection(ctrl);

            const keep = selected.filter(title => !title.startsWith('!'));
            showNotification(`keep = ${keep.join(', ')}`);
            try {
                await currentDoc();
                //await createNewDoc();
            } catch (error) {
                showNotification(`${error}`)
            }

            async function currentDoc() {
                const allRT = context.document.getContentControls();
                allRT.load(props);
                await context.sync();
                const selectCtrls = getSelectCtrls(allRT.items);//!We need to retrieve all the selected items again because we may have added new ctrls by cloning the 'Duplicate' ctrls
                for (const ctrl of selectCtrls) {
                    if (keep.includes(ctrl.title)) continue;
                    ctrl.select();
                    ctrl.cannotDelete = false;
                    showNotification(`Deleted Ctrl: ${ctrl.title}`)
                    ctrl.delete(false);
                }
                await context.sync();
            };

            async function createNewDoc() {
                return;//!Desactivating working with new document created from template until we find a solution to the context issue
                const template = await getTemplate() as Base64URLString;
                console.log(template);
                if (!template) return showNotification('Failed to create the template');
                const newDoc = context.application.createDocument(template);
                const all = newDoc.contentControls;
                all.load(['title', 'tag']);
                await newDoc.context.sync();

                showNotification(`All ctrls from newDoc = : ${all.items.map(c => c.title).join(', ')}`);


                all.items.map(ctrl => {
                    if (keep.includes(ctrl.title)) return;
                    ctrl.cannotDelete = false;
                    ctrl.delete(false);
                });
                await newDoc.context.sync()
                newDoc.open();
            }
        });
    }

    async function promptForSelection(ctrl: ContentControl) {
        try {
            ctrl.select();
            await showSelectPrompt([ctrl]);
        } catch (error) {
            showNotification(`Error from promptForSelection() = ${error}`)
        }
    }

    async function showSelectPrompt(selectCtrls: ContentControl[]) {
        const blocks: (selectBlock | undefined)[] = [];
        try {
            for (const ctrl of selectCtrls) {
                if (processed(ctrl.id)) continue;//!We must escape the ctrls that have already been processed
                if (ctrl.tag === RTDuplicateTag) {
                    await duplicateBlock(ctrl.id);
                    continue
                };
                const addBtn = selectCtrls.indexOf(ctrl) + 1 === selectCtrls.length;
                blocks.push(await insertPromptBlock(ctrl.id, addBtn) || undefined);
            }
            await btnOnClick(blocks)
        } catch (error) {
            return showNotification(`Error from showSelectPrompt() = ${error}`)
        }
    }

    async function insertPromptBlock(id: number, addBtn: boolean): Promise<selectBlock | void> {
        try {
            return await wordRun();
        } catch (error) {
            return showNotification(`Error from insertPromptBlock() = ${error}`)
        }

        async function wordRun() {
            return await Word.run(async (context) => {
                const ctrl = context.document.contentControls.getById(id);
                ctrl.load(props);
                const label = labelRange(ctrl, RTSiTag);
                label.select();
                await context.sync();
                const text = label.text;
                // showNotification(`CtrlSi.text = ${text}`);
                label.font.hidden = true;
                await context.sync();
                return { ctrl, ...appendHTMLElements(text, ctrl.title, addBtn) } as selectBlock;//The checkBox will have as id the title of the "select" contentcontrol}
            });
        }
    }

    function appendHTMLElements(text: string, id: string, addBtn: boolean = false): selectBlock {
        const container = createHTMLElement('div', 'promptContainer', '', USERFORM) as HTMLDivElement;
        const checkBox = createHTMLElement('input', 'checkBox', '', container, id) as HTMLInputElement;
        createHTMLElement('label', 'label', text, container) as HTMLParagraphElement;
        checkBox.type = 'checkbox';
        if (!addBtn) return { container, checkBox };
        const btns = createHTMLElement('div', 'btns', '', container);
        const btnNext = createHTMLElement('button', 'btnOK', 'Next', btns) as HTMLButtonElement;
        return { container, checkBox, btnNext }
    }

    function btnOnClick(blocks: (selectBlock | undefined)[]): Promise<string[]> {
        return new Promise((resolve, reject) => {
            const btn = blocks.find(container => container?.btnNext)?.btnNext;
            !btn ? resolve(selected) : btn.onclick = processBlocks;
            async function processBlocks() {
                const values: [ContentControl | undefined, boolean][] =
                    blocks
                        .filter(block => block?.ctrl && block.checkBox)
                        //@ts-ignore
                        .map(block => [block.ctrl, block.checkBox.checked]);
                blocks.forEach(block => block?.container.remove());//We start by removing all the containers
                for (const [ctrl, checked] of values) {
                    if (!ctrl) continue;
                    const subOptions = await getSubOptions(ctrl.id, checked);
                    if (checked)
                        await isSelected(ctrl.id, subOptions);
                    else isNotSelected(ctrl.id, subOptions);
                }
                resolve(selected);
            };
        });
    }

    async function isSelected(id: number, subOptions: ContentControl[] | undefined) {
        selected.push(`${id}`);
        if (subOptions) await showSelectPrompt(subOptions);
    };

    function isNotSelected(id: number, subOptions: ContentControl[]) {
        const exclude = (id: number) => `!${id}`;
        selected.push(exclude(id));
        subOptions
            .forEach(ctrl => selected.push(exclude(ctrl.id)));
        console.log(selected)
    };

    async function getSubOptions(id: number, directChildren: boolean, children?: ContentControl[]) {
        if (!children) children = await getChildren();
        if (!directChildren) return getSelectCtrls(children);
        return getSelectCtrls(children).filter(c => c.parentContentControl.id === id);//!We need to make sure we get only the direct children of the ctrl and not all the nested ctrls
        async function getChildren() {
            return Word.run(async (context) => {
                const ctrl = context.document.getContentControls().getById(id);
                const children = ctrl.getContentControls();
                children.load([...props, 'parentContentControl'])
                await context.sync();
                return children.items.filter(c => c.id !== id);//!We must exclude the ctrl itself which in some cases may be returned as part of its children due to a bug in Word API
            })
        }
    }


    async function duplicateBlock(id: number) {
        const replace = Word.InsertLocation.replace;
        const after = Word.InsertLocation.after;
        try {
            await insertClones();
        } catch (error) {
            showNotification(`${error}`)
        }

        async function insertClones() {
            await Word.run(async (context) => {
                const ctrl = context.document.contentControls.getById(id);
                ctrl.load(props);
                const label = labelRange(ctrl, RTSectionTag);
                await context.sync();
                if (!label.text) return showNotification("No lable text");
                ctrl.select();
                const message = `How many ${label.text} parties are there?`;
                const answer = Number(await promptForInput(message));
                if (isNaN(answer))
                    return showNotification(`The provided text cannot be converted into a number: ${answer}`);
                const title = `${getCtrlTitle(ctrl.tag, id)}-Cloned ${answer}`
                ctrl.title = title;//!We must update the title in case it is no matching the id in the template.
                const ctrlContent = ctrl.getOoxml();
                label.font.hidden = true;
                await context.sync();
                for (let i = 1; i < answer; i++)
                    ctrl.getRange().insertOoxml(ctrlContent.value, after);
                const clones = ctrl.context.document.getContentControls().getByTitle(title);
                clones.load(props);
                label.font.hidden = true;
                await context.sync();
                const items = clones.items;//!clones.items.entries() caused the for loop to fail in scriptLab. The reason is unknown
                try {
                    for (const clone of items)
                        await processClone(clone.id, items.indexOf(clone) + 1);
                } catch (error) {
                    showNotification(`Error from processClone() = ${error}`)
                }
            });
        }

        async function processClone(id: number, i: number) {
            await Word.run(async (context) => {
                const clone = context.document.contentControls.getById(id);
                clone.load(props);
                const label = labelRange(clone, RTSectionTag);
                const children = clone.getContentControls();
                children.load(props);
                await context.sync();
                clone.title = `${getCtrlTitle(clone.tag, clone.id)}-${i}`;
                const text = `${label.text} ${i}`;
                label.insertText(text, replace);
                label.font.hidden = true;
                await context.sync();
                const subOptions = await getSubOptions(clone.id, true);//!We select only the direct select ctrls children
                const div = createHTMLElement('div', '', text, USERFORM, '', false);
                await showSelectPrompt(subOptions);
                div.remove();
            });

        };

    }

    function labelRange(parent: ContentControl, tag: string) {
        const ctrl = getFirstByTag(parent, tag)
        const range = ctrl.getRange('Content');
        ctrl.cannotEdit = false;
        range.font.hidden = false;
        range.load(['text']);
        return range;
    }

    function getFileURL() {
        let url;
        Office.context.document.getFilePropertiesAsync(undefined, (result) => {
            if (result.error) return;
            url = result.value.url;
        });
        return url
    }

    async function getTemplate() {
        try {
            return await getDocumentBase64();
        } catch (error) {
            showNotification(`Failed to create new Doc: ${error}`)
        }
    }
};

async function promptForInput(question: string, deflt?: string, fun?: Function): Promise<string | void> {
    if (!question) return '';
    const container = createHTMLElement('div', 'promptContainer', '', USERFORM);
    const prompt = createHTMLElement('div', 'prompt', '', container);
    const ask = createHTMLElement('p', 'ask', question, prompt);
    const input = createHTMLElement('input', 'answer', '', prompt) as HTMLInputElement;
    const btns = createHTMLElement('div', 'btns', '', prompt);
    const btnOK = createHTMLElement('button', 'btnOK', 'OK', btns);
    const btnCancel = createHTMLElement('button', 'btnCancel', 'Cancel', btns);
    if (deflt) input.value = deflt;
    return new Promise((resolve, reject) => {
        btnCancel.onclick = () => reject(container.remove());
        btnOK.onclick = () => {
            const answer = input.value;
            console.log('user answer = ', answer);
            container.remove();
            if (fun) fun(answer);
            resolve(answer)
        };
    });

};

function getCtrlTitle(tag: string, id: number) {
    return `${tag}&${id}`
}

function getFirstByTag(range: Word.Range | ContentControl, tag: string) {
    return range.getContentControls().getByTag(tag).getFirst();
}

/**
 * Asynchronously gets the entire document content as a Base64 string.
 * This function handles multi-slice documents by requesting each slice in parallel.
 * @returns A Promise that resolves with the Base64-encoded document content.
 */
async function getDocumentBase64(): Promise<Base64URLString> {
    const failed = (result: Office.AsyncResult<Office.File | Office.Slice>) => result.status !== Office.AsyncResultStatus.Succeeded;
    const sliceSize = 16 * 1024;//!We need not to exceed the Maximum call stack limit when the slices will be passed to String.FromCharCode()
    return new Promise((resolve, reject) => {
        // Step 1: Request the document as a compressed file.
        Office.context.document.getFileAsync(
            Office.FileType.Compressed,
            { sliceSize: sliceSize },
            (fileResult) => processFile(fileResult)
        );

        function processFile(fileResult: Office.AsyncResult<Office.File>) {
            if (failed(fileResult))
                return reject(fileResult.error);

            const file = fileResult.value;
            const sliceCount = file.sliceCount;
            const slices: number[][] = [];

            getSlice();

            function getSlice() {
                file.getSliceAsync(slices.length, (sliceResult) => processSlice(sliceResult));
            }

            function processSlice(sliceResult: Office.AsyncResult<Office.Slice>) {
                try {
                    if (failed(sliceResult))
                        return file.closeAsync(() => reject(sliceResult.error));

                    slices.push(sliceResult.value.data);

                    if (slices.length < sliceCount) return getSlice();

                    const binaryString = slices.map(slice => String.fromCharCode(...slice)).join('');
                    const base64String = btoa(binaryString);
                    file.closeAsync(() => resolve(base64String));

                } catch (error) {
                    showNotification(`${error}, succeeded = ${sliceResult.status}, loaded = ${slices.length}`)
                }


            }
        }
    });
}


async function deleteAllNotSelected(selected: string[], wdDoc: Word.Document | Word.DocumentCreated) {
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

function createHTMLElement(tag: string, css: string, innerText: string, parent?: HTMLElement | Document, id?: string, append: boolean = true) {
    const el = document.createElement(tag);
    if (innerText) el.innerText = innerText;
    if (css) el.classList.add(css);
    if (id) el.id = id;
    if (!parent) return el;
    append ? parent.appendChild(el) : parent.prepend(el);
    return el
}

function showNotification(message: string, clear: boolean = false) {
    if (clear) NOTIFICATION.innerHTML = '';
    createHTMLElement('p', 'notification', message, NOTIFICATION, '', true);
}

async function setCanBeEditedForAllSelectCtrls(edit: boolean = true) {
    await Word.run(async (context) => {
        const ctrls = context.document
            .contentControls;
        ctrls.load(['title', 'tag']);
        await context.sync();
        ctrls.items.forEach(ctrl => {
            if (OPTIONS.indexOf(ctrl.tag) > -1)
                ctrl.cannotEdit = edit;
        });
        await context.sync();
    });
}

function deleteCtrlById() {
    Word.run(async (context) => {
        let title = await promptForInput('Provide the title or the id of the control. If you provide the title, the id will be extracted from it');
        if (!title) return showNotification(`You did not provide a valid id or title: ${title}`);
        if (title.includes('&')) title = title.split('&')[1];
        const id = Number(title);
        if (isNaN(id)) return showNotification(`The id could not be extracted from the title: ${title}`);
        const ctrl = context.document.contentControls.getById(id);
        const ctrls = ctrl.getContentControls();
        ctrls.load(['tag', 'id']);
        await context.sync();
        console.log('Ctrls = ', ctrls.items.map(c => c.tag));
        ctrl.cannotDelete = false;
        ctrls.items.forEach(ctrl => ctrl.cannotDelete = false);
        ctrl.delete(false);
        await context.sync();
    });
}

function updateAllCtrlsTitles() {
    Word.run(async (context) => {
        const ctrls = context.document.getContentControls();
        ctrls.load(['title', 'tag', 'id']);
        await context.sync();
        ctrls.items
            .filter(ctrl => ctrl.tag)
            .forEach(ctrl => ctrl.title = getCtrlTitle(ctrl.tag, ctrl.id));
        await context.sync();
    });
}
async function selectAllCtrlsByTag(tag: string, color: string) {
    Word.run(async (context) => {
        const ctrls = context.document.getContentControls();
        ctrls.load(['tag']);
        await context.sync();
        const sameTag = ctrls.items.filter(c => c.tag === tag);
        await setCtrlsColor(sameTag, color);
        await setCtrlsFontColor(sameTag, color);
        await context.sync();
    });
}
async function setCtrlsColor(ctrls: ContentControl[], color: string) {
    ctrls.forEach(ctrl => ctrl.color = color);
}
function setCtrlsFontColor(ctrls: ContentControl[], color: string) {
    ctrls.forEach(c => c.getRange().font.color = color);
}

function setRangeStyle(objs: (ContentControl | Word.Paragraph)[], style: string) {
    objs.forEach(o => o.getRange().style = style);
}

async function finalizeContract() {
    const tags = [RTSiTag, RTDescriptionTag, RTObsTag, RTDropDownTag];
    await removeRTs(tags);
}

async function removeRTs(tags: string[]) {
    Word.run(async (context) => {
        for (const tag of tags) {
            const ctrls = context.document.getContentControls().getByTag(tag);
            ctrls.load('tag');
            await context.sync();
            ctrls.items.forEach(ctrl => {
                ctrl.select();
                ctrl.cannotDelete = false;
                ctrl.delete(tag === RTDropDownTag)
            });
            await context.sync();
        }
    });
}

async function changeAllSameTagCtrlsCannEdit(tag: string, edit: boolean) {
    Word.run(async (context) => {
        const ctrls = context.document.getContentControls().getByTag(tag);
        ctrls.load('tag');
        await context.sync();
        ctrls.items.forEach(ctrl => ctrl.cannotEdit = edit);
        await context.sync()
    });
}