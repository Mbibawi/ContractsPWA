const OPTIONS = ['RTSelect', 'RTShow', 'RTEdit'],
    StylePrefix = 'Contrat_',
    RTFieldTag = 'RTField',
    RTDropDownTag = 'RTList',
    RTDropDownColor = '#991c63',
    RTDuplicateTag = 'RTRepeat',
    RTSectionTag = 'RTSection',
    RTSectionStyle = `${StylePrefix}${RTSectionTag}`,
    RTSelectTag = 'RTSelect',
    RTOrTag = 'RTOr',
    RTObsTag = 'RTObs',
    RTObsStyle = `${StylePrefix}${RTObsTag}`,
    RTDescriptionTag = 'RTDesc',
    RTDescriptionStyle = `${StylePrefix}${RTDescriptionTag}`,
    RTSiTag = 'RTSi',
    RTSiStyles = ['0', '1', '2', '3', '4'].map(n => `${StylePrefix}${RTSiTag}${n}cm`);
const version = "v10.7";

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

function showBtns(btns: Btn[], append = true, on: string = 'click') {
    return btns.map(btn => insertBtn(btn, append, on))
};

function mainUI() {
    if (!USERFORM) return;
    const p = document.createElement('p');
    p.innerText = version;
    USERFORM.insertAdjacentElement('beforebegin', p);
    USERFORM.innerHTML = '';
    const main: Btn[] =
        [[customizeContract, 'Customize Contract'], [prepareTemplate, 'Prepare Template'], [finalizeContract, 'Finalize Contract']];
    const btns = showBtns(main);
    const back = [goBack, 'Go Back'] as Btn;
    btns.forEach(btn => btn?.addEventListener('click', () => insertBtn(back, false)));
    function goBack() {
        USERFORM.innerHTML = '';
        document.getElementById('stylesList')?.remove();
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
        [insertDroDownListAll, 'Insert DropDown List For All Matches'],
        [insertRTSiAll, 'Insert RT Si For All'],
        [insertRTSectionAll, 'Insert RT Section For All'],
        [insertRTDescription, 'Insert RT Description For All'],
    ] as [Function, string][];

    showBtns(btns);
    showBtns([[() => customizeContract(true), 'Show Nested Options Tree']], true);
    showStylesList();

    function showStylesList() {
        const id = 'stylesList';
        if (document.getElementById(id)) return;
        Word.run(async (context) => {
            const allStyles = context.document.getStyles();
            allStyles.load(['nameLocal']);
            await context.sync();
            const styles = allStyles.items.filter(style => style.nameLocal.startsWith(StylePrefix));
            if (!styles.length) return;
            document.createElement('select')
            const container = createHTMLElement('div', '', '', undefined, id) as HTMLDivElement;
            USERFORM.insertAdjacentElement('beforebegin', container);
            const select = createHTMLElement('select', '', '', container) as HTMLSelectElement;
            styles.forEach(style => {
                const option = createHTMLElement('option', '', style.nameLocal.split(StylePrefix)[1], select) as HTMLOptionElement;
                option.value = style.nameLocal;
            });

            select.onmouseenter = async () => {
                const range = await getSelectionRange();
                if (!range) return;
                const value = Array.from(select.options).find(o => o.value === range?.style)?.value || range.style;
                if (value) select.value = value;
                range.untrack();
            }

            select.onchange = async () => {
                const range = await getSelectionRange();
                if (!range) return;
                range.style = select.value;
                range.untrack();
                await range.context.sync();
            }
        })

    }
}

function insertBtn([fun, label]: Btn, append: boolean = true, on: string = 'click') {
    if (!USERFORM) return;
    const htmlBtn = document.createElement('button');
    append ? USERFORM.appendChild(htmlBtn) : USERFORM.prepend(htmlBtn);
    htmlBtn.innerText = label;
    htmlBtn.addEventListener(on, () => fun());
    return htmlBtn
}

/**
 * Wraps every occurrence of text formatted with a specific character style in a rich text content control.
 *
 * @param style The name of the character style to find (e.g., "Emphasis", "Strong", "MyCustomStyle").
 * @returns A Promise that resolves when the operation is complete.
 */
async function findTextAndWrapItWithContentControl(styles: string[], title: string, tag: string, cannotEdit: boolean, cannotDelete: boolean) {

    return await Word.run(async (context) => {
        const ctrls: ContentControl[] = [];
        const { search, matchWildcards } = await searchs();
        if (!search?.length) return showNotification('The provided search string is not valid');
        if (!styles?.length) return showNotification(`The styles[] has 0 length, no styles are included, the function will return`);

        for (const find of search) {
            const matches = await searchString(find, context, matchWildcards);
            if (!matches?.items.length) continue;
            matches.load(['style', 'text', 'parentContentControlOrNullObject']);
            await context.sync();
            const ranges = matches.items.filter(range => styles.includes(range.style));
            showNotification(`Found ${ranges.length} ranges matching the search string. First range text = ${ranges[0].text}`);
            ctrls.push(...await insertCtrls(ranges))
        };

        return ctrls;

        async function insertCtrls(ranges: Word.Range[]) {
            const ctrls: ContentControl[] = [];
            for (const range of ranges) {
                const parent = range.parentContentControlOrNullObject;
                parent.load('tag');
                await context.sync();
                if (parent.tag === tag) continue;
                try {
                    const ctrl = await insertContentControl(range, title, tag, ranges.indexOf(range), RichText, range.style, cannotEdit, cannotDelete);
                    if (ctrl) ctrls.push(ctrl);
                } catch (error) {
                    showNotification(`Error from insertCtrls() while inserting the contentControl() in the matching range. Error = ${error}`)
                }
            }
            return ctrls
        }
    });

    async function searchs() {
        const separator = '_&_';
        const search = (await promptForInput(`Provide the search string. You can provide more than one string to search by separated by ${separator} witohout space`, separator))?.split(separator) as string[];
        const matchWildcards = await promptConfirm('Match Wild Cards');

        if (!styles) styles = (await promptForInput(`Provide the styles that that need to be matched separated by ","`))?.split(',') || [];
        return { search, matchWildcards }
    }
}

async function searchString(search: string, context: Word.RequestContext, matchWildcards: boolean, replaceWith?: string): Promise<Word.RangeCollection> {
    const searchResults = context.document.body.search(search, { matchWildcards: matchWildcards });
    searchResults.load(['style', 'text']);
    searchResults.track();
    await context.sync();
    if (!replaceWith) return searchResults;
    for (const match of searchResults.items)
        match.insertText(replaceWith, Word.InsertLocation.replace);
    await context.sync();
    return await searchString(replaceWith, context, false)
}

async function insertRTDescription(selection: boolean = false, style: string = `${StylePrefix}Normal`) {
    NOTIFICATION.innerHTML = '';
    let ctrls: (ContentControl | undefined)[] | void;
    if (selection) {
        const range = await getSelectionRange();
        if (!range) return showNotification('No Text Was selected !');
        ctrls = [await insertContentControl(range, RTDescriptionTag, RTDescriptionTag, 0, RichText, RTDescriptionStyle, true, true)];
    }
    else ctrls = await findTextAndWrapItWithContentControl([RTDescriptionStyle], RTDescriptionTag, RTDescriptionTag, true, true);

    if (!ctrls?.length) return;
    const ids = ctrls.map(c => c?.id || 0);

    await insertFieldCtrl(ids, style);

}
async function insertFieldCtrl(ids: number[], style: string) {
    style = '';
    await Word.run(async (context) => {
        for (const id of ids) {
            if (!id) continue;
            const ctrl = context.document.getContentControls().getById(id);
            await context.sync();
            try {
                await insert(ctrl);
            } catch (error) {
                showNotification(`Error inserting field: ctrl.id = ${ctrl?.id}, error: ${error}`);
                continue
            }
        }
        await context.sync();
    });

    async function insert(ctrl: ContentControl) {
        const start = ctrl.getRange(Word.RangeLocation.before);
        const field = await insertContentControl(start, RTFieldTag, RTFieldTag, 0, RichText, style, false, false, '[*]');
        if (!field) return;
        field.getRange('Content').font.bold = true
    }
}
function insertRTSiAll() {
    insertForAllParags(RTSiStyles, RTSiTag)
}
function insertRTSectionAll() {
    insertForAllParags([RTSectionStyle], RTSectionTag)
}
async function insertForAllParags(Styles: string[], tag: string) {
    NOTIFICATION.innerHTML = '';
    await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load(['style', 'text', 'parentContentControlOrNullObject']);
        await context.sync();
        const parags = paragraphs.items
            .filter(p => Styles.includes(p.style));
        for (const parag of parags) {
            parag.select();
            const style = parag.style;
            try {
                const parent = parag.parentContentControlOrNullObject;
                parent.load(['tag']);
                await context.sync();
                if (parent.tag === tag) continue;//We escape paragraphs already wraped in a contentcontrol with the same tag
                console.log(`range style: ${parag.style} & text = ${parag.text}`);
                await insertContentControl(parag.getRange('Content'),
                    tag, tag, parags.indexOf(parag), RichText, style);
            } catch (error) {
                console.log(`Error from insertForAllParags() when trying to wrap the paragraph : ${parag.text}. Error :\n${error}`);
                continue
            }
        }
        await context.sync();

    })
}

async function insertDroDownListAll() {
    NOTIFICATION.innerHTML = '';
    const range = await getSelectionRange();
    if (!range) return;
    range.load(["text"]);
    const bookmark = 'temporaryBookmark';
    range.getRange(Word.RangeLocation.start).insertBookmark(bookmark);//!We must select the begining of the range otherwise insertContentControl() will fail. This is due to the shitty Word js api as usual
    await range.context.sync();
    const text = range.text;
    const find = text.split('/').join('');

    await Word.run(async (context) => {
        try {
            const matches = await searchString(find, context, false, text);
            for (const match of matches.items)
                await insertDropDownList(match, matches.items.indexOf(match) + 1);
            context.document.getBookmarkRange(bookmark).select();
            context.document.deleteBookmark(bookmark);
            await context.sync();
        } catch (error) {
            showNotification(`Error from insertDropDownList = ${error}`)
        }
    });
}
async function insertDropDownList(range: Word.Range | void, index: number = 0) {
    if (!range) range = await getSelectionRange();
    if (!range) return;
    range.load(["text"]);
    range.track();
    await range.context.sync();
    const options = range.text.split("/");
    if (!options.length) return showNotification("No options");
    showNotification(options.join());

    const ctrl = await insertContentControl(range, RTDropDownTag, RTDropDownTag, index, dropDownList, null, false, true);
    if (!ctrl) return;
    ctrl.dropDownListContentControl.deleteAllListItems();
    options.forEach(option => ctrl.dropDownListContentControl.addListItem(option));
    setCtrlsFontColor([ctrl], RTDropDownColor);
    setCtrlsColor([ctrl], RTDropDownColor);
    await ctrl.context.sync();
}
async function insertContentControl(range: Word.Range, title: string, tag: string, index: number = 1, type: Word.ContentControlType, style: string | null, cannotEdit: boolean = true, cannotDelete: boolean = true, placeHolder?: string): Promise<Word.ContentControl | undefined> {
    range.select();
    const styles = range.context.document.getStyles();
    styles.load(['nameLocal', 'type']);
    // Insert a rich text content control around the found range.
    //@ts-expect-error
    const ctrl = range.insertContentControl(type);
    ctrl.load(['id']);
    ctrl.track();
    await range.context.sync();
    // Set properties for the new content control.
    if (ctrl.id) console.log(`the newly created ContentControl id = ${ctrl.id} `);
    try {
        ctrl.select();
        ctrl.title = getCtrlTitle(title, ctrl.id);
        ctrl.tag = tag;
        ctrl.appearance = Word.ContentControlAppearance.boundingBox;
        if (placeHolder) ctrl.placeholderText = placeHolder;
        const foundStyle = styles.items.find(s => s.nameLocal === style);
        if (style && foundStyle?.type === Word.StyleType.character)
            ctrl.style = style;
        if (style) ctrl.getRange().style = style;
        ctrl.cannotDelete = cannotDelete;
        ctrl.cannotEdit = cannotEdit;//!This must come at the end after the style has been set.
        await ctrl.context.sync();
        showNotification(`Wrapped text in range ${index} with a content control.`);
        return ctrl;
    } catch (error) {
        showNotification(`There was an error while setting the properties of the newly crated contentcontrol by insertContentControl(): ${error}.`);
        return undefined
    }

}

async function getSelectionRange() {
    return await Word.run(async (context) => {
        const range = context.document
            .getSelection()
            .getRange('Content');
        range.load(['style', 'isEmpty']);
        range.track();
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


async function customizeContract(showNested: boolean = false) {
    USERFORM.innerHTML = '';
    const selected: string[] = [];
    const not: string = 'RTDelete';
    const processed = (id: number) => selected.find(t => t.includes(id.toString()));
    const TAGS = [...OPTIONS, RTDuplicateTag];
    const getSelectCtrls = (ctrls: ContentControl[]) => ctrls.filter(ctrl => TAGS.includes(ctrl.tag));
    const props = ['id', 'tag', 'title'];
    if (showNested) return await showNestedOptionsTree();

    await loopSelectCtrls();

    async function loopSelectCtrls() {
        await Word.run(async (context) => {
            const allRT = context.document.getContentControls();
            allRT.load(props);
            await context.sync();
            const selectCtrls = getSelectCtrls(allRT.items);
            for (const ctrl of selectCtrls)
                await promptForSelection(ctrl);

            const keep = selected
                .filter(title => !title.startsWith('!'))
                .map(title => Number(title));
            console.log(`keep = ${keep.join(',\n')}`);
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
                const selectCtrls = getSelectCtrls(allRT.items);
                const ids: Set<number> = new Set();

                for (const ctrl of selectCtrls) {
                    const nested = ctrl.getContentControls();
                    nested.load(['id','tag']);
                    await context.sync();
                    const nestedIds = nested.items.map(c => c.id);
                    const escape = keep.filter(id => nestedIds.includes(Number(id))) //!This means that ctrl has amongst its nested  contentcontrols one or more contentcontrols that we do not want to delete. We will hence keep the parent
                    const ctrls = [...nested.items];
                    if (!escape.length) ctrls.push(ctrl);// => it means ctrl hasn't any nested ctrl that we don't want to delete, so we can safely delet ctrl and its nested ctrls.

                    ctrls.forEach(c => {
                        const cannotDelete = keep.includes(c.id);
                        c.cannotDelete = cannotDelete;
                        if(!cannotDelete) c.cannotEdit = cannotDelete;//!we must set cannotEdit to false if the ctrl is to be deleted otherwise we will get an error from the shitty Word js api
                    });
                    
                    if (escape.length || ctrl.tag === RTDuplicateTag) continue;
                    if (keep.includes(ctrl.id)) continue;
                    ids.add(ctrl.id);
                }

                await context.sync();

                await deleteCtrls(ids);
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
                    if (keep.includes(ctrl.id)) return;
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
            await showSelectPrompt([ctrl]);
        } catch (error) {
            showNotification(`Error from promptForSelection() = ${error}`)
        }
    }

    async function showSelectPrompt(selectCtrls: ContentControl[]) {
        const blocks: selectBlock[] = [];
        const subOptions = async (id: number, direct: boolean) => await showSelectPrompt(await getSubOptions(id, direct));
        try {
            for (const ctrl of selectCtrls) {
                if (processed(ctrl.id)) continue;//!We must escape the ctrls that have already been processed
                ctrl.select();
                if (ctrl.tag === RTDuplicateTag) {
                    await duplicateBlock(ctrl.id);
                    continue
                };
                const addBtn = selectCtrls.indexOf(ctrl) + 1 === selectCtrls.length;
                const block = await insertPromptBlock(ctrl.id, addBtn);
                if (!block) continue;
                blocks.push(block);
                if (!block.container) await subOptions(ctrl.id, true);
                else if (!block.checkBox && block.btnNext) {
                    //!This is the case where selectCtrl has no "ctrlSi" contentControl as a direct child. We will await the user to click the button in order to process all the already displayed elements of selectCtrls[] until this point. Then, we will process the selectCtrl separetly before moving to the next selectCtrl in selectCtrls[]
                    await btnOnClick(blocks, block.btnNext);//We must await the user to click the button in order to process all the already displayed elements/options of selectCtrls[].
                    await subOptions(ctrl.id, true);//!We select only the direct select ctrls children
                }
                else if (block.btnNext) await btnOnClick(blocks, block.btnNext);//This is the case where btnNext was added because we reached the end of selectCtrls[] (addBtn = true). We then need to await the user to click the button in order to process all the already displayed elements/options of selectCtrls[].
            }
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
                const label = await labelRange(ctrl, RTSiTag);
                await context.sync();
                if (!label)
                    return !addBtn ? appendHTMLElements('') : { container: undefined };//!If this is not the last element in selectCtrls (addBtn = false) We will return a container with only a button to be clicked, otherwise, we will return a slectBlock with undefined container
                label.select();
                const text = label.text || `The ctrl label was found but no text could be retrieved ! ctrl title = ${ctrl.title}`;
                label.font.hidden = true;
                await context.sync();
                return appendHTMLElements(text, id.toString(), addBtn) as selectBlock;//The checkBox will have as id the title of the "select" contentcontrol}
            });
        }
    }

    function appendHTMLElements(text: string, id?: string, addBtn: boolean = false): selectBlock {
        const container = createHTMLElement('div', 'promptContainer', '', USERFORM) as HTMLDivElement;
        if (!id) return { container, btnNext: btn() }//!We return a container with a button with no checkBox
        const option = createHTMLElement('div', 'select', '', container);
        const checkBox = createHTMLElement('input', 'checkBox', '', option, id) as HTMLInputElement;//!We must give the checkBox the id of the selectCtrl because the id will be later used to retrieve the selectCtrl and process its children
        checkBox.type = 'checkbox';
        if (selected.includes(id)) checkBox.checked = true;//!Normaly this should never happen
        createHTMLElement('label', 'label', text, option) as HTMLParagraphElement;
        if (!addBtn) return { container, checkBox };
        return { container, checkBox, btnNext: btn() };

        function btn() {
            const btns = createHTMLElement('div', 'btns', '', container);
            return createHTMLElement('button', 'btnOK', 'Next', btns) as HTMLButtonElement;
        }
    }

    function btnOnClick(blocks: selectBlock[], btn: HTMLButtonElement): Promise<string[]> {
        return new Promise((resolve, reject) => {
            !btn ? resolve(selected) : btn.onclick = processBlocks;
            async function processBlocks() {
                const checkBoxes: [string, boolean][] =
                    blocks
                        .filter(block => block.checkBox)
                        //@ts-ignore
                        .map(block => [block.checkBox.id, block.checkBox.checked]);
                blocks.forEach(block => block.container?.remove());//We remove all the containers from the DOM

                for (const [id, checked] of checkBoxes) {
                    const subOptions = await getSubOptions(Number(id), checked);
                    if (checked)
                        await isSelected(id, subOptions);
                    else isNotSelected(id, subOptions);
                }
                resolve(selected);
            };
        });
    }

    async function isSelected(id: string, subOptions: ContentControl[] | undefined) {
        selected.push(id);
        if (subOptions) await showSelectPrompt(subOptions);
    };

    function isNotSelected(id: string | number, subOptions: ContentControl[]) {
        const exclude = (id: string | number) => `!${id}`;
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
            await insertClones(id);
        } catch (error) {
            showNotification(`${error}`)
        }

        async function insertClones(id: number) {
            await Word.run(async (context) => {
                const ctrl = context.document.contentControls.getById(id);
                ctrl.load(props);
                const label = await labelRange(ctrl, RTSectionTag);
                if (!label) return;
                await context.sync();
                if (!label.text) return showNotification("No lable text");
                ctrl.select();
                const message = `Combien de ${label.text} y'a-t-il ?`;
                let answer = Number(await promptForInput(message, '1'));
                if (isNaN(answer)) {
                    showNotification(`The provided text cannot be converted into a number: ${answer}`);
                    return await insertClones(id);
                } else if (answer < 1) return isNotSelected(id, await getSubOptions(id, false));
                const title = `${getCtrlTitle(ctrl.tag, id)}-Cloned ${answer}`
                ctrl.title = title;//!We must update the title in case it is no matching the id in the template.
                const ctrlContent = ctrl.getOoxml();
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
                const label = await labelRange(clone, RTSectionTag);
                await context.sync();
                if (!label) return;
                clone.title = `${getCtrlTitle(clone.tag, clone.id)}-${i}`;
                const text = `${label.text} ${i}`;
                label.insertText(text, replace);
                label.font.hidden = true;
                await context.sync();
                const div = createHTMLElement('div', '', text, USERFORM, '', false);
                await showSelectPrompt(await getSubOptions(clone.id, true));//!We select only the direct select ctrls children
                div.remove();
            });

        };

    }

    async function labelRange(parent: ContentControl, tag: string) {
        const ctrl = getFirstByTag(parent, tag);
        ctrl.load(['id', 'parentContentControl']);
        await parent.context.sync();
        if (ctrl.parentContentControl.id !== parent.id) return undefined;//!The label ctrl must be a direct child of the parent ctrl
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

    async function showNestedOptionsTree() {
        const selection = await getSelectionRange();
        if (!selection) return prepareTemplate();
        const ctrls = selection.getContentControls();
        ctrls.load(props);
        await selection.context.sync();
        const ctrl = ctrls.items[0];
        if (!ctrl.id) return failed('The selection is not inside a content control');
        if (!TAGS.includes(ctrl.tag)) return failed(`Ctrl is not a select control. Its tag is ${ctrl.tag}`);
        const subOptions = await getSubOptions(ctrl.id, true);
        await showSelectPrompt(subOptions);
        prepareTemplate();
        function failed(message: string) {
            showNotification(message);
            prepareTemplate();
        }
    }
};

async function deleteCtrls(ids: Set<number>) {
    await Word.run(async (context) => {
        for (const id of ids) {
            const ctrls = context.document.getContentControls();
            ctrls.load(['id', 'cannotDelete']);
            await context.sync();
            const ctrl = ctrls.items.find(ctrl => ctrl.id === id);
            if (!ctrl || ctrl.cannotDelete) continue;
            ctrl.delete(false);
            showNotification(`found and deleted ctrl with id = ${id}`)
        }
        await context.sync();
    })
}
async function promptForInput(question: string, deflt?: string, fun?: Function, cancel: boolean = true): Promise<string | void> {
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
    const tags = [RTSiTag, RTDescriptionTag, RTObsTag, RTSectionTag];
    Word.run(async (context) => {
        const allCtrls = context.document.getContentControls();
        allCtrls.load(['tag', 'title']);
        await context.sync();
        allCtrls.items.forEach(ctrl => {
            if (!tags.includes(ctrl.tag))
                return ctrl.appearance = Word.ContentControlAppearance.hidden;
            ctrl.select();
            ctrl.cannotDelete = false;
            ctrl.delete(ctrl.tag === RTDropDownTag)//!We keep the content of the dropdown ctrls
        });
        await context.sync();
    });

}


