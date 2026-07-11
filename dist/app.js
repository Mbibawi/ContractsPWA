/// <reference types="./types.d.ts" />
const version = "v11.13.2";
let USERFORM, NOTIFICATION;
const goHome = [() => mainUI(false), 'Home', 'Return to the main menu of the app'];
Office.onReady((info) => {
    // Check that we loaded into Word
    if (info.host !== Office.HostType.Word)
        return showAlert('This addin is designed to work on Word only');
    USERFORM = document.getElementById('userFormSection');
    NOTIFICATION = document.getElementById('notification');
    mainUI();
});
function mainUI(showVersion = true) {
    if (!USERFORM)
        return;
    USERFORM.innerHTML = '';
    NOTIFICATION.innerHTML = '';
    insertVersion();
    new EditContract().showMainBtn();
    new WordFileds().showMainBtn();
    function insertVersion() {
        if (!showVersion)
            return;
        const p = document.createElement('p');
        p.innerText = version;
        USERFORM.insertAdjacentElement('beforebegin', p);
    }
}
function showNotification(message, clear = false) {
    if (clear)
        NOTIFICATION.innerHTML = '';
    createHTMLElement('p', 'notification', message, NOTIFICATION, '', true);
}
function showAlert(message, clear = false) {
    const [modal, window] = getModalContainer(USERFORM, 'Alert', 'alert', false);
    createHTMLElement('p', '', message, window, '', true);
    const btn = createHTMLElement('button', '', 'OK', window, '', true);
    btn.onclick = () => { modal.remove(); };
}
function createHTMLElement(tag, css, textContent, parent, id, append = true) {
    const el = document.createElement(tag);
    if (textContent)
        el.textContent = textContent;
    if (css)
        el.classList = css;
    if (id)
        el.id = id;
    if (!parent)
        return el;
    append ? parent.appendChild(el) : parent.prepend(el);
    return el;
}
function getModalContainer(parent, textContent, id, append = true) {
    const modal = createHTMLElement('div', 'modal', textContent, parent, id, append);
    const window = createHTMLElement('div', 'modal-window', '', modal, '', append);
    return [modal, window];
}
function insertBtn([fun, label, hint], append = true, on = 'click') {
    if (!USERFORM)
        return;
    const wrapper = document.createElement('div');
    wrapper.style.position = 'relative';
    wrapper.style.display = 'inline-block';
    const htmlBtn = document.createElement('button');
    wrapper.appendChild(htmlBtn);
    append ? USERFORM.appendChild(wrapper) : USERFORM.prepend(wrapper);
    htmlBtn.innerText = label;
    htmlBtn.addEventListener(on, () => fun());
    addHint();
    return htmlBtn;
    function addHint() {
        if (!hint)
            return;
        const hintBox = document.createElement('div');
        hintBox.innerText = hint;
        hintBox.classList = 'hintBox';
        wrapper.appendChild(hintBox);
        htmlBtn.onmouseover = (e) => hideElement(hintBox, false, e);
        htmlBtn.onmouseout = (e) => hideElement(hintBox, true, e);
    }
}
function hideElement(element, hide, e) {
    e?.stopPropagation();
    if (hide)
        element.style.display = 'none';
    else
        element.style.display = 'block';
}
class WordContentCtrls {
    constructor() {
        //ContentControl tags
        this.StylePrefix = 'Contrat_';
        this.RTFieldTag = 'RTField';
        this.RTDropDownTag = 'RTList';
        this.RTDropDownColor = '#991c63';
        this.RTCloneTag = 'RTRepeat';
        this.RTSectionTag = 'RTSection'; //This tag is a contentcontrol which contains a text to be displayed (like a lable or a title) other than for choosing a specifc case (RTSi)
        this.RTSelectTag = 'RTSelect';
        this.RTOrTag = 'RTOr';
        this.RTObsTag = 'RTObs';
        this.RTDescriptionTag = 'RTDesc';
        this.RTSiTag = 'RTSi';
        this.RTDeleteTag = '>>>> DeleteCtrl >>>>';
        //Stylesreadonly 
        this.RTSectionStyle = `${this.StylePrefix}${this.RTSectionTag}`;
        this.RTObsStyle = `${this.StylePrefix}${this.RTObsTag}`;
        this.RTSiStyles = ['0', '1', '2', '3', '4'].map(n => `${this.StylePrefix}${this.RTSiTag}${n}cm`);
        this.RTDescriptionStyle = `${this.StylePrefix}${this.RTDescriptionTag}`;
        //Types
        this.richText = Word.ContentControlType.richText;
        this.richTextInline = Word.ContentControlType.richTextInline;
        this.richTextParag = Word.ContentControlType.richTextParagraphs;
        this.comboBox = Word.ContentControlType.comboBox;
        this.checkBox = Word.ContentControlType.checkBox;
        this.dropDownList = Word.ContentControlType.dropDownList;
        this.bounding = Word.ContentControlAppearance.boundingBox;
        this.hidden = Word.ContentControlAppearance.hidden;
        this.OPTIONS = [this.RTSelectTag, 'RTShow', 'RTEdit'];
    }
    /**
     *
     * @param ctrls
     */
    async insertFields(ctrls) {
        for (const ctrl of ctrls) {
            try {
                const start = ctrl.getRange(Word.RangeLocation.before);
                const field = await this.insertSingleFiled(start, ctrls.indexOf(ctrl));
                if (field)
                    field.untrack();
                ctrl.untrack();
                await ctrl.context.sync();
            }
            catch (error) {
                showNotification(`Error inserting field: ctrl.id = ${ctrl?.id}, error: ${error}`);
                continue;
            }
        }
        ;
    }
    async insertSingleFiled(range, i = 0, style = '') {
        if (!range) {
            range = (await this.getSelectionRange())?.getRange(Word.RangeLocation.start);
            if (!range)
                return console.log('could not retrieve the range to insert the field contentcontrol');
        }
        ;
        const field = await this.insertContentControl(range, this.RTFieldTag, this.RTFieldTag, i, this.richText, style, false, false, '[*]', ['id']);
        if (!field)
            return;
        // field.onExited.add(() => updateAllFields(field));
        field.font.bold = true;
        return field;
    }
    /**
     *
     * @param range The range which will be wraped by the inserted ContentControl
     * @param title The title of the ContentControl
     * @param tag   The tag of the ContentControl
     * @param index A number that will just be used in case the function is called for the creation of a serie (or an Array) of CpontentControls and we want to console.log the index of each ContentControl in the array after its creation.
     * @param type The new ContentControl's type
     * @param style The style of the ContentControl
     * @param cannotEdit Sets the boolean value of the cannotEdit property of the ContentControl
     * @param cannotDelete Sets the boolean value of the cannotDelete property of the ContentControl
     * @param placeHolder If provided, the ContentControl text will be initiated with the provided placeholder
     * @param props The new ContentControl's properties that we might want to get loaded after its creation
     * @returns The newly created Word.ContentControl object
     */
    async insertContentControl(range, title, tag, index = 1, type, style, cannotEdit = true, cannotDelete = true, placeHolder, props = []) {
        range.select();
        const ctrl = range.insertContentControl(type);
        ctrl.load(['id']);
        ctrl.track();
        await range.context.sync();
        // Set properties for the new content control.
        if (ctrl.id)
            console.log(`the newly created ContentControl id = ${ctrl.id} `);
        try {
            ctrl.select();
            ctrl.title = this.getCtrlTitle(title, ctrl.id);
            ctrl.tag = tag;
            ctrl.appearance = Word.ContentControlAppearance.boundingBox;
            if (placeHolder)
                ctrl.placeholderText = placeHolder;
            if (style)
                ctrl.getRange().style = style;
            ctrl.cannotDelete = cannotDelete;
            ctrl.cannotEdit = cannotEdit; //!This must come at the end after the style has been set.
            if (props.length) {
                //If the props agrument is passed, we assume the user intends to use the ContenControl object when returned by the function. Therefor we track it otherwise it will be garbage collected and it will not be able to work with it when returned
                ctrl.load(props);
                ctrl.track();
            }
            ;
            await range.context.sync();
            showNotification(`Wrapped text in range ${index} with a content control.`);
            return ctrl;
        }
        catch (error) {
            showNotification(`There was an error while setting the properties of the newly crated contentcontrol by insertContentControl(): ${error}.`);
            return undefined;
        }
    }
    async wrapSelectionWithContentControl(title, tag, type, style, cannotEdit, cannotDelete) {
        const range = await this.getSelectionRange();
        if (!range)
            return;
        if (!style && this.RTSiStyles.includes(range.style))
            style = range.style || null;
        await this.insertContentControl(range, title, tag, 0, type, style, cannotEdit, cannotDelete);
    }
    async getSelectionRange(props = [], fun) {
        return await Word.run(async (context) => {
            const range = context.document
                .getSelection()
                .getRange('Content');
            range.load(['style', 'isEmpty', ...props]);
            range.track();
            await context.sync();
            //if (range.isEmpty) return showAlert('The selection range is empty, you must select a text to continue');
            if (fun)
                await fun(range);
            return range;
        });
    }
    getFirstByTag(range, tag) {
        return range.getContentControls().getByTag(tag).getFirst();
    }
    getCtrlTitle(tag, id) {
        return `${tag}&${id}`;
    }
    async setCtrlsColor(ctrls, color) {
        ctrls.forEach(ctrl => ctrl.color = color);
    }
    setCtrlsFontColor(ctrls, color) {
        ctrls.forEach(c => c.getRange().font.color = color);
    }
    async setFieldTitle(title, field) {
        await Word.run(async (context) => {
            let range;
            if (!field) {
                range = await this.getSelectionRange();
                if (!range)
                    return;
                range.load('parentContentControlOrNullObject');
                await context.sync();
                field = range.parentContentControlOrNullObject;
            }
            field.load(['title', 'tag']);
            await context.sync();
            if (field?.tag !== this.RTFieldTag)
                return console.log('field tag !== RTFieldTag');
            field.title = `${this.RTFieldTag}&${title}`;
            await context.sync();
        });
    }
    async updateAllContentControlIDs() {
        const tags = [
            this.RTFieldTag,
            this.RTDropDownTag,
            this.RTDropDownColor,
            this.RTCloneTag,
            this.RTSectionTag,
            this.RTSelectTag,
            this.RTOrTag,
            this.RTObsTag,
            this.RTDescriptionTag,
            this.RTSiTag
        ];
        await Word.run(async (context) => {
            const ctrls = context.document.getContentControls();
            ctrls.load(['tag', 'title', 'id']);
            await context.sync();
            const relevant = ctrls.items.filter(ctrl => tags.includes(ctrl.tag));
            relevant.forEach(ctrl => ctrl.title = this.getCtrlTitle(ctrl.tag, ctrl.id));
            await context.sync();
        });
    }
    async updateAllFields(field) {
        await Word.run(async (context) => {
            let range;
            if (field) {
                field.load(['tag', 'title']);
                range = field.getRange('Content');
                range.load(['text']);
            }
            const ctrls = context.document.getContentControls();
            ctrls.load(['tag', 'title', 'id']);
            await context.sync();
            const tag = field?.tag || this.RTFieldTag;
            const fields = ctrls.items.filter(c => c.tag === tag);
            if (!field)
                return fields.filter(f => f.title !== this.getCtrlTitle(tag, f.id)).map(f => f.title);
            if (!range?.text)
                return console.log('could not retrieve the range from the filed');
            const sameTitle = fields.filter(f => f.title === field.title);
            for (const f of sameTitle)
                f.getRange().insertText(range.text, Word.InsertLocation.replace);
            await context.sync();
        });
    }
}
export class EditContract extends WordContentCtrls {
    constructor() {
        super(...arguments);
        this._stylesListId = 'stylesList';
        this._fields = new WordFileds();
        this.main = [
            [this.customizeContract, 'Customize Contract', 'Selecting and editing a contract template'],
            [this.prepareTemplate, 'Prepare Template', 'Creates a contract template'],
            [this.finalizeContract, 'Finalize Contract', 'Removing all the unwanted contentcontrols and issues the final versiofn of the contract'],
            [this.lockUnlockAll, 'Remove Cannot Delete For All', 'Toggels the cannot be deleted setting of all the contentcontrols in the document'],
            goHome,
        ];
        this.goBack = [() => {
                USERFORM.innerHTML = '';
                this.stylesList?.remove();
                this.showBtns(this.main);
            }, 'Go Back', 'Return to the previous menu'];
    }
    get stylesList() { return document.getElementById(this._stylesListId); }
    showMainBtn() {
        insertBtn([() => this.showBtns(this.main), 'Edit Contracts', undefined], false);
    }
    showBtns(btns = this.main, append = true, on = 'click') {
        USERFORM.innerHTML = '';
        const htmlBtns = btns.map(([fun, label, hint]) => insertBtn([fun.bind(this), label, hint], append, on));
        if (btns === this.main) {
            htmlBtns
                .slice(0, -1) //We exclude the goHome htmBtn
                .forEach(btn => btn?.addEventListener('click', () => [this.goBack, goHome].forEach(btn => insertBtn(btn, false))));
        }
        return htmlBtns;
    }
    ;
    prepareTemplate() {
        USERFORM.innerHTML = '';
        const searchString = this.searchString.bind(this), getSelectionRange = this.getSelectionRange.bind(this), insertContentControl = this.insertContentControl.bind(this), insertFields = this.insertFields.bind(this), setCtrlsColor = this.setCtrlsColor.bind(this), setCtrlsFontColor = this.setCtrlsFontColor.bind(this), promptForInput = this.promptForInput.bind(this), promptConfirm = this.promptConfirm.bind(this);
        const siTag = this.RTSiTag, selectTag = this.RTSelectTag, sectionTag = this.RTSectionTag, descTag = this.RTDescriptionTag, stylePrefix = this.StylePrefix, richText = this.richText, dorpDownTag = this.RTDropDownTag;
        const descStyle = this.RTDescriptionStyle, siStyle = this.RTSiStyles, sectionStyle = this.RTSectionStyle, dropDownColor = this.RTDropDownColor, dropDownList = this.dropDownList;
        const wrapRange = this.wrapSelectionWithContentControl.bind(this);
        function wrap(title, tag, type, style, cannotEdit, cannotDelete, label, hint) {
            return [
                () => wrapRange(title, tag, type, style, cannotEdit, cannotDelete),
                label,
                hint
            ];
        }
        ;
        const single = (tag, other) => `Inserts a single ${tag} contentcontrol at the begining of the selected range. ${other}If no range is selected, it will return.`;
        const all = (style, tag) => `Wraps all the pragraphs having as style ${style}, in a ${tag} contrentcontrol}`;
        const btns = [
            wrap(this.RTSiTag, this.RTSiTag, this.richText, this.RTSiStyles[0], true, true, 'Insert Single RT Si', single(this.RTSiTag)),
            wrap(this.RTSelectTag, this.RTSelectTag, this.richText, null, false, true, 'Insert Single RT Select', single(this.RTSelectTag, 'Any such contentControl is a container. Each contentcontrol having the same tag within its range, will be considered as an option to select or to exclude')),
            [insertRTBlock_Select_Si, 'Insert RT Select & Si Block', 'Finds the first paragraph formatted with any of the RTSiStyles. Wraps this paragraph in a RTSi ContentControl, Then wraps the whol selected range in a RTSelect ContentControl.'],
            [insertDropDownList, 'Insert a Dropdown List from selection', 'Creates a dropwdown list from the selected string. The options to choose from must be separated by "/"'],
            [() => insertRTDescription(true), 'Insert Single RT Description', single(this.RTDescriptionTag)],
            [this.insertSingleFiled, 'Insert ContentControl Field', single(this.RTFieldTag)],
            [this._fields.insertNewFILLINField, 'Insert FILLIN Field', single(this.RTFieldTag)],
            wrap(this.RTSectionTag, this.RTSectionTag, this.richText, this.RTSectionTag, true, true, 'Insert Single RT Section', single(this.RTSectionTag)),
            //wrap(this.RTOrTag, this.RTOrTag, this.richText, null, false, true, 'Insert Single RT OR', single(this.RTOrTag, 'need to check what it does')),
            wrap(this.RTCloneTag, this.RTCloneTag, this.richText, null, false, true, 'Insert Single RT Dublicate Block', single(this.RTCloneTag, 'need to check what it does')),
            wrap(this.RTObsTag, this.RTObsTag, this.richText, this.RTObsTag, true, true, 'Insert Single RT Obs', single(this.RTObsTag)),
            [insertDropDownListAll, 'Insert DropDown List For All Matches', 'It will check the document for all the strings matching the "/" separated values of the selected range and will convert them into drowpdown lists. The matching strings do not need to include the "/" mark'],
            [insertRTSiAll, 'Insert RT Si For All', all(this.RTSiStyles.join(' or '), this.RTSiTag)],
            [insertRTSectionAll, 'Insert RT Section For All', all(this.RTSectionStyle, this.RTSectionTag)],
            [insertRTDescription, 'Insert RT Description For All', all(this.RTDescriptionStyle, this.RTDescriptionTag)],
            [() => this.customizeContract(true), 'Show Nested Options Tree', 'Lists all the selection options in the document'],
            [this.updateAllContentControlIDs, 'Update ContentControl Titles', 'Updates the titles of all the ContentControls in the document'],
        ];
        this.showBtns(btns);
        if (!this.stylesList)
            showStylesList(this._stylesListId);
        function showStylesList(id) {
            Word.run(async (context) => {
                const allStyles = context.document.getStyles();
                allStyles.load(['nameLocal']);
                await context.sync();
                const styles = allStyles.items.filter(style => style.nameLocal.startsWith(stylePrefix));
                if (!styles.length)
                    return;
                const container = createHTMLElement('div', '', '', undefined, id);
                USERFORM.insertAdjacentElement('beforebegin', container);
                const select = createHTMLElement('select', '', '', container);
                styles.forEach(style => {
                    const option = createHTMLElement('option', '', style.nameLocal.split(stylePrefix)[1], select);
                    option.value = style.nameLocal;
                });
                select.onmouseenter = async () => {
                    const range = await getSelectionRange();
                    if (!range)
                        return;
                    const value = Array.from(select.options).find(o => o.value === range?.style)?.value || range.style;
                    range.untrack();
                    await range.context.sync();
                    if (value)
                        select.value = value;
                };
                select.onchange = async () => {
                    const range = await getSelectionRange();
                    if (!range)
                        return;
                    range.style = select.value;
                    range.untrack();
                    await range.context.sync();
                };
            });
        }
        async function insertRTDescription(selection = false, style = `${stylePrefix}Normal`) {
            NOTIFICATION.innerHTML = '';
            let ctrls = [];
            if (selection) {
                const range = await getSelectionRange();
                if (!range)
                    return;
                ctrls.push(await insertContentControl(range, descTag, descTag, 0, richText, descStyle, true, true, undefined, ['id']));
            }
            else
                ctrls = await findTextAndWrapItWithContentControl([descStyle], descTag, descTag, true, true);
            if (!ctrls?.length)
                return;
            //const ids = ctrls.map(c => c?.id || 0);
            await insertFields(ctrls.filter(ctrl => ctrl !== undefined));
        }
        async function addOrUpdateCustomXml(id, text, root = 'contractFields', prefix = 'contract', nameSpace = 'contract-namespace') {
            await Word.run(async (context) => {
                const newNode = `<${prefix}:${id} >${text}</${id}>`;
                const xmlParts = context.document.customXmlParts;
                let customPart = xmlParts.getByNamespace(nameSpace).getOnlyItemOrNullObject();
                customPart.load('isNullObject');
                await context.sync();
                if (customPart.isNullObject) {
                    // Create new XML part with root node
                    const newXml = `<${root} xmlns:${prefix}="${nameSpace}">${newNode}</${root}>`;
                    customPart = xmlParts.add(newXml);
                    await context.sync();
                    console.log("Created new custom XML part.");
                }
                else {
                    // Append new node to existing part
                    const xmlText = customPart.getXml();
                    await context.sync();
                    const xmlDoc = getXmlRootNode(xmlText);
                    const rootNode = xmlDoc.getElementsByName(root)[0];
                    if (!rootNode)
                        return console.warn("Root node not found.");
                    const newElement = xmlDoc.createElementNS(nameSpace, `${prefix}:${id}`);
                    newElement.textContent = text;
                    rootNode.appendChild(newElement);
                    await updateCustomXml(context, xmlParts, xmlDoc, customPart);
                }
            });
        }
        function getXmlRootNode(xmlText) {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlText.value, "application/xml");
            return xmlDoc;
        }
        async function updateCustomXml(context, xmlParts, xmlDoc, oldXmlPart) {
            const serializer = new XMLSerializer();
            const updatedXml = serializer.serializeToString(xmlDoc);
            oldXmlPart.delete(); // Remove old part
            xmlParts.add(updatedXml); // Add updated part
            await context.sync();
            console.log("Appended new node to existing XML part.");
        }
        async function insertRTBlock_Select_Si() {
            const range = await getSelectionRange();
            if (!range)
                return;
            //Wraping the range with ContentControl "RTSelect"
            const ctrl = await insertContentControl(range, selectTag, selectTag, undefined, richText, null, false, false, undefined, ['id']);
            if (!ctrl)
                return showAlert('Failed to insert the RTSelect ContentControl');
            try {
                const _ctrl = range.context.document.contentControls.getById(ctrl.id);
                _ctrl.load(['paragraphs', 'paragraphs/style']);
                await range.context.sync();
                const si = _ctrl.paragraphs.items.find(p => siStyle.includes(p.style));
                if (!si)
                    return showAlert('No paragraph styled with on of the "RTSi" styles was found in the selected range');
                si.track();
                //Wraping the paragraph with ContentControl "RTSi"
                await insertContentControl(si, siTag, siTag, undefined, richText, si.style, true, true);
                [range, ctrl, si].forEach(obj => obj.untrack());
                await range.context.sync();
            }
            catch (error) {
                console.log(error.debugInfo || error);
            }
        }
        function insertRTSiAll() {
            insertForAllParags(siStyle, siTag);
        }
        function insertRTSectionAll() {
            insertForAllParags([sectionStyle], sectionTag);
        }
        async function insertForAllParags(Styles, tag) {
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
                        if (parent.tag === tag)
                            continue; //We escape paragraphs already wraped in a contentcontrol with the same tag
                        console.log(`range style: ${parag.style} & text = ${parag.text}`);
                        await insertContentControl(parag.getRange('Content'), tag, tag, parags.indexOf(parag), richText, style);
                    }
                    catch (error) {
                        console.log(`Error from insertForAllParags() when trying to wrap the paragraph : ${parag.text}. Error :\n${error}`);
                        continue;
                    }
                }
                await context.sync();
            });
        }
        async function insertDropDownListAll() {
            NOTIFICATION.innerHTML = '';
            const range = await getSelectionRange();
            if (!range)
                return;
            range.load(["text"]);
            const bookmark = 'temporaryBookmark';
            range.getRange(Word.RangeLocation.start).insertBookmark(bookmark); //!We must select the begining of the range otherwise insertContentControl() will fail. This is due to the shitty Word js api as usual
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
                }
                catch (error) {
                    showNotification(`Error from insertDropDownList = ${error}`);
                }
            });
        }
        async function insertDropDownList(range, index = 0) {
            if (!range)
                range = await getSelectionRange();
            if (!range)
                return;
            range.load(["text", 'parentContentControlOrNullObject']);
            await range.context.sync();
            const parent = range.parentContentControlOrNullObject;
            parent.load('tag');
            await range.context.sync();
            if (parent.tag === dorpDownTag)
                return;
            const options = range.text.split("/");
            if (!options.length)
                return showNotification("No options");
            showNotification(options.join());
            const ctrl = await insertContentControl(range, dorpDownTag, dorpDownTag, index, dropDownList, null, false, true, undefined, ['id']);
            if (!ctrl)
                return;
            ctrl.dropDownListContentControl.deleteAllListItems();
            options.forEach(option => ctrl.dropDownListContentControl.addListItem(option));
            setCtrlsFontColor([ctrl], dropDownColor);
            setCtrlsColor([ctrl], dropDownColor);
            range.untrack();
            ctrl.untrack();
            await ctrl.context.sync();
        }
        /**
 * Wraps every occurrence of text formatted with a specific character style in a rich text content control.
 *
 * @param style The name of the character style to find (e.g., "Emphasis", "Strong", "MyCustomStyle").
 * @returns A Promise that resolves when the operation is complete.
 */
        async function findTextAndWrapItWithContentControl(styles, title, tag, cannotEdit, cannotDelete) {
            const { search, matchWildcards } = await searchs();
            if (!styles?.length)
                return showNotification(`The styles[] has 0 length, no styles are included, the function will return`);
            if (!search?.length)
                return showNotification('The provided search string is not valid');
            return await Word.run(async (context) => {
                const ctrls = [];
                for (const find of search) {
                    const matches = await searchString(find, context, matchWildcards);
                    if (!matches?.items.length)
                        continue;
                    matches.load(['style', 'text', 'parentContentControlOrNullObject']);
                    await context.sync();
                    const ranges = matches.items.filter(range => styles.includes(range.style));
                    showNotification(`Found ${ranges.length} ranges matching the search string. First range text = ${ranges[0].text}`);
                    ctrls.push(...await insertCtrls(ranges, context));
                }
                ;
                return ctrls;
            });
            async function insertCtrls(ranges, context) {
                const ctrls = [];
                for (const range of ranges) {
                    const parent = range.parentContentControlOrNullObject;
                    parent.load('tag');
                    await context.sync();
                    if (parent.tag === tag)
                        continue;
                    try {
                        const ctrl = await insertContentControl(range, title, tag, ranges.indexOf(range), richText, range.style, cannotEdit, cannotDelete, undefined, ['id']);
                        if (ctrl)
                            ctrls.push(ctrl);
                    }
                    catch (error) {
                        showNotification(`Error from insertCtrls() while inserting the contentControl() in the matching range. Error = ${error}`);
                    }
                }
                return ctrls;
            }
            async function searchs() {
                const separator = '_&_';
                const search = (await promptForInput(`Provide the search string. You can provide more than one string to search by separated by ${separator} witohout space`, separator))?.split(separator);
                const matchWildcards = await promptConfirm('Match Wild Cards');
                if (!styles)
                    styles = (await promptForInput(`Provide the styles that that need to be matched separated by ","`))?.split(',').map(style => style.trim()) || [];
                return { search, matchWildcards };
            }
        }
        async function associateFieldWithMain(mainID, subID, root = 'contractFields', prefix = 'contract', nameSpace = 'contract-namespace') {
            await Word.run(async (context) => {
                const ctrls = context.document.getContentControls();
                const main = ctrls.getById(mainID);
                const sub = ctrls.getById(subID);
                //const newNode = `<${prefix}:${ id } >${ text }</${id}>`;
                const xmlParts = context.document.customXmlParts;
                let customPart = xmlParts.getByNamespace(nameSpace).getOnlyItemOrNullObject();
                customPart.load('isNullObject');
                await context.sync();
                if (!customPart)
                    return console.log('customXmlPart not found');
            });
        }
    }
    ;
    async customizeContract(showNested = false) {
        USERFORM.innerHTML = '';
        const deleteCtrl = this.RTDeleteTag;
        const processed = [];
        const add = (id) => !processed.includes(id) ? processed.push(id) : processed;
        const RTDuplicateTag = this.RTCloneTag, RTSiTag = this.RTSiTag, RTSectionTag = this.RTSectionTag;
        const TAGS = [...this.OPTIONS, this.RTCloneTag];
        //const escape = (id: number) => processed.find(element => element.endsWith(id.toString()));
        const getSelectCtrls = (ctrls) => ctrls.filter(ctrl => TAGS.includes(ctrl.tag));
        const promptForInput = this.promptForInput.bind(this), getCtrlTitle = this.getCtrlTitle.bind(this), getFirstByTag = this.getFirstByTag.bind(this), getDocumentBase64 = this.getDocumentBase64.bind(this), getSelectionRange = this.getSelectionRange.bind(this), prepareTemplate = this.prepareTemplate.bind(this);
        const props = ['id', 'tag', 'title'];
        await loopSelectCtrls();
        async function loopSelectCtrls() {
            await Word.run(async (context) => {
                if (showNested)
                    return await showNestedOptionsTree(context);
                const allRT = context.document.getContentControls();
                allRT.load(props);
                await context.sync();
                const selectCtrls = getSelectCtrls(allRT.items);
                try {
                    for (const ctrl of selectCtrls)
                        await promptForSelection([ctrl], context);
                    await context.sync();
                    USERFORM.innerHTML = '';
                    const btn = createHTMLElement('button', '', 'Delete all the unselected cases ?', USERFORM, 'deleteAll', false);
                    btn.onclick = async () => {
                        await Word.run(async (context) => {
                            await deleteUnselected(context);
                            btn.disabled = true;
                            btn.textContent = 'Deleted';
                        });
                    };
                }
                catch (error) {
                    showNotification(`Error from promptForSelection() = ${error}`);
                }
                ;
            });
        }
        async function deleteUnselected(context) {
            try {
                await currentDoc();
                //await createNewDoc();
            }
            catch (error) {
                showNotification(`${error}`);
            }
            async function currentDoc() {
                const toDelete = context.document.getContentControls().getByTitle(deleteCtrl);
                toDelete.load(props);
                await context.sync();
                console.log(`toDelete = ${toDelete.items.map(ctrl => ctrl.id).join(',\n')}`);
                for (const ctrl of toDelete.items) {
                    try {
                        //if (ctrl.tag === RTDuplicateTag) continue;
                        unprotect(ctrl);
                        ctrl.delete(false);
                        await context.sync();
                    }
                    catch (error) {
                        console.log(`Error from deleting controls = ${error}. This is most probably caused by the contentcontrol (its id = ${ctrl.id}) having been already deleted with its parent`);
                    }
                }
            }
            ;
            async function createNewDoc() {
                await Word.run(async (context) => {
                    return; //!Desactivating working with new document created from template until we find a solution to the context issue
                    const template = await getTemplate();
                    console.log(template);
                    if (!template)
                        return showAlert('Failed to create the template');
                    const newDoc = context.application.createDocument(template);
                    const all = newDoc.contentControls;
                    all.load(['title', 'tag']);
                    await newDoc.context.sync();
                    showNotification(`All ctrls from newDoc = : ${all.items.map(c => c.title).join(', ')}`);
                    all.items.map(ctrl => {
                        if (ctrl.title === deleteCtrl)
                            return;
                        ctrl.cannotDelete = false;
                        ctrl.delete(false);
                    });
                    await newDoc.context.sync();
                    newDoc.open();
                });
            }
            function unprotect(ctrl) {
                ctrl.cannotDelete = false;
                ctrl.cannotEdit = false;
            }
        }
        async function subOptions(ctrl) {
            return await promptForSelection(await getDirectChildren(ctrl), ctrl.context);
        }
        ;
        async function getDirectChildren(ctrl) {
            const children = ctrl.getContentControls();
            children.load([...props, 'parentContentControl', 'parentContentControl.id', 'parentContentControl.tag']);
            await ctrl.context.sync();
            return children.items.filter(c => c.parentContentControl.id === ctrl.id);
        }
        async function promptForSelection(selectCtrls, context) {
            const blocks = [];
            try {
                await processCtrls();
            }
            catch (error) {
                return showNotification(`Error from showSelectPrompt() = ${error}`);
            }
            async function processCtrls() {
                for (const ctrl of selectCtrls) {
                    const id = ctrl.id;
                    if (processed.includes(id))
                        continue; //!We must escape the ctrls that have already been processed
                    ctrl.select();
                    add(id);
                    if (ctrl.tag === RTDuplicateTag) {
                        await duplicateBlock(ctrl, context);
                        continue;
                    }
                    ;
                    const isLast = selectCtrls.indexOf(ctrl) === selectCtrls.length - 1; //We check if this is the last contentcontrol in the array
                    const label = await labelRange(ctrl, RTSiTag);
                    const block = insertPromptBlock(ctrl, isLast, label);
                    if (!block)
                        continue;
                    blocks.push(block);
                    if (!block.wraper)
                        await subOptions(ctrl);
                    else if (!block.checkBox && block.btnNext) {
                        //!We excluded this case for the moment as a test
                        //!This is the case where selectCtrl has no "ctrlSi" contentControl as a direct child. We will await the user to click the button in order to process all the already displayed elements of selectCtrls[] until this point. Then, we will process the selectCtrl separetly before moving to the next selectCtrl in selectCtrls[]
                        await btnOnClick(block.btnNext, blocks, context); //We must await the user to click the button in order to process all the already displayed elements/options of selectCtrls[].
                        await subOptions(ctrl); //!We select only the direct select ctrls children
                    }
                    else if (block.btnNext)
                        await btnOnClick(block.btnNext, blocks, context); //This is the case where btnNext was added because we reached the end of selectCtrls[] (addBtn = true). We then need to await the user to click the button in order to process all the already displayed elements/options of selectCtrls[].
                }
            }
        }
        /**
         *
         * @param id The id of the contentControl containig the options
         * @param isLast If true, a button will be appended at the end
         * @returns
         */
        function insertPromptBlock(ctrl, isLast, label) {
            try {
                return showSelectUI();
            }
            catch (error) {
                return showNotification(`Error from insertPromptBlock() = ${error}`);
            }
            function showSelectUI() {
                if (!label) {
                    //A select ctrl that does not have a direct RTSiTag nested ctrl, is a container for other selectCtrls that need to be displayed. They are meant to offer multiple options for the user to select only one of them. But this is not always the case
                    return { wraper: undefined }; //!If this is not the last element in selectCtrls (addBtn == false) We will return a container with only a button to be clicked in order to move to the next select ctrl in the array
                    //return appendHTMLElements('');//!If this is not the last element in selectCtrls (addBtn == false) We will return a container with only a button to be clicked in order to move to the next select ctrl in the array
                    if (!isLast) {
                    }
                    else
                        return { wraper: undefined }; //!If this is the last element in selectCtrls array (addBtn == true), we will return a slectBlock with undefined container (which means that we will display the options nested within the select ctrl, but since this is the last ctrl in the selectCtrls array, we do not need to show a btnNext because when the nested options (i.e., the nested selectCtrls) will be displayed, a btnNext will be automatically inserted)
                }
                const text = label.text || `The ctrl label was found but no text could be retrieved ! ctrl title = ${ctrl.title}`;
                label.font.hidden = true;
                return appendHTMLElements(text, ctrl, isLast); //The checkBox will have as id the title of the "select" contentcontrol}
                ;
            }
        }
        function appendHTMLElements(text, ctrl, isLast = false) {
            const wraper = createHTMLElement('div', 'promptContainer', '', USERFORM);
            if (!ctrl)
                return { wraper, btnNext: btn() }; //!We return a container with a button with no checkBox
            const id = ctrl.id;
            const option = createHTMLElement('div', 'select', '', wraper);
            const chkbox = createHTMLElement('input', 'checkBox', '', option); //!We must give the checkBox the id of the selectCtrl because the id will be later used to retrieve the selectCtrl and process its children
            chkbox.type = 'checkbox';
            if (processed.includes(id))
                chkbox.checked = true; //!Normaly this should never happen
            createHTMLElement('label', 'label', text, option);
            if (!isLast)
                return { wraper, checkBox: { chkbox, ctrl } };
            return { wraper, checkBox: { chkbox, ctrl }, btnNext: btn() };
            function btn() {
                const btns = createHTMLElement('div', 'btns', '', wraper);
                return createHTMLElement('button', 'btnOK', 'Next', btns);
            }
        }
        function btnOnClick(btn, blocks, context) {
            return new Promise((resolve, reject) => {
                if (!btn)
                    resolve(false);
                btn.onclick = () => processBlocks();
                async function processBlocks() {
                    const checkBoxes = blocks
                        .filter(block => block.checkBox)
                        .map(block => [block.checkBox.ctrl, block.checkBox.chkbox.checked]);
                    blocks.forEach(block => block.wraper.remove()); //We remove all the containers from the DOM
                    for (const [ctrl, checked] of checkBoxes) {
                        const id = ctrl.id;
                        add(id);
                        const subOptions = await getSubOptions(ctrl, checked, context);
                        if (checked)
                            await isSelected(subOptions, context);
                        else
                            isNotSelected([ctrl, ...subOptions]);
                    }
                    resolve(true);
                }
                ;
            });
        }
        async function isSelected(subOptions, context) {
            //id = `${escapePrefix}${id}`;
            if (subOptions?.length)
                await promptForSelection(subOptions, context);
        }
        ;
        /**
         *
         * @param subOptions This is an array of all the contentControl children of the main control, including the main control itself
         * @param context
         */
        function isNotSelected(subOptions) {
            subOptions
                .forEach(ctrl => {
                ctrl.title = deleteCtrl;
                add(ctrl.id);
            }); //We are adding the "keep" prefix to the ids of the subOptions ctrls on purpose. This is because the parent ctrl will be deleted given that its id is added without the prefix. Hence all its children will (i.e., the subOptions) will be deleted as well with the parent ctrl. Adding the ids of each children, we will unnecesarily burden the list with a great number of ids, that will in all cases be deleted. 
        }
        ;
        async function getSubOptions(ctrl, directChildren, context) {
            const children = await getChildren();
            if (!directChildren)
                return children;
            return children.filter(c => c.parentContentControl.id === ctrl.id); //!We need to make sure we get only the direct children of the ctrl and not all the nested ctrls
            async function getChildren() {
                const children = ctrl.getContentControls();
                children.load([...props, 'parentContentControl']);
                await context.sync();
                return getSelectCtrls(children.items).filter(c => c.id !== ctrl.id); //!We must exclude the ctrl itself which in some cases may be returned as part of its children due to a bug in Word API
            }
        }
        /**
         *
         * @param id
         */
        async function duplicateBlock(ctrl, context) {
            const replace = Word.InsertLocation.replace;
            const after = Word.InsertLocation.after;
            try {
                await insertClones(ctrl);
            }
            catch (error) {
                showNotification(`${error}`);
            }
            async function insertClones(ctrl) {
                //const ctrl = context.document.contentControls.getById(id);
                //ctrl.load(props);
                const label = await labelRange(ctrl, RTSectionTag);
                //await context.sync();
                //if (!label) return;
                if (!label?.text)
                    return showNotification("No lable text");
                ctrl.select();
                const message = `Combien de ${label.text} y'a-t-il ?`;
                let answer = Number(await promptForInput(message, '1'));
                if (isNaN(answer)) {
                    showAlert(`The provided text cannot be converted into a number: ${answer}`);
                    return await insertClones(ctrl);
                }
                else if (answer < 1)
                    return isNotSelected([ctrl, ...await getSubOptions(ctrl, false, context)]);
                const title = `${getCtrlTitle(ctrl.tag, ctrl.id)}-Cloned ${answer}`;
                ctrl.title = title; //!We must update the title in case it is no matching the id in the template.
                const ctrlContent = ctrl.getOoxml();
                const range = ctrl.getRange();
                await context.sync();
                for (let i = 1; i < answer; i++)
                    range.insertOoxml(ctrlContent.value, after);
                const clones = ctrl.context.document.getContentControls().getByTitle(title);
                clones.load(props);
                label.font.hidden = true;
                await context.sync();
                const items = clones.items; //!clones.items.entries() caused the for loop to fail in scriptLab. The reason is unknown
                try {
                    for (const clone of items)
                        await processClone(clone, items.indexOf(clone) + 1);
                }
                catch (error) {
                    showNotification(`Error from processClone() = ${error}`);
                }
            }
            async function processClone(clone, i) {
                //const clone = context.document.contentControls.getById(id);
                clone.load(props);
                const label = await labelRange(clone, RTSectionTag);
                await clone.context.sync();
                if (!label)
                    return;
                clone.title = `${getCtrlTitle(clone.tag, clone.id)}-${i}`;
                const text = `${label} ${i}`;
                label.insertText(text, replace);
                await context.sync();
                const div = createHTMLElement('div', '', text, USERFORM, '', false);
                await promptForSelection(await getSubOptions(clone, true, context), context); //!We select only the direct select ctrls children
                div.remove();
                ;
            }
            ;
        }
        async function labelRange(parent, tag) {
            const ctrl = getFirstByTag(parent, tag);
            ctrl?.load(['id', 'parentContentControl']);
            const range = ctrl?.getRange('Content');
            range?.load(['text']);
            ctrl.cannotEdit = false;
            range.font.hidden = false;
            await parent.context.sync();
            if (ctrl?.parentContentControl?.id !== parent.id)
                return undefined; //!The label ctrl must be a direct child of the parent ctrl
            return range;
        }
        function getFileURL() {
            let url;
            Office.context.document.getFilePropertiesAsync(undefined, (result) => {
                if (result.error)
                    return;
                url = result.value.url;
            });
            return url;
        }
        async function getTemplate() {
            try {
                return await getDocumentBase64();
            }
            catch (error) {
                showNotification(`Failed to create new Doc: ${error}`);
            }
        }
        async function showNestedOptionsTree(context) {
            const selection = await getSelectionRange();
            if (!selection)
                return prepareTemplate();
            const ctrls = selection.getContentControls();
            ctrls.load(props);
            await selection.context.sync();
            const ctrl = ctrls.items[0];
            if (!ctrl.id)
                return failed('The selection is not inside a content control');
            if (!TAGS.includes(ctrl.tag))
                return failed(`Ctrl is not a select control. Its tag is ${ctrl.tag}`);
            const subOptions = await getSubOptions(ctrl, true, context);
            await promptForSelection(subOptions, context);
            prepareTemplate();
            function failed(message) {
                showNotification(message);
                prepareTemplate();
            }
        }
    }
    ;
    async finalizeContract() {
        const toDelte = this.RTDeleteTag;
        const remove = [this.RTSiTag, this.RTDescriptionTag, this.RTObsTag, this.RTSectionTag, toDelte]; //The contentcontrol and its content will be deleted.
        const content = [this.RTSelectTag, this.RTCloneTag]; //We will delete the contentControl but keep its content
        const styles = [...this.RTSiStyles, this.RTSectionStyle, this.RTObsStyle, this.RTDescriptionStyle];
        await Word.run(async (context) => {
            const allCtrls = context.document.getContentControls();
            allCtrls.load(['tag', 'title']);
            await context.sync();
            for (const ctrl of allCtrls.items) {
                ctrl.cannotDelete = false; //!We must remove the cannotDelete from all ctrls because this will prevent deleting the line on which there is a hidden contentcontrol
                if (!ctrl?.tag)
                    continue;
                if (remove.includes(ctrl.tag) || ctrl.title === toDelte)
                    ctrl.delete(false);
                else if (content.includes(ctrl.tag))
                    ctrl.delete(true);
                else
                    ctrl.appearance = Word.ContentControlAppearance.hidden;
            }
            await context.sync();
            const body = context.document.body.getRange();
            body.load('paragraphs');
            await context.sync();
            const parags = body.paragraphs;
            parags.load('style');
            await context.sync();
            parags.items
                .filter(p => styles.includes(p.style))
                .forEach(p => p.style = `${this.StylePrefix}Normal`);
            await context.sync();
        });
    }
    async searchString(search, context, matchWildcards, replaceWith) {
        const searchResults = context.document.body.search(search, { matchWildcards: matchWildcards });
        searchResults.load(['style', 'text']);
        searchResults.track();
        await context.sync();
        if (!replaceWith)
            return searchResults;
        for (const match of searchResults.items)
            match.insertText(replaceWith, Word.InsertLocation.replace);
        await context.sync();
        return await this.searchString(replaceWith, context, false);
    }
    async promptConfirm(question, fun) {
        if (!question)
            question = 'No question was provided !!!';
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
        function confirm(confirm) {
            container.remove();
            if (fun)
                fun(confirm);
            return confirm;
        }
    }
    async promptForInput(question, deflt, fun, cancel = true) {
        if (!question)
            return '';
        const container = createHTMLElement('div', 'promptContainer', '', USERFORM);
        const prompt = createHTMLElement('div', 'prompt', '', container);
        const ask = createHTMLElement('p', 'ask', question, prompt);
        const input = createHTMLElement('input', 'answer', '', prompt);
        const btns = createHTMLElement('div', 'btns', '', prompt);
        const btnOK = createHTMLElement('button', 'btnOK', 'OK', btns);
        const btnCancel = createHTMLElement('button', 'btnCancel', 'Cancel', btns);
        if (deflt)
            input.value = deflt;
        return new Promise((resolve, reject) => {
            btnCancel.onclick = () => reject(container.remove());
            btnOK.onclick = () => {
                const answer = input.value;
                console.log('user answer = ', answer);
                container.remove();
                if (fun)
                    fun(answer);
                resolve(answer);
            };
        });
    }
    ;
    /**
     * Asynchronously gets the entire document content as a Base64 string.
     * This function handles multi-slice documents by requesting each slice in parallel.
     * @returns A Promise that resolves with the Base64-encoded document content.
     */
    async getDocumentBase64() {
        const failed = (result) => result.status !== Office.AsyncResultStatus.Succeeded;
        const sliceSize = 16 * 1024; //!We need not to exceed the Maximum call stack limit when the slices will be passed to String.FromCharCode()
        return new Promise((resolve, reject) => {
            // Step 1: Request the document as a compressed file.
            Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: sliceSize }, (fileResult) => processFile(fileResult));
            function processFile(fileResult) {
                if (failed(fileResult))
                    return reject(fileResult.error);
                const file = fileResult.value;
                const sliceCount = file.sliceCount;
                const slices = [];
                getSlice();
                function getSlice() {
                    file.getSliceAsync(slices.length, (sliceResult) => processSlice(sliceResult));
                }
                function processSlice(sliceResult) {
                    try {
                        if (failed(sliceResult))
                            return file.closeAsync(() => reject(sliceResult.error));
                        slices.push(sliceResult.value.data);
                        if (slices.length < sliceCount)
                            return getSlice();
                        const binaryString = slices.map(slice => String.fromCharCode(...slice)).join('');
                        const base64String = btoa(binaryString);
                        file.closeAsync(() => resolve(base64String));
                    }
                    catch (error) {
                        showNotification(`${error}, succeeded = ${sliceResult.status}, loaded = ${slices.length}`);
                    }
                }
            }
        });
    }
    async deleteAllNotSelected(selected, wdDoc) {
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
    async setCanBeEditedForAllSelectCtrls(edit = true) {
        await Word.run(async (context) => {
            const ctrls = context.document
                .contentControls;
            ctrls.load(['title', 'tag']);
            await context.sync();
            ctrls.items.forEach(ctrl => {
                if (this.OPTIONS.indexOf(ctrl.tag) > -1)
                    ctrl.cannotEdit = edit;
            });
            await context.sync();
        });
    }
    setRangeStyle(objs, style) {
        objs.forEach(o => o.getRange().style = style);
    }
    async lockUnlockAll(unlock = false, tags = []) {
        await Word.run(async (context) => {
            const all = context.document.getContentControls();
            all.load(['tag', 'id']);
            await context.sync();
            let ctrls = all.items;
            if (tags.length)
                ctrls = ctrls.filter(c => tags.includes(c.tag));
            for (const ctrl of ctrls)
                ctrl.cannotDelete = unlock;
            await context.sync();
        });
    }
}
;
export class WordFileds extends WordContentCtrls {
    constructor() {
        super(...arguments);
        this._fillIn = Word.FieldType.fillIn;
    }
    showMainBtn() {
        insertBtn([() => this.showButtons(), 'Edit FILLIN Fields', 'Displays the user interface for editing the existing FILLIN fiels, or inserting new FILLIN fields'], true);
    }
    showButtons() {
        USERFORM.innerHTML = '';
        insertBtn([async () => await this.showInputs(), 'Edit The FILLIN Fields', 'Shows the interface to edit the FILLIN fiels in the document'], true);
        insertBtn([() => this.insertNewFILLINField(), 'Insert new FILLIN filed', 'Inserts a new FILLIN field in the selected range. It replaces the selected text with the FILLIN field'], true);
        insertBtn(goHome);
    }
    async showInputs() {
        USERFORM.innerHTML = '';
        await Word.run(async (context) => {
            const fields = context.document.body.fields;
            fields.load(['code', 'result', 'type']);
            await context.sync();
            const fillIn = fields.items.filter(field => field.type === this._fillIn);
            if (!fillIn.length)
                return console.log("no FILLIN fields were found");
            fillIn.forEach(field => field.result.load('text'));
            await context.sync();
            const inputs = fillIn.map((field, index) => {
                const code = field.code;
                console.log("field code = " + code);
                const match = code.match(/(?:FILLIN|ASK)\s+"([^"]+)"/i);
                if (!match || !match.length)
                    return undefined;
                const lable = match[1];
                if (!lable) {
                    console.log('could not extract label from code = ' + code);
                    return undefined;
                }
                ;
                const div = createHTMLElement('div', '', '', USERFORM, '', true);
                createHTMLElement('label', '', lable, div, '', true);
                const input = createHTMLElement('input', '', '', div, `FILLIN_${index.toString()}`, true);
                input.value = field.result.text;
                return [input, field];
            }).filter(item => item !== undefined);
            const showBtns = this.showButtons.bind(this);
            await awaitPromise();
            async function awaitPromise() {
                return new Promise((resolve) => {
                    const edit = async (cancel = false) => {
                        if (cancel)
                            return resolve(showBtns());
                        for (const [input, field] of inputs) {
                            if (!input.value)
                                continue;
                            if (!field)
                                return console.log('field not found');
                            field.result.insertText(input.value, Word.InsertLocation.replace);
                            console.log('Modified field = ' + field.code);
                        }
                        await context.sync();
                        resolve(showBtns());
                    };
                    insertBtn([() => edit(false), 'Update All Fileds From Inputs', 'Parses the values of the inputs, and updates the corresponding fields'], true);
                    insertBtn([() => edit(true), 'Cancel and go back', 'Cancels the editing session'], false); //We insert the goHome navigation button on top of all the inputs
                });
            }
            ;
        });
    }
    async insertNewFILLINField() {
        const create = createHTMLElement;
        const [modal, window] = getModalContainer(USERFORM, '', 'newField', false);
        showDialogue();
        function showDialogue() {
            const labels = ['Provide the FILLIN field prompt', 'Provide the FILLIN default value'];
            const inputs = labels.map(label => {
                create('label', '', label, window, undefined, true);
                return create('input', '', '', window, '', true);
            });
            const btn = create('button', '', 'Insert FILLIN Field', window, 'ok', true);
            btn.onclick = () => onClick(inputs[0].value, inputs[1].value || '[*]');
        }
        async function onClick(question, deflt) {
            const type = Word.FieldType.empty; //!We chose the empty field on purpose
            await Word.run(async (context) => {
                const range = context.document.getSelection().getRange(Word.RangeLocation.whole);
                const field = range.insertField(Word.InsertLocation.replace, type);
                field.code = `FILLIN "${question}"  \\d ${deflt}  \\* MERGEFORMAT`;
                field.updateResult();
                modal.remove();
                await context.sync();
            });
        }
    }
}
//# sourceMappingURL=app.js.map