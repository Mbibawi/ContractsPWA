/// <reference types="./types.d.ts" />
const version = "v11.9.8";
let USERFORM, NOTIFICATION;
const goHome = [() => mainUI(false), 'Home', 'Return to the main menu of the app'];
Office.onReady((info) => {
    // Check that we loaded into Word
    if (info.host !== Office.HostType.Word)
        return showNotification('This addin is designed to work on Word only');
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
    async insertFields(ids, style) {
        await Word.run(async (context) => {
            for (const id of ids) {
                if (!id)
                    continue;
                const ctrl = context.document.getContentControls().getById(id);
                await context.sync();
                try {
                    const start = ctrl.getRange(Word.RangeLocation.before);
                    await this.insertSingleFiled(start, ids.indexOf(id));
                }
                catch (error) {
                    showNotification(`Error inserting field: ctrl.id = ${ctrl?.id}, error: ${error}`);
                    continue;
                }
            }
            await context.sync();
        });
    }
    async insertSingleFiled(range, i = 0, style = '') {
        if (!range) {
            range = (await this.getSelectionRange())?.getRange(Word.RangeLocation.start);
            if (!range)
                return console.log('could not retrieve the range to insert the field contentcontrol');
        }
        ;
        const field = await this.insertContentControl(range, this.RTFieldTag, this.RTFieldTag, i, this.richText, style, false, false, '[*]');
        if (!field)
            return;
        // field.onExited.add(() => updateAllFields(field));
        field.font.bold = true;
    }
    async insertContentControl(range, title, tag, index = 1, type, style, cannotEdit = true, cannotDelete = true, placeHolder) {
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
        if (ctrl.id)
            console.log(`the newly created ContentControl id = ${ctrl.id} `);
        try {
            ctrl.select();
            ctrl.title = this.getCtrlTitle(title, ctrl.id);
            ctrl.tag = tag;
            ctrl.appearance = Word.ContentControlAppearance.boundingBox;
            if (placeHolder)
                ctrl.placeholderText = placeHolder;
            const foundStyle = styles.items.find(s => s.nameLocal === style);
            if (style && foundStyle?.type === Word.StyleType.character)
                ctrl.style = style;
            if (style)
                ctrl.getRange().style = style;
            ctrl.cannotDelete = cannotDelete;
            ctrl.cannotEdit = cannotEdit; //!This must come at the end after the style has been set.
            await ctrl.context.sync();
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
        if (this.RTSiStyles.includes(range.style))
            style = range.style;
        await this.insertContentControl(range, title, tag, 0, type, style, cannotEdit, cannotDelete);
    }
    async getSelectionRange() {
        return await Word.run(async (context) => {
            const range = context.document
                .getSelection()
                .getRange('Content');
            range.load(['style', 'isEmpty']);
            range.track();
            await context.sync();
            if (range.isEmpty)
                return showNotification('The selection range is empty');
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
        const wrapSelection = this.wrapSelectionWithContentControl.bind(this), getSelectionRange = this.getSelectionRange.bind(this), StylePrefix = this.StylePrefix;
        function wrap(title, tag, type, style, cannotEdit, cannotDelete, label, hint) {
            return [
                () => wrapSelection(title, tag, type, style, cannotEdit, cannotDelete),
                label,
                hint
            ];
        }
        ;
        const single = (tag, other) => `Inserts a single ${tag} contentcontrol at the begining of the selected range. ${other}If no range is selected, it will return.`;
        const all = (style, tag) => `Wraps all the pragraphs having as style ${style}, in a ${tag} contrentcontrol}`;
        const btns = [
            wrap(this.RTSiTag, this.RTSiTag, this.richText, this.RTSiStyles[0], true, true, 'Insert Single RT Si', single(this.RTSiTag)),
            [() => this.insertRTDescription(true), 'Insert Single RT Description', single(this.RTDescriptionTag)],
            wrap(this.RTSelectTag, this.RTSelectTag, this.richText, null, false, true, 'Insert Single RT Select', single(this.RTSelectTag, 'Any such contentControl is a container. Each contentcontrol having the same tag within its range, will be considered as an option to select or to exclude')),
            wrap(this.RTSectionTag, this.RTSectionTag, this.richText, this.RTSectionTag, true, true, 'Insert Single RT Section', single(this.RTSectionTag)),
            //wrap(this.RTOrTag, this.RTOrTag, this.richText, null, false, true, 'Insert Single RT OR', single(this.RTOrTag, 'need to check what it does')),
            wrap(this.RTCloneTag, this.RTCloneTag, this.richText, null, false, true, 'Insert Single RT Dublicate Block', single(this.RTCloneTag, 'need to check what it does')),
            [this.insertDropDownList, 'Insert a Dropdown List from selection', 'Creates a dropwdown list from the selected string. The options to choose from must be separated by "/"'],
            wrap(this.RTObsTag, this.RTObsTag, this.richText, this.RTObsTag, true, true, 'Insert Single RT Obs', single(this.RTObsTag)),
            [this.insertDropDownListAll, 'Insert DropDown List For All Matches', 'It will check the document for all the strings matching the "/" separated values of the selected range and will convert them into drowpdown lists. The matching strings do not need to include the "/" mark'],
            [this.insertRTSiAll, 'Insert RT Si For All', all(this.RTSiStyles.join(' or '), this.RTSiTag)],
            [this.insertRTSectionAll, 'Insert RT Section For All', all(this.RTSectionStyle, this.RTSectionTag)],
            [this.insertRTDescription, 'Insert RT Description For All', all(this.RTDescriptionStyle, this.RTDescriptionTag)],
            [this.insertSingleFiled, 'Insert ContentControl Field', single(this.RTFieldTag)],
            [this._fields.insertNewFILLINField, 'Insert FILLIN Field', single(this.RTFieldTag)],
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
                const styles = allStyles.items.filter(style => style.nameLocal.startsWith(StylePrefix));
                if (!styles.length)
                    return;
                const container = createHTMLElement('div', '', '', undefined, id);
                USERFORM.insertAdjacentElement('beforebegin', container);
                const select = createHTMLElement('select', '', '', container);
                styles.forEach(style => {
                    const option = createHTMLElement('option', '', style.nameLocal.split(StylePrefix)[1], select);
                    option.value = style.nameLocal;
                });
                select.onmouseenter = async () => {
                    const range = await getSelectionRange();
                    if (!range)
                        return;
                    const value = Array.from(select.options).find(o => o.value === range?.style)?.value || range.style;
                    if (value)
                        select.value = value;
                    range.untrack();
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
    }
    /**
     * Wraps every occurrence of text formatted with a specific character style in a rich text content control.
     *
     * @param style The name of the character style to find (e.g., "Emphasis", "Strong", "MyCustomStyle").
     * @returns A Promise that resolves when the operation is complete.
     */
    async findTextAndWrapItWithContentControl(styles, title, tag, cannotEdit, cannotDelete) {
        const insertContentControl = this.insertContentControl.bind(this), promptForInput = this.promptForInput, promptConfirm = this.promptConfirm, RichText = this.richText;
        const { search, matchWildcards } = await searchs();
        if (!styles?.length)
            return showNotification(`The styles[] has 0 length, no styles are included, the function will return`);
        if (!search?.length)
            return showNotification('The provided search string is not valid');
        return await Word.run(async (context) => {
            const ctrls = [];
            for (const find of search) {
                const matches = await this.searchString(find, context, matchWildcards);
                if (!matches?.items.length)
                    continue;
                matches.load(['style', 'text', 'parentContentControlOrNullObject']);
                await context.sync();
                const ranges = matches.items.filter(range => styles.includes(range.style));
                showNotification(`Found ${ranges.length} ranges matching the search string. First range text = ${ranges[0].text}`);
                ctrls.push(...await insertCtrls(ranges));
            }
            ;
            return ctrls;
            async function insertCtrls(ranges) {
                const ctrls = [];
                for (const range of ranges) {
                    const parent = range.parentContentControlOrNullObject;
                    parent.load('tag');
                    await context.sync();
                    if (parent.tag === tag)
                        continue;
                    try {
                        const ctrl = await insertContentControl(range, title, tag, ranges.indexOf(range), RichText, range.style, cannotEdit, cannotDelete);
                        if (ctrl)
                            ctrls.push(ctrl);
                    }
                    catch (error) {
                        showNotification(`Error from insertCtrls() while inserting the contentControl() in the matching range. Error = ${error}`);
                    }
                }
                return ctrls;
            }
        });
        async function searchs() {
            const separator = '_&_';
            const search = (await promptForInput(`Provide the search string. You can provide more than one string to search by separated by ${separator} witohout space`, separator))?.split(separator);
            const matchWildcards = await promptConfirm('Match Wild Cards');
            if (!styles)
                styles = (await promptForInput(`Provide the styles that that need to be matched separated by ","`))?.split(',') || [];
            return { search, matchWildcards };
        }
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
    async insertRTDescription(selection = false, style = `${this.StylePrefix}Normal`) {
        NOTIFICATION.innerHTML = '';
        let ctrls;
        if (selection) {
            const range = await this.getSelectionRange();
            if (!range)
                return showNotification('No Text Was selected !');
            ctrls = [await this.insertContentControl(range, this.RTDescriptionTag, this.RTDescriptionTag, 0, this.richText, this.RTDescriptionStyle, true, true)];
        }
        else
            ctrls = await this.findTextAndWrapItWithContentControl([this.RTDescriptionStyle], this.RTDescriptionTag, this.RTDescriptionTag, true, true);
        if (!ctrls?.length)
            return;
        const ids = ctrls.map(c => c?.id || 0);
        await this.insertFields(ids, style);
    }
    async addOrUpdateCustomXml(id, text, root = 'contractFields', prefix = 'contract', nameSpace = 'contract-namespace') {
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
                const xmlDoc = this.getXmlRootNode(xmlText);
                const rootNode = xmlDoc.getElementsByName(root)[0];
                if (!rootNode)
                    return console.warn("Root node not found.");
                const newElement = xmlDoc.createElementNS(nameSpace, `${prefix}:${id}`);
                newElement.textContent = text;
                rootNode.appendChild(newElement);
                await this.updateCustomXml(context, xmlParts, xmlDoc, customPart);
            }
        });
    }
    getXmlRootNode(xmlText) {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText.value, "application/xml");
        return xmlDoc;
    }
    async updateCustomXml(context, xmlParts, xmlDoc, oldXmlPart) {
        const serializer = new XMLSerializer();
        const updatedXml = serializer.serializeToString(xmlDoc);
        oldXmlPart.delete(); // Remove old part
        xmlParts.add(updatedXml); // Add updated part
        await context.sync();
        console.log("Appended new node to existing XML part.");
    }
    async associateFieldWithMain(mainID, subID, root = 'contractFields', prefix = 'contract', nameSpace = 'contract-namespace') {
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
    insertRTSiAll() {
        this.insertForAllParags(this.RTSiStyles, this.RTSiTag);
    }
    insertRTSectionAll() {
        this.insertForAllParags([this.RTSectionStyle], this.RTSectionTag);
    }
    async insertForAllParags(Styles, tag) {
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
                    await this.insertContentControl(parag.getRange('Content'), tag, tag, parags.indexOf(parag), this.richText, style);
                }
                catch (error) {
                    console.log(`Error from insertForAllParags() when trying to wrap the paragraph : ${parag.text}. Error :\n${error}`);
                    continue;
                }
            }
            await context.sync();
        });
    }
    async insertDropDownListAll() {
        NOTIFICATION.innerHTML = '';
        const range = await this.getSelectionRange();
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
                const matches = await this.searchString(find, context, false, text);
                for (const match of matches.items)
                    await this.insertDropDownList(match, matches.items.indexOf(match) + 1);
                context.document.getBookmarkRange(bookmark).select();
                context.document.deleteBookmark(bookmark);
                await context.sync();
            }
            catch (error) {
                showNotification(`Error from insertDropDownList = ${error}`);
            }
        });
    }
    async insertDropDownList(range, index = 0) {
        if (!range)
            range = await this.getSelectionRange();
        if (!range)
            return;
        range.load(["text", 'parentContentControlOrNullObject']);
        await range.context.sync();
        const parent = range.parentContentControlOrNullObject;
        parent.load('tag');
        await range.context.sync();
        if (parent.tag === this.RTDropDownTag)
            return;
        const options = range.text.split("/");
        if (!options.length)
            return showNotification("No options");
        showNotification(options.join());
        const ctrl = await this.insertContentControl(range, this.RTDropDownTag, this.RTDropDownTag, index, this.dropDownList, null, false, true);
        if (!ctrl)
            return;
        ctrl.dropDownListContentControl.deleteAllListItems();
        options.forEach(option => ctrl.dropDownListContentControl.addListItem(option));
        this.setCtrlsFontColor([ctrl], this.RTDropDownColor);
        this.setCtrlsColor([ctrl], this.RTDropDownColor);
        await ctrl.context.sync();
        range.untrack();
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
    async customizeContract(showNested = false) {
        USERFORM.innerHTML = '';
        const escapePrefix = '>>>>escape>>>>';
        const processed = [];
        const RTDuplicateTag = this.RTCloneTag, RTSiTag = this.RTSiTag, RTSectionTag = this.RTSectionTag;
        const TAGS = [...this.OPTIONS, this.RTCloneTag];
        const escape = (id) => processed.find(element => element.endsWith(id.toString()));
        const getSelectCtrls = (ctrls) => ctrls.filter(ctrl => TAGS.includes(ctrl.tag));
        const promptForInput = this.promptForInput.bind(this), getCtrlTitle = this.getCtrlTitle.bind(this), getFirstByTag = this.getFirstByTag.bind(this), getDocumentBase64 = this.getDocumentBase64.bind(this), getSelectionRange = this.getSelectionRange.bind(this), prepareTemplate = this.prepareTemplate.bind(this);
        const props = ['id', 'tag', 'title'];
        if (showNested)
            return await showNestedOptionsTree();
        await loopSelectCtrls();
        async function loopSelectCtrls() {
            await Word.run(async (context) => {
                const allRT = context.document.getContentControls();
                allRT.load(props);
                await context.sync();
                const selectCtrls = getSelectCtrls(allRT.items);
                try {
                    for (const ctrl of selectCtrls)
                        await promptForSelection([ctrl]);
                    await deleteUnselected();
                }
                catch (error) {
                    showNotification(`Error from promptForSelection() = ${error}`);
                }
                ;
            });
        }
        async function deleteUnselected() {
            const toDelete = processed
                .filter(id => !id.startsWith(escapePrefix))
                .map(id => Number(id));
            console.log(`toDelete = ${toDelete.join(',\n')}`);
            try {
                await currentDoc();
                processed.length = 0; //We remove any element in selected
                //await createNewDoc();
            }
            catch (error) {
                showNotification(`${error}`);
            }
            async function currentDoc() {
                await Word.run(async (context) => {
                    const allRT = context.document.getContentControls();
                    allRT.load(props);
                    await context.sync();
                    const selectCtrls = getSelectCtrls(allRT.items);
                    for (const ctrl of selectCtrls) {
                        if (ctrl.tag === RTDuplicateTag) {
                            unprotect(ctrl); //!This important, otherwise it will not be possible to delete any of the nested ctrls, and we will get an error from the shitty Word api
                            continue;
                        }
                        if (!toDelete.includes(ctrl.id))
                            continue;
                        const nested = ctrl.getContentControls();
                        nested.load(['id']);
                        await context.sync();
                        [ctrl, ...nested.items].forEach(c => unprotect(c));
                    }
                    for (const id of toDelete) {
                        const ctrl = context.document.getContentControls().getById(id);
                        if (!ctrl)
                            continue;
                        ctrl.select();
                        ctrl.delete(false);
                    }
                    await context.sync();
                });
                async function filterIds(ids) {
                    //!I got hard time to get this to work. Be careful before making any change.
                    //! We need to make sure that the array of ids of the ctrls to be deleted does not include the ids of any nested ctrl of any of the ids in the array. For example: if the array contains the id of ctrl x, the id of any ctrl nested within the range of ctrl x must be removed from the array
                    return await Word.run(async (context) => {
                        const ctrls = context.document.getContentControls();
                        ctrls.load(['id']);
                        await context.sync();
                        for (const id of ids) {
                            const ctrl = ctrls.getById(id);
                            const nested = ctrl.getContentControls();
                            nested.load('id');
                            await context.sync();
                            const nestedIds = nested.items.filter(c => c.id !== id).map(c => c.id);
                            ids = ids.filter(i => !nestedIds.includes(i)); //!we remove any nested ctrls from the toDelete array
                        }
                        return ids;
                    });
                }
            }
            ;
            async function createNewDoc() {
                await Word.run(async (context) => {
                    return; //!Desactivating working with new document created from template until we find a solution to the context issue
                    const template = await getTemplate();
                    console.log(template);
                    if (!template)
                        return showNotification('Failed to create the template');
                    const newDoc = context.application.createDocument(template);
                    const all = newDoc.contentControls;
                    all.load(['title', 'tag']);
                    await newDoc.context.sync();
                    showNotification(`All ctrls from newDoc = : ${all.items.map(c => c.title).join(', ')}`);
                    all.items.map(ctrl => {
                        if (toDelete.includes(ctrl.id))
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
        async function promptForSelection(selectCtrls) {
            const blocks = [];
            const subOptions = async (id) => await promptForSelection(await getSubOptions(id, true));
            try {
                for (const ctrl of selectCtrls) {
                    if (escape(ctrl.id))
                        continue; //!We must escape the ctrls that have already been processed
                    ctrl.select();
                    if (ctrl.tag === RTDuplicateTag) {
                        await duplicateBlock(ctrl.id);
                        continue;
                    }
                    ;
                    const isLast = selectCtrls.indexOf(ctrl) === selectCtrls.length - 1; //We check if this is the last contentcontrol in the array
                    const block = await insertPromptBlock(ctrl.id, isLast);
                    if (!block)
                        continue;
                    blocks.push(block);
                    if (!block.wraper)
                        await subOptions(ctrl.id);
                    else if (!block.checkBox && block.btnNext) {
                        //!We excluded this case for the moment as a test
                        //!This is the case where selectCtrl has no "ctrlSi" contentControl as a direct child. We will await the user to click the button in order to process all the already displayed elements of selectCtrls[] until this point. Then, we will process the selectCtrl separetly before moving to the next selectCtrl in selectCtrls[]
                        await btnOnClick(blocks, block); //We must await the user to click the button in order to process all the already displayed elements/options of selectCtrls[].
                        await subOptions(ctrl.id); //!We select only the direct select ctrls children
                    }
                    else if (block.btnNext)
                        await btnOnClick(blocks, block); //This is the case where btnNext was added because we reached the end of selectCtrls[] (addBtn = true). We then need to await the user to click the button in order to process all the already displayed elements/options of selectCtrls[].
                }
            }
            catch (error) {
                return showNotification(`Error from showSelectPrompt() = ${error}`);
            }
        }
        /**
         *
         * @param id The id of the contentControl containig the options
         * @param isLast If true, a button will be appended at the end
         * @returns
         */
        async function insertPromptBlock(id, isLast) {
            try {
                return await showSelectUI();
            }
            catch (error) {
                return showNotification(`Error from insertPromptBlock() = ${error}`);
            }
            async function showSelectUI() {
                return await Word.run(async (context) => {
                    const ctrl = context.document.contentControls.getById(id);
                    ctrl.load(props);
                    const label = await labelRange(ctrl, RTSiTag);
                    await context.sync();
                    if (!label) {
                        //A select ctrl that does not have a direct RTSiTag nested ctrl, is a container for other selectCtrls that need to be displayed. They are meant to offer multiple options for the user to select only one of them. But this is not always the case
                        isSelected(id.toString()); //!We add it to the selectedCtrls array to avoid it being processed again.
                        return { wraper: undefined }; //!If this is not the last element in selectCtrls (addBtn == false) We will return a container with only a button to be clicked in order to move to the next select ctrl in the array
                        //return appendHTMLElements('');//!If this is not the last element in selectCtrls (addBtn == false) We will return a container with only a button to be clicked in order to move to the next select ctrl in the array
                        if (!isLast) {
                        }
                        else
                            return { wraper: undefined }; //!If this is the last element in selectCtrls array (addBtn == true), we will return a slectBlock with undefined container (which means that we will display the options nested within the select ctrl, but since this is the last ctrl in the selectCtrls array, we do not need to show a btnNext because when the nested options (i.e., the nested selectCtrls) will be displayed, a btnNext will be automatically inserted)
                    }
                    label.select();
                    const text = label.text || `The ctrl label was found but no text could be retrieved ! ctrl title = ${ctrl.title}`;
                    label.font.hidden = true;
                    await context.sync();
                    return appendHTMLElements(text, id.toString(), isLast); //The checkBox will have as id the title of the "select" contentcontrol}
                });
            }
        }
        function appendHTMLElements(text, id, isLast = false) {
            const wraper = createHTMLElement('div', 'promptContainer', '', USERFORM);
            if (!id)
                return { wraper, btnNext: btn() }; //!We return a container with a button with no checkBox
            const option = createHTMLElement('div', 'select', '', wraper);
            const checkBox = createHTMLElement('input', 'checkBox', '', option, id); //!We must give the checkBox the id of the selectCtrl because the id will be later used to retrieve the selectCtrl and process its children
            checkBox.type = 'checkbox';
            if (processed.includes(id))
                checkBox.checked = true; //!Normaly this should never happen
            createHTMLElement('label', 'label', text, option);
            if (!isLast)
                return { wraper, checkBox };
            return { wraper, checkBox, btnNext: btn() };
            function btn() {
                const btns = createHTMLElement('div', 'btns', '', wraper);
                return createHTMLElement('button', 'btnOK', 'Next', btns);
            }
        }
        function btnOnClick(blocks, block) {
            return new Promise((resolve, reject) => {
                !block.btnNext ? resolve(processed) : block.btnNext.onclick = () => processBlocks();
                btnDelete(); //We append a button which will delete all the unselected contentecontrols
                async function processBlocks(deleteSelected = false) {
                    const checkBoxes = blocks
                        .filter(block => block.checkBox)
                        //@ts-ignore
                        .map(block => [block.checkBox.id, block.checkBox.checked]);
                    blocks.forEach(block => block.wraper.remove()); //We remove all the containers from the DOM
                    for (const [id, checked] of checkBoxes) {
                        const subOptions = await getSubOptions(Number(id), checked);
                        if (checked)
                            await isSelected(id, subOptions);
                        else
                            isNotSelected(id, subOptions);
                    }
                    if (deleteSelected)
                        await deleteUnselected();
                    resolve(processed);
                }
                ;
                function btnDelete() {
                    const btn = createHTMLElement('button', '', 'Delete Unselected', block.wraper);
                }
            });
        }
        async function isSelected(id, subOptions) {
            id = `${escapePrefix}${id}`;
            processed.push(id);
            if (subOptions)
                await promptForSelection(subOptions);
        }
        ;
        function isNotSelected(id, subOptions) {
            processed.push(id.toString());
            subOptions
                .forEach(ctrl => isSelected(ctrl.id.toString())); //!We are adding the "keep" prefix to the ids of the subOptions ctrls on purpose. This is because the parent ctrl will be deleted given that its id is added without the prefix. Hence all its children will (i.e., the subOptions) will be deleted as well with the parent ctrl. Adding the ids of each children, we will unnecesarily burden the list with a great number of ids, that will in all cases be deleted. 
            console.log(processed);
        }
        ;
        async function getSubOptions(id, directChildren, children) {
            if (!children)
                children = await getChildren();
            if (!directChildren)
                return getSelectCtrls(children);
            return getSelectCtrls(children).filter(c => c.parentContentControl.id === id); //!We need to make sure we get only the direct children of the ctrl and not all the nested ctrls
            async function getChildren() {
                return Word.run(async (context) => {
                    const ctrl = context.document.getContentControls().getById(id);
                    const children = ctrl.getContentControls();
                    children.load([...props, 'parentContentControl']);
                    await context.sync();
                    return children.items.filter(c => c.id !== id); //!We must exclude the ctrl itself which in some cases may be returned as part of its children due to a bug in Word API
                });
            }
        }
        /**
         *
         * @param id
         */
        async function duplicateBlock(id) {
            const replace = Word.InsertLocation.replace;
            const after = Word.InsertLocation.after;
            try {
                await insertClones(id);
            }
            catch (error) {
                showNotification(`${error}`);
            }
            async function insertClones(id) {
                await Word.run(async (context) => {
                    const ctrl = context.document.contentControls.getById(id);
                    ctrl.load(props);
                    const label = await labelRange(ctrl, RTSectionTag);
                    if (!label)
                        return;
                    await context.sync();
                    if (!label.text)
                        return showNotification("No lable text");
                    ctrl.select();
                    const message = `Combien de ${label.text} y'a-t-il ?`;
                    let answer = Number(await promptForInput(message, '1'));
                    if (isNaN(answer)) {
                        showNotification(`The provided text cannot be converted into a number: ${answer}`);
                        return await insertClones(id);
                    }
                    else if (answer < 1)
                        return isNotSelected(id, await getSubOptions(id, false));
                    const title = `${getCtrlTitle(ctrl.tag, id)}-Cloned ${answer}`;
                    ctrl.title = title; //!We must update the title in case it is no matching the id in the template.
                    const ctrlContent = ctrl.getOoxml();
                    await context.sync();
                    for (let i = 1; i < answer; i++)
                        ctrl.getRange().insertOoxml(ctrlContent.value, after);
                    const clones = ctrl.context.document.getContentControls().getByTitle(title);
                    clones.load(props);
                    label.font.hidden = true;
                    await context.sync();
                    const items = clones.items; //!clones.items.entries() caused the for loop to fail in scriptLab. The reason is unknown
                    try {
                        for (const clone of items)
                            await processClone(clone.id, items.indexOf(clone) + 1);
                    }
                    catch (error) {
                        showNotification(`Error from processClone() = ${error}`);
                    }
                });
            }
            async function processClone(id, i) {
                await Word.run(async (context) => {
                    const clone = context.document.contentControls.getById(id);
                    clone.load(props);
                    const label = await labelRange(clone, RTSectionTag);
                    await context.sync();
                    if (!label)
                        return;
                    clone.title = `${getCtrlTitle(clone.tag, clone.id)}-${i}`;
                    const text = `${label.text} ${i}`;
                    label.insertText(text, replace);
                    label.font.hidden = true;
                    await context.sync();
                    const div = createHTMLElement('div', '', text, USERFORM, '', false);
                    await promptForSelection(await getSubOptions(clone.id, true)); //!We select only the direct select ctrls children
                    div.remove();
                });
            }
            ;
        }
        async function labelRange(parent, tag) {
            const ctrl = getFirstByTag(parent, tag);
            ctrl.load(['id', 'parentContentControl']);
            await parent.context.sync();
            if (ctrl.parentContentControl.id !== parent.id)
                return undefined; //!The label ctrl must be a direct child of the parent ctrl
            const range = ctrl.getRange('Content');
            ctrl.cannotEdit = false;
            range.font.hidden = false;
            range.load(['text']);
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
        async function showNestedOptionsTree() {
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
            const subOptions = await getSubOptions(ctrl.id, true);
            await promptForSelection(subOptions);
            prepareTemplate();
            function failed(message) {
                showNotification(message);
                prepareTemplate();
            }
        }
    }
    ;
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
    async finalizeContract() {
        const remove = [this.RTSiTag, this.RTDescriptionTag, this.RTObsTag, this.RTSectionTag]; //The contentcontrol and its content will be deleted.
        const content = [this.RTSelectTag, this.RTCloneTag]; //We will delete the contentControl but keep its content
        const styles = [...this.RTSiStyles, this.RTSectionStyle, this.RTObsStyle, this.RTDescriptionStyle];
        await Word.run(async (context) => {
            const allCtrls = context.document.getContentControls();
            allCtrls.load(['tag', 'title']);
            await context.sync();
            const ids = allCtrls.items.map(c => c.id);
            for (const id of ids) {
                const ctrls = context.document.getContentControls();
                ctrls.load('id');
                await context.sync();
                const ctrl = ctrls.items.find(c => c.id === id);
                if (!ctrl)
                    return;
                ctrl.load(['tag', 'title']);
                await context.sync();
                ctrl.cannotDelete = false; //!We must remove the cannotDelete from all ctrls because this will prevent deleting the line on which there is a hidden contentcontrol
                if (!ctrl?.tag)
                    return;
                if (remove.includes(ctrl.tag))
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
        insertBtn([() => this.showInputs(), 'Edit The FILLIN Fields', 'Shows the interface to edit the FILLIN fiels in the document'], true);
        insertBtn([() => this.insertNewFILLINField(), 'Insert new FILLIN filed', 'Inserts a new FILLIN field in the selected range. It replaces the selected text with the FILLIN field'], true);
    }
    async showInputs() {
        USERFORM.innerHTML = '';
        await Word.run(async (context) => {
            const fields = context.document.body.fields;
            fields.load(["code", "result"]);
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
                const l = document.createElement('label');
                const input = document.createElement('input');
                input.id = `FILLIN_${index.toString()}`;
                input.value = field.result.text;
                const div = document.createElement('div');
                div.append(l, input);
                USERFORM.appendChild(div);
                l.textContent = lable;
                //input.onchange = () => this.editField(index, input);
                return [input, index];
            }).filter(item => item !== undefined);
            insertBtn([() => this.editAllFields(inputs), 'Update All Fileds From Inputs', 'Parses the values of the inputs, and updates the corresponding fields'], true);
        });
        insertBtn(goHome, false); //We insert the goHome navigation button on top of all the inputs
    }
    async editAllFields(inputs) {
        await Word.run(async (context) => {
            const fields = context.document.body.fields;
            fields.load(['code', 'result']);
            await context.sync();
            for (const [input, index] of inputs) {
                if (!input.value)
                    continue;
                const field = fields.items[index];
                if (!field)
                    return console.log('field not found');
                field.result.insertText(input.value, Word.InsertLocation.replace);
                await context.sync();
                console.log('Modified field = ' + field.code);
            }
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