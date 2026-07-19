/// <reference types="./types.d.ts" />
const version = "v11.16.7";
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
    element('p', 'notification', message, NOTIFICATION, '', true);
}
function showAlert(message) {
    showNotification(`Displayed Alert: ${message}`);
    const { modal, window } = getModalContainer(USERFORM, 'Alert', 'alert', false);
    element('p', '', message, window, '', true);
    const btn = element('button', '', 'OK', window, '', true);
    btn.onclick = () => { modal.remove(); };
}
function element(tag, css, textContent, parent, id, append = true) {
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
async function promptForInput(questions, deflt, fun, parent) {
    if (!questions?.length)
        return [];
    const { modal, window } = getModalContainer(parent ?? USERFORM);
    const inputs = questions.map(question => {
        const div = element('div', '', '', window, undefined, true);
        element('label', 'question', question, div, undefined, true);
        const input = element('input', 'answer', deflt || '', div, undefined, true);
        if (deflt)
            input.value = deflt;
        return input;
    });
    const btns = element('div', 'btns', '', window);
    const btnOK = element('button', 'btnOK', 'OK', btns);
    const btnCancel = element('button', 'btnCancel', 'Cancel', btns);
    return new Promise((resolve, reject) => {
        btnCancel.onclick = () => reject(modal.remove());
        btnOK.onclick = () => {
            const answers = inputs.map(input => input.value || '');
            console.log('user answers = ', answers.join('|'));
            modal.remove();
            if (fun)
                fun(answers);
            answers.length === 1 ? resolve(answers[0]) : resolve(answers);
        };
    });
}
;
async function promptConfirm(question, fun) {
    if (!question)
        question = 'No question was provided !!!';
    const { modal, window } = getModalContainer(USERFORM, 'Confirm Deletion');
    const prompt = element('div', 'prompt', '', window);
    element('p', 'ask', question, prompt);
    const btns = element('div', 'btns', '', prompt);
    const btnOK = element('button', 'btnOK', 'OK', btns);
    const btnNo = element('button', 'btnCancel', 'NO', btns);
    return new Promise((resolve, reject) => {
        btnOK.onclick = () => resolve(confirm(true));
        btnNo.onclick = () => resolve(confirm(false));
    });
    function confirm(confirm) {
        modal.remove();
        if (fun)
            fun(confirm);
        return confirm;
    }
}
function getModalContainer(parent, textContent, id, append = true) {
    const modal = element('div', 'modal', '', parent, id, append);
    if (textContent)
        element('p', '', textContent, modal, '', true);
    const window = element('div', 'modal-window', '', modal, '', append);
    return { modal, window };
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
        this.RTCloneTag = 'RTRepeat';
        this.RTSectionTag = 'RTSection'; //This tag is a contentcontrol which contains a text to be displayed (like a lable or a title) other than for choosing a specifc case (RTSi)
        this.RTSelectTag = 'RTSelect';
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
        try {
            range.select();
            const ctrl = range.insertContentControl(type);
            ctrl.load(['id', ...props.filter(prop => prop !== 'id')]);
            range.context.trackedObjects.add(ctrl); //!We must track the object before range.context.sync() is called otherwise it will be lost.
            await range.context.sync();
            console.log(`the newly created ContentControl id = ${ctrl.id} `);
            // Set properties for the new content control.
            ctrl.title = this.getCtrlTitle(title, ctrl.id);
            ctrl.tag = tag;
            ctrl.appearance = Word.ContentControlAppearance.boundingBox;
            if (placeHolder)
                ctrl.placeholderText = placeHolder;
            if (style)
                ctrl.getRange().style = style;
            ctrl.cannotDelete = cannotDelete;
            ctrl.cannotEdit = cannotEdit; //!This must come at the end after the style has been set.
            ctrl.select();
            await range.context.sync();
            showNotification(`Wrapped text in range ${index} with a content control.`);
            return ctrl;
        }
        catch (error) {
            showNotification(`There was an error while setting the properties of the newly crated contentcontrol by insertContentControl(): ${error.debugInfo || error}.`);
            return undefined;
        }
    }
    async wrapSelectionWithContentControl(title, tag, type, style, cannotEdit, cannotDelete) {
        const range = await this.getSelectionRange();
        if (!range)
            return;
        if (!style && this.RTSiStyles.includes(range.style))
            style = range.style || null;
        const ctrl = await this.insertContentControl(range, title, tag, 0, type, style, cannotEdit, cannotDelete);
        if (!ctrl)
            return;
        if (tag === this.RTCloneTag) {
            const style = this.RTSectionStyle;
            tag = this.RTSectionTag;
            await Word.run(range, async (context) => {
                const ctrlRange = ctrl.getRange('Content');
                ctrlRange.load(['paragraphs', 'paragraphs/style']);
                await context.sync();
                const p = ctrlRange.paragraphs.items.find(p => p.style === style);
                if (!p) {
                    ctrl.cannotDelete = false;
                    ctrl.delete(true);
                    await context.sync();
                    return showAlert('The Repeat ContentControl must have a RTSection ContentControl, we did not have any such styled paragraphs within the selected range.');
                }
                await this.insertContentControl(p, tag, tag, 1, this.richText, null, true, true);
            });
        }
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
            this.RTCloneTag,
            this.RTSectionTag,
            this.RTSelectTag,
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
        const searchString = this.searchString.bind(this), getSelectionRange = this.getSelectionRange.bind(this), insertContentControl = this.insertContentControl.bind(this), insertFields = this.insertFields.bind(this), setCtrlsColor = this.setCtrlsColor.bind(this), setCtrlsFontColor = this.setCtrlsFontColor.bind(this), insertAskField = this._fields.insertAskField.bind(this);
        const siTag = this.RTSiTag, selectTag = this.RTSelectTag, sectionTag = this.RTSectionTag, descTag = this.RTDescriptionTag, stylePrefix = this.StylePrefix, richText = this.richText, dorpDownTag = this.RTDropDownTag;
        const descStyle = this.RTDescriptionStyle, siStyle = this.RTSiStyles, sectionStyle = this.RTSectionStyle, dropDownList = this.dropDownList;
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
            [insertBlockAmountWithFILLINField, 'Insert Amount Block', 'Inserts a ContentControl block containing a FILLIN field associated with a bookmark for the amount in figures and in text'],
            [insertDropDownList, 'Insert a Dropdown List from selection', 'Creates a dropwdown list from the selected string. The options to choose from must be separated by "/"'],
            [() => insertRTDescription(true), 'Insert Single RT Description', single(this.RTDescriptionTag)],
            [this.insertSingleFiled, 'Insert ContentControl Field', single(this.RTFieldTag)],
            [this._fields.insertFIllINField, 'Insert FILLIN Field', single(this.RTFieldTag)],
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
                const container = element('div', '', '', undefined, id);
                USERFORM.insertAdjacentElement('beforebegin', container);
                const select = element('select', '', '', container);
                styles.forEach(style => {
                    const option = element('option', '', style.nameLocal.split(stylePrefix)[1], select);
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
            try {
                await Word.run(range, async (context) => {
                    //Wraping the range with ContentControl "RTSelect"
                    const _ctrl = await insertContentControl(range, selectTag, selectTag, undefined, richText, null, false, false, undefined, ['id']);
                    if (!_ctrl)
                        return showAlert('Failed to insert the RTSelect ContentControl');
                    const ctrl = context.document.contentControls.getById(_ctrl.id);
                    ctrl.load(['paragraphs', 'paragraphs/style']);
                    await context.sync();
                    const si = ctrl.paragraphs.items.find(p => siStyle.includes(p.style));
                    if (!si)
                        return showAlert('No paragraph styled with on of the "RTSi" styles was found in the selected range');
                    //Wraping the paragraph with ContentControl "RTSi"
                    await insertContentControl(si, siTag, siTag, undefined, richText, si.style, true, true);
                    [range, ctrl, si].forEach(obj => obj.untrack());
                    await context.sync();
                });
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
            const dropDownColor = '#991c63';
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
        async function insertBlockAmountWithFILLINField() {
            await Word.run(async (context) => {
                try {
                    const range = context.document.getSelection().getRange(Word.RangeLocation.start);
                    const wraper = await insertContentControl(range, 'Block Amount Associated to FILLINField', 'BlockAmount', 0, richText, null, false, false);
                    if (!wraper)
                        throw new Error('Failed to insert a contentControl in the selected range');
                    const wraperRange = wraper.getRange();
                    wraperRange.insertText('[*]', Word.RangeLocation.start);
                    const bookmarkName = await insertAskField(wraperRange.getRange(Word.RangeLocation.start));
                    if (!bookmarkName)
                        throw new Error('The bookmark name returned after trying to insert the ASK field is undefined');
                    getField(`{REF ${bookmarkName}} \\h \\*cardText`, wraper); //amout as text
                    wraperRange.insertText(' euros (', Word.InsertLocation.end);
                    getField(`REF ${bookmarkName} \\h`, wraper); // amount in figures
                    wraperRange.insertText(' €', Word.InsertLocation.end);
                    await context.sync();
                }
                catch (error) {
                    showAlert(`An error occured ${error.debugInfo || error}`);
                }
            });
            function getField(code, ctrl) {
                const field = ctrl.getRange().insertField(Word.InsertLocation.end, Word.FieldType.empty);
                if (!field)
                    throw new Error(`Failed to insert a field with field code = ${code}`);
                field.code = code;
                return field;
            }
            ;
        }
        /**
         * Wraps every occurrence of text formatted with a specific character style in a rich text content control.
         *
         * @param styles any array of styles that we will check that the range style is included in this array of styles.
         * @param title the title of the content control.
         * @param tag the tag of the content control.
         * @param cannotEdit true if the content control cannot be edited.
         * @param cannotDelete true if the content control cannot be deleted.
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
                const search = (await promptForInput([`Provide the search string. You can provide more than one string to search by separated by ${separator} witohout space`], separator))?.split(separator);
                const matchWildcards = await promptConfirm('Match Wild Cards');
                if (!styles)
                    styles = (await promptForInput([`Provide the styles that that need to be matched separated by ","`]))?.split(',').map(style => style.trim()) || [];
                return { search, matchWildcards };
            }
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
    }
    ;
    async customizeContract(showNested = false) {
        const RTClone = this.RTCloneTag, RTSiTag = this.RTSiTag, RTSectionTag = this.RTSectionTag, RTSelect = this.RTSelectTag;
        const getCtrlTitle = this.getCtrlTitle.bind(this), prepareTemplate = this.prepareTemplate.bind(this), finalizeContract = this.finalizeContract;
        const selectCtrls = [];
        await loopSelectCtrls();
        async function loopSelectCtrls() {
            await Word.run(async (context) => {
                if (showNested)
                    return await showNestedOptionsTree(context); //!This must come before selectCtrls is populated. Will populate it from the selection
                selectCtrls.push(...await fetchSelectCtrls(context));
                try {
                    for (const ctrl of selectCtrls)
                        await promptForSelection([ctrl], context);
                    USERFORM.innerHTML = '';
                    if (!await promptConfirm('Do you want to delete the unselected contentcontrols?\nWARNING: If you click NO, all the excluded options will be lost and you will have to start over again'))
                        return;
                    await deleteUnselected(context);
                    if (await promptConfirm('Finihsed deleting the unselected options.\nDo you want to finalize the contract by removing the other ContentControls like RTSi, RTDescription ?'))
                        await finalizeContract();
                }
                catch (error) {
                    showNotification(`Error from promptForSelection() = ${error}`);
                }
                ;
            });
        }
        async function labelRange(id, context) {
            const label = context.document.contentControls.getById(id);
            await context.sync();
            if (!label)
                return showAlert('The lable was not found, it was probably deleted at some point');
            label.cannotEdit = false; //!WARNING, we must unlock the cannotEdit before unhidding the font
            label.font.hidden = false; //!WARNING, this must come before range.load('text')
            label.load(['text']);
            label.font.hidden = true;
            label.cannotEdit = true;
            await context.sync();
            return label;
        }
        ;
        async function promptForSelection(ctrls, context, clear = true) {
            if (!ctrls?.length)
                return;
            try {
                await processCtrls();
            }
            catch (error) {
                return showNotification(`Error from showSelectPrompt() = ${error}`);
            }
            async function processCtrls() {
                if (clear)
                    USERFORM.innerHTML = ''; //We clear the form before populating it
                const blocks = [];
                for (const ctrl of ctrls) {
                    if (ctrl.processed)
                        continue; //!WE MUST escape the ctrls that have already been processed.
                    ctrl.processed = true;
                    if (!ctrl?.hasLabel)
                        await promptForSelection(subOptions(ctrl), context); //When a 'RTSelect' ContentControl  does not have a lable (which is a 'RTSi' or 'RTSection' ContentControl) it means that this ContentControl is a mere wraper for sub 'RTSelect' ContentControls, each representing an option from which the user must choose. Hence, we do not need to prompt the user to decide whether to keep or delete this select section 
                    else if (ctrl.tag === RTSelect && ctrl.hasLabel.tag === RTSectionTag) {
                        //When an RTSelect ContentControl has as label a 'RTSection' ContentControl, it means that  the RTSelect ContentControl is a wraper for sub 'RTSelect' ContentControls, but it has a lable that needs to be be displayed to the user to explain to him under which section the options are displayed.
                        await insertLabel(ctrl.hasLabel.id);
                        await promptForSelection(subOptions(ctrl), context, false);
                    }
                    else
                        await showPromptBlock(ctrl);
                }
                async function showPromptBlock(ctrl) {
                    const isLast = ctrls.indexOf(ctrl) === ctrls.length - 1; //We check if this is the last contentcontrol in the array
                    const block = await insertHtml(ctrl, isLast);
                    if (!block)
                        return;
                    blocks.push(block);
                    if (block.btnNext)
                        await btnOnClick(blocks, context); //This is the case where btnNext was added because we reached the end of ctrls[] (isLast = true). We then need to await the user to click the button in order to process all the already displayed html elements/options of ctrls[].
                    async function insertHtml(ctrl, isLast) {
                        const wraper = insertWraper(USERFORM);
                        const option = element('div', 'select', '', wraper);
                        const checkBox = element('input', 'checkBox', '', option); //!We must give the checkBox the id of the selectCtrl because the id will be later used to retrieve the selectCtrl and process its children
                        checkBox.type = 'checkbox';
                        const label = await insertLabel(ctrl.hasLabel.id, option);
                        if (!label)
                            return wraper.remove();
                        if (!isLast)
                            return { wraper, checkBox, ctrl };
                        return { wraper, checkBox, ctrl, btnNext: btn() };
                        function btn() {
                            const btns = element('div', 'btns', '', wraper);
                            return element('button', 'btnOK', 'Next', btns);
                        }
                    }
                    ;
                }
                ;
                function insertWraper(parent) {
                    return element('div', 'promptContainer', '', parent);
                }
                ;
                async function insertLabel(id, wraper) {
                    if (!wraper)
                        wraper = insertWraper(USERFORM);
                    const label = await labelRange(id, context);
                    if (!label)
                        return showAlert("The Label ContentControl could not be found");
                    const text = label.text || 'Label text could not be retrived';
                    return element('label', 'label', text, wraper);
                }
                ;
            }
        }
        /**
         *
         * @param id
         */
        async function cloneSelectBlock(ctrl, context) {
            const after = Word.InsertLocation.after;
            try {
                await insertClones(ctrl);
            }
            catch (error) {
                showNotification(`${error}`);
            }
            async function insertClones(ctrl) {
                if (ctrl?.hasLabel?.tag !== RTSectionTag)
                    return showAlert('The Clone does not have any RTSection label !'); //'RTSelect' ContentControls who are meant to be cloned, must have a direct ContentControl child having as tag 'RTSection' which contains the title of the block to be replicated/cloned (e.g. "Seller")
                const label = await labelRange(ctrl.hasLabel.id, context);
                if (!label?.text)
                    return showAlert("InsertClones() failed: The ContentControl to be replicated/cloned, must have a direct ContentControl child having as tag 'RTSection'. No such tag was found");
                const message = `Combien de ${label.text} y'a-t-il ?`;
                let answer = Number(await promptForInput([message], '1'));
                if (isNaN(answer)) {
                    showAlert(`The provided text cannot be converted into a number: ${answer}`);
                    return await insertClones(ctrl); //reprompting the user
                }
                else if (answer < 2)
                    return;
                else if (answer < 1)
                    return isNotSelected(ctrl);
                try {
                    const original = context.document.contentControls.getById(ctrl.id);
                    if (!original)
                        throw new Error('InsertClones() failed: We could not retrive the original ContentControl to be replicated.');
                    original.select();
                    original.title = `${getCtrlTitle(ctrl.tag, ctrl.id)}-Cloned ${answer} times`; //We give it a unique title by which we will retrieve the colnes that we will create.
                    const range = original.getRange();
                    const Ooxml = original.getOoxml();
                    await context.sync();
                    for (let i = 1; i < answer; i++)
                        range.insertOoxml(Ooxml.value, after);
                    const clones = context.document.contentControls.getByTitle(original.title);
                    clones.load('id');
                    await context.sync();
                    if (!clones?.items.length)
                        throw new Error('Failed to retrieve the clones');
                    const selectCtrlItems = (await fetchSelectCtrls(context, clones));
                    //!The newly inserted clones and their nested contentControls are not inlcuded in selectCtrls[], which means that subOptions() will never be able to retrieve them, and they will not be deleted or manipulated through the selectCtrls[]. So we need to add them to selectCtrls[]; 
                    const index = selectCtrls.indexOf(ctrl) + 1;
                    selectCtrlItems.reverse(); //We reverse it in order to insert the newClones right after the existing one in the right order
                    for (const selectCtrl of selectCtrlItems) {
                        if (selectCtrl.id === ctrl.id)
                            continue; //We escape the original block since it is already in selectCtrls[];
                        const nested = clones.items.find(c => c.id === selectCtrl.id).getContentControls();
                        const nestedSelectCtrls = await fetchSelectCtrls(context, nested); //we get the selectCtrl representation of their nested contentControls
                        nestedSelectCtrls.unshift(selectCtrl);
                        selectCtrls.splice(index, 0, ...nestedSelectCtrls);
                    }
                    ;
                    selectCtrlItems.reverse(); //We must reverse the array again
                    for (const clone of selectCtrlItems)
                        await processClone(clone, label.text, selectCtrlItems.indexOf(clone) + 1);
                }
                catch (error) {
                    showNotification(`Error from processClone() = ${error}`);
                }
            }
            async function processClone(clone, text, i) {
                clone.processed = true; //!IMPORTANT The newly inserted clones have never went through promptForSelection(). Their 'processed' prop has never been set to true. They will be processed again when passed to promptForSelection() although they have already been processed here. We need to set their processed prop to true in order to avoid this.
                if (clone.hasLabel?.tag !== RTSectionTag)
                    return showAlert('The clone does not have an RTSection label !'); //Normally this should never occur.
                const label = await labelRange(clone.hasLabel.id, context);
                if (!label)
                    throw new Error('Failed to retrive the label of the Clone'); //This should never occur.
                text = `${text}-${i}`;
                label.cannotEdit = false; //!WARNING, we must set cannotEdit to false before modifing the text, otherwise we will get an error.
                label.insertText(text, Word.InsertLocation.replace);
                label.cannotEdit = true;
                const ctrl = context.document.contentControls.getById(clone.id);
                ctrl.select();
                ctrl.title = `${getCtrlTitle(clone.tag, clone.id)}-${i}`;
                await context.sync();
                USERFORM.innerHTML = ''; //!We need to clear the userform html here because promptForselection() will not do it.
                element('div', '', text, USERFORM, '', true);
                await promptForSelection(subOptions(clone), context, false);
            }
            ;
        }
        async function isSelected(ctrl, context) {
            if (ctrl.tag === RTClone)
                await cloneSelectBlock(ctrl, context); //We need to prompt the user to decide if he wants to clone/copy this block
            else
                await promptForSelection(subOptions(ctrl), context, ctrl.tag === RTClone);
        }
        ;
        function isNotSelected(ctrl) {
            ctrl.delete = true;
            subOptions(ctrl).forEach(c => c.processed = true);
        }
        ;
        async function deleteUnselected(context) {
            try {
                await process();
                //await createNewDoc();
            }
            catch (error) {
                showNotification(`${error}`);
            }
            async function process() {
                const ctrls = context.document.contentControls;
                ctrls.load(['id']);
                await context.sync();
                const ids = selectCtrls.filter(c => c.delete).map(c => c.id);
                const toDelete = ctrls.items.filter(ctrl => ids.includes(ctrl.id));
                console.log(`toDelete ids = ${toDelete.map(ctrl => ctrl.id)}`);
                for (const ctrl of toDelete) {
                    try {
                        //if (ctrl.tag === RTDuplicateTag) continue;
                        const nested = ctrl.getContentControls();
                        nested.load('id');
                        await context.sync();
                        nested.items.forEach(c => c.cannotDelete = false);
                        ctrl.cannotDelete = false;
                        ctrl.delete(false);
                        await context.sync();
                    }
                    catch (error) {
                        console.log(`Error from deleting controls = ${error}. This is most probably caused by the contentcontrol (its id = ${ctrl.id}) having been already deleted with its parent`);
                    }
                }
            }
            ;
        }
        function subOptions(ctrl) {
            return ctrl.nested
                .map(child => selectCtrls.find(c => c.id === child.id)) //We get the full properties of each of ctrl children (children elements only contain the id and the tag)
                .filter(c => c !== undefined); //we remove undefined elements
        }
        async function fetchSelectCtrls(context, allRT) {
            const props = ['id', 'title', 'tag', 'parentContentControlOrNullObject/id', 'contentControls/id', 'contentControls/tag', 'contentControls/parentContentControl/id'];
            const labelTags = [RTSectionTag, RTSiTag];
            if (!allRT)
                allRT = context.document.getContentControls();
            allRT.load(props);
            await context.sync();
            const ctrls = getSelectCtrls(allRT.items)
                .map(ctrl => selectCtrl(ctrl));
            console.log(ctrls);
            return ctrls;
            function selectCtrl(ctrl) {
                return {
                    id: ctrl.id,
                    tag: ctrl.tag,
                    title: ctrl.title,
                    parent: ctrl.parentContentControlOrNullObject.id,
                    nested: getNested(ctrl),
                    hasLabel: hasLabel(ctrl),
                    processed: false,
                    delete: false,
                };
            }
            ;
            function getNested(ctrl) {
                //This is to cover the case of newly inserted clones, where the suboptions of the new clone is not already in the selectCtrls[];
                return getSelectCtrls(directChildren(ctrl))
                    .map(c => { return { id: c.id, tag: c.tag }; });
            }
            function hasLabel(ctrl) {
                const label = directChildren(ctrl)
                    .find(child => labelTags.includes(child.tag));
                if (!label)
                    return undefined;
                return { id: label.id, tag: label.tag };
            }
            function directChildren(ctrl) {
                return ctrl.contentControls.items
                    .filter(nested => nested.parentContentControl.id === ctrl.id); /*!we keep only one level of children*/
            }
            function getSelectCtrls(ctrls) {
                return ctrls.filter(ctrl => [RTSelect, RTClone].includes(ctrl.tag));
            }
        }
        ;
        function btnOnClick(blocks, context) {
            return new Promise((resolve) => {
                const btn = blocks.find(block => block.btnNext)?.btnNext;
                btn.onclick = async () => resolve(await processBlocks());
            });
            async function processBlocks() {
                const checkBoxes = blocks
                    .map(block => [block.ctrl, block.checkBox.checked]);
                blocks.forEach(block => block.wraper.remove()); //We remove all the containers from the DOM
                for (const [ctrl, checked] of checkBoxes) {
                    if (checked)
                        await isSelected(ctrl, context);
                    else
                        isNotSelected(ctrl);
                }
                return true;
            }
            ;
        }
        async function showNestedOptionsTree(context) {
            const ctrls = context.document.getSelection().contentControls;
            await context.sync();
            selectCtrls.length = 0;
            selectCtrls.push(...await fetchSelectCtrls(context, ctrls));
            if (!ctrls.items.length)
                return showAlert('You must select a range containing the RTSelect ContentContrls you want to show its tree'); //!This MUST COME AFTER selectCtrls.push() because ctrls.items are not available until fetchSelectCtrls() loads their properties.
            for (const ctrl of selectCtrls)
                await promptForSelection([ctrl], context);
            if (await promptConfirm('Do you want to delete the unselected contentcontrols?'))
                await deleteUnselected(context);
            prepareTemplate();
        }
    }
    async finalizeContract() {
        const remove = [this.RTSiTag, this.RTDescriptionTag, this.RTObsTag, this.RTSectionTag]; //The contentcontrol and its content will be deleted.
        const keepContent = [this.RTSelectTag, this.RTCloneTag]; //We will delete the contentControl but keep its content
        const styles = [...this.RTSiStyles, this.RTSectionStyle, this.RTObsStyle, this.RTDescriptionStyle];
        const hide = await promptConfirm('Do you want to hide the ContentControls that will not be deleted?');
        await Word.run(async (context) => {
            const allCtrls = context.document.getContentControls();
            allCtrls.load(['id', 'tag', 'title']);
            const body = context.document.body.getRange();
            body.load(['paragraphs', 'paragraphs/style']);
            await context.sync();
            try {
                for (const ctrl of allCtrls.items) {
                    ctrl.cannotDelete = false; //!We must remove the cannotDelete from all ctrls because this will prevent deleting the line on which there is a hidden contentcontrol
                    if (remove.includes(ctrl?.tag))
                        ctrl.delete(false);
                    else if (keepContent.includes(ctrl?.tag))
                        ctrl.delete(true);
                    else if (hide)
                        ctrl.appearance = Word.ContentControlAppearance.hidden;
                }
                body.paragraphs.items
                    .filter(p => styles.includes(p.style))
                    .forEach(p => {
                    p.style = `${this.StylePrefix}Normal`;
                    p.delete();
                });
                await context.sync();
            }
            catch (error) {
                showAlert(`Error while deleting contentcontrol:\n Error: ${error.debugInfo}`);
            }
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
        insertBtn([() => this.insertFIllINField(), 'Insert new FILLIN filed', 'Inserts a new FILLIN field in the selected range. It replaces the selected text with the FILLIN field'], true);
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
                const div = element('div', '', '', USERFORM, '', true);
                element('label', '', lable, div, '', true);
                const input = element('input', '', '', div, `FILLIN_${index.toString()}`, true);
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
    async insertAskField(range) {
        try {
            const type = 'ASK';
            const labels = ['Provide the name of the bookmark without spaces', 'Provide the user prompt'];
            const values = await this.insertField(labels, type, range);
            if (!values)
                throw new Error(`insretField() failed`);
            const { field, answers } = values;
            if (!field)
                throw new Error('The field is undefined');
            const bookmarkName = answers[0]?.replaceAll(' ', '') || undefined;
            if (!bookmarkName)
                throw new Error('The bookmark name returned is not valid !');
            field.code = `${type} "${bookmarkName}"`;
            field.updateResult();
            return { askField: field, bookmarkName };
        }
        catch (error) {
            showAlert(`insertAskField() failed with error: ${error.debugInfo || error}`);
        }
    }
    ;
    async insertFIllINField(range) {
        try {
            const type = 'FILLIN';
            const labels = ['Provide the FILLIN user prompt', 'Provide the FILLIN default value'];
            const values = (await this.insertField(labels, type, range));
            if (!values)
                throw new Error('Failed to insert the field');
            const { field, answers } = values;
            if (!field)
                throw new Error('insertField() failed. No field object was returned');
            const prompt = answers[0];
            field.code = `${type} "${prompt}"  \\d ${answers[1] || '[*]'}  \\* MERGEFORMAT`;
            field.updateResult();
            return field;
        }
        catch (error) {
            showAlert(`insertFILLINField() failed with error: ${error.debugInfo || error}`);
        }
    }
    ;
    async insertField(labels, type, range) {
        const create = element;
        const { modal, window } = getModalContainer(USERFORM, '', 'newField', false);
        return await showDialogue();
        async function showDialogue() {
            const answers = await promptForInput(labels, undefined, undefined, window);
            const btn = create('button', '', `Insert ${type} Field`, window, 'ok', true);
            return new Promise((resolve) => btn.onclick = () => resolve(onClick(answers)));
        }
        async function onClick(answers) {
            if (answers?.length)
                return undefined;
            const [answer1, answer2] = answers;
            if (!answer1 || !answer2)
                return undefined;
            return await Word.run(async (context) => {
                try {
                    if (!range)
                        range = context.document.getSelection().getRange(Word.RangeLocation.whole);
                    const field = range.insertField(Word.InsertLocation.start, Word.FieldType.empty);
                    modal.remove();
                    field.track();
                    await context.sync();
                    return { field, answers };
                }
                catch (error) {
                    console.log(error.debugInfo || error);
                }
            });
        }
        ;
    }
}
//# sourceMappingURL=app.js.map