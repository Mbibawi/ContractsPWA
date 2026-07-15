/// <reference types="./office.d.ts" />

type ContentControl = Word.ContentControl;
type ContentControlType = Word.ContentControlType.richText | Word.ContentControlType.comboBox | Word.ContentControlType.checkBox | Word.ContentControlType.dropDownList | Word.ContentControlType.comboBox

type Btn = [Function, string, string | undefined];

type promptBlock = {
    wraper: HTMLDivElement,
    checkBox: HTMLInputElement,
    ctrl: selectCtrl,
    btnNext?: HTMLButtonElement
};

type selectCtrl = {
    id: number;
    tag: string;
    title: string;
    parent: number | undefined;
    processed: boolean;
    delete: boolean;
    nested: { id: number, tag: string }[],
    hasLabel: { id: number, tag: string } | undefined
}