/// <reference types="./office.d.ts" />

type ContentControl = Word.ContentControl;
type ContentControlType = Word.ContentControlType.richText  | Word.ContentControlType.comboBox | Word.ContentControlType.checkBox | Word.ContentControlType.dropDownList | Word.ContentControlType.comboBox

type Btn = [Function, string, string | undefined];

type selectBlock = { wraper?: HTMLDivElement, checkBox?: { chkbox: HTMLInputElement, ctrl: ContentControl }, btnNext?: HTMLButtonElement };