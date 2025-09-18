type ContentControl = Word.ContentControl;
type ContentControlType = Word.ContentControlType.richText | Word.ContentControlType.richTextInline | Word.ContentControlType.richTextParagraphs | Word.ContentControlType.comboBox | Word.ContentControlType.checkBox | Word.ContentControlType.dropDownList | Word.ContentControlType.datePicker | Word.ContentControlType.richTextTable | Word.ContentControlType.richTextTableCell;

type Btn = [Function, string];

type selectBlock = {container: HTMLDivElement, checkBox?: HTMLInputElement, btnNext?: HTMLButtonElement};