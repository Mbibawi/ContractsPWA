type RichText = Word.ContentControl & { type: Word.ContentControlType.richText };

type contentControl = {id:number, title?:string, tag?:string, delete?:boolean, content?:string|null}