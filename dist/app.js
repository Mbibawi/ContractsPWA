"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
//@ts-ignore
const Word = __importStar(require("@microsoft/office-js/word"));
//@ts-ignore
Office.onReady((info) => {
    // Check that we loaded into Word
    if (info.host === Office.HostType.Word) {
        const btnEditWord = document.getElementById("edit");
        if (btnEditWord)
            btnEditWord.ondblclick = () => sayHello('Contracts App Works');
        getRichTextContentControlTitles()
            .then(titels => console.log('Found titles = ', titels));
    }
});
function sayHello(sentence) {
    //@ts-ignore
    return Word.run((context) => {
        // insert a paragraph at the start of the document.
        const paragraph = context.document.body.insertParagraph(sentence, Word.InsertLocation.start);
        // sync the context to run the previous API call, and return.
        return context.sync();
    });
}
async function getRichTextContentControlTitles() {
    return Word.run(async (context) => {
        // 1. Grab the collection of all content controls in the document
        const allControls = context.document.contentControls;
        // 2. Queue up a load for each control’s title and type
        allControls.load("items/title,type");
        // 3. Execute the queued commands
        await context.sync();
        // 4. Filter to only Rich Text controls and collect their titles
        const titles = allControls.items
            .filter(cc => cc.type === Word.ContentControlType.richText)
            .map(cc => cc.title);
        // 5. (Optional) Log or return the titles
        console.log("Rich Text Content Control Titles:", titles);
        return titles;
    });
}
// Example invocation
getRichTextContentControlTitles()
    .then(titles => {
    // Do something with the titles array
})
    .catch(error => console.error(error));
//# sourceMappingURL=app.js.map