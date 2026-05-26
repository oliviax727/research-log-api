"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
console.log("Hello world!");
async function fetchFile(file) {
    const response = await fetch(file);
    const data = await response.json();
    const response2 = await fetch(data.download_url);
    const text = await response2.text();
    console.log("Content of index.ts:");
    console.log(text);
}
async function importFile(file) {
    const module = await import(file);
    console.log(module);
}
importFile("https://api.github.com/repos/oliviax727/research-log-api/contents/index.ts");
//# sourceMappingURL=test.js.map