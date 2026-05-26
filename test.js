"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
async function importFile(file) {
    const response = await fetch(file);
    const data = await response.json();
    const jsContent = await fetch(data.download_url).then(r => r.text());
    const dataUrl = `data:text/javascript,${encodeURIComponent(jsContent)}`;
    const module = await import(dataUrl);
    console.log(module);
}
importFile("https://api.github.com/repos/oliviax727/research-log-api/contents/index.js");
//# sourceMappingURL=test.js.map