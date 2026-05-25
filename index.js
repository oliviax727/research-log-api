"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
console.log("Hello world!");
async function main() {
    const response = await fetch("https://api.github.com/repos/oliviax727/research-log-api/contents/test-module.ts");
    const data = await response.json();
    console.log(data.download_url);
    const response2 = await fetch(data.download_url);
    const text = await response2.text();
    console.log("Content of test-module.ts:");
    console.log(text);
}
main();
//# sourceMappingURL=index.js.map