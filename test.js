async function importFile(fileUrl) {
    const response = await fetch(fileUrl);
    if (!response.ok) {
        throw new Error(`Fetch failed: ${response.status} ${response.statusText}`);
    }
    const data = (await response.json());
    const researchLogModule = (await import(data.download_url));
    console.log(researchLogModule);
}
importFile("https://api.github.com/repos/oliviax727/research-log-api/contents/index.ts");
export {};
//# sourceMappingURL=test.js.map