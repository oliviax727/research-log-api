console.log("Hello world!");

async function fetchFile(file: string): Promise<void> {
    const response = await fetch(file);
    const data = await response.json();

    const response2 = await fetch(data.download_url);
    const text = await response2.text();

    console.log("Content of index.ts:");
    console.log(text);
}

async function importFile(file: string): Promise<void> {
    const module = await import(file);
    console.log(module);
}

importFile("https://api.github.com/repos/oliviax727/research-log-api/contents/index.ts");