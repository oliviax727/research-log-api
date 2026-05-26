

async function importFile(file: string): Promise<void> {
    const response = await fetch(file);
    const data = await response.json();

    const module = await import(data.download_url);
    console.log(module);
}

importFile("https://api.github.com/repos/oliviax727/research-log-api/contents/index.ts");
