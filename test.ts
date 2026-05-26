async function importFile(file: string): Promise<void> {
    const response = await fetch(file);
    const data = await response.json();

    const jsContent = await fetch(data.download_url).then(r => r.text());
    const dataUrl = `data:text/javascript,${encodeURIComponent(jsContent)}`;

    const module = await import(dataUrl);
    console.log(jsContent);
}

importFile("https://api.github.com/repos/oliviax727/research-log-api/contents/index.js");
