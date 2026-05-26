interface GitHubContent {
  download_url: string;
}

interface ResearchLogModule {
  ResearchLog: {
    runFunction: (workbook: any, functionString: string) => void;
    addEntry?: (...args: unknown[]) => unknown;
    shiftEntry?: (...args: unknown[]) => unknown;
    popEntry?: (...args: unknown[]) => unknown;
    clockIn?: (...args: unknown[]) => unknown;
    clockOut?: (...args: unknown[]) => unknown;
  };
}

async function importFile(fileUrl: string): Promise<void> {
  const response = await fetch(fileUrl);
  if (!response.ok) {
    throw new Error(`Fetch failed: ${response.status} ${response.statusText}`);
  }

  const data = (await response.json()) as GitHubContent;

  const researchLogModule = (await import(data.download_url)) as ResearchLogModule;
  console.log(researchLogModule);
}

importFile("https://api.github.com/repos/oliviax727/research-log-api/contents/index.ts");
