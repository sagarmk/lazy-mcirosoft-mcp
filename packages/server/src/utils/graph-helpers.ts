import type { Client } from "@microsoft/microsoft-graph-client";

interface PagedResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

export async function getAllPages<T>(
  client: Client,
  endpoint: string,
  maxPages: number = 10
): Promise<T[]> {
  const results: T[] = [];
  let url: string | undefined = endpoint;
  let page = 0;

  while (url && page < maxPages) {
    const response: PagedResponse<T> = await client.api(url).get();
    results.push(...response.value);
    url = response["@odata.nextLink"];
    page++;
  }

  return results;
}

export function formatGraphError(error: unknown): string {
  if (error instanceof Error) {
    const graphError = error as Error & { statusCode?: number; code?: string; body?: string };
    if (graphError.statusCode) {
      return `Graph API Error (${graphError.statusCode}): ${graphError.code || graphError.message}`;
    }
    return error.message;
  }
  return String(error);
}
