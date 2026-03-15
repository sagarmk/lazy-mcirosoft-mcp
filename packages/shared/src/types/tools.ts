export type ToolCategory =
  | "users"
  | "mail"
  | "calendar"
  | "contacts"
  | "files"
  | "teams"
  | "sharepoint"
  | "planner"
  | "onenote"
  | "groups";

export interface ToolParameter {
  name: string;
  type: string;
  description: string;
  required: boolean;
}

export interface ToolInfo {
  name: string;
  category: ToolCategory;
  description: string;
  keywords: string[];
  parameters: ToolParameter[];
}

export interface ToolSearchResult {
  tools: ToolInfo[];
  total: number;
  query: string;
}

export interface ToolExecuteRequest {
  tool_name: string;
  parameters: Record<string, unknown>;
}
