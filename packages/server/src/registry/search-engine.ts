import type { ToolDefinition } from "./tool-definition.js";
import type { ToolCategory } from "lazy-ms-graph-mcp-shared";

interface SearchOptions {
  category?: ToolCategory;
  limit?: number;
}

interface ScoredTool {
  tool: ToolDefinition;
  score: number;
}

export function searchTools(
  tools: ToolDefinition[],
  query: string,
  options: SearchOptions = {}
): ToolDefinition[] {
  const { category, limit = 10 } = options;
  const queryLower = query.toLowerCase();
  const tokens = queryLower.split(/\s+/).filter(Boolean);

  let candidates = tools;
  if (category) {
    candidates = candidates.filter((t) => t.category === category);
  }

  const scored: ScoredTool[] = candidates.map((tool) => {
    let score = 0;
    const nameLower = tool.name.toLowerCase();
    const descLower = tool.description.toLowerCase();

    // Exact name match
    if (nameLower === queryLower) score += 100;

    // Name contains full query
    if (nameLower.includes(queryLower)) score += 50;

    // Category match
    if (tool.category.toLowerCase() === queryLower) score += 40;

    for (const token of tokens) {
      // Name contains token
      if (nameLower.includes(token)) score += 30;

      // Category match per token
      if (tool.category.toLowerCase().includes(token)) score += 20;

      // Keyword match
      for (const kw of tool.keywords) {
        if (kw.toLowerCase().includes(token)) score += 15;
      }

      // Description match
      if (descLower.includes(token)) score += 10;

      // Parameter name match
      const shape = tool.parameters.shape;
      for (const paramName of Object.keys(shape)) {
        if (paramName.toLowerCase().includes(token)) score += 5;
      }
    }

    return { tool, score };
  });

  return scored
    .filter((s) => s.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, limit)
    .map((s) => s.tool);
}
