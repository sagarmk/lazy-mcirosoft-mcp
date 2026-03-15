import type { ToolCategory } from "lazy-ms-graph-mcp-shared";
import type { ToolDefinition } from "./tool-definition.js";
import { searchTools } from "./search-engine.js";

export class ToolRegistry {
  private tools = new Map<string, ToolDefinition>();

  register(tool: ToolDefinition): void {
    if (this.tools.has(tool.name)) {
      throw new Error(`Tool '${tool.name}' is already registered`);
    }
    this.tools.set(tool.name, tool);
  }

  get(name: string): ToolDefinition | undefined {
    return this.tools.get(name);
  }

  search(query: string, options?: { category?: ToolCategory; limit?: number }): ToolDefinition[] {
    return searchTools(Array.from(this.tools.values()), query, options);
  }

  getCategories(): ToolCategory[] {
    const categories = new Set<ToolCategory>();
    for (const tool of this.tools.values()) {
      categories.add(tool.category);
    }
    return Array.from(categories);
  }

  getAll(): ToolDefinition[] {
    return Array.from(this.tools.values());
  }

  size(): number {
    return this.tools.size;
  }
}
