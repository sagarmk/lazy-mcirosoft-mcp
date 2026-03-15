import type { z } from "zod";
import type { Client } from "@microsoft/microsoft-graph-client";
import type { ToolCategory, ToolInfo, ToolParameter } from "lazy-ms-graph-mcp-shared";

export interface ToolDefinition {
  name: string;
  category: ToolCategory;
  description: string;
  keywords: string[];
  parameters: z.ZodObject<any>;
  handler: (client: Client, params: any) => Promise<unknown>;
}

export function toolToInfo(tool: ToolDefinition): ToolInfo {
  const shape = tool.parameters.shape;
  const params: ToolParameter[] = Object.entries(shape).map(([name, schema]) => {
    const zodSchema = schema as z.ZodTypeAny;
    return {
      name,
      type: getZodTypeName(zodSchema),
      description: zodSchema.description || "",
      required: !zodSchema.isOptional(),
    };
  });

  return {
    name: tool.name,
    category: tool.category,
    description: tool.description,
    keywords: tool.keywords,
    parameters: params,
  };
}

function getZodTypeName(schema: z.ZodTypeAny): string {
  const typeName = schema._def?.typeName;
  if (!typeName) return "unknown";

  const typeMap: Record<string, string> = {
    ZodString: "string",
    ZodNumber: "number",
    ZodBoolean: "boolean",
    ZodArray: "array",
    ZodObject: "object",
    ZodOptional: "optional",
    ZodEnum: "enum",
  };

  if (typeName === "ZodOptional") {
    const inner = (schema as z.ZodOptional<any>)._def.innerType;
    return getZodTypeName(inner);
  }

  return typeMap[typeName] || typeName;
}
