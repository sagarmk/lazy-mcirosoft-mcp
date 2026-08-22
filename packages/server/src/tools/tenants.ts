import { z } from "zod";
import type { ToolRegistry } from "../registry/tool-registry.js";
import type { TenantManager } from "../auth/tenant-manager.js";

// These tools operate on the server's tenant profiles rather than the Graph
// API, so their handlers close over the TenantManager and ignore the Graph
// client argument (execute_tool passes null for this category when no
// credentials are configured).
export function registerTenantTools(registry: ToolRegistry, manager: TenantManager): void {
  registry.register({
    name: "tenants_list",
    category: "tenants",
    description:
      "List all configured Entra ID (Azure AD) tenant profiles. Shows each profile's name, description (when to use it), tenant ID, client ID, and which one is the default. Pass a profile name as the 'tenant' argument of execute_tool to run any tool against that tenant.",
    keywords: ["tenants", "tenant", "entra", "azure", "directory", "list", "profiles", "accounts", "organizations"],
    parameters: z.object({}),
    handler: async () => {
      const tenants = manager.list();
      if (tenants.length === 0) {
        return {
          tenants: [],
          total: 0,
          hint: "No tenants configured. Set MSGRAPH_CLIENT_ID/MSGRAPH_CLIENT_SECRET/MSGRAPH_TENANT_ID for one tenant, or MSGRAPH_TENANTS (JSON) / MSGRAPH_TENANT_<NAME>_* variables for several.",
        };
      }
      return { tenants, total: tenants.length, default: manager.getDefaultName() };
    },
  });

  registry.register({
    name: "tenants_get_current",
    category: "tenants",
    description: "Show the default Entra ID tenant profile — the one used when execute_tool is called without a 'tenant' argument.",
    keywords: ["tenants", "tenant", "current", "default", "active", "entra", "azure", "which"],
    parameters: z.object({}),
    handler: async () => {
      const name = manager.getDefaultName();
      if (!name) {
        return { error: "No tenants configured." };
      }
      return manager.list().find((t) => t.isDefault);
    },
  });

  registry.register({
    name: "tenants_switch",
    category: "tenants",
    description:
      "Switch the default Entra ID tenant profile for this session. Subsequent execute_tool calls without a 'tenant' argument will use it. Does not persist across server restarts.",
    keywords: ["tenants", "tenant", "switch", "change", "select", "default", "entra", "azure", "set"],
    parameters: z.object({
      name: z.string().describe("Name of the tenant profile to make the default (see tenants_list)."),
    }),
    handler: async (_client, params: { name: string }) => {
      const summary = manager.setDefault(params.name);
      return { switched: true, default: summary };
    },
  });
}
