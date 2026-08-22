import type { Client } from "@microsoft/microsoft-graph-client";
import type { TenantConfig } from "lazy-ms-graph-mcp-shared";
import { createMsalClient } from "./msal-client.js";
import { createGraphClient } from "./graph-client.js";

export interface TenantSummary {
  name: string;
  tenantId: string;
  clientId: string;
  isDefault: boolean;
  /** When to use this tenant — surfaced to the AI via tenants_list. */
  description?: string;
}

interface TenantEntry {
  config: TenantConfig;
  client: Client | null;
}

/**
 * Holds every configured Entra ID tenant profile and lazily builds one
 * MSAL + Graph client pair per tenant on first use. Lookups are
 * case-insensitive; the default tenant can be changed at runtime
 * (session-scoped, never persisted).
 */
export class TenantManager {
  private entries = new Map<string, TenantEntry>();
  private defaultName: string | null = null;

  constructor(tenants: TenantConfig[], defaultTenant?: string) {
    for (const tenant of tenants) {
      const name = tenant.name.trim().toLowerCase();
      if (!name) throw new Error("Tenant profile has an empty name");
      if (this.entries.has(name)) throw new Error(`Duplicate tenant profile '${name}'`);
      this.entries.set(name, { config: { ...tenant, name }, client: null });
    }

    if (defaultTenant) {
      this.setDefault(defaultTenant);
    } else if (this.entries.size > 0) {
      this.defaultName = this.entries.keys().next().value ?? null;
    }
  }

  size(): number {
    return this.entries.size;
  }

  names(): string[] {
    return Array.from(this.entries.keys());
  }

  has(name: string): boolean {
    return this.entries.has(name.trim().toLowerCase());
  }

  getDefaultName(): string | null {
    return this.defaultName;
  }

  list(): TenantSummary[] {
    return Array.from(this.entries.values()).map(({ config }) => ({
      name: config.name,
      tenantId: config.tenantId,
      clientId: config.clientId,
      isDefault: config.name === this.defaultName,
      ...(config.description ? { description: config.description } : {}),
    }));
  }

  setDefault(name: string): TenantSummary {
    const key = name.trim().toLowerCase();
    const entry = this.entries.get(key);
    if (!entry) throw new Error(this.unknownTenantMessage(key));
    this.defaultName = key;
    return {
      name: entry.config.name,
      tenantId: entry.config.tenantId,
      clientId: entry.config.clientId,
      isDefault: true,
      ...(entry.config.description ? { description: entry.config.description } : {}),
    };
  }

  /** Graph client for the named tenant, or the default tenant when omitted. */
  getClient(name?: string): Client {
    const key = name?.trim().toLowerCase() || this.defaultName;
    if (!key) {
      throw new Error(
        "No Entra ID tenants configured. Set MSGRAPH_CLIENT_ID/MSGRAPH_CLIENT_SECRET/MSGRAPH_TENANT_ID, or MSGRAPH_TENANTS for multiple tenants."
      );
    }
    const entry = this.entries.get(key);
    if (!entry) throw new Error(this.unknownTenantMessage(key));
    if (!entry.client) {
      const msal = createMsalClient(entry.config);
      entry.client = createGraphClient(msal);
    }
    return entry.client;
  }

  private unknownTenantMessage(name: string): string {
    const available = this.names().join(", ") || "(none)";
    return `Unknown tenant '${name}'. Available tenants: ${available}`;
  }
}
