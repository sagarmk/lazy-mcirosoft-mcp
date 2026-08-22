import { readFileSync } from "fs";
import type { TenantConfig } from "lazy-ms-graph-mcp-shared";

export interface ResolvedTenants {
  tenants: TenantConfig[];
  defaultTenant?: string;
}

export interface ResolveOptions {
  /** Reads a tenants JSON file; injectable for tests. Defaults to fs.readFileSync. */
  readFile?: (path: string) => string;
  /** Receives non-fatal config warnings (skipped entries, shadowed duplicates). */
  warn?: (message: string) => void;
}

const NAMED_ENV_PATTERN = /^MSGRAPH_TENANT_(.+)_(CLIENT_ID|CLIENT_SECRET|TENANT_ID|DESCRIPTION)$/;

/**
 * Collect Entra ID tenant profiles from the environment. Sources, in
 * precedence order (first definition of a name wins):
 *
 *   1. MSGRAPH_TENANTS        — inline JSON: an array of
 *      {name, clientId, clientSecret, tenantId} or an object map of
 *      name -> {clientId, clientSecret, tenantId}
 *   2. MSGRAPH_TENANTS_FILE   — path to a JSON file with the same shape
 *   3. MSGRAPH_TENANT_<NAME>_CLIENT_ID / _CLIENT_SECRET / _TENANT_ID triples
 *   4. MSGRAPH_CLIENT_ID / MSGRAPH_CLIENT_SECRET / MSGRAPH_TENANT_ID
 *      (or the MICROSOFT_MCP_* fallbacks) — registered under the name "default"
 *
 * Names are case-insensitive and stored lowercase. Malformed explicit config
 * (bad JSON, missing fields in a JSON entry, unknown MSGRAPH_DEFAULT_TENANT)
 * throws; an incomplete named-env triple is skipped with a warning.
 */
export function resolveTenants(env: NodeJS.ProcessEnv, options: ResolveOptions = {}): ResolvedTenants {
  const readFile = options.readFile ?? ((path: string) => readFileSync(path, "utf-8"));
  const warn = options.warn ?? (() => {});

  const tenants: TenantConfig[] = [];
  const seen = new Set<string>();

  const add = (tenant: TenantConfig, source: string) => {
    const name = tenant.name.trim().toLowerCase();
    if (!name) {
      throw new Error(`Tenant entry from ${source} has an empty name`);
    }
    if (seen.has(name)) {
      warn(`Tenant '${name}' from ${source} is shadowed by an earlier definition; ignoring it`);
      return;
    }
    seen.add(name);
    tenants.push({ ...tenant, name });
  };

  const inlineJson = env.MSGRAPH_TENANTS;
  if (inlineJson) {
    for (const tenant of parseTenantsJson(inlineJson, "MSGRAPH_TENANTS")) {
      add(tenant, "MSGRAPH_TENANTS");
    }
  }

  const filePath = env.MSGRAPH_TENANTS_FILE;
  if (filePath) {
    let raw: string;
    try {
      raw = readFile(filePath);
    } catch (error) {
      throw new Error(
        `Cannot read MSGRAPH_TENANTS_FILE '${filePath}': ${error instanceof Error ? error.message : String(error)}`
      );
    }
    for (const tenant of parseTenantsJson(raw, `MSGRAPH_TENANTS_FILE (${filePath})`)) {
      add(tenant, "MSGRAPH_TENANTS_FILE");
    }
  }

  for (const tenant of collectNamedEnvTenants(env, warn)) {
    add(tenant, "environment variables");
  }

  const legacyClientId = env.MSGRAPH_CLIENT_ID || env.MICROSOFT_MCP_CLIENT_ID || "";
  const legacyClientSecret = env.MSGRAPH_CLIENT_SECRET || env.MICROSOFT_MCP_CLIENT_SECRET || "";
  const legacyTenantId = env.MSGRAPH_TENANT_ID || env.MICROSOFT_MCP_TENANT_ID || "";
  if (legacyClientId && legacyClientSecret && legacyTenantId) {
    add(
      { name: "default", clientId: legacyClientId, clientSecret: legacyClientSecret, tenantId: legacyTenantId },
      "MSGRAPH_CLIENT_ID/MSGRAPH_CLIENT_SECRET/MSGRAPH_TENANT_ID"
    );
  }

  let defaultTenant = env.MSGRAPH_DEFAULT_TENANT?.trim().toLowerCase() || undefined;
  if (defaultTenant && !seen.has(defaultTenant)) {
    throw new Error(
      `MSGRAPH_DEFAULT_TENANT is '${defaultTenant}' but no such tenant is configured. Available: ${
        tenants.map((t) => t.name).join(", ") || "(none)"
      }`
    );
  }
  if (!defaultTenant) {
    defaultTenant = seen.has("default") ? "default" : tenants[0]?.name;
  }

  return { tenants, defaultTenant };
}

function parseTenantsJson(raw: string, source: string): TenantConfig[] {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    throw new Error(`${source} is not valid JSON: ${error instanceof Error ? error.message : String(error)}`);
  }

  const entries: Array<{ name: unknown; value: unknown }> = Array.isArray(parsed)
    ? parsed.map((value) => ({ name: (value as Record<string, unknown>)?.name, value }))
    : parsed && typeof parsed === "object"
      ? Object.entries(parsed).map(([name, value]) => ({ name, value }))
      : [];

  if (entries.length === 0) {
    throw new Error(`${source} must be a non-empty JSON array or object map of tenants`);
  }

  return entries.map(({ name, value }, index) => {
    const record = (value ?? {}) as Record<string, unknown>;
    const tenant: TenantConfig = {
      name: typeof name === "string" ? name : "",
      clientId: typeof record.clientId === "string" ? record.clientId : "",
      clientSecret: typeof record.clientSecret === "string" ? record.clientSecret : "",
      tenantId: typeof record.tenantId === "string" ? record.tenantId : "",
      ...(typeof record.description === "string" && record.description ? { description: record.description } : {}),
    };
    const missing = (["name", "clientId", "clientSecret", "tenantId"] as const).filter((key) => !tenant[key]);
    if (missing.length > 0) {
      throw new Error(`${source} entry ${index + 1} is missing: ${missing.join(", ")}`);
    }
    return tenant;
  });
}

function collectNamedEnvTenants(env: NodeJS.ProcessEnv, warn: (message: string) => void): TenantConfig[] {
  const partial = new Map<string, Partial<TenantConfig>>();

  for (const [key, value] of Object.entries(env)) {
    const match = key.match(NAMED_ENV_PATTERN);
    if (!match || !value) continue;
    const name = match[1].toLowerCase();
    const entry = partial.get(name) ?? {};
    if (match[2] === "CLIENT_ID") entry.clientId = value;
    else if (match[2] === "CLIENT_SECRET") entry.clientSecret = value;
    else if (match[2] === "DESCRIPTION") entry.description = value;
    else entry.tenantId = value;
    partial.set(name, entry);
  }

  const tenants: TenantConfig[] = [];
  for (const [name, entry] of partial) {
    if (entry.clientId && entry.clientSecret && entry.tenantId) {
      tenants.push({
        name,
        clientId: entry.clientId,
        clientSecret: entry.clientSecret,
        tenantId: entry.tenantId,
        ...(entry.description ? { description: entry.description } : {}),
      });
    } else {
      const upper = name.toUpperCase();
      warn(
        `Tenant '${name}' is incomplete: set all of MSGRAPH_TENANT_${upper}_CLIENT_ID, MSGRAPH_TENANT_${upper}_CLIENT_SECRET, MSGRAPH_TENANT_${upper}_TENANT_ID`
      );
    }
  }
  return tenants;
}
