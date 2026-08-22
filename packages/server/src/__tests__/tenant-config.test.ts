import { test } from "node:test";
import assert from "node:assert/strict";
import { resolveTenants } from "../config/tenant-config.js";

const LEGACY_ENV = {
  MSGRAPH_CLIENT_ID: "legacy-client",
  MSGRAPH_CLIENT_SECRET: "legacy-secret",
  MSGRAPH_TENANT_ID: "legacy-tenant",
};

test("legacy single-tenant env vars become the 'default' profile", () => {
  const { tenants, defaultTenant } = resolveTenants({ ...LEGACY_ENV });
  assert.equal(tenants.length, 1);
  assert.deepEqual(tenants[0], {
    name: "default",
    clientId: "legacy-client",
    clientSecret: "legacy-secret",
    tenantId: "legacy-tenant",
  });
  assert.equal(defaultTenant, "default");
});

test("MICROSOFT_MCP_* fallback vars still work", () => {
  const { tenants } = resolveTenants({
    MICROSOFT_MCP_CLIENT_ID: "a",
    MICROSOFT_MCP_CLIENT_SECRET: "b",
    MICROSOFT_MCP_TENANT_ID: "c",
  });
  assert.equal(tenants.length, 1);
  assert.equal(tenants[0].name, "default");
});

test("no credentials at all yields zero tenants without throwing", () => {
  const { tenants, defaultTenant } = resolveTenants({});
  assert.equal(tenants.length, 0);
  assert.equal(defaultTenant, undefined);
});

test("MSGRAPH_TENANTS accepts a JSON array and lowercases names", () => {
  const env = {
    MSGRAPH_TENANTS: JSON.stringify([
      { name: "Contoso", clientId: "c1", clientSecret: "s1", tenantId: "t1" },
      { name: "fabrikam", clientId: "c2", clientSecret: "s2", tenantId: "t2" },
    ]),
  };
  const { tenants, defaultTenant } = resolveTenants(env);
  assert.deepEqual(tenants.map((t) => t.name), ["contoso", "fabrikam"]);
  assert.equal(defaultTenant, "contoso");
});

test("MSGRAPH_TENANTS accepts an object map keyed by name", () => {
  const env = {
    MSGRAPH_TENANTS: JSON.stringify({
      contoso: { clientId: "c1", clientSecret: "s1", tenantId: "t1" },
    }),
  };
  const { tenants } = resolveTenants(env);
  assert.equal(tenants.length, 1);
  assert.equal(tenants[0].name, "contoso");
  assert.equal(tenants[0].clientId, "c1");
});

test("invalid MSGRAPH_TENANTS JSON throws with the source named", () => {
  assert.throws(() => resolveTenants({ MSGRAPH_TENANTS: "{not json" }), /MSGRAPH_TENANTS is not valid JSON/);
});

test("MSGRAPH_TENANTS entry missing a field throws", () => {
  const env = {
    MSGRAPH_TENANTS: JSON.stringify([{ name: "contoso", clientId: "c1" }]),
  };
  assert.throws(() => resolveTenants(env), /missing: clientSecret, tenantId/);
});

test("per-name MSGRAPH_TENANT_<NAME>_* triples are collected", () => {
  const env = {
    MSGRAPH_TENANT_CONTOSO_CLIENT_ID: "c1",
    MSGRAPH_TENANT_CONTOSO_CLIENT_SECRET: "s1",
    MSGRAPH_TENANT_CONTOSO_TENANT_ID: "t1",
    MSGRAPH_TENANT_FABRIKAM_CLIENT_ID: "c2",
    MSGRAPH_TENANT_FABRIKAM_CLIENT_SECRET: "s2",
    MSGRAPH_TENANT_FABRIKAM_TENANT_ID: "t2",
  };
  const { tenants } = resolveTenants(env);
  assert.deepEqual(tenants.map((t) => t.name).sort(), ["contoso", "fabrikam"]);
});

test("an incomplete per-name triple is skipped with a warning", () => {
  const warnings: string[] = [];
  const env = {
    MSGRAPH_TENANT_CONTOSO_CLIENT_ID: "c1",
  };
  const { tenants } = resolveTenants(env, { warn: (m) => warnings.push(m) });
  assert.equal(tenants.length, 0);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /MSGRAPH_TENANT_CONTOSO_CLIENT_SECRET/);
});

test("plain MSGRAPH_TENANT_ID is not mistaken for a named tenant", () => {
  const { tenants } = resolveTenants({ MSGRAPH_TENANT_ID: "t1" });
  assert.equal(tenants.length, 0);
});

test("legacy vars and named tenants coexist", () => {
  const env = {
    ...LEGACY_ENV,
    MSGRAPH_TENANT_CONTOSO_CLIENT_ID: "c1",
    MSGRAPH_TENANT_CONTOSO_CLIENT_SECRET: "s1",
    MSGRAPH_TENANT_CONTOSO_TENANT_ID: "t1",
  };
  const { tenants, defaultTenant } = resolveTenants(env);
  assert.deepEqual(tenants.map((t) => t.name).sort(), ["contoso", "default"]);
  assert.equal(defaultTenant, "default");
});

test("inline JSON wins over a named-env tenant with the same name", () => {
  const warnings: string[] = [];
  const env = {
    MSGRAPH_TENANTS: JSON.stringify([{ name: "contoso", clientId: "json", clientSecret: "s", tenantId: "t" }]),
    MSGRAPH_TENANT_CONTOSO_CLIENT_ID: "env",
    MSGRAPH_TENANT_CONTOSO_CLIENT_SECRET: "env",
    MSGRAPH_TENANT_CONTOSO_TENANT_ID: "env",
  };
  const { tenants } = resolveTenants(env, { warn: (m) => warnings.push(m) });
  assert.equal(tenants.length, 1);
  assert.equal(tenants[0].clientId, "json");
  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /shadowed/);
});

test("MSGRAPH_DEFAULT_TENANT picks the default (case-insensitive)", () => {
  const env = {
    ...LEGACY_ENV,
    MSGRAPH_TENANT_CONTOSO_CLIENT_ID: "c1",
    MSGRAPH_TENANT_CONTOSO_CLIENT_SECRET: "s1",
    MSGRAPH_TENANT_CONTOSO_TENANT_ID: "t1",
    MSGRAPH_DEFAULT_TENANT: "Contoso",
  };
  const { defaultTenant } = resolveTenants(env);
  assert.equal(defaultTenant, "contoso");
});

test("unknown MSGRAPH_DEFAULT_TENANT throws and lists what exists", () => {
  assert.throws(() => resolveTenants({ ...LEGACY_ENV, MSGRAPH_DEFAULT_TENANT: "nope" }), /Available: default/);
});

test("MSGRAPH_TENANTS_FILE is read and parsed", () => {
  const file = JSON.stringify([{ name: "filetenant", clientId: "c", clientSecret: "s", tenantId: "t" }]);
  const { tenants } = resolveTenants(
    { MSGRAPH_TENANTS_FILE: "/fake/tenants.json" },
    { readFile: (path) => (path === "/fake/tenants.json" ? file : "") }
  );
  assert.equal(tenants.length, 1);
  assert.equal(tenants[0].name, "filetenant");
});

test("unreadable MSGRAPH_TENANTS_FILE throws with the path", () => {
  assert.throws(
    () =>
      resolveTenants(
        { MSGRAPH_TENANTS_FILE: "/missing.json" },
        {
          readFile: () => {
            throw new Error("ENOENT");
          },
        }
      ),
    /Cannot read MSGRAPH_TENANTS_FILE '\/missing\.json'/
  );
});
