import { test } from "node:test";
import assert from "node:assert/strict";
import { TenantManager } from "../auth/tenant-manager.js";
import { ToolRegistry } from "../registry/tool-registry.js";
import { registerTenantTools } from "../tools/tenants.js";

const TWO_TENANTS = [
  { name: "contoso", clientId: "c1", clientSecret: "s1", tenantId: "t1" },
  { name: "fabrikam", clientId: "c2", clientSecret: "s2", tenantId: "t2" },
];

test("first tenant is the default when none is specified", () => {
  const manager = new TenantManager(TWO_TENANTS);
  assert.equal(manager.getDefaultName(), "contoso");
  assert.deepEqual(manager.names(), ["contoso", "fabrikam"]);
  assert.equal(manager.size(), 2);
});

test("explicit default tenant is honored", () => {
  const manager = new TenantManager(TWO_TENANTS, "fabrikam");
  assert.equal(manager.getDefaultName(), "fabrikam");
});

test("list exposes ids but never the client secret", () => {
  const manager = new TenantManager(TWO_TENANTS);
  const listed = manager.list();
  assert.deepEqual(listed, [
    { name: "contoso", tenantId: "t1", clientId: "c1", isDefault: true },
    { name: "fabrikam", tenantId: "t2", clientId: "c2", isDefault: false },
  ]);
  for (const entry of listed) {
    assert.ok(!("clientSecret" in entry));
  }
});

test("setDefault switches and is case-insensitive", () => {
  const manager = new TenantManager(TWO_TENANTS);
  const summary = manager.setDefault("FABRIKAM");
  assert.equal(summary.name, "fabrikam");
  assert.equal(manager.getDefaultName(), "fabrikam");
});

test("setDefault on an unknown tenant throws and lists options", () => {
  const manager = new TenantManager(TWO_TENANTS);
  assert.throws(() => manager.setDefault("nope"), /Available tenants: contoso, fabrikam/);
});

test("duplicate tenant names are rejected", () => {
  assert.throws(() => new TenantManager([TWO_TENANTS[0], { ...TWO_TENANTS[1], name: "Contoso" }]), /Duplicate tenant/);
});

test("getClient with no tenants configured throws a setup hint", () => {
  const manager = new TenantManager([]);
  assert.throws(() => manager.getClient(), /No Entra ID tenants configured/);
});

test("getClient caches one client per tenant", () => {
  const manager = new TenantManager(TWO_TENANTS);
  const a1 = manager.getClient("contoso");
  const a2 = manager.getClient("CONTOSO");
  const b = manager.getClient("fabrikam");
  assert.equal(a1, a2);
  assert.notEqual(a1, b);
});

test("getClient without a name uses the current default", () => {
  const manager = new TenantManager(TWO_TENANTS);
  assert.equal(manager.getClient(), manager.getClient("contoso"));
  manager.setDefault("fabrikam");
  assert.equal(manager.getClient(), manager.getClient("fabrikam"));
});

test("getClient on an unknown tenant throws", () => {
  const manager = new TenantManager(TWO_TENANTS);
  assert.throws(() => manager.getClient("nope"), /Unknown tenant 'nope'/);
});

test("tenant tools register under the 'tenants' category", () => {
  const registry = new ToolRegistry();
  registerTenantTools(registry, new TenantManager(TWO_TENANTS));
  assert.deepEqual(registry.getCategories(), ["tenants"]);
  assert.ok(registry.get("tenants_list"));
  assert.ok(registry.get("tenants_get_current"));
  assert.ok(registry.get("tenants_switch"));
});

test("tenants_list / tenants_switch / tenants_get_current round-trip", async () => {
  const registry = new ToolRegistry();
  const manager = new TenantManager(TWO_TENANTS);
  registerTenantTools(registry, manager);
  const client = null as never;

  const listed = (await registry.get("tenants_list")!.handler(client, {})) as {
    tenants: Array<{ name: string }>;
    default: string;
  };
  assert.equal(listed.tenants.length, 2);
  assert.equal(listed.default, "contoso");

  const switched = (await registry.get("tenants_switch")!.handler(client, { name: "fabrikam" })) as {
    switched: boolean;
    default: { name: string };
  };
  assert.equal(switched.switched, true);
  assert.equal(switched.default.name, "fabrikam");

  const current = (await registry.get("tenants_get_current")!.handler(client, {})) as { name: string };
  assert.equal(current.name, "fabrikam");
});

test("tenants_list with nothing configured returns a setup hint instead of failing", async () => {
  const registry = new ToolRegistry();
  registerTenantTools(registry, new TenantManager([]));
  const result = (await registry.get("tenants_list")!.handler(null as never, {})) as { total: number; hint: string };
  assert.equal(result.total, 0);
  assert.match(result.hint, /MSGRAPH_TENANTS/);
});
