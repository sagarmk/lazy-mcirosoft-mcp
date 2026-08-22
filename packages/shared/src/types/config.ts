export interface AuthConfig {
  clientId: string;
  clientSecret: string;
  tenantId: string;
}

/**
 * A named Entra ID (Azure AD) app registration. The server can hold several of
 * these at once; `name` is the profile key used to select one at execution time.
 */
export interface TenantConfig extends AuthConfig {
  name: string;
  /** Human/AI-facing note on when to use this tenant (shown in tenants_list). */
  description?: string;
}

export interface ServerConfig {
  port: number;
  transport: "stdio" | "sse";
  autoStart: boolean;
}

export interface AppConfig {
  auth: AuthConfig;
  /** Additional named Entra ID tenants beyond `auth` (which stays the default). */
  tenants?: TenantConfig[];
  /** Name of the tenant profile to use when a call does not specify one. */
  defaultTenant?: string;
  server: ServerConfig;
  app: {
    minimizeToTray: boolean;
    showNotifications: boolean;
  };
}

export const DEFAULT_CONFIG: AppConfig = {
  auth: {
    clientId: "",
    clientSecret: "",
    tenantId: "",
  },
  server: {
    port: 3100,
    transport: "stdio",
    autoStart: true,
  },
  app: {
    minimizeToTray: true,
    showNotifications: true,
  },
};
