export interface AuthConfig {
  clientId: string;
  clientSecret: string;
  tenantId: string;
}

export interface ServerConfig {
  port: number;
  transport: "stdio" | "sse";
  autoStart: boolean;
}

export interface AppConfig {
  auth: AuthConfig;
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
