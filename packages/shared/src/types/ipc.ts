export type ServerStatus = "stopped" | "starting" | "running" | "error";

export interface ServerStatusInfo {
  status: ServerStatus;
  uptime?: number;
  pid?: number;
  error?: string;
}

export interface LogEntry {
  timestamp: string;
  level: "info" | "warn" | "error" | "debug";
  message: string;
}

export const IPC_CHANNELS = {
  SERVER_START: "lazy-mcp:server:start",
  SERVER_STOP: "lazy-mcp:server:stop",
  SERVER_STATUS: "lazy-mcp:server:status",
  SERVER_STATUS_CHANGED: "lazy-mcp:server:status-changed",
  SERVER_LOGS: "lazy-mcp:server:logs",
  SERVER_LOG_ENTRY: "lazy-mcp:server:log-entry",
  CONFIG_GET: "lazy-mcp:config:get",
  CONFIG_SET: "lazy-mcp:config:set",
  CONFIG_VALIDATE: "lazy-mcp:config:validate",
  AUTO_START_GET: "lazy-mcp:app:auto-start:get",
  AUTO_START_SET: "lazy-mcp:app:auto-start:set",
  SETUP_GUIDE_GET: "lazy-mcp:app:setup-guide",
  SERVER_PATH_GET: "lazy-mcp:app:server-path",
  INTEGRATION_INSTALL: "lazy-mcp:app:integration-install",
} as const;
