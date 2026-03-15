import type { LogEntry } from "lazy-ms-graph-mcp-shared";

export function log(level: LogEntry["level"], message: string): void {
  const entry: LogEntry = {
    timestamp: new Date().toISOString(),
    level,
    message,
  };
  // Use stderr so we don't interfere with MCP stdio transport
  process.stderr.write(JSON.stringify(entry) + "\n");

  // If spawned by Electron (IPC available), also send via IPC
  if (process.send) {
    process.send({ type: "log", entry });
  }
}

export const logger = {
  info: (msg: string) => log("info", msg),
  warn: (msg: string) => log("warn", msg),
  error: (msg: string) => log("error", msg),
  debug: (msg: string) => log("debug", msg),
};
