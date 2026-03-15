import type { ToolRegistry } from "../registry/tool-registry.js";
import { registerUserTools } from "./users.js";
import { registerMailTools } from "./mail.js";
import { registerCalendarTools } from "./calendar.js";
import { registerContactTools } from "./contacts.js";
import { registerFileTools } from "./files.js";
import { registerTeamsTools } from "./teams.js";
import { registerSharePointTools } from "./sharepoint.js";
import { registerPlannerTools } from "./planner.js";
import { registerOneNoteTools } from "./onenote.js";
import { registerGroupTools } from "./groups.js";

export function registerAllTools(registry: ToolRegistry): void {
  registerUserTools(registry);
  registerMailTools(registry);
  registerCalendarTools(registry);
  registerContactTools(registry);
  registerFileTools(registry);
  registerTeamsTools(registry);
  registerSharePointTools(registry);
  registerPlannerTools(registry);
  registerOneNoteTools(registry);
  registerGroupTools(registry);
}
