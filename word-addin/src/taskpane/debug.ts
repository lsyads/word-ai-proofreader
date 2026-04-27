/* global console, document */

type DebugLevel = "info" | "warn" | "error";

const MAX_DEBUG_LOG_ITEMS = 80;

export function appendDebugLog(
  level: DebugLevel,
  message: string,
  details?: Record<string, unknown>
) {
  writeBrowserConsole(level, message, details);

  const log = document.getElementById("debug-log");
  if (!log) {
    return;
  }

  const item = document.createElement("div");
  item.className = `debug-log-item is-${level}`;
  item.textContent = formatDebugMessage(level, message, details);
  log.appendChild(item);

  while (log.children.length > MAX_DEBUG_LOG_ITEMS) {
    log.removeChild(log.children[0]);
  }

  log.scrollTop = log.scrollHeight;
}

export function clearDebugLog() {
  const log = document.getElementById("debug-log");
  if (log) {
    log.innerHTML = "";
  }
}

function writeBrowserConsole(
  level: DebugLevel,
  message: string,
  details?: Record<string, unknown>
) {
  const payload = details || {};
  if (level === "error") {
    console.error(message, payload);
    return;
  }
  if (level === "warn") {
    console.warn(message, payload);
    return;
  }

  console.info(message, payload);
}

function formatDebugMessage(
  level: DebugLevel,
  message: string,
  details?: Record<string, unknown>
): string {
  const timestamp = new Date().toLocaleTimeString();
  const serializedDetails = details ? ` ${safeStringify(details)}` : "";
  return `${timestamp} [${level.toUpperCase()}] ${message}${serializedDetails}`;
}

function safeStringify(value: unknown): string {
  try {
    return JSON.stringify(value, normalizeUnknown, 2);
  } catch {
    return String(value);
  }
}

function normalizeUnknown(_key: string, value: unknown): unknown {
  if (value instanceof Error) {
    return {
      name: value.name,
      message: value.message,
      stack: value.stack,
    };
  }

  return value;
}
