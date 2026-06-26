// Launcher for the `mcp-obsidian` MCP server (runs `uvx mcp-obsidian`).
//
// Why this exists: mcp-obsidian crashes on startup if it cannot read
// OBSIDIAN_API_KEY (the token issued by the Obsidian "Local REST API" plugin).
// Some MCP hosts (e.g. certain opencode builds) don't forward env vars to the
// child process, so the server was marked "failed". This launcher injects the
// key from a gitignored .secret.json (or the env) and forwards the full env to uvx.
const { spawn } = require("child_process");
const fs = require("fs");
const path = require("path");

function loadKey() {
  if (process.env.OBSIDIAN_API_KEY) return process.env.OBSIDIAN_API_KEY;
  const secretPath = path.join(__dirname, ".secret.json");
  if (fs.existsSync(secretPath)) {
    try {
      return JSON.parse(fs.readFileSync(secretPath, "utf8")).obsidianApiKey;
    } catch (e) {
      console.error("obsidian-mcp: failed to parse .secret.json:", e.message);
    }
  }
  console.error("obsidian-mcp: OBSIDIAN_API_KEY not found.");
  console.error('  Set the env var, or create .secret.json with { "obsidianApiKey": "<token>" } next to this file.');
  process.exit(1);
}

process.env.OBSIDIAN_API_KEY = loadKey();

if (!process.env.UV_CACHE_DIR && process.env.LOCALAPPDATA) {
  process.env.UV_CACHE_DIR = path.join(process.env.LOCALAPPDATA, "uv", "cache");
}

const child = spawn("uvx", ["mcp-obsidian"], { stdio: "inherit", env: process.env });
child.on("error", (e) => {
  console.error("obsidian-mcp: failed to spawn uvx:", e.message);
  process.exit(1);
});
child.on("exit", (code) => process.exit(code ?? 0));
