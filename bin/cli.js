#!/usr/bin/env node

const { spawn } = require("child_process");

// Try uvx first (fast, no venv needed), fall back to pip-installed entry point
const args = process.argv.slice(2);

function tryUvx() {
  const child = spawn("uvx", ["docalyze-mcp-server", ...args], {
    stdio: "inherit",
    shell: process.platform === "win32",
  });
  child.on("error", () => tryPipx());
  child.on("exit", (code) => process.exit(code ?? 1));
}

function tryPipx() {
  const child = spawn("pipx", ["run", "docalyze-mcp-server", ...args], {
    stdio: "inherit",
    shell: process.platform === "win32",
  });
  child.on("error", () => {
    console.error(
      "Error: Could not find uvx or pipx. Install uv (https://docs.astral.sh/uv/) or pipx, then retry."
    );
    process.exit(1);
  });
  child.on("exit", (code) => process.exit(code ?? 1));
}

tryUvx();
