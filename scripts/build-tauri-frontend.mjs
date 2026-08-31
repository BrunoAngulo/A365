import { spawnSync } from "node:child_process";

const result = spawnSync("npx", ["next", "build"], {
  env: {
    ...process.env,
    TAURI_BUILD: "1",
  },
  shell: process.platform === "win32",
  stdio: "inherit",
});

process.exit(result.status ?? 1);
