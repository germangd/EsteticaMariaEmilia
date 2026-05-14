/**
 * Evita EPERM en `.next/trace` cuando queda un `next dev` viejo en 3000/3001
 * y abrís otra terminal: liberamos puertos, borramos trace si existe, y arrancamos Next.
 */
const { spawn } = require("child_process");
const fs = require("fs");
const path = require("path");

const projectRoot = path.join(__dirname, "..");

async function freePorts() {
  const killPort = require("kill-port");
  for (const port of [3000, 3001]) {
    try {
      await killPort(port);
    } catch {
      // sin proceso en ese puerto
    }
  }
}

function tryRemoveNextTrace() {
  try {
    const trace = path.join(projectRoot, ".next", "trace");
    if (fs.existsSync(trace)) fs.unlinkSync(trace);
  } catch {
    // otro proceso o antivirus tiene el archivo
  }
}

async function main() {
  await freePorts();
  tryRemoveNextTrace();

  const child = spawn("npx", ["next", "dev"], {
    stdio: "inherit",
    cwd: projectRoot,
    shell: true,
    env: process.env,
  });

  child.on("exit", (code, signal) => {
    if (signal) process.exit(1);
    process.exit(code ?? 0);
  });
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
