import { mkdir, readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import path from "node:path";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sourcePath = path.join(projectRoot, "dist", "index.html");
const webDirectory = path.join(projectRoot, "dist", "web");
const webPath = path.join(webDirectory, "index.html");

const source = await readFile(sourcePath, "utf8");
const webEntry = source
  .replaceAll('="./', '="../')
  .replace("<title>Art Tour Time</title>", "<title>Art Tour Time · Web</title>");

await mkdir(webDirectory, { recursive: true });
await writeFile(webPath, webEntry);
