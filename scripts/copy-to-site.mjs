import { cpSync, mkdirSync, readdirSync, readFileSync, rmSync, existsSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const here = fileURLToPath(new URL("..", import.meta.url));
const site = join(here, "..", "..", "public", "excel");
const files = [
  ...readdirSync(join(here, "src")).map((f) => [join(here, "src", f), join(site, f)]),
  ...readdirSync(join(here, "assets")).map((f) => [join(here, "assets", f), join(site, "assets", f)]),
  [join(here, "manifest.xml"), join(site, "manifest.xml")],
];

if (process.argv.includes("--check")) {
  const stale = files.filter(([from, to]) => !existsSync(to) || !readFileSync(from).equals(readFileSync(to)));
  if (stale.length) {
    console.error(`public/excel is out of date; run npm run publish-site in clients/excel:\n${stale.map(([f]) => `  ${f}`).join("\n")}`);
    process.exit(1);
  }
  console.log(`public/excel matches clients/excel (${files.length} files)`);
} else {
  rmSync(site, { recursive: true, force: true });
  mkdirSync(join(site, "assets"), { recursive: true });
  files.forEach(([from, to]) => cpSync(from, to));
  console.log(`copied ${files.length} files to public/excel`);
}
