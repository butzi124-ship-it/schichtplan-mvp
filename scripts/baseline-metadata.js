const fs = require("fs");
const path = require("path");
const { createHash } = require("crypto");
const root = "tests/visual-snapshots";
const files = fs.readdirSync(root, { recursive: true }).filter(f => f.endsWith(".png")).sort();
if (files.length !== 33) throw new Error(`Expected 33 reviewed scene/device references, received ${files.length}`);
const provenance = { image: "mcr.microsoft.com/playwright:v1.62.1-noble", playwright: require("@playwright/test/package.json").version,
  sourceCommit: process.env.GITHUB_SHA, workflowRun: process.env.GITHUB_RUN_ID,
  files: files.map(f => ({ file: f.replaceAll("\\", "/"), sha256: createHash("sha256").update(fs.readFileSync(path.join(root, f))).digest("hex") })) };
fs.writeFileSync(path.join(root, "provenance.json"), JSON.stringify(provenance, null, 2) + "\n");
