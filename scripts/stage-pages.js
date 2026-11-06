const fs = require("fs");
const path = require("path");
fs.mkdirSync("_site", { recursive: true });
for (const name of ["schichtplan_mvp.html", "schichtplan_mvp.js"]) fs.copyFileSync(name, path.join("_site", name));
fs.copyFileSync("schichtplan_mvp.html", "_site/index.html");
fs.cpSync("img", "_site/img", { recursive: true });
fs.writeFileSync("_site/qa-commit.txt", (process.env.GITHUB_SHA || "local") + "\n");
