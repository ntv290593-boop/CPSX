// Ghép SKILL.template.md + lib/cpsx_lib.js -> cpsx-master/SKILL.md. Chạy: node skill-src/build.js (từ thư mục _MASTER)
const fs = require("fs"), path = require("path");
const root = path.join(__dirname, "..");
const lib = fs.readFileSync(path.join(root, "cpsx-master/lib/cpsx_lib.js"), "utf8").replace(/\r/g, "");
const ver = (lib.match(/VER: "([^"]+)"/) || [])[1];
const tpl = fs.readFileSync(path.join(__dirname, "SKILL.template.md"), "utf8").replace(/\r/g, "");
const out = tpl.replace("{{LIB}}", () => lib.trimEnd()).replace(/\{\{VER\}\}/g, ver);
fs.writeFileSync(path.join(root, "cpsx-master/SKILL.md"), out);
console.log("SKILL.md", out.length, "chars, lib v" + ver);
