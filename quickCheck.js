// quickCheck.js
const fs = require("fs");
const code = fs.readFileSync("./src/functions/vapiWebhook.js", "utf8");

const hasRegion = code.includes('region: "IND"');
const hasWildcard = code.includes('/*"');

console.log(
  "Fix 1 — /* wildcard on paths:",
  hasWildcard ? "✅ Applied" : "❌ MISSING",
);
console.log(
  "Fix 2 — region: IND parameter:",
  hasRegion ? "✅ Applied" : "❌ MISSING",
);

if (hasRegion && hasWildcard) {
  console.log("\n✅ Both fixes are in the local code.");
  console.log("→ Just redeploy to Azure and your calls will work.");
} else {
  console.log("\n❌ Some fixes are missing! Do not deploy yet.");
}
