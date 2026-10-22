const fs = require('fs');

const content = fs.readFileSync('index_hoshin_matrice_XMatrix4_mise_a_jour.html', 'utf8');

const dbMatch = content.match(/const HOSHIN_DATABASE = (\[[\s\S]*?\]);\n/);
const pbMatch = content.match(/const HOSHIN_PROCESS_BLUEPRINT = (\[[\s\S]*?\]);\n/);

if (!dbMatch || !pbMatch) {
  console.error("Could not find the arrays");
  process.exit(1);
}

const dbCode = dbMatch[1];
const pbCode = pbMatch[1];

const db = eval('(' + dbCode + ')');
const pb = eval('(' + pbCode + ')');

fs.writeFileSync('data/hoshin_content.json', JSON.stringify({leaders: db, processBlueprint: pb}, null, 2), 'utf8');
console.log("Successfully updated data/hoshin_content.json");
