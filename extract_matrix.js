const fs = require('fs');
const contentOld = fs.readFileSync('index.html', 'utf8');
const contentNew = fs.readFileSync('index_hoshin_matrice_XMatrix4_mise_a_jour.html', 'utf8');

const regexMatrix = /const matrixData = \{[\s\S]*?const defaultCorrelations = \{[\s\S]*?\};\n/g;

const matchNew = contentNew.match(regexMatrix);
if (matchNew) {
    const newStr = matchNew[0];
    const matchOld = contentOld.match(regexMatrix);
    if (matchOld) {
        const updated = contentOld.replace(matchOld[0], newStr);
        fs.writeFileSync('index.html', updated, 'utf8');
        console.log("Successfully replaced matrixData and defaultCorrelations in index.html");
    } else {
        console.log("Could not find old matrixData in index.html");
    }
} else {
    console.log("Could not find new matrixData in index_hoshin_matrice_XMatrix4_mise_a_jour.html");
}
