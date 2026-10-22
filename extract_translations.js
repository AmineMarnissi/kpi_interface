const fs = require('fs');
const contentOld = fs.readFileSync('index.html', 'utf8');
const contentNew = fs.readFileSync('index_hoshin_matrice_XMatrix4_mise_a_jour.html', 'utf8');

const regexTrans = /const matrixTranslations = \{[\s\S]*?\/\/ Bascule linguistique/g;

const matchNew = contentNew.match(regexTrans);
if (matchNew) {
    const newStr = matchNew[0];
    const matchOld = contentOld.match(regexTrans);
    if (matchOld) {
        const updated = contentOld.replace(matchOld[0], newStr);
        fs.writeFileSync('index.html', updated, 'utf8');
        console.log("Successfully replaced matrixTranslations in index.html");
    } else {
        console.log("Could not find old matrixTranslations in index.html");
    }
} else {
    console.log("Could not find new matrixTranslations in index_hoshin_matrice_XMatrix4_mise_a_jour.html");
}
