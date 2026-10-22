const fs = require('fs');
const html = fs.readFileSync('index.html', 'utf8');

const dataMatch = html.match(/const DATA_V2\s*=\s*({[\s\S]*?});/);
if (dataMatch) {
  const DATA_V2 = eval('(' + dataMatch[1] + ')');
  const axes = [];
  const objectives = [];
  const initiatives = [];
  const actorRelations = [];
  const kpiRelations = [];

  Object.entries(DATA_V2).forEach(([axisLabel, axis]) => {
    axes.push({ label: axisLabel, count: axis.count });
    Object.entries(axis.objectives).forEach(([objectiveLabel, objective]) => {
      objectives.push({ label: objectiveLabel, axis: axisLabel, count: objective.count });
      Object.entries(objective.initiatives).forEach(([initiativeLabel, initiative]) => {
        initiatives.push({ label: initiativeLabel, axis: axisLabel, objective: objectiveLabel, count: initiative.count });
        Object.entries(initiative.actors).forEach(([actorLabel, actor]) => {
          actorRelations.push({ actor: actorLabel, axis: axisLabel, objective: objectiveLabel, initiative: initiativeLabel, count: actor.count });
          actor.kpis.forEach(kpi => {
            kpiRelations.push({ kpi: kpi.label, actor: actorLabel, axis: axisLabel, objective: objectiveLabel, initiative: initiativeLabel, count: kpi.count });
          });
        });
      });
    });
  });
  
  console.log("axes.length:", axes.length);
  console.log("objectives.length:", objectives.length);
  console.log("initiatives.length:", initiatives.length);
  console.log("actorRelations.length:", actorRelations.length);
  console.log("kpiRelations.length:", kpiRelations.length);

  const uAxes = new Set(axes.map(x=>x.label)).size;
  const uObj = new Set(objectives.map(x=>x.label)).size;
  const uInit = new Set(initiatives.map(x=>x.label)).size;
  console.log("unique axes:", uAxes, "unique obj:", uObj, "unique init:", uInit);
} else {
  console.log("DATA_V2 not found");
}
