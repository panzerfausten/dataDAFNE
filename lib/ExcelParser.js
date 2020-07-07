const ExcelJS = require('exceljs');

const PATHWAY_DATA = {
    "P1":{
        "color":"#1f77b4",
        "url":"http://xake.deib.polimi.it:8081/drupal/omo-turkana/?q=pathway/p389",
        "description":"Structural configuration related to this pathway encompasses the realization of Koysha reservoir and power plant and the full realization of both Kuraz Irrigation Scheme and Private Commercial Agricoltural schemes in the lower Omo. This specific pathway has been selected with the following filtering sequence during the screening exercise in NSL 2019: Fish."
    },
    "P2":{
        "color":"#ff7f0e",
        "url":"http://xake.deib.polimi.it:8081/drupal/omo-turkana/?q=pathway/p3330",
        "description":"Structural configuration related to this pathway encompasses the realization of Koysha reservoir and power plant and the full realization of both Kuraz Irrigation Scheme and Private Commercial Agricoltural schemes in the lower Omo. This specific pathway has been selected with the following filtering sequence during the screening exercise in NSL 2019: Food, Fish, Energy"
    },
    "P3":{
        "color":"#2ca02c",
        "url":"http://xake.deib.polimi.it:8081/drupal/omo-turkana/?q=pathway/p3394",
        "description":"Structural configuration related to this pathway encompasses the realization of Koysha reservoir and power plant and the full realization of both Kuraz Irrigation Scheme and Private Commercial Agricoltural schemes in the lower Omo. This specific pathway has been selected with the following filtering sequence during the screening exercise in NSL 2019: Fish, Food"
    },
    "P4":{
        "color":"#d62728",
        "url":"http://xake.deib.polimi.it:8081/drupal/omo-turkana/?q=pathway/p3428",
        "description":"Structural configuration related to this pathway encompasses the realization of Koysha reservoir and power plant and the full realization of both Kuraz Irrigation Scheme and Private Commercial Agricoltural schemes in the lower Omo. This specific pathway has been selected with the following filtering sequence during the screening exercise in NSL 2019: Irrigation."
    },
    "P5":{
        "color":"#9467bd",
        "url":"http://xake.deib.polimi.it:8081/drupal/omo-turkana/?q=pathway/p3470",
        "description":"Structural configuration related to this pathway encompasses the realization of Koysha reservoir and power plant and the full realization of both Kuraz Irrigation Scheme and Private Commercial Agricoltural schemes in the lower Omo. This specific pathway has been selected with the following filtering sequence during the screening exercise in NSL 2019: Food, Energy"
    },
    "P6":{
        "color":"#8c564b",
        "url":"http://xake.deib.polimi.it:8081/drupal/omo-turkana/?q=pathway/p3607",
        "description":"Structural configuration related to this pathway encompasses the realization of Koysha reservoir and power plant and the full realization of both Kuraz Irrigation Scheme and Private Commercial Agricoltural schemes in the lower Omo."
    },
}
module.exports = {
    process: function(filename){
        return new Promise((resolve,reject) => {
            let data = {};
            var workbook = new ExcelJS.Workbook();
            let headers = [];
            workbook.xlsx.readFile(filename)
                .then(function() {
                    var worksheet = workbook.getWorksheet('Sheet1');
                    let indicators = [];
                    let pathways   = {};

                    let col = worksheet.getColumn(1);
                    worksheet.columns.forEach((item, i) => {
                        let indicator = {};
                        item.eachCell(function(cell, rowNumber) {
                            let header = '';
                            if(i  === 0){
                                header = cell.text.trim();
                                headers.push(cell.text.trim());
                            }else if(i > 0 && rowNumber <=  16){
                                let v = null;
                                if(cell.text.hasOwnProperty('richText')){
                                    v = cell.text.richText[0].text;
                                }else{
                                    if(rowNumber <= 14){
                                        v = cell.text;

                                    }else{
                                        v = isNaN(parseFloat(cell.text)) ? cell.text : parseFloat(cell.text);
                                    }
                                }
                                indicator[headers[rowNumber - 1]] = v;

                            }else{
                            }

                        });
                        if(Object.keys(indicator).length > 0){
                            indicators.push(indicator);
                        }

                    });
                    worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
                      // console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
                      if(rowNumber > 16 && row.values.length > 0 ){
                          let pathway = {};
                          let pathwayName = row.values[1];
                          let data = row.values.slice(2);




                          if(pathwayName.toString().includes('n')){
                            pathwayName = pathwayName.replace('n','');
                            pathways[pathwayName]['name'] = pathwayName
                            pathways[pathwayName]['data'] = row.values.slice(2).map( v => parseFloat(v))
                          }else{
                            pathway['abs'] = row.values.slice(2).map( v => parseFloat(v))
                            pathway['color'] = PATHWAY_DATA[pathwayName].color;
                            pathway['url'] = PATHWAY_DATA[pathwayName].url;
                            pathway['description'] = PATHWAY_DATA[pathwayName].description;
                            pathways[pathwayName] = pathway;
                          }
                      }
                    });
                    data['indicators'] = indicators;
                    data['pathways']   = Object.keys(pathways).map( p => pathways[p]);
                    data['pathways']   = data['pathways'].sort(
                        function(a, b){
                            if(a.name < b.name) { return -1; }
                            if(a.name > b.name) { return 1; }
                            return 0;
                        }
                    )
                    data['metadata'] = {
                        "max": 1000.00,
                        "min": 0
                    }
                    return resolve(data);

                });
        })

    }
}
