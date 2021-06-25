const { GoogleSpreadsheet } = require('google-spreadsheet');
const credentials = require('./credentials.json');

const docId = new GoogleSpreadsheet('1jxlGdmE9L4DG0WB7H_9gVNwk1l5AWXYzijdOFz5ccPM');

async function accessSpreadsheet() {
    await (docId.useServiceAccountAuth({
      client_email: credentials.client_email,
      private_key: credentials.private_key,
    }));
  
    await docId.loadInfo(); // loads document properties and worksheets
    const sheet = docId.sheetsByIndex[0]; // or use doc.sheetsById[id] // gets the first sheet
    await sheet.loadCells('A1:H27'); // loads specific cells

    // filling the sheet
    try { 
        // running the sheet
        for (let index = 3; index < 27; index++) {
            
            // getting cells
            let ditchedClasses = sheet.getCell(index, 2).value;
            let situation = sheet.getCell(index, 6);
            let naf = sheet.getCell(index, 7);
            
            // average
            let p1 = sheet.getCell(index, 3).value;
            let p2 = sheet.getCell(index, 4).value;
            let p3 = sheet.getCell(index, 5).value;
            let average = Math.ceil((p1 + p2 + p3) / 3);

            // rules
            if (average < 50) {
                situation.value = "Reprovado por Nota";
                naf.value = 0;
            } else if (50 <= average && average < 70) {
                situation.value = "Exame Final";
                naf.value = 100 - average; // right way to calculate the grade for the final exam
            } else {
                situation.value = "Aprovado";
                naf.value = 0;
            }

            // failed due to absence
            const percentageMiss = (ditchedClasses * 100) / 60; // 60 classes
            if (percentageMiss > 25) { // more than 25% = more than 15 ditched classes
                situation.value = "Reprovado por Falta";
                naf.value = 0;
            }
        }
    } catch (error) {
        console.log(error);
    }
    await sheet.saveUpdatedCells();
}
accessSpreadsheet();
