// converts the excel file buildings.xsls to the pdf buildings.pdf
const readXlsxFile = require('read-excel-file/node');
const mustache = require('mustache');
const fs = require('fs');
const pdf = require('html-pdf');
// var merge = require('easy-pdf-merge');


// the html that will be templated to display the data
const template = fs.readFileSync('./template.html', "utf8");

// hardcoded parameters: if the excel file changes, change these to adapt
const ROWS_TO_IGNORE = [0, 1, 13, 23, 43, 58, 68, 73, 78, 79];
const COLUMN_NAMES = [
  'capacity',
  'population',
  'fuelStation',
  'greenPlate',
  'drinkOptions',
  'food',
  'occupancy',
  'electricalMeter',
  'waterMeter',
  'gasMeter',
  'freestyleOrTraditional',
  'secureStorage',
  'elevators',
  'electricalClosets',
  'custodialClosets',
  'highDensityStorage',
  'badgeAccessType',
  'leedRating',
  'dateOpened',
  'sumOfNetArea',
  'leasedOrOwned',
  'buildingCondition',
  'options',
  'landlord',
  'previousName',
  'buildingCode',
  'address',
  'conferenceRooms',
  'showers',
  'dryCleaning',
  'giftShops',
  'performanceBar',
  'reception',
  'bikeRack',
  'quietRooms',
  'mothersRooms',
  'custodialWorkOrders',
  'maintenanceWorkOrders',
  'avWorkOrders',
  'avgTurnoverRate',
  'fitnessMembershipVsUtilization',
  'costPerSeat',
  'depreciation'
]
const PAGE_NUMBER_OVERRIDE = null;

// read the data from the excel file
const buildings = [];
// File path.
readXlsxFile('./buildings.xlsx').then((rows) => {
  // `rows` is an array of rows
  // each row being an array of cells.
  let pdfNum = 0;
  const filenames = [];
  const pageNumbers = PAGE_NUMBER_OVERRIDE || 1000000;
  for (let i = 0; i < rows.length && i < pageNumbers; i++) {
    const row = rows[i];
    if (!ROWS_TO_IGNORE.includes(i)) {
      const building = buildingFrom(row);
      // get the html from the template and the data
      const html = mustache.render(template, building);
      // convert it to pdf
      const pdfOptions = {
        "height": "8in",        // allowed units: mm, cm, in, px
        "width": "10.5in",            // allowed units: mm, cm, in, px
      };
      const filename = './pdf/page' + pdfNum + '.pdf';
      filenames.push(filename);
      pdf.create(html, pdfOptions).toFile(filename, function(err, res) {
          if (err) {
              console.log("ERROR: " + err);
          } else {
              console.log(res.filename);
          }
      });
      pdfNum++;
    }
  }
  // merge(filenames, './final.pdf', function(err) {
  //     if (err) {
  //         console.log(err)
  //     } else {
  //         console.log('SUCCESS');
  //     }
  // });
});

const buildingFrom = row => {
  const building = {};
  building.name = row[0];
  for (let i = 0; i < COLUMN_NAMES.length; i++) {
    const columnName = COLUMN_NAMES[i];
    // the first column we care about, capacity, is at index 4 in the spreadsheet
    const columnIndex = i + 4;
    building[columnName] = row[columnIndex];
  }
  return building;
}
