function importKaggleDataset() {
  var kaggleUsername = ''; // Replace with your actual Kaggle username
  var kaggleKey = ''; // Replace with your actual key
  
  var datasetUrl = ''; // Place here the dataset URL on Kaggle
  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(kaggleUsername + ':' + kaggleKey)
  };
  
  var options = {
    "method": "GET",
    "headers": headers,
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(datasetUrl, options);
  if (response.getResponseCode() === 200) {
    var blob = response.getBlob();
    var unzippedFiles = Utilities.unzip(blob);
    var csvData = Utilities.parseCsv(unzippedFiles[0].getDataAsString());
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw Data');
    sheet.clear(); // Clears the existing data
    sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  } else {
    Logger.log("Failed to fetch data. Response code: " + response.getResponseCode());
  }
}

function getFormattedLatestDate(cleanedDataSheet) {
  // Map of month indices to Norwegian month names
  const monthsNorwegian = ["januar", "februar", "mars", "april", "mai", "juni", 
                           "juli", "august", "september", "oktober", "november", "desember"];

  // Fetch all dates from the cleaned data sheet
  const dateRange = cleanedDataSheet.getRange('A2:A' + cleanedDataSheet.getLastRow());
  const dates = dateRange.getValues();
  let latestDate = new Date(0); // Start with the earliest possible date

  // Iterate through all date entries to find the latest
  dates.forEach(row => {
    const currentDate = new Date(row[0]);
    if (currentDate > latestDate) {
      latestDate = currentDate;
    }
  });

  // Extract day, month index, and year from the latest date
  const day = latestDate.getDate();
  const monthIndex = latestDate.getMonth(); // Month index where January is 0
  const year = latestDate.getFullYear();

  // Format the date string using the Norwegian month names
  return day + '. ' + monthsNorwegian[monthIndex] + ' ' + year;
}

function updateChartData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cleanedData = ss.getSheetByName('Cleaned Data');
  const chartsSheet = ss.getSheetByName('Charts');

 // Get the formatted latest date from the Cleaned Data sheet
  const formattedLatestDate = getFormattedLatestDate(cleanedData);

  // Get the missile type from the dropdown in cell E1
  const missileType = chartsSheet.getRange('E1').getValue();

  // Clear previous chart data and metadata areas
  var lastRow = chartsSheet.getLastRow();
  if (lastRow >= 4) {
    chartsSheet.getRange('A4:E' + lastRow).clearContent();
  }
  chartsSheet.getRange('F3:G15').clearContent();

  // Calculate new data
  const data = cleanedData.getRange('A2:F' + cleanedData.getLastRow()).getValues();
  const monthlyLaunches = {};
  let totalMissilesLaunched = 0;
  let totalMissilesDestroyed = 0;
  let totalEfficiency = 0;
  let countMonths = 0;
  


  data.forEach(row => {
    const type = row[2];
    if ((missileType === 'all missile types' && type !== 'shahed drone') ||
        (missileType === 'all missiles and shaheds') ||
        (type === missileType)) {
      const currentDate = new Date(row[0]);
      const monthYear = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MM/yyyy");
      const launched = parseInt(row[4]);
      const destroyed = parseInt(row[5]);

      if (!isNaN(launched) && !isNaN(destroyed)) {
        monthlyLaunches[monthYear] = monthlyLaunches[monthYear] || { launched: 0, destroyed: 0 };
        monthlyLaunches[monthYear].launched += launched;
        monthlyLaunches[monthYear].destroyed += destroyed;
        if (launched > 0) {
          let efficiency = (destroyed / launched * 100);
          totalEfficiency += efficiency;
          countMonths++;
        }
        totalMissilesLaunched += launched;
        totalMissilesDestroyed += destroyed;
      }
    }
  });

  // Calculate average efficiency
const averageEfficiency = totalMissilesDestroyed / totalMissilesLaunched;

  // Prepare data for 'Charts' sheet
  const chartData = Object.keys(monthlyLaunches)
    .sort((a, b) => new Date(a) - new Date(b)) 
    .map(key => {
      const item = monthlyLaunches[key];
      return [key, item.launched, item.destroyed, (item.launched > 0 ? (item.destroyed / item.launched * 100).toFixed(0) : '0') + '%'];
    });

  // Generate dynamic headers based on the selected missile type
  const headerMap = {
    'all missile types': 'Antall russiske missilangrep',
    'ballistic': 'Antall russiske ballistisk missilangrep',
    'cruise supersonic missile': 'Antall russiske cruise supersonisk missilangrep',
    'cruise subsonic missile': 'Antall russiske cruise subsonisk missilangrep',
    'hypersonic': 'Antall russiske hypersonisk missilangrep',
    'shahed drone': 'Antall russiske Shahed-droneangrep',
    'all missiles and shaheds': 'Antall russiske missile- og droneangrep'
  };

  const headers = [['Måned', headerMap[missileType] || '', 'Antall nedskutt', 'Andel nedskutt, %']];
  chartsSheet.getRange('A3:D3').setValues(headers);
  chartsSheet.getRange('A4:D' + (3 + chartData.length)).setValues(chartData);

  // Write text description header
  chartsSheet.getRange('F3').setValue("Tekstbeskrivelse for grafen");
  chartsSheet.getRange('F5').setValue("tekst:");
  chartsSheet.getRange('G5').setValue("Grafen viser " + headerMap[missileType].toLowerCase() + " på Ukraina fra 1. oktober til " + formattedLatestDate + ".");
  chartsSheet.getRange('F6').setValue("undertekst:");
  chartsSheet.getRange('G6').setValue("Det var " + totalMissilesLaunched + " totale missilangrep i perioden. " + totalMissilesDestroyed + " av dem var skutt ned. Det ukrainske luftforsvaret klarte å skyte ned " + averageEfficiency + " av luftmål.");
  chartsSheet.getRange('F7').setValue("Byline:");
  chartsSheet.getRange('G7').setValue("Pavlo Krasnomovets");
  chartsSheet.getRange('F8').setValue("Kilde:");
  chartsSheet.getRange('G8').setValue("Det Ukrainske luftforsvaret, data samlet av CREDIT."); // Credit the author of the dataset
  chartsSheet.getRange('F9').setValue("Kilde lenken:");
  chartsSheet.getRange('G9').setValue("https://www.kaggle.com/datasets/piterfm/massive-missile-attacks-on-ukraine");

  // Write summary statistics
  chartsSheet.getRange('F11').setValue("Sammendragsstatistikk (valgt periode)");
  chartsSheet.getRange('F13').setValue("Totale missilangrep:");
  chartsSheet.getRange('G13').setValue(totalMissilesLaunched);
  chartsSheet.getRange('F14').setValue("Totale missiler nedskutt:");
  chartsSheet.getRange('G14').setValue(totalMissilesDestroyed);
  chartsSheet.getRange('F15').setValue("Gjennomsnitt andel nedskutt:");
  chartsSheet.getRange('G15').setValue(averageEfficiency);
}

function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() === 'Charts' && range.getA1Notation() === 'E1') {
    updateChartData();
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Update Data')
    .addItem('Import Kaggle Dataset', 'importKaggleDataset')
    .addItem('Update Chart Data', 'updateChartData')
    .addToUi();
}