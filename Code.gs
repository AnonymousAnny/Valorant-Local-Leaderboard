// Function to start the process upon sheet opening
async function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Refresh Valorant Stats')
      .addItem('Refresh Valorant Stats', 'refreshValorantData')
      .addToUi();
  
    // Automatically refresh data when the sheet is opened
    refreshValorantData();
  }
  
  // Function to refresh Valorant data
  async function refreshValorantData() {
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      const apiKey = scriptProperties.getProperty('API_KEY');
      if (!apiKey) {
        throw new Error('API key not found.');
      }
      
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const lastRow = sheet.getLastRow();
      const dataRange = sheet.getRange(2, 1, lastRow - 1, 2); // Assuming column A and B for name and tag
      const data = dataRange.getValues();
      
      const requests = data.map((row, index) => {
        const name = row[0];
        const tagWithFirstChar = row[1];
        const tag = tagWithFirstChar.slice(1);
        const region = 'ap';
        const url = `https://api.henrikdev.xyz/valorant/v2/mmr/${region}/${name}/${tag}`;
        const options = {
          "method": "GET",
          "headers": {
            "accept": "application/json",
            "Authorization": `${apiKey}`  // Use the API key retrieved from script properties
          }
        };
        return { url: url, options: options, row: index + 2 }; // Adjust row index to 2-based index
      });
  
      // Batch process requests using asynchronous operations
      const responses = await batchFetch(requests);
  
      // Process responses and update the sheet
      responses.forEach(response => {
        const { row, data, error } = response;
        if (error) {
          Logger.log(`Error fetching row ${row}: ${error}`);
          clearTargetCells(sheet, row, error); // Clear cells F to L and set error message in F
        } else {
          updateSheetWithData(row, data);
        }
      });
  
      // Sort the sheet by ELO after processing all rows
      sortSheetByELO();
  
      // Hide columns I and L
      hideColumns(sheet, ['I', 'L']);
  
      // Manage crown symbols after refreshing data
      manageCrownSymbols();
  
    } catch (error) {
      Logger.log(`Error in refreshValorantData: ${error}`);
      throw error;
    }
  }
  
  // Function to clear target cells and set error message in column F
  function clearTargetCells(sheet, row, error) {
    const range = sheet.getRange(`F${row}:L${row}`);
    const blankValues = [error, '', '', '', '', '', '']; // Set error message in column F
    range.setValues([blankValues]);
  }
  
  // Function to update sheet with fetched data
  function updateSheetWithData(row, data) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const currentData = data.current_data;
    const highestRank = data.highest_rank;
  
    const currenttierpatched = currentData.currenttierpatched;
    const ranking_in_tier = currentData.ranking_in_tier;
    const elo = currentData.elo;
    const largeImageUrl = currentData.images.large;
    const highestRankPatchedTier = highestRank ? highestRank.patched_tier : '';
    const highestRankSeason = highestRank ? highestRank.season : '';
    const gamesNeededForRating = currentData.games_needed_for_rating;
  
    // Update sheet with fetched data
    sheet.getRange(`F${row}`).setValue(currenttierpatched);
    sheet.getRange(`G${row}`).setFormula(`=IMAGE("${largeImageUrl}")`);
    sheet.getRange(`H${row}`).setValue(ranking_in_tier);
    sheet.getRange(`I${row}`).setValue(elo);
    sheet.getRange(`J${row}`).setValue(highestRankPatchedTier);
    sheet.getRange(`K${row}`).setValue(highestRankSeason);
    sheet.getRange(`L${row}`).setValue(gamesNeededForRating);
  
    // Apply conditional formatting based on games_needed_for_rating
    if (gamesNeededForRating === 1) {
      const range = sheet.getRange(`F${row}:H${row}`);
      range.setBackground('#D3D3D3'); // Light grey color (adjust as needed)
    }
  }
  
  async function batchFetch(requests) {
    const fetchRequests = requests.map(request => {
      return {
        url: request.url,
        headers: request.options.headers,
        muteHttpExceptions: true
      };
    });
  
    try {
      const responses = UrlFetchApp.fetchAll(fetchRequests);
      return responses.map((response, index) => {
        try {
          const json = JSON.parse(response.getContentText());
          if (json.status === 200) {
            return { row: requests[index].row, data: json.data };
          } else {
            return { row: requests[index].row, error: `API Error: ${json.status}` };
          }
        } catch (error) {
          return { row: requests[index].row, error: `JSON Parse Error: ${error.message}` };
        }
      });
    } catch (error) {
      return requests.map(request => {
        return { row: request.row, error: `Fetch Error: ${error.message}` };
      });
    }
  }
  
  
  // Function to sort sheet by ELO
  function sortSheetByELO() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    range.sort([{ column: 9, ascending: false }]);
  }
  
  // Function to hide columns
  function hideColumns(sheet, columns) {
    columns.forEach(column => {
      const columnIndex = column.charCodeAt(0) - 65 + 1; // Convert A to 1, B to 2, etc.
      sheet.hideColumns(columnIndex);
    });
  }
  
  // Function to manage crown symbols in column E
  function manageCrownSymbols() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const crownSymbol = '👑';
  
    // Remove crown symbol from all cells in column E
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(`E2:E${lastRow}`);
    const cellValues = range.getValues();
    const updatedValues = cellValues.map(([value]) => {
      if (value && value.startsWith(crownSymbol)) {
        return [value.replace(crownSymbol, '').trim()];
      }
      return [value];
    });
    range.setValues(updatedValues);
  
    // Add crown symbol to E2 cell if it's not already there
    const cellE2Value = sheet.getRange('E2').getValue();
    if (!cellE2Value.startsWith(crownSymbol)) {
      sheet.getRange('E2').setValue(crownSymbol + ' ' + cellE2Value);
    }
  }
  
