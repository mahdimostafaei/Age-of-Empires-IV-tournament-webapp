// =========================================================================
// SheetDB.gs - Core Database Interaction Module
// =========================================================================

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

/**
 * Helper to get a sheet by its exact snake_case name.
 */
function get_sheet(sheet_name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheet_name);
}

/**
 * Reads an entire sheet and returns an array of objects mapped to the headers.
 * Example return: [{aoe4_world_id: 123, player_name: "TheRealMoly"}, ...]
 */
function get_all_records(sheet_name) {
  const sheet = get_sheet(sheet_name);
  if (!sheet) throw new Error(`Sheet ${sheet_name} not found.`);
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Empty sheet (only headers)
  
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    let record = {};
    headers.forEach((header, index) => {
      record[header] = row[index];
    });
    return record;
  });
}

/**
 * A simple test function to make sure our bridge is working.
 */
function test_database_connection() {
  // Let's try reading the config_civs tab
  const civs = get_all_records('config_civs');
  Logger.log(`Found ${civs.length} civilizations in the database.`);
  if (civs.length > 0) {
    Logger.log("First civ looks like: " + JSON.stringify(civs[0]));
  }
}
/**
 * Appends a new record object to the specified sheet.
 * Keys in record_obj must match the snake_case headers in the sheet.
 */
function append_record(sheet_name, record_obj) {
  const sheet = get_sheet(sheet_name);
  if (!sheet) throw new Error(`Sheet ${sheet_name} not found.`);
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const new_row = headers.map(header => record_obj[header] !== undefined ? record_obj[header] : "");
  
  sheet.appendRow(new_row);
  return { status: "success", message: `Record added to ${sheet_name}` };
}

/**
 * Updates an existing record identified by a primary key column.
 * Example: update_record_by_key('players', 'aoe4_world_id', 16818159, { current_tourney_elo: 1650 })
 */
function update_record_by_key(sheet_name, key_col_name, key_value, update_obj) {
  const sheet = get_sheet(sheet_name);
  if (!sheet) throw new Error(`Sheet ${sheet_name} not found.`);
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: "error", message: "No records found." };
  
  const headers = data[0];
  const key_index = headers.indexOf(key_col_name);
  
  if (key_index === -1) throw new Error(`Key column '${key_col_name}' not found in ${sheet_name}.`);
  
  for (let i = 1; i < data.length; i++) {
    // Loose equality (==) to handle both string and numeric IDs safely
    if (data[i][key_index] == key_value) {
      const row_num = i + 1;
      
      Object.keys(update_obj).forEach(header => {
        const col_index = headers.indexOf(header);
        if (col_index !== -1) {
          sheet.getRange(row_num, col_index + 1).setValue(update_obj[header]);
        }
      });
      
      return { status: "success", message: `Updated row ${row_num} in ${sheet_name}` };
    }
  }
  
  return { status: "not_found", message: `Record with ${key_col_name} = ${key_value} not found.` };
}