use std::{collections::HashMap, sync::Mutex};
use umya_spreadsheet::Spreadsheet;

pub struct SpreadsheetState {
    pub spreadsheets: Mutex<HashMap<String, SpreadsheetInfo>>,
}

pub struct SpreadsheetInfo {
    pub spreadsheet: Spreadsheet,
}
