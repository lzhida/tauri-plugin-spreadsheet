use tauri::{
    plugin::{Builder, TauriPlugin},
    Manager, Runtime,
};

use spreadsheet::{
    close_all_xlsx, close_xlsx, copy_sheet, list_xlsx, new_sheet, new_xlsx, read_xlsx, write_xlsx,
};
use state::SpreadsheetState;
use std::{collections::HashMap, sync::Mutex};
use worksheet::{
    get_sheet_highest_column, get_sheet_highest_column_and_row, get_sheet_highest_row,
    get_value_by_column_and_row, remove_row, set_value_by_column_and_row,
};
mod error;
mod spreadsheet;
mod state;
mod worksheet;

/// 初始化插件。
pub fn init<R: Runtime>() -> TauriPlugin<R> {
    Builder::new("spreadsheet")
        .invoke_handler(tauri::generate_handler![
            close_all_xlsx,
            close_xlsx,
            copy_sheet,
            get_sheet_highest_column,
            get_sheet_highest_column_and_row,
            get_sheet_highest_row,
            get_value_by_column_and_row,
            list_xlsx,
            new_sheet,
            new_xlsx,
            read_xlsx,
            remove_row,
            set_value_by_column_and_row,
            write_xlsx,
        ])
        .setup(|app| {
            app.manage(SpreadsheetState {
                spreadsheets: Mutex::new(HashMap::new()),
            });
            Ok(())
        })
        .build()
}
