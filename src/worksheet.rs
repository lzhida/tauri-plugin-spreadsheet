use crate::error::Error;
use crate::state::SpreadsheetState;
use tauri::{command, AppHandle, Runtime, State, Window};
use umya_spreadsheet::Worksheet;

/// `get_worksheet` 根据 `path` 和 `sheet_name` 获取文件 sheet 实例。
fn get_worksheet<T, F: FnOnce(&mut Worksheet) -> Result<T, Error>>(
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    f: F,
) -> Result<T, Error> {
    match state.spreadsheets.lock() {
        Ok(mut map) => match map.get_mut(&path) {
            Some(spreadsheet) => match spreadsheet.spreadsheet.get_sheet_by_name_mut(&sheet_name) {
                Ok(worksheet) => {
                    return f(worksheet);
                }
                Err(error) => {
                    return Err(Error::String(format!(
                        "获取 {} 文件 {} sheet 失败! {}",
                        &path, &sheet_name, error
                    )))
                }
            },
            None => return Err(Error::String(format!("未获取 {} 文件!", &path))),
        },
        Err(error) => return Err(Error::String(format!("获取文件锁失败! {} ", error))),
    }
}

/// `get_sheet_highest_column` 获取 sheet 最大列数。
#[command]
pub fn get_sheet_highest_column<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
) -> Result<u32, Error> {
    get_worksheet(state, path, sheet_name.clone(), |worksheet| {
        let data = worksheet.get_highest_column();
        println!("获取 sheet {} 列数: {}", &sheet_name, &data);
        Ok(data)
    })
}

/// `get_sheet_highest_column_and_row` 获取 sheet 最大行数和列数。
///
/// # Return value
///
/// `(u32, u32)` - (column, row)
#[command]
pub fn get_sheet_highest_column_and_row<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
) -> Result<(u32, u32), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        let data = worksheet.get_highest_column_and_row();
        println!("获取 sheet 行列 {:?}", data);
        Ok(data)
    })
}

/// `get_sheet_highest_raw` 获取 sheet 最大行数。
#[command]
pub fn get_sheet_highest_row<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
) -> Result<u32, Error> {
    get_worksheet(state, path, sheet_name.clone(), |worksheet| {
        let data = worksheet.get_highest_row();
        println!("获取 sheet {} 行数: {}", &sheet_name, &data);
        Ok(data)
    })
}

/// `get_value_by_column_and_row` 根据行列获取值。
#[command]
pub fn get_value_by_column_and_row<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    local: (u32, u32),
) -> Result<String, Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        let data = worksheet.get_value_by_column_and_row(&local.0, &local.1);
        println!("获取 sheet 位置 {:?} 的值: {}", &local, &data);
        Ok(data)
    })
}

/// `reomove_row` 删除行。
#[command]
pub fn remove_row<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    start: u32,
    length: u32,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        let row_num = worksheet.get_highest_row();
        let mut start = start;
        let mut length = length;
        if start >= row_num {
            start = row_num;
            length = 1;
        } else if start + length > row_num {
            length = row_num - start;
        }

        worksheet.remove_row(&start, &length);
        println!("移除 sheet 行 {}, {}", &start, &length);
        Ok(())
    })
}

/// `set_value_by_column_and_row` 根据行列设置值。
#[command]
pub fn set_value_by_column_and_row<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    local: (u32, u32),
    value: String,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        worksheet
            .get_cell_by_column_and_row_mut(&local.0, &local.1)
            .set_value(value.clone());
        println!("设置 sheet 的位置 {:?} 的值为 {}", &local, value);
        Ok(())
    })
}
