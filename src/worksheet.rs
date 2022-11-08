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

/// `append_column` 附加指定内容的列到表格最后一列后面。
#[command]
pub fn append_column<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    data: Vec<Vec<String>>,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        let mut column_index = worksheet.get_highest_column();
        for column in data {
            column_index += 1;
            for (i, v) in column.iter().enumerate() {
                let row_index: u32 = i as u32 + 1;
                println!("插入数据: {}, {}, {}", &row_index, &column_index, &v);
                worksheet
                    .get_cell_by_column_and_row_mut(&column_index, &row_index)
                    .set_value(v);
            }
        }
        Ok(())
    })
}

/// `append_row` 附加内容行到表格末尾。
#[command]
pub fn append_row<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    data: Vec<Vec<String>>,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name.clone(), |worksheet| {
        let mut row_index = worksheet.get_highest_row();
        for row in data {
            row_index += 1;
            for (i, v) in row.iter().enumerate() {
                let column_index: u32 = i as u32 + 1;
                println!("插入数据: {}, {}, {}", &row_index, &column_index, &v);
                worksheet
                    .get_cell_by_column_and_row_mut(&column_index, &row_index)
                    .set_value(v);
            }
        }
        Ok(())
    })
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

#[command]
pub fn get_collection_by_column<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    range: (u32, u32),
) -> Result<Vec<Vec<String>>, Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        let mut data: Vec<Vec<String>> = Vec::new();
        let (start, end) = range;
        for i in start..end {
            let collection = worksheet.get_collection_by_column(&i);
            let mut temp = vec![String::from(""); collection.len()];
            for (_, item) in collection.iter().enumerate() {
                if let Ok(row) = item
                    .get_coordinate()
                    .get_row_num()
                    .to_string()
                    .parse::<usize>()
                {
                    if row - 1 < temp.len() {
                        temp[row - 1] = item.get_value().to_string();
                    }
                }
            }
            data.push(temp)
        }
        Ok(data)
    })
}

#[command]
pub fn get_collection_by_row<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    range: (u32, u32),
) -> Result<Vec<Vec<String>>, Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        let mut data: Vec<Vec<String>> = Vec::new();
        let (start, end) = range;
        for i in start..end {
            let collection = worksheet.get_collection_by_row(&i);
            let mut temp = vec![String::from(""); collection.len()];
            for (_, item) in collection.iter().enumerate() {
                if let Ok(column) = item
                    .get_coordinate()
                    .get_col_num()
                    .to_string()
                    .parse::<usize>()
                {
                    if column - 1 < temp.len() {
                        temp[column - 1] = item.get_value().to_string();
                    }
                }
            }
            data.push(temp)
        }
        Ok(data)
    })
}

/// `insert_column` 在指定位置插入指定内容的列。
#[command]
pub fn insert_column<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    column_index: u32,
    data: Vec<Vec<String>>,
    is_add: bool,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        if is_add {
            let num_columns: u32 = data.len().try_into().unwrap();
            worksheet.insert_new_column_by_index(&column_index, &num_columns);
            println!("插入 sheet 列 {}, {}", &column_index, &num_columns);
        }
        let mut column_index = column_index;
        for column in data {
            for (i, v) in column.iter().enumerate() {
                let row_index: u32 = i as u32 + 1;
                println!("插入数据: {}, {}, {}", &row_index, &column_index, &v);
                worksheet
                    .get_cell_by_column_and_row_mut(&column_index, &row_index)
                    .set_value(v);
            }
            column_index += 1;
        }
        Ok(())
    })
}

/// `insert_new_column` 在 `column` 位置插入 `num_columns` 列。
#[command]
pub fn insert_new_column<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    column: String,
    num_columns: u32,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        worksheet.insert_new_column(&column, &num_columns);
        println!("插入空白列数据: {}, {}", column, num_columns);
        Ok(())
    })
}

/// `insert_new_column_by_index` 在 `column_index` 位置插入 `num_columns` 列。
#[command]
pub fn insert_new_column_by_index<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    column_index: u32,
    num_columns: u32,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        worksheet.insert_new_column_by_index(&column_index, &num_columns);
        println!("移除列数据: {}, {}", column_index, num_columns);
        Ok(())
    })
}

/// `insert_new_row` 在指定位置插入指定空白行。
#[command]
pub fn insert_new_row<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    row_index: u32,
    num_rows: u32,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        worksheet.insert_new_row(&row_index, &num_rows);
        println!("插入 sheet 行 {}, {}", &row_index, &num_rows);
        Ok(())
    })
}

/// `insert_row` 在指定位置插入指定内容的行。
#[command]
pub fn insert_row<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    row_index: u32,
    data: Vec<Vec<String>>,
    is_add: bool,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        if is_add {
            let num_rows: u32 = data.len().try_into().unwrap();
            worksheet.insert_new_row(&row_index, &num_rows);
            println!("插入 sheet 行 {}, {}", &row_index, &num_rows);
        }
        let mut row_index = row_index;
        for row in data {
            for (i, v) in row.iter().enumerate() {
                let column_index: u32 = i as u32 + 1;
                println!("插入数据: {}, {}, {}", &row_index, &column_index, &v);
                worksheet
                    .get_cell_by_column_and_row_mut(&column_index, &row_index)
                    .set_value(v);
            }
            row_index += 1;
        }
        Ok(())
    })
}

/// `remove_column` 删除从 `column` 开始的 `num_columns` 列。
///
/// # Arguments
///
/// - `column` 列数的字符串，如: "B"
/// - `num_columns` 要移除的列数，如："3"
#[command]
pub fn remove_column<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    column: String,
    num_columns: u32,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        worksheet.remove_column(&column, &num_columns);
        println!("移除列数据: {}, {}", column, num_columns);
        Ok(())
    })
}

/// `remove_column_by_index` 删除从 `column_index` 开始的 `num_columns` 列。
///
/// # Arguments
///
/// - `column_index` 列数的索引数，如: "2"
/// - `num_columns` 要移除的列数，如："3"
#[command]
pub fn remove_column_by_index<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
    column_index: u32,
    num_columns: u32,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        worksheet.remove_column_by_index(&column_index, &num_columns);
        println!("移除列数据: {}, {}", column_index, num_columns);
        Ok(())
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
    row_index: u32,
    num_rows: u32,
) -> Result<(), Error> {
    get_worksheet(state, path, sheet_name, |worksheet| {
        let highest_row = worksheet.get_highest_row();
        let mut num_rows = num_rows;
        if row_index < 1 {
            return Err(Error::String(String::from("row_index 参数不能小于 1!")));
        }
        if num_rows < 1 {
            return Err(Error::String(String::from("num_rows 参数不能小于 1!")));
        }
        if row_index >= highest_row {
            return Err(Error::String(String::from(
                "row_index 参数大小超过最大行数!",
            )));
        } else if row_index + num_rows > highest_row {
            num_rows = highest_row - row_index;
        }

        println!("移除 sheet 行 {}, {}", &row_index, &num_rows);
        worksheet.remove_row(&row_index, &num_rows);
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
