use tauri::{
    command,
    plugin::{Builder, TauriPlugin},
    AppHandle, Manager, Runtime, State, Window,
};

use error::Error;
use std::{collections::HashMap, path::Path, sync::Mutex};
use umya_spreadsheet::{new_file, reader::xlsx::read, writer::xlsx::write, Spreadsheet, Worksheet};

mod error;

struct SpreadsheetState {
    spreadsheets: Mutex<HashMap<String, SpreadsheetInfo>>,
}

struct SpreadsheetInfo {
    spreadsheet: Spreadsheet,
}

/// `get_spreadsheet` 将根据指定 **path** 获取 spreadsheet 实例。
///
/// # Example
///
/// ```
/// get_spreadsheet(state, path, |spreadsheet| {
///   // spreadsheet 操作
/// })
/// ```
fn get_spreadsheet<T, F: FnOnce(&mut SpreadsheetInfo) -> Result<T, Error>>(
    state: State<'_, SpreadsheetState>,
    path: String,
    f: F,
) -> Result<T, Error> {
    match state.spreadsheets.lock() {
        Ok(mut map) => {
            if !map.contains_key(&path) {
                return Err(Error::NotFound(path.clone()));
            }
            match map.get_mut(&path) {
                Some(spreadsheet) => return f(spreadsheet),
                None => return Err(Error::String(format!("未获取 {} 文件!", &path))),
            }
        }
        Err(error) => return Err(Error::String(format!("获取锁失败! {}", error))),
    }
}

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

/// `read_xlsx` 读取指定 *path* 的 xlsx 文件。
#[command]
fn read_xlsx<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
) -> Result<String, Error> {
    println!("读取 xlsx 文件!");
    match state.spreadsheets.lock() {
        Ok(mut map) => {
            if map.contains_key(&path) {
                println!("读取过文件!");
                return Ok(path.clone());
            }
            let path = Path::new(&path);
            match read(path.clone()) {
                Ok(spreadsheet) => {
                    let data = SpreadsheetInfo { spreadsheet };
                    map.insert(path.display().to_string(), data);
                    println!("读取 {} 文件成功!", path.display().to_string());
                    return Ok(format!("读取 {} 文件成功!", path.display().to_string()));
                }
                Err(error) => {
                    println!("读取 {} 文件失败! {:?}", path.display().to_string(), error);
                    return Err(Error::String(format!(
                        "读取 {} 文件失败!",
                        path.display().to_string(),
                    )));
                }
            }
        }
        Err(error) => return Err(Error::String(format!("获取文件锁失败! {} ", error))),
    }
}

/// `new_xlsx` 创建指定 **path** 的 xlsx 文件。
#[command]
fn new_xlsx<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
) -> Result<String, Error> {
    match state.spreadsheets.lock() {
        Ok(mut map) => {
            if map.contains_key(&path) {
                let message = format!("新建文件 {} 已存在", &path);
                println!("{}", message);
                return Err(Error::String(message));
            }
            let path = Path::new(&path);
            let book = new_file();
            let data = SpreadsheetInfo { spreadsheet: book };
            map.insert(path.display().to_string(), data);
            let message = format!("新建 xlsx 文件: {}", path.display().to_string());
            return Ok(message);
        }
        Err(error) => return Err(Error::String(format!("获取文件锁失败! {} ", error))),
    }
}

/// `write_xlsx` 写入 xlsx 文件。
#[command]
fn write_xlsx<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
) -> Result<(), Error> {
    get_spreadsheet(state, path.clone(), |spreadsheet| {
        let path = Path::new(&path);
        match write(&spreadsheet.spreadsheet, path) {
            Ok(_) => {
                let message = format!("保存 xlsx 文件 {}", path.display().to_string());
                println!("{}", message);
                return Ok(());
            }
            Err(_) => {
                return Err(Error::String(format!(
                    "保存文件 {} 失败!",
                    path.display().to_string()
                )));
            }
        }
    })
}

/// `get_sheet_highest_column_and_row` 获取 sheet 最大行数和列数。
///
/// # Return value
///
/// `(u32, u32)` - (column, row)
#[command]
fn get_sheet_highest_column_and_row<R: Runtime>(
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

/// `get_sheet_highest_column` 获取 sheet 最大列数。
#[command]
fn get_sheet_highest_column<R: Runtime>(
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

/// `get_sheet_highest_raw` 获取 sheet 最大行数。
#[command]
fn get_sheet_highest_row<R: Runtime>(
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

/// `create_sheet` 创建 sheet。
#[command]
fn new_sheet<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    sheet_name: String,
) -> Result<String, Error> {
    get_spreadsheet(state, path, |spreadsheet| {
        if let Ok(_) = spreadsheet.spreadsheet.new_sheet(&sheet_name) {
            println!("创建新 sheet {}", &sheet_name);
            return Ok(sheet_name);
        } else {
            return Err(Error::String(format!("新建 {} sheet 失败!", &sheet_name)));
        }
    })
}

/// `copy_sheet` 复制 sheet
#[command]
fn copy_sheet<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
    source_sheet_name: String,
    target_sheet_name: String,
) -> Result<(), Error> {
    get_spreadsheet(state, path, |spreadsheet| {
        match spreadsheet
            .spreadsheet
            .get_sheet_by_name(&source_sheet_name)
        {
            Ok(worksheet) => {
                let mut clone_sheet = worksheet.clone();
                clone_sheet.set_name(&target_sheet_name);
                match spreadsheet.spreadsheet.add_sheet(clone_sheet) {
                    Ok(_) => {
                        println!(
                            "复制 sheet {} 到 sheet {}",
                            &source_sheet_name, &target_sheet_name
                        );
                        return Ok(());
                    }
                    Err(error) => {
                        return Err(Error::String(format!(
                            "添加 sheet {} 失败! {}",
                            &target_sheet_name, error
                        )))
                    }
                }
            }
            Err(error) => {
                return Err(Error::String(format!(
                    "获取源 sheet {} 失败! {}",
                    &source_sheet_name, error
                )));
            }
        }
    })
}

/// `set_value_by_column_and_row` 根据行列设置值。
#[command]
fn set_value_by_column_and_row<R: Runtime>(
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

/// `get_value_by_column_and_row` 根据行列获取值。
#[command]
fn get_value_by_column_and_row<R: Runtime>(
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
fn remove_row<R: Runtime>(
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

/// `close_xlsx` 关闭指定 xlsx 文件。
#[command]
fn close_xlsx<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
    path: String,
) -> Result<(), Error> {
    let mut map = state.spreadsheets.lock().expect("mutex poisoned!");
    map.remove(&path);
    println!("删除 xlsx 文件 {}!", &path);
    Ok(())
}

/// `close_all_xlsx` 关闭所有 xlsx 文件。
#[command]
fn close_all_xlsx<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
) -> Result<(), Error> {
    let mut map = state.spreadsheets.lock().expect("mutex poisoned!");
    map.clear();
    println!("删除所有 xlsx 文件");
    Ok(())
}

/// `list_xlsx` 列出所有打开的 xlsx 文件。
#[command]
fn list_xlsx<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
) -> Result<Vec<String>, Error> {
    let map = state.spreadsheets.lock().expect("mutex poisoned!");
    let list = map.keys();
    let mut keys = Vec::new();
    for key in list {
        keys.push(key.clone());
    }
    println!("列表内容: {:?}", &keys);
    Ok(keys)
}

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
