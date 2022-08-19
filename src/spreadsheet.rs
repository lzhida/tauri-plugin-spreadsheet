use crate::error::Error;
use crate::state::{SpreadsheetInfo, SpreadsheetState};
use std::path::Path;
use tauri::State;
use tauri::{command, AppHandle, Runtime, Window};
use umya_spreadsheet::{new_file, reader::xlsx::read, writer::xlsx::write};

/// `get_spreadsheet` 将根据指定 **path** 获取 spreadsheet 实例。
///
/// # Example
///
/// ```
/// get_spreadsheet(state, path, |spreadsheet| {
///   // spreadsheet 操作
/// })
/// ```
pub fn get_spreadsheet<T, F: FnOnce(&mut SpreadsheetInfo) -> Result<T, Error>>(
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

/// `close_all_xlsx` 关闭所有 xlsx 文件。
#[command]
pub fn close_all_xlsx<R: Runtime>(
    _app: AppHandle<R>,
    _window: Window<R>,
    state: State<'_, SpreadsheetState>,
) -> Result<(), Error> {
    let mut map = state.spreadsheets.lock().expect("mutex poisoned!");
    map.clear();
    println!("删除所有 xlsx 文件");
    Ok(())
}

/// `close_xlsx` 关闭指定 xlsx 文件。
#[command]
pub fn close_xlsx<R: Runtime>(
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

/// `copy_sheet` 复制 sheet
#[command]
pub fn copy_sheet<R: Runtime>(
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

/// `list_xlsx` 列出所有打开的 xlsx 文件。
#[command]
pub fn list_xlsx<R: Runtime>(
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

/// `create_sheet` 创建 sheet。
#[command]
pub fn new_sheet<R: Runtime>(
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

/// `new_xlsx` 创建指定 **path** 的 xlsx 文件。
#[command]
pub fn new_xlsx<R: Runtime>(
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

/// `read_xlsx` 读取指定 *path* 的 xlsx 文件。
#[command]
pub fn read_xlsx<R: Runtime>(
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

/// `write_xlsx` 写入 xlsx 文件。
#[command]
pub fn write_xlsx<R: Runtime>(
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
