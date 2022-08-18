# Tauri Plugin spreadsheet

基于 `umya_spreadsheet` 模块编写的 `tauri` 插件，封装 `xlsx` 文件的操作。

## 安装方式

### RUST

在 `src-tauri/Cargo.toml` 文件中添加依赖:

```toml
[dependencies.tauri-plugin-spreadsheet]
git = "https://github.com/tauri-plugin-spreadsheet"
tag = "v0.1.0"
```

在 `src-tauri/src/main.rs` 文件中引用插件:

```RUST
use tauri_plugin_spreadsheet;

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_spreadsheet::init())
        .build()
        .run();
}
```

### WEBVIEW

在 `pacakge.json` 文件中添加依赖:

```json
  "dependencies": {
    "tauri-plugin-spreadsheet-api": "github:lzhida/tauri-plugin-spreadsheet#v0.1.0",
  }
```

或者在项目根目录下执行以下命令:

```
npm install https://github.com/lzhida/tauri-plugin-serialport\#v0.1.0
# or
yarn add https://github.com/lzhida/tauri-plugin-serialport\#v0.1.0
```

## 使用

在 `JavaScript` 或 `TypeScript` 文件中使用:

```TypeScript
import { Spreadsheet } from 'tauri-plugin-spreadsheet-api';
```
