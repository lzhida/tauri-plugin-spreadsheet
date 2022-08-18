import { invoke } from '@tauri-apps/api/tauri';

class Spreadsheet {
  path: string;
  sheetName: string;
  constructor(path: string, sheetName: string) {
    this.path = path;
    this.sheetName = sheetName;
  }

  /**
   * @description: 关闭当前 xlsx 文件
   * @return {*}
   */
  async close() {
    return await invoke('plugin:spreadsheet|close_xlsx', {
      path: this.path,
    });
  }

  /**
   * @description: 关闭所有 xlsx 文件
   * @return {*}
   */
  static async closeAll() {
    return await invoke('plugin:spreadsheet|close_all_xlsx');
  }

  /**
   * @description: 复制 sheet，不添加 sourceSheetName 时，默认使用当前 sheet
   * @param {string} targetSheetName 目标 sheet 名称
   * @param {string} sourceSheetName 源 sheet 名称
   * @return {*}
   */
  async copySheet(targetSheetName: string, sourceSheetName?: string) {
    return await invoke('plugin:spreadsheet|copy_sheet', {
      path: this.path,
      sourceSheetName: sourceSheetName || this.sheetName,
      targetSheetName,
    });
  }

  /**
   * @description: 新建 xlsx 文件
   * @return {*}
   */
  async create() {
    return await invoke('plugin:spreadsheet|new_xlsx', {
      path: this.path,
    });
  }

  /**
   * @description: 获取 sheet 列数
   * @return {*}
   */
  async getSheetColumn() {
    return await invoke('plugin:spreadsheet|get_sheet_highest_column', {
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 获取 sheet 行列范围
   * @return {*}
   */
  async getSheetRange() {
    return await invoke('plugin:spreadsheet|get_sheet_highest_column_and_row', {
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 获取 sheet 行数
   * @return {*}
   */
  async getSheetRow() {
    await invoke('plugin:spreadsheet|get_sheet_highest_row', {
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 获取当前 sheet 指定位置的值
   * @param {number} local 位置，[column, row]
   * @return {*}
   */
  async getValue(local: number[]) {
    return await invoke('plugin:spreadsheet|get_value_by_column_and_row', {
      path: this.path,
      sheetName: this.sheetName,
      local,
    });
  }

  /**
   * @description: 列出所有打开的 xlsx 文件
   * @return {*}
   */
  static async list() {
    return await invoke('plugin:spreadsheet|list_xlsx');
  }

  /**
   * @description: 新建 sheet
   * @param {string} sheetName sheet 名称
   * @return {*}
   */
  async newSheet(sheetName: string) {
    return await invoke('plugin:spreadsheet|new_sheet', {
      path: this.path,
      sheetName: sheetName,
    });
  }

  /**
   * @description: 读取 xlsx 文件
   * @return {*}
   */
  async read() {
    return await invoke('plugin:spreadsheet|read_xlsx', {
      path: this.path,
    });
  }

  /**
   * @description: 删除当前 sheet 指定行
   * @param {number} start 开始行数
   * @param {number} length 要删除的行数
   * @return {*}
   */
  async removeRow(start: number, length: number) {
    return await invoke('plugin:spreadsheet|remove_row', {
      start,
      length,
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 设置当前 sheet 指定位置的值
   * @param {number} local 位置，[column, row]
   * @param {string} value 值
   * @return {*}
   */
  async setValue(local: number[], value: string) {
    return await invoke('plugin:spreadsheet|set_value_by_column_and_row', {
      path: this.path,
      sheetName: this.sheetName,
      local,
      value,
    });
  }

  /**
   * @description: 保存 xlsx 文件
   * @return {*}
   */
  async write() {
    return await invoke('plugin:spreadsheet|write_xlsx', {
      path: this.path,
    });
  }
}

export { Spreadsheet };
