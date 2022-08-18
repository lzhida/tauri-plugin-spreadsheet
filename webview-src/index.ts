import { invoke } from '@tauri-apps/api/tauri';

class Spreadsheet {
  path: string;
  sheetName: string;
  constructor(path: string, sheetName: string) {
    this.path = path;
    this.sheetName = sheetName;
  }

  async read() {
    return await invoke('plugin:spreadsheet|read_xlsx', {
      path: this.path,
    });
  }

  async create() {
    return await invoke('plugin:spreadsheet|new_xlsx', {
      path: this.path,
    });
  }

  async write() {
    return await invoke('plugin:spreadsheet|write_xlsx', {
      path: this.path,
    });
  }

  async getSheetRange() {
    return await invoke('plugin:spreadsheet|get_sheet_highest_column_and_row', {
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  async getSheetRow() {
    await invoke('plugin:spreadsheet|get_sheet_highest_row', {
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  async getSheetColumn() {
    return await invoke('plugin:spreadsheet|get_sheet_highest_column', {
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  async newSheet(sheetName: string) {
    return await invoke('plugin:spreadsheet|new_sheet', {
      path: this.path,
      sheetName: sheetName,
    });
  }

  async copySheet(targetSheetName: string, sourceSheetName?: string) {
    return await invoke('plugin:spreadsheet|copy_sheet', {
      path: this.path,
      sourceSheetName: sourceSheetName || this.sheetName,
      targetSheetName,
    });
  }

  async setValue(local: number[], value: string) {
    return await invoke('plugin:spreadsheet|set_value_by_column_and_row', {
      path: this.path,
      sheetName: this.sheetName,
      local,
      value,
    });
  }

  async getValue(local: number[]) {
    return await invoke('plugin:spreadsheet|get_value_by_column_and_row', {
      path: this.path,
      sheetName: this.sheetName,
      local,
    });
  }

  async removeRow(start: number, length: number) {
    return await invoke('plugin:spreadsheet|remove_row', {
      start,
      length,
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  async close() {
    return await invoke('plugin:spreadsheet|close_xlsx', {
      path: this.path,
    });
  }

  static async closeAll() {
    return await invoke('plugin:spreadsheet|close_all_xlsx');
  }

  static async list() {
    return await invoke('plugin:spreadsheet|list_xlsx');
  }
}

export { Spreadsheet };
