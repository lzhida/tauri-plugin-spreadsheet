import { invoke } from '@tauri-apps/api/tauri';

class Spreadsheet {
  path: string;
  sheetName: string;
  constructor(path: string, sheetName: string) {
    this.path = path;
    this.sheetName = sheetName;
  }

  /**
   * @description: 附加数据到表格最后一列。
   * @param {string[][]} data 二维字符串数组
   * @return {Promise<void>}
   */
  async appendColumn(data: string[][]): Promise<void> {
    return await invoke('plugin:spreadsheet|append_column', {
      data,
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 附加数据到表格最后一行。
   * @param {string[][]} data 二维字符串数组
   * @return {Promise<void>}
   */
  async appendRow(data: string[][]): Promise<void> {
    return await invoke('plugin:spreadsheet|append_row', {
      data,
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 关闭当前 xlsx 文件
   * @return {Promise<void>}
   */
  async close(): Promise<void> {
    return await invoke('plugin:spreadsheet|close_xlsx', {
      path: this.path,
    });
  }

  /**
   * @description: 关闭所有 xlsx 文件
   * @return {Promise<void>}
   */
  static async closeAll(): Promise<void> {
    return await invoke('plugin:spreadsheet|close_all_xlsx');
  }

  /**
   * @description: 复制 sheet，不添加 sourceSheetName 时，默认使用当前 sheet
   * @param {string} targetSheetName 目标 sheet 名称
   * @param {string} sourceSheetName 源 sheet 名称
   * @return {Promise<void>}
   */
  async copySheet(
    targetSheetName: string,
    sourceSheetName?: string,
  ): Promise<void> {
    return await invoke('plugin:spreadsheet|copy_sheet', {
      path: this.path,
      sourceSheetName: sourceSheetName || this.sheetName,
      targetSheetName,
    });
  }

  /**
   * @description: 新建 xlsx 文件
   * @return {Promise<void>}
   */
  async create(): Promise<void> {
    return await invoke('plugin:spreadsheet|new_xlsx', {
      path: this.path,
    });
  }

  /**
   * @description: 获取 sheet 列数
   * @return {Promise<number>}
   */
  async getSheetColumn(): Promise<number> {
    return await invoke('plugin:spreadsheet|get_sheet_highest_column', {
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 获取 sheet 行列范围
   * @return {Promise<number[]>}
   */
  async getSheetRange(): Promise<number[]> {
    return await invoke('plugin:spreadsheet|get_sheet_highest_column_and_row', {
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 获取 sheet 行数
   * @return {Promise<number>}
   */
  async getSheetRow(): Promise<number> {
    return await invoke('plugin:spreadsheet|get_sheet_highest_row', {
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 获取当前 sheet 指定位置的值
   * @param {number} local 位置，[column, row]
   * @return {Promise<string>}
   */
  async getValue(local: number[]): Promise<string> {
    return await invoke('plugin:spreadsheet|get_value_by_column_and_row', {
      path: this.path,
      sheetName: this.sheetName,
      local,
    });
  }

  /**
   * @description: 在指定位置 `columnIndex` 插入数据。
   * @param {number} columnIndex 指定开始列数
   * @param {string[][]} data 插入的数据，二维字符串数组
   * @param {boolean} isAdd 是否新增空白插入，false 则覆盖原来数据
   * @return {Promise<void>}
   */
  async insertColumn(
    columnIndex: number,
    data: string[][],
    isAdd = true,
  ): Promise<void> {
    return await invoke('plugin:spreadsheet|insert_column', {
      data,
      isAdd,
      columnIndex,
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 在指定位置 `columnIndex` 插入 `num_columns` 空白列。
   * @param {number} columnIndex 指定开始列数
   * @param {number} numColumns 插入列数
   * @return {Promise<void>}
   */
  async insertNewColumn(
    columnIndex: number | string,
    numColumns: number,
  ): Promise<void> {
    if (typeof columnIndex === 'number') {
      return await invoke('plugin:spreadsheet|insert_new_column_by_index', {
        columnIndex,
        numColumns,
        path: this.path,
        sheetName: this.sheetName,
      });
    } else if (typeof columnIndex === 'string') {
      return await invoke('plugin:spreadsheet|insert_new_column', {
        column: columnIndex,
        numColumns,
        path: this.path,
        sheetName: this.sheetName,
      });
    } else {
      return Promise.reject('columnIndex 参数类型不对!');
    }
  }

  /**
   * @description: 在指定位置 `rowIndex` 插入数据。
   * @param {number} rowIndex 指定行数
   * @param {string[][]} data 插入的数据，二维字符串数组
   * @param {boolean} isAdd 是否新增空白插入，false 则覆盖原来数据
   * @return {Promise<void>}
   */
  async insertRow(
    rowIndex: number,
    data: string[][],
    isAdd = true,
  ): Promise<void> {
    return await invoke('plugin:spreadsheet|insert_row', {
      data,
      isAdd,
      rowIndex,
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 列出所有打开的 xlsx 文件
   * @return {Promise<string[]>}
   */
  static async list(): Promise<string[]> {
    return await invoke('plugin:spreadsheet|list_xlsx');
  }

  /**
   * @description: 新建 sheet
   * @param {string} sheetName sheet 名称
   * @return {Promise<void>}
   */
  async newSheet(sheetName: string): Promise<void> {
    return await invoke('plugin:spreadsheet|new_sheet', {
      path: this.path,
      sheetName: sheetName,
    });
  }

  /**
   * @description: 读取 xlsx 文件
   * @return {Promise<void>}
   */
  async read(): Promise<void> {
    return await invoke('plugin:spreadsheet|read_xlsx', {
      path: this.path,
    });
  }

  /**
   * @description: 删除 `start` 开始的 `length` 的列数。
   * @param {number} start 开始列数
   * @param {number} length 删除列数
   * @return {Promise<void>}
   */
  async removeColumn(start: number | string, length: number): Promise<void> {
    if (typeof start === 'number') {
      return await invoke('plugin:spreadsheet|insert_new_column_by_index', {
        columnIndex: start,
        numColumns: length,
        path: this.path,
        sheetName: this.sheetName,
      });
    } else if (typeof start === 'string') {
      return await invoke('plugin:spreadsheet|insert_new_column', {
        column: start,
        numColumns: length,
        path: this.path,
        sheetName: this.sheetName,
      });
    } else {
      return Promise.reject('start 参数类型不对!');
    }
  }

  /**
   * @description: 删除当前 sheet 指定行
   * @param {number} rowIndex 开始行数
   * @param {number} numRows 要删除的行数
   * @return {Promise<void>}
   */
  async removeRow(rowIndex: number, numRows: number): Promise<void> {
    return await invoke('plugin:spreadsheet|remove_row', {
      rowIndex,
      numRows,
      path: this.path,
      sheetName: this.sheetName,
    });
  }

  /**
   * @description: 设置当前 sheet 指定位置的值
   * @param {number} local 位置，[column, row]
   * @param {string} value 值
   * @return {Promise<void>}
   */
  async setValue(local: number[], value: string): Promise<void> {
    return await invoke('plugin:spreadsheet|set_value_by_column_and_row', {
      path: this.path,
      sheetName: this.sheetName,
      local,
      value,
    });
  }

  /**
   * @description: 保存 xlsx 文件
   * @return {Promise<void>}
   */
  async write(): Promise<void> {
    return await invoke('plugin:spreadsheet|write_xlsx', {
      path: this.path,
    });
  }
}

export { Spreadsheet };
