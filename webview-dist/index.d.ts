declare class Spreadsheet {
    path: string;
    sheetName: string;
    constructor(path: string, sheetName: string);
    /**
     * @description: 附加数据到表格最后一列。
     * @param {string[][]} data 二维字符串数组
     * @return {Promise<void>}
     */
    appendColumn(data: string[][]): Promise<void>;
    /**
     * @description: 附加数据到表格最后一行。
     * @param {string[][]} data 二维字符串数组
     * @return {Promise<void>}
     */
    appendRow(data: string[][]): Promise<void>;
    /**
     * @description: 关闭当前 xlsx 文件
     * @return {Promise<void>}
     */
    close(): Promise<void>;
    /**
     * @description: 关闭所有 xlsx 文件
     * @return {Promise<void>}
     */
    static closeAll(): Promise<void>;
    /**
     * @description: 复制 sheet，不添加 sourceSheetName 时，默认使用当前 sheet
     * @param {string} targetSheetName 目标 sheet 名称
     * @param {string} sourceSheetName 源 sheet 名称
     * @return {Promise<void>}
     */
    copySheet(targetSheetName: string, sourceSheetName?: string): Promise<void>;
    /**
     * @description: 新建 xlsx 文件
     * @return {Promise<void>}
     */
    create(): Promise<void>;
    /**
     * @description: 获取 sheet 列数
     * @return {Promise<number>}
     */
    getSheetColumn(): Promise<number>;
    /**
     * @description: 获取 sheet 行列范围
     * @return {Promise<number[]>}
     */
    getSheetRange(): Promise<number[]>;
    /**
     * @description: 获取 sheet 行数
     * @return {Promise<number>}
     */
    getSheetRow(): Promise<number>;
    /**
     * @description: 获取当前 sheet 指定位置的值
     * @param {number} local 位置，[column, row]
     * @return {Promise<string>}
     */
    getValue(local: number[]): Promise<string>;
    /**
     * @description: 在指定位置 `columnIndex` 插入数据。
     * @param {number} columnIndex 指定开始列数
     * @param {string[][]} data 插入的数据，二维字符串数组
     * @param {boolean} isAdd 是否新增空白插入，false 则覆盖原来数据
     * @return {Promise<void>}
     */
    insertColumn(columnIndex: number, data: string[][], isAdd?: boolean): Promise<void>;
    /**
     * @description: 在指定位置 `columnIndex` 插入 `num_columns` 空白列。
     * @param {number} columnIndex 指定开始列数
     * @param {number} numColumns 插入列数
     * @return {Promise<void>}
     */
    insertNewColumn(columnIndex: number | string, numColumns: number): Promise<void>;
    /**
     * @description: 在指定位置 `rowIndex` 插入数据。
     * @param {number} rowIndex 指定行数
     * @param {string[][]} data 插入的数据，二维字符串数组
     * @param {boolean} isAdd 是否新增空白插入，false 则覆盖原来数据
     * @return {Promise<void>}
     */
    insertRow(rowIndex: number, data: string[][], isAdd?: boolean): Promise<void>;
    /**
     * @description: 列出所有打开的 xlsx 文件
     * @return {Promise<string[]>}
     */
    static list(): Promise<string[]>;
    /**
     * @description: 新建 sheet
     * @param {string} sheetName sheet 名称
     * @return {Promise<void>}
     */
    newSheet(sheetName: string): Promise<void>;
    /**
     * @description: 读取 xlsx 文件
     * @return {Promise<void>}
     */
    read(): Promise<void>;
    /**
     * @description: 删除 `start` 开始的 `length` 的列数。
     * @param {number} start 开始列数
     * @param {number} length 删除列数
     * @return {Promise<void>}
     */
    removeColumn(start: number | string, length: number): Promise<void>;
    /**
     * @description: 删除当前 sheet 指定行
     * @param {number} rowIndex 开始行数
     * @param {number} numRows 要删除的行数
     * @return {Promise<void>}
     */
    removeRow(rowIndex: number, numRows: number): Promise<void>;
    /**
     * @description: 设置当前 sheet 指定位置的值
     * @param {number} local 位置，[column, row]
     * @param {string} value 值
     * @return {Promise<void>}
     */
    setValue(local: number[], value: string): Promise<void>;
    /**
     * @description: 保存 xlsx 文件
     * @return {Promise<void>}
     */
    write(): Promise<void>;
}
export { Spreadsheet };
