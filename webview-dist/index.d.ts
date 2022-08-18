declare class Spreadsheet {
    path: string;
    sheetName: string;
    constructor(path: string, sheetName: string);
    /**
     * @description: 关闭当前 xlsx 文件
     * @return {*}
     */
    close(): Promise<unknown>;
    /**
     * @description: 关闭所有 xlsx 文件
     * @return {*}
     */
    static closeAll(): Promise<unknown>;
    /**
     * @description: 复制 sheet，不添加 sourceSheetName 时，默认使用当前 sheet
     * @param {string} targetSheetName 目标 sheet 名称
     * @param {string} sourceSheetName 源 sheet 名称
     * @return {*}
     */
    copySheet(targetSheetName: string, sourceSheetName?: string): Promise<unknown>;
    /**
     * @description: 新建 xlsx 文件
     * @return {*}
     */
    create(): Promise<unknown>;
    /**
     * @description: 获取 sheet 列数
     * @return {*}
     */
    getSheetColumn(): Promise<unknown>;
    /**
     * @description: 获取 sheet 行列范围
     * @return {*}
     */
    getSheetRange(): Promise<unknown>;
    /**
     * @description: 获取 sheet 行数
     * @return {*}
     */
    getSheetRow(): Promise<void>;
    /**
     * @description: 获取当前 sheet 指定位置的值
     * @param {number} local 位置，[column, row]
     * @return {*}
     */
    getValue(local: number[]): Promise<unknown>;
    /**
     * @description: 列出所有打开的 xlsx 文件
     * @return {*}
     */
    static list(): Promise<unknown>;
    /**
     * @description: 新建 sheet
     * @param {string} sheetName sheet 名称
     * @return {*}
     */
    newSheet(sheetName: string): Promise<unknown>;
    /**
     * @description: 读取 xlsx 文件
     * @return {*}
     */
    read(): Promise<unknown>;
    /**
     * @description: 删除当前 sheet 指定行
     * @param {number} start 开始行数
     * @param {number} length 要删除的行数
     * @return {*}
     */
    removeRow(start: number, length: number): Promise<unknown>;
    /**
     * @description: 设置当前 sheet 指定位置的值
     * @param {number} local 位置，[column, row]
     * @param {string} value 值
     * @return {*}
     */
    setValue(local: number[], value: string): Promise<unknown>;
    /**
     * @description: 保存 xlsx 文件
     * @return {*}
     */
    write(): Promise<unknown>;
}
export { Spreadsheet };
