declare class Spreadsheet {
    path: string;
    sheetName: string;
    constructor(path: string, sheetName: string);
    read(): Promise<unknown>;
    create(): Promise<unknown>;
    write(): Promise<unknown>;
    getSheetRange(): Promise<unknown>;
    getSheetRow(): Promise<void>;
    getSheetColumn(): Promise<unknown>;
    newSheet(sheetName: string): Promise<unknown>;
    copySheet(targetSheetName: string, sourceSheetName?: string): Promise<unknown>;
    setValue(local: number[], value: string): Promise<unknown>;
    getValue(local: number[]): Promise<unknown>;
    removeRow(start: number, length: number): Promise<unknown>;
    close(): Promise<unknown>;
    static closeAll(): Promise<unknown>;
    static list(): Promise<unknown>;
}
export { Spreadsheet };
