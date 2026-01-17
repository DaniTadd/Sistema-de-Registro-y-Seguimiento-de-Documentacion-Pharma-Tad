// ambient.d.ts - DEFINICIONES MANUALES PARA OFFICE SCRIPTS (VS CODE)
declare namespace ExcelScript {
    
    interface Workbook {
        getWorksheet(name: string): Worksheet | undefined;
    }

    // Copia y reemplaza esto en tu ambient.d.ts
interface Worksheet {
    getName(): string;
    getRange(address?: string): Range;
    getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number): Range;
    getTable(name: string): Table;
    
    // ESTO ES LO QUE TE FALTABA PARA QUE VS CODE NO LLORE:
    getUsedRange(valuesOnly?: boolean): Range; 
    getProtection(): WorksheetProtection;
    
    // Métodos de seguridad
    protect(options?: WorksheetProtectionOptions, password?: string): void; // (Legacy)
    unprotect(password?: string): void; // (Legacy)
}

// Asegúrate de tener estas interfaces al final del archivo también:
interface WorksheetProtection {
    protect(options?: any, password?: string): void;
    unprotect(password?: string): void;
}

interface Range {
    getValue(): string | number | boolean;
    getValues(): (string | number | boolean)[][];
    setValue(value: any): void;
    setValues(values: any[][]): void;
    clear(applyTo?: any): void;
    getFormat(): RangeFormat;
    getLastRow(): Range;
    getRowIndex(): number;
}

interface RangeFormat {
    getFill(): any;
    getFont(): any;
}

    interface WorksheetProtectionOptions {
        allowAutoFilter?: boolean;
        allowDeleteColumns?: boolean;
        allowDeleteRows?: boolean;
        allowFormatCells?: boolean;
        allowFormatColumns?: boolean;
        allowFormatRows?: boolean;
        allowInsertColumns?: boolean;
        allowInsertRows?: boolean;
        allowInsertHyperlinks?: boolean;
        allowSort?: boolean;
        allowSelectLockedCells?: boolean;
        allowSelectUnlockedCells?: boolean;
        allowPivotTables?: boolean;
    }

    interface Table {
        getColumnByName(name: string): TableColumn;
        getRangeBetweenHeaderAndTotal(): Range;
        getHeaderRowRange(): Range;
        getRange(): Range;
        getRowCount(): number;
        getColumns(): TableColumn[];
        addRow(index: number, values: (string | number | boolean)[]): void;
    }

    interface TableColumn {
        getRangeBetweenHeaderAndTotal(): Range;
        getName(): string;
    }

    interface Range {
        getText(): string;
        getValue(): string | number | boolean;
        getValues(): (string | number | boolean)[][];
        setValue(value: string | number | boolean): void;
        setValues(values: (string | number | boolean)[][]): void;
        clear(applyTo?: ClearApplyTo): void;
        getFormat(): RangeFormat;
        getCell(row: number, column: number): Range;
        merge(across: boolean): void;
        getRowCount(): number;
        getRowIndex(): number;
        getRow(rowIndex: number): Range;
        select(): void
    }

    interface RangeFormat {
        getFill(): Fill;
        getFont(): RangeFont;
        // --- PROTECCIÓN DE FORMATO ---
        getProtection(): FormatProtection;
        
        setHorizontalAlignment(alignment: HorizontalAlignment): void;
        setVerticalAlignment(alignment: VerticalAlignment): void;
        setWrapText(wrap: boolean): void;
        autofitColumns(): void;
        autofitRows(): void;
    }

    // --- INTERFAZ PARA BLOQUEO DE CELDAS ---
    interface FormatProtection {
        setLocked(locked: boolean): void;
    }

    interface Fill {
        setColor(color: string): void;
        clear(): void;
    }

    interface RangeFont {
        setBold(bold: boolean): void;
        setColor(color: string): void;
    }

    enum ClearApplyTo {
        contents = "Contents",
        formats = "Formats"
    }

    enum HorizontalAlignment {
        center = "Center",
        left = "Left",
        right = "Right"
    }

    enum VerticalAlignment {
        center = "Center",
        top = "Top",
        bottom = "Bottom"
    }
}