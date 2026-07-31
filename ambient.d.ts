// ambient.d.ts - DEFINICIONES MANUALES PARA OFFICE SCRIPTS (VS CODE)
// Arquitectura: Consolidada para SGC-Engine v3.0

declare namespace ExcelScript {
    
    interface Workbook {
        getActiveWorksheet(): Worksheet;
        getWorksheet(name: string): Worksheet | undefined;
        getWorksheets(): Worksheet[];
        getTable(name: string): Table | undefined;
        getNamedItem(name: string): NamedItem | undefined;
    }

    interface Worksheet {
        getName(): string;
        setVisibility(visibility: SheetVisibility): void;
        getVisibility(): SheetVisibility;
        getRange(address?: string): Range;
        getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number): Range;
        getTable(name: string): Table | undefined;
        getNamedItem(name: string): NamedItem | undefined;
        getUsedRange(valuesOnly?: boolean): Range; 
        getProtection(): WorksheetProtection;
    }

    interface WorksheetProtection {
        protect(options?: WorksheetProtectionOptions, password?: string): void;
        unprotect(password?: string): void;
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
        getName(): string;
        getWorksheet(): Worksheet;
        getColumnByName(name: string): TableColumn;
        getRangeBetweenHeaderAndTotal(): Range;
        getHeaderRowRange(): Range;
        getRange(): Range;
        getRowCount(): number;
        getColumns(): TableColumn[];
        addRow(index: number, values: (string | number | boolean)[]): void;
        getAutoFilter(): AutoFilter;
    }

    interface TableColumn {
        getName(): string;
        getIndex(): number;
        getRangeBetweenHeaderAndTotal(): Range;
    }

    interface NamedItem {
        getName(): string;
        getRange(): Range;
    }

    interface Range {
        getText(): string;
        getValue(): string | number | boolean;
        getValues(): (string | number | boolean)[][];
        setValue(value: string | number | boolean | null): void;
        setValues(values: (string | number | boolean | null)[][]): void;
        clear(applyTo?: ClearApplyTo): void;
        getFormat(): RangeFormat;
        getCell(row: number, column: number): Range;
        merge(across: boolean): void;
        getRowCount(): number;
        getColumnCount(): number;
        getRowIndex(): number;
        getRow(rowIndex: number): Range;
        getLastRow(): Range;
        getIntersection(anotherRange: Range | string): Range;
        getResizedRange(deltaRows: number, deltaColumns: number): Range;
        getOffsetRange(rowOffset: number, columnOffset: number): Range;
        getUsedRange(valuesOnly?: boolean): Range;
        select(): void;
        setNumberFormatLocal(numberFormat: string | string[][]): void;
        getLastCell(): Range;
    }

    interface RangeFormat {
        getFill(): Fill;
        getFont(): RangeFont;
        getProtection(): FormatProtection;
        setHorizontalAlignment(alignment: HorizontalAlignment): void;
        setVerticalAlignment(alignment: VerticalAlignment): void;
        setWrapText(wrap: boolean): void;
        autofitColumns(): void;
        autofitRows(): void;
    }

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

    interface Range {
        // Agregar esta línea dentro de interface Range
        delete(shiftDirection: DeleteShiftDirection): void;
        getText(): string;
        getTexts(): string[][]; // Agrega esta línea
        getValue(): string | number | boolean;
    }

    enum DeleteShiftDirection {
        up = "Up",
        left = "Left"
    }
    
    enum SheetVisibility {
        visible = "Visible",
        hidden = "Hidden",
        veryHidden = "VeryHidden"
    }

    interface FilterCriteria {
        criterion1?: string;
        criterion2?: string;
        color?: string;
        operator?: string;
        filterOn?: string;
        values?: string[] | number[];
        dynamicCriteria?: unknown;
    }

    interface AutoFilter {
        clearCriteria(): void;
        apply(range: Range | string, columnIndex?: number, criteria?: FilterCriteria): void;
    }
}