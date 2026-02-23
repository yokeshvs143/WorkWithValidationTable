import { ReactElement, createElement, useState, useCallback, useEffect, useRef } from "react";

import classNames from "classnames";
import { TableviewContainerProps } from "../typings/TableviewProps";
import Big from "big.js";
import "./ui/Tableview.css";

interface CellObject {
    id: string;
    sequenceNumber: string;
    isBlocked: boolean;
    isMerged: boolean;
    mergeId: string;
    isBlank: boolean;
    rowIndex: number;
    columnIndex: number;
    // isSelected removed — selection is tracked externally via selectedCells Set
}

interface TableRow {
    id: string;
    rowIndex: number;
    cells: CellObject[];
}

interface TableData {
    rows: number;
    columns: number;
    environment: string;
    tableRows: TableRow[];
    mergeSpans?: Record<string, { rowSpan: number; colSpan: number; anchorRow: number; anchorCol: number }>;
    metadata?: {
        createdAt?: string;
        updatedAt?: string;
    };
}

interface MergeSpanInfo {
    rowSpan: number;
    colSpan: number;
    anchorRow: number;
    anchorCol: number;
}

const computeMergeSpans = (rows: TableRow[]): Record<string, MergeSpanInfo> => {
    const groupedByMerge: Record<string, { row: number; col: number }[]> = {};
    rows.forEach(row => {
        row.cells.forEach(cell => {
            if (!cell.isMerged || !cell.mergeId) return;
            if (!groupedByMerge[cell.mergeId]) groupedByMerge[cell.mergeId] = [];
            groupedByMerge[cell.mergeId].push({ row: cell.rowIndex, col: cell.columnIndex });
        });
    });
    const result: Record<string, MergeSpanInfo> = {};
    Object.entries(groupedByMerge).forEach(([mergeId, positions]) => {
        const rows2 = positions.map(p => p.row);
        const cols = positions.map(p => p.col);
        const minRow = Math.min(...rows2);
        const maxRow = Math.max(...rows2);
        const minCol = Math.min(...cols);
        const maxCol = Math.max(...cols);
        result[mergeId] = {
            rowSpan: maxRow - minRow + 1,
            colSpan: maxCol - minCol + 1,
            anchorRow: minRow,
            anchorCol: minCol
        };
    });
    return result;
};

const isCellHidden = (cell: CellObject, mergeSpans: Record<string, MergeSpanInfo>): boolean => {
    if (!cell.isMerged || !cell.mergeId) return false;
    const span = mergeSpans[cell.mergeId];
    if (!span) return false;
    return !(cell.rowIndex === span.anchorRow && cell.columnIndex === span.anchorCol);
};

const getCellSpan = (
    cell: CellObject,
    mergeSpans: Record<string, MergeSpanInfo>
): { rowSpan: number; colSpan: number } => {
    if (!cell.isMerged || !cell.mergeId) return { rowSpan: 1, colSpan: 1 };
    const span = mergeSpans[cell.mergeId];
    if (!span) return { rowSpan: 1, colSpan: 1 };
    if (cell.rowIndex === span.anchorRow && cell.columnIndex === span.anchorCol) {
        return { rowSpan: span.rowSpan, colSpan: span.colSpan };
    }
    return { rowSpan: 1, colSpan: 1 };
};

const Tableview = (props: TableviewContainerProps): ReactElement => {
    const getInitialRows = () => {
        if (props.rowCountAttribute?.status === "available" && props.rowCountAttribute.value) {
            return Number(props.rowCountAttribute.value);
        }
        return 3;
    };

    const getInitialColumns = () => {
        if (props.columnCountAttribute?.status === "available" && props.columnCountAttribute.value) {
            return Number(props.columnCountAttribute.value);
        }
        return 3;
    };

    const [rowCount, setRowCount] = useState<number>(getInitialRows());
    const [columnCount, setColumnCount] = useState<number>(getInitialColumns());
    const [environment, setEnvironment] = useState<string>((props as any).environmentAttribute?.value || "");
    const [tableRows, setTableRows] = useState<TableRow[]>([]);
    const [mergeSpans, setMergeSpans] = useState<Record<string, MergeSpanInfo>>({});
    const [selectedCells, setSelectedCells] = useState<Set<string>>(new Set());
    const [isSelectionMode, setIsSelectionMode] = useState<boolean>(false);

    // Drag selection
    const [isDragging, setIsDragging] = useState<boolean>(false);
    const [dragStartCell, setDragStartCell] = useState<{ row: number; col: number } | null>(null);
    const dragSelectionRef = useRef<Set<string>>(new Set());
    const preSelectionRef = useRef<Set<string>>(new Set());

    const [isInitialLoad, setIsInitialLoad] = useState<boolean>(true);
    const [isSaving, setIsSaving] = useState<boolean>(false);
    const [dataLoaded, setDataLoaded] = useState<boolean>(false);
    const lastSavedDataRef = useRef<string>("");
    const isUserInputRef = useRef<boolean>(false);
    const ignoreAttributeUpdateRef = useRef<boolean>(false);

    // ── Feature flags ─────────────────────────────────────────────────────────
    const hasBlankingEnabled = !!(props as any).enableCellBlanking;
    const hasMergingEnabled = !!props.enableCellMerging;
    const isSelectionAllowed = hasMergingEnabled || hasBlankingEnabled;

    // ── Keep mergeSpans in sync whenever tableRows change ─────────────────────
    useEffect(() => {
        setMergeSpans(computeMergeSpans(tableRows));
    }, [tableRows]);

    // ── Statistics ────────────────────────────────────────────────────────────
    const updateCellStatistics = useCallback(
        (rows: TableRow[], spans: Record<string, MergeSpanInfo>) => {
            const totalCells = rows.reduce((sum, row) => sum + row.cells.length, 0);
            const blockedCells = rows.reduce((sum, row) => sum + row.cells.filter(c => c.isBlocked).length, 0);
            const mergedCells = rows.reduce(
                (sum, row) => sum + row.cells.filter(c => c.isMerged && !isCellHidden(c, spans)).length,
                0
            );
            const blankCells = rows.reduce(
                (sum, row) => sum + row.cells.filter(c => c.isBlank && !isCellHidden(c, spans)).length,
                0
            );

            if (props.totalCellsAttribute?.status === "available") props.totalCellsAttribute.setValue(new Big(totalCells));
            if (props.blockedCellsAttribute?.status === "available") props.blockedCellsAttribute.setValue(new Big(blockedCells));
            if (props.mergedCellsAttribute?.status === "available") props.mergedCellsAttribute.setValue(new Big(mergedCells));
            if ((props as any).blankCellsAttribute?.status === "available") (props as any).blankCellsAttribute.setValue(new Big(blankCells));
        },
        [props.totalCellsAttribute, props.blockedCellsAttribute, props.mergedCellsAttribute]
    );

    // ── Load data ─────────────────────────────────────────────────────────────
    useEffect(() => {
        if (isSaving) return;
        const incomingData = props.useAttributeData?.value || "";
        if (incomingData === lastSavedDataRef.current && lastSavedDataRef.current !== "") return;

        if (incomingData && incomingData !== "") {
            try {
                const tableData: TableData = JSON.parse(incomingData);
                if (tableData.tableRows && tableData.rows > 0 && tableData.columns > 0) {
                    const validatedRows = tableData.tableRows.map((row, idx) => {
                        const rowIndex = idx + 1;
                        return {
                            ...row,
                            id: `row_${rowIndex}`,
                            rowIndex,
                            cells: row.cells.map((cell, cIdx) => {
                                const colIndex = cIdx + 1;
                                const seqNum = cell.sequenceNumber || "-";
                                const validatedCell: CellObject = {
                                    id: `cell_${rowIndex}_${colIndex}`,
                                    sequenceNumber: seqNum,
                                    isBlocked: cell.isBlocked !== undefined ? cell.isBlocked : false,
                                    isMerged: cell.isMerged || false,
                                    mergeId: cell.mergeId || "",
                                    isBlank: cell.isBlank || false,
                                    rowIndex,
                                    columnIndex: colIndex
                                };
                                return validatedCell;
                            })
                        };
                    });

                    const loadedEnv = tableData.environment || (props as any).environmentAttribute?.value || "";
                    setEnvironment(loadedEnv);

                    setRowCount(tableData.rows);
                    setColumnCount(tableData.columns);
                    ignoreAttributeUpdateRef.current = true;
                    if (props.rowCountAttribute?.status === "available") props.rowCountAttribute.setValue(new Big(tableData.rows));
                    if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(tableData.columns));

                    const spans = computeMergeSpans(validatedRows);
                    setMergeSpans(spans);
                    setTableRows(validatedRows);
                    setSelectedCells(new Set());
                    setIsSelectionMode(false);
                    setDataLoaded(true);
                    updateCellStatistics(validatedRows, spans);
                    lastSavedDataRef.current = incomingData;
                    if (isInitialLoad) setTimeout(() => setIsInitialLoad(false), 500);
                }
            } catch (error) {
                console.error("Error loading table from attribute:", error);
                if (isInitialLoad) setTimeout(() => setIsInitialLoad(false), 500);
            }
        } else {
            if (isInitialLoad) setTimeout(() => setIsInitialLoad(false), 500);
        }
    }, [props.useAttributeData?.value, updateCellStatistics, isSaving, isInitialLoad, props.rowCountAttribute, props.columnCountAttribute]);

    // ── Sync environment attribute ────────────────────────────────────────────
    useEffect(() => {
        const envValue = (props as any).environmentAttribute?.value;
        if (!envValue) return;
        setEnvironment(envValue);
        if (tableRows.length > 0) {
            saveToBackend(tableRows, rowCount, columnCount, envValue);
        }
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [(props as any).environmentAttribute?.value]);

    useEffect(() => {
        if (ignoreAttributeUpdateRef.current) { ignoreAttributeUpdateRef.current = false; return; }
        if (props.rowCountAttribute?.status === "available" && props.rowCountAttribute.value != null) {
            const v = Number(props.rowCountAttribute.value);
            if (!isNaN(v) && v > 0 && v <= 100 && v !== rowCount && !isUserInputRef.current) setRowCount(v);
        }
    }, [props.rowCountAttribute?.value, rowCount]);

    useEffect(() => {
        if (ignoreAttributeUpdateRef.current) { ignoreAttributeUpdateRef.current = false; return; }
        if (props.columnCountAttribute?.status === "available" && props.columnCountAttribute.value != null) {
            const v = Number(props.columnCountAttribute.value);
            if (!isNaN(v) && v > 0 && v <= 100 && v !== columnCount && !isUserInputRef.current) setColumnCount(v);
        }
    }, [props.columnCountAttribute?.value, columnCount]);

    const createMergeId = (r1: number, c1: number, r2: number, c2: number) => `${r1}${c1}${r2}${c2}`;

    // ── Reset drag state — call after any merge/blank operation ───────────────
    const resetDragState = useCallback(() => {
        setIsDragging(false);
        setDragStartCell(null);
        dragSelectionRef.current = new Set();
        preSelectionRef.current = new Set();
    }, []);

    // ── Create table ──────────────────────────────────────────────────────────
    const createNewTable = useCallback((rows: number, cols: number) => {
        if (rows <= 0 || cols <= 0) return;
        const newTableRows: TableRow[] = Array.from({ length: rows }, (_, idx) => {
            const rowIndex = idx + 1;
            return {
                id: `row_${rowIndex}`,
                rowIndex,
                cells: Array.from({ length: cols }, (_, cIdx) => {
                    const colIndex = cIdx + 1;
                    return {
                        id: `cell_${rowIndex}_${colIndex}`,
                        sequenceNumber: "-",
                        isBlocked: false,
                        isMerged: false,
                        mergeId: "",
                        isBlank: false,
                        rowIndex,
                        columnIndex: colIndex
                    };
                })
            };
        });
        const spans = computeMergeSpans(newTableRows);
        setMergeSpans(spans);
        setTableRows(newTableRows);
        setSelectedCells(new Set());
        setIsSelectionMode(false);
        setDataLoaded(true);
        saveToBackend(newTableRows, rows, cols);
    }, []);

    useEffect(() => {
        const timer = setTimeout(() => {
            if (!dataLoaded && tableRows.length === 0) createNewTable(rowCount, columnCount);
        }, 100);
        return () => clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [dataLoaded]);

    // ── Save ──────────────────────────────────────────────────────────────────
    const saveToBackend = useCallback(
        (rows: TableRow[], rowCnt: number, colCnt: number, envOverride?: string) => {
            setIsSaving(true);
            const envValue = envOverride !== undefined ? envOverride : environment;
            const tableData: TableData = {
                rows: rowCnt,
                columns: colCnt,
                environment: envValue,
                tableRows: rows,
                metadata: { updatedAt: new Date().toISOString() }
            };
            const jsonData = JSON.stringify(tableData);
            lastSavedDataRef.current = jsonData;
            if (props.useAttributeData?.status === "available") props.useAttributeData.setValue(jsonData);
            if (props.tableDataAttribute?.status === "available") props.tableDataAttribute.setValue(jsonData);
            ignoreAttributeUpdateRef.current = true;
            if (props.rowCountAttribute?.status === "available") props.rowCountAttribute.setValue(new Big(rowCnt));
            if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(colCnt));
            const spans = computeMergeSpans(rows);
            updateCellStatistics(rows, spans);
            if (props.onTableChange?.canExecute) props.onTableChange.execute();
            setTimeout(() => setIsSaving(false), 100);
        },
        [environment, props.useAttributeData, props.tableDataAttribute, props.rowCountAttribute, props.columnCountAttribute, props.onTableChange, updateCellStatistics]
    );

    useEffect(() => {
        if (props.autoSave && tableRows.length > 0 && !isSaving) saveToBackend(tableRows, rowCount, columnCount);
    }, [tableRows, props.autoSave, saveToBackend, isSaving, rowCount, columnCount]);

    useEffect(() => {
        if (tableRows.length > 0) updateCellStatistics(tableRows, mergeSpans);
    }, [tableRows, mergeSpans, updateCellStatistics]);

    // ── Dimensions ────────────────────────────────────────────────────────────
    // ── Generate table: pending flag set on button click, resolved when generateResult attribute updates ──
    const pendingGenerateRef = useRef<boolean>(false);

    const applyDimensions = useCallback(() => {
        const onGenerateTable = (props as any).onGenerateTable;

        if (onGenerateTable?.canExecute) {
            // Mark that a generate is pending, then execute the microflow/nanoflow.
            // The MF/NF must set the generateResult boolean attribute to true or false.
            // The useEffect below watches that attribute and acts on the result.
            pendingGenerateRef.current = true;
            onGenerateTable.execute();
        } else {
            // No action configured — generate directly
            if (isNaN(rowCount) || isNaN(columnCount)) { alert("Please enter valid numbers"); return; }
            if (rowCount <= 0 || columnCount <= 0) { alert("Rows and columns must be positive numbers"); return; }
            if (rowCount > 100 || columnCount > 100) { alert("Maximum 100 rows and 100 columns"); return; }
            ignoreAttributeUpdateRef.current = true;
            if (props.rowCountAttribute?.status === "available") props.rowCountAttribute.setValue(new Big(rowCount));
            if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(columnCount));
            createNewTable(rowCount, columnCount);
        }
    }, [rowCount, columnCount, createNewTable, props.rowCountAttribute, props.columnCountAttribute]);

    // ── Watch generateResult attribute — react after MF/NF sets it ─────────────────
    useEffect(() => {
        const generateResult = (props as any).generateResult;
        if (!pendingGenerateRef.current) return;
        if (generateResult?.status !== "available") return;

        // Consume the pending flag immediately to avoid re-triggering
        pendingGenerateRef.current = false;

        if (generateResult.value === true) {
            // MF/NF returned true — generate the table
            if (isNaN(rowCount) || isNaN(columnCount)) { alert("Please enter valid numbers"); return; }
            if (rowCount <= 0 || columnCount <= 0) { alert("Rows and columns must be positive numbers"); return; }
            if (rowCount > 100 || columnCount > 100) { alert("Maximum 100 rows and 100 columns"); return; }
            ignoreAttributeUpdateRef.current = true;
            if (props.rowCountAttribute?.status === "available") props.rowCountAttribute.setValue(new Big(rowCount));
            if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(columnCount));
            createNewTable(rowCount, columnCount);
            // Reset the result attribute back to false so it's ready for next click
            generateResult.setValue(false);
        }
        // false → MF/NF already showed its own validation feedback; do nothing
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [(props as any).generateResult?.value]);

    // ── Add row ───────────────────────────────────────────────────────────────
    const addRow = useCallback(() => {
        const newRowCount = rowCount + 1;
        if (newRowCount > 100) { alert("Maximum 100 rows"); return; }
        isUserInputRef.current = true;
        setRowCount(newRowCount);
        ignoreAttributeUpdateRef.current = true;
        if (props.rowCountAttribute?.status === "available") props.rowCountAttribute.setValue(new Big(newRowCount));
        setTableRows(prevRows => {
            const newRows = [...prevRows];
            const rowIndex = newRowCount;
            newRows.push({
                id: `row_${rowIndex}`,
                rowIndex,
                cells: Array.from({ length: columnCount }, (_, cIdx) => {
                    const colIndex = cIdx + 1;
                    return {
                        id: `cell_${rowIndex}_${colIndex}`,
                        sequenceNumber: "-",
                        isBlocked: false,
                        isMerged: false,
                        mergeId: "",
                        isBlank: false,
                        rowIndex,
                        columnIndex: colIndex
                    };
                })
            });
            saveToBackend(newRows, newRowCount, columnCount);
            return newRows;
        });
        setTimeout(() => { isUserInputRef.current = false; }, 100);
    }, [rowCount, columnCount, props.rowCountAttribute, saveToBackend]);

    // ── Add column ────────────────────────────────────────────────────────────
    const addColumn = useCallback(() => {
        const newColCount = columnCount + 1;
        if (newColCount > 100) { alert("Maximum 100 columns"); return; }
        isUserInputRef.current = true;
        setColumnCount(newColCount);
        ignoreAttributeUpdateRef.current = true;
        if (props.columnCountAttribute?.status === "available") props.columnCountAttribute.setValue(new Big(newColCount));
        setTableRows(prevRows => {
            const newRows = prevRows.map(row => ({
                ...row,
                cells: [
                    ...row.cells,
                    {
                        id: `cell_${row.rowIndex}_${newColCount}`,
                        sequenceNumber: "-",
                        isBlocked: false,
                        isMerged: false,
                        mergeId: "",
                        isBlank: false,
                        rowIndex: row.rowIndex,
                        columnIndex: newColCount
                    }
                ]
            }));
            saveToBackend(newRows, rowCount, newColCount);
            return newRows;
        });
        setTimeout(() => { isUserInputRef.current = false; }, 100);
    }, [rowCount, columnCount, props.columnCountAttribute, saveToBackend]);

    // ── Cell value change ─────────────────────────────────────────────────────
    const handleCellValueChange = useCallback(
        (rowIndex: number, colIndex: number, newValue: string) => {
            setTableRows(prevRows => {
                const newRows = prevRows.map(row => ({ ...row, cells: row.cells.map(cell => ({ ...cell })) }));
                const targetCell = newRows.find(r => r.rowIndex === rowIndex)?.cells.find(c => c.columnIndex === colIndex);
                if (!targetCell) return prevRows;
                targetCell.sequenceNumber = newValue;
                if (targetCell.mergeId && targetCell.mergeId !== "") {
                    const mergeId = targetCell.mergeId;
                    newRows.forEach(row => row.cells.forEach(cell => {
                        if (cell.mergeId === mergeId) cell.sequenceNumber = newValue;
                    }));
                }
                const spans = computeMergeSpans(newRows);
                updateCellStatistics(newRows, spans);
                if (props.autoSave) saveToBackend(newRows, rowCount, columnCount);
                return newRows;
            });
            if (props.onCellClick?.canExecute) props.onCellClick.execute();
        },
        [props.onCellClick, props.autoSave, updateCellStatistics, saveToBackend, rowCount, columnCount]
    );

    // ── Checkbox (blocked toggle) ─────────────────────────────────────────────
    const handleCheckboxChange = useCallback(
        (rowIndex: number, colIndex: number) => {
            setTableRows(prevRows => {
                const newRows = prevRows.map(row => ({ ...row, cells: row.cells.map(cell => ({ ...cell })) }));
                const targetCell = newRows.find(r => r.rowIndex === rowIndex)?.cells.find(c => c.columnIndex === colIndex);
                if (!targetCell) return prevRows;
                const newBlocked = !targetCell.isBlocked;
                targetCell.isBlocked = newBlocked;
                if (targetCell.mergeId && targetCell.mergeId !== "") {
                    const mergeId = targetCell.mergeId;
                    newRows.forEach(row => row.cells.forEach(cell => {
                        if (cell.mergeId === mergeId) cell.isBlocked = newBlocked;
                    }));
                }
                const spans = computeMergeSpans(newRows);
                updateCellStatistics(newRows, spans);
                if (props.autoSave) saveToBackend(newRows, rowCount, columnCount);
                return newRows;
            });
            if (props.onCellClick?.canExecute) props.onCellClick.execute();
        },
        [props.onCellClick, props.autoSave, updateCellStatistics, saveToBackend, rowCount, columnCount]
    );

    // ── Rectangular selection ─────────────────────────────────────────────────
    const getRectangularSelection = useCallback(
        (startRow: number, startCol: number, endRow: number, endCol: number): Set<string> => {
            const minRow = Math.min(startRow, endRow);
            const maxRow = Math.max(startRow, endRow);
            const minCol = Math.min(startCol, endCol);
            const maxCol = Math.max(startCol, endCol);
            const selection = new Set<string>();
            for (let r = minRow; r <= maxRow; r++)
                for (let c = minCol; c <= maxCol; c++)
                    selection.add(`cell_${r}_${c}`);
            return selection;
        },
        []
    );

    // ── Drag select ───────────────────────────────────────────────────────────
    const handleCellMouseDown = useCallback(
        (rowIndex: number, colIndex: number, event: React.MouseEvent) => {
            if (!isSelectionAllowed) return;
            if ((event.target as HTMLElement).tagName === "INPUT") return;
            event.preventDefault();
            preSelectionRef.current = new Set(selectedCells);
            setIsDragging(true);
            setDragStartCell({ row: rowIndex, col: colIndex });
            setIsSelectionMode(true);
            const cellId = `cell_${rowIndex}_${colIndex}`;
            if (event.shiftKey) {
                dragSelectionRef.current = new Set([cellId]);
                setSelectedCells(prev => { const s = new Set(prev); s.add(cellId); return s; });
            } else {
                dragSelectionRef.current = new Set([cellId]);
                setSelectedCells(new Set([cellId]));
            }
        },
        [selectedCells, isSelectionAllowed]
    );

    const handleCellMouseEnter = useCallback(
        (rowIndex: number, colIndex: number) => {
            if (!isDragging || !dragStartCell) return;
            const dragged = getRectangularSelection(dragStartCell.row, dragStartCell.col, rowIndex, colIndex);
            dragSelectionRef.current = dragged;
            const final = new Set(preSelectionRef.current);
            dragged.forEach(c => final.add(c));
            setSelectedCells(final);
        },
        [isDragging, dragStartCell, getRectangularSelection]
    );

    useEffect(() => {
        const onUp = () => {
            if (isDragging) { setIsDragging(false); setDragStartCell(null); preSelectionRef.current = new Set(); }
        };
        document.addEventListener("mouseup", onUp);
        return () => document.removeEventListener("mouseup", onUp);
    }, [isDragging]);

    const handleCellClick = useCallback(
        (rowIndex: number, colIndex: number, event?: React.MouseEvent) => {
            if (!isSelectionAllowed) {
                if (props.onCellClick?.canExecute) props.onCellClick.execute();
                return;
            }
            if (isDragging) return;
            const cellId = `cell_${rowIndex}_${colIndex}`;
            if (props.onCellClick?.canExecute) props.onCellClick.execute();
            const isCtrlOrCmd = event?.ctrlKey || event?.metaKey;
            if (isSelectionMode) {
                setSelectedCells(prev => {
                    const s = new Set(prev);
                    if (isCtrlOrCmd) {
                        if (s.has(cellId) && s.size > 1) s.delete(cellId); else s.add(cellId);
                    } else {
                        if (s.has(cellId) && s.size === 1) return s;
                        else s.add(cellId);
                    }
                    return s;
                });
            } else {
                setSelectedCells(new Set([cellId]));
                setIsSelectionMode(true);
            }
        },
        [isSelectionMode, isDragging, props.onCellClick, isSelectionAllowed]
    );

    const selectAllCells = useCallback(() => {
        if (!isSelectionAllowed) return;
        const all = new Set<string>();
        tableRows.forEach(row => row.cells.forEach(cell => {
            if (!isCellHidden(cell, mergeSpans)) all.add(cell.id);
        }));
        setSelectedCells(all);
        setIsSelectionMode(true);
    }, [tableRows, mergeSpans, isSelectionAllowed]);

    const clearSelection = useCallback(() => {
        setSelectedCells(new Set());
        setIsSelectionMode(false);
    }, []);

    // ── Merge ─────────────────────────────────────────────────────────────────
    const mergeCells = useCallback(() => {
        if (selectedCells.size < 2) return;
        const positions = Array.from(selectedCells).map(id => {
            const parts = id.replace("cell_", "").split("_");
            return { row: parseInt(parts[0]), col: parseInt(parts[1]) };
        });
        const minRow = Math.min(...positions.map(p => p.row));
        const maxRow = Math.max(...positions.map(p => p.row));
        const minCol = Math.min(...positions.map(p => p.col));
        const maxCol = Math.max(...positions.map(p => p.col));
        if (selectedCells.size !== (maxRow - minRow + 1) * (maxCol - minCol + 1)) {
            alert("Please select a rectangular area to merge"); return;
        }
        setTableRows(prevRows => {
            const newRows = prevRows.map(row => ({ ...row, cells: row.cells.map(cell => ({ ...cell })) }));
            for (let r = minRow; r <= maxRow; r++) {
                for (let c = minCol; c <= maxCol; c++) {
                    const cell = newRows.find(row => row.rowIndex === r)?.cells.find(cl => cl.columnIndex === c);
                    if (cell?.isMerged && cell.mergeId) {
                        const oldId = cell.mergeId;
                        newRows.forEach(row => row.cells.forEach(c2 => {
                            if (c2.mergeId === oldId) {
                                c2.isMerged = false;
                                c2.mergeId = "";
                            }
                        }));
                    }
                }
            }
            const mergeId = createMergeId(minRow, minCol, maxRow, maxCol);
            const topLeft = newRows.find(r => r.rowIndex === minRow)?.cells.find(c => c.columnIndex === minCol);
            if (!topLeft) return prevRows;
            for (let r = minRow; r <= maxRow; r++) {
                for (let c = minCol; c <= maxCol; c++) {
                    const cell = newRows.find(row => row.rowIndex === r)?.cells.find(cl => cl.columnIndex === c);
                    if (!cell) continue;
                    cell.sequenceNumber = topLeft.sequenceNumber;
                    cell.isBlocked = topLeft.isBlocked;
                    cell.isBlank = topLeft.isBlank;
                    cell.isMerged = true;
                    cell.mergeId = mergeId;
                }
            }
            const spans = computeMergeSpans(newRows);
            updateCellStatistics(newRows, spans);
            saveToBackend(newRows, rowCount, columnCount);
            return newRows;
        });
        const sortedRows = Array.from(selectedCells).map(id => parseInt(id.replace("cell_", "").split("_")[0]));
        const sortedCols = Array.from(selectedCells).map(id => parseInt(id.replace("cell_", "").split("_")[1]));
        setSelectedCells(new Set([`cell_${Math.min(...sortedRows)}_${Math.min(...sortedCols)}`]));
        // Reset drag state so next drag starts completely fresh
        resetDragState();
    }, [selectedCells, updateCellStatistics, saveToBackend, rowCount, columnCount, resetDragState]);

    // ── Unmerge ───────────────────────────────────────────────────────────────
    const unmergeCells = useCallback(() => {
        if (selectedCells.size === 0) return;
        const cellId = Array.from(selectedCells)[0];
        const parts = cellId.replace("cell_", "").split("_");
        const rowIndex = parseInt(parts[0]);
        const colIndex = parseInt(parts[1]);
        setTableRows(prevRows => {
            const newRows = prevRows.map(row => ({ ...row, cells: row.cells.map(cell => ({ ...cell })) }));
            const target = newRows.find(r => r.rowIndex === rowIndex)?.cells.find(c => c.columnIndex === colIndex);
            if (!target?.isMerged) return prevRows;
            const mergeId = target.mergeId;
            newRows.forEach(row => row.cells.forEach(cell => {
                if (cell.mergeId === mergeId) {
                    cell.isMerged = false;
                    cell.mergeId = "";
                }
            }));
            const spans = computeMergeSpans(newRows);
            updateCellStatistics(newRows, spans);
            saveToBackend(newRows, rowCount, columnCount);
            return newRows;
        });
        // Reset drag state so next drag starts completely fresh
        resetDragState();
    }, [selectedCells, updateCellStatistics, saveToBackend, rowCount, columnCount, resetDragState]);

    // ── Blank / Unblank ───────────────────────────────────────────────────────
    const blankSelectedCells = useCallback(() => {
        if (selectedCells.size === 0) return;
        setTableRows(prevRows => {
            const newRows = prevRows.map(row => ({ ...row, cells: row.cells.map(cell => ({ ...cell })) }));
            selectedCells.forEach(cellId => {
                const parts = cellId.replace("cell_", "").split("_");
                const rowIndex = parseInt(parts[0]);
                const colIndex = parseInt(parts[1]);
                const cell = newRows.find(r => r.rowIndex === rowIndex)?.cells.find(c => c.columnIndex === colIndex);
                if (!cell || isCellHidden(cell, mergeSpans)) return;
                cell.isBlank = true;
                if (cell.mergeId && cell.mergeId !== "") {
                    const mergeId = cell.mergeId;
                    newRows.forEach(row => row.cells.forEach(c => { if (c.mergeId === mergeId) c.isBlank = true; }));
                }
            });
            const spans = computeMergeSpans(newRows);
            updateCellStatistics(newRows, spans);
            saveToBackend(newRows, rowCount, columnCount);
            return newRows;
        });
        setSelectedCells(new Set());
        setIsSelectionMode(false);
        // Reset drag state so next drag starts completely fresh
        resetDragState();
    }, [selectedCells, mergeSpans, updateCellStatistics, saveToBackend, rowCount, columnCount, resetDragState]);

    const unblankSelectedCells = useCallback(() => {
        if (selectedCells.size === 0) return;
        setTableRows(prevRows => {
            const newRows = prevRows.map(row => ({ ...row, cells: row.cells.map(cell => ({ ...cell })) }));
            selectedCells.forEach(cellId => {
                const parts = cellId.replace("cell_", "").split("_");
                const rowIndex = parseInt(parts[0]);
                const colIndex = parseInt(parts[1]);
                const cell = newRows.find(r => r.rowIndex === rowIndex)?.cells.find(c => c.columnIndex === colIndex);
                if (!cell || isCellHidden(cell, mergeSpans)) return;
                cell.isBlank = false;
                if (cell.mergeId && cell.mergeId !== "") {
                    const mergeId = cell.mergeId;
                    newRows.forEach(row => row.cells.forEach(c => { if (c.mergeId === mergeId) c.isBlank = false; }));
                }
            });
            const spans = computeMergeSpans(newRows);
            updateCellStatistics(newRows, spans);
            saveToBackend(newRows, rowCount, columnCount);
            return newRows;
        });
        setSelectedCells(new Set());
        setIsSelectionMode(false);
        // Reset drag state so next drag starts completely fresh
        resetDragState();
    }, [selectedCells, mergeSpans, updateCellStatistics, saveToBackend, rowCount, columnCount, resetDragState]);

    // ── Styles ────────────────────────────────────────────────────────────────
    const tableStyle = { borderColor: props.tableBorderColor || "#dee2e6" };
    const selectedCellStyle = { backgroundColor: props.selectedCellColor || "#cfe2ff" };
    const mergedCellStyle = { backgroundColor: props.mergedCellColor || "#e3f2fd", borderColor: "#2196f3" };
    const blockedCellStyle = { backgroundColor: "white", borderColor: "#fdd835" };

    const hasSelection = selectedCells.size > 0;

    // ── Render ────────────────────────────────────────────────────────────────
    return (
        <div className={classNames("tableview-container", props.class)} style={props.style}>

            {/* ══ Controls bar ══ */}
            {(props.showGenerateButton || (hasSelection && isSelectionAllowed)) && (
                <div className="tableview-controls">

                    {props.showGenerateButton && (
                        <button className="tableview-btn tableview-btn-primary" onClick={applyDimensions}>
                            Generate Table
                        </button>
                    )}

                    {hasSelection && isSelectionAllowed && (
                        createElement("div", { style: { display: "contents" } },

                            props.showGenerateButton && createElement("div", { className: "tableview-controls-divider" }),

                            createElement("p", { className: "tableview-selection-info" },
                                `${selectedCells.size} cell(s) selected`
                            ),

                            createElement("button", {
                                className: "tableview-btn tableview-btn-info",
                                onClick: selectAllCells,
                                title: "Select all cells"
                            }, "Select All"),

                            hasMergingEnabled && createElement("div", { style: { display: "contents" } },
                                createElement("button", {
                                    className: "tableview-btn tableview-btn-warning",
                                    onClick: mergeCells,
                                    disabled: selectedCells.size < 2
                                }, "Merge Selected"),
                                createElement("button", {
                                    className: "tableview-btn tableview-btn-danger",
                                    onClick: unmergeCells
                                }, "Unmerge")
                            ),

                            hasBlankingEnabled && createElement("div", { style: { display: "contents" } },
                                createElement("button", {
                                    className: "tableview-btn tableview-btn-blank",
                                    onClick: blankSelectedCells,
                                    title: "Hide selected cells visually — data is preserved"
                                }, "Blank"),
                                createElement("button", {
                                    className: "tableview-btn tableview-btn-unblank",
                                    onClick: unblankSelectedCells,
                                    title: "Restore selected blank cells back to normal"
                                }, "Unblank")
                            ),

                            createElement("button", {
                                className: "tableview-btn tableview-btn-secondary",
                                onClick: clearSelection
                            }, "Clear Selection")
                        )
                    )}
                </div>
            )}

            {/* ══ Table ══ */}
            <div className="tableview-table-section">
                {props.showAddColumnButton && (
                    <div className="tableview-add-column-container">
                        <button className="tableview-btn tableview-btn-add" onClick={addColumn} title="Add Column">+</button>
                    </div>
                )}

                <div className="tableview-table-row-wrapper">
                    {props.showAddRowButton && (
                        <div className="tableview-add-row-container">
                            <button className="tableview-btn tableview-btn-add" onClick={addRow} title="Add Row">+</button>
                        </div>
                    )}

                    <div
                        className="tableview-table-wrapper"
                        style={{ userSelect: isDragging ? "none" : "auto" }}
                    >
                        <table
                            className="tableview-table"
                            style={tableStyle}
                            data-rows={rowCount}
                            data-cols={columnCount}
                        >
                            <tbody>
                                {tableRows.map(row => (
                                    <tr key={row.id}>
                                        {row.cells.map(cell => {
                                            if (isCellHidden(cell, mergeSpans)) return null;

                                            const isSelected = selectedCells.has(cell.id);
                                            const { rowSpan, colSpan } = getCellSpan(cell, mergeSpans);

                                            const blankEdgeStyle: React.CSSProperties = {};
                                            if (cell.isBlank) {
                                                const cellToRight = row.cells.find(c => c.columnIndex === cell.columnIndex + 1);
                                                const rowBelow = tableRows.find(r => r.rowIndex === cell.rowIndex + 1);
                                                const cellBelow = rowBelow?.cells.find(c => c.columnIndex === cell.columnIndex);

                                                const isLastBlankRight = !cellToRight || !cellToRight.isBlank;
                                                const isLastBlankBottom = !cellBelow || !cellBelow.isBlank;

                                                if (isLastBlankRight) blankEdgeStyle.borderRight = "1px solid #dee2e6";
                                                if (isLastBlankBottom) blankEdgeStyle.borderBottom = "1px solid #dee2e6";
                                            }

                                            const cellInlineStyle = cell.isBlank
                                                ? blankEdgeStyle
                                                : isSelected
                                                    ? selectedCellStyle
                                                    : cell.isMerged
                                                        ? mergedCellStyle
                                                        : cell.isBlocked
                                                            ? blockedCellStyle
                                                            : {};

                                            return (
                                                <td
                                                    key={cell.id}
                                                    rowSpan={rowSpan}
                                                    colSpan={colSpan}
                                                    className={classNames("tableview-cell", {
                                                        "tableview-cell-merged": cell.isMerged && !cell.isBlank,
                                                        "tableview-cell-selected": isSelected && !cell.isBlank,
                                                        "tableview-cell-blocked": cell.isBlocked && !cell.isBlank,
                                                        "tableview-cell-blank": cell.isBlank,
                                                        "tableview-cell-dragging": isDragging && isSelectionAllowed
                                                    })}
                                                    onClick={e => {
                                                        if (props.enableCheckbox === true) {
                                                            handleCheckboxChange(cell.rowIndex, cell.columnIndex);
                                                        }
                                                        handleCellClick(cell.rowIndex, cell.columnIndex, e);
                                                    }}
                                                    onMouseDown={e => handleCellMouseDown(cell.rowIndex, cell.columnIndex, e)}
                                                    onMouseEnter={() => handleCellMouseEnter(cell.rowIndex, cell.columnIndex)}
                                                    style={cellInlineStyle}
                                                >
                                                    {!cell.isBlank && (
                                                        <div className="tableview-cell-content">
                                                            {props.enableCheckbox ? (
                                                                <input
                                                                    type="checkbox"
                                                                    className="tableview-checkbox"
                                                                    checked={cell.isBlocked}
                                                                    readOnly
                                                                    tabIndex={-1}
                                                                />
                                                            ) : (
                                                                <input
                                                                    type="checkbox"
                                                                    className="tableview-checkbox tableview-checkbox-readonly"
                                                                    checked={cell.isBlocked}
                                                                    readOnly
                                                                    tabIndex={-1}
                                                                />
                                                            )}
                                                            {props.enableCellEditing ? (
                                                                <input
                                                                    type="text"
                                                                    className="tableview-cell-input"
                                                                    value={cell.sequenceNumber}
                                                                    onChange={e => handleCellValueChange(cell.rowIndex, cell.columnIndex, e.target.value)}
                                                                    onClick={e => e.stopPropagation()}
                                                                    onMouseDown={e => e.stopPropagation()}
                                                                    placeholder="#"
                                                                />
                                                            ) : (
                                                                <span
                                                                    className="tableview-cell-value"
                                                                    title={cell.sequenceNumber}
                                                                >
                                                                    {cell.sequenceNumber}
                                                                </span>
                                                            )}
                                                        </div>
                                                    )}
                                                </td>
                                            );
                                        })}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            {/* ══ Info bar ══ */}
            <div className="tableview-info">
                <p><strong>Table:</strong> {rowCount} rows × {columnCount} columns = {rowCount * columnCount} cells</p>
                <p><strong>Blocked:</strong> {tableRows.reduce((s, row) => s + row.cells.filter(c => c.isBlocked).length, 0)}</p>
                <p><strong>Merged:</strong> {tableRows.reduce((s, row) => s + row.cells.filter(c => c.isMerged && !isCellHidden(c, mergeSpans)).length, 0)}</p>
                {hasBlankingEnabled && (
                    <p><strong>Blank:</strong> {tableRows.reduce((s, row) => s + row.cells.filter(c => c.isBlank && !isCellHidden(c, mergeSpans)).length, 0)}</p>
                )}
            </div>
        </div>
    );
};

export default Tableview;