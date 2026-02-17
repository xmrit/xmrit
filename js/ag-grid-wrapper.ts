/**
 * AG Grid wrapper module for Xmrit
 */
import dayjs from "dayjs";
import {
  type GridApi,
  type GridOptions,
  type ColDef,
  type CellValueChangedEvent,
  type IHeaderComp,
  type IHeaderParams,
  createGrid,
  ModuleRegistry,
  AllCommunityModule,
  themeAlpine,
} from "ag-grid-community";

// Register all community modules
ModuleRegistry.registerModules([AllCommunityModule]);

const compactTheme = themeAlpine.withParams({
  fontSize: 12,
  spacing: 3,
  headerFontSize: 12,
});

// ── Types ──

export interface XmritGridConfig<T = any> {
  element: HTMLElement;
  data: T[];
  columnDefs: ColDef<T>[];
  colHeaders?: string[];
  onCellValueChanged?: (event: CellValueChangedEvent<T>) => void;
  minSpareRows?: boolean; // whether to keep a spare empty row at the bottom
  dataSchema?: T;
  /** If true, allow right-click row insert/remove */
  contextMenu?: boolean;
  /** If false, don't allow insert/remove rows from context menu */
  allowInsertRow?: boolean;
  /** If true, show row numbers */
  rowHeaders?: boolean;
}

export interface XmritGrid<T = any> {
  api: GridApi<T>;
  updateSettings(opts: { data?: T[]; colHeaders?: string[] }): void;
  updateData(data: T[]): void;
  getColHeaders(): string[];
  destroy(): void;
}

// ── Context Menu ──

let activeContextMenu: HTMLElement | null = null;

function removeContextMenu() {
  if (activeContextMenu) {
    activeContextMenu.remove();
    activeContextMenu = null;
  }
  document.removeEventListener("click", removeContextMenu);
}

function showContextMenu<T>(
  event: MouseEvent,
  api: GridApi<T>,
  rowIndex: number,
  allowInsertRow: boolean,
  dataSchema: T | undefined,
  ensureSpareRowFn: (() => void) | null
) {
  removeContextMenu();
  event.preventDefault();

  const menu = document.createElement("div");
  menu.className = "xmrit-context-menu";

  type MenuItem = { label: string; action: () => void; disabled?: boolean } | "separator";
  const items: MenuItem[] = [];

  // Row operations
  if (allowInsertRow) {
    items.push({
      label: "Insert row above",
      action: () => {
        const newRow = dataSchema
          ? JSON.parse(JSON.stringify(dataSchema))
          : ({} as any);
        api.applyTransaction({ add: [newRow], addIndex: rowIndex });
        ensureSpareRowFn?.();
      },
    });
    items.push({
      label: "Insert row below",
      action: () => {
        const newRow = dataSchema
          ? JSON.parse(JSON.stringify(dataSchema))
          : ({} as any);
        api.applyTransaction({ add: [newRow], addIndex: rowIndex + 1 });
        ensureSpareRowFn?.();
      },
    });
  }

  items.push({
    label: "Remove row",
    action: () => {
      const rowNode = api.getDisplayedRowAtIndex(rowIndex);
      if (rowNode) {
        api.applyTransaction({ remove: [rowNode.data] });
        ensureSpareRowFn?.();
      }
    },
  });

  items.push("separator");

  // Undo / Redo
  items.push({
    label: "Undo",
    action: () => { api.undoCellEditing(); },
    disabled: api.getCurrentUndoSize() === 0,
  });
  items.push({
    label: "Redo",
    action: () => { api.redoCellEditing(); },
    disabled: api.getCurrentRedoSize() === 0,
  });


  items.forEach((item) => {
    if (item === "separator") {
      const hr = document.createElement("div");
      hr.className = "xmrit-context-menu-separator";
      menu.appendChild(hr);
      return;
    }
    const div = document.createElement("div");
    div.className = "xmrit-context-menu-item";
    if (item.disabled) div.classList.add("disabled");
    div.textContent = item.label;
    div.addEventListener("click", () => {
      if (!item.disabled) item.action();
      removeContextMenu();
    });
    menu.appendChild(div);
  });

  menu.style.left = `${event.clientX}px`;
  menu.style.top = `${event.clientY}px`;
  document.body.appendChild(menu);
  activeContextMenu = menu;

  // Close on next click anywhere
  setTimeout(() => {
    document.addEventListener("click", removeContextMenu);
  }, 0);
}

// ── Clickable Header Component (for main table editable headers) ──

export class ClickableHeader implements IHeaderComp {
  private eGui!: HTMLElement;
  private params!: IHeaderParams;

  init(params: IHeaderParams): void {
    this.params = params;
    this.eGui = document.createElement("div");
    this.eGui.className = "xmrit-clickable-header";
    this.eGui.textContent = params.displayName;
    this.eGui.style.cursor = "pointer";
    this.eGui.style.width = "100%";
    this.eGui.addEventListener("click", () => {
      const onHeaderClick = (params as any).onHeaderClick;
      if (onHeaderClick) {
        onHeaderClick(params.column.getColId(), params.displayName);
      }
    });
  }

  getGui(): HTMLElement {
    return this.eGui;
  }

  refresh(params: IHeaderParams): boolean {
    this.params = params;
    this.eGui.textContent = params.displayName;
    return true;
  }

  destroy(): void {}
}

// ── Spare Row Helper ──

function createEnsureSpareRows<T>(
  api: GridApi<T>,
  dataSchema: T | undefined,
  enabled: boolean
): (() => void) | null {
  if (!enabled) return null;
  return () => {
    const rowCount = api.getDisplayedRowCount();
    if (rowCount === 0) {
      const newRow = dataSchema
        ? JSON.parse(JSON.stringify(dataSchema))
        : ({} as any);
      api.applyTransaction({ add: [newRow] });
      return;
    }
    const lastRow = api.getDisplayedRowAtIndex(rowCount - 1);
    if (!lastRow) return;
    const data = lastRow.data as any;
    // Check if last row has any non-null/non-empty values
    const hasData = Object.values(data).some(
      (v) => v !== null && v !== undefined && v !== ""
    );
    if (hasData) {
      const newRow = dataSchema
        ? JSON.parse(JSON.stringify(dataSchema))
        : ({} as any);
      api.applyTransaction({ add: [newRow] });
    }
  };
}

// ── Factory ──

export function createXmritGrid<T>(config: XmritGridConfig<T>): XmritGrid<T> {
  const {
    element,
    data,
    columnDefs,
    colHeaders,
    onCellValueChanged,
    minSpareRows = false,
    dataSchema,
    contextMenu = true,
    allowInsertRow = true,
    rowHeaders = false,
  } = config;

  let currentColHeaders = colHeaders ? [...colHeaders] : [];

  // Build column defs with headers
  const buildColDefs = (headers: string[]): ColDef<T>[] => {
    return columnDefs.map((col, i) => ({
      ...col,
      headerName: headers[i] ?? col.headerName ?? (col as any).field ?? "",
    }));
  };

  // Add row number column if needed
  const finalColDefs = rowHeaders
    ? [
        {
          headerName: "#",
          valueGetter: "node.rowIndex + 1",
          width: 50,
          editable: false,
          sortable: false,
          suppressMovable: true,
        } as ColDef<T>,
        ...buildColDefs(currentColHeaders),
      ]
    : buildColDefs(currentColHeaders);

  const gridOptions: GridOptions<T> = {
    theme: compactTheme,
    columnDefs: finalColDefs,
    rowData: [...data],
    defaultColDef: {
      flex: 1,
      editable: true,
      sortable: false,
      filter: false,
      suppressMovable: true,
    },
    domLayout: "autoHeight",
    undoRedoCellEditing: true,
    preventDefaultOnContextMenu: true,
    singleClickEdit: true,
    stopEditingWhenCellsLoseFocus: true,
    suppressDragLeaveHidesColumns: true,
    rowDragManaged: false,
    getRowId: undefined, // use index-based identity
    onCellValueChanged: (event) => {
      ensureSpareRowFn?.();
      onCellValueChanged?.(event);
    },
    onCellContextMenu: (event) => {
      if (contextMenu && event.event) {
        showContextMenu(
          event.event as MouseEvent,
          api,
          event.rowIndex!,
          allowInsertRow,
          dataSchema,
          ensureSpareRowFn
        );
      }
    },
  };

  const api = createGrid(element, gridOptions);
  const ensureSpareRowFn = createEnsureSpareRows(api, dataSchema, minSpareRows);

  // Ensure initial spare row
  ensureSpareRowFn?.();

  const grid: XmritGrid<T> = {
    api,
    updateSettings(opts) {
      if (opts.colHeaders) {
        currentColHeaders = [...opts.colHeaders];
      }
      if (opts.data !== undefined) {
        api.setGridOption("rowData", [...opts.data]);
      }
      // Always update column defs when headers change
      if (opts.colHeaders) {
        const newDefs = rowHeaders
          ? [
              {
                headerName: "#",
                valueGetter: "node.rowIndex + 1",
                width: 50,
                editable: false,
                sortable: false,
                suppressMovable: true,
              } as ColDef<T>,
              ...buildColDefs(currentColHeaders),
            ]
          : buildColDefs(currentColHeaders);
        api.setGridOption("columnDefs", newDefs);
      }
      ensureSpareRowFn?.();
    },
    updateData(newData) {
      api.setGridOption("rowData", [...newData]);
      ensureSpareRowFn?.();
    },
    getColHeaders() {
      return [...currentColHeaders];
    },
    destroy() {
      api.destroy();
    },
  };

  return grid;
}

// ── Date Extension Helper (for "Extend dates" button) ──

/**
 * Detects date pattern from existing dates and generates new dates.
 * Uses dayjs for consistent timezone handling with the rest of the app.
 */
export function extendDateSeries(
  dates: string[],
  count: number
): string[] {
  const validDates = dates.filter((d) => d);
  if (validDates.length === 0) return [];

  const dateObjs = validDates.map((d) => dayjs(d, "YYYY-MM-DD"));

  // Detect interval in days between consecutive dates
  let diffDays = 1; // default: one day
  if (dateObjs.length >= 2) {
    diffDays = dateObjs[1].diff(dateObjs[0], "day");
    // Verify consistent pattern
    for (let i = 2; i < dateObjs.length; i++) {
      if (dateObjs[i].diff(dateObjs[i - 1], "day") !== diffDays) {
        // Inconsistent pattern, fall back to diff between last two
        diffDays = dateObjs[dateObjs.length - 1].diff(
          dateObjs[dateObjs.length - 2],
          "day"
        );
        break;
      }
    }
  }

  const result: string[] = [];
  let last = dateObjs[dateObjs.length - 1];
  for (let i = 0; i < count; i++) {
    last = last.add(diffDays, "day");
    result.push(last.format("YYYY-MM-DD"));
  }
  return result;
}
