import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import { createRoot } from "react-dom/client";
import { HierarchyTree as HierarchyTreeComponent, MPCacheEntry } from "./HierarchyTree";

/** Snapshot of a single record's column values – detached from the live dataset. */
interface RecordSnapshot {
    id: string;
    values: Record<string, unknown>;
}

export class HierarchyTree
    implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private root: ReturnType<typeof createRoot>;
    private notifyOutputChanged: () => void;

    private selectedLeafNames: string = "";
    private selectedLeafIds:   string = "";

    // — Accumulator for paged loading —
    private accumulatedRecords: RecordSnapshot[] = [];
    private _seenIds:           Set<string>      = new Set();
    private isLoadingPages:     boolean          = false;
    private pageCount:          number           = 0;

    // — MP Cache —
    // Built ONCE after all pages are loaded. Cleared on refresh.
    // Structure: path → MPCacheEntry
    // Example key: "GBS1|||GBS2|||GBS3"
    private _mpCache:  Map<string, MPCacheEntry> = new Map();
    private _mpBuilt:  boolean                   = false;

    // — Refresh guard —
    private _context:        ComponentFramework.Context<IInputs>;
    private _isRefreshing:   boolean = false;
    private _refreshVersion: number  = 0;

    // ─────────────────────────────────────────────
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.container = container;
        this.notifyOutputChanged = notifyOutputChanged;
        this.container.style.cssText =
            "width:100%; height:100%; box-sizing:border-box;";
        this.root = createRoot(container);
    }

    // ─────────────────────────────────────────────
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Always store latest context so onRefresh callback can use it
        this._context = context;

        // Reset refresh guard — new updateView means refresh completed
        this._isRefreshing = false;

        const dataset = context.parameters.tableGrid;

        // Guard: dataset still loading → show spinner, wait
        if (dataset.loading) {
            this.renderLoading("Loading data...");
            return;
        }

        // Reset accumulator at the START of a fresh load cycle (not mid-paging)
        if (!this.isLoadingPages) {
            this.accumulatedRecords = [];
            this._seenIds           = new Set();
            this.pageCount          = 0;
            // NOTE: _mpCache is NOT cleared here — it gets cleared only on
            // explicit refresh. This prevents unnecessary rebuilds on re-renders
            // triggered by unrelated Power Apps events.
        }

        // Set a large page size to minimise round-trips
        if (dataset.paging && dataset.paging.pageSize !== 5000) {
            dataset.paging.setPageSize(5000);
        }

        // — Snapshot current page's records (deduplicated by record ID) —
        const currentIds  = dataset.sortedRecordIds;
        const prevCount   = this.accumulatedRecords.length;

        for (const rid of currentIds) {
            if (this._seenIds.has(rid)) continue;
            const rec = dataset.records[rid];
            if (!rec) continue;
            const recId = rec.getRecordId();
            if (this._seenIds.has(recId)) continue;

            this._seenIds.add(rid);
            this._seenIds.add(recId);

            const snap: RecordSnapshot = { id: recId, values: {} };

            // Grab every column the dataset exposes
            for (const col of dataset.columns) {
                try {
                    snap.values[col.name] = rec.getValue(col.name);
                } catch { /* column not available */ }
            }

            // Explicitly capture columns needed for tree/output that may not
            // appear in dataset.columns but ARE accessible via getValue()
            const extraCols = ["STRATEGY_NAME", "STRATEGY_ID", "STRATEGY_TYPE_KEY"];
            for (let i = 1; i <= 15; i++) extraCols.push(`GBS_LEVEL_${i}`);
            for (const colName of extraCols) {
                if (snap.values[colName] === undefined) {
                    try {
                        snap.values[colName] = rec.getValue(colName);
                    } catch { /* column not available */ }
                }
            }

            this.accumulatedRecords.push(snap);
        }

        this.pageCount++;
        const newRecordsAdded = this.accumulatedRecords.length > prevCount;

        // — Load next page if available —
        const shouldLoadMore =
            dataset.paging &&
            dataset.paging.hasNextPage &&
            newRecordsAdded;

        if (shouldLoadMore) {
            this.isLoadingPages = true;
            // Progressive render during paging — no cache yet, shows raw records
            this._renderTree(context, this.accumulatedRecords, true);
            dataset.paging.loadNextPage(); // triggers another updateView
            return;
        }

        // — All pages loaded — build MP cache then final render —
        this.isLoadingPages = false;
        this._buildMPCache(this.accumulatedRecords); // ← build/update cache
        this._renderTree(context, this.accumulatedRecords, false);
    }

    // ─────────────────────────────────────────────
    /**
     * Build the Materialized Path cache from all accumulated records.
     *
     * This runs ONCE after paging completes (or on refresh).
     * Instead of scanning 30-35k raw records every time filters change,
     * the tree is built from this pre-parsed cache in O(unique paths).
     *
     * Cache structure example for one row with 3 filled levels:
     *   GBS_LEVEL_1="Alpha", GBS_LEVEL_2="Beta", GBS_LEVEL_3="Gamma"
     *   → "Alpha"         : { depth:0, parentPath:"",           label:"Alpha", ... }
     *   → "Alpha|||Beta"  : { depth:1, parentPath:"Alpha",      label:"Beta",  ... }
     *   → "Alpha|||Beta|||Gamma" : { depth:2, parentPath:"Alpha|||Beta", label:"Gamma", strategyName:"...", ... }
     */
    private _buildMPCache(records: RecordSnapshot[]): void {
        // Always do a full rebuild — called only after all pages loaded
        const newCache = new Map<string, MPCacheEntry>();

        for (const rec of records) {
            const strategyTypeKey = Number(rec.values["STRATEGY_TYPE_KEY"] ?? 0);
            const strategyName    = String(rec.values["STRATEGY_NAME"]     ?? "").trim();
            const strategyId      = String(rec.values["STRATEGY_ID"]       ?? "").trim();
            const level1Value     = String(rec.values["GBS_LEVEL_1"]       ?? "").trim();

            let parentPath = "";
            let lastPath   = "";
            let isLeaf     = false;

            for (let i = 1; i <= 15; i++) {
                const val = String(rec.values[`GBS_LEVEL_${i}`] ?? "").trim();
                if (!val || val.toLowerCase() === "null") break;

                const path  = parentPath ? `${parentPath}|||${val}` : val;
                isLeaf      = false; // will be set true after loop if no deeper level

                if (!newCache.has(path)) {
                    newCache.set(path, {
                        path,
                        label:           val,
                        depth:           i - 1,
                        parentPath,
                        strategyTypeKey,
                        level1Value,
                        recordId:        rec.id,
                        // strategyName/strategyId filled below on leaf
                    });
                }

                parentPath = path;
                lastPath   = path;
                isLeaf     = true;
            }

            // Attach strategy metadata to the leaf node
            if (lastPath && isLeaf) {
                const entry = newCache.get(lastPath);
                if (entry && !entry.strategyName) {
                    entry.strategyName = strategyName;
                    entry.strategyId   = strategyId;
                }
            }
        }

        this._mpCache = newCache;
        this._mpBuilt = true;

        console.log(
            `[HierarchyTree] MP cache built: ${this._mpCache.size} unique paths from ${records.length} records`
        );
    }

    // ─────────────────────────────────────────────
    /** Render the tree component with the given records */
    private _renderTree(
        context: ComponentFramework.Context<IInputs>,
        allRecords: RecordSnapshot[],
        isLoadingMore: boolean
    ): void {
        const strategyTypeKey = Number(
            context.parameters.strategyTypeKey?.raw ?? 0
        );

        this.root.render(
            React.createElement(HierarchyTreeComponent, {
                key:            this._refreshVersion,   // changes only on refresh → full remount
                records:        [...allRecords],         // new reference each render
                totalRows:      allRecords.length,
                maxLevels:      Number(context.parameters.maxLevels?.raw ?? 15),
                filterId:       String(context.parameters.filterId?.raw   ?? "").trim(),
                filterName:     String(context.parameters.filterName?.raw ?? "").trim(),
                strategyTypeKey,
                isLoadingMore,
                // Pass cache only when it's fully built (not during paging)
                // During paging, mpCache is undefined → HierarchyTree falls back
                // to raw records scan (same behaviour as before)
                mpCache: (!isLoadingMore && this._mpBuilt) ? this._mpCache : undefined,
                onSelectionChange: (names: string[], ids: string[]) => {
                    this.selectedLeafNames = names.join(",");
                    this.selectedLeafIds   = ids.join(",");
                    this.notifyOutputChanged();
                },
                onRefresh: () => {
                    // Guard: prevent double-tap / parallel refresh
                    if (this._isRefreshing) return;
                    this._isRefreshing = true;

                    // Increment version → React remounts HierarchyTree cleanly
                    this._refreshVersion++;

                    // ── Full reset ──
                    // Records accumulator cleared
                    this.accumulatedRecords = [];
                    this._seenIds           = new Set();
                    this.pageCount          = 0;
                    this.isLoadingPages     = false;

                    // MP cache cleared — will be rebuilt after fresh data loads
                    this._mpCache = new Map();
                    this._mpBuilt = false;

                    // Trigger Power Apps to re-fetch from SQL via connector
                    this._context.parameters.tableGrid.refresh();
                },
            })
        );
    }

    // ─────────────────────────────────────────────
    /** Show a loading indicator while pages are still coming in */
    private renderLoading(message: string): void {
        this.root.render(
            React.createElement(
                "div",
                {
                    style: {
                        display:        "flex",
                        flexDirection:  "column",
                        alignItems:     "center",
                        justifyContent: "center",
                        height:         "100%",
                        fontFamily:     "Segoe UI, sans-serif",
                        color:          "#605E5C",
                        gap:            "12px",
                    },
                },
                React.createElement("div", {
                    style: {
                        width:        32,
                        height:       32,
                        border:       "3px solid #EDEBE9",
                        borderTop:    "3px solid #0078D4",
                        borderRadius: "50%",
                        animation:    "pcf-spin 1s linear infinite",
                    },
                }),
                React.createElement(
                    "style",
                    null,
                    "@keyframes pcf-spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }"
                ),
                React.createElement(
                    "span",
                    { style: { fontSize: 13 } },
                    message
                )
            )
        );
    }

    // ─────────────────────────────────────────────
    public getOutputs(): IOutputs {
        return {
            selectedLeafNames: this.selectedLeafNames,
            selectedLeafIds:   this.selectedLeafIds,
        };
    }

    public destroy(): void {
        this.root.unmount();
    }
}
