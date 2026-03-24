import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class StrategyTree
  implements ComponentFramework.StandardControl<IInputs, IOutputs> {

  private container: HTMLDivElement;
  private treeDiv: HTMLDivElement;
  private searchInput: HTMLInputElement;
  private context: ComponentFramework.Context<IInputs>;

private readonly LEVEL_COLUMNS = [
  "cr3e9_gbslevel1",
  "cr3e9_gbslevel2",
  "cr3e9_gbslevel3",
  "cr3e9_gbslevel4",
  "cr3e9_gbslevel5",
  "cr3e9_gbslevel6",
  "cr3e9_gbslevel7",
  "cr3e9_gbslevel8",
  "cr3e9_gbslevel9",
  "cr3e9_gbslevel10",
  "cr3e9_gbslevel11",
  "cr3e9_gbslevel12",
  "cr3e9_gbslevel13",
  "cr3e9_gbslevel14",
  "cr3e9_gbslevel15"
];

  constructor() {
    // init
  }

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.context = context;
    this.container = container;
    this.container.style.cssText = `
      width:100%; height:100%; 
      font-family: Segoe UI, sans-serif;
      overflow: auto;
      box-sizing: border-box;
    `;

    // ---- Header ----
    const header = document.createElement("div");
    header.style.cssText = `
      background: #0078d4;
      color: white;
      padding: 12px 16px;
      font-size: 15px;
      font-weight: 600;
      letter-spacing: 0.3px;
    `;
    header.innerText = "Strategy Hierarchy";
    this.container.appendChild(header);

    // ---- Search ----
    const searchWrap = document.createElement("div");
    searchWrap.style.cssText = "padding: 10px 16px; background: #f3f2f1;";
    this.searchInput = document.createElement("input");
    this.searchInput.type = "text";
    this.searchInput.placeholder = "Search...";
    this.searchInput.style.cssText = `
      width: 100%; padding: 7px 12px;
      border: 1px solid #c8c6c4; border-radius: 4px;
      font-size: 13px; box-sizing: border-box;
      outline: none;
    `;
    this.searchInput.addEventListener("input", () => this.filterTree());
    searchWrap.appendChild(this.searchInput);
    this.container.appendChild(searchWrap);

    // ---- Tree Area ----
    this.treeDiv = document.createElement("div");
    this.treeDiv.style.cssText = "padding: 12px 16px;";
    this.container.appendChild(this.treeDiv);

    this.treeDiv.innerHTML = `<div style="color:#a19f9d; font-size:13px;">Loading...</div>`;
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.context = context;
    const dataset = context.parameters.tableGrid;

    if (dataset.loading) {
      this.treeDiv.innerHTML = `<div style="color:#a19f9d; font-size:13px;">Loading...</div>`;
      return;
    }

    if (dataset.sortedRecordIds.length === 0) {
      this.treeDiv.innerHTML = `
        <div style="color:#a19f9d; font-size:13px; padding:20px; text-align:center;">
          No data available. Please connect your data source.
        </div>`;
      return;
    }

    // Use the actual Power Apps dataset
    const tree = this.buildTree(dataset);
    this.renderTree(tree);
  }

  // ----------------------------------------------------------------
  // Wide format (GBS_LEVEL_3 ... GBS_LEVEL_12) → nested tree
  // ----------------------------------------------------------------
  private buildTree(dataset: ComponentFramework.PropertyTypes.DataSet): TreeNode[] {
    // nodeMap key = full path joined by "|||" to avoid collisions
    const nodeMap = new Map<string, TreeNode>();

    dataset.sortedRecordIds.forEach(id => {
      const record = dataset.records[id];
      let parentPath = "";

      this.LEVEL_COLUMNS.forEach(col => {
        let val = "";
        try { val = String(record.getValue(col) ?? "").trim(); } catch { return; }
        if (!val) return;

        const fullPath = parentPath ? `${parentPath}|||${val}` : val;

        if (!nodeMap.has(fullPath)) {
          nodeMap.set(fullPath, {
            label: val,
            fullPath,
            parentPath: parentPath || null,
            children: [],
            expanded: !parentPath, // root level open by default
            visible: true
          });
        }
        parentPath = fullPath;
      });
    });

    // Build tree structure
    const roots: TreeNode[] = [];
    nodeMap.forEach(node => {
      if (!node.parentPath) {
        roots.push(node);
      } else {
        const parent = nodeMap.get(node.parentPath);
        if (parent) parent.children.push(node);
      }
    });

    return roots;
  }

  // ----------------------------------------------------------------
  // Render tree as HTML
  // ----------------------------------------------------------------
  private allNodes: TreeNode[] = [];

  private renderTree(roots: TreeNode[]): void {
    this.allNodes = [];
    this.collectNodes(roots, this.allNodes);
    this.treeDiv.innerHTML = "";
    const ul = this.buildUL(roots, 0);
    this.treeDiv.appendChild(ul);
  }

  private collectNodes(nodes: TreeNode[], arr: TreeNode[]): void {
    nodes.forEach(n => { arr.push(n); this.collectNodes(n.children, arr); });
  }

  private buildUL(nodes: TreeNode[], depth: number): HTMLUListElement {
    const ul = document.createElement("ul");
    ul.style.cssText = `
      list-style: none;
      margin: 0;
      padding-left: ${depth === 0 ? 0 : 18}px;
    `;

    nodes.forEach(node => {
      const li = document.createElement("li");
      li.dataset.path = node.fullPath;
      li.style.cssText = "margin: 2px 0;";

      const row = document.createElement("div");
      row.style.cssText = `
        display: flex; align-items: center;
        padding: 5px 8px; border-radius: 4px;
        cursor: pointer; font-size: 13px;
        color: #323130;
        transition: background 0.15s;
      `;
      row.addEventListener("mouseenter", () => {
        row.style.background = "#edebe9";
      });
      row.addEventListener("mouseleave", () => {
        row.style.background = "transparent";
      });

      // Toggle arrow
      const arrow = document.createElement("span");
      if (node.children.length > 0) {
        arrow.innerText = node.expanded ? "▾" : "▸";
        arrow.style.cssText = `
          margin-right: 6px; font-size: 11px;
          color: #0078d4; min-width: 12px;
          transition: transform 0.15s;
        `;
        row.addEventListener("click", () => {
          node.expanded = !node.expanded;
          arrow.innerText = node.expanded ? "▾" : "▸";
          if (childUL) {
            childUL.style.display = node.expanded ? "block" : "none";
          }
        });
      } else {
        arrow.style.cssText = "margin-right: 6px; min-width: 12px;";
        arrow.innerText = "•";
      }
      row.appendChild(arrow);

      // Label
      const label = document.createElement("span");
      label.innerText = node.label;
      row.appendChild(label);

      li.appendChild(row);

      // Children
      let childUL: HTMLUListElement | null = null;
      if (node.children.length > 0) {
        childUL = this.buildUL(node.children, depth + 1);
        childUL.style.display = node.expanded ? "block" : "none";
        li.appendChild(childUL);
      }

      ul.appendChild(li);
    });

    return ul;
  }

  // ----------------------------------------------------------------
  // Search / Filter
  // ----------------------------------------------------------------
  private filterTree(): void {
    const query = this.searchInput.value.trim().toLowerCase();
    if (!query) {
      // Reset all
      this.allNodes.forEach(n => { n.visible = true; n.expanded = false; });
      if (this.allNodes.length > 0) {
        this.renderTree(this.getRoots());
      }
      return;
    }

    // Mark matching nodes and their ancestors visible
    this.allNodes.forEach(n => { n.visible = false; });

    this.allNodes.forEach(n => {
      if (n.label.toLowerCase().includes(query)) {
        n.visible = true;
        // Mark all ancestors visible too
        let path = n.parentPath;
        while (path) {
          const ancestor = this.allNodes.find(a => a.fullPath === path);
          if (ancestor) { ancestor.visible = true; ancestor.expanded = true; }
          path = ancestor?.parentPath ?? null;
        }
      }
    });

    this.treeDiv.innerHTML = "";
    const filtered = this.buildFilteredUL(this.getRoots(), 0);
    this.treeDiv.appendChild(filtered);
  }

  private buildFilteredUL(nodes: TreeNode[], depth: number): HTMLUListElement {
    const ul = document.createElement("ul");
    ul.style.cssText = `list-style:none; margin:0; padding-left:${depth===0?0:18}px;`;

    nodes.filter(n => n.visible).forEach(node => {
      const li = document.createElement("li");
      const row = document.createElement("div");
      row.style.cssText = `
        display:flex; align-items:center; padding:5px 8px;
        border-radius:4px; font-size:13px; color:#323130;
      `;

      const arrow = document.createElement("span");
      arrow.style.cssText = "margin-right:6px; min-width:12px; color:#0078d4; font-size:11px;";
      arrow.innerText = node.children.length > 0 ? "▾" : "•";
      row.appendChild(arrow);

      const label = document.createElement("span");
      const query = this.searchInput.value.trim().toLowerCase();
      const idx = node.label.toLowerCase().indexOf(query);
      if (idx >= 0) {
        label.innerHTML =
          node.label.substring(0, idx) +
          `<mark style="background:#fff100; padding:0;">${node.label.substring(idx, idx + query.length)}</mark>` +
          node.label.substring(idx + query.length);
      } else {
        label.innerText = node.label;
      }
      row.appendChild(label);
      li.appendChild(row);

      if (node.children.length > 0) {
        const childUL = this.buildFilteredUL(node.children, depth + 1);
        li.appendChild(childUL);
      }

      ul.appendChild(li);
    });

    return ul;
  }

  private getRoots(): TreeNode[] {
    return this.allNodes.filter(n => n.parentPath === null);
  }

  public getOutputs(): IOutputs { return {}; }
  public destroy(): void {
    this.treeDiv.innerHTML = "";
    this.allNodes = [];
  }
}

// ---- Type ----
interface TreeNode {
  label: string;
  fullPath: string;
  parentPath: string | null;
  children: TreeNode[];
  expanded: boolean;
  visible: boolean;
}