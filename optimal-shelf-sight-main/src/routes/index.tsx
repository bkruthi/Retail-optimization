import { createFileRoute } from "@tanstack/react-router";
import * as React from "react";
import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import Solver from "javascript-lp-solver";
import {
  BarChart3,
  Box,
  Check,
  ChevronRight,
  CircleAlert,
  Database,
  FileSpreadsheet,
  Gauge,
  Info,
  LayoutDashboard,
  LineChart,
  PackageSearch,
  PanelLeft,
  Play,
  Settings2,
  SlidersHorizontal,
  Upload,
} from "lucide-react";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "RetailOptima — Retail Decision Support" },
      { name: "description", content: "A focused interface for retail inventory and product optimization." },
      { property: "og:title", content: "RetailOptima — Retail Decision Support" },
      { property: "og:description", content: "A focused interface for retail inventory and product optimization." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: RetailOptimizer,
});

type RawRow = Record<string, unknown>;
type Product = {
  StockCode: string;
  Description: string;
  unit_price: number;
  cost: number;
  margin: number;
  demand: number;
  storage_unit: number;
  storage_tier: string;
  revenue: number;
  orders: number;
};
type CleanRow = {
  InvoiceNo: string;
  StockCode: string;
  Description: string;
  Quantity: number;
  InvoiceDate: Date;
  UnitPrice: number;
  CustomerID: string;
  Country: string;
  Revenue: number;
};
type ModelResult = {
  status: string;
  rows: Product[];
  total_profit: number;
  budget_used: number;
  storage_used: number;
  budget_limit: number;
  storage_limit: number;
  count_selected?: number;
  max_products?: number;
};
type TierBase = { tiny: number; small: number; medium: number; large: number };
type SizeModifiers = { MINI: number; TINY: number; SMALL: number; TRAVEL: number; LARGE: number; JUMBO: number; GIANT: number };

const REQUIRED_COLUMNS = [
  "InvoiceNo", "StockCode", "Description", "Quantity", "InvoiceDate",
  "UnitPrice", "CustomerID", "Country",
];
const NON_PRODUCT_CODES = new Set([
  "POST", "DOT", "M", "D", "S", "C2", "PADS", "B", "AMAZONFEE", "CRUK", "BANK CHARGES",
]);
const TIER_BASE_DEFAULTS = { tiny: 0.3, small: 0.6, medium: 1, large: 1.8 };
const SIZE_DEFAULTS = { MINI: 0.5, TINY: 0.5, SMALL: 0.6, TRAVEL: 0.6, LARGE: 1.6, JUMBO: 2.2, GIANT: 2.2 };
const TIER_KEYWORDS: Record<string, string[]> = {
  tiny: ["NECKLACE", "BRACELET", "EARRINGS", "RING", "CARD", "STICKER", "MAGNET", "CHARM", "KEYRING", "RIBBON", "PEN", "TISSUE", "DOILY", "DOILIES"],
  small: ["MUG", "CUP", "BOWL", "CANDLE", "JAR", "TIN", "BOTTLE", "COASTER", "EGG", "KNOB", "DISH", "HOOK", "TOY", "REEL", "PENCIL"],
  medium: ["BOX", "BAG", "HOLDER", "FRAME", "CLOCK", "SIGN", "GARLAND", "BUNTING", "NOTEBOOK", "WRAP", "DECORATION", "ORNAMENT", "TOWEL", "APRON", "PLATE", "TRAY", "NAPKIN", "KIT", "GLASS", "TILE", "TUBE", "BIN", "PARASOL", "POT", "HEART", "TREE", "LIGHT", "STICKERS"],
  large: ["CUSHION", "LAMP", "LANTERN", "VASE", "BASKET", "BUCKET", "RUG", "THROW", "MIRROR", "CABINET", "DRAWER", "SHELF", "STOOL", "CHAIR", "STORAGE", "PLANTER", "CUTLERY", "UMBRELLA", "DOORMAT", "PILLOW", "GARDEN"],
};

const navItems = [
  { id: "Dashboard", icon: LayoutDashboard },
  { id: "Inventory", icon: Box },
  { id: "Product Mix", icon: SlidersHorizontal },
  { id: "Promotion", icon: PackageSearch },
  { id: "Analysis", icon: BarChart3 },
] as const;

function numberValue(value: unknown) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function parseDate(value: unknown) {
  const date = new Date(String(value));
  return Number.isNaN(date.getTime()) ? null : date;
}

function cleanTransactions(rows: RawRow[]) {
  const seen = new Set<string>();
  const output: CleanRow[] = [];
  for (const row of rows) {
    const signature = JSON.stringify(row);
    if (seen.has(signature)) continue;
    seen.add(signature);
    const stockCode = String(row["StockCode"] ?? "").trim().toUpperCase();
    const invoiceNo = String(row["InvoiceNo"] ?? "");
    const description = String(row["Description"] ?? "").trim();
    const quantity = numberValue(row["Quantity"]);
    const unitPrice = numberValue(row["UnitPrice"]);
    const date = parseDate(row["InvoiceDate"]);
    if (!stockCode || NON_PRODUCT_CODES.has(stockCode)) continue;
    if (invoiceNo.startsWith("C") || invoiceNo.startsWith("A")) continue;
    if (!description || quantity <= 0 || unitPrice <= 0 || !date) continue;
    output.push({
      InvoiceNo: invoiceNo,
      StockCode: stockCode,
      Description: description,
      Quantity: quantity,
      InvoiceDate: date,
      UnitPrice: unitPrice,
      CustomerID: String(row["CustomerID"] ?? ""),
      Country: String(row["Country"] ?? ""),
      Revenue: quantity * unitPrice,
    });
  }
  return output;
}

function storageSizing(description: string, tierBase: Record<string, number>, modifiers: Record<string, number>, packCap: number) {
  const tokens = new Set((description.toUpperCase().match(/[A-Z]+/g) ?? []));
  const matchesKeyword = (keyword: string) => {
    const key = keyword.trim().toUpperCase();
    if (!key) return false;
    return tokens.has(key) || tokens.has(`${key}S`) || tokens.has(`${key}ES`) ||
      (key.endsWith("Y") && tokens.has(`${key.slice(0, -1)}IES`));
  };

  let tier = "medium";
  for (const candidate of ["tiny", "small", "medium", "large"]) {
    if (TIER_KEYWORDS[candidate]?.some(matchesKeyword)) {
      tier = candidate;
      break;
    }
  }

  let sizeModifier = 1;
  for (const [keyword, multiplier] of Object.entries(modifiers)) {
    if (matchesKeyword(keyword)) {
      sizeModifier = multiplier;
      break;
    }
  }

  const packMatch = description.toUpperCase().match(/(?:PACK OF|BOX OF|SET OF|SET\/|SET|PACK)\s*(\d{1,3})/i);
  const count = packMatch?.[1] ? Number(packMatch[1]) : 1;
  const packMultiplier = count > 1 ? Math.min(1 + 0.5 * Math.log10(count), packCap) : 1;
  return { tier, value: (tierBase[tier] ?? 1) * sizeModifier * packMultiplier };
}

function aggregateProducts(data: CleanRow[], costRatio: number, topN: number, tierBase: Record<string, number>, modifiers: Record<string, number>, packCap: number) {
  const groups = new Map<string, { descriptions: Map<string, number>; unitPrice: number; demand: number; revenue: number; invoices: Set<string> }>();
  for (const row of data) {
    const current = groups.get(row.StockCode) ?? { descriptions: new Map(), unitPrice: 0, demand: 0, revenue: 0, invoices: new Set() };
    current.descriptions.set(row.Description, (current.descriptions.get(row.Description) ?? 0) + 1);
    current.unitPrice += row.UnitPrice;
    current.demand += row.Quantity;
    current.revenue += row.Revenue;
    current.invoices.add(row.InvoiceNo);
    groups.set(row.StockCode, current);
  }
  return [...groups.entries()]
    .map(([stockCode, group]) => {
      const description = [...group.descriptions.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? stockCode;
      const unitPrice = group.unitPrice / Math.max(1, [...group.descriptions.values()].reduce((sum: number, count: number) => sum + count, 0));
      const sizing = storageSizing(description, tierBase, modifiers, packCap);
      const cost = unitPrice * costRatio;
      return { StockCode: stockCode, Description: description, unit_price: unitPrice, cost, margin: unitPrice - cost, demand: Math.round(group.demand), storage_unit: sizing.value, storage_tier: sizing.tier, revenue: group.revenue, orders: group.invoices.size };
    })
    .sort((a, b) => b.revenue - a.revenue)
    .slice(0, topN);
}

function solveProductMix(products: Product[], budget: number, storage: number): ModelResult {
  const model = { optimize: "profit", opType: "max", constraints: { budget: { max: budget }, storage: { max: storage } } as Record<string, Record<string, number>>, variables: {} as Record<string, Record<string, number>> };
  const variables = model.variables;
  products.forEach((product, index) => { variables[`x${index}`] = { profit: product.margin, budget: product.cost, storage: product.storage_unit, [`demand${index}`]: 1 }; model.constraints[`demand${index}`] = { max: product.demand }; });
  const solution = Solver.Solve(model);
  const rows = products.map((product, index) => ({ ...product, recommended: numberValue(solution[`x${index}`]) })).filter((row) => row.recommended > 0).map((row) => ({ ...row, expected: row.recommended * row.margin, spend: row.recommended * row.cost, usedStorage: row.recommended * row.storage_unit })).sort((a, b) => b.expected - a.expected);
  return { status: String(solution["feasible"] === false ? "Infeasible" : "Optimal"), rows: rows.map(({ recommended: RecommendedQty, expected: ExpectedProfit, spend: SpendUsed, usedStorage: StorageUsed, ...product }) => ({ ...product, RecommendedQty, ExpectedProfit, SpendUsed, StorageUsed } as Product)), total_profit: rows.reduce((sum, row) => sum + row.expected, 0), budget_used: rows.reduce((sum, row) => sum + row.spend, 0), storage_used: rows.reduce((sum, row) => sum + row.usedStorage, 0), budget_limit: budget, storage_limit: storage };
}

function solvePromotion(products: Product[], budget: number, storage: number, maxProducts: number): ModelResult {
  const model = { optimize: "profit", opType: "max", constraints: { budget: { max: budget }, storage: { max: storage }, count: { max: maxProducts } }, variables: {} as Record<string, Record<string, number>>, ints: {} as Record<string, number> };
  const variables = model.variables;
  const ints = model.ints;
  products.forEach((product, index) => { const key = `y${index}`; variables[key] = { profit: product.margin, budget: product.cost, storage: product.storage_unit, count: 1 }; ints[key] = 1; });
  const solution = Solver.Solve(model);
  const rows = products.filter((_, index) => Boolean(numberValue(solution[`y${index}`]))).sort((a, b) => b.margin - a.margin);
  return { status: String(solution["feasible"] === false ? "Infeasible" : "Optimal"), rows, total_profit: rows.reduce((sum, row) => sum + row.margin, 0), budget_used: rows.reduce((sum, row) => sum + row.cost, 0), storage_used: rows.reduce((sum, row) => sum + row.storage_unit, 0), budget_limit: budget, storage_limit: storage, count_selected: rows.length, max_products: maxProducts };
}

function money(value: number) { return `£${value.toLocaleString("en-GB", { maximumFractionDigits: 0 })}`; }
function decimal(value: number) { return value.toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

function RetailOptimizer() {
  const [page, setPage] = useState<(typeof navItems)[number]["id"]>("Dashboard");
  const [data, setData] = useState<CleanRow[] | null>(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [costRatio, setCostRatio] = useState(0.6);
  const [topN, setTopN] = useState(20);
  const [tierBase, setTierBase] = useState<TierBase>(TIER_BASE_DEFAULTS);
  const [modifiers, setModifiers] = useState<SizeModifiers>(SIZE_DEFAULTS);
  const [packCap, setPackCap] = useState(2);
  const [mixBudget, setMixBudget] = useState(5000);
  const [mixStorage, setMixStorage] = useState(500);
  const [promoBudget, setPromoBudget] = useState(5000);
  const [promoStorage, setPromoStorage] = useState(500);
  const [maxProducts, setMaxProducts] = useState(8);
  const [mixResult, setMixResult] = useState<ModelResult | null>(null);
  const [promoResult, setPromoResult] = useState<ModelResult | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const products = useMemo(() => data ? aggregateProducts(data, costRatio, topN, tierBase, modifiers, packCap) : [], [data, costRatio, topN, tierBase, modifiers, packCap]);
  const stats = useMemo(() => data ? { revenue: data.reduce((sum, row) => sum + row.Revenue, 0), quantity: data.reduce((sum, row) => sum + row.Quantity, 0), orders: new Set(data.map((row) => row.InvoiceNo)).size, products: new Set(data.map((row) => row.StockCode)).size, customers: new Set(data.map((row) => row.CustomerID).filter(Boolean)).size, countries: new Set(data.map((row) => row.Country).filter(Boolean)).size } : null, [data]);

  async function handleBuffer(buffer: ArrayBuffer, sourceName: string) {
    setError("");
    try {
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0] ?? ""];
      if (!sheet) throw new Error("The selected file has no readable worksheet.");
      const rows = XLSX.utils.sheet_to_json<RawRow>(sheet, { defval: "" });
      const missing = REQUIRED_COLUMNS.filter((column) => !rows.some((row) => Object.prototype.hasOwnProperty.call(row, column)));
      if (missing.length) throw new Error(`Missing required columns: ${missing.join(", ")}`);
      const cleaned = cleanTransactions(rows);
      if (!cleaned.length) throw new Error("No valid transaction rows remain after preprocessing.");
      setData(cleaned); setFileName(sourceName); setPage("Dashboard"); setMixResult(null); setPromoResult(null);
    } catch (caught) { setError(caught instanceof Error ? caught.message : "Unable to read this file."); }
  }

  async function handleFile(file: File) {
    await handleBuffer(await file.arrayBuffer(), file.name);
  }

  async function handleSample() {
    try {
      const response = await fetch("/sample_online_retail.csv");
      if (!response.ok) throw new Error("Bundled sample dataset could not be loaded.");
      await handleBuffer(await response.arrayBuffer(), "sample_online_retail.csv");
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Unable to load the bundled sample dataset.");
    }
  }

  const activeIndex = navItems.findIndex((item) => item.id === page);
  return (
    <div className="min-h-screen bg-zinc-50 font-sans text-zinc-900">
      <div className="flex min-h-screen">
        <aside className="hidden w-64 shrink-0 flex-col border-r border-zinc-200 bg-sidebar lg:flex">
          <div className="p-6"><div className="flex items-center gap-3"><div className="grid size-8 place-items-center rounded-md bg-zinc-900 text-xs font-semibold text-primary-foreground">RO</div><div><div className="text-sm font-semibold tracking-tight">RetailOptima</div><div className="font-mono text-[10px] uppercase tracking-wider text-zinc-500">Operations Research</div></div></div></div>
          <nav className="flex-1 space-y-0.5 px-3">
            {navItems.map(({ id, icon: Icon }) => <Button key={id} variant="ghost" onClick={() => setPage(id)} className={`h-10 w-full justify-start gap-3 rounded-md px-3 text-sm font-medium ${page === id ? "bg-zinc-900 text-primary-foreground hover:bg-zinc-800 hover:text-primary-foreground" : "text-zinc-500 hover:bg-zinc-200/60 hover:text-zinc-900"}`}><Icon className="size-4" />{id}</Button>)}
          </nav>
          <div className="m-6 rounded-lg bg-accent-amber p-4 ring-1 ring-zinc-300/40"><div className="mb-2 font-mono text-[10px] uppercase tracking-wider text-zinc-500">System status</div><div className="flex items-center gap-2 text-xs font-medium"><span className={`size-2 rounded-full ${data ? "bg-chart-2" : "bg-chart-4"}`} />{data ? "Dataset loaded" : "Awaiting dataset"}</div></div>
        </aside>
        <main className="min-w-0 flex-1">
          <header className="flex min-h-14 items-center justify-between border-b border-zinc-200 bg-card px-5 py-3 lg:px-8"><div className="flex items-center gap-3"><PanelLeft className="size-4 text-zinc-400 lg:hidden" /><span className="text-xs font-medium text-zinc-500">Main / <span className="text-zinc-900">{page}</span></span></div><div className="flex items-center gap-3"><span className="hidden font-mono text-[10px] uppercase tracking-wider text-zinc-400 sm:inline">Retail decision support</span><Button variant="outline" size="sm" onClick={() => inputRef.current?.click()}><Upload />Upload data</Button></div></header>
          <input ref={inputRef} className="hidden" type="file" accept=".csv,.xlsx,.xls" onChange={(event) => { const file = event.target.files?.[0]; if (file) void handleFile(file); }} />
          <div className="mx-auto max-w-7xl space-y-8 p-5 lg:p-8">
            <div className="lg:hidden"><div className="mb-3 flex gap-1 overflow-x-auto pb-1">{navItems.map(({ id, icon: Icon }) => <Button key={id} variant="ghost" size="sm" onClick={() => setPage(id)} className={page === id ? "bg-zinc-900 text-primary-foreground hover:bg-zinc-800 hover:text-primary-foreground" : "text-zinc-500"}><Icon />{id}</Button>)}</div></div>
            {error && <div className="flex items-start gap-3 rounded-md border border-destructive/20 bg-destructive/5 p-4 text-sm text-destructive"><CircleAlert className="mt-0.5 size-4 shrink-0" /><span>{error}</span></div>}
            {page === "Dashboard" && <DashboardView data={data} fileName={fileName} stats={stats} onUpload={() => inputRef.current?.click()} onSample={handleSample} />}
            {page === "Inventory" && <InventoryView data={data} products={products} fileName={fileName} onUpload={() => inputRef.current?.click()} tierBase={tierBase} setTierBase={setTierBase} modifiers={modifiers} setModifiers={setModifiers} packCap={packCap} setPackCap={setPackCap} costRatio={costRatio} setCostRatio={setCostRatio} topN={topN} setTopN={setTopN} />}
            {page === "Product Mix" && <OptimizationView kind="mix" products={products} budget={mixBudget} storage={mixStorage} maxProducts={maxProducts} setBudget={setMixBudget} setStorage={setMixStorage} setMaxProducts={setMaxProducts} result={mixResult} onRun={() => setMixResult(solveProductMix(products, mixBudget, mixStorage))} />}
            {page === "Promotion" && <OptimizationView kind="promo" products={products} budget={promoBudget} storage={promoStorage} maxProducts={maxProducts} setBudget={setPromoBudget} setStorage={setPromoStorage} setMaxProducts={setMaxProducts} result={promoResult} onRun={() => setPromoResult(solvePromotion(products, promoBudget, promoStorage, maxProducts))} />}
            {page === "Analysis" && <AnalysisView data={data} stats={stats} products={products} />}
          </div>
        </main>
      </div>
    </div>
  );
}

function DashboardView({ data, fileName, stats, onUpload, onSample }: { data: CleanRow[] | null; fileName: string; stats: { revenue: number; quantity: number; orders: number; products: number; customers: number; countries: number } | null; onUpload: () => void; onSample: () => void }) {
  return <>
    <section className="grid gap-8 lg:grid-cols-[minmax(0,2fr)_minmax(280px,1fr)]"><div className="space-y-4"><div className="font-mono text-[11px] uppercase tracking-[0.18em] text-zinc-500">System dashboard</div><h1 className="max-w-3xl text-3xl font-medium leading-tight tracking-tight lg:text-4xl">Intelligent resource allocation for retail environments.</h1><p className="max-w-2xl text-base leading-relaxed text-zinc-600">Use real transaction data to support product mix and promotional selection decisions under purchasing, storage, and product-count constraints.</p></div><div className="flex min-h-44 flex-col items-center justify-center rounded-lg bg-accent-blue p-6 text-center ring-1 ring-black/5"><Database className="mb-4 size-6 text-zinc-500" />{data ? <><p className="text-sm font-medium text-zinc-900">Dataset ready</p><p className="mt-1 max-w-xs truncate font-mono text-xs text-zinc-500">{fileName}</p></> : <><p className="text-sm font-medium text-zinc-500">No active dataset detected.</p><div className="mt-4 flex flex-wrap justify-center gap-2"><Button variant="outline" className="bg-card" onClick={onUpload}><Upload />Upload data</Button><Button variant="ghost" onClick={onSample}>Use sample dataset</Button></div></>}</div></section>
    {stats && <div className="grid grid-cols-2 gap-px overflow-hidden rounded-lg bg-zinc-200 ring-1 ring-zinc-200 sm:grid-cols-5"><Stat label="Revenue" value={money(stats.revenue)} /><Stat label="Transactions" value={stats.orders.toLocaleString()} /><Stat label="Products" value={stats.products.toLocaleString()} /><Stat label="Customers" value={stats.customers.toLocaleString()} /><Stat label="Countries" value={stats.countries.toLocaleString()} /></div>}
    <div className="grid gap-px overflow-hidden rounded-xl bg-zinc-200 ring-1 ring-zinc-200 md:grid-cols-2"><ModelCard label="Model A" title="Product Mix Optimization" type="Linear Programming" description="Determine recommended stock quantities per product subject to purchasing budget and warehouse capacity limits." lines={["Maximize Σ marginᵢ × xᵢ", "Σ costᵢ × xᵢ ≤ Budget"]} tone="blue" /><ModelCard label="Model B" title="Promotional Product Selection" type="0–1 Integer Programming / Knapsack" description="Select products for promotion under promotional budget, storage/display capacity, and maximum product count." lines={["Maximize Σ marginᵢ × yᵢ", "yᵢ ∈ {0,1}"]} tone="green" /></div>
    <section className="space-y-4"><SectionHeading icon={LineChart} title="Data readiness" meta={data ? "Cleaned transaction rows available" : "Awaiting data aggregation"} /><div className="rounded-xl border border-zinc-200 bg-card p-12 text-center"><div className="mx-auto max-w-sm"><div className="mx-auto mb-5 grid size-12 place-items-center rounded-full bg-zinc-100 text-zinc-400"><Gauge className="size-5" /></div><p className="text-sm italic text-zinc-400">{data ? "Open Inventory or Analysis to inspect the uploaded data." : "Upload transaction data to begin."}</p></div></div></section>
    <StorageOverview />
  </>;
}

function InventoryView({ data, products, fileName, onUpload, tierBase, setTierBase, modifiers, setModifiers, packCap, setPackCap, costRatio, setCostRatio, topN, setTopN }: { data: CleanRow[] | null; products: Product[]; fileName: string; onUpload: () => void; tierBase: TierBase; setTierBase: React.Dispatch<React.SetStateAction<TierBase>>; modifiers: SizeModifiers; setModifiers: React.Dispatch<React.SetStateAction<SizeModifiers>>; packCap: number; setPackCap: (value: number) => void; costRatio: number; setCostRatio: (value: number) => void; topN: number; setTopN: (value: number) => void }) {
  if (!data) return <EmptyPage title="Inventory" description="Upload the retail transaction CSV or XLSX file to build the product-level catalog." onUpload={onUpload} />;
  return <><PageIntro eyebrow="Data source" title="Inventory" description={`${fileName} · ${data.length.toLocaleString()} cleaned transaction rows`} /><div className="grid gap-8 xl:grid-cols-[minmax(0,1fr)_300px]"><section className="min-w-0 space-y-4"><SectionHeading icon={Box} title="Product catalog" meta={`Top ${products.length} by historical revenue`} /><div className="overflow-x-auto rounded-lg border border-zinc-200 bg-card"><table className="w-full min-w-[980px] text-left text-xs"><thead className="border-b border-zinc-200 bg-zinc-100 font-mono uppercase tracking-wider text-zinc-500"><tr>{["StockCode", "Description", "Unit Price", "Cost", "Margin", "Demand", "Storage / Unit", "Size Tier", "Revenue", "Orders"].map((heading) => <th key={heading} className="px-3 py-3 font-medium">{heading}</th>)}</tr></thead><tbody className="divide-y divide-zinc-100">{products.map((product) => <tr key={product.StockCode} className="hover:bg-accent-blue/40"><td className="whitespace-nowrap px-3 py-3 font-mono text-zinc-500">{product.StockCode}</td><td className="max-w-72 px-3 py-3 font-medium">{product.Description}</td><td className="whitespace-nowrap px-3 py-3">{money(product.unit_price)}</td><td className="whitespace-nowrap px-3 py-3">{money(product.cost)}</td><td className="whitespace-nowrap px-3 py-3">{money(product.margin)}</td><td className="whitespace-nowrap px-3 py-3">{product.demand.toLocaleString()}</td><td className="whitespace-nowrap px-3 py-3">{decimal(product.storage_unit)}</td><td className="px-3 py-3 capitalize">{product.storage_tier}</td><td className="whitespace-nowrap px-3 py-3">{money(product.revenue)}</td><td className="px-3 py-3">{product.orders.toLocaleString()}</td></tr>)}</tbody></table></div></section><StorageConfig tierBase={tierBase} setTierBase={setTierBase} modifiers={modifiers} setModifiers={setModifiers} packCap={packCap} setPackCap={setPackCap} costRatio={costRatio} setCostRatio={setCostRatio} topN={topN} setTopN={setTopN} /></div></>;
}

function OptimizationView({ kind, products, budget, storage, maxProducts, setBudget, setStorage, setMaxProducts, result, onRun }: { kind: "mix" | "promo"; products: Product[]; budget: number; storage: number; maxProducts: number; setBudget: (value: number) => void; setStorage: (value: number) => void; setMaxProducts: (value: number) => void; result: ModelResult | null; onRun: () => void }) {
  const isMix = kind === "mix";
  if (!products.length) return <EmptyPage title={isMix ? "Product Mix Optimization" : "Promotional Product Selection"} description="Upload transaction data and build the inventory product list before running the model." />;
  return <><PageIntro eyebrow={isMix ? "Model A · Linear Programming" : "Model B · 0–1 Integer Programming"} title={isMix ? "Product Mix Optimization" : "Promotional Product Selection"} description={isMix ? "Maximize total margin while respecting purchasing budget, storage capacity, and historical demand limits." : "Maximize selected-product margin while respecting promotional budget, storage/display capacity, and the maximum product count."} /><section className="rounded-xl border border-zinc-200 bg-card p-6"><div className="grid gap-5 md:grid-cols-3"><Field label={isMix ? "Purchasing Budget" : "Promotional Budget"} value={budget} onChange={setBudget} prefix="£" /><Field label="Storage Capacity" value={storage} onChange={setStorage} suffix="units" />{isMix ? <div className="space-y-2"><label className="font-mono text-[10px] uppercase tracking-wider text-zinc-500">Model parameters</label><div className="flex h-9 items-center rounded-md border border-zinc-200 bg-zinc-50 px-3 text-sm text-zinc-600">Cost ratio is configured in Inventory</div></div> : <Field label="Maximum Number of Products" value={maxProducts} onChange={setMaxProducts} />}</div><div className="mt-6 flex flex-wrap items-center justify-between gap-4 border-t border-zinc-100 pt-5"><div className="flex items-center gap-2 text-xs text-zinc-500"><Info className="size-4" />{isMix ? "xᵢ ≥ 0 and xᵢ ≤ historical demandᵢ" : "yᵢ ∈ {0,1} and Σ yᵢ ≤ maximum product count"}</div><Button onClick={onRun}><Play />Run {isMix ? "Product Mix" : "Promotion"} Optimization</Button></div></section>{result ? <ResultView kind={kind} result={result} products={products} /> : <div className="rounded-xl border border-dashed border-zinc-300 bg-card p-12 text-center text-sm italic text-zinc-400">Run the model to calculate an actual recommendation from the uploaded product data.</div>}</>;
}

function ResultView({ kind, result, products }: { kind: "mix" | "promo"; result: ModelResult; products: Product[] }) { const isMix = kind === "mix"; const rows = result.rows as (Product & { RecommendedQty?: number; ExpectedProfit?: number; SpendUsed?: number; StorageUsed?: number })[]; return <section className="space-y-4"><SectionHeading icon={Check} title="Optimization result" meta={`Solver status: ${result.status}`} /><div className="grid grid-cols-2 gap-px overflow-hidden rounded-lg bg-zinc-200 ring-1 ring-zinc-200 sm:grid-cols-4"><Stat label={isMix ? "Expected Profit" : "Total Profit"} value={money(result.total_profit)} tone="green" /><Stat label="Budget Used" value={`${money(result.budget_used)} / ${money(result.budget_limit)}`} /><Stat label="Storage Used" value={`${decimal(result.storage_used)} / ${decimal(result.storage_limit)}`} tone="blue" /><Stat label={isMix ? "Products with Qty" : "Products Selected"} value={isMix ? rows.length.toLocaleString() : `${result.count_selected ?? 0} / ${result.max_products ?? 0}`} /></div><div className="overflow-x-auto rounded-lg border border-zinc-200 bg-card"><table className="w-full min-w-[720px] text-left text-xs"><thead className="border-b border-zinc-200 bg-zinc-100 font-mono uppercase tracking-wider text-zinc-500"><tr>{["StockCode", "Description", "Unit Price", "Cost", "Margin", ...(isMix ? ["Recommended Qty", "Expected Profit", "Spend Used", "Storage Used"] : ["Storage / Unit", "Demand", "Orders"])].map((heading) => <th key={heading} className="px-3 py-3 font-medium">{heading}</th>)}</tr></thead><tbody className="divide-y divide-zinc-100">{rows.map((product) => <tr key={product.StockCode}><td className="px-3 py-3 font-mono text-zinc-500">{product.StockCode}</td><td className="max-w-72 px-3 py-3 font-medium">{product.Description}</td><td className="px-3 py-3">{money(product.unit_price)}</td><td className="px-3 py-3">{money(product.cost)}</td><td className="px-3 py-3">{money(product.margin)}</td>{isMix ? <><td className="px-3 py-3 font-mono">{decimal(product.RecommendedQty ?? 0)}</td><td className="px-3 py-3">{money(product.ExpectedProfit ?? 0)}</td><td className="px-3 py-3">{money(product.SpendUsed ?? 0)}</td><td className="px-3 py-3">{decimal(product.StorageUsed ?? 0)}</td></> : <><td className="px-3 py-3">{decimal(product.storage_unit)}</td><td className="px-3 py-3">{product.demand.toLocaleString()}</td><td className="px-3 py-3">{product.orders.toLocaleString()}</td></>}</tr>)}</tbody></table></div>{isMix && <Sensitivity result={result} products={products} />}</section>; }

function Sensitivity({ result, products }: { result: ModelResult; products: Product[] }) { const points = [0.5, 0.75, 1, 1.25, 1.5].map((factor) => { const output = solveProductMix(products, result.budget_limit * factor, result.storage_limit); return { budget: result.budget_limit * factor, profit: output.total_profit }; }); const max = Math.max(...points.map((point) => point.profit), 1); return <div className="rounded-xl border border-zinc-200 bg-card p-6"><div className="mb-5 flex items-center justify-between"><div><h3 className="text-sm font-semibold">Budget Sensitivity</h3><p className="mt-1 text-xs text-zinc-500">Product Mix re-solved from approximately 0.5× to 1.5× budget.</p></div><LineChart className="size-4 text-zinc-400" /></div><div className="flex h-36 items-end gap-3 border-b border-l border-zinc-200 px-4 pb-0 pt-4">{points.map((point) => <div key={point.budget} className="flex flex-1 flex-col items-center justify-end gap-2"><div className="w-full max-w-16 rounded-t-sm bg-accent-blue" style={{ height: `${Math.max(8, (point.profit / max) * 100)}%` }} title={`${money(point.profit)} maximum profit`} /><span className="font-mono text-[10px] text-zinc-500">{Math.round(point.budget / result.budget_limit * 100)}%</span></div>)}</div><div className="mt-3 flex items-center justify-between font-mono text-[10px] uppercase tracking-wider text-zinc-400"><span>Budget</span><span>Maximum Profit</span></div></div>; }

function AnalysisView({ data, stats, products }: { data: CleanRow[] | null; stats: { revenue: number; quantity: number; orders: number; products: number; customers: number; countries: number } | null; products: Product[] }) { if (!data || !stats) return <EmptyPage title="Analysis" description="Upload transaction data to generate the descriptive analysis supported by the project." />; const byCountry = [...new Map(data.map((row) => [row.Country, data.filter((item) => item.Country === row.Country).reduce((sum, item) => sum + item.Revenue, 0)])).entries()].sort((a, b) => b[1] - a[1]).slice(0, 8); const byMonth = [...new Map(data.map((row) => { const month = `${row.InvoiceDate.getFullYear()}-${String(row.InvoiceDate.getMonth() + 1).padStart(2, "0")}`; return [month, data.filter((item) => item.InvoiceDate.getFullYear() === row.InvoiceDate.getFullYear() && item.InvoiceDate.getMonth() === row.InvoiceDate.getMonth()).reduce((sum, item) => sum + item.Revenue, 0)]; })).entries()].sort((a, b) => a[0].localeCompare(b[0])); return <><PageIntro eyebrow="Descriptive statistics" title="Analysis" description="Simple transaction and product analysis generated from the cleaned uploaded data." /><div className="grid grid-cols-2 gap-px overflow-hidden rounded-lg bg-zinc-200 ring-1 ring-zinc-200 sm:grid-cols-4"><Stat label="Revenue" value={money(stats.revenue)} /><Stat label="Quantity Sold" value={stats.quantity.toLocaleString()} /><Stat label="Transactions" value={stats.orders.toLocaleString()} /><Stat label="Unique Products" value={stats.products.toLocaleString()} /></div><div className="grid gap-8 lg:grid-cols-2"><ChartPanel title="Monthly Sales Trend" icon={LineChart}>{byMonth.map(([month, value]) => <BarRow key={month} label={month} value={value} max={Math.max(...byMonth.map(([, amount]) => amount))} />)}</ChartPanel><ChartPanel title="Sales by Country" icon={BarChart3}>{byCountry.map(([country, value]) => <BarRow key={country} label={country} value={value} max={byCountry[0]?.[1] ?? 1} />)}</ChartPanel></div><ChartPanel title="Top Products by Revenue" icon={PackageSearch}>{products.slice(0, 8).map((product) => <BarRow key={product.StockCode} label={product.Description} value={product.revenue} max={products[0]?.revenue ?? 1} />)}</ChartPanel></>; }

function StorageConfig({ tierBase, setTierBase, modifiers, setModifiers, packCap, setPackCap, costRatio, setCostRatio, topN, setTopN }: { tierBase: TierBase; setTierBase: React.Dispatch<React.SetStateAction<TierBase>>; modifiers: SizeModifiers; setModifiers: React.Dispatch<React.SetStateAction<SizeModifiers>>; packCap: number; setPackCap: (value: number) => void; costRatio: number; setCostRatio: (value: number) => void; topN: number; setTopN: (value: number) => void }) { return <aside className="space-y-5"><div className="rounded-lg bg-accent-green p-5 ring-1 ring-black/5"><div className="mb-4 flex items-center gap-2"><Settings2 className="size-4 text-zinc-600" /><h2 className="text-sm font-semibold">Storage Configuration</h2></div><p className="mb-5 text-xs leading-relaxed text-zinc-600">Rule-based storage sizing reads category keywords, size language, and pack language from each description. Products without a matching category use the medium tier.</p><div className="space-y-3">{Object.entries(tierBase).map(([tier, value]) => <label key={tier} className="flex items-center justify-between gap-3 text-xs capitalize text-zinc-600">{tier} tier<Input className="h-8 w-24 bg-card text-right font-mono text-xs" type="number" min="0.1" step="0.1" value={value} onChange={(event) => setTierBase((current) => ({ ...current, [tier]: numberValue(event.target.value) }))} /></label>)}</div><div className="mt-5 border-t border-black/10 pt-5"><label className="mb-2 block font-mono text-[10px] uppercase tracking-wider text-zinc-500">Pack multiplier cap</label><Input className="bg-card font-mono text-xs" type="number" min="1" step="0.1" value={packCap} onChange={(event) => setPackCap(numberValue(event.target.value))} /></div></div><div className="rounded-lg border border-zinc-200 bg-card p-5"><div className="mb-4 flex items-center gap-2"><SlidersHorizontal className="size-4 text-zinc-500" /><h2 className="text-sm font-semibold">Model parameters</h2></div><Field label="Cost ratio" value={costRatio} onChange={setCostRatio} step={0.05} /><Field label="Top-N products" value={topN} onChange={setTopN} /></div><div className="rounded-lg border border-zinc-200 bg-card p-5"><div className="mb-3 flex items-center gap-2"><Info className="size-4 text-zinc-400" /><span className="text-xs font-medium">Size-word multipliers</span></div><div className="space-y-2">{Object.entries(modifiers).map(([key, value]) => <label key={key} className="flex items-center justify-between gap-2 font-mono text-[10px] uppercase tracking-wider text-zinc-500">{key}<Input className="h-7 w-20 bg-zinc-50 text-right font-mono text-[10px]" type="number" min="0.1" step="0.1" value={value} onChange={(event) => setModifiers((current) => ({ ...current, [key]: numberValue(event.target.value) }))} /></label>)}</div></div></aside>; }

function StorageOverview() { return <section className="border-t border-zinc-200 pt-8"><div className="grid gap-6 md:grid-cols-[190px_1fr]"><div><h2 className="text-xs font-semibold uppercase tracking-wider">Storage config</h2><p className="mt-2 text-xs leading-relaxed text-zinc-500">Editable rules for tier-based storage sizing and unit multipliers.</p></div><div className="grid grid-cols-2 gap-3 sm:grid-cols-4">{Object.entries(TIER_BASE_DEFAULTS).map(([tier, value]) => <div key={tier} className="rounded border border-zinc-200 bg-zinc-100 p-3"><span className="block font-mono text-[10px] uppercase text-zinc-400">{tier} tier</span><span className="mt-1 block font-mono text-sm">{value.toFixed(1)}</span></div>)}</div></div></section>; }
function PageIntro({ eyebrow, title, description }: { eyebrow: string; title: string; description: string }) { return <section><div className="mb-3 font-mono text-[11px] uppercase tracking-[0.18em] text-zinc-500">{eyebrow}</div><h1 className="text-3xl font-medium tracking-tight">{title}</h1><p className="mt-2 max-w-2xl text-sm leading-relaxed text-zinc-500">{description}</p></section>; }
function SectionHeading({ icon: Icon, title, meta }: { icon: typeof LineChart; title: string; meta?: string }) { return <div className="flex items-center justify-between gap-4"><div className="flex items-center gap-2"><Icon className="size-4 text-zinc-400" /><h2 className="text-sm font-semibold">{title}</h2></div>{meta && <span className="font-mono text-[10px] uppercase tracking-wider text-zinc-400">{meta}</span>}</div>; }
function EmptyPage({ title, description, onUpload }: { title: string; description: string; onUpload?: () => void }) { return <><PageIntro eyebrow="Data source required" title={title} description={description} /><div className="rounded-xl border border-dashed border-zinc-300 bg-card p-16 text-center"><FileSpreadsheet className="mx-auto mb-4 size-7 text-zinc-400" /><p className="text-sm font-medium">Upload transaction data to begin</p><p className="mx-auto mt-2 max-w-sm text-xs leading-relaxed text-zinc-500">The required schema is InvoiceNo, StockCode, Description, Quantity, InvoiceDate, UnitPrice, CustomerID, and Country.</p>{onUpload && <Button className="mt-6" onClick={onUpload}><Upload />Select CSV / XLSX file</Button>}</div></>; }
function Stat({ label, value, tone }: { label: string; value: string; tone?: "blue" | "green" }) { return <div className={`bg-card p-4 ${tone === "blue" ? "bg-accent-blue" : tone === "green" ? "bg-accent-green" : ""}`}><div className="font-mono text-[10px] uppercase tracking-wider text-zinc-500">{label}</div><div className="mt-1 text-lg font-medium tracking-tight">{value}</div></div>; }
function ModelCard({ label, title, type, description, lines, tone }: { label: string; title: string; type: string; description: string; lines: string[]; tone: "blue" | "green" }) { return <div className={`space-y-6 bg-card p-7 ${tone === "blue" ? "bg-accent-blue/40" : "bg-accent-green/40"}`}><div className="flex items-center justify-between gap-4"><span className="rounded bg-zinc-100 px-2 py-0.5 font-mono text-[10px] uppercase tracking-wider text-zinc-600">{label}</span><span className="text-right text-xs text-zinc-400">{type}</span></div><div><h2 className="text-lg font-medium">{title}</h2><p className="mt-2 max-w-lg text-sm leading-relaxed text-zinc-500">{description}</p></div><div className="space-y-3 border-t border-zinc-200/70 pt-4">{lines.map((line) => <div key={line} className="flex justify-between gap-4 text-xs"><span className="text-zinc-500">{line.startsWith("Maximize") ? "Objective" : line.startsWith("y") ? "Variable" : "Constraint"}</span><span className="font-mono text-right text-zinc-900">{line}</span></div>)}</div></div>; }
function Field({ label, value, onChange, prefix, suffix, step = 100 }: { label: string; value: number; onChange: (value: number) => void; prefix?: string; suffix?: string; step?: number }) { return <label className="block space-y-2"><span className="font-mono text-[10px] uppercase tracking-wider text-zinc-500">{label}</span><div className="flex h-9 items-center rounded-md border border-zinc-200 bg-card px-3 focus-within:ring-1 focus-within:ring-ring">{prefix && <span className="mr-2 text-xs text-zinc-400">{prefix}</span>}<input className="min-w-0 flex-1 bg-transparent text-sm outline-none" type="number" min="0" step={step} value={value} onChange={(event) => onChange(numberValue(event.target.value))} />{suffix && <span className="ml-2 text-xs text-zinc-400">{suffix}</span>}</div></label>; }
function ChartPanel({ title, icon, children }: { title: string; icon: typeof LineChart; children: React.ReactNode }) { return <section className="rounded-xl border border-zinc-200 bg-card p-6"><SectionHeading icon={icon} title={title} /><div className="mt-6 space-y-3">{children}</div></section>; }
function BarRow({ label, value, max }: { label: string; value: number; max: number }) { return <div className="grid grid-cols-[minmax(100px,1fr)_minmax(100px,2fr)_auto] items-center gap-3 text-xs"><span className="truncate text-zinc-600" title={label}>{label}</span><div className="h-2 rounded-sm bg-zinc-100"><div className="h-full rounded-sm bg-accent-blue" style={{ width: `${Math.max(2, (value / max) * 100)}%` }} /></div><span className="font-mono text-zinc-500">{money(value)}</span></div>; }
