# RetailOptima — Retail Product & Inventory Optimization

A polished React/TanStack UI for the **Smart Retail Product & Inventory Optimization System**. This version uses the visual layout, navigation, cards, tables, and interaction style of `optimal-shelf-sight-main`, while preserving the core functionality of the updated `retail-optimization` project.

## What is included

- CSV / XLSX / XLS transaction upload
- Bundled UCI Online Retail sample dataset
- Transaction cleaning and validation
- Product-level aggregation by StockCode
- Dynamic storage sizing:
  - Tiny / Small / Medium / Large tiers
  - size-word modifiers
  - pack/set multipliers
  - editable storage parameters
- Product Mix Optimization using Linear Programming
- Promotional Product Selection using 0–1 Integer Programming / Knapsack
- Budget and storage constraints
- Historical-demand upper bounds for product mix
- Maximum product-count constraint for promotions
- Budget sensitivity analysis
- Descriptive analysis:
  - revenue
  - quantity
  - transactions
  - products
  - customers
  - countries
  - monthly sales trend
  - sales by country
  - top products by revenue
- Responsive desktop/mobile navigation
- Client-side processing; uploaded data is processed in the browser

## Run locally

### Prerequisites

- Node.js 20+ recommended
- npm 10+ recommended

### Install

```bash
npm install
```

If you prefer Bun:

```bash
bun install
```

### Start development server

```bash
npm run dev
```

Then open the local URL shown by Vite.

### Production build

```bash
npm run build
```

### Preview production build

```bash
npm run preview
```

### Lint

```bash
npm run lint
```

## Dataset schema

The application expects these columns:

```text
InvoiceNo
StockCode
Description
Quantity
InvoiceDate
UnitPrice
CustomerID
Country
```

The bundled `public/sample_online_retail.csv` can be loaded directly from the Dashboard.

## Optimization models

### Model A — Product Mix

Maximizes expected profit:

```text
max Σ margin_i × x_i

subject to:
Σ cost_i × x_i ≤ Budget
Σ storage_i × x_i ≤ Storage Capacity
0 ≤ x_i ≤ historical demand_i
```

### Model B — Promotional Selection

Maximizes selected-product margin:

```text
max Σ margin_i × y_i

subject to:
Σ cost_i × y_i ≤ Budget
Σ storage_i × y_i ≤ Storage Capacity
Σ y_i ≤ Maximum Products
y_i ∈ {0,1}
```

## Project structure

```text
src/
  routes/
    index.tsx          # Main application and UI
  components/ui/       # Reusable UI components
  styles.css           # Design system / theme
public/
  sample_online_retail.csv
```

## Notes

- Procurement cost is modeled as a configurable percentage of observed selling price because the Online Retail dataset does not contain procurement cost.
- Storage is a deterministic text-derived proxy, not measured physical volume.
- Product storage rules are editable from the Inventory page.
- The optimization operates on the top-N products by historical revenue to keep the decision model compact and interpretable.

## Verification

The project source was reviewed after combining the UI and functionality. The optimization, preprocessing, storage-sizing, upload, sample-data, navigation, and analysis paths are wired into the React application.

If dependency installation fails because of network restrictions, retry `npm install` when internet/package-registry access is available before running the development or production build.
