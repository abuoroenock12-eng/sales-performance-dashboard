import React, { useState, useMemo, useEffect, useCallback } from "react";
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, ShoppingBag, Package, Users, BarChart3, Settings,
  Search, Bell, Moon, Sun, Menu, X, ChevronDown, ChevronUp, ChevronLeft,
  ChevronRight, ChevronsLeft, ArrowUp, ArrowDown, Download, RotateCcw,
  DollarSign, ShoppingCart, TrendingUp, CreditCard, Percent, Filter,
  ArrowUpDown, FileSpreadsheet, FileText, Boxes, CircleUser, Sparkles,
} from "lucide-react";

/* ============================================================================
   1. STATIC REFERENCE DATA
   ========================================================================= */

const CATEGORIES = [
  "Electronics", "Apparel & Fashion", "Home & Kitchen", "Sports & Outdoors",
  "Beauty & Personal Care", "Toys & Games", "Books & Stationery", "Automotive & Tools",
];

const REGIONS = [
  { name: "North America", weight: 34 },
  { name: "Europe", weight: 26 },
  { name: "Asia Pacific", weight: 22 },
  { name: "Latin America", weight: 11 },
  { name: "Middle East & Africa", weight: 7 },
];

const PAYMENT_METHODS = [
  { name: "Credit Card", weight: 38 },
  { name: "Debit Card", weight: 21 },
  { name: "PayPal", weight: 19 },
  { name: "Digital Wallet", weight: 13 },
  { name: "Bank Transfer", weight: 9 },
];

const STATUSES = [
  { name: "Completed", weight: 64 },
  { name: "Processing", weight: 13 },
  { name: "Pending", weight: 10 },
  { name: "Cancelled", weight: 8 },
  { name: "Refunded", weight: 5 },
];

const SALES_REPS = [
  "Maria Chen", "James Okafor", "Priya Nair", "Lucas Bergström", "Fatima Al-Sayed",
  "Ethan Walker", "Sofia Marchetti", "Daniel Kim", "Grace Achieng", "Noah Bennett",
  "Isabela Rocha", "Ravi Deshmukh",
];

const PRODUCTS_BY_CATEGORY = {
  "Electronics": [
    ["Wireless Noise-Cancelling Headphones", 89, 249],
    ["4K Ultra HD Streaming Stick", 29, 69],
    ["Portable Bluetooth Speaker", 24, 129],
    ["Smart Fitness Watch", 59, 219],
    ["27-inch QHD Monitor", 179, 389],
    ["Mechanical Keyboard", 39, 149],
    ["Wireless Charging Pad", 15, 45],
    ["Compact Mirrorless Camera", 349, 899],
  ],
  "Apparel & Fashion": [
    ["Organic Cotton T-Shirt", 12, 28],
    ["Slim Fit Denim Jeans", 32, 78],
    ["Merino Wool Sweater", 45, 110],
    ["Running Performance Jacket", 55, 140],
    ["Leather Chelsea Boots", 89, 210],
    ["Classic Trench Coat", 95, 240],
    ["Everyday Canvas Sneakers", 38, 85],
    ["Silk Blend Scarf", 22, 60],
  ],
  "Home & Kitchen": [
    ["Stainless Steel Cookware Set", 79, 199],
    ["Programmable Espresso Machine", 129, 349],
    ["Memory Foam Pillow", 22, 55],
    ["Ceramic Non-Stick Skillet", 18, 48],
    ["Robot Vacuum Cleaner", 149, 399],
    ["Air Fryer, 6-Quart", 59, 129],
    ["Egyptian Cotton Sheet Set", 35, 95],
    ["Stand Mixer", 119, 329],
  ],
  "Sports & Outdoors": [
    ["Adjustable Dumbbell Set", 69, 199],
    ["4-Person Camping Tent", 89, 229],
    ["Trail Running Shoes", 55, 129],
    ["Insulated Water Bottle", 14, 34],
    ["Yoga Mat, Extra Thick", 19, 45],
    ["Carbon Fiber Bike Frame", 299, 799],
    ["Folding Kayak", 249, 599],
    ["Resistance Band Kit", 12, 32],
  ],
  "Beauty & Personal Care": [
    ["Vitamin C Serum", 18, 42],
    ["Ceramic Hair Straightener", 29, 79],
    ["Electric Toothbrush", 24, 69],
    ["Hydrating Face Moisturizer", 15, 38],
    ["Aromatherapy Diffuser Set", 22, 55],
    ["Professional Nail Kit", 19, 49],
    ["Men's Grooming Kit", 26, 62],
    ["Mineral Sunscreen SPF 50", 12, 26],
  ],
  "Toys & Games": [
    ["1000-Piece Jigsaw Puzzle", 12, 26],
    ["Building Block Set, 800pc", 29, 69],
    ["Remote Control Race Car", 24, 79],
    ["Strategy Board Game", 22, 48],
    ["Plush Animal Collection", 10, 28],
    ["Beginner Coding Robot Kit", 49, 119],
    ["Art & Craft Studio Set", 19, 45],
    ["Wooden Train Set", 32, 74],
  ],
  "Books & Stationery": [
    ["Hardcover Notebook, Set of 3", 9, 22],
    ["Fountain Pen Gift Set", 18, 45],
    ["Desk Organizer Set", 14, 32],
    ["Leather Journal", 16, 38],
    ["Watercolor Paint Set", 12, 29],
    ["Planner & Calendar Bundle", 15, 34],
    ["Premium Sketchbook", 10, 24],
    ["Fine-Tip Marker Set, 60pc", 13, 28],
  ],
  "Automotive & Tools": [
    ["Cordless Drill Driver Kit", 49, 129],
    ["Digital Tire Pressure Gauge", 8, 19],
    ["Car Dash Camera", 34, 89],
    ["All-Weather Floor Mats", 29, 69],
    ["Portable Jump Starter", 45, 109],
    ["Mechanic's Tool Set, 120pc", 59, 149],
    ["Microfiber Detailing Kit", 16, 38],
    ["LED Work Light", 12, 32],
  ],
};

const FIRST_NAMES = [
  "Alex", "Jordan", "Taylor", "Morgan", "Casey", "Riley", "Jamie", "Cameron",
  "Avery", "Reese", "Hana", "Liam", "Olivia", "Noah", "Emma", "Mateo",
  "Amara", "Kwame", "Yuki", "Diego", "Priya", "Omar", "Ines", "Lars",
  "Nadia", "Sven", "Aisha", "Carlos", "Mei", "Tariq",
];
const LAST_NAMES = [
  "Johnson", "Williams", "Brown", "Garcia", "Martinez", "Nguyen", "Kim",
  "Patel", "Silva", "Andersson", "Rossi", "Müller", "Dubois", "Kowalski",
  "Osei", "Haddad", "Yamamoto", "Ivanov", "Costa", "Fernandez", "O'Brien",
  "Novak", "Adeyemi", "Santos", "Bakker", "Larsen", "Choi", "Hassan",
];
const COMPANY_WORDS_A = ["Nova", "Bright", "Summit", "Crest", "Cobalt", "Lumen", "Vertex", "Ridge", "Harbor", "Aster"];
const COMPANY_WORDS_B = ["Retail", "Logistics", "Holdings", "Trading Co.", "Group", "Supply", "Partners", "Ventures", "Distribution", "Works"];

/* ============================================================================
   2. SEEDED RANDOM DATA GENERATION
   ========================================================================= */

function mulberry32(seed) {
  return function () {
    seed |= 0; seed = (seed + 0x6D2B79F5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function weightedPick(rand, items) {
  const total = items.reduce((s, i) => s + i.weight, 0);
  let r = rand() * total;
  for (const item of items) {
    if (r < item.weight) return item.name;
    r -= item.weight;
  }
  return items[items.length - 1].name;
}

function buildCustomerPool(rand, count) {
  const names = new Set();
  while (names.size < count) {
    if (rand() < 0.72) {
      names.add(`${FIRST_NAMES[Math.floor(rand() * FIRST_NAMES.length)]} ${LAST_NAMES[Math.floor(rand() * LAST_NAMES.length)]}`);
    } else {
      names.add(`${COMPANY_WORDS_A[Math.floor(rand() * COMPANY_WORDS_A.length)]} ${COMPANY_WORDS_B[Math.floor(rand() * COMPANY_WORDS_B.length)]}`);
    }
  }
  return Array.from(names);
}

function generateSalesData(count) {
  const rand = mulberry32(1337);
  const customers = buildCustomerPool(rand, 420);
  const today = new Date("2026-08-06T00:00:00");
  const rangeDays = 395; // ~13 months of history
  const records = [];

  for (let i = 0; i < count; i++) {
    // Skew toward more recent dates to simulate month-over-month growth
    const skew = Math.pow(rand(), 0.62);
    const dayOffset = Math.floor(skew * rangeDays);
    const date = new Date(today);
    date.setDate(date.getDate() - dayOffset);
    // Trim weekends slightly by re-rolling ~30% of the time
    if ((date.getDay() === 0 || date.getDay() === 6) && rand() < 0.35) {
      date.setDate(date.getDate() - (1 + Math.floor(rand() * 2)));
    }

    const category = CATEGORIES[Math.floor(rand() * CATEGORIES.length)];
    const productList = PRODUCTS_BY_CATEGORY[category];
    const [productName, priceMin, priceMax] = productList[Math.floor(rand() * productList.length)];
    const unitPrice = Math.round((priceMin + rand() * (priceMax - priceMin)) * 100) / 100;
    const quantity = 1 + Math.floor(rand() * rand() * 9);
    const totalAmount = Math.round(unitPrice * quantity * 100) / 100;

    records.push({
      id: `ORD-${(100000 + i).toString()}`,
      date,
      dateStr: date.toISOString().slice(0, 10),
      customer: customers[Math.floor(rand() * customers.length)],
      product: productName,
      category,
      quantity,
      unitPrice,
      totalAmount,
      region: weightedPick(rand, REGIONS),
      salesRep: SALES_REPS[Math.floor(rand() * SALES_REPS.length)],
      paymentMethod: weightedPick(rand, PAYMENT_METHODS),
      status: weightedPick(rand, STATUSES),
    });
  }
  return records.sort((a, b) => b.date - a.date);
}

/* ============================================================================
   3. FORMATTING HELPERS
   ========================================================================= */

const fmtCurrency = (v, compact = false) =>
  new Intl.NumberFormat("en-US", {
    style: "currency", currency: "USD",
    notation: compact ? "compact" : "standard",
    maximumFractionDigits: compact ? 1 : 2,
  }).format(v || 0);

const fmtNumber = (v, compact = false) =>
  new Intl.NumberFormat("en-US", {
    notation: compact ? "compact" : "standard",
    maximumFractionDigits: compact ? 1 : 0,
  }).format(v || 0);

const fmtDate = (d) =>
  new Intl.DateTimeFormat("en-US", { month: "short", day: "2-digit", year: "numeric" }).format(d);

const MONTH_LABELS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/* ============================================================================
   4. THEME TOKENS (light / dark) — standard Tailwind palette only
   ========================================================================= */

function useTheme(isDark) {
  return {
    appBg: isDark ? "bg-slate-950" : "bg-slate-50",
    surface: isDark ? "bg-slate-900" : "bg-white",
    surfaceAlt: isDark ? "bg-slate-800" : "bg-slate-50",
    border: isDark ? "border-slate-800" : "border-slate-200",
    borderStrong: isDark ? "border-slate-700" : "border-slate-300",
    textPrimary: isDark ? "text-slate-100" : "text-slate-900",
    textSecondary: isDark ? "text-slate-400" : "text-slate-500",
    textMuted: isDark ? "text-slate-500" : "text-slate-400",
    hover: isDark ? "hover:bg-slate-800" : "hover:bg-slate-100",
    sidebarBg: isDark ? "bg-slate-950" : "bg-blue-950",
    sidebarBorder: isDark ? "border-slate-800" : "border-blue-900",
    sidebarText: "text-blue-200",
    sidebarTextMuted: "text-blue-400",
    inputBg: isDark ? "bg-slate-800" : "bg-white",
    ring: "focus:ring-2 focus:ring-blue-500 focus:border-blue-500",
  };
}

const CHART_BLUES = ["#2563EB", "#0EA5E9", "#60A5FA", "#1D4ED8", "#38BDF8", "#93C5FD", "#1E40AF", "#7DD3FC"];
const STATUS_COLORS = {
  Completed: { bg: "bg-emerald-100", text: "text-emerald-700", dot: "bg-emerald-500", bgDark: "bg-emerald-500/15", textDark: "text-emerald-400" },
  Processing: { bg: "bg-blue-100", text: "text-blue-700", dot: "bg-blue-500", bgDark: "bg-blue-500/15", textDark: "text-blue-400" },
  Pending: { bg: "bg-amber-100", text: "text-amber-700", dot: "bg-amber-500", bgDark: "bg-amber-500/15", textDark: "text-amber-400" },
  Cancelled: { bg: "bg-red-100", text: "text-red-700", dot: "bg-red-500", bgDark: "bg-red-500/15", textDark: "text-red-400" },
  Refunded: { bg: "bg-purple-100", text: "text-purple-700", dot: "bg-purple-500", bgDark: "bg-purple-500/15", textDark: "text-purple-400" },
};

/* ============================================================================
   5. AGGREGATION HELPERS
   ========================================================================= */

const isLostRevenue = (status) => status === "Cancelled" || status === "Refunded";

function computeKpis(rows) {
  const totalOrders = rows.length;
  const totalUnits = rows.reduce((s, r) => s + r.quantity, 0);
  const totalRevenue = rows.reduce((s, r) => s + (isLostRevenue(r.status) ? 0 : r.totalAmount), 0);
  const totalCustomers = new Set(rows.map((r) => r.customer)).size;
  const revenueOrders = rows.filter((r) => !isLostRevenue(r.status)).length;
  const avgOrderValue = revenueOrders ? totalRevenue / revenueOrders : 0;

  // Monthly buckets for growth + sparklines
  const byMonth = {};
  rows.forEach((r) => {
    const key = `${r.date.getFullYear()}-${String(r.date.getMonth() + 1).padStart(2, "0")}`;
    if (!byMonth[key]) byMonth[key] = { revenue: 0, orders: 0, units: 0, customers: new Set() };
    if (!isLostRevenue(r.status)) byMonth[key].revenue += r.totalAmount;
    byMonth[key].orders += 1;
    byMonth[key].units += r.quantity;
    byMonth[key].customers.add(r.customer);
  });
  const months = Object.keys(byMonth).sort();
  const last = months[months.length - 1];
  const prev = months[months.length - 2];
  const growth = last && prev && byMonth[prev].revenue > 0
    ? ((byMonth[last].revenue - byMonth[prev].revenue) / byMonth[prev].revenue) * 100
    : 0;

  const sparkRevenue = months.map((m) => ({ m, value: byMonth[m].revenue }));
  const sparkOrders = months.map((m) => ({ m, value: byMonth[m].orders }));
  const sparkUnits = months.map((m) => ({ m, value: byMonth[m].units }));
  const sparkCustomers = months.map((m) => ({ m, value: byMonth[m].customers.size }));
  const sparkAov = months.map((m) => ({ m, value: byMonth[m].orders ? byMonth[m].revenue / byMonth[m].orders : 0 }));

  return {
    totalOrders, totalUnits, totalRevenue, totalCustomers, avgOrderValue, growth,
    sparkRevenue, sparkOrders, sparkUnits, sparkCustomers, sparkAov,
  };
}

function computeMonthlyTrend(rows) {
  const byMonth = {};
  rows.forEach((r) => {
    const key = `${r.date.getFullYear()}-${String(r.date.getMonth() + 1).padStart(2, "0")}`;
    if (!byMonth[key]) byMonth[key] = { revenue: 0, orders: 0, y: r.date.getFullYear(), mo: r.date.getMonth() };
    if (!isLostRevenue(r.status)) byMonth[key].revenue += r.totalAmount;
    byMonth[key].orders += 1;
  });
  return Object.keys(byMonth).sort().map((k) => ({
    label: `${MONTH_LABELS[byMonth[k].mo]} '${String(byMonth[k].y).slice(2)}`,
    revenue: Math.round(byMonth[k].revenue),
    orders: byMonth[k].orders,
  }));
}

function computeCategoryBreakdown(rows) {
  const byCat = {};
  rows.forEach((r) => {
    if (!byCat[r.category]) byCat[r.category] = { revenue: 0, orders: 0 };
    if (!isLostRevenue(r.status)) byCat[r.category].revenue += r.totalAmount;
    byCat[r.category].orders += 1;
  });
  return Object.entries(byCat)
    .map(([name, v]) => ({ name, revenue: Math.round(v.revenue), orders: v.orders }))
    .sort((a, b) => b.revenue - a.revenue);
}

function computeRegionBreakdown(rows) {
  const byRegion = {};
  rows.forEach((r) => {
    if (!byRegion[r.region]) byRegion[r.region] = { revenue: 0, orders: 0 };
    if (!isLostRevenue(r.status)) byRegion[r.region].revenue += r.totalAmount;
    byRegion[r.region].orders += 1;
  });
  return Object.entries(byRegion)
    .map(([name, v]) => ({ name, revenue: Math.round(v.revenue), orders: v.orders }))
    .sort((a, b) => b.revenue - a.revenue);
}

function computeTopProducts(rows) {
  const byProduct = {};
  rows.forEach((r) => {
    if (!byProduct[r.product]) byProduct[r.product] = { revenue: 0, units: 0 };
    if (!isLostRevenue(r.status)) byProduct[r.product].revenue += r.totalAmount;
    byProduct[r.product].units += r.quantity;
  });
  return Object.entries(byProduct)
    .map(([name, v]) => ({ name, revenue: Math.round(v.revenue), units: v.units }))
    .sort((a, b) => b.revenue - a.revenue)
    .slice(0, 10)
    .reverse();
}

function computeDailyPerformance(rows, days = 30) {
  const byDay = {};
  rows.forEach((r) => {
    if (!byDay[r.dateStr]) byDay[r.dateStr] = { revenue: 0, orders: 0 };
    if (!isLostRevenue(r.status)) byDay[r.dateStr].revenue += r.totalAmount;
    byDay[r.dateStr].orders += 1;
  });
  return Object.keys(byDay).sort().slice(-days).map((d) => ({
    label: d.slice(5),
    revenue: Math.round(byDay[d].revenue),
    orders: byDay[d].orders,
  }));
}

/* ============================================================================
   6. SMALL REUSABLE UI PIECES
   ========================================================================= */

function ChartCard({ title, subtitle, t, isDark, height = 300, children, right }) {
  return (
    <div className={`rounded-2xl border ${t.border} ${t.surface} p-5 shadow-sm transition-colors`}>
      <div className="flex items-start justify-between mb-4 gap-3">
        <div>
          <h3 className={`text-sm font-semibold ${t.textPrimary}`}>{title}</h3>
          {subtitle && <p className={`text-xs ${t.textSecondary} mt-0.5`}>{subtitle}</p>}
        </div>
        {right}
      </div>
      <div style={{ height }}>
        <ResponsiveContainer width="100%" height="100%">
          {children}
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function CustomTooltip({ active, payload, label, isDark, formatter }) {
  if (!active || !payload || !payload.length) return null;
  return (
    <div className={`rounded-lg border px-3 py-2 text-xs shadow-lg ${isDark ? "bg-slate-800 border-slate-700 text-slate-100" : "bg-white border-slate-200 text-slate-800"}`}>
      {label && <p className="font-semibold mb-1">{label}</p>}
      {payload.map((p, i) => (
        <p key={i} style={{ color: p.color }}>
          {p.name}: <span className="font-semibold">{formatter ? formatter(p.value, p.name) : p.value}</span>
        </p>
      ))}
    </div>
  );
}

const KPI_ICONS = { revenue: DollarSign, units: Boxes, orders: ShoppingCart, customers: Users, aov: CreditCard, growth: Percent };

function KpiCard({ t, isDark, icon, label, value, delta, sparkData, color = "#2563EB", suffix }) {
  const Icon = KPI_ICONS[icon] || DollarSign;
  const positive = delta === undefined || delta >= 0;
  return (
    <div className={`relative overflow-hidden rounded-2xl border ${t.border} ${t.surface} p-4 shadow-sm hover:shadow-md transition-all duration-200`}>
      <div className="absolute top-0 left-0 right-0 h-1" style={{ background: `linear-gradient(90deg, ${color}, ${color}55)` }} />
      <div className="flex items-center justify-between mb-3">
        <div className={`h-9 w-9 rounded-xl flex items-center justify-center`} style={{ backgroundColor: `${color}1A` }}>
          <Icon size={17} style={{ color }} />
        </div>
        {delta !== undefined && (
          <span className={`inline-flex items-center gap-0.5 text-xs font-semibold px-1.5 py-0.5 rounded-md ${positive ? "text-emerald-600 bg-emerald-50" : "text-red-600 bg-red-50"} ${isDark ? (positive ? "bg-emerald-500/10" : "bg-red-500/10") : ""}`}>
            {positive ? <ArrowUp size={11} /> : <ArrowDown size={11} />}
            {Math.abs(delta).toFixed(1)}%
          </span>
        )}
      </div>
      <p className={`text-xs font-medium ${t.textSecondary} mb-1`}>{label}</p>
      <p className={`text-xl font-bold ${t.textPrimary} tracking-tight`}>{value}{suffix}</p>
      {sparkData && sparkData.length > 1 && (
        <div className="h-8 mt-2 -mx-1">
          <ResponsiveContainer width="100%" height="100%">
            <AreaChart data={sparkData}>
              <defs>
                <linearGradient id={`spark-${label.replace(/\s/g, "")}`} x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={color} stopOpacity={0.35} />
                  <stop offset="100%" stopColor={color} stopOpacity={0} />
                </linearGradient>
              </defs>
              <Area type="monotone" dataKey="value" stroke={color} strokeWidth={1.75} fill={`url(#spark-${label.replace(/\s/g, "")})`} isAnimationActive={false} />
            </AreaChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

function Select({ t, value, onChange, options, label }) {
  return (
    <label className="flex flex-col gap-1 min-w-0">
      <span className={`text-[11px] font-medium ${t.textSecondary}`}>{label}</span>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className={`text-sm rounded-lg border ${t.border} ${t.inputBg} ${t.textPrimary} px-2.5 py-1.5 outline-none ${t.ring} transition-colors cursor-pointer`}
      >
        {options.map((o) => (
          <option key={o} value={o}>{o}</option>
        ))}
      </select>
    </label>
  );
}

/* ============================================================================
   7. SIDEBAR
   ========================================================================= */

const NAV_ITEMS = [
  { label: "Dashboard", icon: LayoutDashboard, active: true },
  { label: "Orders", icon: ShoppingBag },
  { label: "Products", icon: Package },
  { label: "Customers", icon: Users },
  { label: "Reports", icon: BarChart3 },
  { label: "Settings", icon: Settings },
];

function Sidebar({ t, collapsed, mobileOpen, setMobileOpen }) {
  return (
    <>
      {mobileOpen && (
        <div className="fixed inset-0 bg-black/50 z-40 lg:hidden" onClick={() => setMobileOpen(false)} />
      )}
      <aside
        className={`${t.sidebarBg} border-r ${t.sidebarBorder} fixed lg:sticky top-0 h-screen z-50 transition-all duration-200 flex flex-col
        ${collapsed ? "lg:w-[76px]" : "lg:w-60"}
        ${mobileOpen ? "translate-x-0 w-64" : "-translate-x-full lg:translate-x-0 w-64"}`}
      >
        <div className="flex items-center gap-2.5 px-5 h-16 border-b border-white/10 shrink-0">
          <div className="h-8 w-8 rounded-lg bg-blue-500 flex items-center justify-center shrink-0">
            <Sparkles size={16} className="text-white" />
          </div>
          {!collapsed && <span className="text-white font-semibold text-[15px] tracking-tight truncate">Sales Performance Dashboard</span>}
          <button className="ml-auto lg:hidden text-blue-300" onClick={() => setMobileOpen(false)}>
            <X size={18} />
          </button>
        </div>

        <nav className="flex-1 px-3 py-4 space-y-1 overflow-y-auto">
          {NAV_ITEMS.map((item) => (
            <button
              key={item.label}
              className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors
              ${item.active ? "bg-blue-600 text-white shadow-sm" : `${t.sidebarText} hover:bg-white/5 hover:text-white`}`}
              title={collapsed ? item.label : undefined}
            >
              <item.icon size={18} className="shrink-0" />
              {!collapsed && <span className="truncate">{item.label}</span>}
            </button>
          ))}
        </nav>

        <div className={`px-3 py-4 border-t border-white/10 ${collapsed ? "hidden lg:hidden" : ""}`}>
          <div className="rounded-xl bg-white/5 p-3">
            <p className="text-[11px] font-semibold text-blue-100 mb-1">Data refresh</p>
            <p className={`text-[11px] ${t.sidebarTextMuted}`}>Synced from order management system every 15 minutes.</p>
          </div>
        </div>
      </aside>
    </>
  );
}

/* ============================================================================
   8. HEADER
   ========================================================================= */

function Header({ t, isDark, setIsDark, search, setSearch, setMobileOpen, collapsed, setCollapsed, liveLabel }) {
  const [notifOpen, setNotifOpen] = useState(false);
  const notifications = [
    { title: "Monthly target reached", body: "North America region hit 102% of its revenue goal.", time: "2h ago" },
    { title: "New large order", body: "Order ORD-104821 placed for $4,280.00.", time: "5h ago" },
    { title: "Refund processed", body: "Refund issued for order ORD-102113.", time: "1d ago" },
  ];

  return (
    <header className={`sticky top-0 z-30 ${t.surface} border-b ${t.border} h-16 flex items-center gap-3 px-4 lg:px-6 transition-colors`}>
      <button className={`lg:hidden ${t.textSecondary}`} onClick={() => setMobileOpen(true)}>
        <Menu size={22} />
      </button>
      <button className={`hidden lg:flex ${t.textSecondary} ${t.hover} p-1.5 rounded-lg`} onClick={() => setCollapsed(!collapsed)}>
        <ChevronsLeft size={18} className={`transition-transform ${collapsed ? "rotate-180" : ""}`} />
      </button>

      <div className="relative flex-1 max-w-md">
        <Search size={16} className={`absolute left-3 top-1/2 -translate-y-1/2 ${t.textMuted}`} />
        <input
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="Search orders, customers, products…"
          className={`w-full text-sm rounded-lg border ${t.border} ${t.inputBg} ${t.textPrimary} pl-9 pr-3 py-2 outline-none ${t.ring} transition-colors`}
        />
      </div>

      <div className="hidden md:flex items-center gap-1.5 ml-auto mr-1">
        <span className="relative flex h-2 w-2">
          <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75" />
          <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500" />
        </span>
        <span className={`text-xs ${t.textSecondary}`}>Live &middot; updated {liveLabel}</span>
      </div>

      <button
        onClick={() => setIsDark(!isDark)}
        className={`p-2 rounded-lg ${t.textSecondary} ${t.hover} transition-colors`}
        title="Toggle theme"
      >
        {isDark ? <Sun size={18} /> : <Moon size={18} />}
      </button>

      <div className="relative">
        <button onClick={() => setNotifOpen(!notifOpen)} className={`p-2 rounded-lg ${t.textSecondary} ${t.hover} relative transition-colors`}>
          <Bell size={18} />
          <span className="absolute top-1.5 right-1.5 h-1.5 w-1.5 rounded-full bg-red-500" />
        </button>
        {notifOpen && (
          <>
            <div className="fixed inset-0 z-10" onClick={() => setNotifOpen(false)} />
            <div className={`absolute right-0 mt-2 w-72 rounded-xl border ${t.border} ${t.surface} shadow-lg z-20 overflow-hidden`}>
              <div className={`px-4 py-3 border-b ${t.border}`}>
                <p className={`text-sm font-semibold ${t.textPrimary}`}>Notifications</p>
              </div>
              {notifications.map((n, i) => (
                <div key={i} className={`px-4 py-3 border-b ${t.border} last:border-0 ${t.hover} cursor-pointer transition-colors`}>
                  <p className={`text-sm font-medium ${t.textPrimary}`}>{n.title}</p>
                  <p className={`text-xs ${t.textSecondary} mt-0.5`}>{n.body}</p>
                  <p className={`text-[11px] ${t.textMuted} mt-1`}>{n.time}</p>
                </div>
              ))}
            </div>
          </>
        )}
      </div>

      <div className={`flex items-center gap-2 pl-2 border-l ${t.border}`}>
        <div className="h-8 w-8 rounded-full bg-blue-600 flex items-center justify-center text-white text-xs font-semibold">AE</div>
        <div className="hidden sm:block leading-tight">
          <p className={`text-xs font-semibold ${t.textPrimary}`}>Abuoro Enock</p>
          <p className={`text-[11px] ${t.textSecondary}`}>Sales Ops</p>
        </div>
      </div>
    </header>
  );
}

/* ============================================================================
   9. FILTER BAR
   ========================================================================= */

function FilterBar({ t, filters, setFilters, options, resultCount }) {
  const update = (key) => (val) => setFilters((f) => ({ ...f, [key]: val }));
  const reset = () => setFilters({
    dateFrom: options.minDate, dateTo: options.maxDate, region: "All Regions",
    category: "All Categories", salesRep: "All Reps", customer: "All Customers",
    paymentMethod: "All Payment Methods",
  });

  return (
    <div className={`rounded-2xl border ${t.border} ${t.surface} p-4 shadow-sm`}>
      <div className="flex items-center gap-2 mb-3">
        <Filter size={15} className="text-blue-600" />
        <h3 className={`text-sm font-semibold ${t.textPrimary}`}>Filters</h3>
        <span className={`text-xs ${t.textSecondary}`}>&middot; {fmtNumber(resultCount)} orders match</span>
        <button onClick={reset} className="ml-auto flex items-center gap-1 text-xs font-medium text-blue-600 hover:text-blue-700">
          <RotateCcw size={12} /> Reset
        </button>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-7 gap-3">
        <label className="flex flex-col gap-1">
          <span className={`text-[11px] font-medium ${t.textSecondary}`}>From</span>
          <input type="date" value={filters.dateFrom} min={options.minDate} max={options.maxDate}
            onChange={(e) => update("dateFrom")(e.target.value)}
            className={`text-sm rounded-lg border ${t.border} ${t.inputBg} ${t.textPrimary} px-2.5 py-1.5 outline-none ${t.ring}`} />
        </label>
        <label className="flex flex-col gap-1">
          <span className={`text-[11px] font-medium ${t.textSecondary}`}>To</span>
          <input type="date" value={filters.dateTo} min={options.minDate} max={options.maxDate}
            onChange={(e) => update("dateTo")(e.target.value)}
            className={`text-sm rounded-lg border ${t.border} ${t.inputBg} ${t.textPrimary} px-2.5 py-1.5 outline-none ${t.ring}`} />
        </label>
        <Select t={t} label="Region" value={filters.region} onChange={update("region")} options={["All Regions", ...options.regions]} />
        <Select t={t} label="Category" value={filters.category} onChange={update("category")} options={["All Categories", ...options.categories]} />
        <Select t={t} label="Sales Rep" value={filters.salesRep} onChange={update("salesRep")} options={["All Reps", ...options.reps]} />
        <label className="flex flex-col gap-1 min-w-0">
          <span className={`text-[11px] font-medium ${t.textSecondary}`}>Customer</span>
          <input list="customer-options" value={filters.customer === "All Customers" ? "" : filters.customer}
            placeholder="All Customers"
            onChange={(e) => update("customer")(e.target.value || "All Customers")}
            className={`text-sm rounded-lg border ${t.border} ${t.inputBg} ${t.textPrimary} px-2.5 py-1.5 outline-none ${t.ring}`} />
          <datalist id="customer-options">
            {options.customers.map((c) => <option key={c} value={c} />)}
          </datalist>
        </label>
        <Select t={t} label="Payment Method" value={filters.paymentMethod} onChange={update("paymentMethod")} options={["All Payment Methods", ...options.paymentMethods]} />
      </div>
    </div>
  );
}

/* ============================================================================
   10. DATA TABLE
   ========================================================================= */

const TABLE_COLUMNS = [
  { key: "id", label: "Order ID" },
  { key: "dateStr", label: "Order Date" },
  { key: "customer", label: "Customer" },
  { key: "product", label: "Product" },
  { key: "category", label: "Category" },
  { key: "quantity", label: "Qty" },
  { key: "unitPrice", label: "Unit Price" },
  { key: "totalAmount", label: "Total" },
  { key: "region", label: "Region" },
  { key: "salesRep", label: "Sales Rep" },
  { key: "paymentMethod", label: "Payment" },
  { key: "status", label: "Status" },
];

function StatusBadge({ status, isDark }) {
  const c = STATUS_COLORS[status] || STATUS_COLORS.Completed;
  return (
    <span className={`inline-flex items-center gap-1.5 px-2 py-0.5 rounded-full text-xs font-medium ${isDark ? `${c.bgDark} ${c.textDark}` : `${c.bg} ${c.text}`}`}>
      <span className={`h-1.5 w-1.5 rounded-full ${c.dot}`} />
      {status}
    </span>
  );
}

function DataTable({ t, isDark, rows }) {
  const [search, setSearch] = useState("");
  const [sort, setSort] = useState({ key: "dateStr", dir: "desc" });
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    let out = rows;
    if (q) {
      out = rows.filter((r) =>
        r.id.toLowerCase().includes(q) ||
        r.customer.toLowerCase().includes(q) ||
        r.product.toLowerCase().includes(q) ||
        r.category.toLowerCase().includes(q) ||
        r.region.toLowerCase().includes(q) ||
        r.salesRep.toLowerCase().includes(q)
      );
    }
    return out;
  }, [rows, search]);

  const sorted = useMemo(() => {
    const arr = [...filtered];
    arr.sort((a, b) => {
      let av = a[sort.key], bv = b[sort.key];
      if (typeof av === "string") { av = av.toLowerCase(); bv = bv.toLowerCase(); }
      if (av < bv) return sort.dir === "asc" ? -1 : 1;
      if (av > bv) return sort.dir === "asc" ? 1 : -1;
      return 0;
    });
    return arr;
  }, [filtered, sort]);

  useEffect(() => { setPage(1); }, [search, rows]);

  const totalPages = Math.max(1, Math.ceil(sorted.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const pageRows = sorted.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  const toggleSort = (key) => setSort((s) => ({ key, dir: s.key === key && s.dir === "asc" ? "desc" : "asc" }));

  const exportCsv = () => {
    const headers = TABLE_COLUMNS.map((c) => c.label);
    const lines = [headers.join(",")];
    sorted.forEach((r) => {
      const row = [r.id, r.dateStr, r.customer, r.product, r.category, r.quantity, r.unitPrice, r.totalAmount, r.region, r.salesRep, r.paymentMethod, r.status]
        .map((v) => `"${String(v).replace(/"/g, '""')}"`).join(",");
      lines.push(row);
    });
    const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "sales_data.csv"; a.click();
    URL.revokeObjectURL(url);
  };

  const exportExcel = () => {
    const data = sorted.map((r) => ({
      "Order ID": r.id, "Order Date": r.dateStr, "Customer Name": r.customer, "Product": r.product,
      "Category": r.category, "Quantity": r.quantity, "Unit Price": r.unitPrice, "Total Amount": r.totalAmount,
      "Region": r.region, "Sales Representative": r.salesRep, "Payment Method": r.paymentMethod, "Order Status": r.status,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sales");
    XLSX.writeFile(wb, "sales_data.xlsx");
  };

  return (
    <div className={`rounded-2xl border ${t.border} ${t.surface} shadow-sm overflow-hidden`}>
      <div className={`flex flex-col sm:flex-row sm:items-center gap-3 p-4 border-b ${t.border}`}>
        <h3 className={`text-sm font-semibold ${t.textPrimary}`}>Sales Orders</h3>
        <span className={`text-xs ${t.textSecondary}`}>{fmtNumber(sorted.length)} records</span>
        <div className="relative sm:ml-auto w-full sm:w-64">
          <Search size={14} className={`absolute left-3 top-1/2 -translate-y-1/2 ${t.textMuted}`} />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search table…"
            className={`w-full text-sm rounded-lg border ${t.border} ${t.inputBg} ${t.textPrimary} pl-8 pr-3 py-1.5 outline-none ${t.ring}`}
          />
        </div>
        <div className="flex gap-2">
          <button onClick={exportCsv} className={`flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg border ${t.border} ${t.textPrimary} ${t.hover} transition-colors`}>
            <FileText size={13} /> CSV
          </button>
          <button onClick={exportExcel} className="flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg bg-blue-600 text-white hover:bg-blue-700 transition-colors">
            <FileSpreadsheet size={13} /> Excel
          </button>
        </div>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className={`${t.surfaceAlt} border-b ${t.border}`}>
              {TABLE_COLUMNS.map((col) => (
                <th key={col.key} onClick={() => toggleSort(col.key)}
                  className={`text-left px-4 py-2.5 font-semibold text-xs uppercase tracking-wide ${t.textSecondary} whitespace-nowrap cursor-pointer select-none ${t.hover}`}>
                  <span className="inline-flex items-center gap-1">
                    {col.label}
                    {sort.key === col.key ? (sort.dir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />) : <ArrowUpDown size={11} className="opacity-40" />}
                  </span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {pageRows.map((r) => (
              <tr key={r.id} className={`border-b ${t.border} ${t.hover} transition-colors`}>
                <td className={`px-4 py-2.5 font-medium text-blue-600 whitespace-nowrap`}>{r.id}</td>
                <td className={`px-4 py-2.5 ${t.textSecondary} whitespace-nowrap`}>{r.dateStr}</td>
                <td className={`px-4 py-2.5 ${t.textPrimary} whitespace-nowrap`}>{r.customer}</td>
                <td className={`px-4 py-2.5 ${t.textPrimary} max-w-[220px] truncate`} title={r.product}>{r.product}</td>
                <td className={`px-4 py-2.5 ${t.textSecondary} whitespace-nowrap`}>{r.category}</td>
                <td className={`px-4 py-2.5 ${t.textPrimary}`}>{r.quantity}</td>
                <td className={`px-4 py-2.5 ${t.textSecondary} whitespace-nowrap`}>{fmtCurrency(r.unitPrice)}</td>
                <td className={`px-4 py-2.5 font-semibold ${t.textPrimary} whitespace-nowrap`}>{fmtCurrency(r.totalAmount)}</td>
                <td className={`px-4 py-2.5 ${t.textSecondary} whitespace-nowrap`}>{r.region}</td>
                <td className={`px-4 py-2.5 ${t.textSecondary} whitespace-nowrap`}>{r.salesRep}</td>
                <td className={`px-4 py-2.5 ${t.textSecondary} whitespace-nowrap`}>{r.paymentMethod}</td>
                <td className="px-4 py-2.5 whitespace-nowrap"><StatusBadge status={r.status} isDark={isDark} /></td>
              </tr>
            ))}
            {pageRows.length === 0 && (
              <tr><td colSpan={12} className={`text-center py-10 text-sm ${t.textSecondary}`}>No orders match your search or filters.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      <div className={`flex flex-col sm:flex-row items-center gap-3 px-4 py-3 border-t ${t.border}`}>
        <div className={`text-xs ${t.textSecondary}`}>
          Showing {sorted.length === 0 ? 0 : (currentPage - 1) * pageSize + 1}–{Math.min(currentPage * pageSize, sorted.length)} of {fmtNumber(sorted.length)}
        </div>
        <label className="flex items-center gap-1.5 sm:ml-4">
          <span className={`text-xs ${t.textSecondary}`}>Rows:</span>
          <select value={pageSize} onChange={(e) => { setPageSize(Number(e.target.value)); setPage(1); }}
            className={`text-xs rounded-md border ${t.border} ${t.inputBg} ${t.textPrimary} px-1.5 py-1 outline-none`}>
            {[10, 25, 50, 100].map((n) => <option key={n} value={n}>{n}</option>)}
          </select>
        </label>
        <div className="flex items-center gap-1 sm:ml-auto">
          <button disabled={currentPage === 1} onClick={() => setPage(1)}
            className={`p-1.5 rounded-md border ${t.border} disabled:opacity-30 ${t.hover} ${t.textSecondary}`}><ChevronsLeft size={14} /></button>
          <button disabled={currentPage === 1} onClick={() => setPage((p) => p - 1)}
            className={`p-1.5 rounded-md border ${t.border} disabled:opacity-30 ${t.hover} ${t.textSecondary}`}><ChevronLeft size={14} /></button>
          <span className={`text-xs px-2 ${t.textPrimary}`}>Page {currentPage} of {totalPages}</span>
          <button disabled={currentPage === totalPages} onClick={() => setPage((p) => p + 1)}
            className={`p-1.5 rounded-md border ${t.border} disabled:opacity-30 ${t.hover} ${t.textSecondary}`}><ChevronRight size={14} /></button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
   11. MAIN APP
   ========================================================================= */

export default function SalesDashboard() {
  const [isDark, setIsDark] = useState(false);
  const [collapsed, setCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [headerSearch, setHeaderSearch] = useState("");
  const [liveTick, setLiveTick] = useState(0);

  const t = useTheme(isDark);
  const allRows = useState(() => generateSalesData(5000))[0];

  const dateBounds = useMemo(() => {
    const dates = allRows.map((r) => r.dateStr);
    return { min: dates[dates.length - 1], max: dates[0] };
  }, [allRows]);

  const [filters, setFilters] = useState({
    dateFrom: "", dateTo: "", region: "All Regions", category: "All Categories",
    salesRep: "All Reps", customer: "All Customers", paymentMethod: "All Payment Methods",
  });

  useEffect(() => {
    setFilters((f) => ({ ...f, dateFrom: dateBounds.min, dateTo: dateBounds.max }));
  }, [dateBounds]);

  useEffect(() => {
    const interval = setInterval(() => setLiveTick((n) => n + 1), 15000);
    return () => clearInterval(interval);
  }, []);

  const filterOptions = useMemo(() => ({
    minDate: dateBounds.min, maxDate: dateBounds.max,
    regions: REGIONS.map((r) => r.name),
    categories: CATEGORIES,
    reps: SALES_REPS,
    customers: Array.from(new Set(allRows.map((r) => r.customer))).sort(),
    paymentMethods: PAYMENT_METHODS.map((p) => p.name),
  }), [allRows, dateBounds]);

  const filteredRows = useMemo(() => {
    if (!filters.dateFrom || !filters.dateTo) return allRows;
    return allRows.filter((r) => {
      if (r.dateStr < filters.dateFrom || r.dateStr > filters.dateTo) return false;
      if (filters.region !== "All Regions" && r.region !== filters.region) return false;
      if (filters.category !== "All Categories" && r.category !== filters.category) return false;
      if (filters.salesRep !== "All Reps" && r.salesRep !== filters.salesRep) return false;
      if (filters.customer !== "All Customers" && r.customer !== filters.customer) return false;
      if (filters.paymentMethod !== "All Payment Methods" && r.paymentMethod !== filters.paymentMethod) return false;
      return true;
    });
  }, [allRows, filters]);

  const tableRows = useMemo(() => {
    const q = headerSearch.trim().toLowerCase();
    if (!q) return filteredRows;
    return filteredRows.filter((r) =>
      r.id.toLowerCase().includes(q) || r.customer.toLowerCase().includes(q) || r.product.toLowerCase().includes(q)
    );
  }, [filteredRows, headerSearch]);

  const kpis = useMemo(() => computeKpis(filteredRows), [filteredRows]);
  const monthlyTrend = useMemo(() => computeMonthlyTrend(filteredRows), [filteredRows]);
  const categoryBreakdown = useMemo(() => computeCategoryBreakdown(filteredRows), [filteredRows]);
  const regionBreakdown = useMemo(() => computeRegionBreakdown(filteredRows), [filteredRows]);
  const topProducts = useMemo(() => computeTopProducts(filteredRows), [filteredRows]);
  const dailyPerf = useMemo(() => computeDailyPerformance(filteredRows), [filteredRows]);

  const liveLabel = liveTick === 0 ? "just now" : `${liveTick * 15}s ago`;

  return (
    <div className={`min-h-screen ${t.appBg} transition-colors duration-200 flex font-sans`}>
      <Sidebar t={t} collapsed={collapsed} mobileOpen={mobileOpen} setMobileOpen={setMobileOpen} />

      <div className="flex-1 min-w-0 flex flex-col">
        <Header
          t={t} isDark={isDark} setIsDark={setIsDark}
          search={headerSearch} setSearch={setHeaderSearch}
          setMobileOpen={setMobileOpen} collapsed={collapsed} setCollapsed={setCollapsed}
          liveLabel={liveLabel}
        />

        <main className="flex-1 p-4 lg:p-6 space-y-5 max-w-[1600px] w-full mx-auto">
          <div>
            <h1 className={`text-xl font-bold ${t.textPrimary} tracking-tight`}>Sales Overview</h1>
            <p className={`text-sm ${t.textSecondary} mt-0.5`}>Performance across all channels, {fmtDate(new Date(filters.dateFrom || dateBounds.min))} – {fmtDate(new Date(filters.dateTo || dateBounds.max))}</p>
          </div>

          <FilterBar t={t} filters={filters} setFilters={setFilters} options={filterOptions} resultCount={filteredRows.length} />

          {/* KPI cards */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-4">
            <KpiCard t={t} isDark={isDark} icon="units" label="Total Sales (units)" value={fmtNumber(kpis.totalUnits, true)} color="#2563EB" sparkData={kpis.sparkUnits} />
            <KpiCard t={t} isDark={isDark} icon="revenue" label="Total Revenue" value={fmtCurrency(kpis.totalRevenue, true)} color="#0EA5E9" sparkData={kpis.sparkRevenue} />
            <KpiCard t={t} isDark={isDark} icon="orders" label="Total Orders" value={fmtNumber(kpis.totalOrders, true)} color="#1D4ED8" sparkData={kpis.sparkOrders} />
            <KpiCard t={t} isDark={isDark} icon="customers" label="Total Customers" value={fmtNumber(kpis.totalCustomers, true)} color="#38BDF8" sparkData={kpis.sparkCustomers} />
            <KpiCard t={t} isDark={isDark} icon="aov" label="Avg. Order Value" value={fmtCurrency(kpis.avgOrderValue)} color="#7C3AED" sparkData={kpis.sparkAov} />
            <KpiCard t={t} isDark={isDark} icon="growth" label="Monthly Sales Growth" value={`${kpis.growth >= 0 ? "+" : ""}${kpis.growth.toFixed(1)}`} suffix="%" color={kpis.growth >= 0 ? "#059669" : "#DC2626"} delta={undefined} />
          </div>

          {/* Row: Monthly trend + donut */}
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
            <div className="lg:col-span-2">
              <ChartCard t={t} isDark={isDark} title="Monthly Sales Trend" subtitle="Revenue and order volume by month" height={300}>
                <LineChart data={monthlyTrend} margin={{ left: 4, right: 8, top: 4, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={isDark ? "#1e293b" : "#e2e8f0"} vertical={false} />
                  <XAxis dataKey="label" tick={{ fontSize: 11, fill: isDark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} />
                  <YAxis yAxisId="left" tick={{ fontSize: 11, fill: isDark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} tickFormatter={(v) => fmtCurrency(v, true)} />
                  <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11, fill: isDark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} tickFormatter={(v) => fmtNumber(v, true)} />
                  <Tooltip content={<CustomTooltip isDark={isDark} formatter={(v, name) => name === "revenue" ? fmtCurrency(v) : fmtNumber(v)} />} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Line yAxisId="left" type="monotone" dataKey="revenue" name="Revenue" stroke="#2563EB" strokeWidth={2.5} dot={false} activeDot={{ r: 5 }} />
                  <Line yAxisId="right" type="monotone" dataKey="orders" name="Orders" stroke="#38BDF8" strokeWidth={2} dot={false} activeDot={{ r: 5 }} strokeDasharray="4 3" />
                </LineChart>
              </ChartCard>
            </div>
            <ChartCard t={t} isDark={isDark} title="Sales Distribution" subtitle="Revenue share by category" height={300}>
              <PieChart>
                <Pie data={categoryBreakdown} dataKey="revenue" nameKey="name" innerRadius={62} outerRadius={95} paddingAngle={2}>
                  {categoryBreakdown.map((entry, i) => (
                    <Cell key={entry.name} fill={CHART_BLUES[i % CHART_BLUES.length]} />
                  ))}
                </Pie>
                <Tooltip content={<CustomTooltip isDark={isDark} formatter={(v) => fmtCurrency(v)} />} />
              </PieChart>
            </ChartCard>
          </div>

          {/* Row: category + region bars */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
            <ChartCard t={t} isDark={isDark} title="Sales by Product Category" subtitle="Revenue by category" height={300}>
              <BarChart data={categoryBreakdown} margin={{ left: 4, right: 8, top: 4, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={isDark ? "#1e293b" : "#e2e8f0"} vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 10, fill: isDark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} angle={-20} textAnchor="end" height={60} interval={0} />
                <YAxis tick={{ fontSize: 11, fill: isDark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} tickFormatter={(v) => fmtCurrency(v, true)} />
                <Tooltip content={<CustomTooltip isDark={isDark} formatter={(v) => fmtCurrency(v)} />} />
                <Bar dataKey="revenue" name="Revenue" radius={[6, 6, 0, 0]}>
                  {categoryBreakdown.map((entry, i) => <Cell key={entry.name} fill={CHART_BLUES[i % CHART_BLUES.length]} />)}
                </Bar>
              </BarChart>
            </ChartCard>
            <ChartCard t={t} isDark={isDark} title="Revenue by Region" subtitle="Geographic distribution" height={300}>
              <BarChart data={regionBreakdown} layout="vertical" margin={{ left: 8, right: 24, top: 4, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={isDark ? "#1e293b" : "#e2e8f0"} horizontal={false} />
                <XAxis type="number" tick={{ fontSize: 11, fill: isDark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} tickFormatter={(v) => fmtCurrency(v, true)} />
                <YAxis type="category" dataKey="name" tick={{ fontSize: 11, fill: isDark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} width={130} />
                <Tooltip content={<CustomTooltip isDark={isDark} formatter={(v) => fmtCurrency(v)} />} />
                <Bar dataKey="revenue" name="Revenue" radius={[0, 6, 6, 0]} fill="#2563EB" />
              </BarChart>
            </ChartCard>
          </div>

          {/* Row: top products */}
          <ChartCard t={t} isDark={isDark} title="Top 10 Best-Selling Products" subtitle="Ranked by revenue" height={360}>
            <BarChart data={topProducts} layout="vertical" margin={{ left: 8, right: 24, top: 4, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={isDark ? "#1e293b" : "#e2e8f0"} horizontal={false} />
              <XAxis type="number" tick={{ fontSize: 11, fill: isDark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} tickFormatter={(v) => fmtCurrency(v, true)} />
              <YAxis type="category" dataKey="name" tick={{ fontSize: 11, fill: isDark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} width={210} />
              <Tooltip content={<CustomTooltip isDark={isDark} formatter={(v, name) => name === "Revenue" ? fmtCurrency(v) : fmtNumber(v)} />} />
              <Bar dataKey="revenue" name="Revenue" radius={[0, 6, 6, 0]} fill="#1D4ED8" />
            </BarChart>
          </ChartCard>

          {/* Row: daily performance */}
          <ChartCard t={t} isDark={isDark} title="Daily Sales Performance" subtitle="Last 30 days in the selected range" height={280}>
            <AreaChart data={dailyPerf} margin={{ left: 4, right: 8, top: 4, bottom: 0 }}>
              <defs>
                <linearGradient id="dailyFill" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor="#2563EB" stopOpacity={0.35} />
                  <stop offset="100%" stopColor="#2563EB" stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke={isDark ? "#1e293b" : "#e2e8f0"} vertical={false} />
              <XAxis dataKey="label" tick={{ fontSize: 10, fill: isDark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: isDark ? "#94a3b8" : "#64748b" }} axisLine={false} tickLine={false} tickFormatter={(v) => fmtCurrency(v, true)} />
              <Tooltip content={<CustomTooltip isDark={isDark} formatter={(v) => fmtCurrency(v)} />} />
              <Area type="monotone" dataKey="revenue" name="Revenue" stroke="#2563EB" strokeWidth={2} fill="url(#dailyFill)" />
            </AreaChart>
          </ChartCard>

          {/* Data table */}
          <DataTable t={t} isDark={isDark} rows={tableRows} />

          <footer className={`text-center text-xs ${t.textMuted} pb-4`}>
            Sales Performance Dashboard &middot; Demo dataset of {fmtNumber(allRows.length)} synthetic orders &middot; For illustrative purposes only
          </footer>
        </main>
      </div>
    </div>
  );
}
