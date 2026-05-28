// ============================================================
// WinTek LIVE — corrected performance stats
// Close rate = (Sold + NG + Rescinded) / actual runs
// Money figures stay funded-only (Sold) so revenue never inflates
// ============================================================

// --- Status rule sets (match MASTER dropdown values exactly) ---

// A presentation actually ran (these form the close-rate denominator)
const RUN_STATUSES = [
  "Sold",
  "NG",
  "Rescinded",
  "Not Sold",
  "No Sale - Follow Up",
];

// Counts as a close (the numerator) — NG & Rescinded count as sales
const CLOSE_STATUSES = ["Sold", "NG", "Rescinded"];

// Only these pay — used for revenue, windows, PPU, commission
const FUNDED_STATUSES = ["Sold"];

// Never ran — excluded from run/close math entirely
// (Cancelled, No Show, Rescheduled, One Leg)

// --- Helpers ---
const norm = (s) => String(s || "").trim();
const money = (v) => Number(String(v).replace(/[^0-9.\-]/g, "")) || 0;
const num = (v) => Number(String(v).replace(/[^0-9.\-]/g, "")) || 0;

// `appointments` = array of MASTER rows.
// Adjust the field reads below if your row keys differ.
function computeStats(appointments) {
  const status = (r) => norm(r.status ?? r.Status ?? r["Status"]);

  const runs = appointments.filter((r) => RUN_STATUSES.includes(status(r)));
  const closes = appointments.filter((r) => CLOSE_STATUSES.includes(status(r)));
  const funded = appointments.filter((r) => FUNDED_STATUSES.includes(status(r)));

  const ngCount = appointments.filter((r) => status(r) === "NG").length;

  const revenue = funded.reduce(
    (sum, r) => sum + money(r.total_sale ?? r["Total Sale"]),
    0
  );
  const windows = funded.reduce(
    (sum, r) => sum + num(r.sold_windows ?? r["Sold Windows"]),
    0
  );
  const commission = funded.reduce(
    (sum, r) => sum + money(r.commission ?? r["Commission"]),
    0
  );

  const closeRate = runs.length ? (closes.length / runs.length) * 100 : 0;
  const avgPPU = windows ? revenue / windows : 0;

  return {
    incomeYTD: commission,         // $2,124.00   (funded commission only)
    salesClosed: closes.length,    // 2           (Sold + NG + Rescinded)
    ngCount,                       // 0
    apptsRun: runs.length,         // 9           (close-rate denominator)
    closeRate,                     // 22.2%
    totalRevenue: revenue,         // $17,700.00  (funded only)
    windowsSold: windows,          // 11
    avgPPU,                        // $1,609.09
  };
}

// ---- Sanity check against current data (16 dated rows) ----
// Runs (9):    1 Sold, 4 Not Sold, 3 No Sale-Follow Up, 1 Rescinded
// Closes (2):  Carney (Sold) + Hollingshead (Rescinded); NG = 0
// Excluded (7): 4 Cancelled, 1 No Show, 1 Rescheduled, 1 One Leg
// => Close rate = 2 / 9 = 22.2%   (was showing 1 / 16 = 6.3%)
