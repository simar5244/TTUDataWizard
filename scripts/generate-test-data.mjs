/**
 * DataWizard Test Data Generator
 * Produces two heavy Excel files for end-to-end testing:
 *   test-data/salesforce_export.xlsx   — 800 rows, Salesforce-style CRM export
 *   test-data/target_crm.xlsx          — 600 rows, "destination" CRM sheet
 *   test-data/financials_q1.xlsx       — 1000 rows, financial data (for charts)
 */

import * as XLSX from "xlsx";
import { mkdir, writeFile } from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT_DIR = path.join(__dirname, "..", "test-data");

// ── Helpers ─────────────────────────────────────────────────────────────────

const rnd = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
const pick = (arr) => arr[rnd(0, arr.length - 1)];
const money = (min, max) => Math.round(Math.random() * (max - min) + min);
const fmt = (n, d = 2) => parseFloat(n.toFixed(d));

function randDate(start, end) {
  const s = new Date(start).getTime();
  const e = new Date(end).getTime();
  return new Date(s + Math.random() * (e - s)).toISOString().slice(0, 10);
}

const FIRST_NAMES = [
  "James","Mary","John","Patricia","Robert","Jennifer","Michael","Linda",
  "William","Barbara","David","Elizabeth","Richard","Susan","Joseph","Jessica",
  "Thomas","Sarah","Charles","Karen","Christopher","Lisa","Daniel","Nancy",
  "Matthew","Betty","Anthony","Margaret","Mark","Sandra","Donald","Ashley",
  "Steven","Dorothy","Paul","Kimberly","Andrew","Emily","Kenneth","Donna",
  "Kevin","Michelle","Brian","Carol","George","Amanda","Timothy","Melissa",
  "Ronald","Deborah","Edward","Stephanie","Jason","Rebecca","Jeffrey","Sharon",
  "Ryan","Laura","Jacob","Cynthia","Gary","Kathleen","Nicholas","Amy","Eric",
  "Angela","Jonathan","Shirley","Stephen","Anna","Larry","Brenda","Justin",
  "Pamela","Scott","Emma","Brandon","Nicole","Benjamin","Helen","Samuel",
  "Samantha","Raymond","Katherine","Gregory","Christine","Frank","Debra",
  "Alexander","Rachel","Patrick","Carolyn","Jack","Janet","Dennis","Maria",
];

const LAST_NAMES = [
  "Smith","Johnson","Williams","Brown","Jones","Garcia","Miller","Davis",
  "Rodriguez","Martinez","Hernandez","Lopez","Gonzalez","Wilson","Anderson",
  "Thomas","Taylor","Moore","Jackson","Martin","Lee","Perez","Thompson","White",
  "Harris","Sanchez","Clark","Ramirez","Lewis","Robinson","Walker","Young",
  "Allen","King","Wright","Scott","Torres","Nguyen","Hill","Flores","Green",
  "Adams","Nelson","Baker","Hall","Rivera","Campbell","Mitchell","Carter",
  "Roberts","Gomez","Phillips","Evans","Turner","Diaz","Parker","Cruz",
  "Edwards","Collins","Reyes","Stewart","Morris","Morales","Murphy","Cook",
  "Rogers","Gutierrez","Ortiz","Morgan","Cooper","Peterson","Bailey","Reed",
  "Kelly","Howard","Ramos","Kim","Cox","Ward","Richardson","Watson","Brooks",
  "Chavez","Wood","James","Bennett","Gray","Mendoza","Ruiz","Hughes","Price",
  "Alvarez","Castillo","Sanders","Patel","Myers","Long","Ross","Foster",
];

const COMPANIES = [
  "Acme Corp","Globex","Initech","Umbrella Corp","Stark Industries",
  "Wayne Enterprises","Oscorp","Massive Dynamic","Soylent Corp","Weyland-Yutani",
  "Cyberdyne Systems","Nakatomi Trading","Tyrell Corporation","Multi National United",
  "Rekall Inc","BiffCo Enterprises","Momcorp","Planet Express","Virtucon",
  "Vandelay Industries","Prestige Worldwide","Dunder Mifflin","Bluth Company",
  "Wonka Industries","Hammerdown Industrial","Northside Solutions","Pacific Rim Analytics",
  "Apex Digital","BlueSky Systems","RedWood Technologies","SilverPeak Group",
  "IronBridge Corp","GoldRush Mining","SteelWave Manufacturing","CopperHead Finance",
  "NeonPath Media","SolarFlare Energy","AquaVault Finance","TerraNova Agriculture",
  "CloudBridge SaaS","FrostByte Computing","VaporWave Studios","NightOwl Security",
  "DawnBreaker Logistics","SunrisePharma","MidnightOil Consulting","ThunderRidge Mining",
  "WhiteLabel Brands","BlackRock Analytics","PurpleHaze Creative","OrangeGrove Retail",
];

const INDUSTRIES = [
  "Technology","Finance","Healthcare","Manufacturing","Retail","Education",
  "Real Estate","Energy","Media","Consulting","Logistics","Agriculture",
  "Pharmaceuticals","Insurance","Automotive","Aerospace","Construction","Hospitality",
];

const STAGES = [
  "Prospecting","Qualification","Needs Analysis","Value Proposition",
  "Id. Decision Makers","Perception Analysis","Proposal/Price Quote",
  "Negotiation/Review","Closed Won","Closed Lost",
];

const COUNTRIES = [
  "United States","Canada","United Kingdom","Germany","France","Australia",
  "Japan","Singapore","Netherlands","Sweden","Switzerland","Brazil","India",
  "South Korea","Mexico","Spain","Italy","New Zealand","Norway","Denmark",
];

const TITLES = [
  "CEO","CTO","CFO","VP of Sales","VP of Marketing","Director of Operations",
  "Sales Manager","Account Executive","Business Development Rep","Head of IT",
  "Chief Revenue Officer","Product Manager","Engineering Manager","Finance Director",
  "Procurement Manager","Operations Lead","Marketing Manager","Data Analyst",
];

const STATES = [
  "CA","NY","TX","FL","IL","PA","OH","GA","NC","MI",
  "NJ","VA","WA","AZ","MA","TN","IN","MO","MD","WI",
];

const SOURCES = [
  "Web","Referral","Event","Cold Call","Email Campaign",
  "LinkedIn","Partner","Trial","Inbound","Webinar",
];

// ── Sheet 1: Salesforce Export ───────────────────────────────────────────────

function makeSalesforceRow(i) {
  const first = pick(FIRST_NAMES);
  const last  = pick(LAST_NAMES);
  const company = pick(COMPANIES);
  const domain  = company.toLowerCase().replace(/[^a-z]/g, "") + ".com";

  return {
    "Contact ID":         `SF-${String(i).padStart(6, "0")}`,
    "First Name":         first,
    "Last Name":          last,
    "Full Name":          `${first} ${last}`,
    "Email":              `${first.toLowerCase()}.${last.toLowerCase()}@${domain}`,
    "Phone":              `(${rnd(200,999)}) ${rnd(200,999)}-${rnd(1000,9999)}`,
    "Mobile":             `(${rnd(200,999)}) ${rnd(200,999)}-${rnd(1000,9999)}`,
    "Job Title":          pick(TITLES),
    "Department":         pick(["Sales","Marketing","Engineering","Finance","Operations","HR","IT","Legal"]),
    "Account Name":       company,
    "Account ID":         `ACC-${String(rnd(10000,99999)).padStart(6,"0")}`,
    "Industry":           pick(INDUSTRIES),
    "Annual Revenue":     money(500_000, 50_000_000),
    "Number of Employees": rnd(10, 50000),
    "Country":            pick(COUNTRIES),
    "State":              pick(STATES),
    "City":               pick(["New York","Los Angeles","Chicago","Houston","Phoenix","Philadelphia","San Antonio","San Diego","Dallas","San Jose","Austin","Jacksonville","Fort Worth","Columbus","Charlotte"]),
    "Postal Code":        String(rnd(10000, 99999)),
    "Lead Source":        pick(SOURCES),
    "Lead Status":        pick(["New","Working","Nurturing","Qualified","Unqualified","Converted"]),
    "Opportunity Name":   `${company} - ${pick(["Q1","Q2","Q3","Q4"])} ${rnd(2024,2026)} Deal`,
    "Opportunity ID":     `OPP-${String(i).padStart(6,"0")}`,
    "Stage":              pick(STAGES),
    "Amount":             money(5_000, 2_000_000),
    "Probability %":      rnd(5, 95),
    "Expected Revenue":   money(1_000, 1_000_000),
    "Close Date":         randDate("2024-01-01", "2026-12-31"),
    "Type":               pick(["New Business","Existing Business","Renewal","Upsell","Cross-sell"]),
    "Next Step":          pick(["Schedule demo","Send proposal","Follow up call","Contract review","Awaiting signature","None"]),
    "Description":        `${pick(["Priority","High-value","Strategic","Inbound","Warm"])} opportunity via ${pick(SOURCES)}.`,
    "Created Date":       randDate("2022-01-01", "2025-06-01"),
    "Last Modified Date": randDate("2025-01-01", "2026-04-01"),
    "Last Activity Date": randDate("2025-06-01", "2026-04-01"),
    "Owner":              `${pick(FIRST_NAMES)} ${pick(LAST_NAMES)}`,
    "Owner Email":        `owner${rnd(1,99)}@salesteam.com`,
    "Campaign":           pick(["Q1 Outbound","Spring Launch","Partner Program","Webinar Series","","Cold Outreach","ABM Wave 1"]),
    "Rating":             pick(["Hot","Warm","Cold",""]),
    "NPS Score":          rnd(0, 10),
    "Contract Value":     money(10_000, 5_000_000),
    "Discount %":         fmt(Math.random() * 30),
    "ARR":                money(5_000, 500_000),
    "MRR":                money(500, 50_000),
  };
}

// ── Sheet 2: Target CRM ───────────────────────────────────────────────────────

function makeTargetRow(i) {
  const first = pick(FIRST_NAMES);
  const last  = pick(LAST_NAMES);
  const company = pick(COMPANIES);
  const domain  = company.toLowerCase().replace(/[^a-z]/g, "") + ".io";

  return {
    "Record ID":          `CRM-${String(i).padStart(5,"0")}`,
    "Name":               `${first} ${last}`,
    "Primary Email":      `${first.toLowerCase()}_${last.toLowerCase()}@${domain}`,
    "Secondary Email":    `${first.toLowerCase()}@gmail.com`,
    "Direct Phone":       `+1-${rnd(200,999)}-${rnd(200,999)}-${rnd(1000,9999)}`,
    "Role":               pick(TITLES),
    "Organization":       company,
    "Org Revenue":        money(1_000_000, 100_000_000),
    "Headcount":          rnd(5, 100000),
    "Vertical":           pick(INDUSTRIES),
    "Geo":                pick(COUNTRIES),
    "Region":             pick(["AMER","EMEA","APAC","LATAM"]),
    "Deal Name":          `${company} ${rnd(2024,2026)} Initiative`,
    "Deal Value":         money(10_000, 3_000_000),
    "Recurring Revenue":  money(1_000, 500_000),
    "Pipeline Stage":     pick(["Lead","Prospect","Demo Scheduled","Proposal Sent","Negotiating","Won","Lost","On Hold"]),
    "Win Probability":    rnd(0, 100),
    "Expected Close":     randDate("2025-01-01", "2027-06-30"),
    "Assigned To":        `${pick(FIRST_NAMES)} ${pick(LAST_NAMES)}`,
    "Team":               pick(["East","West","Enterprise","SMB","Partner","EMEA","APAC"]),
    "Source Channel":     pick(SOURCES),
    "Date Added":         randDate("2023-01-01", "2025-12-31"),
    "Last Updated":       randDate("2025-01-01", "2026-04-01"),
    "Tags":               pick(["priority","follow-up","champion","executive-sponsor","budget-confirmed","no-budget","competitor","renewal","","upsell"]),
    "Health Score":       rnd(1, 100),
    "NPS":                rnd(-100, 100),
    "Engagement Score":   fmt(Math.random() * 10),
    "Last Email Open":    randDate("2025-06-01", "2026-04-01"),
    "Meetings Held":      rnd(0, 20),
    "Notes":              pick(["Strong interest","Budget approved","Decision next quarter","Evaluating competitors","Champion identified","On hold","Needs executive alignment",""]),
  };
}

// ── Sheet 3: Financials Q1 ────────────────────────────────────────────────────

const PRODUCTS = [
  "DataWizard Pro","DataWizard Starter","DataWizard Enterprise","API Add-on",
  "Storage Expansion","SSO Integration","Priority Support","Custom Connector",
  "Training Package","Professional Services","White-label License","Data Audit",
];

const COST_CENTERS = ["Engineering","Sales","Marketing","Operations","HR","Legal","Finance","IT"];

function makeFinancialRow(i) {
  const month = rnd(1, 12);
  const year  = pick([2024, 2025, 2026]);
  const qty   = rnd(1, 500);
  const unit  = money(50, 5000);
  const revenue = qty * unit;
  const cogs    = fmt(revenue * (0.2 + Math.random() * 0.4));
  const gross   = fmt(revenue - cogs);
  const opex    = money(1000, 50000);
  const ebitda  = fmt(gross - opex);

  return {
    "Transaction ID":     `TXN-${String(i).padStart(7,"0")}`,
    "Date":               `${year}-${String(month).padStart(2,"0")}-${String(rnd(1,28)).padStart(2,"0")}`,
    "Year":               year,
    "Quarter":            `Q${Math.ceil(month / 3)}`,
    "Month":              month,
    "Month Name":         ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][month-1],
    "Product":            pick(PRODUCTS),
    "Category":           pick(["SaaS","Professional Services","Support","License","Infrastructure"]),
    "Cost Center":        pick(COST_CENTERS),
    "Sales Rep":          `${pick(FIRST_NAMES)} ${pick(LAST_NAMES)}`,
    "Region":             pick(["AMER","EMEA","APAC","LATAM"]),
    "Country":            pick(COUNTRIES),
    "Customer Segment":   pick(["Enterprise","Mid-Market","SMB","Startup"]),
    "Units Sold":         qty,
    "Unit Price":         unit,
    "Revenue":            revenue,
    "COGS":               cogs,
    "Gross Profit":       gross,
    "Gross Margin %":     fmt((gross / revenue) * 100),
    "OpEx":               opex,
    "EBITDA":             ebitda,
    "EBITDA Margin %":    fmt((ebitda / revenue) * 100),
    "Net Revenue":        fmt(revenue * (0.85 + Math.random() * 0.1)),
    "Discount Given":     money(0, Math.floor(revenue * 0.2)),
    "Tax":                fmt(revenue * 0.08),
    "Invoice Status":     pick(["Paid","Pending","Overdue","Draft","Cancelled"]),
    "Payment Method":     pick(["Credit Card","Wire Transfer","ACH","Check","Net 30","Net 60"]),
    "Days to Payment":    rnd(0, 90),
    "Churn Risk":         pick(["Low","Medium","High","Critical"]),
    "LTV":                money(10_000, 2_000_000),
    "CAC":                money(500, 20_000),
    "LTV:CAC Ratio":      fmt(rnd(1, 15) + Math.random()),
    "MRR":                money(500, 100_000),
    "ARR":                money(6_000, 1_200_000),
    "Expansion Revenue":  money(0, 50_000),
    "Churn Revenue":      money(0, 20_000),
    "Net New MRR":        money(-5000, 50_000),
    "Headcount (Dept)":   rnd(2, 200),
    "Budget Allocated":   money(50_000, 5_000_000),
    "Budget Spent":       money(20_000, 4_500_000),
    "Budget Variance %":  fmt((Math.random() * 40) - 20),
    "Forecast":           money(100_000, 10_000_000),
    "Actual vs Forecast %": fmt((Math.random() * 40) - 20),
  };
}

// ── Build & Write workbooks ──────────────────────────────────────────────────

async function buildWorkbook(name, rowFn, count) {
  console.log(`  Building ${name} (${count} rows)…`);
  const rows = Array.from({ length: count }, (_, i) => rowFn(i + 1));
  const ws = XLSX.utils.json_to_sheet(rows);

  // Auto column widths
  const cols = Object.keys(rows[0]);
  ws["!cols"] = cols.map((k) => {
    const max = Math.max(k.length, ...rows.slice(0, 50).map((r) => String(r[k] ?? "").length));
    return { wch: Math.min(max + 2, 40) };
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, name.replace(/[^a-zA-Z0-9 ]/g, "").slice(0, 31));

  const filePath = path.join(OUT_DIR, `${name}.xlsx`);
  XLSX.writeFile(wb, filePath);
  console.log(`  ✓ ${filePath}`);
  return filePath;
}

// ── Extra: multi-sheet workbook for dashboard testing ─────────────────────────

async function buildMultiSheet() {
  const name = "dashboard_test_data.xlsx";
  console.log(`  Building ${name} (multi-sheet)…`);
  const wb = XLSX.utils.book_new();

  // Monthly Revenue summary
  const monthly = [];
  const years = [2024, 2025, 2026];
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  for (const year of years) {
    for (const [mi, mon] of months.entries()) {
      const rev = money(800_000, 4_000_000);
      monthly.push({
        Period: `${mon} ${year}`,
        Year: year,
        Month: mon,
        "Month #": mi + 1,
        Revenue: rev,
        "Gross Profit": fmt(rev * (0.5 + Math.random() * 0.25)),
        EBITDA: fmt(rev * (0.1 + Math.random() * 0.25)),
        MRR: money(50_000, 400_000),
        "New Customers": rnd(10, 200),
        "Churned Customers": rnd(0, 30),
        "Net New Customers": rnd(5, 170),
        "Active Seats": rnd(500, 10_000),
        "Support Tickets": rnd(20, 500),
        "NPS": rnd(20, 80),
      });
    }
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(monthly), "Monthly Revenue");

  // By Region
  const regions = ["AMER","EMEA","APAC","LATAM"];
  const byRegion = regions.flatMap((r) =>
    Array.from({ length: 24 }, (_, i) => ({
      Region: r,
      Quarter: `Q${(i % 4) + 1} ${2024 + Math.floor(i / 8)}`,
      Revenue: money(200_000, 2_000_000),
      Deals: rnd(5, 100),
      "Win Rate %": rnd(20, 75),
      "Avg Deal Size": money(10_000, 200_000),
      "Pipeline": money(500_000, 10_000_000),
      "Headcount": rnd(10, 500),
    }))
  );
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(byRegion), "By Region");

  // Product Mix
  const byProduct = PRODUCTS.flatMap((p) =>
    months.map((m) => ({
      Product: p,
      Month: m,
      Units: rnd(10, 2000),
      Revenue: money(5_000, 500_000),
      "Avg Price": money(200, 5_000),
      "Gross Margin %": rnd(40, 90),
      "Customer Satisfaction": fmt(3 + Math.random() * 2),
    }))
  );
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(byProduct), "Product Mix");

  // Customer Cohorts
  const cohorts = [];
  for (let y = 2022; y <= 2025; y++) {
    for (let q = 1; q <= 4; q++) {
      cohorts.push({
        "Cohort": `Q${q} ${y}`,
        "Customers Acquired": rnd(20, 300),
        "M1 Retention %": rnd(70, 99),
        "M3 Retention %": rnd(60, 95),
        "M6 Retention %": rnd(50, 90),
        "M12 Retention %": rnd(40, 85),
        "M24 Retention %": rnd(30, 80),
        "LTV": money(5_000, 100_000),
        "CAC": money(500, 10_000),
        "Payback Months": rnd(3, 36),
        "Expansion Revenue %": fmt(Math.random() * 40),
      });
    }
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cohorts), "Cohort Analysis");

  const filePath = path.join(OUT_DIR, name);
  XLSX.writeFile(wb, filePath);
  console.log(`  ✓ ${filePath}`);
}

// ── Main ─────────────────────────────────────────────────────────────────────

async function main() {
  await mkdir(OUT_DIR, { recursive: true });
  console.log("\n🧙 DataWizard Test Data Generator\n");

  await buildWorkbook("salesforce_export", makeSalesforceRow, 800);
  await buildWorkbook("target_crm", makeTargetRow, 600);
  await buildWorkbook("financials_q1", makeFinancialRow, 1000);
  await buildMultiSheet();

  console.log(`\n✅ Done! Files written to: test-data/\n`);
  console.log("Files:");
  console.log("  salesforce_export.xlsx  — 800 rows × 38 cols (Salesforce-style CRM)");
  console.log("  target_crm.xlsx         — 600 rows × 30 cols (target CRM sheet)");
  console.log("  financials_q1.xlsx      — 1000 rows × 39 cols (P&L / revenue data)");
  console.log("  dashboard_test_data.xlsx — 4 sheets: Monthly/Region/Product/Cohort\n");
}

main().catch(console.error);
