# Excel Sales Analytics Dashboard

> Production-grade sales performance analysis system built entirely in Excel

![Dashboard Preview](screenshots/05_dashboard.png)

## 🎯 Project Overview

An end-to-end analytics dashboard demonstrating professional Excel capabilities:
- Data quality auditing and cleaning
- Multi-sheet architecture with cross-references
- 40+ formulas including advanced lookups and conditional aggregation
- Auto-updating charts and KPI cards
- Production-ready design with audit trails

**Skills Demonstrated:** Data cleaning • Formula engineering • Conditional formatting • Dashboard design • Cross-sheet referencing • INDEX/MATCH • SUMIFS • Error handling

---

## 📊 File Structure

The workbook contains 5 interconnected sheets:

### 1. Raw_Data (🔴 Red Tab)
- **Purpose:** Original source data with intentional quality issues
- **Never edited** — preserved as audit trail
- **Contains:** 20 transactions across 4 regions, 6 products, 6 sales reps

![Raw Data](screenshots/01_raw_data.png)

**Data quality issues present:**
- ❌ Missing values (Units Sold, Sales Rep)
- ❌ Negative prices
- ❌ Text inconsistencies (case sensitivity)
- ❌ Trailing spaces
- ❌ Wrong calculations
- ❌ Duplicate entries

---

### 2. Cleaned_Data (🟢 Green Tab)
- **Purpose:** Production-ready dataset with all errors corrected
- **Key formulas:** `=IF(OR(F2="";G2="");"";F2*G2)` for dynamic Total Sales
- **Flagging system:** Missing values documented in Notes column

![Cleaned Data](screenshots/02_cleaned_data.png)

**Cleaning techniques applied:**
- TRIM() for space removal
- Find & Replace for format standardization
- IF/OR wrappers for error prevention
- Regional format handling (comma vs period decimals)

---

### 3. KPI_Engine (🔵 Blue Tab)
- **Purpose:** Centralized calculation hub
- **40+ formulas** across 4 analysis sections
- **All metrics** pull from Cleaned_Data dynamically

![KPI Engine](screenshots/03_kpi_engine.png)

**Sections:**

**A. Overall Business Metrics**
- Total Revenue: `=SUM(Cleaned_Data!H:H)`
- Avg Transaction Value: `=AVERAGE(Cleaned_Data!H:H)`
- Total Units: `=SUM(Cleaned_Data!F:F)`
- Avg Customer Rating: `=AVERAGE(Cleaned_Data!I:I)`

**B. Regional Performance**
- Revenue by Region: `=SUMIF(Cleaned_Data!D:D;A13;Cleaned_Data!H:H)`
- Units by Region: `=SUMIF(Cleaned_Data!D:D;A13;Cleaned_Data!F:F)`
- Transactions: `=COUNTIF(Cleaned_Data!D:D;A13)`
- Best/Worst Region: INDEX/MATCH on MAX/MIN revenue

**C. Product Performance**
- Revenue by Product: SUMIF across 6 products
- Error handling: `=IFERROR(B25/C25;"N/A")` for divisions

**D. Sales Rep Performance**
- Revenue by Rep with rankings
- Conditional formatting highlights top/bottom performers

---

### 4. Pivot_Analysis (🟠 Orange Tab)
- **Purpose:** Manual cross-tabulation (Region × Product revenue matrix)
- **Core formula:** `=SUMIFS(Cleaned_Data!H:H;Cleaned_Data!D:D;$A4;Cleaned_Data!E:E;B$3)`
- **Mixed references** allow one formula to populate entire 4×6 grid

![Pivot Analysis](screenshots/04_pivot_analysis.png)

**Why manual instead of Pivot Table?**
- Demonstrates understanding of underlying logic
- Full control over layout and formatting
- Shows formula proficiency (SUMIFS with multiple criteria)

---

### 5. Dashboard (🟣 Purple Tab)
- **Purpose:** Executive summary with visual KPIs
- **Auto-updating charts** linked to KPI_Engine
- **Interactive elements** (if slicers added)

![Dashboard](screenshots/05_dashboard.png)

**Components:**
- KPI cards (Total Revenue, Best Region, Top Product, Top Rep)
- 3 charts: Revenue by Region, Product Performance, Rep Leaderboard
- Conditional formatting for at-a-glance insights

---

## 🔗 Data Flow Architecture

┌─────────────┐
│  Raw_Data   │  (Never edited - audit trail)
└──────┬──────┘
       │
       ↓
┌─────────────┐
│Cleaned_Data │  (Formulas + Flags)
└──────┬──────┘
       │
       ↓
┌─────────────┐
│ KPI_Engine  │  (40+ calculations)
└──────┬──────┘
       │
       ├────→ Pivot_Analysis (Cross-tabs)
       │
       └────→ Dashboard (Visual summary)

---

## 🧠 Key Technical Concepts

### 1. Cross-Sheet Referencing
All formulas use explicit sheet references:
```excel
=SUMIF(Cleaned_Data!D:D;A13;Cleaned_Data!H:H)
```
Changes to source data propagate automatically.

### 2. Mixed Cell References
```excel
=SUMIFS(Cleaned_Data!H:H;Cleaned_Data!D:D;$A4;Cleaned_Data!E:E;B$3)
```
- `$A4` locks column, row changes when copied down
- `B$3` locks row, column changes when copied right
- One formula works for entire matrix (24 cells)

### 3. INDEX/MATCH Lookups
```excel
=INDEX(A13:A16;MATCH(MAX(B13:B16);B13:B16;0))
```
Finds best-performing region dynamically. Preferred over VLOOKUP for:
- Bidirectional lookup capability
- Column-number independence
- Better performance at scale

### 4. Error Handling
```excel
=IFERROR(B25/C25;"N/A")
```
Prevents #DIV/0! errors when denominators are zero.

### 5. Regional Excel Format Handling
- German Excel: `,` = decimal, `;` = function separator
- US Excel: `.` = decimal, `,` = function separator
- All formulas adjusted accordingly

---

## 📈 Charts & Visualizations

### Chart 1: Revenue by Region (Column Chart)
- **Data source:** KPI_Engine!A13:B16
- **Auto-updates** when regional revenue changes

### Chart 2: Product Performance (Horizontal Bar)
- **Data source:** KPI_Engine!A25:B30
- **Sorted** by revenue descending

### Chart 3: Sales Rep Leaderboard (Column Chart)
- **Data source:** KPI_Engine!A35:B40
- **Conditional coloring** (top = green, bottom = red)

**Advanced features ready to add:**
- Combo charts (Revenue + Rating)
- Sparklines (in-cell trends)
- Slicers (interactive filtering)

---

## 🎯 Business Insights Delivered

**Regional Analysis:**
- **South** leads in revenue (€52,680) with highest volume (92 units)
- **West** has highest avg sale price (€667) indicating premium product mix
- **North** underperforms — potential for growth or resource reallocation

**Product Analysis:**
- **Laptop** dominates (47% of revenue)
- **Keyboard** high volume (70 units) but low revenue (budget segment)
- **Headphones** zero sales — discontinue or investigate why

**Rep Performance:**
- **Bob** top performer (€38,856 revenue)
- **Unknown** rep with €19,999 in sales — data quality issue to investigate
- Performance spread suggests coaching opportunity for bottom tier

---

## 🛠️ Skills Demonstrated

| Category | Skills |
|----------|--------|
| **Data Quality** | Error detection, systematic auditing, flagging vs filling missing data |
| **Formula Logic** | SUM, AVERAGE, SUMIF, SUMIFS, COUNTIF, AVERAGEIF, INDEX/MATCH, IFERROR |
| **Architecture** | Multi-sheet design, cross-references, modular calculations |
| **Visualization** | Charts, conditional formatting, dashboard design |
| **Production Practices** | Audit trails, documentation, maintaiability, error handling |
| **Technical** | Mixed cell references, dynamic ranges, regional settings |

---

## 💼 Interview Talking Points

### "How do you handle missing data?"
*"I follow the principle: flag, don't fill. Missing values are documented in a notes column while cells remain blank. Formulas use IF/OR wrappers to prevent errors. This preserves data integrity and lets stakeholders decide whether to track down real values or exclude from analysis. Inventing data creates untraceable downstream errors."*

### "Why use formulas instead of hardcoded values?"
*"Formulas create a living dashboard. When source data changes, everything recalculates automatically. Hardcoded values become stale and require manual updates. This is the difference between a one-time report and a production system."*

### "Explain your use of INDEX/MATCH."
*"INDEX/MATCH is more flexible than VLOOKUP — it works bidirectionally, doesn't break when columns are inserted, and performs better at scale. For example, to find the best-performing region, MATCH identifies the position of the maximum revenue, and INDEX retrieves the corresponding region name."*

---

## 🚀 Future Enhancements

**Next features to add:**
- [ ] Power Query for automated data refresh
- [ ] Slicers for interactive filtering
- [ ] VBA macro for one-click report generation
- [ ] Dynamic arrays (FILTER, SORT, UNIQUE)
- [ ] Waterfall chart showing revenue contribution
- [ ] Combo charts (Revenue vs Customer Satisfaction)
- [ ] Sparklines for trend visualization

**Scalability path:**
- Export to CSV → Python pandas
- Convert formulas → SQL queries
- Dashboard → Power BI / Tableau
- Automation → Scheduled pipelines

---

## 📚 Learning Resources

**Concepts mastered in this project:**
- Data cleaning workflows
- Cross-sheet formula architecture
- Conditional aggregation (SUMIF/SUMIFS/COUNTIF)
- Advanced lookups (INDEX/MATCH)
- Error handling strategies
- Regional Excel format differences
- Dashboard design principles

---


🎨 DATA FLOW DIAGRAM


┌──────────────────────────────────────────────────────────┐
│                     DATA PIPELINE                         │
└──────────────────────────────────────────────────────────┘

┌─────────────────┐
│   RAW_DATA      │
│   (Source)      │
│  - 20 rows      │
│  - 10 columns   │
│  - Errors 🔴    │
└────────┬────────┘
         │
         │ Manual cleaning
         │ + Formulas
         ↓
┌─────────────────┐
│ CLEANED_DATA    │
│  - Trimmed      │
│  - Standardized │
│  - Flagged ⚠️   │
│  - Formula H col│
└────────┬────────┘
         │
         │ Cross-sheet
         │ references
         ↓
┌─────────────────┐
│  KPI_ENGINE     │
│  40+ formulas   │
│  ┌───────────┐  │
│  │ Overall   │  │
│  │ Regional  │  │
│  │ Product   │  │
│  │ Rep       │  │
│  └───────────┘  │
└────┬──────┬─────┘
     │      │
     │      └─────────────┐
     ↓                    ↓
┌────────────┐    ┌──────────────┐
│PIVOT_ANALYSIS│    │  DASHBOARD   │
│ SUMIFS matrix│    │ Charts + KPIs│
│ 4×6 grid     │    │ Visual summary│
└────────────┘    └──────────────┘
