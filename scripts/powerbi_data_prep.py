"""
=============================================================
POWER BI DATA PREP — Procore Submittal & RFI Transformer
=============================================================
Tailored for API CPMC Project — Bird Construction
Matches exact Procore export column headers.

Your Procore Submittal columns:
  Project Name, Title, Type, Ball in Court,
  Ball In Court Due Date, Final Due Date, Status,
  Overdue, Days Overdue

Your Procore RFI columns:
  Rfi #, Subject, Ball in Court, Date Created,
  Due Date, Status, Overdue, Due Date Variance

Usage:
  1. Place your Procore exports in data/raw/
  2. Run: python powerbi_data_prep.py
  3. Import data/processed/PowerBI_Procore_Data.xlsx into Power BI

pip install pandas openpyxl
=============================================================
"""

import pandas as pd
from datetime import datetime, timedelta
import os

# ============================================================
# CONFIGURATION
# ============================================================
SUBMITTAL_FILE = "../data/raw/Open Submittals - Final (2).xlsx"
RFI_FILE = "../data/raw/Open RFIs - Final (2).xlsx"
OUTPUT_FILE = "../data/processed/PowerBI_Procore_Data.xlsx"

SUBMITTAL_OVERDUE_DAYS = 14
RFI_OVERDUE_DAYS = 10

TODAY = datetime.now()


# ============================================================
# COLUMN MAPPING — Your exact Procore headers → Standard names
# ============================================================
SUBMITTAL_COL_MAP = {
    # --- YOUR ACTUAL PROCORE SUBMITTAL COLUMNS ---
    "Project Name": "Project_Name",
    "Title": "Title",
    "Type": "Submittal_Type",
    "Ball in Court": "Ball_in_Court",
    "Ball In Court": "Ball_in_Court",
    "Ball In Court Due Date": "Due_Date",
    "Ball in Court Due Date": "Due_Date",
    "Final Due Date": "Final_Due_Date",
    "Status": "Status",
    "Overdue": "Overdue_Flag",
    "Days Overdue": "Days_Overdue",
    # --- ADDITIONAL PROCORE VARIATIONS (future-proofing) ---
    "Number": "Submittal_ID",
    "#": "Submittal_ID",
    "Submittal Number": "Submittal_ID",
    "Submittal No": "Submittal_ID",
    "Submittal No.": "Submittal_ID",
    "No.": "Submittal_ID",
    "Subject": "Title",
    "Description": "Title",
    "Spec Section": "Spec_Section",
    "Specification Section": "Spec_Section",
    "Responsible Contractor": "Contractor",
    "Subcontractor": "Contractor",
    "Company": "Contractor",
    "Received From": "Contractor",
    "Submitted On": "Date_Created",
    "Created Date": "Date_Created",
    "Date Submitted": "Date_Created",
    "Due Date": "Due_Date",
    "Required Date": "Due_Date",
    "Date Returned": "Date_Closed",
    "Closed Date": "Date_Closed",
    "Reviewer": "Reviewer",
    "Approver": "Reviewer",
}

RFI_COL_MAP = {
    # --- YOUR ACTUAL PROCORE RFI COLUMNS ---
    "Rfi #": "RFI_ID",
    "RFI #": "RFI_ID",
    "Rfi#": "RFI_ID",
    "Subject": "Subject",
    "Ball in Court": "Ball_in_Court",
    "Ball In Court": "Ball_in_Court",
    "Date Created": "Date_Created",
    "Date Initiated": "Date_Created",
    "Due Date": "Due_Date",
    "Status": "Status",
    "Overdue": "Overdue_Flag",
    "Due Date Variance": "Due_Date_Variance",
    # --- ADDITIONAL PROCORE VARIATIONS (future-proofing) ---
    "Number": "RFI_ID",
    "#": "RFI_ID",
    "RFI Number": "RFI_ID",
    "Description": "Subject",
    "Question": "Subject",
    "Title": "Subject",
    "Discipline": "Discipline",
    "Category": "Discipline",
    "Responsible Contractor": "Contractor",
    "Subcontractor": "Contractor",
    "Initiated By": "Contractor",
    "From": "Contractor",
    "Created By": "Contractor",
    "Priority": "Priority",
    "Importance": "Priority",
    "Response Due": "Due_Date",
    "Required Date": "Due_Date",
    "Date Closed": "Date_Closed",
    "Closed Date": "Date_Closed",
    "Date Answered": "Date_Closed",
    "Cost Impact": "Cost_Impact",
    "Schedule Impact": "Schedule_Impact",
}


# ============================================================
# HELPER FUNCTIONS
# ============================================================
def load_and_map(filepath, col_map):
    """Load Excel/CSV and auto-map columns."""
    if not os.path.exists(filepath):
        print(f"  ⚠️  File not found: {filepath}")
        return None

    ext = filepath.lower().split(".")[-1]
    if ext in ("xlsx", "xls"):
        df = pd.read_excel(filepath, engine="openpyxl")
    else:
        df = pd.read_csv(filepath)

    df.columns = df.columns.str.strip()

    # Map columns
    mapped = {}
    for orig in df.columns:
        if orig in col_map:
            mapped[orig] = col_map[orig]
    df = df.rename(columns=mapped)

    print(f"  ✅ Loaded {len(df)} rows from {filepath}")
    print(f"     Mapped columns: {list(mapped.values())}")
    unmapped = [c for c in df.columns if c not in mapped.values()]
    if unmapped:
        print(f"     Unmapped columns (kept as-is): {unmapped}")
    return df


def enrich_submittals(df):
    """Add calculated fields for Power BI — tailored to your Procore format."""

    # Parse any date columns present
    for col in ["Date_Created", "Due_Date", "Date_Closed", "Final_Due_Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", utc=True).dt.tz_localize(None)

    # ---- Days Open ----
    # Your Procore export has "Days Overdue" (mapped to Days_Overdue) — use it directly
    if "Days_Overdue" in df.columns:
        df["Days_Open"] = pd.to_numeric(df["Days_Overdue"], errors="coerce").fillna(0).astype(int)
    elif "Date_Created" in df.columns:
        df["Days_Open"] = df.apply(
            lambda r: (r.get("Date_Closed", pd.NaT) - r["Date_Created"]).days
            if pd.notna(r.get("Date_Closed"))
            else (TODAY - r["Date_Created"]).days
            if pd.notna(r["Date_Created"]) else 0,
            axis=1
        )
    elif "Due_Date" in df.columns:
        df["Days_Open"] = df["Due_Date"].apply(
            lambda x: max(0, (TODAY - x).days) if pd.notna(x) else 0
        )
    else:
        df["Days_Open"] = 0

    # ---- Is Open flag ----
    open_keywords = ["open", "pending", "revise", "draft", "submitted", "in review"]
    if "Status" in df.columns:
        df["Is_Open"] = df["Status"].apply(
            lambda s: any(k in str(s).lower() for k in open_keywords)
        )
    else:
        df["Is_Open"] = True  # If no status, assume open since this is "Open Submittals" export

    # ---- Overdue flag ----
    # Your Procore export has "Overdue" column (mapped to Overdue_Flag) — use it directly
    if "Overdue_Flag" in df.columns:
        df["Is_Overdue"] = df["Overdue_Flag"].apply(
            lambda x: str(x).strip().lower() in ["yes", "true", "1", "overdue", "y"]
        )
    else:
        df["Is_Overdue"] = df["Is_Open"] & (df["Days_Open"] > SUBMITTAL_OVERDUE_DAYS)

    # ---- Aging Bucket (for Power BI slicers) ----
    df["Aging_Bucket"] = pd.cut(
        df["Days_Open"],
        bins=[-1, 7, 14, 21, 30, 9999],
        labels=["0-7 days", "8-14 days", "15-21 days", "22-30 days", "30+ days"]
    )

    # ---- Time Intelligence ----
    # Use whichever date column is available
    date_col = None
    for c in ["Date_Created", "Due_Date", "Final_Due_Date"]:
        if c in df.columns and df[c].notna().any():
            date_col = c
            break

    if date_col and pd.api.types.is_datetime64_any_dtype(df[date_col]):
        df["Created_Week"] = df[date_col].dt.isocalendar().week.astype("Int64")
        df["Created_Month"] = df[date_col].dt.to_period("M").astype(str)
        df["Created_Year"] = df[date_col].dt.year.astype("Int64")

    # ---- Ball in Court as "Contractor" equivalent ----
    # Your export doesn't have a separate Contractor column
    # Ball in Court serves this purpose for tracking responsibility
    if "Ball_in_Court" in df.columns and "Contractor" not in df.columns:
        df["Contractor"] = df["Ball_in_Court"]

    # Item type label
    df["Item_Type"] = "Submittal"

    return df


def enrich_rfis(df):
    """Add calculated fields for Power BI — tailored to your Procore format."""

    # Parse any date columns present
    for col in ["Date_Created", "Due_Date", "Date_Closed"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", utc=True).dt.tz_localize(None)

    # ---- Days Open ----
    # Your Procore export has "Due Date Variance" (mapped to Due_Date_Variance)
    if "Due_Date_Variance" in df.columns:
        df["Days_Open"] = pd.to_numeric(df["Due_Date_Variance"], errors="coerce").fillna(0).abs().astype(int)
    elif "Date_Created" in df.columns:
        df["Days_Open"] = df.apply(
            lambda r: (r.get("Date_Closed", pd.NaT) - r["Date_Created"]).days
            if pd.notna(r.get("Date_Closed"))
            else (TODAY - r["Date_Created"]).days
            if pd.notna(r["Date_Created"]) else 0,
            axis=1
        )
    elif "Due_Date" in df.columns:
        df["Days_Open"] = df["Due_Date"].apply(
            lambda x: max(0, (TODAY - x).days) if pd.notna(x) else 0
        )
    else:
        df["Days_Open"] = 0

    # ---- Is Open flag ----
    open_keywords = ["open", "pending", "overdue", "draft", "in review"]
    if "Status" in df.columns:
        df["Is_Open"] = df["Status"].apply(
            lambda s: any(k in str(s).lower() for k in open_keywords)
        )
    else:
        df["Is_Open"] = True  # "Open RFIs" export implies all are open

    # ---- Overdue flag ----
    if "Overdue_Flag" in df.columns:
        df["Is_Overdue"] = df["Overdue_Flag"].apply(
            lambda x: str(x).strip().lower() in ["yes", "true", "1", "overdue", "y"]
        )
    else:
        df["Is_Overdue"] = df["Is_Open"] & (df["Days_Open"] > RFI_OVERDUE_DAYS)

    # ---- Aging Bucket ----
    df["Aging_Bucket"] = pd.cut(
        df["Days_Open"],
        bins=[-1, 5, 10, 15, 21, 9999],
        labels=["0-5 days", "6-10 days", "11-15 days", "16-21 days", "21+ days"]
    )

    # ---- Time Intelligence ----
    date_col = None
    for c in ["Date_Created", "Due_Date"]:
        if c in df.columns and df[c].notna().any():
            date_col = c
            break

    if date_col and pd.api.types.is_datetime64_any_dtype(df[date_col]):
        df["Created_Week"] = df[date_col].dt.isocalendar().week.astype("Int64")
        df["Created_Month"] = df[date_col].dt.to_period("M").astype(str)
        df["Created_Year"] = df[date_col].dt.year.astype("Int64")

    # ---- Ball in Court as "Contractor" equivalent ----
    if "Ball_in_Court" in df.columns and "Contractor" not in df.columns:
        df["Contractor"] = df["Ball_in_Court"]

    # Item type label
    df["Item_Type"] = "RFI"

    return df


def create_lookup_tables(df_sub, df_rfi):
    """Create dimension/lookup tables for Power BI star schema."""

    # Contractor / Ball in Court lookup
    parties = set()
    for df in [df_sub, df_rfi]:
        if df is not None:
            for col in ["Contractor", "Ball_in_Court"]:
                if col in df.columns:
                    parties |= set(df[col].dropna().unique())
    df_contractors = pd.DataFrame({
        "Contractor": sorted(parties),
        "Contractor_ID": range(1, len(parties) + 1)
    })

    # Status lookup
    statuses = set()
    for df in [df_sub, df_rfi]:
        if df is not None and "Status" in df.columns:
            statuses |= set(df["Status"].dropna().unique())
    df_statuses = pd.DataFrame({
        "Status": sorted(statuses),
        "Status_ID": range(1, len(statuses) + 1)
    })

    # Date table (for Power BI time intelligence)
    all_dates = []
    for df in [df_sub, df_rfi]:
        if df is not None:
            for col in ["Date_Created", "Due_Date", "Date_Closed", "Final_Due_Date"]:
                if col in df.columns:
                    valid = pd.to_datetime(df[col], errors="coerce").dropna()
                    all_dates.extend(valid.tolist())

    if all_dates:
        min_date = min(all_dates)
        max_date = max(max(all_dates), pd.Timestamp(TODAY))
        date_range = pd.date_range(start=min_date, end=max_date, freq="D")
        df_dates = pd.DataFrame({
            "Date": date_range,
            "Year": date_range.year,
            "Quarter": date_range.quarter,
            "Month": date_range.month,
            "Month_Name": date_range.strftime("%B"),
            "Week": date_range.isocalendar().week.astype(int),
            "Day_of_Week": date_range.strftime("%A"),
            "Is_Weekend": date_range.weekday >= 5,
        })
    else:
        df_dates = pd.DataFrame()

    return df_contractors, df_statuses, df_dates


def create_dax_measures():
    """Generate DAX measures for Power BI."""
    measures = """
============================================================
POWER BI DAX MEASURES — Copy these into Power BI
============================================================

--- SUBMITTAL MEASURES ---

Total Submittals = COUNTROWS(Submittals)

Open Submittals = CALCULATE(
    COUNTROWS(Submittals),
    Submittals[Is_Open] = TRUE
)

Overdue Submittals = CALCULATE(
    COUNTROWS(Submittals),
    Submittals[Is_Overdue] = TRUE
)

Submittal Closure Rate =
    DIVIDE(
        CALCULATE(COUNTROWS(Submittals), Submittals[Is_Open] = FALSE),
        COUNTROWS(Submittals),
        0
    )

Avg Submittal Days Open =
    AVERAGE(Submittals[Days_Open])

--- RFI MEASURES ---

Total RFIs = COUNTROWS(RFIs)

Open RFIs = CALCULATE(
    COUNTROWS(RFIs),
    RFIs[Is_Open] = TRUE
)

Overdue RFIs = CALCULATE(
    COUNTROWS(RFIs),
    RFIs[Is_Overdue] = TRUE
)

RFI Closure Rate =
    DIVIDE(
        CALCULATE(COUNTROWS(RFIs), RFIs[Is_Open] = FALSE),
        COUNTROWS(RFIs),
        0
    )

Avg RFI Response Time =
    AVERAGE(RFIs[Days_Open])

--- COMBINED MEASURES ---

Total Open Items = [Open Submittals] + [Open RFIs]

Total Overdue Items = [Overdue Submittals] + [Overdue RFIs]

Overall Health Score =
    1 - DIVIDE(
        [Total Overdue Items],
        [Open Submittals] + [Open RFIs],
        0
    )

--- CONDITIONAL FORMATTING MEASURE ---

Overdue Alert Color =
    SWITCH(
        TRUE(),
        [Days_Open] > 30, "#EF4444",
        [Days_Open] > 21, "#F59E0B",
        [Days_Open] > 14, "#FBBF24",
        "#10B981"
    )
"""
    return measures


# ============================================================
# MAIN EXECUTION
# ============================================================
def main():
    print("=" * 60)
    print("  POWER BI DATA PREP — Procore Transformer")
    print(f"  Run Date: {TODAY.strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)

    # Load files
    print("\n📂 Loading files...")
    df_sub = load_and_map(SUBMITTAL_FILE, SUBMITTAL_COL_MAP)
    df_rfi = load_and_map(RFI_FILE, RFI_COL_MAP)

    if df_sub is None and df_rfi is None:
        print("\n❌ No files found. Place your Procore exports in the same folder.")
        return

    # Enrich data
    print("\n🔧 Enriching data with calculated fields...")
    if df_sub is not None:
        df_sub = enrich_submittals(df_sub)
        print(f"  ✅ Submittals: {len(df_sub)} rows, {df_sub['Is_Overdue'].sum()} overdue")
        print(f"     Final columns: {list(df_sub.columns)}")
    if df_rfi is not None:
        df_rfi = enrich_rfis(df_rfi)
        print(f"  ✅ RFIs: {len(df_rfi)} rows, {df_rfi['Is_Overdue'].sum()} overdue")
        print(f"     Final columns: {list(df_rfi.columns)}")

    # Create lookup tables
    print("\n📊 Building lookup tables for star schema...")
    df_contractors, df_statuses, df_dates = create_lookup_tables(df_sub, df_rfi)
    print(f"  ✅ Contractors/Parties: {len(df_contractors)}")
    print(f"  ✅ Statuses: {len(df_statuses)}")
    print(f"  ✅ Date table: {len(df_dates)} days")

    # Write to Excel (multi-sheet for Power BI)
    print(f"\n💾 Writing to {OUTPUT_FILE}...")
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        if df_sub is not None:
            df_sub.to_excel(writer, sheet_name="Submittals", index=False)
        if df_rfi is not None:
            df_rfi.to_excel(writer, sheet_name="RFIs", index=False)
        df_contractors.to_excel(writer, sheet_name="Dim_Contractors", index=False)
        df_statuses.to_excel(writer, sheet_name="Dim_Statuses", index=False)
        if not df_dates.empty:
            df_dates.to_excel(writer, sheet_name="Dim_Dates", index=False)

        # DAX measures as reference sheet
        dax = create_dax_measures()
        dax_df = pd.DataFrame({"DAX_Measures": dax.split("\n")})
        dax_df.to_excel(writer, sheet_name="DAX_Reference", index=False)

    print(f"  ✅ Saved: {OUTPUT_FILE}")
    print(f"\n{'=' * 60}")
    print("  NEXT STEPS IN POWER BI:")
    print("  1. Open Power BI Desktop")
    print(f"  2. Get Data → Excel → select {OUTPUT_FILE}")
    print("  3. Load all sheets (Submittals, RFIs, Dim_*)")
    print("  4. Set up relationships in Model view:")
    print("     • Submittals[Contractor] → Dim_Contractors[Contractor]")
    print("     • RFIs[Contractor] → Dim_Contractors[Contractor]")
    print("     • Submittals[Date_Created] → Dim_Dates[Date]")
    print("     • RFIs[Date_Created] → Dim_Dates[Date]")
    print("  5. Copy DAX measures from the DAX_Reference sheet")
    print("  6. Build visuals (see recommended layout below)")
    print(f"{'=' * 60}")

    print("""
  📐 RECOMMENDED POWER BI DASHBOARD LAYOUT:
  ──────────────────────────────────────────
  PAGE 1: Executive Overview
    • KPI cards: Open / Overdue / Closure Rate (Submittals & RFIs)
    • Donut chart: Status distribution
    • Stacked bar: Items by Ball in Court
    • Line chart: Cumulative open items over time

  PAGE 2: Submittal Deep Dive
    • Table: Full submittal log with conditional formatting
    • Bar chart: Days Open by Ball in Court
    • Treemap: Responsibility breakdown
    • Slicer: Ball in Court, Status, Aging Bucket

  PAGE 3: RFI Deep Dive
    • Table: Full RFI log with conditional formatting
    • Bar chart: RFIs by Ball in Court
    • Timeline: RFI creation trend over time
    • Slicer: Status, Aging Bucket

  PAGE 4: Bottleneck Analysis
    • Scatter: Days Open vs Due Date
    • Funnel: Open → Pending → Closed flow
    • Gauge: Overall Health Score
    """)


if __name__ == "__main__":
    main()