import io
from datetime import datetime
from typing import Dict, Any, List

import pandas as pd
import numpy as np
import streamlit as st

# Use matplotlib for PDF creation and charts.  Matplotlib is widely available and
# allows us to render charts and multiple pages in a single PDF without relying
# on the reportlab dependency.  We import the PdfPages backend to append each
# figure as a page in the report.
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

# ---- Hard-coded mappings (from your uploaded files at build time) ----
CRAFT_ORDER = [
    "Turns",
    "EAF Mech Days",
    "EAF Elec Days",
    "AOD Mech Days",
    "AOD Elec Days",
    "Alloy Mech Days",
    "Caster Mech Days",
    "Caster Elec Days",
    "WTP Mech Days",
    "Baghouse Mech Days",
    "Preheater Elec Days",
    "Segment Shop",
    "Utilities Mech Days",
    "HVAC Elec Days",
]

# Address book mapping of employee numbers to name and craft description.
# This data is embedded so the application works without external dependencies.
ADDRESS_BOOK = [
    {"AddressBookNumber": "1103079", "Name": "CONKEL, JOHNATHON J", "Craft Description": "Alloy Mech Days"},
    {"AddressBookNumber": "817150", "Name": "PETERS, JESSE DANIEL", "Craft Description": "AOD Elec Days"},
    {"AddressBookNumber": "648991", "Name": "JONES, TERRELL D.", "Craft Description": "AOD Elec Days"},
    {"AddressBookNumber": "136792", "Name": "MCKINNEY, CHRIS ALVIE", "Craft Description": "AOD Mech Days"},
    {"AddressBookNumber": "1142730", "Name": "CHRISTERSON, NATHANIEL BENJAMEN", "Craft Description": "Baghouse Mech Days"},
    {"AddressBookNumber": "1150094", "Name": "WRIGHT, KEVIN BRADLEY", "Craft Description": "Baghouse Mech Days"},
    {"AddressBookNumber": "1064305", "Name": "DALTON II, JEFFERY WAYNE", "Craft Description": "Caster Elec Days"},
    {"AddressBookNumber": "1115109", "Name": "GANDER, ANTHONY T", "Craft Description": "Caster Elec Days"},
    {"AddressBookNumber": "1055943", "Name": "HEFFELMIRE, RONALD SCOTT", "Craft Description": "Caster Mech Days"},
    {"AddressBookNumber": "1112813", "Name": "KOONS, ANDREW LEWIS ALAN", "Craft Description": "Caster Mech Days"},
    {"AddressBookNumber": "95636", "Name": "MORRISON, GEORGE D.", "Craft Description": "Caster Mech Days"},
    {"AddressBookNumber": "586013", "Name": "DENNIS, SHAWN MICHAEL", "Craft Description": "EAF Elec Days"},
    {"AddressBookNumber": "1137121", "Name": "STEWART, THOMAS JASON", "Craft Description": "EAF Mech Days"},
    {"AddressBookNumber": "1106595", "Name": "WASH, MICHAEL DAVID", "Craft Description": "EAF Mech Days"},
    {"AddressBookNumber": "178909", "Name": "LEMASTER, DANIEL M.", "Craft Description": "HVAC Elec Days"},
    {"AddressBookNumber": "1115133", "Name": "BROCK, TREVOR COLE", "Craft Description": "Preheater Elec Days"},
    {"AddressBookNumber": "133760", "Name": "BRIGHTWELL, JEFFERY W.", "Craft Description": "Segment Shop"},
    {"AddressBookNumber": "173665", "Name": "CRAIG, JAMES D.", "Craft Description": "Segment Shop"},
    {"AddressBookNumber": "336719", "Name": "DEEN, ALAN J.", "Craft Description": "Segment Shop"},
    {"AddressBookNumber": "1151409", "Name": "DEMAREE, MATTHEW CHRISTOPHER", "Craft Description": "Segment Shop"},
    {"AddressBookNumber": "848802", "Name": "KLOSS, CHARLES W.", "Craft Description": "Segment Shop"},
    {"AddressBookNumber": "95644", "Name": "SMITH, JAMES M.", "Craft Description": "Segment Shop"},
    {"AddressBookNumber": "1104469", "Name": "WATSON, JACOB LEYTON", "Craft Description": "Segment Shop"},
    {"AddressBookNumber": "1103976", "Name": "BAUGHMAN, THOMAS BRUCE", "Craft Description": "Turns"},
    {"AddressBookNumber": "1066095", "Name": "HELTON, MICHAEL AJ", "Craft Description": "Turns"},
    {"AddressBookNumber": "44231", "Name": "REED, BRIAN L.", "Craft Description": "Turns"},
    {"AddressBookNumber": "1167030", "Name": "STROUD, MATTHEW T.", "Craft Description": "Turns"},
    {"AddressBookNumber": "164380", "Name": "WARREN, MARK L.", "Craft Description": "Turns"},
    {"AddressBookNumber": "185050", "Name": "WHOBREY, BRADLEY G.", "Craft Description": "Turns"},
    {"AddressBookNumber": "1103132", "Name": "WILLIAMS II, STEVEN FOSTER", "Craft Description": "Turns"},
    {"AddressBookNumber": "1106747", "Name": "BANTA, BENJAMIN GAYLE", "Craft Description": "Turns"},
    {"AddressBookNumber": "1165384", "Name": "BOHART, WILLIAM M.", "Craft Description": "Turns"},
    {"AddressBookNumber": "1144250", "Name": "CAREY, JOSEPH MICHAEL", "Craft Description": "Turns"},
    {"AddressBookNumber": "770363", "Name": "CORMAN, DAVID H.", "Craft Description": "Turns"},
    {"AddressBookNumber": "1149608", "Name": "DIEDERICH, JOSEPH W", "Craft Description": "Turns"},
    {"AddressBookNumber": "193471", "Name": "GRAY, DENNIS C.", "Craft Description": "Turns"},
    {"AddressBookNumber": "109866", "Name": "HOWARD, LARRY D.", "Craft Description": "Turns"},
    {"AddressBookNumber": "1141761", "Name": "PHILLIPS, TIMOTHY CRAIG RYAN", "Craft Description": "Turns"},
    {"AddressBookNumber": "272006", "Name": "SEE, JOHN JOURDAN", "Craft Description": "Turns"},
    {"AddressBookNumber": "106260", "Name": "SPILLMAN, WILLIAM H.", "Craft Description": "Turns"},
    {"AddressBookNumber": "1131299", "Name": "STEWART, BRADFORD LEE", "Craft Description": "Turns"},
    {"AddressBookNumber": "1109876", "Name": "STOKES, MATHEW DAVID", "Craft Description": "Turns"},
    {"AddressBookNumber": "234448", "Name": "THOMAS, CODY JORDAN", "Craft Description": "Turns"},
    {"AddressBookNumber": "1107352", "Name": "WATKINS, KENNETH EDWARD", "Craft Description": "Turns"},
    {"AddressBookNumber": "1096665", "Name": "ATWELL, TALON BRADLEY", "Craft Description": "Turns"},
    {"AddressBookNumber": "108986", "Name": "ROGERS, CHARLES D.", "Craft Description": "Utilities Mech Days"},
    {"AddressBookNumber": "206092", "Name": "TURNER, SHANE M.", "Craft Description": "Utilities Mech Days"},
    {"AddressBookNumber": "1089377", "Name": "ROSE, CAMERON CHASE", "Craft Description": "WTP Mech Days"},
]

# Mapping from work order type codes to descriptive names.
TYPE_MAP: Dict[str, str] = {
    "0": "Break In",
    "1": "Standard Corrective",
    "2": "Material Repair TMJ Order",
    "3": "Capital Project",
    "4": "Urgent Corrective",
    "5": "Emergency Order",
    "6": "PM Restore/Replace",
    "7": "PM Inspection",
    "8": "Follow Up Work Order",
    "9": "Standing W.O. - Do not Delete",
    "B": "Marketing",
    "C": "Cost Improvement",
    "D": "Design Work - ETO",
    "E": "Plant Work - ETO",
    "G": "Governmental/Regulatory",
    "M": "Model W.O. - Eq Mgmt",
    "N": "Template W.O. - CBM Alerts",
    "P": "Project",
    "R": "Rework Order",
    "S": "Shop Order",
    "T": "Tool Order",
    "W": "Case",
    "X": "General Work Request",
    "Y": "Follow Up Work Request",
    "Z": "System Work Request",
}

# Columns used for display and included in the PDF and data downloads.
DISPLAY_COLUMNS: List[str] = ["Name", "Work Order #", "Sum of Hours", "Type", "CostCenter", "Description", "Problem"]

# Columns expected in the uploaded Excel export.  If any are missing they will be
# created with missing values.
REQUIRED_TIME_COLUMNS: List[str] = ["AddressBookNumber", "Name", "Production Date", "OrderNumber", "Sum of Hours.", "Hours Estimated", "Status", "Type", "PMFrequency", "Description", "Problem", "Department", "Location", "Equipment", "PM Number", "PM", "CostCenter"]

def _find_header_row(df_raw: pd.DataFrame) -> int:
    """Locate the header row in an Excel export by searching for a row that
    contains the expected column names.  Raises if the header cannot be found.
    """
    first_col = df_raw.columns[0]
    mask = df_raw[first_col].astype(str).str.strip() == "AddressBookNumber"
    idx = df_raw.index[mask].tolist()
    if idx:
        return idx[0]
    # Fallback: search the first few rows for known column names.
    for i in range(min(10, len(df_raw))):
        row_vals = df_raw.iloc[i].astype(str).str.strip().tolist()
        if "AddressBookNumber" in row_vals and "Production Date" in row_vals:
            return i
    raise ValueError("Could not locate header row containing 'AddressBookNumber'.")

def _read_excel_twice(file) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Read the uploaded Excel file twice: once without headers and once with
    pandas' default header detection.  This allows us to locate the header row and
    preserve dtypes."""
    data = file.read()
    df_raw = pd.read_excel(io.BytesIO(data), header=None, dtype=str)
    df_hdr = pd.read_excel(io.BytesIO(data))
    return df_raw, df_hdr

def load_timeworkbook(file_like) -> pd.DataFrame:
    """Load and clean the Time on Work Order export."""
    df_raw, _ = _read_excel_twice(file_like)
    header_row = _find_header_row(df_raw)
    file_like.seek(0)
    df = pd.read_excel(file_like, header=header_row)
    # Drop unnamed columns
    df = df.loc[:, ~df.columns.astype(str).str.contains(r"^Unnamed")]
    # === Ensure CostCenter column exists and is cleaned ===
    try:
        if "CostCenter" not in df.columns:
            if df.shape[1] > 13:
                df["CostCenter"] = df.iloc[:, 13]
            else:
                df["CostCenter"] = pd.NA
        df["CostCenter"] = df["CostCenter"].astype(str).fillna("").str.strip().replace({"nan": ""})
    except Exception:
        df["CostCenter"] = pd.NA
    # === End CostCenter ensure ===

    # Ensure all required columns exist
    for c in REQUIRED_TIME_COLUMNS:
        if c not in df.columns:
            df[c] = pd.NA
    # === Custom: Map CostCenter (column N) into 'CC' from the SAME uploaded sheet ===
    try:
        if "CostCenter" in df.columns:
            df["CC"] = df["CostCenter"].astype(str).fillna("").str.strip().replace({"nan": ""})
        else:
            # Fallback by position: column N (14th column, index 13) from the already-read df
            if df.shape[1] > 13:
                _series = df.iloc[:, 13].astype(str)
                df["CC"] = _series.fillna("").str.strip().replace({"nan": ""})
            else:
                df["CC"] = pd.NA
    except Exception:
        df["CC"] = pd.NA
    # === End Custom ===

    df["AddressBookNumber"] = df["AddressBookNumber"].astype(str).str.strip()
    if "Production Date" in df.columns:
        df["Production Date"] = pd.to_datetime(df["Production Date"], errors="coerce").dt.date
    # Normalize 'Sum of Hours' column names and values
    if "Sum of Hours" in df.columns:
        df["Sum of Hours"] = pd.to_numeric(df["Sum of Hours"], errors="coerce")
    elif "Sum of Hours." in df.columns:
        df["Sum of Hours"] = pd.to_numeric(df["Sum of Hours."], errors="coerce")
    elif "Hours" in df.columns:
        df["Sum of Hours"] = pd.to_numeric(df["Hours"], errors="coerce")
    else:
        df["Sum of Hours"] = pd.NA
    # Normalize work order number into 'Work Order #'
    if "Work Order Number" in df.columns:
        base_wo = df["Work Order Number"]
    elif "OrderNumber" in df.columns:
        base_wo = df["OrderNumber"]
    elif "WO Number" in df.columns:
        base_wo = df["WO Number"]
    elif "WorkOrderNumber" in df.columns:
        base_wo = df["WorkOrderNumber"]
    else:
        base_wo = pd.Series([pd.NA] * len(df))
    df["Work Order #"] = base_wo.astype(str).str.replace(r"\.0$", "", regex=True)
    if "Problem" not in df.columns:
        df["Problem"] = pd.NA
    return df

def get_craft_order_df() -> pd.DataFrame:
    """Return a DataFrame with the craft order."""
    return pd.DataFrame({"Craft Description": CRAFT_ORDER})

def get_address_book_df() -> pd.DataFrame:
    """Return a cleaned DataFrame of the address book."""
    df = pd.DataFrame(ADDRESS_BOOK)[["AddressBookNumber", "Name", "Craft Description"]]
    df["AddressBookNumber"] = df["AddressBookNumber"].astype(str).str.strip()
    df["Name"] = df["Name"].astype(str).str.strip()
    df["Craft Description"] = df["Craft Description"].astype(str).str.strip()
    return df

def _apply_craft_category(df: pd.DataFrame, order_df: pd.DataFrame) -> pd.DataFrame:
    """Assign categories to the craft description column so crafts are ordered consistently."""
    order = order_df["Craft Description"].tolist()
    seen: set[str] = set()
    ordered: List[str] = []
    for c in order:
        if c not in seen:
            ordered.append(c)
            seen.add(c)
    categories = ordered + ["Unassigned"]
    df["Craft Description"] = df["Craft Description"].fillna("Unassigned")
    df["Craft Description"] = pd.Categorical(df["Craft Description"], categories=categories, ordered=True)
    return df

def _clean_code(value: object) -> str | None:
    """Normalize numeric work order codes to strings and return None for missing values."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if s == "" or s.lower() == "nan":
        return None
    if s.endswith(".0"):
        s = s[:-2]
    return s

def _map_type(value):
    """Map a work order type code to its descriptive name using TYPE_MAP."""
    key = _clean_code(value)
    if key is None:
        return pd.NA
    out = TYPE_MAP.get(key, key)
    # Clean up the inspection maintenance wording
    if isinstance(out, str) and out.strip().lower() == "inspection maintenance order":
        return "PM Inspection"
    return out

def prepare_report_data(
    time_df: pd.DataFrame,
    addr_df: pd.DataFrame,
    craft_order_df: pd.DataFrame,
    selected_date,
) -> Dict[str, Any]:
    """Prepare the report data for the selected production date.

    Returns a dictionary with keys:
      - ``groups``: a list of tuples (craft_name, payload) where payload contains a
        'detail' DataFrame with the selected columns.
      - ``full_detail``: the complete filtered DataFrame of selected columns.
      - ``unmapped_people``: a list of address book entries that could not be mapped.
    """
    f = time_df[time_df["Production Date"] == selected_date].copy()
    f["AddressBookNumber"] = f["AddressBookNumber"].astype(str).str.strip()
    addr_df["AddressBookNumber"] = addr_df["AddressBookNumber"].astype(str).str.strip()
    merged = f.merge(
        addr_df[["AddressBookNumber", "Craft Description", "Name"]].rename(columns={"Name": "AB_Name"}),
        on="AddressBookNumber",
        how="left",
    )
    merged["Name"] = merged["Name"].fillna(merged["AB_Name"])
    merged = merged.drop(columns=["AB_Name"])
    # Identify unmapped people
    unmapped: List[Dict[str, Any]] = []
    mask_unmapped = merged["Craft Description"].isna() | (merged["Craft Description"].astype(str).str.len() == 0)
    if mask_unmapped.any():
        unmapped = (
            merged.loc[mask_unmapped, ["AddressBookNumber", "Name"]]
            .drop_duplicates()
            .to_dict("records")
        )
    merged.loc[mask_unmapped, "Craft Description"] = "Unassigned"
    merged = _apply_craft_category(merged, craft_order_df)
    # === Custom: Ensure 'CC' is populated after merge (fallback to 'CostCenter' if needed) ===
    try:
        if "CC" not in merged.columns or merged["CC"].isna().all():
            if "CostCenter" in merged.columns:
                merged["CC"] = merged["CostCenter"].astype(str).fillna("").str.strip().replace({"nan": ""})
    except Exception:
        pass
    # === End Custom ===

    # Ensure display columns exist
    for col in DISPLAY_COLUMNS:
        if col not in merged.columns:
            merged[col] = pd.NA
    merged["Type"] = merged["Type"].apply(_map_type)
    merged = merged.sort_values(["Craft Description", "Name", "Work Order #"])
    merged["Sum of Hours"] = pd.to_numeric(merged["Sum of Hours"], errors="coerce").round(2)
    groups_payload: List = []
    for craft in list(merged["Craft Description"].cat.categories):
        g_detail = merged[merged["Craft Description"] == craft][DISPLAY_COLUMNS].copy()
        if g_detail.empty:
            continue
        groups_payload.append((str(craft), {"detail": g_detail}))
    full_detail = merged[DISPLAY_COLUMNS].copy()
    return {"groups": groups_payload, "full_detail": full_detail, "unmapped_people": unmapped}

def _auto_height(df: pd.DataFrame) -> int:
    """Calculate a table height for the Streamlit dataframe component."""
    rows = len(df) + 1
    row_px = 35
    header_px = 40
    return min(header_px + rows * row_px, 20000)

def _df_for_pdf(df: pd.DataFrame) -> pd.DataFrame:
    """Format numeric columns for inclusion in the PDF."""
    out = df.copy()
    # Format "Sum of Hours" as a string with two decimal places.  Convert to
    # numeric first so non-numeric values become NaN and then to formatted
    # string.  Leave as-is if missing.
    if "Sum of Hours" in out.columns:
        out["Sum of Hours"] = (
            pd.to_numeric(out["Sum of Hours"], errors="coerce")
            .fillna(0)
            .map(lambda x: f"{x:.2f}")
        )
    # Normalize cost center column for the PDF.  If a column named "CC" is
    # missing but "CostCenter" exists, rename it to "CC".  If both exist
    # concurrently, prefer the existing "CC" column and drop "CostCenter".
    if "CC" not in out.columns:
        if "CostCenter" in out.columns:
            out = out.rename(columns={"CostCenter": "CC"})
    else:
        # Drop any redundant CostCenter column when CC is present
        if "CostCenter" in out.columns:
            out = out.drop(columns=["CostCenter"])
    # Reorder columns to ensure a consistent layout in the PDF.  Include the
    # cost center (CC) column between Type and Description.
    desired_order = [
        "Name",
        "Work Order #",
        "Sum of Hours",
        "Type",
        "CC",
        "Description",
        "Problem",
    ]
    # Filter out columns not present and drop duplicates if any
    order = [col for col in desired_order if col in out.columns]
    out = out[order]
    return out

# -----------------------------------------------------------------------------
# PDF helper functions
#
# The following helpers render figures for the PDF export.  To more closely
# mirror the Streamlit dashboard layout, each craft will have a summary page
# containing the bar chart of hours by work order type, some key metrics,
# and a breakdown table.  Additional pages are then added for the detailed
# work order table, splitting the rows across multiple pages as needed.  These
# helpers return matplotlib Figure objects which the main ``build_pdf``
# function writes to the PdfPages instance.
def _create_summary_figure(df_detail: pd.DataFrame, craft_name: str, max_type_count: int | None = None) -> plt.Figure:
    """
    Create a summary figure for a single craft.  The layout consists of a bar
    chart showing hours by work order type, a metrics text block, and a
    breakdown table of hours and percentages by type.  This attempts to
    approximate the Streamlit dashboard view.

    :param df_detail: DataFrame with columns including 'Type' and 'Sum of Hours'.
    :param craft_name: Name of the craft group.
    :return: A matplotlib Figure ready for saving to the PDF.
    """
    # Prepare aggregated data
    df = df_detail.copy()
    df["Sum of Hours"] = pd.to_numeric(df["Sum of Hours"], errors="coerce").fillna(0.0)
    agg = (
        df.groupby("Type", dropna=False)["Sum of Hours"]
        .sum()
        .reset_index()
        .rename(columns={"Sum of Hours": "hours"})
        .sort_values("hours", ascending=False)
    )
    total = float(agg["hours"].sum())
    agg["percent"] = np.where(total > 0, (agg["hours"] / total) * 100.0, 0.0)
    top_type = agg.iloc[0]["Type"] if not agg.empty else "-"
    top_pct = agg.iloc[0]["percent"] if not agg.empty else 0.0

    # Create figure and axes using a grid layout
    fig = plt.figure(figsize=(11, 8.5))
    fig.suptitle(f"{craft_name} — Summary", fontsize=16, y=0.98)
    # Bar chart axis occupies the left portion of the page
    ax_bar = fig.add_axes([0.05, 0.35, 0.55, 0.5])
    default_color = "#333333"
    colors_list = []
    for t in agg["Type"]:
        if isinstance(t, str):
            colors_list.append(_TYPE_COLORS.get(str(t), default_color))
        else:
            colors_list.append(default_color)
    # Render bars with a fixed width so a single bar does not occupy the
    # entire chart width.  Positions are spaced uniformly along the x-axis.
    positions = np.arange(len(agg))
    # Use a fixed bar width so that bars retain the same thickness regardless
    # of how many categories are present.  A value less than 1.0 ensures
    # spacing between bars when the number of categories is small.  Add a thin
    # black outline around each bar for visual separation.
    bar_width = 0.8
    ax_bar.bar(
        positions,
        agg["hours"],
        color=colors_list,
        width=bar_width,
        edgecolor="black",
        linewidth=0.5,
    )
    # If a maximum type count is supplied, fix the x-axis limits so that the
    # bar chart maintains the same overall width across different groups.  When
    # fewer types are present than ``max_type_count``, additional empty
    # positions ensure the bars do not stretch to fill the axis.  If not
    # provided, the x-axis will scale automatically.
    if max_type_count is not None:
        try:
            max_val = int(max_type_count)
        except Exception:
            max_val = len(agg)
        # Ensure the axis accommodates at least the number of actual categories
        max_val = max(max_val, len(agg))
        ax_bar.set_xlim(-0.5, max_val - 0.5)
    ax_bar.set_xticks(positions)
    ax_bar.set_xticklabels(agg["Type"].astype(str), rotation=45, fontsize=8)
    ax_bar.set_title("Hours by Work Order Type", fontsize=12)
    ax_bar.set_xlabel("Type")
    ax_bar.set_ylabel("Hours")
    # Add horizontal margins to avoid bars touching the axes
    ax_bar.margins(x=0.15)
    # Reposition metrics to the right side above the breakdown table to avoid overlap
    fig.text(0.65, 0.82, f"Total Hours: {total:,.2f}", fontsize=10, va="top")
    fig.text(0.65, 0.79, f"Top Type: {top_type}", fontsize=10, va="top")
    fig.text(0.65, 0.76, f"% in Top Type: {top_pct:.1f}%", fontsize=10, va="top")
    # Breakdown table axis occupies the right portion below the metrics
    # Allocate the breakdown table to the lower half of the right side.  Using a height
    # of 0.4 (from 0.35 to 0.75) leaves room above for the metrics.
    ax_tbl = fig.add_axes([0.65, 0.35, 0.3, 0.4])
    ax_tbl.axis("off")
    tbl_data = [
        [str(row["Type"]), f"{row['hours']:.2f}", f"{row['percent']:.1f}%"]
        for _, row in agg.iterrows()
    ]
    col_labels = ["Type", "Hours", "%"]
    table = ax_tbl.table(
        cellText=tbl_data,
        colLabels=col_labels,
        cellLoc="left",
        loc="upper left",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.auto_set_column_width(col=list(range(len(col_labels))))
    return fig

def _create_detail_table_figures(
    df_detail: pd.DataFrame,
    craft_name: str,
    rows_per_page: int = 20,
    shade_by_work_order: bool = False,
) -> List[plt.Figure]:
    """
    Split a detailed DataFrame into multiple figures for inclusion in the PDF.

    This function dynamically determines how many rows fit on a single page
    based on the amount of text that appears in the Description and Problem
    columns.  Rows that contain wrapped text consume more vertical space than
    those with single-line text.  By calculating the height required for each
    row and breaking pages only when the accumulated height exceeds the
    available space, tables will never run off the bottom of the page.  In
    addition, header labels are wrapped to fit within their columns and the
    Type column is color coded with bold white text to match the dashboard.

    :param df_detail: DataFrame with the detail rows to display.
    :param craft_name: Name of the craft group.
    :param rows_per_page: A nominal maximum number of rows per page.  This
        value is used to determine a base row height.  Pages may contain
        fewer rows if wrapping causes rows to require additional height.
    :param shade_by_work_order: When True, apply alternating row shading based
        on groups of identical Work Order # values.  Shading toggles between
        shaded and unshaded each time the Work Order # changes in the sorted
        data.  Only non-Type columns receive the shading; the Type column
        retains its color-coded formatting.
    :return: List of matplotlib Figures.
    """
    import textwrap

    # Fixed relative widths for each of the seven columns.  These values sum
    # to 1.0 and correspond to the column order produced by ``_df_for_pdf``:
    # Name, Work Order #, Sum of Hours, Type, CC, Description, Problem.
    # Column widths have been updated per user request:
    #   Name          : 18% (0.18)
    #   Work Order #  : 6%  (0.06)
    #   Sum of Hours  : 5%  (0.05)
    #   Type          : 14% (0.14)
    #   CC            : 3%  (0.03)
    #   Description   : 17% (0.17)
    #   Problem       : 37% (0.37)
    col_widths = [0.18, 0.06, 0.05, 0.14, 0.03, 0.17, 0.37]

    figures: List[plt.Figure] = []
    # Convert to a simple DataFrame for PDF output
    df_pdf = _df_for_pdf(df_detail)
    col_labels = df_pdf.columns.tolist()

    # Wrap header labels to prevent truncation.  Insert manual line breaks for
    # specific headers and use a small width for others so multi-word labels
    # wrap naturally.
    header_wrapped: List[str] = []
    for lbl in col_labels:
        if lbl == "Work Order #":
            header_wrapped.append("Work Order\n#")
        elif lbl == "Sum of Hours":
            header_wrapped.append("Sum of\nHours")
        else:
            wrapped = textwrap.fill(
                lbl,
                width=10,
                break_long_words=False,
                break_on_hyphens=False,
            )
            header_wrapped.append(wrapped)

    # Identify the index of the "Type" column so we can color it later.
    try:
        type_col_index = col_labels.index("Type")
    except ValueError:
        type_col_index = None

    # Prepare wrapped DataFrame for both content and line counting.  Wrap
    # Description and Problem columns using a narrower width to ensure
    # long text breaks across multiple lines within the PDF column.  We use
    # a width of 50 characters which has been empirically found to fit
    # comfortably within the Problem column at a small font size.
    df_wrapped = df_pdf.copy()
    for col in ["Description", "Problem"]:
        if col in df_wrapped.columns:
            df_wrapped[col] = (
                df_wrapped[col]
                .astype(str)
                .apply(lambda x: textwrap.fill(x, width=50, break_long_words=True, break_on_hyphens=True))
            )

    # Determine line counts for each row based on the wrapped Description and
    # Problem columns.  The number of lines in a row is the maximum of the
    # line counts in these two columns (or 1 if both are empty).
    line_counts: List[int] = []
    for i in range(len(df_wrapped)):
        max_lines = 1
        for col in ["Description", "Problem"]:
            if col in df_wrapped.columns:
                val = str(df_wrapped.iloc[i][col])
                lines = val.count("\n") + 1 if val else 1
                if lines > max_lines:
                    max_lines = lines
        line_counts.append(max_lines)

    # Compute shading flags if requested.  Alternating shading toggles on group
    # boundaries.  The first group is shaded and subsequent groups toggle
    # shading status.  We convert work order values to strings to avoid
    # mismatches due to numeric conversion.
    shading_flags: List[bool] = [False] * len(df_wrapped)
    if shade_by_work_order and "Work Order #" in df_pdf.columns:
        prev_value = None
        shade = False
        for idx in range(len(df_pdf)):
            curr_value = str(df_pdf.iloc[idx]["Work Order #"])
            if curr_value != prev_value:
                # toggle shade when work order changes
                shade = not shade
                prev_value = curr_value
            shading_flags[idx] = shade

    # Determine header line count to set header row height.  The tallest
    # header cell determines the number of lines.
    header_line_counts = [lbl.count("\n") + 1 for lbl in header_wrapped]
    header_lines = max(header_line_counts) if header_line_counts else 1

    # Compute a base height per line.  We allocate 85% of the axis height to
    # the table body and header combined.  The nominal rows_per_page is used
    # to determine how tall a single-line row should be.  If all rows were
    # single-line and there were exactly rows_per_page rows, everything
    # would fit.  Extra lines in wrapped rows reduce the number of rows per
    # page automatically.
    base_height = 0.85 / (rows_per_page + 1)
    header_height = base_height * header_lines

    # Build pages by accumulating rows until the accumulated height exceeds
    # the available space.  The available vertical space for data rows on
    # each page is 0.85 (the same space used later for the axes).  We start
    # with the header height already allocated.
    pages: List[List[int]] = []  # each entry is a list of row indices
    current_page: List[int] = []
    current_height = header_height
    for idx, row_lines in enumerate(line_counts):
        row_height = base_height * row_lines
        # If adding this row would exceed the available space, start a new page.
        if current_height + row_height > 0.85 and current_page:
            pages.append(current_page)
            current_page = [idx]
            current_height = header_height + row_height
        else:
            current_page.append(idx)
            current_height += row_height
    # Append the last page
    if current_page:
        pages.append(current_page)

    # Generate a figure for each page
    for page_num, page_rows in enumerate(pages, start=1):
        # Slice the wrapped DataFrame to only include rows on this page
        chunk = df_wrapped.iloc[page_rows].copy()
        # Create the figure and axis for this page.  Landscape orientation is
        # used to match the rest of the report.  The title reflects the
        # current page number and the total number of pages.
        fig = plt.figure(figsize=(11, 8.5))
        fig.suptitle(
            f"{craft_name} — Detail (Page {page_num}/{len(pages)})",
            fontsize=14,
            y=0.95,
        )
        ax = fig.add_axes([0.05, 0.05, 0.90, 0.85])
        ax.axis("off")
        # Build the table with fixed column widths.  Provide the wrapped
        # headers and content.  All cells are left aligned.
        table = ax.table(
            cellText=chunk.values,
            colLabels=header_wrapped,
            cellLoc="left",
            loc="upper left",
            colWidths=col_widths,
        )
        table.auto_set_font_size(False)
        table.set_fontsize(7)
        # Set header row height based on header_lines
        for col_idx in range(len(header_wrapped)):
            table[(0, col_idx)].set_height(header_height)
        # Set row heights and apply color formatting for the Type column
        for r_idx, row_idx in enumerate(page_rows):
            lines = line_counts[row_idx]
            row_height = base_height * lines
            # Set height for each cell in the row
            for col_idx in range(len(header_wrapped)):
                table[(r_idx + 1, col_idx)].set_height(row_height)
            # Apply color coding to Type column
            if type_col_index is not None:
                try:
                    type_value = str(chunk.iloc[r_idx]["Type"])
                except Exception:
                    type_value = ""
                color = _TYPE_COLORS.get(type_value, "#FFFFFF")
                cell = table[(r_idx + 1, type_col_index)]
                cell.set_facecolor(color)
                # Set text color and weight
                text = cell.get_text()
                text.set_color("white")
                text.set_weight("bold")

            # Apply alternating shading for non-Type cells if enabled
            if shade_by_work_order and 0 <= row_idx < len(shading_flags):
                if shading_flags[row_idx]:
                    # Choose a very light blue for shading; skip Type column
                    shade_color = "#eaf3ff"
                    for col_idx in range(len(header_wrapped)):
                        if type_col_index is not None and col_idx == type_col_index:
                            continue
                        cell = table[(r_idx + 1, col_idx)]
                        # Only override facecolor if no color is already set
                        # Use set_facecolor directly to apply shading
                        cell.set_facecolor(shade_color)
        figures.append(fig)
    return figures

def build_pdf(
    report: Dict[str, Any],
    date_label: str,
    sort_option: str = "Name",
    shade_by_work_order: bool | None = None,
) -> bytes:
    """
    Construct a PDF report with summary charts for each craft group along with
    overall statistics.  This implementation uses matplotlib to render bar charts
    and writes each figure to a multipage PDF via PdfPages.  A title page and an
    optional overall summary page are included before the per-craft pages.

    :param report: The dictionary produced by ``prepare_report_data`` with keys
        ``groups`` (list of (craft_name, payload) tuples) and
        ``full_detail`` (DataFrame containing all records).
    :param date_label: The selected date string used in the report title and
        output file name.
    :return: Bytes representing the PDF document.
    """
    buffer = io.BytesIO()
    # Determine the maximum number of unique work order types across all groups.  This
    # value is used to fix the x-axis width of bar charts so that bar width
    # remains consistent across groups even when some groups contain fewer
    # categories.  We gather all unique 'Type' values from the detail
    # DataFrames in the report and count them.  If no types are found, fall
    # back to the number of keys in ``_TYPE_COLORS``.
    all_types: set[str] = set()
    for craft_name, payload in report.get("groups", []):
        df_detail = payload.get("detail")
        if isinstance(df_detail, pd.DataFrame) and not df_detail.empty:
            # Coerce to string and drop NA
            try:
                types_list = df_detail["Type"].dropna().astype(str).unique().tolist()
                all_types.update(types_list)
            except Exception:
                pass
    if not all_types:
        all_types = set(_TYPE_COLORS.keys())
    max_type_count = len(all_types)
    # Determine whether to shade rows by work order based on sort_option if
    # shade_by_work_order is not explicitly provided.  Explicitly passing
    # shade_by_work_order overrides this inference.
    if shade_by_work_order is None:
        shade_by_work_order = sort_option == "Work Order # (descending)"

    with PdfPages(buffer) as pdf:
        # Title page
        fig_title = plt.figure(figsize=(11, 8.5))
        # Title page heading
        fig_title.text(0.5, 0.6, "Daily Work Order Log", fontsize=24, ha="center")
        fig_title.text(0.5, 0.5, f"Report for {date_label}", fontsize=16, ha="center")
        plt.axis("off")
        pdf.savefig(fig_title)
        plt.close(fig_title)

        # Overall summary page
        full_detail = report.get("full_detail")
        if isinstance(full_detail, pd.DataFrame) and not full_detail.empty:
            # Sort overall detail if sort_option requires Work Order # sorting.
            # Note: sorting overall summary does not change aggregated counts,
            # but ensures consistency if the summary chart uses the detail.
            df_full_sorted = full_detail.copy()
            if sort_option == "Work Order # (descending)":
                try:
                    df_full_sorted["Work Order #"] = pd.to_numeric(df_full_sorted["Work Order #"], errors="coerce")
                    df_full_sorted = df_full_sorted.sort_values(by="Work Order #", ascending=False)
                except Exception:
                    pass
            else:
                try:
                    df_full_sorted = df_full_sorted.sort_values(by="Name")
                except Exception:
                    pass
            summary_fig = _create_summary_figure(df_full_sorted, "Overall Summary", max_type_count)
            pdf.savefig(summary_fig)
            plt.close(summary_fig)

            # Do not include full-detail pages for the overall summary.  The
            # detail tables will be presented under each craft section.

        # Per-craft pages
        for craft_name, payload in report.get("groups", []):
            df_detail = payload.get("detail")
            if df_detail is None or df_detail.empty:
                continue
            # Sort detail rows according to the selected table sorting option
            df_sorted = df_detail.copy()
            if sort_option == "Work Order # (descending)":
                try:
                    df_sorted["Work Order #"] = pd.to_numeric(df_sorted["Work Order #"], errors="coerce")
                    df_sorted = df_sorted.sort_values(by="Work Order #", ascending=False)
                except Exception:
                    # Fallback to string sorting if numeric conversion fails
                    df_sorted = df_sorted.sort_values(by="Work Order #", ascending=False)
            else:
                try:
                    df_sorted = df_sorted.sort_values(by="Name")
                except Exception:
                    pass
            # Summary page for this craft with fixed bar width
            craft_summary_fig = _create_summary_figure(df_sorted, craft_name, max_type_count)
            pdf.savefig(craft_summary_fig)
            plt.close(craft_summary_fig)
            # Detail table pages for this craft
            craft_detail_figs = _create_detail_table_figures(
                df_sorted,
                craft_name,
                shade_by_work_order=shade_by_work_order,
            )
            for fig in craft_detail_figs:
                pdf.savefig(fig)
                plt.close(fig)
    buffer.seek(0)
    return buffer.read()

# === Mini-dashboard helpers for Streamlit ===
import altair as alt

_TYPE_COLORS = {
    "Break In": "#d62728",
    "Standard Corrective": "#1f77b4",
    "Urgent Corrective": "#ff7f0e",
    "Emergency Order": "#d62728",
    "PM Restore/Replace": "#2ca02c",
    "PM Inspection": "#2ca02c",
    "Follow Up Work Order": "#d4c720",
    "Project": "#9467bd",
}

def _craft_dashboard_block(df_detail: pd.DataFrame) -> None:
    """Render dashboard metrics and chart for a single craft in the Streamlit app."""
    if df_detail is None or df_detail.empty:
        return
    df = df_detail.copy()
    df["Sum of Hours"] = pd.to_numeric(df["Sum of Hours"], errors="coerce").fillna(0.0)
    agg = (
        df.groupby("Type", dropna=False)["Sum of Hours"]
        .sum()
        .reset_index()
        .rename(columns={"Sum of Hours": "hours"})
        .sort_values("hours", ascending=False)
    )
    total = float(agg["hours"].sum())
    if total <= 0:
        total = 0.0
    agg["percent"] = np.where(total > 0, (agg["hours"] / total) * 100.0, 0.0)
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Hours", f"{total:,.2f}")
    top_type = agg.iloc[0]["Type"] if not agg.empty else "-"
    top_pct = agg.iloc[0]["percent"] if not agg.empty else 0.0
    c2.metric("Top Type", f"{top_type}")
    c3.metric("% in Top Type", f"{top_pct:.1f}%")
    # Create a base chart with explicit padding so single bars do not fill the
    # entire width of the chart.  The scale paddings ensure there is space
    # around bars even when only one type is present.  Tooltips show the
    # type and hours when hovering.
    base = alt.Chart(agg).encode(
        x=alt.X(
            "Type:N",
            sort="-y",
            title="Work Order Type",
            scale=alt.Scale(paddingInner=0.3, paddingOuter=0.2),
        ),
        tooltip=[alt.Tooltip("Type:N"), alt.Tooltip("hours:Q", format=".2f")],
    )
    color_scale = alt.Scale(domain=list(_TYPE_COLORS.keys()), range=list(_TYPE_COLORS.values()))
    st.caption("Hours by Work Order Type")
    # Set a fixed bar size so that bars do not automatically stretch across
    # the available width.  Doubling the previous thickness results in a bar
    # size of 50 pixels for a more substantial appearance.
    # Reduce overall chart width by 25% relative to the default container width.  The
    # bar thickness remains unchanged (size=50).  Setting an explicit width
    # property prevents the chart from automatically expanding to fill the
    # available horizontal space.
    chart = base.mark_bar(size=50).encode(
        y=alt.Y("hours:Q", title="Hours"),
        color=alt.Color("Type:N", scale=color_scale),
    )
    # Render the chart inside a narrower column so it does not span the full width.
    # Using a 3:1 column ratio leaves 25% of the space empty, effectively
    # shrinking the x-axis.  The bar thickness is unchanged (size=50).
    chart_col, _ = st.columns([3, 1])
    chart_col.altair_chart(chart, use_container_width=True)
    # Prepare breakdown table showing only hours with exactly two decimal places.  Convert
    # numeric values to strings so Streamlit will left-align them.  This avoids the
    # default right alignment for numeric columns in st.dataframe.
    breakdown_df = agg[["Type", "hours"]].copy()
    # Format Hours with two decimal places as a string
    breakdown_df["Hours"] = breakdown_df["hours"].astype(float).map(lambda x: f"{x:.2f}")
    breakdown_df = breakdown_df[["Type", "Hours"]]
    st.caption("Breakdown (Hours)")
    # Display the breakdown table in a narrower column (75% width) similar to
    # the bar chart above.  This uses a 3:1 column layout to leave 25% of the
    # horizontal space empty.  The table is shown without an index.
    tbl_col, _ = st.columns([3, 5])
    tbl_col.dataframe(breakdown_df, use_container_width=True, hide_index=True)

def _style_breakdown(df: pd.DataFrame):
    """Return a pandas Styler so 'Type' is right-aligned and 'hours' is left-aligned."""
    if df is None or df.empty:
        return df
    try:
        styler = df.style
        # Right-align the 'Type' column
        if "Type" in df.columns:
            styler = styler.set_properties(subset=["Type"], **{"text-align": "right"})
        # Left-align all other numeric columns
        numeric_cols = [c for c in df.columns if c != "Type"]
        if numeric_cols:
            styler = styler.set_properties(subset=numeric_cols, **{"text-align": "left"})
        # Build header styles: the first column (Type) right aligned, others left
        header_styles = []
        for idx, col in enumerate(df.columns):
            align = "right" if col == "Type" else "left"
            header_styles.append({"selector": f"th.col_heading.level0.col{idx}", "props": [("text-align", align)]})
        styler = styler.set_table_styles(header_styles, overwrite=False)
        return styler
    except Exception:
        # Fallback: return the raw DataFrame if something goes wrong with styling
        return df

# === Table styling for Type colors ===
def _style_types(df: pd.DataFrame):
    """Apply background colors based on the Type to the DataFrame in Streamlit."""
    if df is None or df.empty or "Type" not in df.columns:
        return df
    def _style_cell(v):
        color = _TYPE_COLORS.get(str(v), None)
        return f"background-color: {color}; color: white; font-weight: 600;" if color else ""
    try:
        return df.style.applymap(_style_cell, subset=["Type"])
    except Exception:
        # Fallback: return unstyled if Styler isn't supported
        return df

# === Streamlit page configuration and UI ===
# Set the Streamlit page configuration and main title.  The title has been
# updated from "Work Order Reporting App" to reflect its purpose as a daily
# log of work orders.
st.set_page_config(page_title="Daily Work Order Log", layout="wide")
st.title("Daily Work Order Log")

with st.sidebar:
    st.header("Upload file")
    time_file = st.file_uploader("Time on Work Order (.xlsx) – REQUIRED", type=["xlsx"], key="time")

if not time_file:
    st.sidebar.info("⬆️ Upload the **Time on Work Order** export to proceed.")
    st.stop()

try:
    time_df = load_timeworkbook(time_file)
    craft_df = get_craft_order_df()
    addr_df = get_address_book_df()
except Exception as e:
    st.sidebar.error(f"File load error: {e}")
    st.stop()

if "Production Date" not in time_df.columns or time_df["Production Date"].dropna().empty:
    st.sidebar.error("No valid 'Production Date' values found in the Time on Work Order file.")
    st.stop()

dates = sorted(pd.to_datetime(time_df["Production Date"]).dt.date.unique())
date_labels = [datetime.strftime(pd.to_datetime(d), "%m/%d/%Y") for d in dates]
label_to_date = dict(zip(date_labels, dates))
selected_label = st.selectbox("Select Production Date", options=date_labels, index=len(date_labels) - 1)
selected_date = label_to_date[selected_label]

report = prepare_report_data(time_df, addr_df, craft_df, selected_date)

# Sidebar controls for sorting detail tables.  This does not affect the
# summary metrics or charts.  The user can choose to sort by Name (the
# default) or by Work Order # in descending order.
st.sidebar.subheader("Table Sorting")
sort_option = st.sidebar.radio(
    "Sort tables by", options=["Name", "Work Order # (descending)"], index=0
)


# Removed export section that embedded a PDF download button.  The original code
# injected a custom HTML button and JavaScript to capture the dashboard using
# html2canvas and jsPDF, hiding the Streamlit sidebar while taking a snapshot
# and downloading the resulting PNG as a PDF.  This functionality has been
# removed to disable dashboard PDF downloads.


st.markdown(f"### Report for {selected_label}")

col_cfg = {
    "Name": st.column_config.TextColumn("Name", width=200),
    "Work Order #": st.column_config.TextColumn("Work Order #", width=50),
    "Sum of Hours": st.column_config.NumberColumn("Sum of Hours", format="%.2f", width=50),
    "Type": st.column_config.TextColumn("Type", width=200),
    "CostCenter": st.column_config.TextColumn("CC", width=45),
    "Cost Center": st.column_config.TextColumn("CC", width=45),
    "CC": st.column_config.TextColumn("CC", width=45),
    # Reduce the Description column width by 45% (from 300 to 165) and
    # allocate that space to the Problem column.  Other columns remain
    # unchanged.
    "Description": st.column_config.TextColumn("Description", width=165),
    "Problem": st.column_config.TextColumn("Problem", width=555),
}

for craft_name, payload in report["groups"]:
    st.markdown(f"#### {craft_name}")
    df_detail = payload["detail"]
    # Always pass the full detail to the dashboard block (metrics & charts)
    _craft_dashboard_block(df_detail)
    # Sort the table according to the selected option.  Sorting does not
    # impact the metrics or charts above.
    if sort_option == "Name":
        df_display = df_detail.sort_values(by="Name", ascending=True)
    else:
        # Sort by Work Order # descending; ensure the column is numeric if possible
        try:
            df_display = df_detail.copy()
            df_display["Work Order #"] = pd.to_numeric(df_display["Work Order #"], errors="coerce")
            df_display = df_display.sort_values(by="Work Order #", ascending=False)
        except Exception:
            # Fallback: simple string sort descending
            df_display = df_detail.sort_values(by="Work Order #", ascending=False)
    # Prepare a Styler with type-based coloring.  If sorting by Work Order #,
    # compute alternating shading based on groups of identical work order
    # numbers.  The shading toggles when the work order changes.  Only
    # non-Type columns receive shading; the Type column retains its own
    # color coding.
    styler = _style_types(df_display)
    if sort_option == "Work Order # (descending)" and not df_display.empty and "Work Order #" in df_display.columns:
        # Compute shading flags in the sorted order.  Toggle shading each
        # time the value of Work Order # changes.  Use string representation
        # to ensure consistent grouping across numeric and string types.
        shade_flags = []
        prev_val = None
        shade = False
        for val in df_display["Work Order #"].astype(str).tolist():
            if val != prev_val:
                shade = not shade
                prev_val = val
            shade_flags.append(shade)
        # Create a lookup mapping from the DataFrame's index to shading flag
        shade_map = dict(zip(df_display.index, shade_flags))

        def highlight_groups(row: pd.Series) -> list[str]:
            """Return a list of style strings for a row based on the shading flag."""
            apply_shade = shade_map.get(row.name, False)
            styles: list[str] = []
            for col in row.index:
                # Skip shading for the Type column so color coding remains
                if apply_shade and col != "Type":
                    styles.append("background-color: #eaf3ff")
                else:
                    styles.append("")
            return styles

        try:
            styler = styler.apply(highlight_groups, axis=1)
        except Exception:
            # If styling fails, fall back to unshaded display
            pass

    st.dataframe(
        styler,
        use_container_width=True,
        hide_index=True,
        height=_auto_height(df_display),
        column_config=col_cfg,
    )
    st.markdown("---")

# Removed CSV export for filtered detail.  Previously, the application
# created a CSV from the ``full_detail`` DataFrame and exposed it via
# ``st.download_button``.  Removing this block disables the CSV export.

# -----------------------------------------------------------------------------
# Static snapshot export
#
# To enable emailing a snapshot of the current dashboard, we provide an option
# to generate a PDF of the current report using the existing ``build_pdf``
# helper.  When this download button is clicked, the PDF is generated in
# memory and sent to the user as a downloadable file.  This allows the user
# to share the snapshot via email or other channels outside the app.  Note
# that this implementation does not include any built‑in emailing functionality.
pdf_bytes = build_pdf(report, selected_label, sort_option)
st.sidebar.download_button(
    "Download static snapshot (PDF)",
    data=pdf_bytes,
    file_name=f"workorder_snapshot_{selected_label.replace('/', '-')}.pdf",
    mime="application/pdf",
)
