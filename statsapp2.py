import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
from PIL import Image

# ---------------- Streamlit page config ----------------

st.set_page_config(page_title="Advanced & Overview Stats Organizer", layout="wide")
st.title("Advanced & Overview Stats Organizer")
st.write(
    "Upload a **KenPom-style Advanced Stats CSV** (with jersey + player columns and ORtg, "
    "%Poss, %Shots, eFG%, TS%, OR%, DR%, ARate, TORate, Blk%, Stl%, FC/40, FD/40, FTRate).\n\n"
    "Optionally, upload a **CBB Overview CSV** (tsPct, fg2Pct, fg3Pct, usagePct, pfP40, pfEff, etc.) "
    "to generate a matching OVERVIEW DOCX.\n\n"
    "The app will sort each category, add ranks, color headers with your logo color, "
    "and export Word docs in a two-column format."
)

# ---------------- Helpers ----------------

def clean_player_name(name: str) -> str:
    """
    Cleans KenPom names by removing junk like 'National Rank' etc.
    """
    import re

    if not isinstance(name, str):
        name = str(name)

    s = name.replace("\n", " ")
    # Remove 'National Rank...' (case-insensitive, optional spaces)
    s = re.sub(r"(?i)national\s*rank.*$", "", s)
    # Remove trailing digits like '1' in 'Boozer1'
    s = re.sub(r"\d+$", "", s)
    # Collapse whitespace
    s = " ".join(s.split())
    return s.strip()


def clean_jersey(j):
    """
    Convert jersey values like '1', '01', '1.0', '1.00', ' 1.0' → '1'.
    If it's not numeric, return stripped string.
    """
    try:
        return str(int(float(str(j).strip())))
    except Exception:
        return str(j).strip()


def load_and_clean_kenpom_csv(uploaded_file):
    """
    Loads a KenPom-style advanced stats CSV:
      - First two columns are Unnamed: 0 (jersey), Unnamed: 1 (name)
    """
    df = pd.read_csv(uploaded_file)

    # Clean whitespace from headers
    df.columns = [c.strip() for c in df.columns]

    # Jersey + player
    df["Jersey"] = df["Unnamed: 0"].astype(str).str.strip()
    df["Player"] = df["Unnamed: 1"].astype(str).str.strip()
    df["Player"] = df["Player"].apply(clean_player_name)

    # Drop raw cols
    df = df.drop(columns=["Unnamed: 0", "Unnamed: 1"])

    # Remove category header rows (blank jerseys)
    df = df[df["Jersey"].notna() & (df["Jersey"].str.len() > 0)].copy()

    # Clean jersey
    df["Jersey"] = (
        df["Jersey"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    allowed_columns = [
        "Jersey",
        "Player",
        "ORtg",
        "%Poss",
        "%Shots",
        "eFG%",
        "TS%",
        "OR%",
        "DR%",
        "ARate",
        "TORate",
        "Blk%",
        "Stl%",
        "FC/40",
        "FD/40",
        "FTRate",
    ]

    df = df[[c for c in df.columns if c in allowed_columns]].copy()

    numeric_cols = [c for c in df.columns if c not in ["Jersey", "Player"]]

    for col in numeric_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        df[col] = pd.to_numeric(df[col], errors="coerce")

        # Truncate to 1 decimal place (KenPom style)
        df[col] = df[col].apply(
            lambda x: float(int(x * 10)) / 10 if pd.notna(x) else x
        )

    return df


def load_and_clean_overview_csv(uploaded_file):
    """
    Loads a CBB Analytics-style overview CSV and returns:
      Jersey, Player, and the overview stats used in the layout.
    Expected raw columns include:
      fullName, jerseyNum, tsPct, fgaP40, fg2Pct, fg3Pct, ftPct,
      fga3Rate, usagePct, ftaRate, orbPct, drbPct, stlPct,
      astPct, astTov, astUsage, tovPct, pfP40, pfEff, blkPct
    """
    df = pd.read_csv(uploaded_file)

    # Standardize to our naming
    df = df.rename(
        columns={
            "fullName": "Player",
            "jerseyNum": "Jersey",
        }
    )

    df["Jersey"] = df["Jersey"].apply(clean_jersey)
    df["Player"] = df["Player"].astype(str).str.strip()

    needed_cols = [
        "Jersey",
        "Player",
        "tsPct",
        "fgaP40",
        "fg2Pct",
        "astPct",
        "astTov",
        "tovPct",
        "astUsage",
        "drbPct",
        "blkPct",
        "fg3Pct",
        "ftPct",
        "fga3Rate",
        "usagePct",
        "ftaRate",
        "orbPct",
        "stlPct",
        "pfP40",
        "pfEff",
    ]

    cols_present = [c for c in needed_cols if c in df.columns]
    df = df[cols_present].copy()

    for col in df.columns:
        if col in ["Jersey", "Player"]:
            continue
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def remove_table_borders(table):
    """
    Remove all visible borders from a python-docx table.
    """
    tbl = table._element

    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    tblBorders = tblPr.find(qn("w:tblBorders"))
    if tblBorders is None:
        tblBorders = OxmlElement("w:tblBorders")
        tblPr.append(tblBorders)

    border_names = ["top", "left", "bottom", "right", "insideH", "insideV"]

    for border_name in border_names:
        border = tblBorders.find(qn(f"w:{border_name}"))
        if border is None:
            border = OxmlElement(f"w:{border_name}")
            tblBorders.append(border)
        border.set(qn("w:val"), "nil")


def format_overview_value(col: str, value: float) -> str:
    """
    Convert overview stats into correct display formats.

    Percent logic:
        - If value < 2     → treat as decimal percentage (1.111 → 111.1%)
        - If value ≥ 2     → already percent (48.7 → 48.7%, 140 → 140.0%)

    Ratios:
        - astTov, astUsage, pfEff → two decimals + 'x'

    Per-40:
        - fgaP40, pfP40 → one decimal
    """
    if pd.isna(value):
        return ""

    # Ratio stats
    if col in ["astTov", "astUsage", "pfEff"]:
        return f"{value:.2f}x"

    # Per-40 stats
    if col in ["fgaP40", "pfP40"]:
        return f"{value:.1f}"

    # Percent-like stats
    percent_cols = [
        "tsPct", "fg2Pct", "fg3Pct", "ftPct",
        "fga3Rate", "usagePct", "ftaRate",
        "orbPct", "drbPct", "stlPct", "tovPct",
    ]
    if col in percent_cols:
        # If <2, treat as decimal (0.487 → 48.7%), (1.111 → 111.1%)
        if value < 2:
            value = value * 100
        return f"{value:.1f}%"

    # Fallback
    return f"{value:.1f}"


def get_darkest_color_from_logo(logo_bytes):
    """
    Given raw logo bytes, return an RGBColor corresponding to a dark
    NON-BACKGROUND color in the image.

    - Ignores fully transparent pixels
    - Ignores near-black pixels (background like #000000)
    - Among the remaining pixels, picks the lowest-luminance color
    - If everything is near-black, falls back to including black
    """
    try:
        img = Image.open(BytesIO(logo_bytes)).convert("RGBA")
    except Exception:
        return None

    # Downsize for speed
    img = img.resize((64, 64))

    def find_dark_color(ignore_near_black: bool):
        darkest_rgb = None
        darkest_lum = float("inf")

        for r, g, b, a in img.getdata():
            if a == 0:
                continue  # ignore fully transparent

            # Optionally skip near-black background pixels
            if ignore_near_black and r < 15 and g < 15 and b < 15:
                continue

            # Perceptual luminance
            lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
            if lum < darkest_lum:
                darkest_lum = lum
                darkest_rgb = (r, g, b)

        return darkest_rgb

    # First try: ignore near-black (background)
    darkest_rgb = find_dark_color(ignore_near_black=True)

    # Fallback: if everything was near-black, allow black too
    if darkest_rgb is None:
        darkest_rgb = find_dark_color(ignore_near_black=False)

    if darkest_rgb is None:
        return None

    return RGBColor(darkest_rgb[0], darkest_rgb[1], darkest_rgb[2])


# ---------------- Streamlit UI ----------------

uploaded_file = st.file_uploader(
    "Upload a KenPom-style **Advanced Stats** CSV",
    type=["csv"],
    help="Row 0 must have ORtg, %Poss, %Shots, eFG%, TS%, OR%, DR%, ARate, "
         "TORate, Blk%, Stl%, FC/40, FD/40, FTRate (starting at column 3+).",
)

overview_file = st.file_uploader(
    "Upload a **CBB Overview** CSV (tsPct, fg2Pct, fg3Pct, usagePct, pfP40, pfEff, etc.)",
    type=["csv"],
    key="overview_csv",
    help="This will be formatted into an OVERVIEW DOCX similar to your overview template.",
)

team_name = st.text_input(
    "Team name for DOCX titles",
    value="",
    help="Shown as '<TEAM> ADVANCED STATISTICS' and '<TEAM> OVERVIEW STATISTICS' in the documents.",
)

logo_file = st.file_uploader(
    "Optional team logo (appears at top of both DOCX files; header color matches its darkest non-black color)",
    type=["png", "jpg", "jpeg"],
    key="logo",
)

title_text = (team_name or "").strip().upper()
safe_team_name = (title_text or "TEAM").replace(" ", "_")

# Read logo once, reuse bytes + compute header color
logo_bytes = None
header_color = None

if logo_file is not None:
    logo_bytes = logo_file.read()
    if logo_bytes:
        header_color = get_darkest_color_from_logo(logo_bytes)

# ---------- ADVANCED STATS PIPELINE (Download Button #1) ----------

if uploaded_file is not None:
    try:
        df_stats = load_and_clean_kenpom_csv(uploaded_file)
    except Exception as e:
        st.error(f"Error reading advanced stats CSV: {e}")
        st.stop()

    st.subheader("Parsed Advanced Stats Table")
    st.dataframe(df_stats, use_container_width=True)

    categories = [
        ("ORtg", "Offensive Rating"),
        ("FTRate", "Free-Throw Rate"),

        ("%Poss", "% of Possessions"),
        ("eFG%", "Effective FG%"),

        ("%Shots", "% of Shots"),
        ("TS%", "True Shooting %"),

        ("OR%", "Offensive Rebound %"),
        ("DR%", "Defensive Rebound %"),

        ("TORate", "Turnover Rate"),
        ("ARate", "Assist Rate"),

        ("FD/40", "Fouls Drawn per 40"),
        ("FC/40", "Fouls Committed per 40"),

        ("Blk%", "Block %"),
        ("Stl%", "Steal %"),
    ]

    # Build Advanced DOCX
    doc = Document()

    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)

    # Logo
    if logo_bytes:
        logo_stream = BytesIO(logo_bytes)
        pic = doc.add_picture(logo_stream, width=Inches(1.2))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Title
    title_paragraph = doc.add_paragraph()
    run = title_paragraph.add_run(f"{title_text} ADVANCED STATISTICS")
    run.bold = True
    run.font.size = Pt(14)
    if header_color is not None:
        run.font.color.rgb = header_color
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.space_after = Pt(4)

    # Pair categories into left/right
    pairs = []
    for i in range(0, len(categories), 2):
        if i + 1 < len(categories):
            pairs.append((categories[i], categories[i + 1]))
        else:
            pairs.append((categories[i], None))

    for left_cat, right_cat in pairs:
        table = doc.add_table(rows=1, cols=2)
        table.autofit = True
        remove_table_borders(table)

        left_cell = table.rows[0].cells[0]
        right_cell = table.rows[0].cells[1]

        # LEFT CATEGORY
        if left_cat is not None:
            col, title = left_cat
            df_sorted = (
                df_stats.sort_values(by=col, ascending=False)
                if col in df_stats.columns
                else None
            )

            p = left_cell.add_paragraph()
            r = p.add_run(title)
            r.bold = True
            r.font.size = Pt(16)
            if header_color is not None:
                r.font.color.rgb = header_color
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(11)

            if df_sorted is not None:
                for rank, (_, row) in enumerate(df_sorted.iterrows(), start=1):
                    value = row[col]
                    if pd.isna(value):
                        continue
                    jersey = clean_jersey(row["Jersey"])
                    name = str(row["Player"])

                    if abs(value - int(value)) < 1e-6:
                        val_str = f"{int(value)}"
                    else:
                        val_str = f"{value:.1f}"

                    suffix = "" if col in ["ORtg", "FC/40", "FD/40"] else "%"

                    pl = left_cell.add_paragraph()
                    pr = pl.add_run(f"{rank}. #{jersey} {name} – {val_str}{suffix}")
                    pr.font.size = Pt(11)
                    pl.paragraph_format.space_before = Pt(0)
                    pl.paragraph_format.space_after = Pt(0)
                    pl.paragraph_format.line_spacing = Pt(11)

        # RIGHT CATEGORY
        if right_cat is not None:
            col, title = right_cat
            df_sorted = (
                df_stats.sort_values(by=col, ascending=False)
                if col in df_stats.columns
                else None
            )

            p = right_cell.add_paragraph()
            r = p.add_run(title)
            r.bold = True
            r.font.size = Pt(16)
            if header_color is not None:
                r.font.color.rgb = header_color
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(11)

            if df_sorted is not None:
                for rank, (_, row) in enumerate(df_sorted.iterrows(), start=1):
                    value = row[col]
                    if pd.isna(value):
                        continue
                    jersey = clean_jersey(row["Jersey"])
                    name = str(row["Player"])

                    if abs(value - int(value)) < 1e-6:
                        val_str = f"{int(value)}"
                    else:
                        val_str = f"{value:.1f}"

                    suffix = "" if col in ["ORtg", "FC/40", "FD/40"] else "%"

                    pl = right_cell.add_paragraph()
                    pr = pl.add_run(f"{rank}. #{jersey} {name} – {val_str}{suffix}")
                    pr.font.size = Pt(11)
                    pl.paragraph_format.space_before = Pt(0)
                    pl.paragraph_format.space_after = Pt(0)
                    pl.paragraph_format.line_spacing = Pt(11)

        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(0)
        sp.paragraph_format.line_spacing = Pt(0.25)

    docx_buffer = BytesIO()
    doc.save(docx_buffer)
    docx_buffer.seek(0)

    filename = f"{safe_team_name}_advanced_statistics.docx"

    st.download_button(
        label="Download Advanced Stats DOCX",
        data=docx_buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key="download_advanced",
    )
else:
    st.info("⬆️ Upload an **Advanced Stats** CSV to generate the advanced stats DOCX.")


# ---------- OVERVIEW PIPELINE (Download Button #2) ----------

if overview_file is not None:
    try:
        df_overview = load_and_clean_overview_csv(overview_file)
    except Exception as e:
        st.error(f"Error reading overview CSV: {e}")
        st.stop()

    st.subheader("Parsed Overview Table")
    st.dataframe(df_overview, use_container_width=True)

    # Left column categories
    left_overview = [
        ("tsPct", "True Shooting %"),
        ("fgaP40", "Field Goal Attempts per 40 (FGA/40)"),
        ("fg2Pct", "Two-Point %"),
        ("astPct", "Assist %"),
        ("astTov", "Assist-to-Turnover"),
        ("tovPct", "Turnover %"),
        ("astUsage", "Assist/Usage"),
        ("drbPct", "Defensive Rebound %"),
        ("blkPct", "Block %"),
    ]

    # Right column categories
    right_overview = [
        ("fg3Pct", "Three-Point %"),
        ("ftPct", "Free Throw %"),
        ("fga3Rate", "Three-Point Attempt Rate"),
        ("usagePct", "Usage %"),
        ("ftaRate", "Free Throw Attempt Rate"),
        ("orbPct", "Offensive Rebound %"),
        ("stlPct", "Steal %"),
        ("pfP40", "Personal Fouls per 40"),
        ("pfEff", "PF Efficiency"),
    ]

    overview_pairs = []
    max_len = max(len(left_overview), len(right_overview))
    for i in range(max_len):
        left = left_overview[i] if i < len(left_overview) else None
        right = right_overview[i] if i < len(right_overview) else None
        overview_pairs.append((left, right))

    overview_doc = Document()

    style = overview_doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)

    # Logo again
    if logo_bytes:
        logo_stream = BytesIO(logo_bytes)
        pic = overview_doc.add_picture(logo_stream, width=Inches(1.2))
        last_paragraph = overview_doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    title_paragraph = overview_doc.add_paragraph()
    run = title_paragraph.add_run(f"{title_text} OVERVIEW STATISTICS")
    run.bold = True
    run.font.size = Pt(14)
    if header_color is not None:
        run.font.color.rgb = header_color
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.space_after = Pt(4)

    # Build two-column overview layout with ranks
    for left_cat, right_cat in overview_pairs:
        table = overview_doc.add_table(rows=1, cols=2)
        table.autofit = True
        remove_table_borders(table)

        left_cell = table.rows[0].cells[0]
        right_cell = table.rows[0].cells[1]

        # LEFT
        if left_cat is not None:
            col, title = left_cat
            p = left_cell.add_paragraph()
            r = p.add_run(title)
            r.bold = True
            r.font.size = Pt(16)
            if header_color is not None:
                r.font.color.rgb = header_color
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(11)

            if col in df_overview.columns:
                df_sorted = df_overview.sort_values(by=col, ascending=False)
                for rank, (_, row) in enumerate(df_sorted.iterrows(), start=1):
                    value = row[col]
                    if pd.isna(value):
                        continue
                    jersey = clean_jersey(row["Jersey"])
                    name = str(row["Player"])
                    val_str = format_overview_value(col, value)

                    pl = left_cell.add_paragraph()
                    pr = pl.add_run(f"{rank}. #{jersey} {name} – {val_str}")
                    pr.font.size = Pt(11)
                    pl.paragraph_format.space_before = Pt(0)
                    pl.paragraph_format.space_after = Pt(0)
                    pl.paragraph_format.line_spacing = Pt(11)

        # RIGHT
        if right_cat is not None:
            col, title = right_cat
            p = right_cell.add_paragraph()
            r = p.add_run(title)
            r.bold = True
            r.font.size = Pt(16)
            if header_color is not None:
                r.font.color.rgb = header_color
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(11)

            if col in df_overview.columns:
                df_sorted = df_overview.sort_values(by=col, ascending=False)
                for rank, (_, row) in enumerate(df_sorted.iterrows(), start=1):
                    value = row[col]
                    if pd.isna(value):
                        continue
                    jersey = clean_jersey(row["Jersey"])
                    name = str(row["Player"])
                    val_str = format_overview_value(col, value)

                    pl = right_cell.add_paragraph()
                    pr = pl.add_run(f"{rank}. #{jersey} {name} – {val_str}")
                    pr.font.size = Pt(11)
                    pl.paragraph_format.space_before = Pt(0)
                    pl.paragraph_format.space_after = Pt(0)
                    pl.paragraph_format.line_spacing = Pt(11)

        sp = overview_doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(0)
        sp.paragraph_format.line_spacing = Pt(0.25)

    overview_buffer = BytesIO()
    overview_doc.save(overview_buffer)
    overview_buffer.seek(0)

    overview_filename = f"{safe_team_name}_overview_statistics.docx"

    st.download_button(
        label="Download Overview DOCX",
        data=overview_buffer,
        file_name=overview_filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key="download_overview",
    )

st.markdown(
    "<p style='text-align: center; font-size: 12px; color: gray;'>By Ari Jacobs</p>",
    unsafe_allow_html=True,
)
