import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO

# ---------------- Streamlit page config ----------------

st.set_page_config(page_title="KenPom Advanced Stats Organizer", layout="wide")
st.title("KenPom Advanced Stats Organizer")
st.write(
    "Upload a **CSV** exported from your spreadsheet with jersey, name, and an advanced "
    "stats header row (ORtg, %Poss, %Shots, eFG%, TS%, OR%, DR%, ARate, TORate, Blk%, "
    "Stl%, FC/40, FD/40, FTRate). The app will organize it into category rankings, "
    "show a preview, and export a Word (.docx) file in your preferred format."
)

# ---------------- Helpers ----------------


def load_and_clean_kenpom_csv(uploaded_file):
    """
    Loads a normal CSV that already has proper headers in row 0.
    This version does NOT shift rows down.
    """

    df = pd.read_csv(uploaded_file)

    # Clean whitespace
    df.columns = [c.strip() for c in df.columns]
    df["Jersey"] = df["Unnamed: 0"].astype(str).str.strip()
    df["Player"] = df["Unnamed: 1"].astype(str).str.strip()

    df = df.drop(columns=["Unnamed: 0", "Unnamed: 1"])

    # Convert numeric columns
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

    return df


def clean_jersey(j):
    """
    Convert jersey values like '1', '01', '1.0', '1.00', ' 1.0' → '1'.
    If it's not numeric, return stripped string.
    """
    try:
        return str(int(float(str(j).strip())))
    except Exception:
        return str(j).strip()


def format_category_block(df: pd.DataFrame, col: str, title: str) -> str:
    """
    Markdown block preview for one category, sorted descending.
    Uses clean_jersey() so jersey numbers don't show as 1.0.
    """
    if col not in df.columns:
        return f"**{title}**\n(Stat '{col}' not found in CSV.)"

    df_sorted = df.sort_values(by=col, ascending=False)
    lines = [f"**{title}**"]
    for _, row in df_sorted.iterrows():
        value = row[col]
        if pd.isna(value):
            continue
        jersey = clean_jersey(row["Jersey"])
        name = str(row["Player"])

        if abs(value - int(value)) < 1e-6:
            val_str = f"{int(value)}"
        else:
            val_str = f"{value:.1f}"

        if col in ["ORtg", "FC/40", "FD/40"]:
            suffix = ""
        else:
            suffix = "%"

        lines.append(f"#{jersey} {name} – {val_str}{suffix}")

    return "\n".join(lines)


def remove_table_borders(table):
    """
    Remove all visible borders from a python-docx table.
    Works on all python-docx versions (avoids tblBorders attribute errors).
    """
    tbl = table._element

    # Create <w:tblPr> if it doesn't exist
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    # Create or find <w:tblBorders>
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


# ---------------- Streamlit UI ----------------

logo_file = st.file_uploader(
    "Optional team logo (will appear at top of DOCX)",
    type=["png", "jpg", "jpeg"],
    key="logo",
)

uploaded_file = st.file_uploader(
    "Upload a KenPom-style CSV",
    type=["csv"],
    help="Use a CSV where row 0 has ORtg, %Poss, %Shots, eFG%, TS%, OR%, DR%, ARate, "
         "TORate, Blk%, Stl%, FC/40, FD/40, FTRate in columns 3+.",
)

team_name = st.text_input(
    "Team name for DOCX title",
    value="",
    help="This will appear at the top of the Word document as '<TEAM> ADVANCED STATISTICS'.",
)

if uploaded_file is not None:
    try:
        df_stats = load_and_clean_kenpom_csv(uploaded_file)
    except Exception as e:
        st.error(f"Error reading CSV: {e}")
        st.stop()

    st.subheader("Parsed Advanced Stats Table")
    st.dataframe(df_stats, use_container_width=True)

    # Category mappings: column name -> pretty title
    # Ordered to match your left/right layout screenshot
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

    # ---------- On-page markdown preview ----------
    st.subheader("Organized Advanced Stats (Preview)")
    combined_blocks = []
    for col, title in categories:
        block = format_category_block(df_stats, col, title)
        st.markdown(block)
        st.markdown("---")
        combined_blocks.append(block)

    # Optional TXT download of all categories
    all_text = "\n\n".join(combined_blocks)
    st.download_button(
        label="Download Organized Stats as .txt",
        data=all_text,
        file_name="organized_advanced_stats.txt",
        mime="text/plain",
    )

    # ---------- Build DOCX (logo + two-column layout, tight vertical spacing) ----------
    doc = Document()

    # Base font
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)

    title_text = (team_name or "").strip().upper()

    # Optional logo at the top (smaller width)
    if logo_file is not None:
        logo_bytes = logo_file.read()
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
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.space_after = Pt(4)

    # Pair up categories into (left, right)
    pairs = []
    for i in range(0, len(categories), 2):
        if i + 1 < len(categories):
            pairs.append((categories[i], categories[i + 1]))
        else:
            pairs.append((categories[i], None))

    for left_cat, right_cat in pairs:
        # 1-row, 2-column table for each pair
        table = doc.add_table(rows=1, cols=2)
        table.autofit = True
        remove_table_borders(table)

        left_cell = table.rows[0].cells[0]
        right_cell = table.rows[0].cells[1]

        # ---- LEFT CATEGORY ----
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
            r.font.size = Pt(16)   # bigger section title
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(11)

            if df_sorted is not None:
                for _, row in df_sorted.iterrows():
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
                    pr = pl.add_run(f"#{jersey} {name} – {val_str}{suffix}")
                    pr.font.size = Pt(11)
                    pl.paragraph_format.space_before = Pt(0)
                    pl.paragraph_format.space_after = Pt(0)
                    pl.paragraph_format.line_spacing = Pt(11)   # tighter vertical spacing

        # ---- RIGHT CATEGORY ----
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
            r.font.size = Pt(16)   # bigger section title
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(11)

            if df_sorted is not None:
                for _, row in df_sorted.iterrows():
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
                    pr = pl.add_run(f"#{jersey} {name} – {val_str}{suffix}")
                    pr.font.size = Pt(11)
                    pl.paragraph_format.space_before = Pt(0)
                    pl.paragraph_format.space_after = Pt(0)
                    pl.paragraph_format.line_spacing = Pt(11)   # tighter vertical spacing

        # Minimal spacer between each pair of categories
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(0)
        sp.paragraph_format.line_spacing = Pt(0.25)

    # Save to in-memory buffer and expose download button
    docx_buffer = BytesIO()
    doc.save(docx_buffer)
    docx_buffer.seek(0)

    safe_team_name = (title_text or "TEAM").replace(" ", "_")
    filename = f"{safe_team_name}_advanced_statistics.docx"

    st.download_button(
        label="Download Advanced Stats DOCX",
        data=docx_buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

else:
    st.info("⬆️ Upload a CSV exported from your KenPom/advanced stats sheet to begin.")

