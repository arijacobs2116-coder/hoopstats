import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
from PIL import Image
import re

# ---------------- Streamlit page config ----------------

st.set_page_config(page_title="College Basketball Stats Organizer", layout="wide")
st.title("KenPom Stats & CBBAnalytics Overview Stats Organizer")
st.write(
    "Copy the **advanced stats table directly from KenPom**, paste the raw text below, "
    "and the app will parse it into clean advanced stats.\n\n"
    "Optional: upload a **CBB Overview CSV** (tsPct, fg2Pct, fg3Pct, usagePct, pfP40, pfEff, etc.) "
    "to generate a matching OVERVIEW DOCX.\n\n"
    "The app sorts each category, adds ranks, colors titles and headers with your logo color, "
    "and exports Word docs in a two-column format.\n\n"
    "**Team name and team logo are required before generating any DOCX.**"
)

# ---------------- Helpers ----------------

def clean_player_name(name: str) -> str:
    """
    Cleans KenPom names by removing junk like 'National Rank' etc.
    """
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


def _final_clean_kenpom_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Shared final cleaning logic for KenPom advanced stats DataFrame.
    Assumes df already has 'Jersey' and 'Player' columns populated.
    """
    # Remove category header rows (blank jerseys)
    df = df[df["Jersey"].notna() & (df["Jersey"].str.len() > 0)].copy()

    # Clean jersey formatting
    df["Jersey"] = (
        df["Jersey"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    # Clean player names (remove National Rank junk)
    df["Player"] = df["Player"].apply(clean_player_name)

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


def load_and_clean_kenpom_csv(uploaded_file):
    """
    Fallback: Loads a KenPom-style advanced stats CSV:
      - First two columns are Unnamed: 0 (jersey), Unnamed: 1 (name), etc.
    """
    df = pd.read_csv(uploaded_file)

    # Clean whitespace from headers
    df.columns = [c.strip() for c in df.columns]

    # Jersey + player
    if "Unnamed: 0" in df.columns and "Unnamed: 1" in df.columns:
        df["Jersey"] = df["Unnamed: 0"].astype(str).str.strip()
        df["Player"] = df["Unnamed: 1"].astype(str).str.strip()
        df = df.drop(columns=["Unnamed: 0", "Unnamed: 1"])
    else:
        # Fallback if headers were renamed
        jersey_col = None
        player_col = None
        for c in df.columns:
            lc = c.lower()
            if "jersey" in lc or lc in ("#", "no", "number"):
                jersey_col = c
            if "player" in lc or "name" in lc:
                player_col = c

        if jersey_col is None or player_col is None:
            raise ValueError("Could not find jersey/name columns in advanced stats CSV.")

        df["Jersey"] = df[jersey_col].astype(str).str.strip()
        df["Player"] = df[player_col].astype(str).str.strip()

    df = _final_clean_kenpom_df(df)
    return df


def extract_numbers(line: str):
    """Return list of floatable numbers found in a line."""
    nums = re.findall(r"[-+]?\d*\.?\d+", line)
    out = []
    for n in nums:
        try:
            out.append(float(n))
        except ValueError:
            continue
    return out


def parse_kenpom_paste(raw: str) -> pd.DataFrame:
    """
    Parse raw text copied directly from a KenPom player-usage/advanced page.

    Strategy:
      * Treat each PLAYER as a "block" of lines between jersey/name lines.
      * Inside each block:
          - Identify jersey + name.
          - Find height token (like '6-9'), then weight, year (Fr/So/Jr/Sr/Gr).
          - Starting after the year, locate %Min and ORtg using BOTH patterns:
                A: G S %Min ORtg ...
                B: G S %Min %MinRank ORtg ...
          - Everything after ORtg up to the first shooting-split token (like '27-34')
            is "advanced zone".
          - In the advanced zone we exploit the KenPom rule:
                real stats have a decimal point; ranks are plain ints.
            So we take the first 13 tokens that contain a "." as:
                %Poss, %Shots, eFG%, TS%, OR%, DR%, ARate,
                TORate, Blk%, Stl%, FC/40, FD/40, FTRate.
    """

    # Clean + strip empty lines
    lines = [ln.rstrip() for ln in raw.splitlines()]
    lines = [ln for ln in lines if ln.strip()]

    # Find indices of lines that start a player: "12 Cameron Boozer", "8 Darren Harris ..."
    player_starts = []
    for idx, line in enumerate(lines):
        if re.match(r"^\s*\d+\s+[A-Za-z]", line):
            player_starts.append(idx)

    if not player_starts:
        raise ValueError("No player lines found in pasted text (no lines starting with jersey + name).")

    players = []

    def parse_player_block(block_lines):
        text = " ".join(block_lines)
        tokens = text.split()
        if not tokens:
            return None

        # 1) Jersey
        if not tokens[0].isdigit():
            return None
        jersey = tokens[0]

        # 2) Name: from tokens[1] until we hit HEIGHT (e.g., '6-9') or YEAR (Fr/So/etc)
        name_tokens = []
        height_idx = None
        year_tokens = {"Fr", "So", "Jr", "Sr", "Gr", "Fr.", "So.", "Jr.", "Sr."}

        for i in range(1, len(tokens)):
            t = tokens[i]
            if re.match(r"^\d+-\d+$", t) or t in year_tokens:
                height_idx = i if re.match(r"^\d+-\d+$", t) else None
                break
            else:
                name_tokens.append(t)

        if not name_tokens:
            return None

        # build + clean name (KenPom paste specific)
        name = " ".join(name_tokens)
        name = re.sub(r"(?i)\s*national\s*rank.*$", "", name)  # drop any 'National Rank'
        name = re.sub(r"\s*\d+$", "", name)                     # drop trailing digits like '1'
        name = " ".join(name.split())

        # If we didn't find an explicit height token yet, scan again for it
        if height_idx is None:
            for i in range(len(name_tokens) + 1, len(tokens)):
                t = tokens[i]
                if re.match(r"^\d+-\d+$", t):
                    height_idx = i
                    break

        if height_idx is None or height_idx + 2 >= len(tokens):
            # Can't reliably locate height/weight/year; still keep jersey + name
            return {"Jersey": jersey, "Player": name}

        # 3) Year index (Ht, Wt, Yr)
        yr_idx = height_idx + 2

        # 4) Find %Min and ORtg after YEAR using patterns A and B
        percent_min = None
        ortg = None
        ortg_token_index = None

        # Helper: is this token an int?
        def is_int_token(tok):
            return tok.isdigit()

        # We scan tokens after Yr for a candidate %Min
        for i in range(yr_idx + 1, len(tokens) - 1):
            t_i = tokens[i]
            # candidate %Min must be decimal between 0 and 100
            if "." not in t_i:
                continue
            try:
                v_min = float(t_i)
            except ValueError:
                continue
            if not (0.0 <= v_min <= 100.0):
                continue

            # require at least one or two small ints before it (G, S)
            prev_ok = False
            # one previous int
            if i - 1 >= yr_idx + 1 and is_int_token(tokens[i - 1]) and int(tokens[i - 1]) <= 40:
                prev_ok = True
            # or two previous ints
            if (
                i - 2 >= yr_idx + 1
                and is_int_token(tokens[i - 1])
                and is_int_token(tokens[i - 2])
                and int(tokens[i - 1]) <= 40
                and int(tokens[i - 2]) <= 40
            ):
                prev_ok = True

            if not prev_ok:
                continue

            # ---- Pattern A: next token is ORtg (decimal 50–200) ----
            if i + 1 < len(tokens) and "." in tokens[i + 1]:
                try:
                    v_ortg = float(tokens[i + 1])
                    if 50.0 <= v_ortg <= 200.0:
                        percent_min = v_min
                        ortg = v_ortg
                        ortg_token_index = i + 1
                        break
                except ValueError:
                    pass

            # ---- Pattern B: next is rank (int), then ORtg (decimal 50–200) ----
            if (
                i + 2 < len(tokens)
                and is_int_token(tokens[i + 1])
                and "." in tokens[i + 2]
            ):
                try:
                    v_ortg = float(tokens[i + 2])
                    if 50.0 <= v_ortg <= 200.0:
                        percent_min = v_min
                        ortg = v_ortg
                        ortg_token_index = i + 2
                        break
                except ValueError:
                    pass

        if ortg is None or ortg_token_index is None:
            # couldn't find ORtg cleanly
            return None

        # 5) Advanced zone: from token after ORtg until first shooting-split token
        #    (like '8-9', '11-13', etc.) or until 'FTM-A'/'2PM-A'/'3PM-A'.
        adv_start = ortg_token_index + 1
        adv_end = len(tokens)
        for j in range(adv_start, len(tokens)):
            t = tokens[j]
            if "-" in t or t in ("FTM-A", "2PM-A", "3PM-A"):
                adv_end = j
                break

        adv_tokens = tokens[adv_start:adv_end]

        # KenPom rule: real stats have decimal point; ranks are ints.
        stat_tokens = [t for t in adv_tokens if "." in t]

        if not stat_tokens:
            return {"Jersey": jersey, "Player": name, "ORtg": ortg}

        # First 13 decimal tokens are our advanced stats, in order:
        # %Poss, %Shots, eFG%, TS%, OR%, DR%, ARate, TORate,
        # Blk%, Stl%, FC/40, FD/40, FTRate
        stat_tokens = stat_tokens[:13]
        stat_values = []
        for t in stat_tokens:
            try:
                stat_values.append(float(t.replace("%", "")))
            except ValueError:
                stat_values.append(None)

        while len(stat_values) < 13:
            stat_values.append(None)

        stat_cols_order = [
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

        row = {"Jersey": jersey, "Player": name, "ORtg": ortg}
        for col_name, val in zip(stat_cols_order[1:], stat_values):
            row[col_name] = val

        return row

    # Build blocks and parse each
    for idx, start_idx in enumerate(player_starts):
        end_idx = player_starts[idx + 1] if idx + 1 < len(player_starts) else len(lines)
        block_lines = lines[start_idx:end_idx]
        row = parse_player_block(block_lines)
        if row is not None:
            players.append(row)

    if not players:
        raise ValueError("Could not parse any players from the pasted KenPom text.")

    df = pd.DataFrame(players)
    df = _final_clean_kenpom_df(df)
    return df




def load_and_clean_overview_csv(uploaded_file):
    """
    Loads a CBB Analytics-style overview CSV and returns:
      Jersey, Player, and the overview stats used in the layout.
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
        if value < 2:
            value = value * 100
        return f"{value:.1f}%"

    return f"{value:.1f}"


def get_darkest_color_from_logo(logo_bytes):
    """
    Given raw logo bytes, return an RGBColor corresponding to a dark
    NON-BACKGROUND color in the image.
    """
    try:
        img = Image.open(BytesIO(logo_bytes)).convert("RGBA")
    except Exception:
        return None

    img = img.resize((64, 64))

    def find_dark_color(ignore_near_black: bool):
        darkest_rgb = None
        darkest_lum = float("inf")

        for r, g, b, a in img.getdata():
            if a == 0:
                continue

            if ignore_near_black and r < 15 and g < 15 and b < 15:
                continue

            lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
            if lum < darkest_lum:
                darkest_lum = lum
                darkest_rgb = (r, g, b)

        return darkest_rgb

    darkest_rgb = find_dark_color(ignore_near_black=True)
    if darkest_rgb is None:
        darkest_rgb = find_dark_color(ignore_near_black=False)

    if darkest_rgb is None:
        return None

    return RGBColor(darkest_rgb[0], darkest_rgb[1], darkest_rgb[2])


# ---------------- Streamlit UI ----------------

team_name = st.text_input(
    "Team name for DOCX titles (REQUIRED)",
    value="",
    help="Shown as '<TEAM> ADVANCED STATISTICS' and '<TEAM> OVERVIEW STATISTICS' in the documents.",
)

if team_name.strip() == "":
    st.error("❗ Team name is required.")

logo_file = st.file_uploader(
    "Upload Team Logo (REQUIRED)",
    type=["png", "jpg", "jpeg"],
    key="logo",
    help="Logo will appear at the top of both DOCX files, and its darkest non-black color will be used for titles and headers.",
)


if logo_file is None:
    st.error("❗ Team logo is required.")

title_text = (team_name or "").strip().upper()
safe_team_name = (title_text or "TEAM").replace(" ", "_")

logo_bytes = None
header_color = None
if logo_file is not None:
    logo_bytes = logo_file.read()
    if logo_bytes:
        header_color = get_darkest_color_from_logo(logo_bytes)

# =========================================================
#  ADVANCED STATS INPUT SECTION (KENPOM)
# =========================================================

st.markdown("### Advanced Stats Input (KenPom)")

pasted_kenpom = st.text_area(
    "Paste raw KenPom advanced/usage table here:",
    height=300,
    help="Highlight the whole advanced/usage table on KenPom → Copy → Paste here.",
)

uploaded_file = st.file_uploader(
    "Or upload an Advanced Stats CSV",
    type=["csv"],
    help="CSV must include: ORtg, %Poss, %Shots, eFG%, TS%, OR%, DR%, ARate, TORate, Blk%, Stl%, FC/40, FD/40, FTRate",
)

# ---------------- BLUE NOTE FOR ADVANCED STATS ----------------
st.markdown(
    """
<div style="
    background-color:#0A66C2;
    padding:12px;
    border-radius:6px;
    color:white;
    font-size:15px;
    margin-top:10px;
">
⬆️ Paste KenPom text above or upload an <strong>Advanced Stats CSV</strong> to generate the Advanced Stats DOCX
(after entering team name and logo).
</div>
""",
    unsafe_allow_html=True
)

st.markdown("---")  # Divider between sections


# =========================================================
#  OVERVIEW STATS INPUT SECTION (CBB)
# =========================================================

st.markdown("### Overview Stats Input (CBB)")

overview_file = st.file_uploader(
    "Upload a CBB Overview CSV (tsPct, fg2Pct, fg3Pct, usagePct, pfP40, pfEff, etc.)",
    type=["csv"],
    help="This will be used to generate the Overview DOCX.",
)

# ---------------- BLUE NOTE FOR OVERVIEW STATS ----------------
st.markdown(
    """
<div style="
    background-color:#0A66C2;
    padding:12px;
    border-radius:6px;
    color:white;
    font-size:15px;
    margin-top:10px;
">
⬆️ Upload a <strong>CBB Overview CSV</strong> to generate the Overview DOCX
(after entering team name and logo).
</div>
""",
    unsafe_allow_html=True
)



# ---------- ADVANCED STATS PIPELINE ----------

df_stats = None

if pasted_kenpom.strip():
    try:
        df_stats = parse_kenpom_paste(pasted_kenpom)
    except Exception as e:
        st.error(f"Error parsing pasted KenPom text: {e}")
        df_stats = None
elif uploaded_file is not None:
    try:
        df_stats = load_and_clean_kenpom_csv(uploaded_file)
    except Exception as e:
        st.error(f"Error reading advanced stats CSV: {e}")
        df_stats = None

if df_stats is not None:
    if title_text == "":
        st.error("❗ Please enter a team name before generating the Advanced Stats DOCX.")
    elif logo_bytes is None:
        st.error("❗ Please upload a team logo before generating the Advanced Stats DOCX.")
    else:
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

        doc = Document()
        style = doc.styles["Normal"]
        font = style.font
        font.name = "Calibri"
        font.size = Pt(11)

        if logo_bytes:
            logo_stream = BytesIO(logo_bytes)
            doc.add_picture(logo_stream, width=Inches(1.2))
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        title_paragraph = doc.add_paragraph()
        run = title_paragraph.add_run(f"{title_text} ADVANCED STATISTICS")
        run.bold = True
        run.font.size = Pt(14)
        if header_color is not None:
            run.font.color.rgb = header_color
        title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_paragraph.paragraph_format.space_after = Pt(4)

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
    st.info("⬆️ Paste KenPom text above or upload an **Advanced Stats** CSV to generate the Advanced Stats DOCX (after entering team name and logo).")


# ---------- OVERVIEW PIPELINE ----------

if overview_file is not None:
    if title_text == "":
        st.error("❗ Please enter a team name before generating the Overview DOCX.")
    elif logo_bytes is None:
        st.error("❗ Please upload a team logo before generating the Overview DOCX.")
    else:
        try:
            df_overview = load_and_clean_overview_csv(overview_file)
        except Exception as e:
            st.error(f"Error reading overview CSV: {e}")
            st.stop()

        st.subheader("Parsed Overview Table")
        st.dataframe(df_overview, use_container_width=True)

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

        if logo_bytes:
            logo_stream = BytesIO(logo_bytes)
            overview_doc.add_picture(logo_stream, width=Inches(1.2))
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