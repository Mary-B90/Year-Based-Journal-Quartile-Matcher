#!/usr/bin/env python
# coding: utf-8

# # SJR Quartile Extraction and Merging (1999–2024)
This script reads three SJR files for the years 1999–2024 (Computer Science, Psychology, and Business), extracts the journal name and quartile (Q1–Q4) (and rank, if available), merges them into a single dataset, standardizes journal names to correctly identify duplicates, keeps the quartile information for each journal, and finally saves the sorted output to an Excel file.
# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 1999  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 1999  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 1999  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR1999_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2000  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2000  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2000  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2000_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2001  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2001  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2001  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2001_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2002  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2002  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2002  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2002_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2003  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2003  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2003  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2003_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2004  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2004  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2004  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2004_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2005  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2005  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2005  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2005_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2006  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2006  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2006  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2006_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2007  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2007  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2007  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2007_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2008  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2008  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2008  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2008_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2009  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2009  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2009  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2009_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2010  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2010  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2010  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2010_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2011  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2011  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2011  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2011_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2012  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2012  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2012  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2012_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2013  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2013  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2013  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2013_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2014  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2014  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2014  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2014_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2015  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2015  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2015  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2015_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2016  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2016  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2016  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2016_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2017  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2017  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2017  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2017_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2018  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2018  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2018  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2018_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2019  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2019  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2019  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2019_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2020  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2020  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2020  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2020_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2022  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2022  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2022  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2022_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())



# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "scimagojr 2023  Subject Area - Computer Science.xlsx",
    BASE_DIR / "scimagojr 2023  Subject Area - Psychology.xlsx",
    BASE_DIR / "scimagojr 2023  Subject Area - Business, Management and Accounting.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR2023_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
import csv
from io import StringIO
from pathlib import Path

# ================= 1) PATH =================
BASE_DIR = Path("xxxx")

SJR_FILES = [
    BASE_DIR / "Business, Management and Accounting.xlsx",
    BASE_DIR / "Psychology.xlsx",
    BASE_DIR / "Computer Science.xlsx",
]

OUT_XLSX = BASE_DIR / "SJR__ONLY_3FILES_SORTED_with_QRank.xlsx"

Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}

# ================= 2) UTILITIES =================
def norm_title(x):
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def parse_semicolon_xlsx(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    lines = []
    for _, row in raw.iterrows():
        parts = [str(x) for x in row.tolist() if x not in [None, "nan"]]
        if parts:
            lines.append("".join(parts))

    reader = csv.reader(StringIO("\n".join(lines)), delimiter=";", quotechar='"')
    rows = list(reader)

    if not rows:
        raise ValueError("Empty file")

    header = rows[0]
    data = rows[1:]
    return pd.DataFrame(data, columns=header)

def load_scimago(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        cols = {str(c).lower().strip(): c for c in df.columns}
        title_col = cols.get("title")
        q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
        rank_col = cols.get("rank")

        if title_col and q_col:
            out = df[[title_col, q_col]].copy()
            out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
            if rank_col:
                out["SJR_Rank"] = df[rank_col]
            return out
    except Exception:
        pass

    df = parse_semicolon_xlsx(path)
    cols = {str(c).lower().strip(): c for c in df.columns}

    title_col = cols.get("title")
    q_col = cols.get("sjr best quartile") or cols.get("best quartile") or cols.get("quartile")
    rank_col = cols.get("rank")

    if not title_col or not q_col:
        raise ValueError(f"Title / Quartile not found in {path.name}")

    out = df[[title_col, q_col]].copy()
    out.rename(columns={title_col: "Title", q_col: "Quartile"}, inplace=True)
    if rank_col:
        out["SJR_Rank"] = df[rank_col]

    return out

# ================= 3) LOAD FILES =================
print("📂 Reading ONLY these files:")
for f in SJR_FILES:
    print(" -", f.name)

parts = []
for f in SJR_FILES:
    df = load_scimago(f)
    df["Source_File"] = f.name
    parts.append(df)
    print("✅ Loaded:", f.name, "| rows:", len(df))

sjr = pd.concat(parts, ignore_index=True)

# ================= 4) CLEAN + RANK =================
sjr["Quartile"] = sjr["Quartile"].astype(str).str.replace('"', "").str.strip()
sjr = sjr[sjr["Quartile"].isin(Q_ORDER)].copy()

sjr["Q_Rank"] = sjr["Quartile"].map(Q_ORDER).astype(int)
sjr["Title_Clean"] = sjr["Title"].apply(norm_title)

if "SJR_Rank" in sjr.columns:
    sjr["SJR_Rank_num"] = pd.to_numeric(sjr["SJR_Rank"], errors="coerce")

# keep best quartile per journal
sort_cols = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    sort_cols.append("SJR_Rank_num")

sjr = sjr.sort_values(sort_cols).drop_duplicates("Title_Clean", keep="first")

# ================= 5) FINAL SORT =================
final_sort = ["Q_Rank"]
if "SJR_Rank_num" in sjr.columns:
    final_sort.append("SJR_Rank_num")
final_sort.append("Title")

sjr_sorted = sjr.sort_values(final_sort).reset_index(drop=True)

# ================= 6) SAVE =================
sjr_sorted.to_excel(OUT_XLSX, index=False)

print("—" * 50)
print("✅ DONE")
print("Saved to:", OUT_XLSX)
print("Rows:", len(sjr_sorted))
print("Quartile counts:")
print(sjr_sorted["Quartile"].value_counts())


# In[ ]:


import pandas as pd
import re
from pathlib import Path

# ================= 1) PATHS =================
print("START RUNNING...")

MERGED_PATH = Path("xxxx")
SJR_PATH    = Path("xxxx")

SHEET_NAME = "NO_DUPLICATES_KEPT"

# خروجی (بهتره فایل اصلی رو overwrite نکنی)
OUT_PATH = MERGED_PATH.parent / "xxx"

# ستون خروجی جدید
NEW_COL = "Quartile_Matched"

# ================= 2) HELPERS =================
def norm_title(x):
    """Normalize journal titles for matching."""
    if pd.isna(x):
        return ""
    s = str(x).lower()
    s = s.replace("&", " and ")
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"[’'`]", "", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def find_journal_column(columns):
    """
    Tries to auto-detect the journal/source column.
    Add/remove candidates based on your actual sheet.
    """
    candidates = [
        "journal", "Journal",
        "source title", "Source title", "Source Title",
        "source", "Source",
        "publication", "Publication",
        "journal name", "Journal name", "Journal Name",
    ]
    cols_lower = {c.lower(): c for c in columns}
    for cand in candidates:
        if cand in columns:
            return cand
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None

# ================= 3) CHECK FILES =================
for p in [MERGED_PATH, SJR_PATH]:
    if not p.exists():
        raise FileNotFoundError(f"❌ Missing file: {p}")

# ================= 4) LOAD SJR (REFERENCE) =================
sjr = pd.read_excel(SJR_PATH)

required_cols = {"Title_Clean", "Quartile"}
missing = required_cols - set(sjr.columns)
if missing:
    raise ValueError(f"❌ In SJR file, missing columns: {missing}. Found: {list(sjr.columns)}")

# Build lookup: Title_Clean -> Quartile
sjr_lookup = sjr[["Title_Clean", "Quartile"]].dropna().copy()
sjr_lookup["Title_Clean"] = sjr_lookup["Title_Clean"].astype(str)

# Make a dict for fast mapping
q_map = dict(zip(sjr_lookup["Title_Clean"], sjr_lookup["Quartile"]))

# ================= 5) LOAD MERGED FILE =================
xf = pd.ExcelFile(MERGED_PATH)
if SHEET_NAME not in xf.sheet_names:
    raise ValueError(f"❌ Sheet '{SHEET_NAME}' not found. موجودها: {xf.sheet_names}")

df = pd.read_excel(MERGED_PATH, sheet_name=SHEET_NAME)

# Detect journal column
journal_col = find_journal_column(df.columns)
if not journal_col:
    raise ValueError(
        "❌ I couldn't find the journal column in NO_DUPLICATES_KEPT.\n"
        f"Columns I see are:\n{list(df.columns)}\n\n"
        "Rename your journal column to 'journal' (recommended), or tell me its exact name."
    )

# ================= 6) MATCH QUARTILE =================
df["_journal_clean"] = df[journal_col].apply(norm_title)

# Map quartile; if not found => NOT FOUND
df[NEW_COL] = df["_journal_clean"].map(q_map).fillna("NOT FOUND")

# (Optional) also add numeric Q_Rank if you want
# Q_ORDER = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}
# df["Q_Rank_Matched"] = df[NEW_COL].map(Q_ORDER).fillna("NOT FOUND")

# Drop helper col
df.drop(columns=["_journal_clean"], inplace=True)

# ================= 7) SAVE (COPY ALL SHEETS) =================
with pd.ExcelWriter(OUT_PATH, engine="openpyxl") as writer:
    for sh in xf.sheet_names:
        if sh == SHEET_NAME:
            df.to_excel(writer, sheet_name=sh, index=False)
        else:
            pd.read_excel(MERGED_PATH, sheet_name=sh).to_excel(writer, sheet_name=sh, index=False)

# ================= 8) SUMMARY =================
matched = (df[NEW_COL] != "NOT FOUND").sum()
not_found = (df[NEW_COL] == "NOT FOUND").sum()

print("✅ DONE")
print("Output:", OUT_PATH)
print("Sheet:", SHEET_NAME)
print("Journal column used:", journal_col)
print("Matched:", matched)
print("NOT FOUND:", not_found)
counts = df[NEW_COL].value_counts(dropna=False)

# تضمین اینکه ترتیب همیشه Q1..Q4..NOT FOUND باشه
order = ["Q1", "Q2", "Q3", "Q4", "NOT FOUND"]
print("✅ DONE")
print("Output:", OUT_PATH)
print("Sheet:", SHEET_NAME)
print("Journal column used:", journal_col)
print("-" * 40)
print("Counts by Quartile:")
for k in order:
    print(f"{k}: {int(counts.get(k, 0))}")
print("-" * 40)
print(f"TOTAL ROWS: {len(df)}")


# # Year-Based SJR Quartile Matching and Update

# This script opens the main Excel file second filter.xlsx (sheet: rank filter) and reads the year and journal name from each row. For every year between 1999 and 2023, it locates the corresponding SJR file (SJR{year}_QRank.xlsx) in the 1999–2023 folder. Journal names are normalized (lowercased and spacing standardized) to ensure accurate matching. The script then matches each journal to its quartile (Q1–Q4) from the SJR file of the same year and writes the result into a new column called Quartile_Matched in the main Excel file. Finally, it updates (replaces) the rank filter sheet in the original Excel file.
# 
# In short, if a row has year = 2007, the quartile is taken only from SJR2007_QRank.xlsx, not from any other year.

# In[ ]:


import pandas as pd
import re
from pathlib import Path

# ================= Paths =================
SECOND_FILTER = Path("xxx")
SHEET_NAME = "rank filter"

SJR_DIR = Path("xxx")

# ================= Columns in SECOND FILTER =================
YEAR_COL = "year"
JOURNAL_COL = "journal"
OUT_COL = "Quartile_Matched"

# ================= Columns in SJR files =================
SJR_TITLE_COL = "Title"
SJR_QUARTILE_COL = "Quartile"

# ================= Helpers =================
def norm_journal(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().lower()
    s = re.sub(r"\s+", " ", s)  # collapse multiple spaces
    return s

def find_col(df: pd.DataFrame, target: str) -> str:
    """Find a column name case-insensitively."""
    target_l = target.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == target_l:
            return c
    raise KeyError(f"Column '{target}' not found. Available: {list(df.columns)}")

# ================= Load main sheet =================
df = pd.read_excel(SECOND_FILTER, sheet_name=SHEET_NAME)

YEAR_COL_REAL = find_col(df, YEAR_COL)
JOURNAL_COL_REAL = find_col(df, JOURNAL_COL)

# ensure output column exists
if OUT_COL not in df.columns:
    df[OUT_COL] = pd.NA

# normalize journals once
df["_jnorm"] = df[JOURNAL_COL_REAL].map(norm_journal)

# year to int (Excel sometimes reads as float)
df["_year_int"] = pd.to_numeric(df[YEAR_COL_REAL], errors="coerce").astype("Int64")

# ================= Year-wise matching =================
for year in range(1999, 2024):
    mask = df["_year_int"] == year
    if not mask.any():
        continue

    sjr_file = SJR_DIR / f"SJR{year}_QRank.xlsx"
    if not sjr_file.exists():
        continue

    sjr_df = pd.read_excel(sjr_file)

    # enforce SJR columns (case-insensitive)
    sjr_title_col = find_col(sjr_df, SJR_TITLE_COL)
    sjr_quart_col = find_col(sjr_df, SJR_QUARTILE_COL)

    sjr_df["_jnorm"] = sjr_df[sjr_title_col].map(norm_journal)
    sjr_map = dict(zip(sjr_df["_jnorm"], sjr_df[sjr_quart_col]))

    df.loc[mask, OUT_COL] = df.loc[mask, "_jnorm"].map(sjr_map)

# ================= Cleanup + Save =================
df.drop(columns=["_jnorm", "_year_int"], inplace=True, errors="ignore")

with pd.ExcelWriter(SECOND_FILTER, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

