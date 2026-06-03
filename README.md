# Product Similarity Finder

Identifies "like" product numbers within a product list by comparing every entry against others using fuzzy string matching. Results are written to a formatted Excel file showing matched pairs and their similarity scores.

---

## Files

| File | Description |
|---|---|
| `product_similarity_app.py` | GUI application (start here) |
| `product_similarity.py` | Command-line script (also used by the app) |

---

## Requirements

**Python:** `C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python39_64\python.exe`

> The system `python` command is broken on this machine. Always use the full path above.

**Packages** (install once):
```powershell
& "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python39_64\python.exe" -m pip install pandas rapidfuzz xlsxwriter openpyxl
```

---

## Running the Application

### GUI (recommended)

```powershell
cd "C:\Users\ejdiguilio\source\repos\PythonApplication3"
& "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python39_64\python.exe" product_similarity_app.py
```

Settings are saved automatically and restored on next launch.

### Command Line

```powershell
cd "C:\Users\ejdiguilio\source\repos\PythonApplication3"

# Default run (ProductList.xlsx, 95% threshold)
& "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python39_64\python.exe" product_similarity.py --no-header

# Custom threshold
& "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python39_64\python.exe" product_similarity.py --no-header --threshold 90

# Custom window size
& "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python39_64\python.exe" product_similarity.py --no-header --window 300

# Different input file
& "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python39_64\python.exe" product_similarity.py --input "C:\path\to\file.xlsx" --no-header
```

---

## Input File

**Default:** `OneDrive - Vector Security\Desktop\ProductList.xlsx`

| Requirement | Detail |
|---|---|
| Format | `.xlsx` or `.xls` |
| Structure | Single column of product numbers or descriptions |
| Header row | None — check **No header row** in the app or pass `--no-header` on the CLI |

The file currently contains **78,076 unique products**.

---

## Settings

### GUI

| Setting | Default | Description |
|---|---|---|
| File | `ProductList.xlsx` | Path to the input Excel file |
| Sheet | `0` | Sheet index (0 = first) or sheet name |
| No header row | ✓ Checked | Check when the file has no column headers |
| Similarity Threshold | `95` | Minimum score (0–100) to flag two products as "like" |
| Fuzzy Scorer | WRatio | Algorithm used to compute similarity (see below) |
| Window Size | `200` | Neighbors compared per item on large lists |
| Output Folder | *(blank)* | Leave blank to save output alongside the input file |

### CLI flags

| Flag | Default | Description |
|---|---|---|
| `--input` | `ProductList.xlsx` | Path to input file |
| `--sheet` | `0` | Sheet index or name |
| `--no-header` | off | Pass when file has no header row |
| `--threshold` | `95` | Similarity threshold 0–100 |
| `--window` | `200` | Sliding-window size for large lists |
| `--output-dir` | same as input | Folder to write the output file |
| `--log-level` | `INFO` | Console verbosity: `DEBUG`, `INFO`, `WARNING` |

---

## Fuzzy Scorer Guide

The scorer controls how similarity between two product strings is measured. All scores are on a 0–100 scale.

| Scorer | How it works | Best for |
|---|---|---|
| **WRatio** *(default)* | Tries all scorers below and returns the highest | General use — safe choice when you are unsure |
| **ratio** | Character-by-character overlap (Levenshtein) | Exact part numbers where order matters |
| **partial_ratio** | Finds the best-matching substring of the longer string | Part numbers with added prefixes or suffixes (e.g. `ADC-V730` vs `ADC-V730-W`) |
| **token_sort_ratio** | Sorts words alphabetically before comparing | Descriptions where word order varies (e.g. `CAT5 BLACK CABLE` vs `BLACK CABLE CAT5`) |
| **token_set_ratio** | Ignores repeated words; most permissive scorer | Descriptions with extra or duplicate words |

**Tip:** Start with **WRatio** at **95%**. If you are getting too many false positives, switch to **ratio**. If you are missing obvious matches (e.g. same part number with a suffix), try **partial_ratio**.

---

## Output File

Output is saved as a timestamped Excel file:
```
ProductSimilarity_YYYYMMDD_HHMMSS.xlsx
```

Saved to the same folder as the input file by default (Desktop).

### Tab 1 — Similar Pairs

| Column | Description |
|---|---|
| Product A | First product in the matched pair |
| Product B | Second product in the matched pair |
| Similarity Score | Score from 0–100 (higher = more similar) |

- Sorted by score descending
- Rows with score ≥ 99% are highlighted green (near-exact duplicates)
- Auto-filter and frozen header row enabled

### Tab 2 — Summary

Run metadata: date, threshold, scorer used, total unique products, total matches found, and count of near-exact (≥99%) matches.

---

## How the Algorithm Works

### Small lists (≤ 5,000 items)
Full pairwise comparison using a vectorized similarity matrix (`cdist`). Fast and exhaustive.

### Large lists (> 5,000 items) — Sorted Neighborhood Method
At 95%+ similarity, two product strings are nearly identical and will always appear close together when sorted alphabetically. Rather than comparing every possible pair (3 billion pairs for 78K products), the algorithm:

1. Sorts all products alphabetically
2. Compares each item against only its next **N** neighbors (default N = 200)
3. Records any pair scoring at or above the threshold

This reduces comparisons from **3,048,000,000** to **~15,600,000** — a 200× speedup — while capturing all real matches.

**Increasing the window:** If you lower the threshold below 90%, consider increasing `--window` to 400–500 to avoid missing matches that are farther apart in sorted order.

---
