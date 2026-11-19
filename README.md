# ppn-try

Python tools for cleaning up Excel templates.

## `clean_template.py`

Makes a cleaned **copy** of a workbook (the original is never touched):

1. Removes hidden sheets that no visible sheet references (old monthly copies).
2. Removes external workbook links that nothing uses anymore.
3. Removes broken / unused defined names (`#REF!`, external, belonging to removed sheets).
4. Drops `calcChain.xml` and orphaned parts (comments, tables, drawings of removed sheets).

It edits the `.xlsx` package directly instead of re-saving through openpyxl, so
tables, threaded comments, data validation and formatting in the kept sheets are
left exactly as they were. Standard library only - no install needed.

```bash
python clean_template.py input/MyFile.xlsx output/MyFile_CLEAN.xlsx
```

Keep a hidden sheet anyway:

```bash
python clean_template.py input/MyFile.xlsx --keep-hidden "BPP (Dec M+3),TPPI (Okt)"
```

Excel files are git-ignored because this repository is public.

## `simplify_setting.py`

Run after `clean_template.py`. Simplifies the **Setting** sheet:

- **Tanda 1**: instead of the wide K1..K6 Kelas/Volume pairs, two small tables:
  `Sheet kilang | Kilang | Status` and one need per row `Sheet kilang | Kelas | Volume | Status`.
  Status shows `✔ OK` or a plain-language warning (sheet not found, volume not a
  column header, missing input, duplicate).
- **Kilang table**: Tabel MM columns are chosen by name from a dropdown instead of typing column numbers.
- Helper columns are hidden. The refinery sheets' tanda-1 formulas are repointed to the new list;
  everything else is unchanged.

```bash
python clean_template.py input/MyFile.xlsx output/MyFile_CLEAN.xlsx
python simplify_setting.py output/MyFile_CLEAN.xlsx output/MyFile_SIMPLE.xlsx --cib-max 600
```

Needs `openpyxl` (`pip install openpyxl`).
