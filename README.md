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
