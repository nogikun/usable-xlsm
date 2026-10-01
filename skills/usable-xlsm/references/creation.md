# New XLSM candidates

Use this path only for a **new** workbook. Preserve existing workbooks through
the staged Windows update path. `create` and `build-project` never execute VBA,
overwrite an existing destination, or certify compilation/runtime behavior.

## Preferred: reuse a known working binary project

Extract a seed once and reuse it for repeatable builds:

```text
uv run --project skills/usable-xlsm usable-xlsm extract-project --workbook seed.xlsm --output work/vbaProject.bin
uv run --project skills/usable-xlsm usable-xlsm create --spec workbook.json --vba-project work/vbaProject.bin --output candidate.xlsm --target-os windows
```

The seed's VBA source is reused unchanged. This cannot embed newly edited `.bas`
source. Update the seed through Windows Excel first if the source must change.
Signed seeds are refused by `extract-project`; use the approved signing flow.
Raw binaries must also come from a known, reviewed seed. Never infer trust from
embedding a binary into a workbook you just generated.

Use a small JSON specification (rows, literal cells, widths, Form Control
buttons). Literal strings beginning with `=` remain literal strings:

```json
{
  "vba_name": "ThisWorkbook",
  "sheets": [{
    "name": "入力",
    "vba_name": "Sheet1",
    "rows": [["品名", "数量"], ["サンプル", 2]],
    "cells": {"A5": "ボタンを押してください"},
    "column_widths": {"A:A": 24},
    "buttons": [{"cell": "D2", "macro": "Module1.Hello", "caption": "実行"}]
  }]
}
```

Match workbook/sheet codenames to the seed's document modules. Visible sheet
names are independent of VBA codenames. XlsxWriter supports Form Control
buttons; ActiveX/UserForm designers are not generated here. Check button
procedure names and behavior in the target runtime before delivering a result.

## Experimental: build a binary project from source

The bundled `vba_project.py` implements MS-OVBA compression and CFB writing; do
not improvise another compressor during each request. Opt in explicitly:

```text
uv run --project skills/usable-xlsm usable-xlsm build-project --source work/vba --output work/vbaProject.bin --experimental --target-os windows
uv run --project skills/usable-xlsm usable-xlsm create --spec workbook.json --source work/vba --experimental --output candidate.xlsm --target-os windows
```

The `create --source` path generates document codenames from the specification.
For standalone `build-project`, repeated `--document-name` options replace the
default `ThisWorkbook`, `Sheet1`; put the workbook codename first.

Limits: standard `.bas` modules plus generated empty workbook/sheet document
modules only; ASCII component names of at most 31 characters; source must encode
losslessly as CP932; bounded CFB v3 size (109 FAT sectors). Classes, UserForms,
external references, signatures, compiled caches and workbook event code are
not generated. An incompressible short final chunk is refused rather than
padded/truncated. Modules that need additional type-library references require
an Excel-created seed or explicit manual reference setup.

The binary is re-extracted and source-compared before writing, then the XLSM ZIP
and embedded project are checked. These checks establish structural integrity
only. `runtime_verified` stays false. Open/compile/save a trusted **copy** in
Excel, run actual tests and buttons, and validate the target OS before release.
Windows-worker readiness, trust and audit checks still apply to that candidate;
never bypass a preflight finding because the candidate came from this builder.

Official format and writer references:
- https://xlsxwriter.readthedocs.io/working_with_macros.html
- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/3d07f2c3-dee0-4ae3-b91f-3e32b789c534
- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/ef7087ac-3974-4452-aab2-7dba2214d239
- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/ec9b2c7b-c68e-4246-8d68-7eef2458180b
