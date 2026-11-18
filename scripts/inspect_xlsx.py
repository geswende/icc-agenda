"""Diagnostic : affiche les cellules non vides de chaque onglet."""
import sys, openpyxl
for path in sys.argv[1:]:
    wb = openpyxl.load_workbook(path, data_only=True)
    print(f"########## {path}")
    for ws in wb.worksheets:
        print(f"=== onglet '{ws.title}' dim={ws.dimensions} fusions={len(ws.merged_cells.ranges)}")
        n = 0
        for row in ws.iter_rows(max_row=70):
            for c in row:
                if c.value not in (None, ""):
                    v = str(c.value).replace("\n", " / ")[:70]
                    fill = c.fill.fgColor.rgb if c.fill and c.fill.fill_type else ""
                    print(f"  {c.coordinate}: {v!r} {fill if fill not in ('00000000','') else ''}")
                    n += 1
                    if n > 140: break
            if n > 140: print("  ..."); break
