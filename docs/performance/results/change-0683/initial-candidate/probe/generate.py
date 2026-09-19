#!/usr/bin/env python3
"""Deterministic, marker-free XLSX fixtures; stdlib only, no production writer."""
import hashlib
import json
from pathlib import Path
import sys
import zipfile

SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'
CASES = ['dense', 'sparse', 'long-inline', 'shared-string', 'formula', 'late-fallback', 'late-refusal']

def column(n):
    result = ''
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result

def generate(out, rows=256, columns=256):
    out.mkdir(parents=True, exist_ok=True)
    manifest = []
    for case in CASES:
        parts = {}
        overrides = [('workbook.xml', 'sheet.main+xml'), ('worksheets/sheet1.xml', 'worksheet+xml')]
        if case == 'shared-string':
            overrides.append(('sharedStrings.xml', 'sharedStrings+xml'))
        parts['[Content_Types].xml'] = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + ''.join(f'<Override PartName="/xl/{name}" ContentType="{CT}{kind}"/>' for name, kind in overrides) + '</Types>'
        parts['_rels/.rels'] = f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{OFFICE}/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        parts['xl/workbook.xml'] = f'<workbook xmlns="{SML}" xmlns:r="{OFFICE}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
        relations = f'<Relationship Id="rId1" Type="{OFFICE}/worksheet" Target="worksheets/sheet1.xml"/>'
        if case == 'shared-string':
            relations += f'<Relationship Id="rId2" Type="{OFFICE}/sharedStrings" Target="sharedStrings.xml"/>'
            parts['xl/sharedStrings.xml'] = f'<sst xmlns="{SML}" count="{rows*columns}" uniqueCount="512">' + ''.join(f'<si><t>shared-{i:04}</t></si>' for i in range(512)) + '</sst>'
        parts['xl/_rels/workbook.xml.rels'] = f'<Relationships xmlns="{PKG}">{relations}</Relationships>'
        xml = [f'<worksheet xmlns="{SML}"><dimension ref="A1:{column(columns)}{rows}"/><sheetData>']
        n = k = 0
        for row in range(1, rows + 1):
            xml.append(f'<row r="{row}">')
            cols = range(1, columns+1) if case in ['dense', 'shared-string', 'late-fallback', 'late-refusal'] else ([1, columns] if case == 'sparse' and row % 7 == 0 else ([1, 2] if case == 'formula' else [1]))
            for col in cols:
                n += 1
                address = f'{column(col)}{row}'
                if case == 'shared-string':
                    k += 1
                    xml.append(f'<c r="{address}" t="s"><v>{(row+col)%512}</v></c>')
                elif case == 'long-inline' or (case == 'sparse' and col != 1):
                    text = ('x' * 4096) if case == 'long-inline' else 'sparse'
                    xml.append(f'<c r="{address}" t="inlineStr"><is><t>{text}-{row}</t></is></c>')
                elif case == 'formula' and col == 1:
                    xml.append(f'<c r="{address}"><f>ROW()+{row}</f><v>{row*2}</v></c>')
                else:
                    xml.append(f'<c r="{address}"><v>{row*10000+col}</v></c>')
            xml.append('</row>')
        xml.append('</sheetData>')
        if case == 'late-fallback':
            xml.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        xml.append('<lateTail>' if case == 'late-refusal' else '</worksheet>')
        parts['xl/worksheets/sheet1.xml'] = ''.join(xml)
        (out/f'{case}.xml').write_text(parts['xl/worksheets/sheet1.xml'])
        path = out / f'{case}.xlsx'
        with zipfile.ZipFile(path, 'w') as archive:
            for name, value in parts.items():
                info = zipfile.ZipInfo(name, date_time=(2000, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                archive.writestr(info, value.encode(), compresslevel=6)
        manifest.append(dict(case=case, rows=rows, columns=columns, records=n, shared_references=k, worksheet_bytes=len(parts['xl/worksheets/sheet1.xml'].encode()), archive_bytes=path.stat().st_size, sha256=hashlib.sha256(path.read_bytes()).hexdigest()))
    (out/'manifest.json').write_text(json.dumps(manifest, indent=2)+'\n')

if __name__ == '__main__':
    generate(Path(sys.argv[1]))
