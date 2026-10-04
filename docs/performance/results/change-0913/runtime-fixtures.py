"""Retain and admit native SVG inputs required at runtime by existing tests."""
import hashlib
import json
from pathlib import Path

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
SOURCE=Path('/tmp/litchi-0913-src')

def main():
    names=['3rdparty/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/svg.docx','3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf164835_nonDummyLineHeight.docx']
    names += ['3rdparty/libreoffice-core/sd/qa/unit/data/pptx/'+name for name in ['tdf169496_hidden_graphic.pptx','tdf163852.pptx','tdf164622.pptx']]
    rows=[]
    for name in names:
        data=(ROOT/name).read_bytes();dst=SOURCE/name;assert not dst.exists();dst.parent.mkdir(parents=True,exist_ok=True);dst.write_bytes(data)
        retained=P/'quality-inputs'/Path(name).name
        with retained.open('xb') as out:out.write(data)
        rows.append({'path':name,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest(),'retained':str(retained.relative_to(ROOT)),'origin':'existing local native test fixture'})
    with (P/'runtime-fixtures.json').open('x') as out:json.dump({'status':'pass','added':rows},out,indent=2);out.write('\n')
    print('Admitted',len(rows),'native runtime fixtures')

if __name__=='__main__':main()
