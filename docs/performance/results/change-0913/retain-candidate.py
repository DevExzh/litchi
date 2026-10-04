"""Retain exact candidate sources for historical replay after adoption or rejection."""
import hashlib
import json
from pathlib import Path

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def desc(f):return {'path':str(f.relative_to(ROOT)),'bytes':f.stat().st_size,'sha256':hashlib.sha256(f.read_bytes()).hexdigest()}
def main():
    assert not (P/'candidate-retention.json').exists()
    manifest=json.loads((P/'source-v1.json').read_text()); inventory={r['path']:r for r in manifest['files']}
    revision=json.loads((P/'supplemental-owned-direct-adversarial-v2.json').read_text())
    rows=[]
    for name in manifest['overlay_paths']+[revision['path']]:
        original=desc(ROOT/name);expected=revision['current'] if name==revision['path'] else inventory[name];assert original==expected
        target=P/'candidate-source'/name;target.parent.mkdir(parents=True,exist_ok=True)
        with target.open('xb') as out:out.write((ROOT/name).read_bytes())
        rows.append({'logical_path':name,'original':original,'retained':desc(target)})
    with (P/'candidate-retention.json').open('x') as out:json.dump({'status':'pass','purpose':'Historical candidate bytes; no dependency on later mutable root sources.','files':rows},out,indent=2);out.write('\n')
    print('Retained five exact candidate source files')
if __name__=='__main__':main()
