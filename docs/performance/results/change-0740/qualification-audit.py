"""Replay all qualification report projections, including failed-profile reports."""
import json
from contract import P,read,projection
if __name__=='__main__':
    rows=[]
    for folder in ['qualification','qualification-uncompressed','qualification-wide','qualification-fp']:
        receipts=read(folder+'/manifest.json')
        for i,r in enumerate(receipts):
            row={'case':r['case'],'samples':r['samples'],'warmups':0,'lane':'native'}
            if folder=='qualification-fp':row['build']='fp'
            assert projection(read(f'{folder}/{i:02d}.json'),row)==read('oracle.json')[r['case']]
            rows.append({'folder':folder,'index':i,'mode':r['mode'],'projection':'passed'})
    assert len(rows)==13
    (P/'qualification-audit.json').write_text(json.dumps({'status':'passed','rows':rows,'scope':'logical report projection only; does not validate corrupt or unqualified profile attribution'},indent=2)+'\n')
    print('PASS 13 qualification report projections; failed profile dispositions unchanged')
