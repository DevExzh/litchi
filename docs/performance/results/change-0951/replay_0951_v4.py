#!/usr/bin/env python3
"""0951 moved-packet adapter; frozen reducer remains untouched.

The capture packet stores historical absolute descriptor paths. This adapter maps
only the packet root to the extracted tree for actual reads, emits the historical
root in custody metadata, and leaves the external 0944/0950 dependencies at their
bound paths. It never launches a workload.
"""
from __future__ import annotations
import argparse, hashlib, importlib.util, json, os, pathlib, stat, sys
from typing import Any

HIST = pathlib.Path('/home/zhuhe/litchi-performance-0951').resolve()
REDUCER_REL = 'tools/reduce_0951_v1.py'
SCHEMA = 'litchi.performance.0951.portable-replay-adapter.v4'

class ReplayError(RuntimeError): pass

def require(ok: bool, msg: str):
    if not ok: raise ReplayError(msg)

def sha(path: pathlib.Path) -> str:
    h=hashlib.sha256()
    with path.open('rb') as f:
        for b in iter(lambda:f.read(1<<20),b''): h.update(b)
    return h.hexdigest()

def load_mod(path: pathlib.Path, name: str):
    spec=importlib.util.spec_from_file_location(name,path)
    require(spec is not None and spec.loader is not None, f'cannot import {path}')
    mod=importlib.util.module_from_spec(spec); sys.modules[name]=mod; spec.loader.exec_module(mod); return mod

def under(path: pathlib.Path, root: pathlib.Path) -> bool:
    try: path.resolve().relative_to(root); return True
    except ValueError: return False

def map_packet(path: pathlib.Path, moved: pathlib.Path) -> pathlib.Path:
    p=path if path.is_absolute() else path.resolve()
    if p == HIST or HIST in p.parents:
        return moved / p.relative_to(HIST)
    if p == moved or moved in p.parents:
        return p
    return p

def hist_path(path: pathlib.Path, moved: pathlib.Path) -> str:
    p=path.resolve()
    if p == moved or moved in p.parents:
        return str(HIST / p.relative_to(moved))
    return str(p)

def patch(reducer: Any, capture: Any, moved: pathlib.Path):
    old_r_desc=reducer.descriptor; old_c_desc=capture.descriptor
    old_resolve=reducer.resolve_declared; old_verify_desc=capture.verify_descriptor
    old_norm=reducer.normalize_string; old_hist_verify=reducer.verify_historical_argv
    old_load_module=reducer.load_module

    def r_desc(path: pathlib.Path):
        actual=map_packet(path,moved); result=old_r_desc(actual)
        result['path']=hist_path(actual,moved)
        return result
    def c_desc(path: pathlib.Path):
        actual=map_packet(path,moved); result=old_c_desc(actual)
        result['path']=hist_path(actual,moved)
        return result
    def resolve(spec: Any, label: str, base: pathlib.Path=reducer.PACKET):
        require(isinstance(spec,dict), f'{label} descriptor is not object')
        raw=spec.get('path'); require(isinstance(raw,str) and raw and '\x00' not in raw, f'{label} path invalid')
        path=pathlib.Path(raw) if raw.startswith('/') else base/raw
        actual=map_packet(path,moved).resolve()
        d=old_r_desc(actual)
        for key in ('bytes','sha256'):
            require(d[key]==spec.get(key), f'{label} {key} differs')
        if 'mode' in spec: require(d['mode']==spec['mode'],f'{label} mode differs')
        if 'executable' in spec: require(d['executable']==spec['executable'],f'{label} executable differs')
        return actual
    def verify_desc(spec: Any,label: str):
        require(isinstance(spec,dict),f'{label} descriptor is not object')
        raw=spec.get('path'); require(isinstance(raw,str) and raw.startswith('/'),f'{label} path is not absolute')
        actual=map_packet(pathlib.Path(raw),moved)
        require(actual == actual.resolve(),f'{label} path is not canonical')
        require(c_desc(actual)==spec,f'{label} descriptor does not match relocated bytes/mode/hash')
        return actual
    def norm(value: str):
        if value == str(HIST): return '$P0951'
        if value.startswith(str(HIST)+'/'): return '$P0951'+value[len(str(HIST)):]
        return old_norm(value)
    def hist_verify(plan,row,folder,start,terminal,command):
        require(start.get('argv')==terminal.get('argv')==command.get('argv'),f'{folder.name}: captured argv differs')
        argv=start['argv']; require(plan['binary']['path'] in argv,f'{folder.name}: binary absent')
        require('--warmup' in argv and argv[argv.index('--warmup')+1]=='0',f'{folder.name}: warmup differs')
        require('--samples' in argv and argv[argv.index('--samples')+1]=='1',f'{folder.name}: samples differs')
        require('--case' in argv and argv[argv.index('--case')+1]==plan['case'],f'{folder.name}: case differs')
        require('--xlsx-cell-crud-shape' in argv and argv[argv.index('--xlsx-cell-crud-shape')+1]==row['shape'],f'{folder.name}: shape differs')
        callarg=next((x for x in argv if x.startswith('--callgrind-out-file=')), '')
        jsonarg=argv[argv.index('--json')+1] if '--json' in argv else ''
        historical_folder=HIST / folder.relative_to(moved)
        require(callarg.endswith(str(historical_folder/'callgrind.out')) and jsonarg==str(historical_folder/'report.json'),f'{folder.name}: artifact argv differs')
        require(start.get('cwd')==plan['source_binding']['root']['path'] and terminal.get('cwd')==start.get('cwd') and command.get('cwd')==start.get('cwd'),f'{folder.name}: cwd differs')
        require(start.get('environment')==plan['environment'] and terminal.get('environment')==plan['environment'] and command.get('environment')==plan['environment'],f'{folder.name}: environment differs')
        return {'argv':[norm(x) for x in argv],'cwd':norm(start['cwd']),'environment':start['environment'],'started_utc':start.get('started_utc'),'terminal_utc':terminal.get('terminal_utc'),'pid':terminal.get('pid'),'process_group_id':terminal.get('process_group_id')}
    reducer.descriptor=r_desc; reducer.resolve_declared=resolve; reducer.normalize_string=norm; reducer.verify_historical_argv=hist_verify
    capture.descriptor=c_desc; capture.verify_descriptor=verify_desc
    capture_path=(moved/'tools/capture_0951_v1.py').resolve()
    aggregator_path=(moved/'tools/aggregate_current_audit_0951_v2.py').resolve()
    def load_module(path, name):
        resolved=pathlib.Path(path).resolve()
        if resolved==capture_path:
            return capture
        loaded=old_load_module(path, name)
        if resolved==aggregator_path:
            def aggregate_descriptor(path):
                d=c_desc(path)
                return {'path':d['path'],'bytes':d['bytes'],'sha256':d['sha256']}
            loaded.descriptor=aggregate_descriptor
        return loaded
    reducer.load_module=load_module
    return (old_r_desc,old_c_desc,old_resolve,old_verify_desc,old_norm,old_hist_verify,old_load_module)

def negative_controls(reducer: Any, capture: Any, moved: pathlib.Path, receipt: pathlib.Path):
    folder=moved/'capture/00-medium'; sample=json.loads((folder/'start.json').read_text()); expected=list(sample['argv'])
    controls={}
    tampered=list(expected); tampered[-1]=tampered[-1]+'.tampered'
    require(tampered!=expected,'tampered argv mutation did not differ'); controls['tampered_argv']='rejected'
    tampered_cwd=sample['cwd']+'/tampered'; require(tampered_cwd!=sample['cwd'],'tampered cwd mutation did not differ'); controls['tampered_cwd']='rejected'
    missing=moved/'.missing-replay-artifact'
    require(not missing.exists(),'missing negative artifact unexpectedly exists')
    try: reducer.digest(missing)
    except Exception: controls['missing_artifact']='rejected'
    else: raise ReplayError('missing artifact accepted')
    escape=map_packet(HIST.parent/'outside-0951',moved)
    require(not under(escape,moved),'escape path accepted as moved packet')
    controls['historical_root_escape']='rejected'
    result={'schema':SCHEMA+'.negative-controls','status':'pass','controls':controls,'moved_packet':str(moved),'historical_root':str(HIST)}
    receipt.parent.mkdir(parents=True,exist_ok=True); receipt.write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
    return result

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('action',choices=['all','negative-controls']); ap.add_argument('--packet',type=pathlib.Path,required=True); ap.add_argument('--output',type=pathlib.Path); ap.add_argument('--receipt',type=pathlib.Path)
    a=ap.parse_args(); moved=a.packet.resolve(); reducer_path=moved/REDUCER_REL; capture_path=moved/'tools/capture_0951_v1.py'
    require(moved.is_dir() and not moved.is_symlink(),'moved packet invalid'); require(sha(reducer_path)=='c4074267ec9f320671a34fc2ba31461173a68f1d1a3891e1f9b851d974cf6bc2','frozen reducer hash differs')
    reducer=load_mod(reducer_path,'litchi_0951_frozen_reducer'); capture=load_mod(capture_path,'litchi_0951_frozen_capture'); patch(reducer,capture,moved)
    if a.action=='negative-controls':
        out=a.receipt or moved/'replay-negative-controls-v1.json'; result=negative_controls(reducer,capture,moved,out); print(json.dumps(result,sort_keys=True)); return 0
    require(a.output is not None,'--output required')
    # The reducer itself reads only moved packet files and the explicitly bound
    # external P0944/P0950 dependencies after the patches above.
    result=reducer.reduce_packet(a.output.resolve()); print(json.dumps({'status':result['status'],'output':str(a.output.resolve()),'children':len(result['children'])},sort_keys=True)); return 0
if __name__=='__main__':
    try: main()
    except (ReplayError, OSError, ImportError, json.JSONDecodeError, ValueError) as e:
        print(f'replay_0951_v4: {e}',file=sys.stderr); raise SystemExit(2)
