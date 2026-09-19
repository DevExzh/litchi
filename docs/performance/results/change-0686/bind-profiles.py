#!/usr/bin/env python3
import hashlib,json
from pathlib import Path
P=Path(__file__).resolve().parent
def sha(f):return hashlib.sha256(f.read_bytes()).hexdigest()
profiles={}
for phase in ['baseline','candidate']:
 command=json.loads((P/(phase+'-profile-command.json')).read_text())
 diagnostic=json.loads((P/'diagnostics'/phase/'manifest.json').read_text())
 assert command['binary_sha256']==diagnostic['binary_sha256']
 profiles[phase]=dict(diagnostic_manifest_sha256=sha(P/'diagnostics'/phase/'manifest.json'),files={f.name:sha(f) for f in sorted(P.glob(phase+'-profile*'))})
(P/'profiles-manifest.json').write_text(json.dumps(profiles,indent=2)+'\n')
