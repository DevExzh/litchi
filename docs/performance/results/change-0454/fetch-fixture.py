#!/usr/bin/env python3
"""Fetch the pinned, unchanged external QA fixture into exclusive task scratch."""
import base64,hashlib,json,urllib.request
from pathlib import Path
ROOT=Path(__file__).resolve().parent
URL='https://git.libreoffice.org/core/+/39ee11db03368a18bc1746a6312307c4cf1b4bdc/sd/qa/unit/data/smoketest.pptx?format=TEXT'
TARGET=Path('/tmp/litchi-goal-0454-native-copy/external-smoketest.pptx')
assert TARGET.parent.is_dir() and not TARGET.exists()
with urllib.request.urlopen(URL,timeout=30) as response:raw=base64.b64decode(response.read(),validate=True)
assert len(raw)==29956 and hashlib.sha256(raw).hexdigest()=='88a4755fa90815802c8f439c9e0488772e5e7d8db63cfd0326e4d3f35fdeaa44'
blob=hashlib.sha1(b'blob '+str(len(raw)).encode()+b'\0'+raw).hexdigest();assert blob=='e0cfe49009c9d735b5dd6ea774dda2e7a6710ae8'
TARGET.write_bytes(raw)
v={'status':'pass','url':URL,'commit':'39ee11db03368a18bc1746a6312307c4cf1b4bdc','git_blob':blob,'path':str(TARGET),'bytes':len(raw),'sha256':hashlib.sha256(raw).hexdigest(),'driver_sha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),'scope':'Unmodified official LibreOffice QA fixture; original producer/save chain unknown. Fetched artifact remains outside tracked source; no new fixture license or native application claim.'}
with (ROOT/'external-fixture.json').open('x') as f:f.write(json.dumps(v,indent=2)+'\n')
print(json.dumps(v))
