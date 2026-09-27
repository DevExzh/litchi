"""Independent expanded-name witnesses for every changed generated document.
Expat tokenizes XML with namespace processing off; this checker resolves prefix
scope as text, independently of Litchi's URI IDs and MCE branch selection.
Reserved aliases are compatibility inputs, not namespace-conformance claims.
"""
from pathlib import Path
import json,hashlib,subprocess,xml.parsers.expat,collections
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def witnesses(data):
 parser=xml.parsers.expat.ParserCreate();parser.ordered_attributes=True
 stack=[{'xml':'http://www.w3.org/XML/1998/namespace','xmlns':'http://www.w3.org/2000/xmlns/'}];duplicates=[];unbound=[];events=0
 def start(name,attrs):
  nonlocal events
  events+=1;scope=stack[-1].copy();pairs=list(zip(attrs[::2],attrs[1::2]))
  for key,value in pairs:
   if key=='xmlns':scope['']=value
   elif key.startswith('xmlns:'):scope[key[6:]]=value
  seen={}
  for key,_ in pairs:
   if key=='xmlns' or key.startswith('xmlns:'):continue
   if ':' in key:
    prefix,local=key.split(':',1)
    if prefix not in scope:unbound.append({'event':events,'name':key,'prefix':prefix});continue
    uri=scope[prefix]
   else:uri,local='',key
   expanded=(uri,local)
   if expanded in seen:duplicates.append({'event':events,'first':seen[expanded],'second':key,'namespace':uri,'local':local})
   seen[expanded]=key
  stack.append(scope)
 def end(name):stack.pop()
 parser.StartElementHandler=start;parser.EndElementHandler=end
 error=None
 try:parser.Parse(data,True)
 except xml.parsers.expat.ExpatError as e:error=str(e)
 return {'duplicates':duplicates,'unbound_attributes':unbound,'tokenizer_error':error}
def main():
 d=P/'measure-0';before=read(d/'differential-before.json')['results'];after=read(d/'differential-after.json')['results']
 changes={k:{'before':v,'after':after[k]} for k,v in before.items() if v!=after[k]}
 inputs=P/'changed-inputs';assert not inputs.exists();inputs.mkdir()
 source=(P/'probe-src/main.rs').read_text();start=source.index('fn generated_document(');end=source.index('#[derive(Clone, Copy)]\nenum AliasPairClass',start)
 generator='use std::fmt::Write as _;\nconst MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";\nconst XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";\nconst XMLNS_NAMESPACE: &str = "http://www.w3.org/2000/xmlns/";\n'+source[start:end]+'''\nfn main() -> Result<(), Box<dyn std::error::Error>> {
 let a: Vec<String> = std::env::args().collect();
 let seed: u64 = a[2].parse()?;
 let xml = if a[1] == "aliasing" { aliasing_document(seed) } else { generated_document(seed) };
 std::fs::write(&a[3],xml)?;
 Ok(())
}
'''
 g=P/'generator.rs';g.write_text(generator);binary=Path('/tmp/litchi-0776-generator');assert not binary.exists()
 cmd=['rustc','--edition','2024','-O',str(g),'-o',str(binary)]
 with (P/'generator-build.log').open('w') as f:r=subprocess.run(cmd,stdout=f,stderr=subprocess.STDOUT)
 assert r.returncode==0
 receipt={'command':cmd,'source_sha256':sha(g),'binary_sha256':sha(binary),'log_sha256':sha(P/'generator-build.log')}
 documents={};groups=collections.Counter()
 for key,item in changes.items():
  parts=key.split('::');groups[parts[0]]+=1
  if parts[0]=='alias_pair':
   assert key=='alias_pair::invalid-expanded-duplicate::aliased'
   assert item['after']=='invalid-control:codec_rejected=true:stream_rejected=true';continue
  assert parts[0] in ['aliasing','generated'],('unexpected fixture change',key)
  assert parts[2].startswith('mce_'),('unexpected non-tree change',key)
  doc='::'.join(parts[:2])
  if doc not in documents:
   path=inputs/(parts[0]+'-'+parts[1]+'.xml');subprocess.run([str(binary),parts[0],str(int(parts[1])),str(path)],check=True)
   documents[doc]={'input':str(path.relative_to(P)),'sha256':sha(path),**witnesses(path.read_bytes())}
  witness=documents[doc]
  if item['after']=='err:non-conformant markup compatibility XML: duplicate attribute':assert witness['duplicates'],(key,witness)
  elif item['after'].startswith('err:non-conformant markup compatibility XML: unbound prefix '):assert item['after'][len('err:non-conformant markup compatibility XML: unbound prefix '):] in {x['prefix'] for x in witness['unbound_attributes']},(key,witness)
  else:raise AssertionError(('unclassified behavior change',key,item))
 binary.unlink();receipt['binary_removed']=True
 result={'comparisons':len(before),'changes':changes,'groups':dict(groups),'documents':documents,'generator':receipt,'before_outcomes':dict(collections.Counter('ok' if x['before'].startswith('ok:') else 'err' for x in changes.values()))}
 (P/'triage.json').write_text(json.dumps(result,indent=2)+'\n');print(json.dumps({'changes':len(changes),'groups':dict(groups),'documents':len(documents)}))
if __name__=='__main__':main()
