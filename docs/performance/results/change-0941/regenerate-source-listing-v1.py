"""Recreate the compact source file listing from a retained immutable closure."""
import hashlib,json,pathlib,sys
root=pathlib.Path(sys.argv[1]); output=pathlib.Path(sys.argv[2])
with output.open('x') as stream:
 for p in sorted(root.rglob('*')):
  if p.is_file():
   row={'bytes':p.stat().st_size,'path':p.relative_to(root).as_posix(),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
   stream.write(json.dumps(row,sort_keys=True,separators=(',',':'))+'\n')
