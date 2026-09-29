"""Decode independent executable mappings from existing perf records only."""
import driver as d
import runner
import capture

def main():
 admission=d.read(d.P/'profile-admission.json')
 assert admission['status']=='pass'
 for n,v in admission['inputs'].items():assert d.desc(d.P/n)==v,n
 binary=d.read(d.P/'build-fp.json')['binary']
 assert d.desc(binary['path'])==binary
 label='elf-program-headers-fp'
 result=d.run('fp',label,['readelf','-lW',binary['path']])
 assert result['exit_code']==0
 rows=[]
 for row in d.read(d.P/'profiles.json')['rows']:
  raw=row['raw'];assert d.desc(raw['path'])==raw
  label='mappings-'+row['label'];dest=d.TARGET/(label+'.script')
  result=runner.run('fp',label,['perf','script','--no-inline','--ns','--show-mmap-events','--show-task-events','-i',raw['path']],dest)
  assert result['exit_code']==0
  rows.append(dict(label=row['label'],raw=raw,decoded_plain=d.desc(dest),decoded=capture.pack(dest),receipt=d.desc(d.P/f'commands/{label}/receipt.json')))
 d.write(d.P/'mappings.json',dict(status='pass',rows=rows,binary=binary,elf=d.desc(d.P/'commands/elf-program-headers-fp/output.log'),elf_receipt=d.desc(d.P/'commands/elf-program-headers-fp/receipt.json')))

if __name__=='__main__':main()
