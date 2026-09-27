"""Record release resolver machine code from retained before/after executables."""
import subprocess
import custody as c

out=c.P/'assembly'
assert not out.exists()
out.mkdir()
rows=[]
for leg in ('before','after'):
    build=c.read(c.P/f'build-{leg}/build.json')
    binary=build['binaries']['native']
    assert c.artifact(binary['path'])==binary
    command=['nm','--defined-only',binary['path']]
    result=subprocess.run(command,text=True,capture_output=True,check=True)
    symbols=[line for line in result.stdout.splitlines() if 'notes8resolved' in line or 'notes15known_namespace' in line]
    (out/f'{leg}.symbols.txt').write_text('\n'.join(symbols)+'\n')
    records=[]
    for index,line in enumerate(symbols):
        symbol=line.split()[-1]
        cmd=['objdump','-d','--no-show-raw-insn',f'--disassemble={symbol}',binary['path']]
        result=subprocess.run(cmd,text=True,capture_output=True,check=True)
        path=out/f'{leg}-{index}.asm.txt';path.write_text(result.stdout)
        records.append({'symbol':symbol,'command':cmd,'assembly':c.artifact(path)})
    rows.append({'leg':leg,'binary':binary,'symbol_command':command,'symbols':c.artifact(out/f'{leg}.symbols.txt'),'records':records})
c.write(out/'receipt.json',rows)
print('Recorded release assembly for both legs')
