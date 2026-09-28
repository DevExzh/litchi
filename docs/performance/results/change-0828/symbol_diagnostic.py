"""Root-only post-abort static diagnostic; never starts the workload or profiler."""
import re
import subprocess
import time

import custody as c

P = c.P
out = P / 'symbol-diagnostic'
assert not out.exists()
build = c.read(P / 'build.json')
frozen = c.read(build['frozen_inputs']['path'])
c.stable(frozen)
binary = build['binaries']['fp']['artifact']
assert c.artifact(binary['path']) == binary
plan = c.read(P / 'plan.json')
owners = {'edit_region': plan['perf']['owner'], **plan['perf']['phase_owners']}
assert not any((P / name).exists() for name in ('qualification', 'native', 'perf'))
out.mkdir()
receipts = []


def run(label, command):
    stdout = out / (label + '.txt')
    stderr = out / (label + '.log')
    started = time.time()
    with stdout.open('x') as output, stderr.open('x') as error:
        result = subprocess.run(command, cwd=c.ROOT, stdout=output, stderr=error)
    receipts.append({'label': label, 'command': command, 'started': started,
                     'ended': time.time(), 'exit_code': result.returncode,
                     'stdout': c.artifact(stdout), 'stderr': c.artifact(stderr)})
    c.write(out / 'receipts.json', receipts)
    assert result.returncode == 0
    return stdout.read_text()


def rows(text):
    result = []
    pattern = re.compile(r'^\s*([0-9a-fA-F]+)\s+([0-9a-fA-F]+)\s+(\S)\s+(.+)$')
    for line in text.splitlines():
        match = pattern.fullmatch(line)
        if match:
            result.append({'address': int(match[1], 16), 'size': int(match[2], 16),
                           'type': match[3], 'symbol': match[4]})
    return result


raw = rows(run('nm-raw', ['nm', '-S', '--defined-only', binary['path']]))
demangled = rows(run('nm-demangled', ['nm', '-C', '-S', '--defined-only', binary['path']]))
proofs = {}
for label, owner in owners.items():
    found = [row for row in demangled if row['symbol'] == owner]
    assert len(found) == 1 and found[0]['size'] > 0
    row = found[0]
    matching = [item for item in raw if all(item[key] == row[key]
                for key in ('address', 'size', 'type'))]
    assert len(matching) == 1
    start, end = row['address'], row['address'] + row['size']
    command = ['objdump', '--demangle=rust', '--wide', '--line-numbers', '-d',
               f'--start-address={start}', f'--stop-address={end}', binary['path']]
    assembly = run(label, command)
    addresses = [int(match[1], 16) for line in assembly.splitlines()
                 if (match := re.match(r'^\s*([0-9a-fA-F]+):\s', line))]
    assert addresses and addresses[0] == start
    assert all(start <= address < end for address in addresses)
    assert f'<{owner}>:' in assembly
    assert re.search(r'push\s+%rbp', assembly)
    assert re.search(r'mov\s+%rsp,%rbp', assembly)
    assert re.search(r'\bcall\b', assembly)
    proofs[label] = {'owner': owner, 'raw': matching[0], 'demangled': row,
                     'instruction_count': len(addresses),
                     'assembly': c.artifact(out / (label + '.txt')),
                     'frame_pointer_and_call_verified': True}

c.stable(frozen)
c.write(out / 'result.json', {
    'schema': 'litchi.performance.0828.symbol-diagnostic.v1', 'status': 'pass',
    'scope': 'post-abort static diagnostic only; not the frozen symbol admission gate',
    'workload_executed': False, 'capture_authorized': False,
    'binary': binary, 'build': c.artifact(P / 'build.json'),
    'driver': c.artifact(P / 'symbol_diagnostic.py'),
    'receipts': c.artifact(out / 'receipts.json'), 'proofs': proofs,
})
print('0828 post-abort address-bounded symbol diagnostic PASS: four wrappers; zero workloads')
