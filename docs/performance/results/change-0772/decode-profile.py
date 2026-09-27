"""Decode the retained capture using this perf version's symoff field."""
import gzip
import hashlib
import json
import subprocess
from build import P, ROOT, sha

if __name__ == '__main__':
    folder = P / 'profile-qualification'
    capture = json.loads((folder / 'manifest.json').read_text())
    data = folder / 'perf.data'
    assert sha(data) == capture['files'][str(data.relative_to(P))]
    command = ['perf', 'script', '--show-lost-events', '-i', str(data),
               '-F', 'comm,pid,tid,time,event,period,ip,sym,symoff,dso']
    output = folder / 'decoded-stacks.txt.gz'
    digest = hashlib.sha256()
    size = 0
    with (folder / 'decode.stderr').open('w') as err, output.open('wb') as stream:
        with gzip.GzipFile(filename='', mode='wb', fileobj=stream, mtime=0) as zipped:
            process = subprocess.Popen(command, cwd=ROOT, stdout=subprocess.PIPE, stderr=err)
            while chunk := process.stdout.read(1024 * 1024):
                zipped.write(chunk)
                digest.update(chunk)
                size += len(chunk)
            code = process.wait()
    receipt = {'command': command, 'exit': code, 'perf_data_sha256': sha(data),
               'output': str(output.relative_to(P)), 'encoded_sha256': sha(output),
               'decoded_sha256': digest.hexdigest(), 'decoded_bytes': size,
               'stderr_sha256': sha(folder / 'decode.stderr')}
    (folder / 'decode.json').write_text(json.dumps(receipt, indent=2) + '\n')
    assert code == 0, receipt
    print(f'PASS decoded {size} bytes from original capture; no workload rerun')
