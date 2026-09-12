"""Serial CFB/XLS quality on unchanged measured source."""
import json,re
from run import HERE,TARGET,FOLDER,run,sha,write
COMMANDS = [('cfb-tests', ['cargo', 'test', '--locked', '-p', 'litchi-cfb', '--all-features', '--', '--test-threads=2']), ('cfb-no-default-tests', ['cargo', 'test', '--locked', '-p', 'litchi-cfb', '--no-default-features', '--', '--test-threads=2']), ('xls-tests', ['cargo', 'test', '--locked', '-p', 'litchi-xls', '--all-features', '--', '--test-threads=2']), ('owner-clippy', ['cargo', 'clippy', '--locked', '-p', 'litchi-cfb', '-p', 'litchi-xls', '--all-features', '--lib', '--', '-D', 'warnings']), ('owner-rustdoc', ['cargo', 'doc', '--locked', '-p', 'litchi-cfb', '-p', 'litchi-xls', '--all-features', '--no-deps'])]
if __name__ == '__main__':
    rows=[]
    for name,command in COMMANDS:
        name='check-'+name
        run(name,['env','CARGO_TARGET_DIR='+str(TARGET),'CARGO_BUILD_JOBS=2','CARGO_INCREMENTAL=0','RUSTDOCFLAGS=-D warnings']+command)
        receipt=FOLDER/(name+'.receipt.json')
        count=sum(int(n) for n in re.findall(r'test result: ok\. (\d+) passed;',(FOLDER/(name+'.stdout')).read_text()))
        rows.append(dict(name=name,receipt_sha256=sha(receipt),executed_tests=count))
    write(HERE/'quality-summary.json',dict(status='pass',checks=rows,executed_tests=sum(r['executed_tests'] for r in rows),source_manifest_sha256=sha(FOLDER/'source-manifest.json'),checks_script_sha256=sha(HERE/'checks.py')))
