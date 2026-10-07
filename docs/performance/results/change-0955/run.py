import datetime, hashlib, json, os, pathlib, signal, subprocess, sys, time
root = pathlib.Path('/home/zhuhe/litchi-performance-0955')
label, *command = sys.argv[1:]
record = root / (label + '.json')
log = root / (label + '.log')
if record.exists() or log.exists():
    raise SystemExit('refusing existing receipt/log')
env = dict(os.environ, CARGO_TARGET_DIR=str(root / 'target'), CARGO_BUILD_JOBS='2', CARGO_INCREMENTAL='0', LC_ALL='C', TZ='UTC')
for key in ('RUSTFLAGS', 'RUSTDOCFLAGS', 'CARGO_ENCODED_RUSTFLAGS'):
    env.pop(key, None)
start = time.monotonic()
r = {'command': command, 'cwd': os.getcwd(), 'started_utc': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'environment': {k: env.get(k) for k in ('CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','RUSTFLAGS','RUSTDOCFLAGS','CARGO_ENCODED_RUSTFLAGS','LC_ALL','TZ')}, 'deadline_seconds': 900}
(root / (label + '-start.json')).write_text(json.dumps(r, indent=2) + '\n')
with log.open('xb') as out:
    proc = subprocess.Popen(command, env=env, stdout=out, stderr=subprocess.STDOUT, start_new_session=True)
    try:
        rc = proc.wait(timeout=900)
        r['timeout'] = False
    except subprocess.TimeoutExpired:
        os.killpg(proc.pid, signal.SIGKILL)
        rc = proc.wait()
        r['timeout'] = True
r.update(returncode=rc, seconds=time.monotonic()-start, log={'path':str(log),'bytes':log.stat().st_size,'sha256':hashlib.sha256(log.read_bytes()).hexdigest()})
record.write_text(json.dumps(r,indent=2)+'\n')
print(json.dumps(r))
raise SystemExit(rc if rc >= 0 else 1)
