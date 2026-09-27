"""Root-only serial Heaptrack decode with retained command and input custody."""
import subprocess
import sys
import time

import custody as c

leg = sys.argv[1]
assert leg in ("before", "after")
out = c.P / f"heaptrack-{leg}"
assert not (out / "decode.json").exists()
receipts = c.read(out / "receipts.json")
assert len(receipts) == 1 and receipts[0]["exit_code"] == 0
traces = receipts[0]["traces"]
assert len(traces) == 1
trace = traces[0]
assert c.artifact(trace["path"]) == trace
histogram = out / "histogram"
log = out / "print.log"
assert not histogram.exists() and not log.exists()
command = ["heaptrack_print", "-f", trace["path"], "-H", str(histogram),
           "-n", "15", "-s", "3"]
started = time.time()
with log.open("w") as stream:
    result = subprocess.run(command, cwd=c.ROOT, stdout=stream, stderr=subprocess.STDOUT)
record = {"command": command, "exit_code": result.returncode,
          "started": started, "ended": time.time(), "trace": trace,
          "log": c.artifact(log)}
if histogram.exists():
    record["histogram"] = c.artifact(histogram)
c.write(out / "decode.json", record)
assert c.artifact(trace["path"]) == trace
assert result.returncode == 0, log
assert histogram.is_file()
print(leg, "Heaptrack decode complete", flush=True)
