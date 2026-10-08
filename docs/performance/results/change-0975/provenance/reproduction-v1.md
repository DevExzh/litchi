# 0975 bounded reproduction references

This packet records a failed native pilot. Packet preparation runs no new
0975 workload, build, capture, retry, or reduction command.

The bounded wrapper checks can be replayed against the sealed external source
tree:

```text
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0975/tools/run_0975.py --self-test
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0975/tools/run_0975.py --check-plan /home/zhuhe/litchi-performance-0975/plans/plan-0975-pre-capture-v3.json
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0975/tools/plan_binding_0975.py --plan /home/zhuhe/litchi-performance-0975/plans/plan-0975-pre-capture-v3.json
```

The launch-boundary helper uses a fresh temporary output path. Do not point
`--output` at the sealed receipt:

```text
replay_dir=$(mktemp -d /tmp/litchi-0975-boundary-replay.XXXXXX)
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0975/tools/launch_boundary_regression_0975_v1.py --output "$replay_dir/launch-boundary-replay.json"
```

Both boundary receipts intercept only `subprocess.Popen`; they report exact
argv construction and cleanup without starting a process, workload, build, or
capture. The historical native-run template was:

```text
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0975/tools/run_0975.py --run --plan /home/zhuhe/litchi-performance-0975/plans/plan-0975-pre-capture-v3.json --go /home/zhuhe/litchi-performance-0975/receipts/root-go-0975-v1.json
```

Do not rerun that command from this packet. The recorded run already stopped
on child 02's nonzero raw loss, with children 03–12 not launched and no retry.
The final independent audit and cleanup receipt are included in the sealed
packet; this historical command is retained for context and must not be rerun.

The large perf data, mmap dumps, packet, decoder logs, retained executable,
and compiled objects remain external. Their path, byte count, SHA-256, and
source receipt are recorded in the external final descriptor [`external-artifact-descriptor-final-v1.json`](external-artifact-descriptor-final-v1.json).
