# 0974 bounded reproduction references

This packet records a prelaunch tooling failure. Packet preparation ran no
0974 workload, build, native capture, retry, or new `--run` command.

The final-v4 review recorded these offline checks against the external sealed
source tree:

```text
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0974/tools/run_0974_v3.py --self-test
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0974/tools/run_0974_v3.py --check-plan /home/zhuhe/litchi-performance-0974/plans/plan-0974-pre-capture-v4.json
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0974/tools/reduce_0974_v4.py --self-test
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0974/tools/reduce_0974_v4.py --admission-self-test --plan /home/zhuhe/litchi-performance-0974/plans/plan-0974-pre-capture-v4.json
```

The v5 boundary regression is reproducible as an offline sentinel test. Use a
fresh temporary directory and a fresh output path; do not point `--output` at
the sealed receipt:

```text
replay_dir=$(mktemp -d /tmp/litchi-0974-boundary-replay.XXXXXX)
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0974/tools/launch_boundary_regression_0974_v1.py --argv-spec /home/zhuhe/litchi-performance-0974/analysis/native-record-spec-v2.json --output "$replay_dir/launch-boundary-replay.json"
```

The authoritative v3 receipt records one `Popen` sentinel call with the
independently derived canonical argv, no process, no perf capture, and restored
patch state. The independent v5 repair review binds that receipt to the v5
helper and plan and reports `offline_repair_pass_not_capture_admission`.

The recorded v4 `--run` attempt failed in command construction with
`RecursionError` before native perf argv construction or `subprocess.Popen`;
child 01's capture directory remained empty, no Office workload started, and
no retry occurred. `root-terminal-0974-v1.json`, the corrective audit, and the
v1 boundary regression receipt bind this outcome. The v5 offline repair does
not supersede the corrective audit's `no_admission` decision and does not
authorize a native rerun.

Large capability probe outputs, lifecycle-control perf data, retained 0973
raw/decoder artifacts, the retained target executable, and all compiled
objects remain external. Their path, byte count, SHA-256, status, and source
receipt are recorded in the external final descriptor [`external-artifact-descriptor-final-v1.json`](external-artifact-descriptor-final-v1.json).
