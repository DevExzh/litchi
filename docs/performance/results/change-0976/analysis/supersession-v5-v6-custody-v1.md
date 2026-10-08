# 0976 v5 to v6 custody supersession

Status: bounded custody note, written after the v6 helper and plan freeze.
No perf process, control, workload, build, or native capture was started.

The v5 JSON artifacts remain byte-preserved at their recorded paths:

* `receipts/helper-seal-0976-v5.json` — SHA-256
  `727aea6589e70fdabefce226e35476dde156fdcc3de699930e2b346e6de4984e`
* `plans/control-plan-0976-v5.json` — SHA-256
  `541e2e4d708421a47ec63660e7cadaa56ab9ba9e3595f2139262a6aa5346dbfd`
* `receipts/preflight-0976-v3.json` — SHA-256
  `710d7e2e12c0dab3272083de5106d9a9fcfd188f592c7c0140185eb3d67b61e7`

Their helper rows recorded these v5 source hashes:

* `tools/affinity_bridge_0976_v2.py` — `51bb0483957ea8e9ba656aa886a554e765aa90ae40ec8c98155165b54faf79b7`
* `tools/stdlib_cpu_burn_0976_v2.py` — `7b7397d015e6ad4b11c15306549766775ca4ed03768cd2f6baa7e6138cc68044`

Those two paths were subsequently corrected for v6 custody and now contain
different bytes. The exact v5 source bytes are not asserted to be recoverable
from the current filesystem. Therefore v5 is historical documentation only;
its plan cannot be revalidated against the current helper paths and must not
be used for execution or as current custody evidence.

The authoritative current artifacts are:

* `receipts/helper-seal-0976-v6.json` — SHA-256
  `1ca27c37ff5ad54d870496249b4015a2f8cddef76c3af1c44f7c1a3ffe7f1e62`
* `plans/control-plan-0976-v6.json` — SHA-256
  `92b40f397df1aa2c3ea8f3cb78e938d9e32d250e97990ffed4478b4da70e9991`
* `receipts/preflight-0976-v4.json` — SHA-256
  `6662670f69ac473d8a76cfdac596f4798f76440ae3753328514c5128aac50542`

The v6 plan binds continuous event mode, absolute `/usr/bin/taskset` and
`/usr/bin/python3.14`, their hashes, the corrected bridge/target bytes, the
control/native gate partition, and the 20 ms observation-only wording.

## Authorized execution form

There is no hidden runner executable. After independent review and a root GO
receipt bind the v6 plan SHA, execute the JSON `command` array from
`plans/control-plan-0976-v6.json` directly with `cwd` set to
`/home/zhuhe/litchi-performance-0976` and the plan's whitelisted environment:

```text
PATH=/usr/bin:/bin
LANG=C
LC_ALL=C
TZ=UTC
PYTHONDONTWRITEBYTECODE=1
LD_PRELOAD absent
```

The command must be launched as an argv array, not through a shell. It starts
with `/usr/bin/taskset --cpu-list 31 /usr/bin/perf record --per-thread`, uses
`instructions:u` period `1000000`, `-m 16M`, `--call-graph dwarf,32768`, and
executes the bound bridge. Preserve stdout/stderr, command/cwd/environment,
terminal status, bridge receipt, target receipt, perf data, and header under
the exclusive v6 output root. Check the v6 plan before launch and stop on any
failure; no retry or replacement output is permitted.
