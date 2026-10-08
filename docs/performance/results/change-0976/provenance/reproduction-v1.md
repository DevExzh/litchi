# 0976 bounded reproduction references

This packet records one harmless stdlib affinity control and bounded offline
decoder-transition evidence. Packet preparation runs no control, Office
workload, build, replay, or decoder command.

The bounded v6 plan and bridge self-tests can be checked against the sealed
external source tree without starting perf or a workload:

```text
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0976/tools/prepare_control_plan_0976_v6.py --check-plan /home/zhuhe/litchi-performance-0976/plans/control-plan-0976-v6.json
PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0976/tools/affinity_bridge_0976_v2.py --self-test
```

The historical v6 control command is the exact `command` array in
`plans/control-plan-0976-v6.json`; it was executed once under the root GO and
its outputs are retained externally under
`control/affinity-control-v6/`. Do not rerun it from this packet. The control
receipt and independent audit report 4,685 samples, CPU31 sample placement,
one observed perf leader TID moved from CPU31 to CPU30, raw loss zero, and
terminal lost samples zero. The audit bounds the claim to the observed TID and
15 polling observations; those observations do not prove continuous placement
or a separate reader thread.

The frozen v2 C shim and guard-driver sources, static/dynamic guard helpers,
decoder-empty analyzer, the historical corrected/no-op execution helpers, the
authorized v3 translation helper, and posthoc validators are copied as bounded
replay references. The v3 translated outputs remain external. The checker-only
failure and posthoc-v1 receipts are historical; the posthoc-v2 receipt passes
equal 6,940-sample canonical streams and the bounded cache lifecycle checks.
The final-v2 0973 corrected/noop posthoc-v3 receipt passes retained no-op
identity controls with the known eight historical lost events and no cache
admission. The independent review is read-only and offline-only. No compiled
artifact is an authoritative repository result.

The sealed shim build recipe is:

```text
cc -shared -fPIC -O2 -std=c11 -Wall -Wextra -Werror -pthread -I/home/zhuhe/litchi-performance-0972/header-v2/root/usr/include -Wl,--version-script=/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.map -o /home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.so /home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.c -ldl -pthread
```

The dynamic guard driver and `dladdr` stub use the exact `driver_argv` and
`stub_argv` arrays in `receipts/dynamic-cache-guard-0976-v5.json`. These are
recorded build recipes; packaging does not rerun them. The offline replay
helpers must run in a fresh external output directory and must refuse an
existing receipt/output path. Run the final v3 translation helper first, then
the v2 translation posthoc validator. For the 0973 regression, execute
`replay_final_v2_0973_regression.py` in a fresh namespace to produce the
corrected replay, execute `validate_and_noop_final_v2_0973.py` in a fresh
namespace to produce the no-op output, and then run
`validate_and_noop_final_v2_0973_v3.py` as the posthoc validator. The v1/v2
checker failures are retained as receipts and must not be treated as fresh
outputs or retried against canonical paths.

The large control perf file, replay caches, mmap dumps, decoder logs, shared
objects, executables, and retained prior-capture artifacts remain external.
Their paths, byte counts, SHA-256 values, and regeneration references are
bound in `external-artifact-descriptor-final-v1.json`; the packet manifest
binds every copied file.
