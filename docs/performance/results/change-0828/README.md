# 0828 — aborted PPTX phase profile

The frozen symbol admission failed after five fresh probe quality gates and
two release builds passed. **Zero measurement reports / zero timed samples**
were collected. Production and benchmark runtime are unchanged.

`abort.json` binds the traceback, frozen driver, partial symbol outputs and
separate post-abort static diagnostic. Address-bounded disassembly verifies all
four wrapper functions, but does not replace frozen admission. The original
`decode.py` remains unchanged. Planned capture/readers are retained as prepared
code; they produced no numerical profile result.

See [the report](../../0828-pptx-edit-phase-profile.md), `results-review.md`,
`symbol-failure-review.md`, `quality.json`, `build.json` and `cleanup.json`.
Every offline attempt is retained under `reader-attempts`. The owned target is
removed; no scratch or temporary binary remains.

Replay with `python3 -B docs/performance/results/change-0828/abort_validate.py --final`
and `python3 -B docs/performance/results/change-0828/seal.py --check-head`
from the repository root. Do not rerun the immutable failed batch in place.
