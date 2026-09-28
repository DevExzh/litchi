# 0819 reader corrections

The independent reader review found three producer-schema mismatches and five
bounded coverage gaps. The unfrozen readers now address them as follows:

- `ordinary_save.corpus.published_sha256` is retained as the single admitted
  reference digest; `ordinary_save.published_sha256` is the per-sample vector.
  Edit phases therefore require an empty vector, while publication phases
  require one admitted digest per sample.
- Procfs controls bind the producer's exact `never_subtracted` token in
  `fixed_32_empty_adjacent_procfs_snapshot_pairs_acquired_before_warmups_and_never_subtracted`.
  Null best-effort procfs deltas remain retained as nulls.
- `atomic_publication_steps` must equal the producer's complete default
  `ATOMIC_STEPS` string, including `sync_all` and parent-directory sync; a
  FileOnly or NoSync description cannot satisfy the check.
- Raw means replay the sorted-sample Rust Welford update. The reader and
  validator use that value with a `1e-12` comparison tolerance rather than
  Python's distinct summation algorithm.
- Phase timing scopes, atomic publication steps, default durability omission,
  format save/sink entry points, and PPTX's `to_bytes` counting boundary are
  checked. Counting phases must retain `sample_byte_split`; other phases must
  omit it.
- Admission selectors bind their format, phase, and input fields. Qualification
  admission, capture commands, scratch markers, completion witnesses, and
  cleanup source/directory rows are replayed.
- Cleanup is omitted from derived `analysis.json`, so pre-cleanup and
  post-cleanup replay has identical bytes. `validate.py --final` requires the
  cleanup witness and replays `seal.json` when it exists; this permits
  `seal.py --write` to create the seal after validation.

Static checks performed without Cargo, binaries, captures, or workloads:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -B - <<'PY'
from pathlib import Path
for name in ("analyze.py", "validate.py"):
    compile(Path("docs/performance/results/change-0819", name).read_text(), name, "exec")
import sys
sys.path.insert(0, "docs/performance/results/change-0819")
import analyze
cv = analyze.load_custody()
assert len(cv["current_source"]["files"]) == 9197
assert len(cv["current_tool"]) == 87
PY
```

The quality process was still live when these corrections were made; no
heavy offline reader execution was performed.

## Offline replay follow-up

Once the capture and quality handles were terminal, the first bounded write
attempt exposed one additional schema mismatch before producing derived
files. The independent artifact audit stores five flat `policy_outputs` rows
(`default`, `full`, `file-only`, `no-sync`, and `stream`); its rows carry their
own `sha256` values and do not repeat the manifest's top-level
`published_sha256`. The reader now binds the complete five-policy set and
digest equality to the manifest output, while retaining the audit's
`source_equals_output` check.

The report producer records the default filesystem envelope as
`["warm", "cold-requested"]`, with fresh child/process isolation and an
owned filesystem root. The reader binds those controls along with each
lane's sample and warmup counts; it leaves unrelated harness configuration
fields available for forward-compatible inspection.

Replay evidence retained in the packet:

- `reader-attempt-0.log`: the initial audit-shape failure;
- `reader-attempt-1.log`: the corrected audit check followed by the
  configuration-shape failure;
- `reader-attempt-2.log`: accepted write of 108 reports and 2,244 samples;
- `reader-check.log`: deterministic replay of all four derived outputs;
- `reader-validation.log`: default validator acceptance.

The successful replay contains 12 native summary rows, eight spread-flagged
rows, and four tail-flagged rows. Static in-memory compilation and derived
cardinality checks pass without Cargo, binary, or capture execution.

Post-cleanup validation exposed a reader ordering mistake: sorting dictionary
representations can order identical descriptors differently when JSON key
ordering differs from construction order. The validator now sorts both exact
descriptor lists by their unique binary path before comparing dictionaries.
No binary identity, cleanup evidence, or raw capture changed. Root retained
`final-validation-attempt-0.log` before this correction.
