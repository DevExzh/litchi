# Root verification of the frozen source-profile pair

The retained before/after bundles passed independent receipt verification and
the separate harness review. Each contains 117 process receipts and 2,340
samples. Root rehashed all 4,852 manifest inputs per snapshot against the
retained detached worktrees, checked the three owner/codec source hashes
against the named Git commits, verified metadata identity, recomputed report
statistics, and compared corpus SHA-256 identities across snapshots. Only the
two intended XLSX host source/test paths differ in the normalized manifests.

Reproduce the root check with the frozen checkouts and the retained harness
files at their original relative paths:

```sh
python3 -B root_verify.py \
  --before-root /path/to/before-worktree \
  --after-root /path/to/after-worktree \
  --output /path/to/root-verification.json
```

The original runner hashed its fresh binaries before cleaning its build
targets. Root checked the retained binary-hash/provenance receipts; root did
not independently rebuild or rehash a second live binary. The small-output
lanes count returned errors without matching a specific error variant. The
context checks prove presence and sharing, but do not assert exact namespace
binding counts. These limitations do not invalidate the recorded successful
read/export/edit observations, but must accompany any use of refusal lanes.

The evidence concerns the named source scanner revisions only. It does not
approve the in-progress XLSX lifecycle planner, establish native Office
acceptance, or claim a whole-package speedup. Raw-source retention is one
metric alongside allocator and RSS receipts, not a substitute for them.

The large inputs are namespace-heavy source fixtures. SVG media payload size
and package-level shared-resource targets remain lifecycle work. The measured
README is retained unchanged so its build-manifest hash can be reproduced.
