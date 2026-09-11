# Contextual source-profile receipts

`before/` and `after/` are sealed raw receipt bundles from detached frozen
worktrees:

| bundle | source snapshot | manifest SHA-256 | binary SHA-256 |
| --- | --- | --- | --- |
| `before/` | `8b5838c591775990747b2cbce82fb2eea372b58b` | `048b6a34a1e00a8978e36fc640bdfcae967497ab26281ad895822c27bcbd7a51` | `1c0df3243a09ec02fc7699d99eccedd7413d4a62ab2eda5fc1392ba20b5aea0d` |
| `after/` | `fc44c4e6c945ab07ded7447f40670d898839eeb3` | `622d4b9bda1e1699cf2387b03eb574ea288dd12b693cef2189c080c4d1198081` | `7060662a0cbd0508ee6ea330df0afaa4575fbe12a8a72a8a719615dbe8a43122` |

Each bundle contains 39 lanes × 3 fresh process receipts, each with 2 warmups
and 20 measured samples, plus `/usr/bin/time -v` output, build provenance,
metadata, manifests, commands, a generated report, and `verification.json`.
The verification files are the authoritative pass receipts from the detached
worktrees; they intentionally say exploratory source-bound evidence.

`comparison.md` is a mechanical side-by-side index of retained raw-source
bytes, context-storage identity, and allocation/time medians. It is a review
aid rather than an optimization claim. Raw JSON and timing files remain the
source of truth.

The scope and bounded observations are described in the parent
`contextual-performance/report.md`; it keeps the scanner evidence separate
from XLSX lifecycle and native acceptance claims.

`normalized-manifest.txt` is the cross-snapshot check. It admits only the two
frozen host source/test file hashes; all other manifest inputs normalize to the
same digest `aacacce9411a6d94bc7211b2f6f511ac2d9109bb4155ca683ecaf48e15539c94`.

To regenerate the index from this checkout:

```text
python3 docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/compare.py \
  --before docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/results/before \
  --after docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/results/after \
  --output docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/results/comparison.md
```

Root can verify copied receipts without the detached source trees with:

```text
python3 docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/verify_receipts.py \
  --results docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/results/before \
  --snapshot before-8b5838c59 \
  --expected-commit 8b5838c591775990747b2cbce82fb2eea372b58b
python3 docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/verify_receipts.py \
  --results docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/results/after \
  --snapshot after-fc44c4e6c \
  --expected-commit fc44c4e6c945ab07ded7447f40670d898839eeb3
```

The detached-worktree command remains the stronger check because the original
`verify.py` also rehashes every manifest input against its frozen source tree:

```text
python3 /var/tmp/litchi-xlsx-contextual-before/docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/verify.py --snapshot before-8b5838c59 --expected-commit 8b5838c591775990747b2cbce82fb2eea372b58b
python3 /var/tmp/litchi-xlsx-contextual-after/docs/report/spec-gap-validation-evidence/xlsx-svg-source/contextual-performance/verify.py --snapshot after-fc44c4e6c --expected-commit fc44c4e6c945ab07ded7447f40670d898839eeb3
```
