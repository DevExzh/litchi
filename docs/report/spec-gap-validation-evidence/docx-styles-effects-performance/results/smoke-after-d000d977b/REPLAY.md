# Frozen source replay

Verify and fetch the retained source bundle from a repository that already
contains its prerequisite:

```bash
git bundle verify captured-source-16cf1e3c6.bundle
git fetch captured-source-16cf1e3c6.bundle \
  16cf1e3c612a14bd3517cd1b57dd57fce2fb6c35:refs/heads/docx-styles-effects-projection-reuse-16cf1e3c6
git worktree add /var/tmp/docx-styles-effects-projection-reuse-16cf1e3c6 \
  refs/heads/docx-styles-effects-projection-reuse-16cf1e3c6
git -C /var/tmp/docx-styles-effects-projection-reuse-16cf1e3c6 rev-parse HEAD
```

The final command must print
`16cf1e3c612a14bd3517cd1b57dd57fce2fb6c35`. The bundle requires
`d000d977b99e03f8542c7dae74acf767a91b1feb`; it contains the two identity-only
pin commits after that production source. The raw smoke receipts are retained
beside this file. The bundle does not contain the receipts or the disposable
Cargo target; the successful runner removed that target.

The runner command used an external results path and a disjoint target path:

```bash
PROFILE_MODE=smoke PYTHONDONTWRITEBYTECODE=1 TMPDIR=/var/tmp \
DOCX_STYLES_EFFECTS_RESULTS=/var/tmp/litchi-docx-styles-effects-projection-reuse-smoke-results-16cf1e3c6-20260912-b \
DOCX_STYLES_EFFECTS_TARGET_DIR=/var/tmp/litchi-docx-styles-effects-projection-reuse-smoke-target-16cf1e3c6-20260912-b \
bash docs/report/spec-gap-validation-evidence/docx-styles-effects-performance/run_current_smoke.sh
```

That target was removed by the successful runner. Replaying source alone does
not recreate the historical absolute paths embedded in raw receipts. A new
execution requires fresh disjoint output and target paths and must be treated
as a new capture.
