# Profile quiet-gate record: not started

The authorized matched profile was held before launch on 2026-09-12. This is
an operator gate record, not a failed profile run and not performance
evidence.

The clean isolated source checkout was `463334f1a32de687ffb2b3357e55816dcfd3102e`.
Its production/profile source pin is
`d000d977b99e03f8542c7dae74acf767a91b1feb`; its current correctness-smoke
prerequisite is retained at commit
`c4ac516353591f88a9b797349002d05737614e78`, under
`results/smoke-after-d000d977b`.

The planned command was:

```bash
cd /var/tmp/litchi-docx-styles-effects-profile-projection-reuse-d000-c4ac-pins
PROFILE_FROZEN=1 \
DOCX_STYLES_EFFECTS_PROFILE_API_WIRED=1 \
PYTHONDONTWRITEBYTECODE=1 TMPDIR=/var/tmp \
DOCX_STYLES_EFFECTS_PROFILE_RESULTS=/var/tmp/litchi-docx-styles-effects-profile-results-463334f1a-20260912-a \
DOCX_STYLES_EFFECTS_PROFILE_TARGET_DIR=/var/tmp/litchi-docx-styles-effects-profile-target-463334f1a-20260912-a \
bash docs/report/spec-gap-validation-evidence/docx-styles-effects-performance/run_profile.sh
```

The planned matrix is 31 rows, three fresh processes per row, two warmups per
process, and 20 measured samples per process: 186 warmups and 1,860 measured
samples. Neither planned external path existed when this record was written;
no profile process, result receipt, Cargo target, or timing sample was
created.

The one-shot quiet-window census at `2026-09-12T13:33:27Z` observed this
competing workload:

```text
PID 2615682  /home/zhuhe/.rustup/toolchains/1.95.0-x86_64-unknown-linux-gnu/bin/cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline --target-dir /home/zhuhe/litchi-goal-0531-target
```

The profile was therefore not launched. The workload was left untouched. A
future capture requires a fresh census showing no Cargo, rustc, rustdoc, or
other profile/smoke process, followed by explicit root serial authorization.
