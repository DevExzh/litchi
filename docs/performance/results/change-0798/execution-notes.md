# 0798 execution notes

This packet retains four setup failures before the diagnostic capture lanes.
The failures are evidence about the harness setup and are kept separate from
the later control/instrumented results. Production source was preserved during
all of these attempts.

The first lock-generation attempt failed during offline dependency resolution.
The probe feature incorrectly forwarded `litchi-opc/attribute-census` instead
of enabling the probe's optional `litchi-opc` dependency. The recorder did not
retain the numeric process exit status, so the retained evidence preserves
`original_returncode: null` and `returncode_recorded: false`; it must not be
rewritten as an ordinary nonzero exit. The failed manifest, template, lock,
and error log remain under `lock-generation-failed-0/`. A later corrected lock
generation is retained separately in `lock-generation.json` and
`probe-src/Cargo.lock`.

The first isolated hook-quality attempt completed lock generation and failed
the test compilation with E0618. A local variable named `tag` shadowed the
diagnostic test helper function on its second use. Its exact logs and mirror
source are retained under `quality-failed-0/` and
`hook-test-src-failed-0/`. The next attempt passed the tests and failed Clippy
on the diagnostic `Drop` hook's needless borrow. That exact source and log are
retained under `quality-failed-1/` and `hook-test-src-failed-1/`. The final
quality attempt passed the tests and Clippy gates. The canonical-copy
filesystem test is intentionally skipped because this mirror contains only the
OPC owner; the shared canonical test source is still archived.

The first control-probe build passed rustfmt, release build, and release check,
then failed release Clippy. The failure was confined to inherited diagnostic
allocation helpers that are unused in this probe and to a lazy doc-list
continuation in the inherited probe header. The retained relocation archive
contains the exact pre-fix probe source and all four logs. The fix was limited
to a scoped `dead_code` allowance on the inherited allocation module and a
blank documentation paragraph; it did not change the workflow or census
logic. The subsequent control build passed all four gates.

The retained receipt paths point to the original live directories. Every audit
resolves them through the corresponding relocation witness before checking the
recorded byte count and SHA-256. `failure_audit.py --write` materializes the
deterministic `failure-audit.json`; `--check` reconstructs the same result and
fails closed on a missing archive, changed hash, altered exit sequence, or a
rewritten unknown lock-generation return code.

This note records setup custody only. All 45 captures are now terminal, the
final before and after builds passed all four gates, and post-cleanup root
validation passed. Root owns the remaining packet analysis and final seal. No
native capture or heavy replay was performed by this audit.
