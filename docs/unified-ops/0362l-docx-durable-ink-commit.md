# 0362l — Durable DOCX Ink replay returns the commit triad

- Parent: `d0406be017` on the coordinator trunk.
- Status: correctness validated; matched performance and integration pending.

`Package::apply_durable_ink` and `apply_durable_ink_with_options` previously
returned only a Package. They now return `ink::Commit`, the shared
`litchi_core::edit::Commit<Package, ink::Patch>` with package, patch and
Diagnostics accessors. Every successful consumer migrates in this source
commit, including the durable codec's own inverse test. No alias or deprecated
entry point is added.

Changed forward and inverse replay retain the existing typed-intent, logical
source/target fingerprint, inverse proof reconstruction and semantic
publication checks. They return the checked in-memory patch from the ordinary
`apply_ink` path. Exact no-ops reuse the already admitted edit state to produce
a no-op patch and unchanged diagnostics. Changed replay releases its initial
admission state before hashing/reconstruction and clones the output package
only for publication. Dropping a returned commit leaves the input unchanged.

Two boundary regressions were reproduced before repair. Default durable replay
reset the owner's retained ink policy to defaults; it now takes the inventory
policy from `semantic_limits().ink_value()`, as `edit_ink` does. A mismatched
logical source fingerprint previously returned generic Invalid; it now returns
PatchConflict with a fixed, content-free reason. This remains invalid_input at
the File classification boundary. Explicit EditOptions still select the
caller's requested replay/reconstruction bounds.

The new integration coverage checks the Commit type, all triad accessors,
changed/inverse/no-op diagnostics, forward and inverse replay of returned
patches onto independently reopened content-matching packages, exact restored
members, unchanged source bytes, exact/one-under retained payload budgets and
typed stale-source errors. All 57 Ink integration tests and 1,489 library tests
pass. Rustdoc passes 170 tests with 31 existing ignored examples; all-target,
all-feature Clippy with warnings denied and formatting pass. Both fresh DOCX
glossary profiles pass. [Final validation](results/0362l/validation.json)
receipts include the failed lock setup, the initial test compilation correction,
the two reproduced assertion failures, and all successful runs.

The six-route matched probe in `results/0362l/` measures forward, inverse and
no-op replay, each with and without publication, on one 1,000-paragraph source.
The exact same probe is built on both sources; the `before` feature adapts only
the Package/Commit extraction. Fixture construction and durable wire decoding
precede measurement. Each measured result performs its first annotation read;
publication routes also serialize it. Untimed preflight now additionally
reopens independent content-matching base and target packages, checks forward,
inverse and no-op replay against exact package bytes, reads through a fixed
4-KiB reader, and writes publication through a fixed-state 8-KiB digest sink.
The stream receipt proves the bounded sink/read path without retaining output
bytes. Native and allocation validation outputs are compared before a timing
runner can start. Counting runs and guarded native runs are separate. No
timing or allocation acceptance is claimed until those builds and comparisons
finish. The qualification runner requires the final baseline and candidate
refs at execution time, records source/probe/lock/compiler/binary digests,
rejects stale binary paths, and keeps all raw artifacts outside the checkout.
The final workspace inventory, remaining W8 replay families and integration
gate stay required.
