# 0362ud — combined DOCX, DOC and facade validation

All seven checks on assembly `8d50e23436` passed. The source was held fixed
through the run; this receipt does not qualify later incoming branches.

The five semantic groups passed 183 tests: 19 variables unit tests, 52 related
DOCX integration tests, one DOC deferred-edit atomicity regression, 35 DOC and
DOCX facade tests, and 76 facade error-classification tests. Strict all-target,
all-feature Clippy passed for `litchi-docx`, `litchi-doc` and `litchi-ops`.
The locked all-target, all-feature performance harness check also passed.

The commands, exact source hash, durations, external log paths, sizes and
SHA-256 values are in [validation.json](results/0362ud/validation.json).
The lane used one Cargo job and the coordinator's capped target. Completed
debug test executables were hashed and retired by the existing runner; its
external receipts retain those hashes. No release or measurement binary was
removed. The largest recorded target size before a group was 18,801,268,035
bytes, below the 20 GB cap.

This is combined semantic and compilation evidence. Native performance,
corpus dispositions, final inventory/glossary closure and the full gate on
the eventual assembled head remain required.
