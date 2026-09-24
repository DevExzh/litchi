# Bounded ODF namespace semantics

The shared helper compares XML 1.0-normalized namespace values while leaving lexical input untouched. Plain UTF-8 values use a borrowed fast path; values requiring normalization are sized before allocation under a 64 KiB lexical/output ceiling. Malformed entities, illegal XML characters, raw attribute delimiters, reserved bindings and named empty namespace bindings fail closed. Default namespace undeclaration remains valid.

Independent helper-only review approved the exact hashes in source.json. Root isolated validation passed 516 tests across 25 targets (one ignored), all-target strict Clippy, and warning-denied rustdoc. Commands, exit codes, compressed logs and the resolved lockfile are retained. A dedicated target and TMPDIR=/var/tmp were used without inherited warning suppressions.

This batch exposes UTF-8 namespace comparison and declaration-validation helpers. It does not establish complete XML grammar validation, non-UTF-8 host support, or fix parser admission that rejects escaped reserved bindings before calling the helpers. ODS/ODT host integration and that admission adapter remain separate work. No performance claim is made.
