# XLSX form-control source payload and parser budgets

The leaf retains source-backed form-control XML in shared payloads, makes property clones share nested storage, and accounts parser work, cancellation, retained buffers and temporary namespace capacity. Namespace Vec reservations stay active through Arc publication. Root namespace Arc metadata is covered by fixed Properties storage; distinct child contexts include their own Arc header. Exact and one-byte-under tests cover those boundaries and writer validation.

Independent review approved the exact two-file leaf hashes in source.json. Root isolated validation passed 1,171 XLSX library tests and 18 form-control property integration tests, strict Clippy for those targets, and warning-denied rustdoc. Commands, exit codes, compressed logs and the resolved lockfile are retained. A dedicated target and TMPDIR=/var/tmp were used without inherited warning suppressions.

This is the parser-owned accounting contract, not an allocator-exact census of every nested Arc metadata allocation inside namespace/opaque values or caller-side setters. Form-control owner-reader integration remains a separate unapproved batch. No runtime performance improvement is claimed by these correctness and resource gates.
