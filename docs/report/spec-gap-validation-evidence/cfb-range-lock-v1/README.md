# CFB range-lock validation

Validated against isolated baseline `29ad6a094` plus the exact files hashed in
`receipt.json`: 308 unit/integration tests and 14 doctests passed; one existing
doctest remains ignored. Strict library/test Clippy, formatting and whitespace
checks passed. The full crate run preceded the addition of `tests/range_lock.rs`;
its two tests were then compiled and run separately against the same production
files, and the final Clippy run included them.

The sparse integration fixture streams approximately 2 GiB per case while
retaining only metadata. It covers adjacent metadata-only FAT capacity boundaries,
a payload chain crossing the reserved sector, zero publication, actual reopening
under explicit limits, exact input admission, and malformed inbound FAT pointers.
Unit tests also verify buffered write shape and reader/allocator boundary rules.
This is correctness and I/O-shape evidence, not a throughput or native producer
interoperability claim.

Default reader limits remain 2 GiB. Explicit low-level and shared reader limits
admit v4 sources up to a documented 32 GiB resource ceiling. V3 input remains
limited to 2 GiB independently of caller limits. This resource profile does not
claim support up to the format's theoretical maximum. Input ceilings are not
independent FAT/MiniFAT allocation budgets: hostile table counts can request much
more decoded metadata than a legal minimal FAT. Separate structural table quotas
remain a follow-up hardening task; this batch makes no total-memory claim.

Independent read-only review found no remaining range-lock blocker. The separate
whole-file maximum-count consistency and fresh-root creation-time policies remain
follow-up work. Earlier validation exposed an inode-reuse test race, corrected in
`868f31873`, and nondeterministic storage ordering, isolated in `29ad6a094`.

Validation logs are stored as deterministic gzip files with their exact original bytes.
