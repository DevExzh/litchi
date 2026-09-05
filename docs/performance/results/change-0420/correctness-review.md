# Final source and correctness review

Candidate `81c70058a` passed independent source review with no blocker. Storage
selection occurs after the target's physical/package parsing and before
built-in Part construction. Exact URI/content type, donor Arc/visible-byte
consistency, target byte equality and no-greater capacity gate reuse. Parts,
source XML records and preservation records receive the same selected Arc.
The target raw archive establishes its own exact-source authority.

PPTX opts in only when the original destination retains clean owned-source
authorization. Dirty destinations, custom parts and altered save options keep
the previous reopen path. Detached candidate capture, patch, fingerprint,
physical and semantic publication checks remain in place.

Validation on the final production implementation:

- OPC/PPTX all-feature unit, integration and doc tests: 1,286 passed, three
  ignored. The only subsequent change strengthened a test fixture.
- Final focused ownership tests: seven passed, including a matching callback
  positive control before the input-limit refusal test.
- Clippy all-feature/all-target/no-deps passed with the established three
  command-local exemptions: `chunks_exact_to_as_chunks`, `clone_on_copy`,
  `needless_lifetimes`.
- Rustdoc for both crates with `-D warnings` passed.
- Scoped rustfmt and staged diff whitespace checks passed.

The focused tests cover XML/binary sharing across preservation holders,
target-only metadata and exact output, byte/capacity mismatch fallback,
inconsistent custom donors, mutation isolation, limit refusal before donor
callbacks, and malformed ZIP typed errors. The first focused compile failed
because the new module omitted a `PackageWriter` import; its raw log is
retained alongside the corrected passing runs.

The advanced OPC API may invoke a custom donor's trusted `blob_arc()` method;
that method can allocate or have side effects. The PPTX gate excludes custom
parts. Full target decompression still precedes reuse. These tests and the
capacity guard do not establish an aggregate live-memory or local-peak bound.
