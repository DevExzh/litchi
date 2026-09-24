XLSX plain cell-tag allocation reduction

Unprefixed worksheet cells with no attributes or only an unqualified r attribute retain source spans without allocating an owned XML tag. Other cell tags use the existing representation. Edited cells reopen with the expected values and formulas, and untouched neighboring XML is preserved.

On one dense-wide one-percent commit-and-save workload, matched normal ABBA runs measured pooled p50 of 322.807 ms for the candidate versus 341.841 ms for control (5.568% lower). Separate allocator ABBA runs measured 20.559% fewer allocation calls and 7.079% fewer allocated bytes. Region peak live bytes increased by four bytes; no peak-memory reduction is claimed. This shared-host measurement does not establish a suite-wide gain.

Both builds use historical base 995bdaf09352297bde17e6ed7d986360f8c10134; the candidate adds the four production changes. Two additional test changes are absent from the binary dependency graphs. Full source/build/capture provenance, eight raw reports, runner, analyzer, and root gates are in validation-bundle.tar.zst. Extract it into an empty directory and run python3 -B verify.py within the extracted bundle. The verifier runs no benchmark and needs no binary.

Root independently extracted the archive, recomputed the reported medians and allocation values, verified the six source paths against the current checkout, and passed the full XLSX suite (1,396 ordinary tests and two doctests), strict Clippy, rustdoc, and scoped formatting.

The raw corpus ZIP is not embedded. Corpus reproduction uses the deterministic generator at the recorded source revision and archive SHA-256 5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714. The portable verifier checks recorded identities without regenerating the ZIP.
