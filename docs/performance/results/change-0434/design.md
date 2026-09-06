# 0434: ODS scalar-text serialization investigation

The accepted ADR tree remains `c950b6c8be822561b498d7bbe87c460873dcbf49`.
The baseline is `91de70866df67bafd7d2f4c8931269ea540e1a1d`, whose production
streaming code and binaries equal the 0433 after-streaming implementation.
That retained large whole-process profile samples `ExecutionContext::consume`
at 24.41% self overhead through text serialization. Source inspection confirms
that every ordinary Unicode scalar performs a separate work-budget charge and
copy into the common row buffer. The profile also includes the untimed oracle;
it is a hypothesis source, not operation-local attribution.

The candidate groups contiguous, already-validated ordinary UTF-8 text into
borrowed spans of at most 256 bytes. Entity/reference writes keep their original
lexical forms. Each scanned character still checks cancellation, and each span
checks cancellation before charging work and writing. No additional owned text
buffer, production dependency, unsafe code, or execution-policy type is needed.
The existing common fragment audit and ZIP publication boundary stay in place.

A span that cannot fit the row window or cumulative Work budget falls back to
the original scalar writes before setting any failure state. Failed aggregate
budget consumption rolls back its ancestor charges; the scalar fallback then
preserves the first failing scalar, observed limit, buffer prefix, and charged
work for deterministic limit failures. A cancellation during a pending span
may leave its scanned prefix uncharged and unserialized; Work accounts for
encoded XML bytes. The incomplete row is never published and the outer sink's
accepted-byte progress remains authoritative. Cooperative cancellation does
not promise identical race timing or failed-operation internal progress.

Differential tests must compare the old scalar encoder with the candidate over
Unicode/escaping and span boundaries, local and ancestor Work limits, row-window
limits, failure state, and charged counters. Public integration checks retain
exact/one-under budgets, cancellation, short writes, and sequential failure
progress. Formal captures must require exact deterministic archive/XML/output
hashes and semantic identities across the two implementations before comparing
normal latency or allocator observations. The baseline binaries are preserved
before source changes; later ABBA capture records runtime checkout identity
separately from their retained build revision and source hashes.

This batch investigates one measured ODS authoring path. Bounded ODT/ODP fresh
authoring, existing-document append, native compatibility breadth, rich content,
and the wider non-iWork goal remain separate work.
