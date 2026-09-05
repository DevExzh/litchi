# PPTX media-copy static audit

This is a call-path audit at the captured revision, prepared before profiling.
Pass counts describe reachable source paths, not measured CPU percentages.

The owned selector times planning, commit, and final publication separately
(`tools/perf-baseline/src/lib.rs:46430`, `46455`, `46462`). Its source and
destination each contain eight deterministic 2 MiB images (`lib.rs:14182`,
`14218`). Package opening, clones and the output correctness workflow are
outside the reported phase sum.

Planning validates the dependency closure, computes physical fingerprints,
builds a candidate, serializes it and reparses it
(`crates/litchi-pptx/src/opened/cross_copy_plan.rs:774`, `837`, `906`, `1005`).
The eager reparse decompresses every admitted Part
(`crates/litchi-opc/src/pkgreader.rs:530`). Commit repeats graph fingerprints,
physical checks, source/destination captures and planning; it then applies the
patch to a clone and fingerprints the mutated candidate
(`cross_copy_plan.rs:450`). The semantic package fingerprint hashes every Part
payload (`crates/litchi-pptx/src/opened/model.rs:388`).

For this clean owned-source corpus, preserving publication copies untouched
destination records and Deflates the eight copied image payloads
(`crates/litchi-opc/src/pkgwriter.rs:613`, `681`). The source trace reaches four
large generated-image Deflate passes: the initial candidate, commit's replan,
the post-apply physical check, and final publication. The checks have semantic
and source-identity purposes; removing them requires proving equivalent
validation, preservation and refusal behavior.

The source-backed API supports a media-rich copy closure. Its planner stages
copied image bytes (`crates/litchi-pptx/src/presentation/source_cross_copy.rs:1508`).
Publication reruns preparation, checks staged image bytes, constructs a topology
plan and writes it (`source_cross_copy.rs:301`, `1721`, `405`). Untouched members
retain their raw source records while added images are Deflated
(`crates/litchi-opc/src/source_backed.rs:5340`, `6949`).

A next experiment can add an opt-in media-rich source-backed selector using
the same existing source/destination bytes, slide selectors, collision pattern,
sink and semantic/raw-member gates. The existing plain source-backed selector
is not a matched performance counterpart. Even with matched inputs, the owned
commit and source-backed publication contain different internal phases; define
an aligned total-operation boundary before claiming a causal comparison.
