# 0807 reader review

This is a static review of the packet readers after the ordinary control and
profile-wrapper builds. No workload, capture, build, or heavy replay was run
for this review.

The six-document seal path agrees. `seal_packet.py` names the same six
repository documents as `validate.py`, and both resolve them under
`docs/performance/`. The list order differs, but both readers compare the
resulting mappings, so the order cannot change the result. The staged and
committed inventory checks also include the packet seal itself and exclude
only the intended seal, cache, and external production entries.

The three-binary cleanup contract is complete. `cleanup.py` requires all
native, Callgrind, and frame-pointer completion receipts, checks the control
and profile paths plus `profile-fp` against the dedicated 0807 target, hashes
each regular executable before removal, and records all three descriptors in
`cleanup.json`. The replay-side cleanup checks require the same target-removal
witness and exact descriptor set. The target is checked to be a real directory
before the one top-level removal.

The evidence-only source chain is also present. `quality.py` checks the current
source manifest against `build/source.json` before and after formatting and
checks every inherited probe file against `build/probe.json`. The analysis
readers call `source_identity()`, which binds the current revision to
`origin.json` and the complete production file map to the copied 0806
before-source manifest. `seal_packet.py` then refuses any production-byte
difference from that build manifest. The current build manifest records
revision `3624e73236` and 9,196 files.

The inherited-probe binding is explicit. `validate.py` checks every copied
probe file against both `inheritance.json` and the 0791 packet reference;
`build/probe.json` additionally covers the materialized `Cargo.toml`. The
hash-bound legacy readers are checked through the inheritance references before
the five offline analyses run.

One reader-hardening finding remains before final sealing. `quality.py` writes
the source and probe descriptors into `quality.json`, but `validate.py` only
checks `quality.json.format.exit_code` and its log. It does not compare the
stored `quality.json.source` and `quality.json.probe` values with
`build/source.json` and `build/probe.json`, nor replay the expected formatting
command. A stale or edited quality receipt could therefore pass the final
validator if its format exit code and log artifact were made consistent. The
current packet shows no such mismatch, and the writer itself performs the
checks, but the final reader should either validate those fields directly or
provide a `quality.py --check` replay before the seal is written.

The source, document-path, three-binary, and inherited-probe contracts have no
other finding in this bounded review. The seal's production guard is not a
standalone substitute for source inheritance validation; the current final
workflow is sound only when the validator also runs all three analyses that
call `source_identity()`.

Root closed the quality-receipt finding before final replay: the validator now
checks the exact formatting command, source descriptor, probe census and
ordered timestamps, in addition to the successful exit and log identity.
