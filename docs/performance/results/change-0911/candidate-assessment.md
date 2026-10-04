# 0911 owning cold snapshot candidate assessment

This is an evidence-only assessment of the retained 0910 binary. It does not
select or implement a production optimization. The v3 packet is admitted: all
four legs have one setup owner edge from
`litchi_perf_baseline::build_pptx_source_edit_corpus`, one measured owner edge
from `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration`, conserved
raw counters, a zero terminal dump, and passing source, semantic, output, sink,
and immutability gates. The canonical reduction is
`analysis.json` (SHA-256
`ac6fa98aed88d9594fe2bafef7c38cb06f08dc6b84eab574b919fd11bdedf76f`).

## Admitted comparison

The retained executable is the 0910 binary
(`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`), with
zero warmups and one measured operation per leg. Setup is retained for custody
and excluded from the route comparison.

| leg | measured owner Ir | setup owner Ir |
| --- | ---: | ---: |
| materialized-a | 1,000,666,812 | 126,538,787 |
| owned-a | 1,099,868,747 | 126,681,314 |
| owned-b | 1,099,912,679 | 126,577,157 |
| materialized-b | 1,000,594,047 | 126,635,417 |

The two descriptive medians are 1,000,630,429.5 Ir for materialized and
1,099,890,713 Ir for owned, a difference of 99,260,283.5 Ir (+9.92%). They are
guest instruction observations, not native elapsed time, allocation, RSS, or a
statistical speedup estimate.

## Source map and bounded opportunity

The owner is `Package::opened_presentation_transaction`
(`crates/litchi-pptx/src/package/model.rs:296`), which calls
`opened_presentation_with_limits` and takes the cold `capture_with_provenance`
branch when no part-digest memo exists (`package/model.rs:263-285`).
`capture_internal` (`crates/litchi-pptx/src/opened/model.rs:730-877`) then:

1. parses the presentation and strict slide catalog;
2. resolves all slide relationships and captures every slide's MCE projection,
   root/name result, and notes-root proof;
3. loads the notes snapshot from those proofs and builds the slide-name index;
4. clones the OPC graph and computes the complete package fingerprint; and
5. returns through `Transaction::new`, whose package and vector clones share
   payload `Arc`s (`crates/litchi-pptx/src/opened/transaction.rs:42-49`).

The route-specific signal is deferred payload materialization during those
`get_part` calls. `DeferredPayload::decode` indexes the retained ZIP, reads and
inflates the member, checks per-part and aggregate limits, and stores one
`Arc<Vec<u8>>` (`crates/litchi-opc/src/payload.rs:178-193`). `Part::ensure_payload`
forces that cell for both `BlobPart` and `XmlPart`
(`crates/litchi-opc/src/part.rs:353-355, 631-633`). In the measured owning
profiles, the direct `get_part` edges to those forces are:

| direct edge in `get_part` | owned-a | owned-b | materialized-a/b |
| --- | ---: | ---: | ---: |
| `BlobPart::ensure_payload` | 75,192,121 | 75,197,158 | 224 / 224 |
| `XmlPart::ensure_payload` | 23,883,168 | 23,905,655 | 14,742 / 14,742 |

The owning path also contains the ZIP read/inflate/CRC/copy work below those
edges: `DeferredPayload::decode` reaches
`IndexedArchive::read_with_accounting` about 98.9M Ir in each owning dump, and
the whole owner reports 50,868,707/50,874,914 memcpy Ir and
28,254,410/28,254,417 memset Ir. The decode and `ensure_payload` edges are inclusive diagnostics and must
not be added to the owner or fingerprint partitions. The memcpy/memset
function self counts are exclusive; they likewise are not additional costs
on top of those enclosing partitions.

The complete fingerprint is the other large stage. It sorts all parts, feeds
root relationships and non-part members, obtains every part, hashes each
payload, then hashes each part's name, content type, payload digest, and sorted
relationships (`crates/litchi-pptx/src/opened/model.rs:912-1020`). Its inclusive
cost is 922,848,036 / 922,844,051 Ir for materialized-a/b and
1,000,195,725 / 1,000,199,235 Ir for owned-a/b. The common
`sha2::sha256::soft::unroll::compress` self cost is exactly 922,100,520 Ir in
all four measured dumps. The owning-only increment is therefore concentrated
in first access to deferred payloads and its ZIP/inflate/copy accounting, while
the dominant fingerprint/hash work is common in this guest profile.

The bounded candidate to investigate is payload decode/copy handling on the
owning cold path. A future change could assess whether decoding and hashing can
share work or an allocation while retaining the decoded payload required by
the snapshot. It must be measured end-to-end; moving inflation earlier only
changes phase attribution and is not a win. The profile does not justify
skipping the cold fingerprint or weakening lazy decode semantics.

The XML/parser cluster is smaller and shared: slide finishing is about
59.6-59.9M Ir, notes-root loading about 5.6-5.9M Ir, presentation cataloging
about 2.75M Ir, and graph cloning about 1.14-1.18M Ir. The owning
`validate_slide_relationship` inclusive rows are about 22.5M Ir versus 1.12M
Ir materialized because relationship resolution reaches owned `get_part`; these
rows overlap the payload and fingerprint diagnostics. `SlidePart::finish_from_processed`
already keys its root proof on the exact raw observation and still performs root
and limit checks (`crates/litchi-pptx/src/parts/slide.rs:1216-1253`). This is a
later, correctness-sensitive memo/parser investigation, not a current
candidate. This guest profile supplies no basis to prioritize transaction construction
or graph cloning; their native cost share still requires native evidence.

## SHA dispatch limitation

The locked dependency is `sha2 0.11.0`. Its x86 source selects
`x86_sha::compress` only when the `cpufeatures` `shani_cpuid` check succeeds;
otherwise it calls `soft::compress`
(`/home/zhuhe/.cargo/registry/src/index.crates.io-1949cf8c6b5b557f/sha2-0.11.0/src/sha256.rs:50-78`).
The retained host-flags receipt says `sha_ni=true`, but explicitly scopes that
fact to host exposure and records no native backend observation or Valgrind
guest CPUID capture (`host-flags.json`, SHA-256
`7ed849b9c5c4fa1f8859a86ba0bc1f2d0d84af6ddca235632f8c9e9911300e32`). The
Callgrind guest proves only the software path. Therefore no scalar-SHA
replacement, SHA-NI claim for the 0910 native run, or native Amdahl estimate is
admissible. Native dispatch confirmation is required before considering any
hash implementation change.

## Correctness boundaries and next gate

Any payload candidate must preserve ZIP CRC/error behavior, stable
`OnceLock` refusal, `PartBytes` and `TotalPartBytes` limits, retained source
provenance, and the decoded allocation used by later readers. The fingerprint
must continue to cover the same complete package, domain strings, part order,
content types, payload bytes, and sorted relationships; `PartDigests` reuse is
valid only for an allocation whose bytes are proven identical. Root/MCE and
notes proofs may be reused only for the exact source bytes and unchanged limits.
Snapshot and transaction isolation must remain unchanged.

The next evidence gate is a native dispatch check and a native cold snapshot
sample before choosing between payload handling and hashing work. If a bounded
payload change is later implemented, rerun the native owning/materialized
control and all existing source, semantic, output, sink, and immutability gates;
rerun v3 attribution if the owner call path changes. No production source was
modified for this assessment.
