# 0437 ODP candidate oracle handoff

The current provider draft is the source of truth for this draft. One report
is one `odp_buffered_create` invocation for one shape. The report uses
`source.odp_slides` and its exact buffered implementation identity is
`litchi_odp::Builder::add_slide_with_title/build`.

The measured deterministic input has 64, 4,096, or 8,192 slides. Slide `i` has these
separate UTF-8 fields:

- `litchi-perf-odp-buffered-title-{i:05} {variant}`
- `litchi-perf-odp-buffered-body-{i:05} {variant}`

The four-cycle variants are `plain slide`, `Unicode café Δ 中`,
`entities <&> "quoted"`, and `mixed façade Ω <&> value`. The semantic oracle
hashes `litchi-odp-buffered-semantic-v1\\0`, the slide count as u64 little
endian, then title length/title bytes followed by body length/body bytes for
each slide. No presentation separators are hashed.

The logical input byte projection is the presentation text: title UTF-8
bytes, LF, body UTF-8 bytes, then LF LF between slides, with no trailing
separator. `entry_bytes` is the first slide's title + LF + body byte count.
The source summary must independently report `slide_count`, `title_count`,
`body_count`, title/body byte totals, and four-element title/body variant
count arrays.

The source summary also requires true gates for archive member set, manifest
bindings, semantic reopen, immutable styles/meta, runtime output digest,
runtime sink length, page structure, and page geometry. The latter two names
are reserved in this draft for the provider's upcoming explicit geometry/page
gates; they are deliberately fail-closed requirements rather than inferred
from the archive hash.

The package contract is `ODP/ODF/ZIP`, stored `mimetype` plus deflated XML,
with members `mimetype`, `content.xml`, `styles.xml`, `meta.xml`, and
`META-INF/manifest.xml`; the target is `content.xml` and the manifest has
four file entries. The fixed style and metadata pins are:

- styles.xml: 1,960 bytes,
  `d9881e91085516246a19c30d9e5cde39a8b10d7e42120b135f48f5ca8afef8d2`;
- meta.xml: 387 bytes,
  `c7e55a3560c73aa42da85eec4751c3e78b5cc53ff964f50acba6c5cd105e6719`.

The sink retains the existing generic fields: `paragraphs` equals slide
count, `runs` equals two text objects per slide, and `input_bytes` equals the
presentation text projection. Source reads and source-operation vectors remain
explicitly not applicable for fresh authoring.

The protocol matrix is normal/allocator, CPU 2, one worker, 30 retained
samples, 3 warmups, two repeats, and tiny/medium/large shapes. CLI bindings
remain `before-buffered`, `after-buffered`, and future `after-streaming`; only
the buffered producer is currently a typed baseline implementation. No speed,
causality, memory, throughput, or cross-role lexical-byte claim is authorized.

The legacy 32,768-title request remains an intentional aggregate audit-limit
refusal. This measurement scope records that boundary and does not relax the
production limit.

## Candidate streaming extension

`candidate/verify-report.py` and `candidate/protocol.json` preserve the
buffered `source.odp_slides` key set and the two buffered role bindings. The
candidate adds the `after-streaming` selector with the provider API identity
`litchi_odp::streaming::stream_plain_slides_to`; it does not reinterpret the
old `before-buffered` binary as a streaming implementation.

The streaming source consumes the same title/body fixture and must pass the
same semantic digest, page/frame structure, geometry, styles/meta, manifest,
member, output, and reopen gates. The protocol records the optional XML
declaration policy and the stored offset-zero `mimetype` layout as package
metadata. The external report still carries the existing corpus projection
(`title`, LF, `body`, LF LF between slides); the provider's raw input is the
sum of UTF-8 title and body bytes without those separators. The candidate
therefore checks streaming `sink.input_bytes` against that raw sum while
keeping buffered `sink.input_bytes` bound to the corpus projection. If the
streaming source summary exposes the approved optional
`provider_input_text_bytes`, it must equal the same raw sum; an unknown field
is rejected.

The streaming sink must report `retained_output_bytes: 0` and the protocol's
fixed `retained_authoring_window_bytes: 4096`. The role-specific protocol
keys bind the selector, source field, provider report field, input accounting,
and fixed window, so a missing, unknown, or mismatched limit/window cannot be
silently accepted. These checks establish report identity and correctness
scope only; they make no speed, throughput, or fixed-memory claim.

The provider API contract remains the source of truth for any final report
field naming. Before formal capture, the harness must either emit the named
optional provider input field or omit it; it must not substitute a new field
without versioning this candidate protocol and oracle together.
