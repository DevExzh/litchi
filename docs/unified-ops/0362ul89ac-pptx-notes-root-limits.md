# 0362ul89ac — PPTX notes root-limit qualification repair

Status: source-ready repair based on the immutable qualification tip
`ae6f32d202a04e3d6675fa431b2579c0d5f27192`. This leaf was intentionally run
without a Cargo/build/rustdoc slot; the coordinator owns combined validation.

## Repair

The notes XML scanner now selects the root-specific policy before intersecting
the caller's semantic limits:

* `sld` keeps the independent 16 MiB XML and one-million-node ceilings. The
  selected notes profile no longer narrows an owning-slide proof to its
  default 8 MiB/100,000-node values. Opened eager parts retain the separate
  64 MiB `parts::raw_xml` ceiling; source-backed reads continue to use their
  caller-selected `SourceLimits` part ceiling.
* `notes`, `notesMaster`, `theme`, `presentation`, and other notes-closure
  roots still intersect their root ceiling with the selected input/output and
  node limits. Depth, attribute bytes, text, aggregate, and source limits are
  not bypassed.
* Attribute count is governed by `MAX_ATTRIBUTES`, independently of the node
  ceiling; the selected policy's attribute-byte limit remains active.

The source-backed notes closure passes the full slide-root semantic ceiling to
its owning-slide loader while retaining the selected notes policy for the
notes, master, theme, and presentation parts. The stale node-ceiling test now
asserts the typed `XmlNodes` refusal and its observed/maximum pair.

## Source-only checks

`rustfmt --edition 2024 --config skip_children=true --emit stdout` parsing of
both changed Rust files passed with no formatting delta, and `git diff --check`
passed. Cargo/native/full-feature tests, rustdoc, the workspace gate, and ABBA
remain coordinator validation; no generated evidence, fixture, executable, or
`3rdparty/` path is included.

The immutable source tip is reported in the coordinator handoff; this record
is included in that tip.
