# 0362zacb — XLS revision-log length preflight

Status: coordinator follow-up to `0362zaca`. This a-record repairs the
bounded-memory defect found during review: `litchi_xls::raw::Reader::revision_log`
must reject an oversized declared stream before `open_stream` can materialize
its bytes. It does not change the retained-ingress limit rule or defer any
other raw-boundary requirement.

## Repair

`revision_log` now combines the caller's `SourceLimits` with the retained
ingress profile, obtains the Revision Log entry's declared length through
`OleFile::stream_len`, and returns the typed `LimitKind::EntryBytes` error
before opening the stream. Only entries within the effective limit reach the
parser. This preserves the existing typed error and keeps the lazy raw
boundary bounded when a source was admitted under a larger profile but the
operation requests a narrower one.

The regression uses an admitted CFB workbook with a retained 2048-byte entry
profile and a 32-byte Revision Log entry. It requests a 1-byte operation
profile, checks `observed: 32` and `maximum: 1`, and uses a counting reader to
prove that the source read count does not change during the refusal. The test
therefore checks the preflight ordering, not only the error variant.

## Validation

The source packet remains based on commit `9082f27816`; the local ignored lock
retains SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` from
`0362zaca` (the offline-only direct `litchi-ops -> litchi-vba` edge and no
registry version or checksum changes).

```text
source (deleted-coord-artifact)
cargo test --locked --offline --all-features -p litchi-xls \
  --test xls_revision_log -- --nocapture
PASS — 4 passed; 0 failed
```

Rustfmt and `git diff --check` passed. The broader checks recorded in
`0362zaca` were not repeated for this focused preflight-only follow-up.
