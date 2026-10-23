# 0361fe — IWork native limit payloads

This narrow follow-up closes the Numbers native error gap called out in
`0361fd` §“Left open”. The private `NativePopUpError::Limit` and
`NativeReplyError::Limit` variants now retain `LimitKind`, `observed`, and
`maximum` until their respective public facade mappings.

The propagation covers operation-budget charges, IWA constructors and archive
serialization (including Snappy bounds), strict Numbers storage/popup/comment
codec limits, and bounded BNC output. Known resource axes map to the matching
typed `LimitKind`; unknown codec resource variants remain a codec failure, and
no limit path fabricates a resource or count. Conversion helpers run only on
failure, so successful operations gain no allocation or accounting pass.

Focused unit coverage exercises budget and codec/BNC mappings and verifies
that both public facades preserve the payload. After the isolated host run,
run:

```text
cargo test -p litchi-numbers --lib
```

The implementation is based on integration tip `eb718eece7`.
