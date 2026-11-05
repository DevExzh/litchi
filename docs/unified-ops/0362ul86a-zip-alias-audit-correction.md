# 0362ul86a — ZIP alias audit correction

The source review of `0362ul86` found one multiline duplicate introduced by
the mechanical migration: `preserve.rs::map_local_layout_error` matched
`ErrorKind::IO(_)` twice. The second arm was unreachable and did not change
behavior, but it violated the single canonical matcher promised by the lane.
It is removed here.

The multiline duplicate scan now returns no matches across the ZIP, ODF, OPC
and ops consumers. This is a source correction only; no Cargo, rustdoc,
native or performance command has run yet.
