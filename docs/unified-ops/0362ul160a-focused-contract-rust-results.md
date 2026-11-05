# 0362ul160a — focused contract Rust checks pass after 157b

The first focused run in 160 catches a missed test consumer from 157: an
optional typed graph iterator is compared directly with `None`, requiring
an irrelevant `PartialEq` implementation. Correction 157b uses `is_none()`
with the same assertion meaning. It commits on the integration branch as
`49fc1721e2db48fb95500e67378cd82a24e99092`; production behavior is unchanged.
The failed run is preserved and is not relabeled.

A fresh run on that committed source passes all 18 IWA graph tests, all 26
IWA index tests and the single nested owned PPT record fidelity test. Root
verifies both rows' command lines, complete logs and target traces, unchanged
source/locks, clean resource guards and empty post-exit process samples. The
service exits successfully, with no owned builders remaining. The target
ends at 385,684,596 bytes, below its 20 GB ceiling.

Evidence is external in
`(deleted-coord-artifact)`:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `state/00-iwa-graph-index-contracts-attempt-contract157b-v1.log` | 4,092 | `e4aafb955c872523455b1f122a5d18a430cf8592c44ffe63a88b723d9e7f5c9f` |
| `state/01-ppt-owned-record-fidelity-attempt-contract157b-v1.log` | 22,158 | `955b6dde6cd9f2ac7e54ad1bf7d5eae8966695c880ae8a84857a13a433ae2372` |
| `root-terminal-verification.json` | 1,849 | `29d2631b7610b99d684afe3d9a55ea9f9f20d363ea0036ffd14d85dcd15e7883` |

The source-bound plan has SHA-256
`616fbcf87defa25963ff569651aae3fdda0ece6455a6c0364a5d36876b794e6e`;
the driver has SHA-256
`8c304f4481a29293f2a295bfe6df650d3162993365f48fa92dce1dcecf51e86d`.
The commands are `cargo test --offline --locked -p litchi-iwa-graph -p
litchi-iwa-index --lib` and `cargo test --offline --locked -p litchi-ppt
--lib parsers::parser::tests::owned_record_view_preserves_the_complete_clone_traversal
-- --exact --nocapture`, executed serially by the approved capped runner.

This verifies the merged graph/index API consumers and owned record test.
PPT hyperlink repair 158a, full corrected-candidate gates, final census,
fidelity and performance qualification remain open. The build slot stays
with the coordinator for the forthcoming committed PPT repair verification.
Nothing is pushed.
