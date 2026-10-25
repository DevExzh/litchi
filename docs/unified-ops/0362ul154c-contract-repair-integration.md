# 0362ul154c — contract repairs integrated and leaves retired

The candidate integration branch now contains the three mandatory API repairs
returned in 154, their coordinator follow-ups, and the bounded signed-archive
verifier from 154a. Its head is
`f17dda6e6f13629b347bf3e290af9cc70f7dc936`. This is an unqualified candidate;
the PPT hyperlink regression from 154b remains under repair in 158.

Records 155/155a restore `BatchCommit::value()` and `into_value()`, migrate
the Rust consumers, remove `Show::slide`, and require all facade commit
accessors in both the glossary gate and API inventory. Python and CLI result
payloads retain their existing meaning. Records 156 and 157/157a remove the
remaining public deprecated forwarding routes while retaining checked ODT
positions, PPTX preset classification, typed IWA graph operations, the full
immutable reference snapshot, and the complete owned PPT record tree.
`find_records_owned()` has no unused selector argument; its fidelity test
checks nested order and payloads after the parser is dropped.

The three leaf heads are merged without conflicts:

| Repair | Leaf head | Integration merge |
| --- | --- | --- |
| 155/155a | `af2bc03c56f8a26e5e0f841495f081525b690388` | `1995290b705b8d14a5c3cf6b6cd04631c6cad146` |
| 156 | `99acd8f8f424b57c2f0ea34b44c21d9a77059807` | `21c9c3501f6e0afb1d2d7202a36fad9a0aa716a7` |
| 157/157a | `90ed01d38880c7936f2865547d1a2bbe1cb4f669` | `f17dda6e6f13629b347bf3e290af9cc70f7dc936` |

On this combined source, 151 Python tests pass across the glossary, debt
inventory, current-native, DrawingML and signed-native policy/consumer
suites. The external receipt is
`/home/zhuhe/code/litchi-worktrees/.coord/contract-repair-integration-f17dda6-20261025/focused-tests.json`.
Its 251-byte log has SHA-256
`c07d588e60c96671e2b255a38277036da0615aea52d4455e9d02033042da3c5a`.
This checks the source contracts and protocol tooling; it is not Rust,
native, corpus or performance acceptance.

Root verifies each leaf's exact head, ancestry in the integration branch,
and clean tracked state before removal. Each worktree contains only its
approved ignored root lockfile beyond tracked files; its bytes and metadata
are preserved in
`/home/zhuhe/code/litchi-worktrees/.coord/cleanup-contract-repair-leaves-20261025/`.
All three worktrees and their three merged branch refs are removed. The
terminal ledger records 62 registered worktrees and 75 remaining heads;
cumulative unified cleanup is 276 worktrees and 438 refs. The allocated-byte
ledger is not a measurement of newly available space. Active controls,
needed historical baselines and independent office branches remain intact.

The glossary now has 71 profiles after the IWA graph and index additions.
The current raw-census matrix must be refreshed before final capture. All
C1 source bindings and receipts remain historical evidence for their exact
revisions. The corrected source must be frozen only after the production
repairs and matrix updates are complete; final gates, census and native
measurements must bind that revision. No additional deferral is accepted,
and nothing is pushed.
