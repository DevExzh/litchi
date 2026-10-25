# 0362ul154 — mandatory contract audit return

The completion audit finds public API debt in frozen candidate
C=`f8cd398a770f7363f628be40edb3ce4219566c16`. It cannot be accepted as the
final candidate. This is required engineering under 0351, not an additional
exception or deferral. Item 22 remains the only deferred audit.

`BatchCommit` in `crates/litchi-ops/src/batch/execute.rs` exposes
`results()`/`into_results()` instead of `value()`/`into_value()`. Record 0351
§2.5 rule 7 explicitly requires the triad accessors for both facade commit
types. The deviation recorded in 0362 §2 must be removed, with every Rust
consumer migrated in the same commit. Python/CLI result payload semantics
need not change when their backing Rust calls migrate.

The public deprecation census also identifies 22 first-party routes:

| Surface | Count | Required capability-preserving migration |
| --- | ---: | --- |
| ODT `transaction::Edit` raw-index forwarding methods | 11 | Delete the forwarding names; retain the adjacent checked `Position` methods and migrate any consumers. |
| Keynote `Show::slide` | 1 | Use the positional slice lookup or typed selector. |
| PPTX `Effect::to_preset` | 1 | Use `preset_class()`, which preserves its result. |
| PPT `RecordParser::find_records` | 1 | Preserve owned cloned records under `find_records_owned`; migrate consumers and remove the old name. |
| IWA graph raw-ID forwarding methods | 7 | Use the canonical typed-ID methods, with checked conversion at input boundaries. |
| IWA index `ObjectIndex::reference_graph` | 1 | Preserve the complete immutable, clonable graph snapshot through `reference_snapshot()`. |

The no-alias rule applies to these public migration routes (0351 §2.6 and
glossary §1). ODT's 17 `pub(crate)` deprecated helpers and the private PPT
recursive helper are outside the public surface; Soapberry-ZIP's upstream
substrate APIs are outside this format migration. The already scoped W7
native-ID host methods are plain APIs, not these deprecated forwarding
routes. No capability loss is authorized by this audit.

Documentation also needs correction: the glossary's facade detection prose
still says `&OpenOptions`, while the implemented raw detectors require
`&detect::SourceLimits`; 0362's RTF error discussion still says there is no
`Refused`, despite the current `Refused(Refusal)` variant and classification.
The sheet-text re-export and design §16.4 are present. Historical inventory
and disposition evidence remain historical: the current 84-profile matrix
still has no captured entries, captures or lanes. It does not prove final
API closure.

Root establishes three isolated repair leaves from C, with all consumer and
gate migrations required in their commits: `uo2/final-contract-accessors`
(155), `uo2/public-format-aliases` (156), and
`uo2/public-support-aliases` (157). Their setup is source-only and capped at
4 GiB; no worker may start Cargo or native execution while the current full65
service owns the build slot. Root will review and integrate the repairs,
resolve shared-file conflicts, and remove their worktrees after integration.

The completed current-native source binding is valid evidence for C. Root
verifies its exact 461,401-file archive, three role locks, six protocol files,
generated pin and manifest projections. The binding receipt is 9,966 bytes,
SHA-256 `e33527de3467661ab4af3d92550f917fd7957d580a5ea1821f32603a23e7388c`;
the external terminal verification is 3,210 bytes, SHA-256
`e7e856ac4e041e5220d1834063bbb7e09a38e060980cc71034fce9c0dc7d9b27`.
It is not installed as a final binding after this audit return.

DrawingML baseline materialization also completes successfully. Root verifies
all ten distinct roots, their source heads, sparse patterns, three role locks
and eight adapter identities each; four old locks and stat records are
preserved. C, P and common Git configuration remain unchanged throughout the
job. Its external terminal verification is 3,924 bytes, SHA-256
`ac2bfd4814955989b259dfadf8abf4bc51b1a33fff1b47931331f49329e1c7b6`.
These historical baseline roots remain usable for the corrected candidate.

The independently reviewed signed streaming repair proceeds separately in P
(154a), with its consumer tests in the same commit: all 25 integrated policy
and consumer tests pass. Candidate bindings remain pending. The signed v5a
recipe's 15 source checks and corrected containment/provenance guards are
ready; an archive from C cannot substitute for the corrected candidate's
archive. The original failed signed tar remains preserved and unverified.

The current full65 run continues as evidence for C, with 32 completed rows
verified and the composite ops tests in progress at this observation. Final
acceptance, census and native measurements must bind the corrected production
revision. The final candidate will incorporate the reviewed production
repairs and the pending protocol fixes before its source is frozen; final
binding commits then remain separate from that frozen source. Nothing is
pushed.
