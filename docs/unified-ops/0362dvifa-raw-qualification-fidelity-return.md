# DrawingML raw qualification fidelity return

Coordinator review returns `b2da8f8ba60d01ca627af5b15b697ea895541fd9` before merge. Its 39 tooling tests, ten untimed preflights and candidate smoke runs are useful diagnostics, but they do not satisfy the unchanged [0362dvie qualification plan](results/0362dvie/qualification-plan.json), SHA-256 `e0cb370e0399c072de9fe4b141ceff682b87d7936bec5332cedc9a181fc13505`.

The required corrections are concrete:

- SVG contextual preparation uses an empty namespace context with a self-declaring fixture. Exercise inherited namespace bindings and independently check the intended relationship and standalone output.
- Formatcode2 and Theme Family changed-output checks compare production output with a production-generated fixture. Add an independent expected splice and untouched-source proof. Bind the complete Theme-part requested name, id and variant id before running the mutation, rather than deriving the request from its result.
- The timed geometry reader calls a helper that formats and classifies errors. Keep native results under the timer and move error normalization outside it. Preselect Diagram entries and Theme-part families before timed writes and output-cap calls.
- The Diagram output-cap proof invokes both text and drawing-reference edits, while its timed row invokes only the text edit. Make the proof and declared measured operation agree.
- Restore exact retained-source checks for Ink actions in the probe, validator and records. Both the current `ink::actions::write_profile_to` and historical `996651b6b19717a0192dfd4144bccd367fb2ea44` implementation write `profile.source()` exactly. The probe calls that writer, and its diagnostic hashes already show exact source and changed-source output. The claim that this writer canonicalizes XML is incorrect; model equality cannot replace its byte-fidelity obligation.

The independent reviewer withdraws the earlier statement accepting Ink's model-only classification after checking both production implementations. This is a qualification regression in the returned packet; no production Ink capability loss is found.

The original diagnostic receipt remains external at `/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/drawingml-native-dvif-v7/compile-preflight-diagnostics.json`,  with SHA-256 `067d37c3e56e63242ac53e40054aa591b013b19a5ac377ac9a41d987f56827f4`. Its generated driver lock and candidate-only smoke runs do not establish frozen-lock paired performance, allocator or RSS acceptance. The separate owner repair and actual paired qualification remain required. No row, baseline fence or sampling requirement is waived.
