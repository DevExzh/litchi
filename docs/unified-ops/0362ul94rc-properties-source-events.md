# 0362ul94rc — properties inert-event source validation

The 0362ul94rb source boundary covered attributes, text, CDATA, and general
references, but both property readers still discarded Comment, Processing
Instruction, and declaration events without checking their decoded payloads.
That left XML 1.0 controls in retained inert bytes, even though the custom-XML
reader already rejects the same source corruption.

The application scanner and core reader now validate comment payloads, PI
targets and payloads, and declaration bytes in their existing event loops.
The checks borrow the event bytes through the reader decoder and preserve the
single-pass application scan. Valid comments and PIs remain inert and are
covered by exact app splicing and core no-op byte-fidelity tests. Adversarial
source tests cover comment and PI controls in both families; authored-value
classification is unchanged.

This is a source-only repair. No Cargo, build, rustdoc, runtime, native, or
performance command was run; the coordinator must rerun the affected gates.
