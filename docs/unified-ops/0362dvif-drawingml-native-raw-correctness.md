# 0362dvif — DrawingML native raw correctness and timing repair

This finite repair is limited to the raw DrawingML probe, its validator and
qualification source tests. It does not change owner adapters, the paired
runner, or the allocator/RSS harness.

The coordinator's immutable native qualification plan is
`(deleted-coord-artifact)`
(the committed copy is `docs/unified-ops/results/0362dvie/qualification-plan.json`,
SHA-256
`e0cb370e0399c072de9fe4b141ceff682b87d7936bec5332cedc9a181fc13505`). It
retains the six raw profiles and names the refusal/replay rows whose native
results may be timed after their untimed proofs.

The probe now prepares every operation's correctness and detail receipt before
the first warmup. Timed functions black-box only the production result: typed
refusal checking, cap selection, output sizing, error classification, parsing,
digests, and JSON/detail construction are outside the timed region. Raw input,
sink, and Vec rows remain in the matrix, and refusal rows retain explicit
typed output coverage. Chart's bounded partial-output behavior is unchanged.

SVG contextual reads are prepared for both source and changed fixtures. The
normal Vec and bounded-sink rows and the outside-clock fidelity check use the
production contextual read/write route; the standalone reader/writer remains
an explicitly additional outside-clock route over the completed contextual
documents. Its semantic roundtrip is checked separately and is not folded
into the timed contextual workload.

The formatcode2 and Theme Family fixtures are non-canonical source-backed
fixtures. The probe parses each source, mutates the parsed value, and checks
the resulting lexical splice and decoded semantic change. Complete Theme-part
add/replace mutations parse their output and verify the requested family name,
id, and variant id. Input-limit boundary fixtures retain a valid source and
append XML-permitted trailing whitespace before crossing the typed byte limit,
so a malformed-input refusal cannot mask the limit refusal. InkML action output
is reported as exact retained-source plus semantic fidelity because the
writer preserves the parsed profile source, including the changed fixture.

The validator requires the contextual route receipt, the complete Theme-part
identity receipt, and the corrected InkML fidelity mode. No benchmark output,
binary, or large evidence is committed here. Required source validation is:

```sh
python3 -m unittest tools.test_drawingml_raw_qualification
```

The coordinator must run the pinned locked normal and allocator suites on the
assembled candidate before accepting native correctness or performance
claims.
