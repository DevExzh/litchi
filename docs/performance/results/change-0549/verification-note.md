# Verification finalization

The first strict precleanup attempt is preserved in
`verification-attempts/precleanup-01`, including its verifier snapshot.
It rejected the instruction-analysis envelope because the verifier expected
only ten validation keys while the frozen analyzer emits fourteen. The four
omitted keys assert collector instruction/self equality, positive parent
attribution, separation of direct callee counts, and setup/timed separation.
The verifier now requires all fourteen exact keys and requires every boolean
to be true. No capture, analyzer, canonical report, or gate changed.

The results review also clarifies the common guard clock: only `OleFile::open`
is timed; oracle checks and returned-value destruction occur afterward.

The second attempt is also preserved. Root created this finalization note
while its replay inventory was active; the verifier correctly detected a
bundle mutation and stopped. The retry starts with all bundle writes stopped.

The third attempt exposed a stale prior-campaign guard-input scope literal.
The verifier now requires the exact description frozen in this campaign;
script/plan hashes, timestamps, and all measurement gates remain enforced.
