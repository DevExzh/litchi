# 0362ul9 — CFB composition and remaining raw-codec review

Date: 2026-10-24. Source integration remains on the coordinator staging branch.
Only redaction audit 22 is deferred. Nothing is pushed.

## Accepted integration

CFB leaf `bf7f48741bc02d1f5e736a7a732c25808d28af36` merges at
`f1d63257e6f1bc0573887e6fb956a2a1ddc0f8d8`. Records
[0362zad](0362zad-cfb-explicit-limits.md) and
[0362zada](0362zada-cfb-review-returns.md) document the required raw limits,
every consumer migration, four retained-policy review returns, and composed
validation. The first two composed checks expose missing future VBA error
fallbacks and a chart test qualification; both are repaired. The final
all-features/all-targets workspace check passes on index tree
`57abd651f19bc0ee91f24e1f60d2850a41eb7f6c` with root lock `4100b1fbb272`.
The reduced receipts preserve every attempt, source tree and log hash.

The independently reviewed XLSX raw qualification protocol
`5e6555d8406709382e26ecbcdd6cb18ded449e31` merges at `95ce396d83`.
Records [0360adca](0360adca-xlsx-raw-native-protocol.md) and
[0360adcaa](0360adcaa-xlsx-raw-protocol-review.md) bind the true baseline,
real worktrees, fixed input hashes and exact workload/fidelity checks.
All 37 composed Python tests pass. This is preparation, not a native
measurement or performance acceptance.

## Required source work and review returns

DrawingML's owner/typed-error/part-policy packet at `0b8e1744d1` passes
source review after migration of three live report Cargo harnesses. Their
focused compilation receipt remains required before integration. Existing
historical measurements and receipt hashes are not rewritten.

The raw-codec audit finds additional work under 0351 §2.5.3. Five isolated
lanes cover chart; diagram; text, SVG and model3d; theme and ink; and geometry,
color, transform and chart formatcode2. Fixed hard ceilings do not remove a
raw XML codec's required caller policy. Typed in-memory scalar constructors
remain domain conversions; that distinction does not exempt raw read/write
functions. Every lane includes its production, test, example, fuzz and live
report-harness consumers.

Text/media tip `d2c420c26d` passes 238 DrawingML tests, 22 SVG tests and four
consumer checks, but is returned before merge. Review finds a remaining
PPTX model3d raw default twin, child/fragment allocation before policy
admission, per-child node-budget resets, changed SVG contextual bytes,
incorrect output-limit attribution and text decoding before admission.
These are required correctness and fidelity repairs, not accepted exceptions.

PPT native protocol `2fa0f5b538` is also returned: object, external-media and
picture-bullet routes still lack meaningful nonempty owners, and source,
lock, probe, binary and receipt mutation fixtures are incomplete. Its
compile-only mini-crate check is not native qualification. The legacy
protocol at `1324089b4a` retains a separate return: the property baseline
`604822b441` already has the migrated object owner, so the frozen probe's
legacy-owner branch cannot compile. It needs a reviewed consumer adapter
for the current-owner/legacy-property boundary, plus the candidate's new
explicit CFB limit argument. The baseline source and workload remain fixed.

## Cleanup and remaining acceptance

Cleanup thirty-nine removes the merged legacy-raw, ZIP and exact-source-role
worktrees non-forced. All three branches and eleven local payloads (566,586
bytes) remain preserved. Restoration is exercised; the coordinator verifies
payload hashes/modes, branch ancestry, removed paths and the protected trunk
probe. See the [cleanup record](0362-worktree-cleanup-2026-10-23.md).

After the raw source repairs and reviews, the assembled revision still needs
fresh all-68 rustdoc capture and exact inventory dispositions, the full
53-step gate, actual OfficeCLI corpus replay, glossary closure and native
qualification. The historical isolated documentation pass on `49c25ef67f`
remains evidence only for that revision. The 17-lane native catalog is being
checked against the newly changed DrawingML paths and all required OLE
owner/property/history workloads so no earlier performance obligation is
silently lost.
