# 0362ul100 — Reviewed runtime repairs composed

The independently reviewed runtime repairs are now composed privately.
PPTX merge `37c99f4915d9e4efb712a08ab14972bf7504b68c` preserves the
comments chain through `43dbfabf2025d2913f99649b1bb170a8550645e3` and the
normalized notes/consumer/harness parent `19c8f3d2bc`. Its current tip is
`fc5ef75a9e8de76e950ebe5b1230c89074e5fd57`, tree
`7b14c8698730685c64299078b8d6c74068e80b16`.

OOXML merge `668874d68b499c6a7afd9b7da25b2dbc365cce67` preserves all
three parents: prior VBA/Custom XML repairs `8c766ac5c6`, properties
`8da5464633ef4085cd2e2a234e3f463986cf4395`, and ribbon
`19ce1f092c202bacebe1c0129a9f66ea503108fa`. Current tip
`b73ee18a69be4122b9097588f146e2c3b3b8b713` has tree
`227d669ec5364facf088eac1d447dc6c5d6447cc`.

Both compositions are conflict-free. Comments now align reader and writer
aggregate/physical budgets, admit raw attributes before MCE, validate selected
fields before ownership and preserve date-time. Properties classify errors by
provenance while validating all source events without the returned duplicate
scans/clones. Ribbon preserves strict source-splice proofs and adds bounded,
atomic authored XML publication; canonical part names key provenance and
private insertion keeps its cheap path. Records 89ay and 94u describe the
complete reviewed changes and their consumer/regression coverage.

## Formatting and evidence correction

The preceding eight-row PPTX compiler packet on `cc2ab3f79d` stops at its
first formatting row after 5.662 seconds, with target size zero. No Cargo row
runs. The failure and complete root verification are recorded in 89ax; its
formatter-only repair `19c8f3d2bc` normalizes harness module order and one
closure layout. The failed packet remains unchanged.

The root composition script subsequently finds four unformatted comments
paths. The merge commits were made before that failure was handled, so their
original formatting statements were premature. Follow-ups `fc5ef75a9e` and
`b73ee18a69` correct those statements. Root applies repository formatting to
all five changed PPTX Rust paths and verifies them; all thirteen changed
OOXML Rust paths already match and pass the actual check. The source correction
is formatting-only. External formatter patches are retained in
`.coord/composition-format-89ay-94u/`; no passing compiler/runtime result is
inferred from these checks.

## Equivalent measurement harness

Baseline `84ba3b5c33a25b3ff2af1303c812679d16ae2a83`, tree
`be3b4ed6d92b86b30c4e1747194d41794bf34a08`, receives the identical harness
formatting. Root verifies all eight harness tool blobs match final PPTX
`fc5ef75a9e`, and the baseline production crates, root manifest/configuration
and fixtures remain byte-identical to `76371266514feb010ef7478c26e963debcc7fa59`.
This endpoint is never merged into trunk. The external proof is
`.coord/composition-format-89ay-94u/baseline-harness-proof.json`, 988 bytes,
SHA-256 `c95b06b25567474efe626f07ad86c502b06c986d94e27612e75135f06ba489e5`.
No benchmark or performance acceptance is claimed.

Fresh source-bound compiler/runtime/glossary packets are being prepared.
The prior PPTX 1,374-pass / 12-failure and common OOXML 511-pass / 17-failure
captures remain historical. The repairs require actual reruns; source review
alone does not discharge them. Final gate, inventory dispositions, native
fidelity and quiet ABBA remain open. Only audit22 is deferred; nothing is pushed.
