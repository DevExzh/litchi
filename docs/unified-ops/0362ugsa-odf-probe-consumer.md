# 0362ugsa — ODF shared-triad probe strict-check repair

The combined strict consumer check at `3d129677ce` rejects the ODF shared-triad
example's receipt helper for taking fourteen arguments, above the repository
limit of twelve. The earlier XLSX packet compiled this example but did not
run strict Clippy on all operations examples.

The helper now groups its four identity strings into a private `FamilyIdentity`.
All four route callers pass the same strings and values in the same order.
Receipt keys, proofs, assertions, allocation observation and measured closures
are unchanged. The correction does not relax the lint policy.

The diagnostic log remains external:
`/home/zhuhe/code/litchi-worktrees/.coord/budget-validation/composed-xlsx-consumers-3d129677ce/01-clippy-all-features-attempt-1792813427121951447-2920408.log`
(750 bytes, SHA-256
`5c1f66ad8ef52a63de075853d0eed937962da724e9ecf0a91a157dc74d4dda92`).
The check exited 101; the runner verified clean source, no remaining owned
builders and a stable 4,706,624,729-byte target. A new source-bound rerun is
required after this correction. No native qualification is claimed.

The source-bound rerun passes at `ecf598b6377c691196bb99fcf0e3be8166903c92`.
Strict all-target/all-feature Clippy for CLI, operations, facade and Python
exits zero in 21.137 seconds. Three post-exit observations show no owned
builders and a stable 4,706,626,475-byte target. Independent source review
confirms that all four identities, receipt keys, assertions and measured
boundaries are unchanged. [The reduced receipt](results/0362ugsa/consumer-check.json)
binds the command, source, locks, runner policy and complete external logs.
This remains one selected check, not the full 53-step gate.
