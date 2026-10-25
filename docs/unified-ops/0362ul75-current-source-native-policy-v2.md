# 0362ul75 — versioned current-source native policy v2

Status: source-only protocol preparation.  No Cargo build, native probe,
rustdoc run, timing run, or performance claim was made by this record.

This change adds an opt-in v2 policy for the PPT auxiliary routes and the
legacy-current routes.  The historical PPT auxiliary recipe and lock guard,
the historical legacy preparer/validator, and the approved f5 runner remain
unchanged.  The new adapter is in
`tools/native_qualification/current_native_policy_v2.py`; source-only receipt
preparation and consumption are in
`prepare_current_native_inputs_v2.py` and `validate_current_native_v2.py`.

The policy binds the baseline to the original source fence
`f2bbbc8d07ab5da26c180f1ff74813d39f632d54`, Git tree
`2e0d0cdf7f6a345961ceca361d53202200d00970`, tree-list SHA-256
`ce92652c439d810849b74b80e650821bc21de9029ea861cd6bbc13e38d809589`, and
the reviewed 139,441-byte lock
`1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c`.
Toolchain, Cargo configuration, profile (`release`), target
(`x86_64-unknown-linux-gnu`), probe, fixture, and ordered route pins are part
of the same fixed policy object.

The current font candidate is not frozen in this revision.  `efff87...`,
`de1f8...`, and `989da...` are recorded only as three provenance hints, with
three explicit source transitions.  The vendor files and manifests were
verified during source-only review, but that observation is not an identity;
the abbreviated values are not accepted as hashes.  Candidate
commit/tree/lock/probe/fixture fields are therefore empty in the trusted
policy, whose status is `incomplete`.  Preparation may record a blocked
diagnostic receipt, but `guard-lock`, the PPT plan, preparation admission, and
qualification validation all fail closed until a later reviewed commit fills
the exact pins.  There is no caller-provided expected digest or candidate
overlay escape hatch.

The pending trusted-policy document digest at the initial implementation was
`9d4762c41ad20ecbc409dcc2014c92828da5847eb94399387ab2000eca3c2cdf`.
The follow-up archive/route/tooling hardening changes this digest; the current
value is recorded by 0362ul75a.

The receipt envelope uses
`litchi.current-native-policy.v2`,
`litchi.current-native-preparation.v2`, and
`litchi.current-native-qualification.v2`.  Its protocol advertises the
`gate-evidence-catalog.v2` and `signed-native-protocol.v2` bindings and the
fixed `before, after, after, before` order so later adapters can coordinate
without changing historical receipts.

Synthetic verification only:

```text
PYTHONPATH=tools python3 -m unittest tools.test_current_native_policy_v2
```

Result: 5 tests passed.  The test exercises pending-policy refusal, absence of
caller digest parameters, exact route checks, blocked preparation, and a
caller route-overlay refusal.
