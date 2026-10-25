# 0362ul87ah — common properties clear-helper compile repair

The replacement F5 attempt for `a3eff9e1a64e1d303628971167c276a9a6f38a8b`
(tree `3c81061b13ab1e655ddc244ae2eb7361f20aa857`) passed its scoped
formatter row and stopped at row 1, `check-common`. The preserved attempt
state is:

```text
/home/zhuhe/code/litchi-worktrees/.coord/0362ul87a-f5-a3eff9e1a6/
```

The row-1 log exited 101 after 41.5067 seconds because
`properties.rs:567` still called the removed unqualified `clear(package)`:

```text
error[E0425]: cannot find function `clear` in this scope
crates/litchi-ooxml-common/src/properties.rs:567:22
```

Receipt identities are retained here: `result.json` SHA-256
`ffe27fca0bc00dcbb3ff61e95004bfdfc94539b6974aca9e5be96672c261cce9`,
`rows.json` SHA-256
`2417c0b18f8669fbef489d010558948c39f6ffc6e784c1c07569be9a4e4668b6`,
the compiler log SHA-256
`029639ec02ad68ec8c7851390781263fe3a2138a40dc0b447d616ee503aec5c0`,
and the target trace SHA-256
`feed36d656ae932e9dbb8fef0bfea581dbcb30369448a8a62240d4cb22bf49a7`.
The target reached 158,558,033 bytes; no resource or ownership guard fired.

The repair changes only the stale call to the existing private codec helper:

```text
None => (write::clear(package)?, None),
```

An audit found no remaining cross-crate call to the removed free publishers;
the only other `clear`/`write` calls are private codec tests or the shared
Edit implementation. `git diff --check` and scoped rustfmt pass. This record
does not claim a passing replacement F5 run; a fresh source-fenced state is
required before the next coordinator-approved launch.
