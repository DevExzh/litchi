# 0817 build-driver recovery review

This is a read-only review of the failed build attempt 0 and the corrected
build attempt 1. It does not authorize or perform Cargo, binary, exporter, or
workload execution.

## Build-0 boundary

`build-0/failure.json` records a pre-Cargo failure with `cargo_started: false`
and the typed failure `reader omitted plan binaries[*].cargo_bin`. The receipt
has SHA-256
`85a3e8475f547b3ae60495e346388d9eb11741d06f42f5084d1e3c9ff9178c46` and
binds the quality receipt at
`6775012473b9fa34cdd1d9d2ac07917c2db38e36e3fe68730e9827c436d902e9`.
The archived driver is retained at `build-0/build.py` with SHA-256
`ec15cd02d694af65d67ff64e34bc8768f568c596b5c68330f0cfa8992cd29f7c`.

The plan declares `binaries[*].cargo_bin` for all three binaries. A static
diff of the archived and current driver contains exactly one change:

```python
- declared_name = declared.get("cargo_binary", declared.get("binary", declared.get("cargo")))
+ declared_name = declared["cargo_bin"]
```

The corrected driver SHA-256 is
`9fcd5997e1f28688aa3ef7bbcd8e4a85a8f657c46a0f0d624589d36470acbde9`.
The build steps, feature lists, Cargo flags, release profile, target path,
environment, and executable-copy logic are unchanged.

## Applicability of quality evidence

The quality-2 source witness, build-0 source witness, and build-1 source
witness are byte-identical (each SHA-256
`48cf724f511d5738a60881979d6cfb52565595ef8a0d292f2326c3548df928e1`). Each
contains the same production revision `953866d382c8e8248de2463ce317279fdd77d5b4`
and 9,196 production files plus the same 87 tool files. The one-line build
driver correction is packet-driver validation only; it changes no production
or Rust harness source, Cargo command, feature, profile, lock, corpus, host,
or quality test result.

Quality-2 and build-0 frozen inputs are identical, including the pre-correction
build-driver digest. Build-1 frozen inputs differ from those witnesses only at
`build.py`, whose digest is the corrected one above. Thus quality-2 remains
valid as evidence for the unchanged Rust/source state, while build-0 remains
the immutable failed-attempt witness and build-1 must carry the corrected
driver digest.

The quality receipt and quality-2 frozen inputs intentionally retain the
pre-correction driver digest
`ec15cd02d694af65d67ff64e34bc8768f568c596b5c68330f0cfa8992cd29f7c`. Final
offline analysis must preserve that historical quality witness or explicitly
record the post-quality driver-only correction; blindly requiring the current
`custody.driver_hashes()` in the old quality receipt would reject valid quality
evidence solely because this recovery changed `build.py` after quality passed.

Build-1 is still in progress at the time of this review; this document makes
no claim about its Cargo outcome or produced binaries.
