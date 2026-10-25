# 0362ul66f — Numbers-wire malformed fixture repair

Status: source READY for the next scoped F5 retry.

The first actual F5 run at `e118141056cfe6a5a8729ab74c48d215873e97e3`
passed the Numbers-wire check, strict clippy, and 84 of 85 Numbers-wire
tests.  The remaining test failed while constructing its deliberately
malformed owned-cell fixture: `BncCell::encode` now correctly rejects a
three-byte fixed-width field before producing output.

The test now verifies the failed Duration metadata transition is atomic by
capturing the deliberately invalid private field map and prefix directly.
This keeps the malformed-shape coverage without asking the encoder to emit a
wire value that the parser must reject.  Production validation and wire
fidelity are unchanged.

The failed row remains preserved in the external F5 state directory; no later
row was started.
