# 0362ugjad — PPT native probe correctness preflight

Packet `079ad7e3bc277d16ce5cf165e612280fa15b85f0` closes the consumer-reporting
return in 0362ugjac. Repair `416f57cd7a` preserves an unavailable optional
patch endpoint as JSON null; a present endpoint must still match exactly.
The coordinator reviewed the source diff, and independent review verified
the real preparation receipt, build logs, binaries and correctness reports.

Both rebuilt normal release probes pass `--mode verify`: 28 routes across
four fixtures. The 12,094-byte reports are identical, SHA-256
`2daa5587982ac38c4a4068adc03a9c91798a4c98ce95b050e541069ed276d07c`.
The preparation receipt is 12,580 bytes, SHA-256
`0bbb366e920949da8f212b7a0bb2ad1c4499054cd810f5b64605078f58e573a9`.
Record 0362ugjab contains the exact source, lock, probe, binary, log and
target identities, including the preserved earlier failures.

The conflict-free composition passes all ten combined Python modules:
94 tests in 26.914 seconds, 27.168 seconds including process setup. The
tested source index is `4f07dad478cac8411b1af725dce9e9876cf6465b`.
The complete command and log identity are in
[`results/0362ugjad/composed-tests.json`](results/0362ugjad/composed-tests.json).

This accepts the probe repair and paired correctness preflight. It does not
accept native performance: allocator builds, calibration, quiet ABBA and
final source-bound qualification remain required.
