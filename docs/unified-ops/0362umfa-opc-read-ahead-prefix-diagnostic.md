# 0362umfa — OPC final472b prefix diagnostic receipt

This follow-up binds the one allowed build/correctness prefix execution for
0362ume/0362umf.  It is diagnostic only: there is no quiet-host timing,
ABBA, allocator, profiler, or performance-acceptance claim.

The complete reduced receipt is external because the build logs and binaries
must remain outside Git:

```text
/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/
opc-read-ahead-final472b-prep/prefix-diagnostic-receipt-fixture-rerun.json
bytes 8851
sha256 1559c13d0a4cd05eef6b3defaab304a3ecfa55b4b504ce8b2d15d3ea20f2f6e3
```

## First attempt and finite sparse repair

The first exact-prefix attempt used the frozen baseline and exited 101 during
the baseline build.  The compiler reported only these missing inputs:

```text
test-data/poi/test-data/spreadsheet/54016.xls
test-data/rtf/watermark.rtf
```

Both paths are tracked at the frozen baseline commit; they were absent only
because its sparse checkout contained `/Cargo*`, `/crates/**`, and `/tools/**`.
The failed root and log remain preserved externally at
`.../opc-read-ahead-final472b-run` (12,246-byte log, SHA-256
`3dddd2bfbac0270144acc07e2b797f1ad188f6187db8ff0e518f708645b58ba3`).

Only those two tracked paths were added to the sparse materialization.  The
baseline stayed clean with HEAD/tree
`410bd99f00fdfb4161a94215710340f16ea536bb` /
`41850bc694ccfda6f84510611d32fe452248773d`; their bytes and hashes are:

| path | bytes | SHA-256 |
| --- | ---: | --- |
| `test-data/poi/test-data/spreadsheet/54016.xls` | 984576 | `2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a` |
| `test-data/rtf/watermark.rtf` | 69471 | `48d62dcd959e737b06ebb8255780bcaaf1e88056ff9c3d5a21d3ff5cd3ddf9cb` |

These hashes match the final candidate checkout.  No production source or
commit was changed.

## Corrected prefix result

The unchanged prefix SHA-256 is
`9729d18cc321e8954305dd49022c913740474df6d11a2520530c2eb07e385266` and
the original all-in-one recipe remains unchanged at
`5b20c3980e6892c5485d99b28dcc455e3e823c0d8190d0cabd87fee51c6a3ddb`.
The corrected run used fresh root
`.../opc-read-ahead-final472b-run-fixture-rerun-20261024T193000Z`, exited 0,
and stopped before `host_snapshot`/`run_legs`.  Both release builds exited 0:

| endpoint | build log (bytes, SHA-256) | binary (bytes, SHA-256) | target bytes |
| --- | --- | --- | ---: |
| baseline | 11248, `827ac2b34e8a295b3b745745dbf6e8048db90d8b82d4667ca712582edc73feac` | 101339896, `acbc53ac9c05edd156fe31fe9650675ffe57322b3ba4e3943bf94dab0d6bd7b2` | 1403891794 |
| candidate | 39797, `51eae81783bfec018be31805d58bca5f738eb86b0fc46d7da65c2e31c7ed263a` | 101711056, `8acb79ffb76538af0ca67f71b81f88dec34fd572eaa38ddd90cfc5db71fa5092` | 1414209073 |

The four one-sample, zero-warmup correctness reports all exited 0 and passed
the production schema/oracles.  Their external files are 10,856–11,925 bytes
each; the receipt binds every path, byte count, and SHA-256.  No timing result
is inferred from those reports.

The timing guard implementation is committed separately as `fa6562ad17` and
has been returned for lifecycle, monitoring and provenance repairs. It is not
part of this accepted build/correctness record. Those repairs require review
before the coordinator schedules an exclusive quiet-host timing window; the
original all-in-one script must not be used.
