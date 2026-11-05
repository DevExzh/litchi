# 0362ul67c — VBA raw-policy validation and atomic handoff

Status: independent post-604 delta READY. The final source tree was validated
at `c68801c837f91a6126c61c99d3b2f73b3915d942`; the source packet starts at
`b72afa830322972e9af187f896b415e5cb02352c` and its true integration base is
`ef1acb5ef0b776c9a6066029a5cac35daeb0b5eb`. This record is the only addition
after that tested source tree. The coordinator may carry the complete lane as
one atomic commit whose parent is the true base.

## Final scoped run

The approved `190417` safe runner supervised a fresh target with offline and
locked dependency resolution, one build job, one test thread, incremental
builds disabled, debug information disabled, a 20 GB memory cap, a 19 GB live
target guard, and a 20 GB hard target guard. No broad operations suite,
native qualification, rustdoc, benchmark, or macro execution was run.

The exact final receipts are outside the checkout:

| item | path | bytes | SHA-256 |
| --- | --- | ---: | --- |
| manifest | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul67-f5-c68801c837/manifest.json` | 1,471 | `a54a2b4ce295678b67db6f89169bbfe665bc82d9722f05e30e59b3e97f52e6c4` |
| receipts | `/home/zhuhe/code/litchi-worktrees/.coord/0362ul67-f5-c68801c837/receipts.json` | 53,474 | `3b471ba182af4297ff7a446c41eb0651e71cf4af47b796848456c67996434b44` |

The final six rows all passed. The focused test counts are 48 VBA tests,
5 PPT VBA-writer tests, and 1,268 PPT library tests with 1 ignored test.
The final target reached 1,231,682,510 bytes; every row stayed below the
19,000,000,000-byte live guard. The runner, toolchain inputs and locks were
unchanged throughout:

* runner SHA-256: `1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`;
* workspace lock SHA-256: `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`;
* perf-baseline lock SHA-256: `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`;
* native-resave lock SHA-256: `51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378`.

The final row log hashes, in receipt order, are:

```text
82869a77e192e16cd83be3793d6426481b56e0cf91ec5271545d0cda5f34d9bb
f67c2f8714e120467c75f733a8fd6cfa1e02acb3775683c4e698711c4e9283d4
c715adfd33d663c963ca0a538aa54814d088f7e245926f0bf16c87edcfeeff7f
ae04c2af1f690418c01ab286830d388dbe36513249a8507e573a4f23ae6a56ed
933ff75d605522d6911d9d8bc30a7866b0dfa8005d3197429488741392d33e48
9985e02450d3e7482b7c9f06e6937325df966e0bb3c1cab36d9290d154669113
```

## Preserved failed states and repairs

All three compile-failed states remain outside Git for independent review; the
passing rerun did not overwrite them:

| state | failed row | receipt SHA-256 | failed log SHA-256 | repair |
| --- | --- | --- | --- | --- |
| `/home/zhuhe/code/litchi-worktrees/.coord/0362ul67-f5-604bce4fc5` | PPT VBA writer | `61911c4d3d12898f0929592304db7c9859f2e9c01649f65e888bd357e9edf15f` | `86370386d6272f6d10b36560166694366f4a1e8b90fd09cd7faa926d772f3382` | removed invalid `const` declarations around retained policy access |
| `/home/zhuhe/code/litchi-worktrees/.coord/0362ul67-f5-0ca559ca6b` | affected all-target check | `7305bfa913c423cbaefb4fc5b1275195516068d5924b6fc56800f2320c748844` | `195c112ca386621c354fdd06d781c3c9dc82479175f920ae2bcbd9febbfe9a91` | migrated the scoped PPT correctness example to public APIs and current digest formatting |
| `/home/zhuhe/code/litchi-worktrees/.coord/0362ul67-f5-6999f2669f` | no-default VBA feature boundary | `00ae3b710796445e0444dab487d655386ee6eeb412121fee39d47a5a34c96066` | `a1f3c87d4454e61915b31405d4ecb2fdbe65b179d869eeeb2debf859d2e4edb0` | passed the selected `SourceLimits` to both `CurrentUser::parse` call sites |

The final `c688` source tree is identical to the tested tree except for this
record. No evidence, executable, or third-party checkout path is committed.

Coordinator composition keeps the independently reviewed public
`Package`/`Presentation` correctness-example projection from this lane and
the tested parser closure spelling. The latter is equivalent to 0362ul39a
(the bound argument is named `data` instead of `bytes`). The entire PPT
presentation implementation and example remain byte-identical to the tested
leaf. Glossary resolution retains all existing font/ODF/other profiles and
adds the complete VBA profile. The timed benchmark paths are unchanged.
