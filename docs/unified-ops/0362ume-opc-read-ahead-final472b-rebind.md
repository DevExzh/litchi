# 0362ume — OPC managed read-ahead rebind to the final API capture

This is a source-only preparation record.  It rebinds the latency-only OPC
managed read-ahead lane from the stale `49c25ef` candidate in 0362umd to the
final production API capture at `472b6705475b073e7fe08f6e11ee17fb1e29f838`.
The validated c796 packet and its external baseline are preserved; this
record does not relabel their evidence or claim a build, correctness run,
quiet-host run, timing result, allocator result, or acceptance result.

## Frozen endpoints and existing inputs

The baseline remains the exact pre-repair endpoint:

```text
revision 410bd99f00fdfb4161a94215710340f16ea536bb
tree     41850bc694ccfda6f84510611d32fe452248773d
checkout (deleted-coord-artifact)
```

Its detached sparse worktree was independently rechecked clean and reused
under the exact HEAD/tree fence.  The required `tools/perf-baseline` sources,
manifest, nested lock and Rust toolchain are materialized; omitted tracked
files remain index-only and are not treated as source changes.

The candidate is the existing clean final production checkout:

```text
revision 472b6705475b073e7fe08f6e11ee17fb1e29f838
tree     ebc1cee05f124c0c63977aff07a23288c19e7d00
checkout (deleted-coord-artifact)
```

The read-ahead probe and build inputs were checked at both endpoints before
the packet was written:

| path | baseline SHA-256 | final472b SHA-256 |
| --- | --- | --- |
| `tools/perf-baseline/Cargo.toml` | `c944031ad9afbfef854cac48d400d30fdf6135e3f2c83d105bbf014f1d3dab40` | `61e103809753ffc528b2cd2c2e00a1525d40fc4e6360668ad982bec22a60bb73` |
| `tools/perf-baseline/Cargo.lock` | `4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184` | `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c` |
| `tools/perf-baseline/src/docx_managed_read_ahead.rs` | `6592815ab69ea362d0163b4548d141c4e7e07cdd1f387f7075aaa18f207fe75f` | same |
| `tools/perf-baseline/src/main.rs` | `b51dd8de3030ee9fa9be7676b05d36a1203172b91d8b28c4bee8a07dfb65447e` | same |
| `tools/perf-baseline/src/lib.rs` | `10dd4f6f80be3a07f610f76fbbaeef50495972be62e251802b407d9699f64229` | `280d61ae1f714af78be6cecf21a4432a9ea2a27478e89c237b5d358bc02db02a` |
| `tools/perf-baseline/src/pptx_range_source.rs` | `f4b0d6171ef9a318e1d29affa0e593e6c5114f972c76dace13e572611a074921` | same |
| `rust-toolchain.toml` | `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` | same |

The final candidate's rename-bearing production paths were also inspected,
rather than applying a blind hash substitution.  The current hashes are
bound in the external packet for the five DOCX source-backed consumers, three
OPC source-backed files, the rename record and the glossary.  In those
production paths `ArtifactStreamPublication` is absent, while
`ArtifactStreamReport` is defined and re-exported by `litchi-opc`.  The
read-ahead probe itself has no reference to either result name and is byte
identical to the validated c796 probe.

## External execution packet

The new packet is separate from the preserved c796 packet:

```text
packet:
(deleted-coord-artifact)/packet.json
SHA-256 a5f2fc3a3428996949b2bf7b34732e95aed5b9857a85b541acb5b9a1a60ac3e1

commands:
(deleted-coord-artifact)/run-opc-read-ahead-final472b.sh
SHA-256 5b20c3980e6892c5485d99b28dcc455e3e823c0d8190d0cabd87fee51c6a3ddb
```

The script's concrete working directory is the packet preparation directory
above.  It binds the existing baseline checkout, the final472b source,
Rust 1.95.0, the nested locks and the existing
`docx_managed_read_ahead.rs` driver.  It selects CPU 60 as the prepared
default, but repeats the host quiet check before and throughout every
process.  Its fresh run root and baseline/candidate targets are under:

```text
(deleted-coord-artifact)
```

That root, its targets, its temporary directory and its CPU lock were absent
when this packet was prepared.  The script fails closed if the root already
exists.  The protocol remains the 0362umd forward-start/4096 and exact arms,
zero artificial delay, 30 samples after three warmups, and the same generated
source-edit-media-v1 corpus and physical-range oracles.  Allocator output is
not applicable to this latency-only lane and must not be inferred from the
ordinary binary.

No Cargo or native command was run for this rebind.  Remaining work is the
fresh locked release build, one untimed correctness preflight per endpoint and
arm, then the quiet four-leg ABBA capture with all source, report, target-size
and host receipts bound to the final472b packet.
