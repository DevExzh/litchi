# 0779 OPC input allocation packet

Status: candidate rejected; production source restored. Paired capture and replay complete. See [the report](../../0779-opc-bounded-input-growth.md).
The source review replaces the stale ZIP-2 recommendation from 0778 with a
current, directly measured OPC owned-ingress allocation question.

## Reproduction

The probe is standalone. `build.py` materializes `probe-src/Cargo.toml` from its
path template and uses the archived standalone dependency lock. Production and
probe inputs, toolchain/host, exact commands, binary hashes and all raw reports
are retained. `before` uses base `345b81ce8a`; `after` changes only
`crates/litchi-opc/src/phys_pkg.rs` as recorded by the source census.

Capture scripts refuse to overwrite an existing lane or build record. In a
fresh evidence directory with the same source revisions and fixture identities:

1. Build the before native and allocator binaries using `build.py before`.
2. Run `capture.py qualification` and retain baseline native/heaptrack checks.
3. Apply the candidate, run `quality.py`, then `build.py after`.
4. Run `capture.py native`, `allocation`, `small-native`, `small-allocation`
   serially. Plan files give exact sample counts, CPU and process orders.
5. Capture separate heaptrack diagnostics; their receipts retain exact commands.

The initial qualification and baseline-native reports are not pooled into the
paired matrix. Supplemental small-open controls retain their own plan and
lanes. Native timing never uses allocator or profiler elapsed time.

Offline replay requires Python 3, `zstd`, retained artifacts and matching production
source (or the documented disposition-specific source witness):

```sh
python3 -B docs/performance/results/change-0779/validate.py
python3 -B docs/performance/results/change-0779/tables.py --check
```

`analyze.py` creates the initial analysis and refuses to overwrite it; the validator
recomputes it without writes.

The final seal covers the packet except itself. Source/output bytes, failed
attempts if any, and raw process reports are retained. Binary cleanup uses
exact path/size/SHA witnesses; owned temporary resource cleanup is recorded in
the report. Source fixtures live in the repository and are bound by SHA-256.

## Scope

The benchmark exercises `Workbook::open`, A1 edit/commit, and explicit NoSync
save using public APIs. Default Full durability is unchanged. It measures
warm sources and absent destinations on a shared host. Outputs are checked
outside operation regions by reopen/marker checks and against 0778's
independently checked published archive hashes. The open phase performs a
second open for post-clock verification, which is included in whole-process
RSS and heaptrack, not in the operation region.

Requested allocation bytes include realloc's new capacity. They are neither
physical copied bytes nor RSS. The candidate may retain spare input capacity;
that cost is reported separately. No CRUD coverage row, cold/range/concurrency
claim, or whole-program completion claim follows from this packet.
