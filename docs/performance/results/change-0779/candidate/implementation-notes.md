# 0779 `read_limited` candidate

This directory contains a source patch only. The live
`crates/litchi-opc/src/phys_pkg.rs` was not edited by the candidate owner, and
no Cargo command, formatter, native producer, allocator capture, or profiler
was run here. The patch was applied to an isolated temporary checkout only to
regenerate and validate the unified diff; `git apply --check` passes against
the unchanged 0779 source.

The candidate keeps the existing 8 KiB read buffer, the initial bounded
reservation, the `Interrupted` retry loop, the late-I/O error path, the
malicious `Read` count check, and the exact one-byte limit probe. After each
successful read it computes `required = len + read` with checked arithmetic.
Only when `required` exceeds the current capacity does it request a target of

```text
max(required, capacity + max(8192, capacity / 8)).min(maximum)
```

The reservation uses `target - len`, rather than `target - capacity`, because
`Vec::try_reserve_exact` measures its additional argument from the current
length. The requested target is therefore at least the required payload, is
never above the input limit, and leaves at most 8 KiB or 12.5% of the prior
capacity as the policy slack. A checked subtraction reports an impossible
internal capacity state as `InvalidData`; it does not alter any reachable
input refusal ordering.

The focused tests in the patch record payload bytes and every requested read
length, exercise growth beyond 64 KiB, short reads, an initial `Interrupted`,
a late I/O error, an overreported read count, the exact-limit one-byte
sentinel, and the geometric target at small, geometric, near-limit, and
arithmetic-overflow boundaries. They also assert the typed error identity and
message for the late I/O case.

This is a bounded-growth candidate without a compacting copy. The allocator
may return more capacity than the requested target, so the bound is on the
requested capacity and logical retained slack; it does not claim a physical
allocator or RSS bound. The candidate also does not claim that cumulative
allocation-request bytes equal copied bytes. Those questions require the
coordinator's before/after allocator and memory evidence.

The pre-edit source and accepted architecture inputs were checked before
preparing the patch:

| Input | SHA-256 |
| --- | --- |
| `crates/litchi-opc/src/phys_pkg.rs` | `3dfa9ec92c0d090b0bc753ecacc63f263975a470c267b79b03ae1c5898ac3ab8` |
| `docs/adr/0005-io-memory-and-performance.md` | `8279fd2dc0b074aa104bf0c83df6187d12e18756cd8c8d7f25750b2a997cb0aa` |
| `docs/adr/0006-validation-security-and-compatibility.md` | `ae21189eda0acc9524a7c5a57e87bb3566405ede5c015c87105e01337666f44b` |

All 35 entries in `../architecture-inputs.json` matched their recorded
hashes before editing the candidate evidence directory.
