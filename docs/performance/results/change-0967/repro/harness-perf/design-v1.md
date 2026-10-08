# 0967 positive control performance harness

Status: frozen control-harness design, 2026-10-07. This is a separate
external harness; the functional control probe under `harness-control/` is
sealed and unchanged. One-sample smoke controls have checked the lifecycle
and output gates; no formal native or allocator capture has been run.

The harness binds one fixed input before measurement and runs the public DOC
lifecycle in a dedicated `#[inline(never)]` measured wrapper. Input generation,
file reading, source identity checks, reference-route preparation, correctness
oracles, per-sample output hashing, and JSON serialization are outside that
wrapper. The allocator lane uses the same wrapper and operation closure but
reports no elapsed-time claim.

The two binaries have the same source and lockfile:

```text
doc_revision_control_profile_0967
doc_revision_control_profile_0967_alloc
```

The CLI is:

```text
--variant baseline|candidate \
--control generated-positive|dense1024|false-marker|real-positive \
--input PATH \
[--edits 1|32 --retention default] \
--warmup N --samples N
```

`generated-positive`, `dense1024`, and `false-marker` use `scope=edit`, require
`--edits` and `--retention default`, and exercise one `Edit` with paragraph-0
replacements, `commit`, and an output `Vec` copy. `real-positive` uses
`scope=query`, rejects edit flags, and measures a fixed no-mutation route:
bounded snapshot open, public revision query, paragraph query, output-byte
copy, and public `Package` reopen. Its existing paragraph-0 tracked-text
refusal is not relaxed or mutated.

The timed scope is exactly four cells, all with `retention=default`: generated
positive with `edits=1`, dense1024 with `edits=32`, and false-marker with
`edits=1` and `edits=32`. `real-positive` is a functional query/reopen control
only and is never part of the timed or allocator-capture matrix.

Native defaults are `--warmup 5 --samples 80`; its `elapsed_ns` values are
the measured native region and `timing_claim=true`. The quality smoke receipt
is noncomparative, while staging controls any formal claim. The allocator
binary requires `--warmup 0 --samples 1`;
it measures the same operation closure and reports `allocation_calls`,
requested/allocated bytes, deallocated bytes, peak live bytes, and retained
bytes. Any allocator setup before the region is excluded from the region
delta, while allocations freed inside the region remain accounted for by the
process-global counter baseline. Output capacities are reported as
`logical_vec_capacity_bytes`, the Rust `Vec` logical capacity; they are not
allocator-byte measurements.

The report schema is:

```text
litchi.performance.0967.control_profile.v1
```

Every report includes `lane`, `variant`, `control`, `scope`, source bytes and
SHA-256, warmup/sample counts, `timing_claim`, `allocator_timing_claim`, a
literal region definition, reference output identity, exact semantic oracle
status, and `samples`. Native samples carry `elapsed_ns` and no allocator
object with `timing_claim=true`; allocator samples carry no elapsed samples,
one allocator object, and `timing_claim=false`.
The query scope carries fixed revision count and public-reopen status; the edit
scope carries edit count, output identity, paragraph/revision readback, and
reference-output equality.

The fixed control inputs are the external files already qualified by the
functional harness:

| control | bytes | SHA-256 | scope |
| --- | ---: | --- | --- |
| generated-positive | 15,360 | `47e6f76319a8086ebc36719bc9f1d894ad907a02856cfb842ac198d7879d8737` | edit |
| dense1024 | 93,696 | `f29dc1e3e1c05dcf37ce0843551a21ca9423cb283d54599813b21d52fcd76def` | edit |
| false-marker | 15,360 | `67cc08d9c17b3e8336327b51ecd78e6f44b14929037e877559b451918159f4ee` | edit |
| real-positive | 66,048 | `f8c19b5712e12244af5e243ab6b996af6bc32193dad208948a155c4a4d5bb3dd` | query |

The edit controls support `edits=1` and `edits=32`, with fixed equal-length
texts. The generated source remains outside every measured sample. The
staging_protocol owner supplies final paired build manifests and reducer
checks against this schema; this directory owns only the matched control
source and functional quality checks.
