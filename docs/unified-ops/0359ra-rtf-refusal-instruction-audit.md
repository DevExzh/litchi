# 0359ra — RTF refusal instruction audit

- Date: 2026-10-22.
- Status: checkfree; measurement only, with no implementation accepted.
- Scope: the exact refusal-contract probe and the actual 0359b release
  binaries. This record does not qualify the 0359b native timing result as
  noise.

## Deterministic comparison

Callgrind 3.26.0 ran each existing 0359b binary once on CPU 7 with
`taskset -c 7`, `--tool=callgrind --dump-instr=yes --collect-jumps=yes` and
the probe's existing JSONL workload. The profile collected `Ir` only. No
setup or driver work was subtracted. The binaries were not rebuilt and no
new large target was created.

The inputs are the same release binaries used by the native A-B-B-A result:

| Input | Size | SHA-256 |
| --- | ---: | --- |
| 0359b before | 4,124,032 B | `4693bd12232d92a61ef2f437de2d623e345a106288d5d0a5c5638cf27677420a` |
| 0359b after | 4,137,976 B | `0ca9e98098c9e687737c9f446019cf5108f17c3b024cf0f565890d26ff61ec96` |

The source probe is
`crates/litchi-rtf/examples/refusal_contract_bench.rs`, SHA-256
`9d9e91c639e47ccc94f154464f6f97a0dc0002a3b2fb37251f61b17bf71cdb22`.
Its scenario order and loop counts are unchanged: `open_1k` (1),
`noop_1k` (100), `text_commit_1k` (1), `text_commit_save_1k` (1),
`replay_1k` (100), `tail_commit_1k` (1), `unsupported_content` (1,000),
and `protected_commit` (100). Each scenario still takes 34 samples, drops
the first three and reports sample 15 of the sorted remaining 31. The
successful, replay and protected-error dispositions are the same assertions
as 0359b; the unsupported route still returns an error for every empty-span
underline request.

The complete profiles report 627,525,400 instructions before and 627,424,214
after, a decrease of 101,186. This is a deterministic profile fact, not a
native timing result. It demonstrates why the aggregate instruction count or
the whole-probe elapsed time cannot be used to waive the 0.445 ns native
refusal-route increase.

## Refusal route

The empty-span underline route calls
`litchi_rtf::edit::Edit::set_text_underline` 34,000 times (34 samples ×
1,000 iterations). Its direct function cost is:

| Binary | Direct `Ir` | Calls | Direct `Ir` per call |
| --- | ---: | ---: | ---: |
| Before | 1,360,000 | 34,000 | 40 |
| After | 1,496,000 | 34,000 | 44 |

The exact increase is four instructions per refused call, or 136,000
instructions in this function. The corresponding before/after disassemblies
show the fixed reason load and the operation/reason pointer-and-length stores
needed to return the shared `Refusal`; the old RTF-only payload stored only its
compact discriminator and payload. This identifies a concrete representation
cost on the route that produces the refusal. It does not claim that every
other profile delta has that cause.

## Other routes in the same probe

These are direct function costs from the same two profiles. The two anonymous
candidate blocks retain the baseline IDs; their names are the baseline
`cfn` labels in the raw profile.

| Function | Callgrind ID | Calls | Before direct `Ir`/call | After direct `Ir`/call | Delta/call |
| --- | ---: | ---: | ---: | ---: | ---: |
| `Package::from_bytes_with_options` | 1066 | 103 | 46 | 46 | 0 |
| `Package::edit` | 1132 | 37,468 | 23 | 23 | 0 |
| `Edit::replace_text` | 1002 | 3,469 | 163.105 | 164.105 | +1.000 |
| `Edit::commit` | 1018 | 6,869 | 39.040 | 39.040 | 0 |
| `Edit::commit_staged` | 1020 | 6,869 | 514.847 | 517.392 | +2.545 |
| `tail_append::Edit::append_text_paragraphs` | 1140 | 34 | 68 | 68 | 0 |
| `tail_append::Edit::commit` | 1146 | 34 | 429 | 439 | +10 |
| `Package::to_bytes` | 1134 | 34 | 78 | 78 | 0 |
| `Package::apply` | 1138 | 3,400 | 74 | 74 | 0 |

The extra work in the normal edit and tail publication functions is separate
from the four-instruction refusal delta: these are other compiled blocks with
different code layout/code generation in the before and after binaries. The
profile establishes the route-local refusal cost and the presence of unrelated
changes; it does not assign those normal-route deltas to `Refusal` stores.

The compact callgrind receipts are retained outside the repository:

| Artifact | Size | SHA-256 |
| --- | ---: | --- |
| `callgrind-20261022/before.callgrind` | 512,449 B | `f9b8459e0983b6e5e9dcdb45a3f76c8ee5e308733f1b0e025cbd6f227242bc6b` |
| `callgrind-20261022/after.callgrind` | 512,765 B | `fd867758cbaf471f3561a8a4a0dfc488fc39b35084468e0ead6a9e7d0e719334` |
| `callgrind-20261022/before.stdout` | 434 B | `6482df0e82dc8b8d593ee96a7a4260f12f50b4dec5d8799a6ebbc2de2da1dfd3` |
| `callgrind-20261022/after.stdout` | 428 B | `5d1fd0662737f8d79a484f653d1e6b680197aa3ea07f532a40d1a2317c96d7d6` |

The JSONL output retains all eight scenario labels and dispositions. Its
under-Valgrind medians are workload receipts only and are not native
qualification data.

## Native qualification still required for a future optimization

There is no current implementation candidate to qualify. If an aligned
representation or RTF-local producer optimization is reopened, the owner
must rerun the existing native workflow after the source and consumers are
held at the candidate commits:

1. Build each side with the 0359b command, using separate external target
   directories and the capped cargo environment:
   `cargo build --locked --release -p litchi-rtf --example refusal_contract_bench`.
2. Run the exact retained driver
   `(deleted-coord-artifact)`, which
   performs three CPU-7 A-B-B-A trials, six run medians per case, 34 samples,
   three warmups and 31 measured samples. Preserve the JSONL runs, summary,
   process monitor and both binary hashes.
3. Repeat the paired allocation check with
   `heaptrack -o <external-prefix> <binary>`, then summarize with
   `heaptrack_print -f <prefix>.zst -p 0 -a 0 -T 0`. Retain the compressed
   captures outside the repository together with their sizes, hashes and
   regeneration commands.
4. Rerun the relevant core and complete RTF tests, the refusal-kind and fixed
   reason assertions, corpus/fidelity checks, and the API glossary gate before
   any timing result is accepted.

The retained 0359b release binaries, native JSONL receipts, heap summaries,
and this deterministic profile are sufficient to reproduce the audit. No
native threshold waiver is made here.
