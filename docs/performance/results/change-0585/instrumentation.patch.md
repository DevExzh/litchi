# The measurement instrumentation, and why both legs share one binary

The chain-link counts in change 0585 were taken from a **single instrumented
build** carrying two runtime switches, not from two separately compiled trees.
This note records exactly what was applied, so the measurement can be repeated,
and states why the shortcut is sound for this particular quantity.

## Why one binary

A chain-link count is a count of `next_chain_sector` calls. It is exact,
deterministic, and independent of optimization level — the same walk performs the
same number of dependent loads whether or not the compiler inlines around it. It
is therefore the one quantity that does **not** need the program's usual paired
detached-worktree release builds, and measuring it from one binary removes the
risk that two builds differ in some way other than the change under test.

This applies to the link counts only. It would **not** be sound for instruction
counts, cycles, or wall-clock, and none of those are claimed by change 0585.

## The two switches

Applied to `crates/litchi-cfb/src/shared.rs`, removed before commit:

```rust
use std::sync::atomic::AtomicBool;

pub static CHAIN_STEPS: AtomicU64 = AtomicU64::new(0);
pub static COLD_LEG: AtomicBool = AtomicBool::new(false);
pub static CURSOR_COLD: AtomicBool = AtomicBool::new(false);

fn next_chain_sector(table: &[u32], sector: u32, table_name: &str) -> Result<u32, OleError> {
    CHAIN_STEPS.fetch_add(1, AtomicOrdering::Relaxed);
    // ... unchanged body
```

`COLD_LEG` is consulted at the top of `StreamChainHint::resume_from`:

```rust
if COLD_LEG.load(AtomicOrdering::Relaxed) {
    return (start_sector, 0);
}
```

`(start_sector, 0)` is exactly what `resume_from` returns when no hint applies,
so setting `COLD_LEG` reproduces the unhinted walk for **every** caller —
including change 0579's `GlobalsBuffer::ensure`.

`CURSOR_COLD` is consulted only at the two `stream_cursor_at_hinted` call sites,
leaving 0579's `read_stream_range_hinted` resumption in place:

```rust
let resume = if CURSOR_COLD.load(AtomicOrdering::Relaxed) {
    (start_sector, 0)
} else {
    hint.resume_from(self, sid, false, start_sector, ordinal)
};
```

Both statics were re-exported from `crates/litchi-cfb/src/lib.rs` for the probe's
benefit. All of it was removed afterwards and the three production files verified
byte-identical to their pre-instrumentation sha256 digests.

## The three legs

| leg | `COLD_LEG` | `CURSOR_COLD` | reproduces |
| --- | --- | --- | --- |
| `pre-0579` | true | — | the walk before change 0579 |
| `0579-only` | false | true | HEAD before this change |
| `this-change` | false | false | HEAD with this change |

The `pre-0579` leg exists as a **control**, and it validated the whole setup: it
reproduces change 0579's recorded pre-change figure of **5,796** chain links for
an open of `ConditionalFormattingSamples.xls` exactly, and the `0579-only` leg
reproduces that record's post-change **2,099** exactly. A switch that reproduces
a previously published number on both sides is a switch that is measuring the
thing it claims to measure.

`0579-only` → `this-change` is the delta change 0585 reports. Quoting
`pre-0579` → `this-change` instead would credit this change with 0579's saving,
and the record deliberately does not.

## The probes

Two throwaway examples, `crates/litchi-xls/examples/chain_probe.rs` and
`crates/litchi-doc/examples/doc_chain_probe.rs`, both removed before commit.
Each opens a fixture through `litchi_core::OwnedSource`, runs the scenario once
per leg, and prints the link count at the end of the open and at the end of the
scan, together with an FNV-1a digest of the extracted text. The digest is what
makes the differential claim checkable: if a leg produced different text, the
run would say so rather than the reader having to trust that it did not.

`chain-legs.tsv` and `doc-chain-legs.tsv` are those probes' raw output, one line
per fixture per leg. `chain-summary.txt` and `doc-chain-summary.txt` are folded
from them and contain nothing the raw files do not.
