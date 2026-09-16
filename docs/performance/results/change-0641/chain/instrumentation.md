# The chain-link instrumentation, and why all three legs share one binary

Change 0641's chain-link counts were taken from a **single instrumented build**
carrying three runtime switches, following change 0585's method and for the same
reason: a chain-link count is a count of `next_chain_sector` calls, exact,
deterministic and independent of optimization level, so it is the one quantity
that does not need paired detached-worktree release builds, and measuring it
from one binary removes the risk that two builds differ in some way other than
the change under test. It would **not** be sound for instructions, cycles or
wall clock, and none of those are taken this way here.

## The switches

Applied to the worktree, removed before commit. `crates/litchi-cfb/src/shared.rs`:

```rust
pub static CHAIN_STEPS: AtomicU64 = AtomicU64::new(0);

fn next_chain_sector(table: &[u32], sector: u32, table_name: &str) -> Result<u32, OleError> {
    CHAIN_STEPS.fetch_add(1, AtomicOrdering::Relaxed);
    // ... unchanged body
```

re-exported from `crates/litchi-cfb/src/lib.rs`. In
`crates/litchi-xls/src/workbook/source.rs`:

```rust
pub static WORKSHEET_COLD: AtomicBool = AtomicBool::new(false);
pub static SHARED_HINT: AtomicBool = AtomicBool::new(false);
```

`WORKSHEET_COLD` is consulted in `WorksheetScan::new`:

```rust
let cursor = if WORKSHEET_COLD.load(Ordering::Relaxed) {
    cfb.stream_cursor_at(path, start)
} else {
    cfb.stream_cursor_at_hinted(path, start, chain)
}
.map_err(SourceBackedError::from)?;
```

`stream_cursor_at` is `stream_cursor_at_hinted` with a fresh hint, so setting
`WORKSHEET_COLD` reproduces the base's cold construction exactly, for every
caller, while leaving change 0585's shared-string resumption in place.

`SHARED_HINT` is consulted around each sheet in `write_text_to_impl`:

```rust
if SHARED_HINT.load(Ordering::Relaxed) { sheet_chain = strings.chain; }
let collected = scan_text_sheet(.., &mut strings, &mut sheet_chain)?;
if SHARED_HINT.load(Ordering::Relaxed) { strings.chain = sheet_chain; }
```

`StreamChainHint` is `Copy` and holds one position, so copying it into the
worksheet cursor before a sheet and back into the resolver afterwards models
**one** position serving both consumers exactly. That is the arrangement change
0585 argued against in prose without measuring it.

## The three legs

| leg | `WORKSHEET_COLD` | `SHARED_HINT` | reproduces |
| --- | --- | --- | --- |
| `base` | true | false | `c7326f680`, this change's before |
| `hint` | false | false | this change |
| `shared` | false | true | one hint serving resolver and cursor |

**The setup validated itself before it was believed.** The `base` leg reproduces
change 0585's recorded open figure for `ConditionalFormattingSamples.xls` —
**2,099** chain links — exactly, and the open figure is identical on all three
legs for every fixture (14,806 corpus-wide), which is the control: this change
operates entirely inside the scan.

## The probe

`chain_probe.rs`, retained beside this note, was
`crates/litchi-xls/examples/chain_probe.rs` and was removed before commit. It
opens a fixture through `litchi_core::OwnedSource`, runs one scenario per leg,
and prints the link count at the end of the open and at the end of the scenario
together with an FNV-1a digest of the result, so a leg that produced different
output says so rather than the reader having to trust that it did not.

## Removal

The four files the instrumentation touched —
`crates/litchi-cfb/src/{shared.rs,lib.rs}`,
`crates/litchi-xls/src/{lib.rs,workbook/source.rs}` — were verified byte-identical
to their pre-instrumentation sha256 digests after the switches were removed.
