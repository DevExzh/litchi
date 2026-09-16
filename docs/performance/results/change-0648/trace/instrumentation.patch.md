# The temporary instrumentation that produced these traces

The resolve traces in this directory were captured with one `eprintln!` added to
`resolve_shared_string_inner` in `crates/litchi-xls/src/workbook/source.rs` and
removed again before anything was committed. It is reproduced here so the traces
can be regenerated:

```rust
// immediately after the entry's source span is computed
if std::env::var_os("XLS_SST_TRACE").is_some() {
    eprintln!(
        "SSTTRACE\t{}\t{}\t{}\t{}",
        strings.region_start, strings.region_end, first_offset, span
    );
}
```

`region_start` and `region_end` were the bounds of the SST record group, and
`first_offset`/`span` the source offset and length of the entry the resolve
needed. With that line in place:

```sh
XLS_SST_TRACE=1 cursor-probe time owned all-cells 1 <fixture> 2> <fixture>.all-cells.tsv
```

Each `.tsv.gz` here is one such capture: one line per resolve, in resolve order.
`simulate.py` reads them and prices the policies in `policy-simulation.txt` —
six sliding-window ceilings with and without change 0568's growth schedule, a
window sized to the entry's span, and the whole-table policy the change
implements. That table is the evidence for rejecting the sliding window.

The shipped code does not carry the instrumentation and does not carry the
`region_start`/`region_end` field pair either: the final design holds the region
as one `Option<(u64, usize)>`, because a table that does not fit one window
disables the window rather than shrinking it.
