# The attribution probe

`main.rs` counts logical positional reads, read bytes and source-version
observations taken by the public source-backed `litchi_xls` APIs over one
in-process `ReadAt`, freezes each operation's outcome as a digest so a
before/after pair is also a differential, and carries its own counting global
allocator. It is change 0636's probe with an allocation mode, an SST-extent mode
and a string-cell counter added.

Modes:

```text
cursor-probe sweep  FILE...          one TSV line per fixture and operation
cursor-probe detail FILE...          the same, plus validate size histograms
cursor-probe sst    FILE...          each fixture's SST record-group extent
cursor-probe alloc  OP FILE          allocations and allocated bytes for one operation
cursor-probe time   KIND OP N FILE   N timed lifecycles of one operation
```

`KIND` is `owned` (an in-process `ReadAt` over the file's bytes) or `file`
(`litchi_core::FileSource`). `OP` is `validate`, `open`, `list`, `all-cells` or
`full-text`.

Build it against each leg's checkout by writing `Cargo.toml` from
`Cargo.toml.template` with `__TREE__` replaced by that checkout's path, then

```sh
CARGO_TARGET_DIR=<scratch> cargo build --release
```

and stage the binary outside the Cargo target directory before timing anything
(change 0627: a concurrent build relinked one mid-run).
