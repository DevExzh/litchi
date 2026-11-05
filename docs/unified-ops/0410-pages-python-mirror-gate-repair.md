# 0410 — Pages Python mirror gate fixture parity

- Status: bounded gate-repair slice on `uo2/iwa-pages-python-gate-repair`,
  based on coordinator source `3ce2a067742b9606e3c2558adee9e022c87f0420`.
- Scope: `crates/pyo3-litchi/tests/test_crud_mirror.py` fixture setup only.
  No binding runtime, Rust API, assertion, ledger, native queue or capability
  changes.

## Root cause and repair

The Python mirror replays the Rust CRUD checklist's `seed` fixtures. The Rust
source `crates/litchi-ops/tests/crud_matrix.rs` defines a Pages seed with the
same three steps used by the Pages rows:

```text
add /body paragraph {text: Intro}
add / section {name: Appendix}
set /section[2] {text: Tail}
```

The Python `SEEDS` dictionary omitted the `pages` entry, so every Pages seed
row failed before opening a fixture with `KeyError: 'pages'`. The repair adds
that exact seed entry; tracked Pages fixtures remain the source for all pinned
fixture rows.

## Verification

- Full row-38 Python suite: 301 tests passed.
- Focused `crud_mirror` suite: 3 tests passed.
- Python AST syntax check passed.
- Build: `cargo build --locked --offline -p litchi-py`, four jobs,
  incremental disabled; exit 0. Build log is external at
  `(deleted-coord-artifact)/resume-20261030-0410-pages-python-build.log`
  (SHA-256 `e6c0b36c3c04939d81978190e76af0667f23d95114800d6d26570fd7d6218871`).
- Full test log is external at
  `(deleted-coord-artifact)/resume-20261030-0410-pages-python-gate-repair.out`
  (SHA-256 `fc4b5de572129372575fff404920e6f4d3bddd9941b34ec4393a0da04421271b`).
- Lockfile SHA-256: `f6b3eae7157c819fd5c02dfbfe9decaf2b973778fb9659b70b738b591e7f3281`.
- No `3rdparty/` references or generated evidence/binaries were added.

This record proves Python mirror fixture parity only; native application
acceptance and unrelated migration rows remain unchanged.
