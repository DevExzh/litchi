# Reference-metadata integration gates

The candidate is staged into a checkout at the commit in `../baseline.json`.
`stage.py` copies the explicit performance source closure, with the retained
Cargo.lock. The root workspace lock is a distinct input and is never replaced.

`run.py CHECKOUT TARGET` runs package tests, warning-denied all-target Clippy,
warning-denied rustdoc, package formatting, selected-file formatting, crate
boundaries and diff checks. Source manifests cover all workspace Rust/build
inputs plus literal ODS include dependencies before and after the run.
`verify.py` checks the retained logs, exact commands and source closure against
the baseline Git objects plus the selected candidate files. Missing receipts
are failures unless preparation explicitly uses `--allow-pending`.

No gate run or source freeze is claimed by the preparation scripts alone.
