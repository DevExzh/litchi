# Initial diagnostic execution

The initial 12 traces completed successfully, but probe Clippy rejected two
single-pattern `match` expressions. Root replaced them with equivalent
`if let` expressions and rebuilt/reran the full bounded diagnostic. These
receipts and source witness preserve the original run. Paths under
`trace-runs/` in these archived receipts are relative to this directory;
original absolute trace paths identify their locations at execution time.
No timing measurements were made in either execution.
