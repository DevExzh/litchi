# 0362ul87ao — use the local OOXML test result alias

The approved F5 run for `029446d5d9daecdc8c239017dd2f00ca3d272851`
(tree `eeeb0d9c6a587492d40ddda262d3d4f5e608abeb`) passed `fmt-scoped` and
stopped at row 1, `check-common`. The transaction-error consumer migration
was otherwise accepted; the workspace's `-D unused-qualifications` lint
rejected `std::result::Result` in the two test publish helpers.

Both helpers now use the local `Result<bool, crate::properties::Error>` alias.
This is a spelling-only repair: the canonical transaction error type and all
typed error assertions remain unchanged.

The preserved receipts are `rows.json` SHA-256
`927d8cfda4c52e17bb1ca6e7d48312feaae723c31db025035ee3b7c40863d327`,
`result.json` SHA-256
`841520bf2db05a6f58f6479629dd6f1f3263e6aa8adbf5f78703214c80a0754f`, and
the `check-common` compiler log SHA-256
`5fec9506f5f7230f9b18fc749ac583cf79a6d0037b793861c6754b095da85564`.
No retry was run for this repair; the next attempt must bind a fresh source,
target, plan, driver, manifest, and fixture inventory.
