# 0007 — Coordinator defaults for the 0003 open questions

Record 0003 left one finding and five questions for the owner. So that M0a can
start, the coordinator adopts the defaults below. Each is reversible before the
M0a merge, and each is listed under "Awaiting owner confirmation" in
`PROGRESS.md`.

| Item | Default adopted | Why |
|---|---|---|
| F2 | `File` is flat, one variant per format, plus `Locked`/`Signed`. Families (document, presentation, workbook) are accessors. | Seven claimed ODF formats fit no family. Design §4 leaves the variant names to this design. |
| Q1 | Add the Litchi extension codes `limit_exceeded` and `cancelled` next to `unsupported_capability`. | Lets agents tell a budget refusal from bad input and retry correctly. |
| Q2 | `litchi::open` on iWork returns `unsupported_format` with a suggestion to use `litchi::iwork`. | Tier 3, gate R1. |
| Q3 | `retain_source` defaults to `true`. | Writes are required for tiers 1–2. S10 must measure the RSS cost (R5). |
| Q4 | `Signed` wraps every signed input. | ADR 0006's simple, safe reading. |
| Q5 | Steps S0–S10 land on `uo/m0a`, and one `--no-ff` merge brings them into `feat/unified-ops-api`. | Trunk sees one switch; the history keeps the individual steps. |
