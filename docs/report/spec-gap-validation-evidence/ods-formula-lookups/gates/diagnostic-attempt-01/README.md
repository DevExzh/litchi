# Superseded lookup gate attempt 1

All seven gates passed on this snapshot. It is not the accepted final source.
Subsequent review found scratch reservation drop-order errors and a missing
shape-probe depth decrement. The corrected candidate adds a regression with
40 sequential CHOOSE expressions. The exact 61 selected inputs and all gate
receipts are retained here; no performance capture used this snapshot.

Selected inputs are retained in `selected-source.tar.gz` (SHA-256 `8770ff27f7d218fcede4f496082e88b148f53fe2ca191de60bb80843bcfc6ade`).
Its 61 members match `freeze.json`; extract into a checkout at the recorded
baseline to reconstruct the selected inputs.
