# 0537: raw worksheet attribute ownership attribution

The previous batch made progress by measuring and rejecting the CFB cold-error
layout candidate, restoring production, committing its evidence and cleaning
owned storage. This batch follows the OOXML planning priority using sealed
historical profiles. It performs no new build, Rust test or performance capture.

The analyzer replays both prior seals, the exact 0530 planning annotations and
positive selected-parent edges, and checks 13 relevant current source/lockfile
bindings against that measured source. It then follows the raw parser’s cell
attribute decoder one level deeper than the earlier report.

| Historical shape/repeat | Planning Ir | Attribute scan inclusive Ir | Decoder inclusive Ir | Decoder direct allocation-child Ir |
| --- | ---: | ---: | ---: | ---: |
| medium/1 | 125,643,633 | 10,273,383 | 3,871,719 | 1,584,480 |
| dense-sparse/1 | 237,075,784 | 20,032,877 | 7,601,777 | 3,031,751 |
| medium/2 | 125,655,087 | 10,273,571 | 3,871,907 | 1,584,668 |
| dense-sparse/2 | 237,102,562 | 20,034,783 | 7,603,683 | 3,033,657 |

The decoder is nested inside the scan, which is nested inside planning. These
costs must not be added together. Allocation-child Ir is instruction work, not
an operation-local allocation count; collection-off call metadata is excluded
from that inference. The numbers are historical, not fresh timing results.

The distinct draft candidate preserves the decoded `Cow` for transient cell
references and numeric style/metadata fields. The cell type is retained in
`PendingCell`, so it must still become owned at that boundary. XML normalization,
complete checked attribute scanning, duplicate handling and error order remain
required. This is the raw semantic parser, not the rejected 0522 lossless-layout
scanner shortcut, and it does not revive the 0531 MCE search replacement.

Before candidate measurement, add a planning allocation region around the
existing edit_sheets timer, with aligned raw samples and normal-build unavailable
status. The current commit/publication allocation vectors do not cover planning.
Then run independent decoder/attribute-order/preservation guards and fresh matched
native, planning-profile, planning-allocation, commit/publication and eager-read
lanes. Freeze thresholds before captures and measure the final tested source.

No runtime optimization is retained or claimed here. OLE2/OOXML stays first;
ODF is deferred until that optimization goal completes, and iWork is excluded.
