# Next OLE2/OOXML priority

0544 makes progress by proving the event-bound oracle and measuring preflight
cost on valid inputs. Its two-scan helper avoids the large 0543 parser-discard
cliff but still pays enough scan work to fail two cap p50/mean rows; a primary
workflow repeat also fails its gain threshold. Do not re-run unchanged source
until a new mechanism or measured source of variance justifies another campaign.

A reviewed alternative is a coarser single-scan upper bound:
`1 + initial_text + 2 * count(< or &)`. Every markup/reference event begins at
one of those markers, and at most one following text event can precede the next
marker. This can decline earlier with one scan, sacrificing shared traversal
for some larger below-cap sources. It must retain authoritative validation,
source fences and runtime caps. See next-design-review.md for static disposition;
there is no implementation or performance admission for this proposal here.
Static counts expose a material limitation: the current 128 grid has a coarse
bound of 131,593, above the 131,072 cap, so this alternative loses the dense
primary shared path. It is not ready to freeze as the next candidate. Earlier
approximate marker counts were wrong; the independent review supplies the
correct counts. Next investigate a cheaper bound that preserves both primary
shapes, or explicitly measure a different admission/resource tradeoff.

A fresh campaign must preserve current success/error/boundary tests and all
native, allocation, refusal and cap gates. Directly check that the real 96/128
fixtures remain admitted and that 160/164/256 fallbacks have acceptable scan
cost; do not equate conservative correctness with useful performance.

If a shared traversal eventually passes, fresh commit-path attribution should
follow for overall workflow impact. OLE2/OOXML stays ahead of deferred ODF, and
iWork remains outside this workstream. The broad performance goal is unfinished.
