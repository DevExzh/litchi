# 0537: attribute raw XLSX decoding costs and isolate transient ownership

This diagnostic replays sealed 0530 planning evidence and verifies that 13
relevant parser, snapshot, harness and lockfile inputs still match the measured
source. It confirms a distinct raw-parser opportunity: `decode_cell_attribute`
forces decoded values into owned strings even when references and numeric
metadata are consumed immediately. The retained cell type still needs owned
storage across events. Production and harness source are unchanged.

Across medium repeats, decoder inclusive Ir is 3,871,719/3,871,907, with direct
allocation-child Ir 1,584,480/1,584,668. Dense-sparse values are
7,601,777/7,603,683 and 3,031,751/3,033,657 respectively. Those are historical
instruction costs within planning, not fresh latency or allocation counts.
Nested scan, decoder and allocator totals must not be added together.

The [evidence and draft](../results/change-0537/README.md) preserve checked
attribute scanning, normalization, duplicate detection, validation order and
retained ownership. This differs from the rejected 0522 layout-scanner shortcut
and 0531 MCE search candidate; neither is revived.

The source audit also identifies a prerequisite: current allocator observation
starts at commit, after the planning timer. The next measured batch must first
add aligned planning allocation samples, then freeze the tested source and run
fresh matched native, profile, allocation and eager-read guards. Existing
commit/publication vectors cannot prove a planning allocation reduction.

No new Rust build, test, capture or accepted performance result is claimed.
Historical raw/annotation replay and current source bindings are verified; the
candidate remains an unapplied draft. OLE2/OOXML remain active, ODF stays
deferred until their goal completes, and iWork is excluded.
