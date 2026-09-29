# Independent applied-source review

The reviewer found no source blocker in the applied, formatted diff.
`file_entries()` maps the immutable physical-order ID array to corresponding
normalized names. Constructor admission publishes matching index/entry/order
records, and no mutable access can break that relation. The removed missing-ID
branch cannot occur for this immutable valid index. Direct `metadata_for` keeps
the same central metadata and bounds checking.

Both selected OPC loops preserve progress/source checks, checked arithmetic,
relationship classification, reservations, folding, physical order and first
representative behavior. The new iterator allocates no retained state, introduces
no cache/memo, performs no source reads and changes no preservation policy.
Physical IDs stay inside existing ZIP/OPC ownership seams.

The focused ZIP test correctly reverses central-directory order while keeping
physical local order. It exercises directory exclusion, normalization of
`dir/../second.bin`, exact iterator length/exhaustion and ID/name/metadata parity.
The applied patch is separately retained after rustfmt; the pre-format candidate
remains immutable as generation evidence.

Historical attribution was independently clarified: the two selected
`validate_topology_limits` caller-edge costs sum to 8,792,282 and 8,791,686 guest
instructions in the two 4,096-Part captures (11.569% and 11.537% of the owner).
The initially quoted 4.40M / 5.78% was only one edge, not the aggregate. Inclusive
costs overlap other regions and cannot be converted to native phase time.
Fresh native and instruction comparisons remain prerequisites for performance
claims; unchanged allocation counts would be expected for this lookup-only
mechanism, but must still be measured.
