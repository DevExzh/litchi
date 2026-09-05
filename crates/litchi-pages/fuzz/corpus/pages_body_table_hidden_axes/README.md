# Pages body-table hidden-axis corpus

These entries are bounded fixture descriptors consumed by
`pages_body_table_hidden_axes`.  The fourth byte selects the malformed graph
or wire mode; the first three bytes keep the rooted table and requested axis
shape reachable.  Every case is deliberately isolated so a failure identifies
one invariant instead of depending on a mutation sequence.

| Descriptor mode | Case | Invariant exercised |
| ---: | --- | --- |
| 1--7 | missing-root/body/attachment/drawable/model/UID-map/formula-owner | rooted ownership and dependency presence |
| 8--10 | duplicate model/info/formula-owner | singular native owner census |
| 11--12, 32 | duplicate active state; wrong/invalid owner UUID; wrong active UUID | active-state uniqueness and UUID binding |
| 13--15 | UID-map length/permutation; unknown/invalid axis UUID | stable UID-map cardinality and identity |
| 16--20 | wrong direction; wrong message type; dangling reference; duplicate name; shared owner | row/column topology and cross-component ownership |
| 21--26 | duplicate field; wrong wire; truncation; malformed known field-70 group; noncanonical wire; invalid UTF-8 | strict wire handling and source preservation |
| 27--28 | missing filter owner; cross-component alias | dependency presence and type-safe ownership |
| 29--30 | hidden-axis bounds; row/column count overflow | model cardinality and arithmetic bounds |
| 31--34 | all-zero UUID; wrong active UUID; UID-map type 6005; dangling formula dependency | UUID validity, canonical routing, and formula ownership |
| 35--38 | missing/wrong model field-46 metadata; missing table-info field-6 metadata; filter dependency metadata | exact field-path ownership and unsupported filter dependencies |

When physical package ingress admits a malformed descriptor, the harness also
checks the redacted public error class: duplicate names are
`AmbiguousTableName`, cross-component aliases and invalid filter metadata are
`UnsupportedDependency`, and the other malformed graph/wire/UID/metadata
cases are `InvalidSource`.  If physical ingress itself rejects a malformed
wire or archive graph, that bounded archive error is retained as the expected
failure path.

The valid seeds cover the absent-owner nonempty-set refusal
(`ownerless_nonempty_set.hex`, plus `absent_owner_nonempty_row.hex`,
`absent_owner_nonempty_column.hex`, and `absent_owner_nonempty_both_axes.hex`),
the absent-owner no-op (`absent_owner_noop.hex`), existing row/column/both-axis
state, filtered/pivot/locked tables, and each command (clear, reset, and
no-op).  `selector_index_one_existing.hex` keeps a changed existing-owner edit
on the second table.  The maximum-index seed keeps the public selector bounds
path hot; modes 29--30 mutate the model's persisted bounds and counts
independently.  `axis_bounds_out_of_range.hex` additionally requests row three
after shrinking the model to three rows, exercising the inconsistent graph-
bound rejection.
`non_identity_uid_map.hex` enables a valid physical-to-stable UID permutation
for the first table; `non_identity_uid_maps.hex` enables the same permutation
on both tables while editing the second.  Both seeds check both permutation
directions, keeping positional semantic reads and edits dependent on the
canonical 6267 map.

The eighth descriptor byte selects the bounded package-byte profile: low nibble
`0` means exactly the source length, `1` means source length plus one, and
larger values request a correspondingly smaller ceiling.

Every valid fixture also carries unknown scalar, fixed32, fixed64,
length-delimited, and group records in the model, table-info, and hidden-owner
payloads.  The harness checks that those records, plus unknown fields nested in
UUID/reference payloads, survive an existing-owner rewrite byte-for-byte.
