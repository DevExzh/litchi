# Local normative scope

The local authoritative source is `/home/zhuhe/code/litchi/3rdparty/specs/[MS-ODRAWXML]`.

`2 Structures/2.1 Part Enumerations.md` states that nonconforming `brushProperty`
value/unit pairs select a default (lines 202-203). Width and height require an
`xsd:decimal` value and an InkML section 6.4 length unit, with defaults `.053 cm`
and `.001 cm` (lines 204-211). Transparency requires an `xsd:int` from 0 through
255 (lines 216-219). `antiAliased`, `fitToCurve`, and `ignorePressure` require
`xsd:boolean` (lines 244-255). The same section identifies the content as an
InkML subset and references InkML section 6.4 (lines 61 and 206-210). The
implementation follows that referenced length vocabulary: `m`, `cm`, `mm`, `in`,
`pt`, `pc`, `em`, and `ex`; `px` selects the profile default.

The v1 profile work also follows the local EMMA boundary rules in the same file:
`annotationXML` under `traceGroup` must contain EMMA (lines 316-319), EMMA's root
is in the EMMA namespace (339-341), the first child is `interpretation` with a
`context` and `mode="ink"` (343-347), and the optional `group`/sequence/lattice
content is ignored as specified (349-364). The v1 implementation intentionally
scopes projection to those rules and does not claim complete EMMA semantics.
