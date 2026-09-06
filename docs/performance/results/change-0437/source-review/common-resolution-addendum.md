# 0437 common seam review resolution

Follow-up to `actual-draft-review.md`, recorded after the coordinator's
source-only resolution and focused review.

The apparent phase concern is intentionally resolved in favor of the broader
ancestor grammar: fixed `Empty`/balanced child subtrees may occur under any
ancestor that remains open in the prefix. The final insertion path is defined
as the trailing run of `Start` events after the last fixed child closes. Thus
the previously discussed shape

```text
<root><slot><fixed/><inner> ... </inner></slot></root>
```

is valid and inserts at depth three. This matches the handoff contract of
“root/open ancestors + balanced fixed children* + final Start+”; no parser
change is required. The focused coverage now includes positive `Empty`/`End`
under-ancestor behavior and rejects prefixes ending in an `Empty` or `End`,
which have no trailing open insertion path.

The fixed-character-data restriction is deliberate. The opt-in prelude
constructor rejects fixed text, CDATA, and references; its documentation now
states that shell text-byte accounting is zero and character data comes from
fragments. This is a narrower element-only common contract chosen for the ODP
shell, rather than an omitted counter implementation.

Typed refusal coverage was also tightened: malformed prelude cases assert
`Error::InvalidFormat`, while valid shell limit failures continue to assert
the typed XML limit resource, actual value, maximum, and callback ordering.
Shell-depth-dominance and the positive balanced-child/end transition are
covered as well. The strict `try_new` constructor remains unchanged.
