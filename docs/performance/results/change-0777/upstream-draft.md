# Draft for quick-xml maintainers — attribute duplicate-check boundaries

This is an unsent technical draft. It describes the locally resolved
quick-xml 0.41.0 implementation; it is not a vulnerability disclosure or a
claim that a practical large collision family has been constructed.

The checked attribute iterator uses a linear name scan for the first 32
attributes. Above that boundary it stores deterministic 64-bit name hashes as
a pre-filter and scans earlier names on a hash hit to recover duplicate error
positions. The fresh `DefaultHasher` for each name is unkeyed. Distinct names
with equal hashes therefore trigger rescans; the algorithm does not provide a
deterministic subquadratic comparison bound. A normal distinct-name timing
sweep is not a demonstration of a hash-collision attack.

Separately, callers that continue after duplicate errors can repeatedly incur
those rescans with ordinary duplicate names. Recovery after a duplicate can
resume within its quoted value: for example, a repeated `a` with value
`x b='y'` can expose a spurious `b` attribute on subsequent iteration. Callers
that stop at the first error are unaffected by that recovery behavior. Litchi
already uses a separate first-wins policy for its lenient readers.

For fail-fast readers Litchi now keeps quick-xml's first 32 checked results,
disables its duplicate check before the next item, and tracks borrowed names
and source positions in an ordered map. It checks duplicate precedence when an
unchecked lexical value error is returned. This preserves items and errors
through the first error and gives O(n log n) name comparisons; name-byte cost
and allocation policy remain separate concerns. The adapter deliberately does
not promise quick-xml's post-error recovery API.

Possible upstream directions include a keyed collision-resistant pre-filter,
a deterministic ordered fallback, or a dedicated fail-fast iterator with an
explicit recovery contract. Any change should retain exact duplicate position
reporting and distinguish duplicate-before-value errors from malformed names
or missing equals signs. Benchmarks should include ordinary short tags,
32/33-name boundaries, late duplicates, malformed duplicate values, long
shared prefixes, and callers that deliberately continue after errors.

The accompanying 0777 packet contains source-bound unit gates, item/error
comparison probes, reader differentials and separate performance observations.
Their scope and any cost regressions belong in a future submitted report;
no partial capture or historical stale result should be substituted for the
completed packet.
