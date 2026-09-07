# Next measured candidate: avoid repeated cached matching

0461 removes helper/result overhead but retains the O(n*k) cached-prefix scan
for n parsed attributes and k requested fields. Initial and candidate slide
parsing remain measured hotspots. This is a next experiment, not a performance
claim or an implemented index.

The bounded proposal is a private recognized-key enum plus a fixed array of
first-occurrence indices beside the existing raw-attribute vector. Start with
`shape_builder` queries in `codec/parser/codec/xml/semantic.rs`: presentation
class/style-name/placeholder/user-transformed; draw name/style-name/layer/z-index/
transform; and SVG geometry fields. Unrecognized keys retain the current path.
Do not build an eager attribute map or cache decoded values.

On a recognized cached hit, decode that first indexed attribute only. If no
index exists, continue the raw iterator from its current position; a partial
scan never proves absence. Clean exhaustion can be remembered. A cached match
must precede replay of an already-recorded malformed-attribute error, while an
uncached request that reaches that error still fails at the same point. Exact
namespace URI and local-name classification must preserve aliases, rebinding,
unknown prefixes and the fact that default namespaces do not bind attributes.
The existing element namespace classifier is insufficient for all attribute
namespaces and must not be reused without explicit coverage.

Record an index only after the existing successful cache append. In particular,
a freshly reached matched value that fails decoding advances without being
cached or indexed; a previously cached invalid value remains replayable. First
indices never change. Drawing-attribute harvesting keeps source order and its
own decode/error wording; only successfully appended attributes enter indices.
The independent direct raw-attribute oracle, not `Parser::get_attr`, must cover
these sequences, duplicates, malformed input, namespaces and harvest parity.

Two costs need measurement before retention: classifying requested keys and
recorded attributes can offset the avoided scans, and the fixed index increases
per-element stack state even if heap metrics stay flat. Consider a private typed
known-key call for hot sites so constants need not be classified repeatedly;
retain a fully equivalent fallback for uncommon keys. Do not add an unbounded
map or use prefixes as semantic identities. Keep the first implementation
small, safe and independently testable rather than projecting an exact gain.

Inspect generated code for removal of the repeated cached scan, then capture a
new baseline/candidate matrix with a frozen practical gate, tiny-input review,
allocation/peak/RSS evidence and separate mechanism diagnostics. Preserve all
independent candidate validation, no-op and reversible patch checks. 0460's
accepted staging fusion remains the baseline; 0461's rejected factoring does
not need to be retained as an enabler.
