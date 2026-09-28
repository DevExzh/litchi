# 0802 bounded-prefix candidate

Source-only handoff based on `334a22f10185d6c9be291fc16e983baa73e25035`. Production is untouched. The
archive was untested at handoff; later packet receipts are authoritative for
execution and disposition.

The candidate retains 0801's no-replay first two attributes and raw third-key
preflight. Exactly empty raw tails become Done during construction. Whitespace
and malformed tails still follow the parser. No XML Name validation is added.

A successful third item with remaining non-whitespace bytes seeds one boxed
array of 32 borrowed Name slices. Subsequent raw keys are checked before values
against the occupied prefix, once on successful parsing. Positions are derived
from the key's original slice only when needed. Comparisons go through Name's
test counter. At most 32 stored names are scanned per new item.

A unique 33rd item either ends the iterator when only whitespace follows, or
seeds the ordered map from all32 stored slices and that item's key. Every later
item uses bounded ordered-map checking. No earlier value is replayed and no
name is hashed. Clones retain the exact linear or ordered phase.

The OwnCheck enum is stored directly in Phase; each backend has one box,
avoiding a redundant outer allocation. Array initialization and copying, the
32-name scan, and the later map seed remain costs to measure. Iterator size,
allocation behavior and public-workflow effects are not inferred from source.

The 39-case matrix and 18 protected consume cases retain the prior frozen
policy, with a fresh seed802080. This packet can only qualify or reject the
candidate for later workflow/resource/cross-format trials, never adopt it.
