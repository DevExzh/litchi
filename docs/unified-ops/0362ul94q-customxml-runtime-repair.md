# 0362ul94q — Custom XML runtime repair

This follow-up repairs the four Custom XML failures reported by the 0362ul94
runtime qualification at source `3774fe4374dfb66e74b98b1ae742672ed6cad37f`.
The failure receipt is retained outside the checkout at
`.coord/ooxml-combined-0362ul94-f5-runtime-3774fe43-v2/common-tests/0000.log`.

The transactional insert assertion now follows the current typed contract:
an occupied destination is `RefusalKind::UnsupportedTarget`, while the
package and relationship-count atomicity checks remain unchanged.  The MCE
rewrite fixture keeps a tight source profile but admits the bounded growth of
the selected branch's newly inserted schema reference.  The lexical exact
output-limit fixture now changes the URI to a genuinely longer value, so its
exact and one-byte-under checks exercise the serialized-output fence rather
than rejecting an input source that is larger than the shorter output.

The noncanonical properties fixture's foreign `x:itemID` expectation now
matches the source's braced value (`{foreign}`).  The test continues to check
that the expanded-name `p:itemID` and `p:uri` values are the only known
attributes rewritten, while prefixes, quote style, whitespace, comments, and
foreign attributes remain intact.

No Cargo, rustdoc, native, or performance command was run for this source-only
repair.  The next runtime qualification must verify the four rows before any
acceptance claim.
