# 0362ul39a — PPT consumers without encryption

The standalone performance consumer check in the RTF policy lane exposes two
unmigrated PPT calls behind `cfg(not(feature = "encryption"))`. Both eager
and shared presentation reads still pass the one-argument `CurrentUser::parse`
function to `Option::map`, although the raw parser now requires selected
`SourceLimits`. All-feature checks compile the encryption branch and therefore
do not cover these consumers.

Both calls now pass the same retained source limits already used to preflight
and load the owning presentation. No public limit-less twin is restored.
The required non-encryption compiler and standalone consumer rerun remains
pending; the failed RTF receipt is retained at
`(deleted-coord-artifact)`. Its lock provenance is being
reviewed separately, so this record makes no RTF acceptance claim.

This repair changes the two feature-gated consumers together. The shared
parser, encryption-enabled paths, timed operations and selected profile
semantics remain unchanged.
