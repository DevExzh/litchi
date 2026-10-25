# 0362ul111a — restore the PPTX collaboration replay aggregate

This return supersedes the semantic-budget change recorded by
`0362ul111-pptx-comment-mce-replay.md` and is based on
`e130751fc503099c4d5fff74a5746b5099d625e1`.  The earlier `cc2a9701c6`
implementation widened the processed parser's `max_string_bytes` to the
physical input limit and checked only model fields after ownership.  That
violated the established selected aggregate contract, so this return restores
the existing bounded parser call with the recipient's selected limits.

The selected aggregate remains the decoded byte total of every selected XML
attribute name/value and text value, including root namespace declarations and
active extension attributes.  It is charged before an event is owned.  The
independent raw scan still admits physical input, event/depth/node, and
per-attribute limits before MCE; discarded branch bytes therefore cannot evade
physical admission, while they do not consume the selected semantic aggregate.
The production path keeps one MCE projection and one bounded ordinary grammar
pass; no parser or limit is bypassed.

The two inactive-MCE regressions no longer use the stale 128-byte budget.  The
presence fixture independently calculates 251 selected bytes:
`(7+58)+(9+56)+(8+59)` for the three root namespace declarations, 32 for the
five active `cmAuthor` attributes, and `(3+19)` for the selected `p:ext` URI.
The threading fixture independently calculates 288 selected bytes: 197 root
namespace bytes, 13 comment attributes, 4 position attributes, 5 text bytes,
41 threading URI bytes, 15 `timeZoneBias` bytes, and 13 `parentCm` bytes.  Each
fixture keeps inactive values above its cap and asserts both the exact budget
and one byte under it.  These expected values are written from the fixture
contract rather than obtained from the production preflight counter.

No Cargo, build, runtime, rustdoc, native fixture, or benchmark command was
run on this source-only leaf.  Rustfmt, `git diff --check`, and the repository
evidence-size/binary guards are run before the guarded commit.  Runtime and
full-feature qualification remain coordinator validation items.
