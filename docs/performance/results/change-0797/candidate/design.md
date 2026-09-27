# 0797 checked-attribute candidate

This directory is a source-only candidate archive based on commit
`08a406b3197f00678973424b1ab6cf7b53ea299a`. It contains no production edit.
The five helper owners are archived because the workspace keeps the canonical
implementation in `litchi-opc` and exact source copies in the four crates that
cannot depend on it.

The candidate removes the duplicate-name state from the common one-attribute
success path. `CheckedAttributes::new` creates quick-xml's iterator with
duplicate checks disabled. On its first `next`, a valid first attribute is
examined only far enough to establish the byte immediately after its borrowed
value. If the remainder of the `BytesStart` is XML whitespace, the iterator is
marked `Done`; there cannot be a duplicate after the only attribute. An empty
tag, a first lexical error, and a first `None` also enter `Done` and remain
fused.

If any non-whitespace byte follows the first value, the first item has already
been returned, so the next call enters a cold `ReplayFirst` phase. That phase
constructs a fresh checked quick-xml iterator over the same immutable tag,
consumes its first item to seed quick-xml's internal duplicate state, and then
dispatches the normal `QuickXml(1)` path for the second item. The existing
32-name handoff and ordered-map check are unchanged after that replay. The
replay consumes the first item internally and never yields it twice.

This preserves the important error boundary. The first item has no prior name,
so disabling its duplicate check cannot change its lexical error. Every tag
with a possible second item is parsed by the checked iterator before that item
is returned. A repeated second name therefore fails before its value is
examined, including an unterminated, unquoted, or otherwise malformed value;
the candidate does not scan a long duplicate value before reporting the
duplicate. The source bytes are immutable while both iterators exist, so the
replayed first item has the same key, value, and position. After any error the
candidate sets `Done`, retaining the existing fused and first-error behavior.

No field is added to `CheckedAttributes`; the two new phase variants carry no
payload. This is intended to keep the iterator layout at the baseline size,
which the root-owned direct probe must measure rather than assume. The
candidate uses no unsafe code, dependency, cache, ambient state, or public API
change. The OLE common copy retains its existing `pub` trait and iterator;
the OPC copy remains `pub`, while sign, XLDM, and XML-minifier retain
`pub(crate)` visibility.

The candidate accepts a deliberate tradeoff. Every tag with a second
attribute pays a tail scan after its first value and parses that first value a
second time on the next call. A malformed second attribute, a missing
separator, and a long duplicate value all take that replay path. Tags with
many attributes still pay the existing bounded handoff and ordered-map work;
the candidate does not alter their worst-case algorithm. This is why the root
must run the 0797 direct-helper preflight and reject the candidate if the
multi-attribute cost outweighs the one-attribute allocation saving.

The archived test additions cover:

- empty tags and one-attribute tags followed by all supported whitespace,
  including repeated `None` calls;
- a second valid, duplicate, malformed, or separator-free attribute;
- clones before the first item and after the first-item/replay boundary;
- an error followed by a valid `tail="ok"` attribute, proving that the error
  remains terminal and the valid recovery tail is not yielded;
- the existing canonical quick-xml differential, duplicate-position,
  malformed-tail, random, and bounded-comparison tests.

The root-owned mirror workspace is expected to compile all five archived
helpers with quick-xml 0.41 and the archived canonical test module. Root also
owns the source-size/oracle gate, direct 0797 timing and counter captures,
quality gates, production restoration audit, and any disposition decision.
