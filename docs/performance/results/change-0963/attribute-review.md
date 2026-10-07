# 0963 — bounded XML attribute candidate review

Status: source and retained-history review complete. No candidate was coded,
built, or benchmarked. Production remains unchanged. The proposed first-two
attribute prefix is the already rejected 0799 family, so it does not provide a
new 0963 implementation candidate.

## What quick-xml allocates

The checked iterator in the local `quick-xml 0.41.0` source stores duplicate
keys in `IterState.keys: Vec<Range<usize>>`; the initial `Vec` and optional hash
set are empty. `Attributes::with_checks(false)` only clears the duplicate-check
flag. It does not install inline storage or otherwise allocate up front.

On the first successfully parsed checked attribute,
`check_for_duplicates` pushes its key range into that `Vec`, which causes the
first allocator growth. A one- or two-attribute tag therefore retains the
allocation until the iterator is dropped. Once 32 keys are present, quick-xml
seeds its hash set and hashes the earlier keys. `Attributes` and its iterator
state derive `Clone`, so cloning after keys have been observed clones that
storage; it is not a shared cursor. The iterator is fused, and an adapter must
also remain fused after the first error.

This confirms the allocation mechanism behind the 0962/0899 attribution. It
does not show that removing the allocation improves a workflow: the duplicate
checker was a small native leaf in the separate 0899 producer population;
that does not bound its current cost. Allocation stack presence is not a
latency saving.

## Candidate identity and prior disposition

The suggested algorithm is exactly the 0799 replay design:

1. begin with an unchecked `Attributes` iterator;
2. parse the first two attributes without the duplicate-key `Vec`;
3. compare the second key explicitly, before parsing its value, so a duplicate
   with a malformed or long value still reports the duplicate first;
4. on a third attribute, or on a non-whitespace tail after the second, create a
   fresh checked iterator, consume the first two attributes again, and continue
   with the existing bounded handoff; and
5. retain the current BTree-based path from the 33rd attribute onward.

The retained 0799 packet measured short-input wins but regressions on distinct
three/four-attribute tags, valid duplicates, and malformed values after the
prefix. It also pays replay and fresh checked-state construction on every
longer tag. 0801 removed replay but retained middle-tag and empty-path
regressions. 0802 added a boxed 32-entry linear array, which increased
iterator state and was rejected together with its protected syntax failures.
0805 combined those techniques and still had material four-attribute and
error-path regressions. The full five-helper workflow in 0806 had no eligible
latency benefit; its large lifecycle candidate/control ratio was 1.031050
(+3.105%, with the confidence interval wholly above one). Its allocation
lane improved capture allocation counts but did not improve latency, so the
workflow gate rejected adoption. 0899 consequently records this family as
low priority without new evidence.

Therefore a fresh 0963 trial of an unchecked first-two prefix would repeat an
existing rejected mechanism. A 32-entry stack or inline array would also need to resolve the state-size
and middle-tag costs seen in the related boxed-prefix 0802 design.

## Compatibility obligations if the family is revisited

Any future implementation would need a byte-for-byte differential oracle for
the checked iterator, including leading `=` names, whitespace around the
equals delimiter, missing or malformed equals, invalid keys and values,
duplicate keys whose value is malformed, and the first error's position. The
duplicate check must precede value scanning exactly where quick-xml does.

The oracle would also have to cover `Clone` at every prefix and fallback
boundary, `FusedIterator` after `None` and after `Err`, borrowed ranges across
the replay, the 33rd-attribute BTree handoff, and hostile long attributes.
The five mirrored litchi consumers must remain equivalent. These obligations
are why a short-input allocation result cannot authorize a local shortcut.

A materially different dependency-level idea would be to retain the checked
quick-xml parser while changing its key store to a two-entry inline container
(for example, a small-vector representation), so no prefix replay occurs.
That would require a quick-xml patch or vendored source, lockfile/source
custody, dependency-wide differential tests, and a new full workflow. It is
deferred: no current evidence justifies that dependency and ownership cost.

## Worksheet scan and namespace ownership

`litchi-xlsx::cell_values::validation::scan_attributes` is narrower than a
global `CheckedAttributes` change. It is shared by the authoritative
`validate_xml` path and the original-source `worksheet_xml_and_parse_source`
observer. It currently consumes `element.checked_attributes()` on every
start/empty element to preserve lexical and duplicate refusals; in the
composed worksheet span it also rejects qualified attributes with local name `id` and prefix other than `xml`. Outside that span it deliberately consumes the result without
attribute-specific policy.

The two 0962 namespace leaves cannot be reused directly here. In quick-xml
0.41, `NsReader::process_event` calls `NamespaceResolver::push` before the
validator sees the event. `push` performs a separate unchecked attribute pass,
retaining only namespace declarations and installing them before
`resolve_event`. `resolve_event` then resolves the element prefix; it does not
return a checked-attribute result. The existing pass can stop at an unchecked
attribute parse error and can return namespace-declaration errors, while
`scan_attributes` must preserve quick-xml duplicate/value error behavior.
Combining them changes install and error ordering unless extra state is added.
The 0812 namespace review rejects that simple fusion, and the historical 0203
namespace-cache experiments regressed or failed their guardrails.

A genuinely distinct, but deferred, diagnostic hypothesis is a validator-local
event sidecar: during the existing namespace pass, retain an exact
quick-xml-compatible attribute-validation result and the worksheet `r:id`
finding, then let `scan_attributes` consume that result instead of walking the
attributes again. This would target only the two worksheet validation routes
and would leave the shared iterator unchanged. It is not implementable by
reusing today's `NamespaceResolver` result: `push` currently discards the
per-attribute observations, `BytesStart` has no such sidecar, and the merged
pass must preserve namespace-install-before-validation behavior. It therefore
requires an explicit API/source change and a differential design before any
code is written. It remains deferred rather than a candidate.

If that diagnostic ever earns a trial, its oracle must compare both validation
routes against the current implementation over duplicate/malformed values,
namespace declarations and reserved prefixes, relationship attributes,
unknown copied elements, MCE events, and namespace-limit failures. It must
also measure the ordinary release workflow on the medium, dense-sparse, and
one-edit cases; source-owner or allocation percentages alone cannot qualify
it. The existing current-source, exact-error, validation-first, and bounded
resource constraints from ADR 0005/0006/0008/0024 remain binding.

## Decision

Do not rerun 0799–0806 under a new name. Keep the bounded 33+ BTree mechanism
and current checked semantics. Defer both the quick-xml inline-key-store idea
and the validator-local namespace sidecar until a new source-level design and
matched evidence isolate a removable second pass without replay or ordering
risk. No production or benchmark artifact was generated by this review.
