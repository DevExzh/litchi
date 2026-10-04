# 0904 XML audit attribute-summary semantic review

This is a source-only qualification of the zero/one-attribute summary idea
from change 0903. It was made against `0283e96b73`, the restored production
tree (whose production source matches the 0903 audit base). No production
source, tests, builds, profiles, or worktrees were changed or run.

## Conclusion

A bounded summary is semantically viable as a candidate only after
`check_start` has completed its whole tag scan and returned success. The only
initially defensible fast cases are:

* zero attributes; or
* exactly one attribute, identified from the raw `AttributeSpan::name`, whose
  bytes are not exactly `xml:space`.

The summary must distinguish `0`, `1 non-xml:space`, and `other`. It must not
be inferred by consuming only the first item of a quick-xml iterator. The
first pass has to finish `check_attribute_layout` and, for the source policy,
the namespace checks in `Namespaces::check_tag`; a second attribute, even a
malformed one, must prevent the fast path.

For zero attributes the replacement for `inspect_attributes` returns the
inherited `Space` unchanged and performs no charge. For one non-`xml:space`
attribute it performs exactly one aggregate `Resource::Attributes` charge
*after* `check_start` succeeds, at the existing tag-start `offset`. The
per-element `Resource::ElementAttributes` check remains in the complete first
pass. A valid summary must therefore leave `State::attributes`, the returned
space, namespace state, observer counters, and the next `enter_element` or
`enter_empty` call identical to the current path.

This is a qualification boundary, not an adoption recommendation. The
existing iterator must remain for two or more attributes and for every
`xml:space` attribute until a separate proof covers their error and state
semantics.

## Why the boundary is safe

`scan` calls `check_start` and then `inspect_attributes` for both `Start` and
`Empty` events (`audit.rs:1945-1980`). `check_start` already reads every
attribute's separator, name, equals sign, quote, and value through
`check_attribute_layout` (`audit.rs:2682-2770`, `2792-2955`). Under
`Policy::SOURCE` it also validates XML names, attribute-value references,
namespace declarations, prefix bindings, and expanded-name uniqueness before
returning. For a successful non-`xml:space` tag, `inspect_attributes` adds
checked-iterator traversal, a redundant per-element limit recheck, and the
aggregate attribute charge. The first-pass bound makes the recheck's result
known for the proposed fast cases; the aggregate charge is the state change
that cannot be omitted.

With at most one raw attribute, the checked iterator cannot produce a
duplicate error. With a non-`xml:space` key, `inspect_attributes` does not
decode the value and does not change the inherited space. Consequently, once
the complete first pass has succeeded, the summary can reproduce the only
state change that matters: one aggregate attribute charge.

The key comparison must be the same raw-byte comparison as the current
`attribute.key.as_ref() == b"xml:space"`. It must not use namespace-expanded
meaning, case folding, or a decoded key. The reader and slice paths both read
UTF-8 input, and quick-xml exposes the raw key bytes here.

The fast path must stay after the entire first pass. Charging the aggregate
budget while the first pass is still visiting attributes changes error order.
For example, with an aggregate attribute limit of zero,
`<root a="1" b/>` must still report the first pass's malformed second
attribute; it must not report an aggregate limit for `a`. Only a fully valid
summary may then report `Resource::Attributes` at the tag offset. Conversely,
an element limit of zero must still be reported by `check_attribute_layout` at
the next attribute's `inner_offset + cursor`, before its name or value is
read. These are distinct from the aggregate charge's existing tag-start
offset.

The summary does not weaken the independent per-element bound. The first pass
continues to check it before reading each attribute. This preserves the
bounded hostile-tag property recorded by change 0764; a future implementation
must not remove both checks merely because ordinary tags are short.

## Policy and state matrix

| Caller/policy | Summary eligibility | State that must remain exact |
| --- | --- | --- |
| `verify_source` / `Policy::SOURCE` | `0` or `1` raw non-`xml:space` after complete well-formed and namespace validation | aggregate attributes, namespace bindings, inherited space, source-window counters and typed errors |
| `verify` / `Policy::COMPACT` | same, but only after compact separator/close checks succeed | compactness refusal ordering and inherited space |
| `verify_authored` / `Policy::AUTHORED` | same as compact | ambiguous-whitespace state before and after the tag; non-`xml:space` must not alter it |
| `verify_reader` and `verify_authored_reader` | same two compact policies; no source-policy streaming entry point exists | physical offsets, token/window limits, chunk-independent report and error identity |
| `audit_window` used by replacement proof | same source-policy rule | `Counters::since`, `Counters::replaced`, depth, namespace scope, balanced window and final markup proof |

XML declarations use `check_attribute_layout` directly and do not call
`inspect_attributes`; the summary must not be applied there. Namespace-only
attributes such as `xmlns:p` may use the one-attribute fast case only after
`check_start` has completed its existing namespace mutation and validation.
An ordinary `xml:lang` attribute is eligible because only the exact raw key
`xml:space` changes the auditor's `Space` state.

For a `Start` event, the fast result must be pushed into `state.spaces` exactly
as the inherited value is today. For an `Empty` event it is immediately
discarded, and namespace closing remains in the same place. A source-window
scan starts with `State::within`; the summary introduces no new state read, so
the existing window proof remains valid only if the aggregate counter is
updated exactly as above.

## Parser discrepancies that prohibit a broader shortcut

The manual first pass and quick-xml are intentionally not interchangeable.
quick-xml 0.41's attribute state machine consumes the first non-whitespace
byte before looking for `=`. Thus a raw attribute beginning with `=` can be
represented as an empty key by quick-xml's iterator, while
`check_attribute_layout` rejects it as a missing name (and the source policy
also rejects it as an invalid QName). This is one reason a probe of
`tag.attributes().next()` cannot establish a summary.

The same distinction appears in the existing compact policy: a bare
`<root a/>` is classified by the manual pass as a compactness separation
defect, while quick-xml's iterator would report its own missing-`=` error.
The first pass refuses it before any summary is eligible. The custom
`CheckedAttributes` handoff after 32 names also preserves quick-xml's unusual
key and duplicate-before-value behavior; it must remain untouched for all
multi-attribute tags. No summary may claim that the manual pass replaces that
iterator contract.

The following inputs are explicit counterexamples to broader rules:

* `<root a="1" a="2"/>`: a first-item-only shortcut would suppress the
  duplicate; the complete summary must classify the tag as `other`.
* `<root a="1" b/>` and `<root a="1" b="unterminated/>`: a summary made
  before the full tag scan would change malformed-tail precedence.
* `<root xml:space="maybe"/>`: layout success is not `xml:space` success;
  the iterator must decode and reject the value.
* `<root xml:space="pre&#120;erve"/>`: compact/authored first-pass layout
  accepts this, but the iterator's normalized value is neither `default` nor
  `preserve` and must fail.
* `<root ="v"/>` and the post-32 `=n0` forms: unusual raw keys exercise the
  quick-xml/custom-helper lexical contract and are not eligible for a
  first-pass replacement.

## Required differential matrix before any code or timing gate

The next candidate needs exact result comparison, including `Report` fields
and complete `Error` values and offsets, for both slice and chunked-reader
entry points where available. The matrix should include:

1. zero, one, two, 31, 32, 33, and per-element-limit cardinalities;
2. ordinary, Unicode, punctuation, leading-`=`, namespace, and `xml:*` raw
   keys;
3. single and double quotes, empty values, escaped values, whitespace around
   `=`, non-ASCII XML whitespace, malformed separators, unquoted values,
   unterminated values, `<`/`&`, and duplicate tails;
4. valid, invalid, escaped, inherited, and reset `xml:space` values;
5. exact and one-under aggregate/per-element limits, plus depth, event, token,
   byte, and text limits to verify that no earlier or later refusal moves;
6. UTF-8 failures and BOM-adjusted offsets;
7. source replacement identical, local-window, original-failure,
   replacement-failure, and window-re-totaling cases; and
8. one-byte, uneven, and prefix-sized `BufRead` chunks for compact and
   authored streaming audits.

The parity harness must retain the existing two-or-more-attribute path and
compare it separately, including the 32/33-name checked-iterator handoff.
Only after these semantic gates pass should a candidate be built and measured
under the ADR 0005 evidence contract. ADR 0006's source publication boundary
requires the complete non-compactness checks, UTF-8/well-formedness,
DOCTYPE refusal, `xml:space` validation, and finite budgets to retain their
typed behavior; this summary is acceptable only if it preserves all of them.
