# Review of the 0437 common seam draft

Sources reviewed read-only:

* `/tmp/litchi-goal-0437-odp-common/generated_xml.rs`
* `/tmp/litchi-goal-0437-odp-common-tests/generated_xml_prelude.rs`
* `/tmp/litchi-goal-0437-odp-audit/design.md`

No build, test, formatter, AST, script, or profile was run. The review below
is based on control-flow inspection and the supplied test cases.

## What is correct

The draft keeps `try_new`'s implementation and strict event grammar unchanged
(`generated_xml.rs:93-114, 598-690`). The new `try_new_with_prelude` is a
separate opt-in method, and the test patch confirms that an `Empty` fixed child
is rejected by the old constructor while accepted by the new one
(`generated_xml_prelude.rs:191-195`).

For a valid final path, the report arithmetic is correct. The new parser takes
`stack.len()` at the first suffix event as the insertion depth
(`generated_xml.rs:720-723`); the reader's `add_fragment` then adds fragment
depth to that value and takes the maximum with the fixed shell audit
(`generated_xml.rs:332-375`). The shell audit is over `prefix + suffix`, so it
includes fixed balanced children, their attributes/events, and their deepest
depth (`generated_xml.rs:460-480`). The prelude tests' direct audit oracle and
the explicit depth-4 fixed shell/depth-6 composed path check exercise this
success path (`generated_xml_prelude.rs:148-188`). EOF subtraction and checked
bytes/events/attributes/text/fragment additions remain unchanged.

Boundary handling is also sound for the cases supplied: offsets are compared
around every quick-xml event, UTF-8 is checked before parsing, and the stack
compares raw qualified names (`generated_xml.rs:692-830`). The patch covers
split names, `/>`, a split UTF-8 attribute value, and split end tags
(`generated_xml_prelude.rs:280-295`). DTD/PI/comment/text/reference/
`xml:space`, extra roots, and suffix non-end events are refused before a
fragment can be emitted.

The publication error boundary is preserved. `prepare` still validates limits
and audits the shell before the callback; writer archive admission is tested
with zero callbacks; first/later producer failures retain their source error,
accepted progress, and poisoned finalization; and writer XML limit failures
retain typed resource/actual/maximum information
(`generated_xml_prelude.rs:297-444`).

## Finding 1 — phase tracking permits a fixed child inside the eventual insertion path

**Severity: contract blocker if “balanced children before the final path” is
literal.**

`parse_prelude_envelope_shape` uses `trailing_starts` only as an end-of-prefix
check. `Event::Empty` and every prefix `Event::End` reset it
(`generated_xml.rs:752-782`), but neither event records that a fixed child was
seen below an ancestor which remains open at the insertion boundary. Therefore
this input is accepted by inspection:

```text
prefix = <root><slot><fixed/><inner>
suffix = </inner></slot></root>
```

The parser observes `root`, `slot`, `fixed/`, and `inner`; its stack at the
boundary is `[root, slot, inner]`, `insertion_depth` is 3, and the final
`trailing_starts` is 1. The suffix matches the stack, so construction succeeds.
But `fixed/` is inside `slot`, an ancestor of the final insertion path. It is
therefore a balanced child *after* the insertion path has begun, contrary to
the design's `balanced fixed children*` followed by `final open insertion path`
grammar. The same shape with a nonempty balanced child (`<fixed></fixed>`)
has the same issue.

The current ODP fixture does not expose this because all fixed children occur
before `<body><presentation>`. Add a negative test with the exact shape above,
and a positive control where the balanced child closes before the final path:

```text
accept: <root><fixed/><slot>     + </slot></root>
reject: <root><slot><fixed/><inner> + </inner></slot></root>
```

Then either make the parser maintain an explicit “final path has begun” phase
(with the API's chosen definition of the first final-path ancestor), or state
that balanced children are allowed under any still-open ancestor and rename the
contract/tests accordingly. The latter is materially broader than the 0437
design and should not be accidental. The current `trailing_starts` check alone
does not enforce the narrow interpretation.

## Finding 2 — fixed character-data policy is narrower than the design text

**Severity: contract decision required; not a counter bug if the refusal is
intentional.**

The draft documents and implements refusal of all fixed `Text`, `CData`, and
reference events (`generated_xml.rs:121-124, 800-804`). The patch tests text,
CDATA, and `&amp;` rejection both at the outer root and inside a balanced
child (`generated_xml_prelude.rs:207-233`).

The design requirement says the common shell report must include “all fixed
bytes/events/attributes/text” (`design.md:164-169`) and distinguishes
“unexpected text” from expected fixed content. If the common seam is intended
to support fixed balanced XML with character data, this draft over-restricts
the public contract and makes `text_bytes` impossible to exercise from a
prelude. The direct audit oracle currently verifies `text_bytes` only from
fragments (`generated_xml_prelude.rs:317-353`).

Choose one contract explicitly before integration:

* If fixed text is allowed inside a balanced fixed subtree, accept it only
  while inside that subtree, continue rejecting text outside the root/between
  fixed siblings/in the insertion path, allow only predefined/numeric
  references, and add a positive fixed-text report/`TextBytes` limit test.
* If the ODP provider deliberately needs an element-only shell, retain this
  refusal but amend the common design/contract to say fixed character data is
  forbidden. Keep a test that proves fixed text is refused and do not claim
  that the seam supports arbitrary fixed text.

This does not affect the current fixture's shell counters, which contain no
fixed text; it affects the stated reusable common API contract.

## Finding 3 — negative constructor tests do not verify typed refusal

**Severity: test gap.**

`assert_rejected` checks only `.is_err()` (`generated_xml_prelude.rs:197-204`).
That proves refusal but not the promised common error category. Add at least
one assertion that malformed shape, QName mismatch, boundary split, forbidden
event, and no insertion path return the crate's typed `Error::InvalidFormat`
(with the generated-XML context). Keep the existing limit tests, which already
verify that a valid shell one-under a caller limit reaches `xml_limit()` with
the exact resource/actual/maximum and zero callbacks for shell preflight.

Do not require every constructor error to be `InvalidFormat`: an immutable
hard-limit violation must retain the existing typed XML-audit limit error, and
allocation failure must retain `Error::Allocation`. The test should distinguish
malformed shape from those cases rather than flattening all errors.

## Additional focused coverage gaps

These are lower-risk gaps rather than observed arithmetic defects:

* Add a positive boundary case with a balanced child ending immediately before
  the final insertion path (`<root><fixed></fixed><slot>`). This protects the
  intended `End`-then-`Start` transition while retaining the existing split
  boundary negatives.
* Add an explicit shell-depth-dominates test where the fixed balanced child is
  deeper than both the insertion path and a shallow fragment. The current
  depth-6 case proves the opposite maximum (path plus fragment), and the direct
  audit equality would catch a wrong result, but a named assertion makes the
  fixed-shell contribution clear.
* Keep the existing repository generated-XML tests in the integration set;
  the prelude patch's strict-constructor check covers only the new `Empty`
  case, while the old tests cover mismatched suffixes, inherited
  `xml:space`, prefix text, producer contracts, and aggregate counters.

## Integration disposition

There is no observed regression in the strict constructor, valid-shell
counter/depth arithmetic, event-boundary checks, or typed writer limit/source
error propagation. Before accepting the draft, resolve Finding 1's grammar
phase explicitly and record the fixed-text decision in the common contract.
Add the typed malformed-constructor assertion and the small positive/negative
phase tests. No production or repository test file was changed by this review.
