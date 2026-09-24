# Matrix `IF` shape and lazy-branch review

Status: accepted bounded profile for the matrix lazy-evaluation boundary. This
file records the normative evidence and the profile decision for an
unselected branch. It does not change the committed array/reference
specification review or claim that the replacement value VM has passed an
implementation gate.

## Primary source

The source is the checked-in ODF 1.4 archive
`3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`.
The reviewed entry is
`part4-formula/OpenDocument-v1.4-os-part4-formula.html`, SHA-256
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
The relevant anchors are:

* `a_3_2_3_Operator_and_Function_Evaluation` (§3.2.3);
* `a_3_3_Non-Scalar_Evaluation__aka_'Array_expressions'_` (§3.3); and
* `a_6_15_4_IF` (§6.15.4).

## Normative rules

Section 3.2.3 establishes eager argument evaluation, with a function-specific
exception:

> “The value of all expression arguments are computed. Exceptions to computation of all arguments are noted in a function's specification.”

It names IF as the example of a function that does not require all argument
expressions to be computed. Section 6.15.4 is explicit:

> “This function only evaluates IfTrue, or IfFalse, and never both; that is to say, it short-circuits.”

Section 3.3 §2.2.2 separately states:

> “The result matrix is rectangular, sized with the maximum number of rows and columns from all non-scalar arguments.”

The specification does not say whether an IF branch that §6.15.4 requires the
evaluator to leave unevaluated is included in that maximum. It therefore does
not resolve the shape of `IF({TRUE()};{1};{2|3})` by itself. The generic
maximum-shape sentence cannot be used to require evaluating or resolving the
unused branch, because that would contradict the explicit IF exception.

## Selected-branch-only profile

For a formula evaluated in matrix context, the result shape is the checked
rectangular shape of the condition together with branches selected by at least
one condition position. A branch selected at no position contributes neither
values nor shape. A selected branch may be evaluated only after the condition
has selected it; its reference geometry and values are then subject to the
normal resolver, work, memory, and cancellation budgets. The result retains
matrix kind even when its dimensions are 1×1.

This gives the following profile results:

```text
IF({TRUE()};{1};{2|3})             -> 1×1 matrix containing Number(1)
IF({TRUE()};1;[Missing.A1:Z100])  -> 1×1 matrix containing Number(1)
```

The second example must not ask a worksheet/reference resolver for the
`Missing` sheet merely to discover the unused branch's dimensions. It must not
read that reference, charge its work or memory, perform its cancellation
checks, or expose a missing-sheet/unsupported-reference failure. Parsing the
branch into a valid AST is a prerequisite for parsing the whole formula and is
distinct from resolving its reference metadata; parser-wide formula limits may
still apply.

If the condition selects the false branch, that branch's shape and reference
resolution become applicable and its resolver or formula errors propagate under
the ordinary function rules. With mixed condition elements, the bounded
profile admits the shapes of branches selected at least once, then applies the
rectangular broadcasting rules to the selected positions. It must not resolve a
branch that is selected nowhere.

A different implementation could inspect both inline-array shapes
syntactically and use the generic maximum, but that is a profile choice rather
than a requirement established by the cited text. Such a profile still cannot
resolve an unused external or worksheet reference to learn its shape, and must
preserve IF's no-evaluation guarantee for its values, errors, resource charges,
and side effects.
