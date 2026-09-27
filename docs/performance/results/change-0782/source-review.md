# Current source selection

The 0781 allocation trace attributes 192 allocations and 7,680,000 requested
bytes to fresh plain-text conversion across three payload writes. The Cow
candidate removed those requests and improved the payload public operation,
but consistently regressed many-short-text write/lifecycle by 12.121%/10.551%.
That candidate was rejected, its source restored, and its target/worktree
removed. Its sealed packet is historical evidence, not this run's baseline.

The subsequent construction audit found only one production non-empty plain
text assignment: a source `ShapeProperties.text` string that lives through
conversion and encoding. Notes and mutation-sensitive rich/centered paths own
paragraphs independently. An optional borrowed slice removes an unused owned
alternative while retaining the same intended copy removal. This is a
concrete design to test, not proof that representation layout caused 0781's
regression.

Current-source inspection and the inherited corrected probe make this a
bounded follow-up with an explicit common-workflow guard. The separate 0780
large-lifecycle attribution question and path-specific exact-size OPC ingress
reservation remain open hypotheses. The now-fixed default CI matrix is not
repaired again. No coverage registry row is promoted by this experiment.
