# 0806 amendment preflight reader review

Review status: **static review passed; no reader change required**. This review
covers `analyze.py` and `root_native_audit.py` only. I inspected their source,
the frozen plan and fixture contracts, the final validator, and the existing
audit descriptor. I did not run either reader, replay native data, invoke Cargo,
build, capture, or profile. The existing `root-native-audit.json` was already
present before this review; I did not create or modify it.

## Numerical calculation

Both readers implement the declared nearest-rank process p50 as
`sorted(samples)[(n + 1) // 2 - 1]`. For the fixed 30 samples this selects the
15th ordered sample, matching the plan's `ceil(n/2)-1` contract. Each row uses
six paired after/before process-p50 ratios in block order. Both implementations
use the fixed seed `806082`, six draws per bootstrap median, 10,000 resamples,
and sorted endpoints 250 and 9749. The root reader additionally rejects
non-finite and non-positive ratios before bootstrapping. The independent
implementations therefore agree on the statistic, pairing, seed, and interval
endpoints while failing closed on malformed timing input.

The primary reader's result is not sufficient by itself: it permits a zero
elapsed sample and does not independently recompute clone hashes. The root
reader rejects non-positive elapsed samples, recomputes every paired row, and
requires the primary `analysis.json` rows to equal its own rows. The packet
validator repeats the row calculation and checks every analysis row before it
accepts `decision.json`.

## Semantic and clone oracles

The primary reader independently reconstructs the expected construct and
consume checksum, accepted count, and error marker from each frozen trace. The
root reader reconstructs all 39 case identities, source encodings, attribute
hashes, error positions, error markers, and sequence hashes from the sealed
fixture shape. It then checks the semantic oracle's quick-xml, baseline, and
candidate traces, iterator sizes, exact expected loop result, and all 30 sample
result fields.

The root reader also computes the exact suffix sequence hash for every clone
advance `[0, 1, 2, 3, 4, 5, 32, 33]`, including the terminal behavior after an
error. It compares each complete clone object, rather than only its boolean
flags. This closes the primary reader's intentionally smaller clone check.
The 0805 cases and fixtures are required to be byte-identical, so the new
reader cannot silently replace the semantic corpus while accepting the same
row shape.

## Custody and final decision binding

The root reader independently binds the five-file constructor rewrite, the
unchanged shared test, both source archives, the sealed probe inputs, both
build source censuses, build chronology, binary identities, all 936 native
receipts and 28,080 samples, receipt order, commands, report/log/RSS paths,
positive RSS values, and the no-profile/no-history scope. It also binds the
application source census and rejects source changes between the before and
after build records. The primary reader checks the same broad custody contract
and the root reader's `--check` mode requires the retained audit JSON to
reproduce byte-for-byte from the current packet.

The final packet validator supplies the last binding: it recomputes all rows,
checks the decision fields and analysis descriptor, invokes
`root_native_audit.py --check`, and checks the retained audit's exact source,
build, capture, cardinality, and policy witnesses. The amendment application
script also verifies the decision's audit descriptor bytes and SHA-256 before
applying any source patch. Consequently, a primary-only analysis or a stale
audit cannot establish the final decision under the documented workflow.

One bounded review note remains: `analyze.py` alone checks clone schedules and
boolean clone flags, and its capture reader accepts non-negative elapsed values
without requiring numeric RSS content. Those are weaker intermediate checks,
not an unclosed final-custody path, because the independent root reader and
the final validator enforce exact clone hashes, positive elapsed/RSS witnesses,
and the decision-to-artifact binding. No amendment to either reader is needed.

## Disposition

The readers are statically suitable for the root-owned preflight handoff. This
review makes no performance, resource, allocator, profile, or production
adoption claim. The root-owned replay and final validator remain the execution
evidence for the retained result.
