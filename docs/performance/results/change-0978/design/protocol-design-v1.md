# 0978 variable-layout decoder integration proposal

This is a protocol proposal only.  It binds the 0978 work to the passing
ADR/GOAL receipt at `receipts/adr-goal-revalidation-0978-v1.json` and keeps
the sealed 0977 runner, raw captures, decoder, reducer, plans, and receipts
immutable.  It does not authorize a build, native command, or new capture.

## Scope

0978 adds one offline variable-layout ABI scan for a retained child packet.
The scanner consumes the exact child `perf.data`, perf mmap dump, and the
existing raw-stack sidecar.  It does not change the perf command, affinity
bridge, same-PID launch chain, current-perf evlist/header policy, or Office
workload.  The existing 0977 decoder remains the base evidence producer; a
small versioned adapter adds the scanner result after the base decoder has
completed.

The scanner must distinguish a valid record with no decoded ABI fields from a
malformed or unaccounted record.  A valid empty record contributes its full
period to an explicit `unknown_period`; it is never silently discarded or
treated as a parser exception.  Malformed records, duplicate sample identity,
missing sample identity, invalid period, and a conservation mismatch fail the
child closed.  The unknown bucket is therefore an accounting result, not a
threshold exception.

## Proposed sealed interface

The scanner module will expose one pure offline entry point after its schema
is sealed:

```text
scan_child(plan, child_dir, base_packet) -> abi_scan
```

The adapter will pass immutable paths and the base packet's exact sample and
period ledger.  It will write one exclusive `<child>.abi-scan.json` sidecar
and add its file binding under `packet.artifacts.abi_scan`; it will add the
same object under `packet.abi_scan`.  Existing 0977 packet fields and gate
semantics remain unchanged.

The minimum proposed object is:

```json
{
  "schema": "litchi.performance.0978.abi_scan.v1",
  "status": "complete",
  "sample_count": 0,
  "total_period": 0,
  "valid_record_count": 0,
  "empty_valid_record_count": 0,
  "unknown_period": 0,
  "resolved_period": 0,
  "period_conservation": {
    "total_period": 0,
    "resolved_period": 0,
    "unknown_period": 0,
    "passed": true
  },
  "policy": "valid empty records are explicit unknown; resolved plus unknown equals total period"
}
```

The scanner author may add bounded histograms or ABI-version counters, but
must preserve these meanings and must not use an `empty` or `unknown` count to
hide lost, malformed, or unparsed periods.  The final field names and schema
hash are pending the reviewer's accounting policy; this proposal is not a
binding schema.

## Minimal wiring

After the scanner and policy receipts are sealed, create fresh 0978 plan and
binding helpers.  Reuse the immutable 0977 v2 runner and its launch/affinity
receipts unless the sealed scanner requires a new child gate.  The preferred
path is for the adapter to fail before returning a packet on scanner errors,
so no runner copy is needed.  If a new runner filename is required for
custody, copy only the thin import/entrypoint and bind its exact SHA; do not
reapply monkey patches or duplicate the process-launch logic.

The 0978 plan should add only these bindings:

```text
base_plan / base_helper_seal       exact 0977 v2 bindings
abi_scanner                        source/helper SHA and schema
abi_policy_receipt                 reviewer-approved accounting policy
abi_adapter                        fresh adapter SHA
adr_goal_revalidation              this batch's passing receipt
```

The plan must retain the exact binary, fixture, perf/tool, decoder shared
object, libdw/libelf, current-perf requalification, and source custody rows
from 0977.  It must set `native_workload_started` and `build_started` false
until root GO.  No receipt may bind a mutable path without bytes and SHA.

## Offline tests before review

Use a temporary copy or hard link of the retained 0977 child-02 artifacts;
never rewrite the retained evidence.  The focused tests should cover:

1. The actual retained child-02 input parses with exact sample count and
   total period from the base packet.
2. A valid empty ABI record increments `empty_valid_record_count` and
   `unknown_period`, while `resolved_period + unknown_period == total_period`.
3. A malformed record, duplicate identity, missing identity, nonpositive or
   altered period, and removed record are each rejected.
4. A 0978 adapter packet preserves the 0977 packet and gate fields, binds the
   exclusive sidecar, and rejects a sidecar/raw/packet hash mismatch.
5. The existing current-perf evlist-v authority and descriptive header
   warning remain unchanged.

The review receipt should bind the scanner source, adapter, policy, test
inputs, actual retained-input output, and all negative-test results.  Only
after that receipt and the final plan are sealed should root consider a GO;
this proposal itself authorizes no workload.
