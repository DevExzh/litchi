# Change 0099: selectable ODF `mimetype` repair planning

Date: 2026-08-14

## Scope

The standalone performance harness now exposes one opt-in
`odf_mimetype_repair_plan` case for the existing narrow generic-ODF repair API.
It does not change the default 36-case/198-record matrix and takes the complete
selectable inventory from 200 to 201 cases.

The deterministic corpus starts from a valid generated ODT, then adds exactly
one nine-byte Extended Timestamp (`0x5455`) field to the local header of the
first, stored `mimetype` member. Central-directory member offsets and the EOCD
directory offset are adjusted without changing any member payload. The generic
validator must report exactly the supported repair issue before the corpus is
admitted.

## Correctness gates

Untimed verification checks:

- preview schema, repair ID, non-destructive intent, source/output lengths and
  SHA-256 fingerprints, issue identity, effects, and deterministic plan JSON;
- exact recovery of the canonical ODT through both the owned patch and the
  forward-only publisher;
- exact inverse restoration of the malformed source artifact;
- stale-source refusal before output and exact no-plan refusal for the already
  canonical package;
- typed one-byte partial-sink progress; and
- member-payload preservation and complete repaired-package reopen inherited
  from the repair planner.

The timed region includes validation, plan construction (including its bounded
full-candidate preflight), and sequential publication to a SHA-256/discard
sink. The sink retains zero output bytes. Each observed write request remains
below 64 KiB, but the harness deliberately reports no retained authoring window
because the planning preflight materializes a bounded complete candidate.

## Verification

The strict focused test and warning/deprecation-denied all-target harness
Clippy pass. A debug one-sample smoke across tiny, medium, and large semantic
shapes produced exact deterministic source/output/plan hashes, 16 writes per
record, zero retained output, and largest writes of 629, 1,106, and 26,878
bytes. The compact artifact is
[`odf-mimetype-repair-smoke-0109-summary.json`](../results/odf-mimetype-repair-smoke-0109-summary.json).

An independent review confirmed ZIP offset surgery, repair eligibility,
reversibility, refusal behavior, timing scope, default-matrix exclusion, and
the absence of a total-memory claim.

## Claim boundary

This is selectable correctness and counter evidence only. The smoke is a dirty
debug build with no warm-up and one sample per shape. It supports no latency,
allocation, peak-memory/RSS, cold-I/O, throughput, or comparative performance
claim. Destructive, structural, XML-semantic, encrypted, signed, and macro
repair remain unsupported.
