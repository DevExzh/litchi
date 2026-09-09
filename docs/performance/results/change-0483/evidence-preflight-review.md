# 0483 evidence preflight review

This began as a read-only review of the 0483 evidence helpers and their
retained receipt protocol. The bounded verifier and pilot-test-fixture changes
described below preserve their prior bytes under content-addressed history;
there were no builds or Git operations.
The selected 0483 target is now explicitly the compare harness
`docx_bounded_tail_append_compare`; the similarly named plain harness is a
separate historical target and is outside this bundle's contract.

## Blocking findings

1. **Harness target mismatch is resolved by explicit selection.**
   `build.py` selects `docx_bounded_tail_append_compare`, `capture.py` emits
   `--route`, and `analyze.py` accepts
   `docx-bounded-tail-append-comparison-v1`. The separately named
   `docx_plain_paragraph_tail_append.rs` accepts `--mode total|phases`, emits
   `docx-plain-paragraph-tail-append-v1`, and has `total_samples`/
   `lifecycle_phases` fields. Keep the selected compare binary/schema/route
   binding explicit in the frozen protocol and never mix the plain records
   into this bundle.

2. **Freeze does not perform the final chronology check before writing.**
   `capture.freeze()` creates `frozen_utc` and runs shape/custody checks, while
   the complete `verify()` chronology check runs only after `protocol.json`
   exists. Concurrent or late receipt/build creation can therefore produce a
   written protocol that is already too old. The coordinator must serialize
   freeze against receipt creation and the pre-freeze path must enforce the
   same condition as the final verifier: every accepted build copy, accepted
   gate/pilot, and fuzz receipt finishes at or before `frozen_utc`; pilots also
   finish before formal captures start.

3. **Git revision is part of the machine-readable result and must be bound.**
   The current build patch records the full `git rev-parse HEAD` in each build,
   the combined binary custody, and the analyzer summary, and the verifier
   checks the 40-hex value and agreement between normal and allocator builds.
   Preserve the before/after HEAD check around each accepted build and the
   exact source-manifest equality already required by accepted gates and fuzz
   custody. If the accepted protocol permits a dirty checkout, add a
   deterministic dirty/status record (or an allowed-path list); otherwise make
   the source-checkpoint commit and clean-tree precondition explicit. A later
   checkout with unchanged source files must not silently change the reported
   revision, and the final protocol/summary must expose the bound value.

## Custody and schema gaps to close before freeze

- Accepted final labels are now operationally run after fuzz preparation with
  `--attempt accepted`, which resolves the earlier `attempt: null` ambiguity.
  Keep the exact suffixed label, full `argv`, source-before/source-after, and
  accepted source digest in the plan. The `env TMPDIR=...` wrapper is retained
  in the exact argv and therefore remains auditable.

- `check_pilot_argv()` accepts the direct 11-argument form and the timed
  18-argument form. It checks that the report argument has the expected suffix,
  but a different absolute path with the same suffix could still be used while
  the verifier reads the in-bundle report. Require the argv report path to
  resolve exactly to the planned `report_path` (and retain/check the resource
  file for the timed form), or make the accepted pilot contract direct-11 only.

- Fuzz custody validates seed ZIPs, hashes, receipt metadata, binary identity,
  and ordering, but does not fully bind the `prepared -> build -> smoke`
  inputs or enforce the recorded build/smoke command shape and corpus inventory.
  At minimum require build `inputs` to equal the prepared receipt, require the
  accepted source/manifest/lock references and fixed ASan build argv, and bind
  smoke argv, seed, max length, timeout, corpus-before inventory, and retained
  post-run inventory to the accepted fuzz plan. Otherwise a structurally valid
  three-receipt set can describe a different fuzz run.

- `check-non-iwork-examples.py` depends on
  `workspace-iwork-exclusion.json`. The Python driver is in the bound helper
  list, but the JSON input must also be content-addressed in the protocol (or
  its exact hash must be recorded in the gate receipt); otherwise its exclusion
  policy can change after the gate. If process profiles are formal evidence,
  bind each `profiles/<attempt>/profile.json`, raw-profile metadata, and its
  binary/gate references as well. If they are diagnostic only, say so in the
  protocol and exclude them from formal claims.

## Statistics and oracle checks

- The analyzer independently derives the main XML SHA-256, opaque payload
  bytes/SHA-256/CRC, route XML, and archive member checks. Preserve those
  checks; the producer's `_corpus` fields alone are not sufficient. Read
  histograms correctly allow requests over 64 KiB, and the sink maximum is
  correctly capped at 16 KiB.

- `validate_sink()` checks histogram sum, accepted bytes, digest, and
  `largest_write <= 16 KiB`, but does not verify bucket bounds or that the
  histogram is consistent with `largest_write`. Add the same lower/upper
  weighted bounds used for reads and reject nonzero `bytes_over_16384` (and
  any impossible largest-write/bucket combination). Otherwise write-size
  statistics can be forged while the archive digest remains valid.

- Allocator conservation is checked. Add the producer invariants for
  `peak_live_bytes_before <= peak_live_bytes_after` and
  `region_peak_live_bytes <= peak_live_bytes_after` if those counters are
  retained as formal evidence; otherwise mark them descriptive rather than
  statistical claims.

Historical/developmental failure classification is structurally sound: every
retained terminal receipt is required to be either an exact accepted plan
label or an explicitly classified developmental record with source-before and
source-after hashes. Keep that exact-union check when adding the accepted
revision and chronology bindings.

## Reviewer status

The pilot path and fuzz custody checks are now tightened in `verify.py`, with
the prior verifier retained under its content-addressed driver-history path.
The new pre-freeze chronology helper and exact evidence-input checks are ready
for `capture.py` to call. Accepted pilots should use the direct 11-argument
form; the timed form is retained only with its fixed report/resource paths.
The 21 existing evidence tests pass, and the retained accepted fuzz records
pass the tightened in-memory custody check (three receipts and 62 seeds).
