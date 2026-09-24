# ODS date and time performance plan

This is the pre-capture plan for the complete ODF 1.4 Part 4 section 6.10
date and time family. The reviewed semantic input is
[contract.md](../contract.md). The matched baseline is production commit
6fa3b8af6a. The current 600e86f4e0 value-inspection evidence context is retained for
matched controls; it is not the date/time candidate identity. The final
date/time candidate source identity will come from the source freeze. The authoritative isolated gate lock is the retained
Cargo.lock with SHA-256
58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3.
The ambient workspace Cargo.lock is not a profile input.

No timing capture has been authorized by this plan. A future runner must
freeze the candidate source, contract, oracle, fixtures, harness, and lock,
then obtain an explicit root handoff before compiling or measuring.

## Scope and matched controls

The candidate-only workload covers all 24 names:

DATE, DATEDIF, DATEVALUE, DAY, DAYS, DAYS360, EASTERSUNDAY, EDATE,
EOMONTH, HOUR, ISOWEEKNUM, MINUTE, MONTH, NETWORKDAYS, NOW, SECOND, TIME,
TIMEVALUE, TODAY, WEEKDAY, WEEKNUM, WORKDAY, YEAR, and YEARFRAC.

Each name has one small deterministic core case with an independent expected
serial, component, formula error, or typed evaluation failure. The core case
is a coverage anchor; the edge and scaling lanes below add the cases that
actually distinguish calendar algorithms, sequence consumption, timestamp
capability, and resource behavior.

The baseline and candidate retain the existing matched evaluator controls from
the predecessor profiles. This includes the existing VALUE date/time/datetime
and mixed-fraction parsing controls, calendar-helper controls, scalar and
reference aggregate controls, and lazy IF/IFERROR/IFNA controls. The exact
control names, formulas, fixtures, read bounds, and output checksums are copied
from the frozen predecessor matrix rather than rewritten for this batch.
Ordinary aggregate and lazy controls remain useful because a date dispatch or
shared calendar helper must not change unrelated evaluator work, allocation,
or cache behavior. No newly added date function is used as a matched control.

If the baseline does not dispatch a date function, its core and date-specific
rows are candidate-only. A baseline Unsupported(Function) result is a
capability observation, not a numeric output and must not be timed or compared
as though it were a formula result. Before/after deltas are reported only for
the matched controls and for any date lane that both revisions pass with the
same exact result contract.

## Candidate matrix

The reviewed matrix is bounded at 34 matched controls and 86 candidate date
cases (120 named cases total). Each side has 68 timed case-phase groups for
the matched controls; the candidate has 240 groups after the date rows are
added. `case-matrix.json` is the authoritative name, formula, exact-output,
shape, and preflight-read declaration for this workload. The shape field denotes
evaluation geometry; expected payloads independently declare result shape. A
reducer evaluated in a 1x1 matrix context may return a scalar number.

| group | planned coverage | read contract |
| --- | --- | --- |
| core-24 | one exact scalar case for every date/time name | 0 for literal and timestamp-backed scalar inputs |
| parser-profile | ISO, en_US, English month-name, datetime, fractional-second, and numeric-fallback DATEVALUE/TIMEVALUE text | 0 for literals; the datetime TIMEVALUE row asserts that only the clock fraction is returned |
| calendar-boundaries | DATE rollover and leap rules, 1900 non-leap behavior, domain bounds, DATEDIF unit boundaries, DAYS360 methods, EDATE/EOMONTH month-end clamping, Easter bounds, signed total-second rounding, WEEKDAY/WEEKNUM modes including fractional-mode refusal, and YEARFRAC bases | 0 unless the case intentionally uses a reference DateParam |
| dateparam-reference | holiday/workweek sequence references, formula-error DateParam, and direct scalar-list refusal | one successful resolver read per admitted scalar reference; list refusal is 0 |
| projected-date | scalar intersection and elementwise DateParam/Offset inputs, including projected IF branches | exact output-coordinate reads; no range flattening |
| holiday-scale | NETWORKDAYS and WORKDAY with 0, 1, 16, 64, and 256 admitted holiday cells | H successful reads per output for an H-cell holiday descriptor |
| workweek-scale | default workweek, supplied seven-element workweek, all-workday and all-off workweeks, and malformed workweek inputs | 7 successful reads per supplied workweek descriptor; a malformed referenced workweek is fully inspected and has its case-specific read bound, while an inline/static-list refusal is 0 |
| interval-scale | NETWORKDAYS spans of 7, 31, 365, and 4096 days; WORKDAY offsets of 0, 1, 64, 256, and 1024 | sequence reads as above; work scales with checked day stepping |
| sequence-projection | projected NETWORKDAYS and WORKDAY with complete holiday/workweek arguments | each projected scalar consumes its complete admitted sequence; expected reads are derived per output, not assumed invariant |
| lazy-and-cache | selected and unselected IF/IFERROR/IFNA branches, complete sequence arguments under projection, repeated explicit timestamp snapshots, and position-sensitive nested MUNIT criteria | unselected branches read 0; cache identity includes position, shape, timestamp, source, and cancellation context |
| timestamp | NOW, TODAY, and no-Year EASTERSUNDAY with an explicit CalculationTimestamp (the Easter row uses 2024-04-01 so it must select Easter 2025), plus the same three calls without one | 0; missing timestamp is typed Unsupported(CalculationClock), never an ambient-clock result |
| refusal-and-failure | direct ReferenceList DateParam/DateSequence refusal, typed read/provider/source failure, resource limits, work limits, cancellation, and retained formula errors | known shape/type refusal is 0; cancellation has one successful read across four sticky repeats |

The holiday and workweek lanes are separate so a result can distinguish date
iteration work from sequence scanning. For a projected sequence case with N
output coordinates, H holiday cells, and W workweek cells, the preflight read
bound is N times (H plus W) when both descriptors are supplied. The harness
records the exact formula-derived bound for each case. It must not silently
credit a cache hit when the contract requires complete sequence consumption.

The interval lanes use checked civil-day arithmetic. They exercise both short
and long spans while keeping the work budget finite. A large interval may
show increasing work and elapsed time without implying additional resolver
reads. Periodic cancellation checkpoints and the configured work budget are
part of the measured behavior.

## Correctness and resource preflight

An independent oracle must validate every core and expanded lane before
timing. It checks exact serial values and fractional days, civil components,
all formula-error codes and payload identity, matrix shape and coordinates,
timestamp-backed results, and typed evaluation failure kinds. It covers the
contract choices for the 1930 two-digit-year pivot, the 1899-12-30 epoch, the
proleptic Gregorian range, MINUTE and SECOND rounding, DATE rollover,
DAYS360 US/European behavior, reversed NETWORKDAYS/YEARFRAC intervals,
WORKDAY zero/nonzero offsets, and WEEKNUM modes 21 and 150.

Each matrix row records the exact formula text and evaluation path (`scalar`
or `value`). The harness emits both fields in its machine-readable preflight
receipt, and the runner compares them row by row before accepting the read
bound. A formula change or an accidental evaluator-path switch therefore
cannot become an unobserved workload change.

The preflight also verifies that:

- every inspected reference cell, including one skipped by conversion, charges
  work, checks the cell limit and cancellation, reads through the resolver, and
  performs the post-read cancellation check;
- retained formula errors do not stop a complete holiday or workweek scan;
- a typed read, source, cancellation, allocation, resource, or source-version
  failure supersedes a retained formula error;
- borrowed Text remains borrowed during parsing and conversion;
- direct ReferenceList and known shape/type refusals occur before any resolver
  read;
- matrix results preserve scalar intersection versus projected elementwise
  behavior;
- NETWORKDAYS and WORKDAY consume complete sequence arguments for every
  projected result; and
- a missing CalculationTimestamp is a typed capability refusal and never
  consults the wall clock, timezone, environment, or workbook metadata.

The preflight must use exact expected values for scalar and matrix lanes. A
formula-error result is represented as the expected formula error, not as a
number or a one-element array merely because the harness uses a common JSON
shape. Cancellation is repeated four times in the same sticky context and
asserts one total successful read per child. Known shape/type refusals and
zero-cell resource refusals assert zero reads; work-limit, provider, source,
and cancellation lanes use their case-specific partial-read bounds. Provider
and source failures remain typed failures and are never caught by
formula-level error handling.

## Measurements

The runner uses the established two-phase process protocol: evaluate a
pre-parsed expression and parse-evaluate the expression in the child. Setup,
oracle work, resolver construction, source fixture construction, and profile
hashing remain outside the evaluate timer. The timed boundary includes the
evaluation, comparison against prepared expected values, checksum of the complete
result, and result drop. Independent oracle calculations run before timing;
per-repeat comparisons remain timed to detect repeat-dependent semantic failures.
Reported latency therefore includes this validation instrumentation. Each phase uses
three warmups and fifteen fresh child samples unless the final handoff records
a different count.

Each raw sample records monotonic elapsed time, repeat count, p50 normalized
elapsed time, allocator calls and requested/released bytes, end-live and peak
live deltas, retained execution-budget bytes, work charges, successful
resolver reads, input and output bytes, and external RSS/HWM where available.
The report preserves p50, p95, p99, min, max, and the exact sample count for
each lane. It reports allocations and bytes separately from evaluator work;
RSS is a process observation rather than an allocation proof.

The date kernels should be measurable as fixed-size civil-date arithmetic with
no ordinary range-sized allocation. Parser lanes expose borrowed input-byte
charging and bounded scratch. Holiday/workweek lanes expose checked
reservation, retained sequence metadata, holiday membership storage, and
drop/release order. Matrix lanes expose output-array storage and projected
descriptor work. Every result includes a checksum so a lower elapsed time
cannot be retained when a lane returned the wrong serial, shape, or error.

Read and work accounting is reported independently. A scalar date calculation
with a 4096-day interval may have high work and zero resolver reads; a
256-cell holiday reference has bounded sequence reads even when its interval
is one day. The report must not infer a causal optimization from equal
allocator counters or equal reads.

Before capture, the review trigger is preregistered at an absolute normalized
latency or RSS delta of 5 percent or more for any matched control group. Every
flagged row reports both absolute baseline and candidate values (nanoseconds
per repeat for latency and KiB for RSS), the signed relative delta, the exact
sample count, and a descriptive uncertainty interval from a seeded bootstrap
over the retained normalized samples. A trigger requires review and explicit
scope; it does not establish a regression cause or justify a favorable rerun.
Baseline samples are collected before candidate samples; CPU affinity is not
pinned and recorded system load is not used as a rejection threshold. These
ordering and shared-host noise limitations accompany every comparison.
Latency and RSS changes remain descriptive until repeated, isolated evidence
supports a scoped claim.

## Capture and source custody

The final source closure must include the evaluator dispatch and date/time
kernel, shared VALUE/calendar parsing helpers, scalar and value/reference
schedulers, matrix/cache code used by projected date functions, all date/time
tests and limits/oracle tests, the reviewed contract, independent oracle and
goldens, feature matrix, native provenance if present, the harness, and the
isolated Cargo.lock. Recursive date/time and calendar child modules must be
included so a module split cannot escape the freeze. The ambient root lock is
not copied into the profile.

Before timing, the runner must record the baseline commit, candidate freeze
hash, contract/oracle/golden hashes, exact source closure, lock hash, Rust
toolchain, target, compiler flags, allocator, CPU/core count, and process-load
check. It must preflight every named case, retain raw JSON and external timing
receipts, hash binary identities before and after capture, and remove temporary
targets and staging artifacts after capture. It must retain any failed setup or
preflight attempt under a diagnostic directory rather than overwriting it.

This plan establishes workload and evidence requirements only. It makes no
production performance claim and does not authorize a build or timing run.
