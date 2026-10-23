# 0362um — OPC workspace consumers

The focused workspace consumer check at `aa7ee42eda` reported unnecessary
qualification errors in `litchi-ooxml-common`'s embedded tests.  The tests
already consume the canonical OPC source policy types, so the imports now
include `SourceCacheLimits` and `SourceReadPolicy`, and all policy constructors
use the imported `SourceLimits`, `SourceCacheLimits`, and `SourceReadPolicy`
names.

The completed v15 workspace check then reported 340 further unnecessary
qualifications across 50 XLSX, PPTX, and DOCX test/source-test consumers. All
diagnosed `litchi_opc::SourceLimits` and `litchi_opc::SourceCacheLimits`
constructors now use their in-scope canonical names. The two XLSX slicer
fixtures also pass the required `SourceLimits::default()` to
`OpcPackage::from_vec`. The four XLSX package test constructors were migrated
to `SourceBackedPackage::from_read_at`, preserving default limits, exact read
policy, cache limits, and managed execution contexts.

The completed v16 workspace check reported one final consumer batch. The 31
XLSB test qualifications now use the imported `SourceLimits`. The XLSX cell
value and row-visibility subtests, PPTX source tests, and XLSX workbook tests
now import the required `SourceLimits`, `SourceCacheLimits`, and
`SourceReadPolicy` from the owning modules. The old private DOCX owner
`write_to_with_policy` helper had no callers after the adapter migration and
was removed rather than silenced.

The policy values and assertions are unchanged.  No aliases, blanket lint
allowances, or new targets were introduced. The independent DOCX production,
XLSB theme, and ODP example diagnostics from that check remain with their
owning workers.

The completed v17 workspace check reported the remaining consumer diagnostics.
All 65 unnecessary qualifications now use the imported policy or core names.
DOCX source-backed test constructors now use the DOCX one-argument
`SourceBackedPackage::from_read_at`; the explicit five-argument constructors
in the OPC consumers retain their limits, cache policy, read policy, and
execution context. SmartArt promotion, XLSB metadata reconstruction, and
XLSB test owners now pass the retained OPC read limits. PPTX metadata uses the
handle's retained limits when it re-admits a revised source package, while an
unrevised owner continues to adopt its retained source graph. The four private
`ooxml_read_limits` accessors are ordinary functions because their state is
behind `Arc`; no const API was required. XLSX streaming subtests now import
`SourceReadPolicy` only in the module that uses it.

The v17 policy values, assertions, source retention, and cache behavior are
unchanged. No concurrent budget changes were copied. Validation remained
limited to formatting, whitespace, source-constructor/qualification audits,
and repository artifact guards; no Cargo target was started.

The completed v21 workspace check reported 60 remaining consumer diagnostics.
The ODF raw validator now imports its OPC `SourceLimits`, and the unified ODP
facade exposes the exact `OpenOptions` retained by its owner for re-admission.
All 35 ODF common qualifications, three XLSB test qualifications, and four
DOCX glossary qualifications now use their in-scope names. The XLSB and DOCX
source-backed tests import `SourceCacheLimits` while preserving their existing
cache and source policies. The ODF writer tests likewise retain the same
`StreamingArchiveLimits`; only redundant module paths were removed.

The v21 policies, assertions, and constructor arguments are unchanged. No
default was widened, no lint allowance or alias was added, and no Cargo target
was started.

The v23 performance-baseline consumer check then exposed the remaining
source-backed harness API drift. DOCX tail-append plans now use the owning
`tail_append` module's `Plan` and `CandidateProof` names, while the publication
helpers retain the deferred package write path. DOCX edit replay and inverse
oracles use the patch's canonical target projection in place of the removed
commit snapshot getter; this keeps the measured lifecycle from materializing a
deferred archive merely to inspect its document identity. The raw DOCX codec
now imports the OPC `Part` trait, uses the three-argument redaction snapshot
constructor, and uses the strict-default hyperlink planner.

The ODF attribution phase now retains the family error type, and buffered ODP,
ODT, and ODS manifest checks pass the owner XML-stream policy explicitly. The
OPC Part-addition harness supplies its required source limits while preserving
the existing cache and exact-read policies. The XLSX allocation harness
disambiguates the filesystem `OpenOptions` from the owner options. Assertions,
workload boundaries, source policies, and output/fidelity checks are unchanged;
no aliases, lint allowances, or Cargo targets were introduced.
