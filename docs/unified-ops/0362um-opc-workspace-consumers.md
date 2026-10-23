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

Validation was limited to formatting, whitespace, no-caller search, and
repository artifact guards as requested; no Cargo target was started.
