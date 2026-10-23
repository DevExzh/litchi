# 0362um — OPC workspace consumers

The focused workspace consumer check at `aa7ee42eda` reported unnecessary
qualification errors in `litchi-ooxml-common`'s embedded tests.  The tests
already consume the canonical OPC source policy types, so the imports now
include `SourceCacheLimits` and `SourceReadPolicy`, and all policy constructors
use the imported `SourceLimits`, `SourceCacheLimits`, and `SourceReadPolicy`
names.

The policy values and assertions are unchanged.  No aliases, blanket lint
allowances, or new targets were introduced.  The independent transaction,
XLS, and ODP diagnostics from that check remain with their owning workers.

Validation was limited to formatting, whitespace, and repository artifact
guards as requested; no Cargo target was started.
