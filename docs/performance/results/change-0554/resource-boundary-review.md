# Name handoff resource boundary

The N mechanism intentionally removes the second normal-entry name allocation. A failure injected at that removed allocation cannot have an identical allocation-call ordinal after the change. This is work elimination, not a promise of allocator-failure schedule equivalence.

The validation pass still creates and validates the first name and its comparison data before graph construction. Existing directory-byte limits, checked extents, fallible directory/graph/cache reservations, and their labels remain in place. Root decoding still uses the fallible `decoded CFB directory name` reserve. Normal-entry transfer performs `Option::take()` and allocates nothing. Error cleanup drops staged public and private owners before final assignment; source review and the focused invalid-directory test cover publication structure, while that test alone does not simulate a later allocation failure.

The pre-existing validation-side `String::from_utf16` is unchanged. Its allocation behavior must not be presented as newly fallible or as full global allocator exhaustion recovery. The removed secondary reserve is an intended allocation-shape change; all remaining typed allocation paths retain their labels and ordering. Existing minifat allocation-overflow and directory-resource tests are included in the CFB test lane. No fault-injection result or allocation-failure schedule parity is claimed.

Measured allocator regions must report actual statuses and all counters. Unavailable or overflow statuses do not become zeros, and operation allocator peaks do not establish process RSS. Retention remains subject to every frozen native, allocator, mechanism, correctness and final-quality gate.
