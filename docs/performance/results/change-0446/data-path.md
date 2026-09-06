# Owned Part-name handoff within the unchanged plain lifecycle

Both builds use the 0445 plain OwnedSource lifecycle. Fresh source/input cloning,
prepared 64 KiB payload and bounded hashing-discard sink exist before entry.
The timer/allocation region includes source-backed catalog opening, one Part/root
relationship plan construction, consuming sequential publication and catalog drop.
Input/payload remain live; digest finalization, source observations, process
endpoint probes and full fixture/gate/report work remain outside elapsed time.

The only planned production change moves the owned decoded Part-name String into
PackURI::new instead of borrowing it and causing Into<String> to clone it. Required
attribute normalization/bounds, URI validation, content-type validation, duplicate
and equivalent-name checks, source freshness, generated-manifest validation and
managed-memory reservations remain unchanged. No parsing pass is skipped.

Source counters are unavailable in both builds. Actual sink/process/allocator
observations remain measured where available. The sink retains no output archive.
Above-entry allocation peaks and endpoint live deltas are distinct from preexisting
fixture/input memory and whole-process RSS. Profiles include setup/gates/warmups/
reporting; run-frame subsets still include untimed preparation/probes. This is a
synthetic OPC one-Part package addition, not semantic Office owner/native/cold/range
or scaling evidence.
