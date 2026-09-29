# 0833 — current filesystem qualification and baseline

Base: `eefaca16e39ace3c1e6219aedbac4fe1c0cd060a`.

The next evidence gap is the intersection of larger unchanged package members,
mandatory versus eager opening, default durable atomic publication, and verified
page-cache state. Record 0772 measured a sequential-sink OPC mutation, not this
filesystem lifecycle. Record 0491 measured an older DOCX full-text lifecycle;
0816 measures delayed caller reads, not page-cache state. None supplies current
measurements for this matrix.

Start with six existing selectors: eager/source-backed OPC open, eager/source-
backed OPC one-Part atomic save, and eager/source-backed PPTX open plus selected-
slide lifecycle. Qualify each on warm and strict `cold-verified` states with one
sample and zero warmups. Keep all failures and ineligibility reasons. Do not
replace stale corpus or output hashes merely to obtain timings. A failed oracle
requires diagnosis before any formal measurements. Qualification timings are
not a baseline or performance claim.

The OPC corpus is four incompressible 4 MiB logical members. The PPTX corpus is
an existing generated media-rich presentation. Existing harness constants and
semantic/readback checks remain the admission authority; no new producer or
native-Office coverage is implied. Warm and cold ZIP archives may differ in
EOCD alignment comments; compare routes within a cache state and retain both
source identities. Cold proves only zero pre-operation resident/dirty/writeback
bytes, observed post-operation residency and positive process read_bytes under
the strict fincore policy. It does not prove physical-device cache temperature.

After successful qualification and offline reader review, freeze a separate
formal schedule before collecting any baseline samples. Measure routes and
cache states as configuration comparisons on unchanged production. Do not infer
a historical before/after speedup. Keep native timing, process RSS, logical reads,
procfs I/O, output bytes, and cache materialization evidence distinct. A native
binary has no allocator observer; this packet cannot claim allocation savings.

Root owns all executions, Git operations and temporary cleanup. Delegated work
is limited to drivers, offline readers and source review. Existing unrelated
review/design/matrix files must remain byte-identical. Previously read GOAL,
CRUD checklist, ADR index and all 32 accepted ADRs match the 35 normative hashes
retained by 0832, checked again before this batch. Production source is unchanged.

ADR mapping: 0003 preserves snapshot/publication contracts; 0005 governs source,
cache, resource and evidence scopes; 0006 preserves all validation and byte
oracles; 0008 requires qualification before performance claims; 0010/0011/0024
retain substrate ownership. No architectural exception or new ADR is needed.

The independent PPTX review also identified a potential discarded notes
snapshot during opened capture. Retain that as an unmeasured follow-up; it is
not authorization to remove notes validation or required fingerprinting.
