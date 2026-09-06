# ZIP first-capture source review

The only production change is in soapberry-zip/src/office.rs. The new explicit
read_entry_precompressed_and_decoded_with_progress returns a decoded Vec and the
same private-field VerifiedPrecompressedEntry accepted by the preservation writer.
It performs one compressed source capture, then one decode of that immutable
capture. The prior expected-byte API dispatches through the same implementation
with a Compare output and preserves its expected-byte equality requirement.

CapturedDecodedOutput distinguishes Compare from Collect. Collect reserves the
entire declared decoded length only after entry/method/local-layout and platform
size checks, before source capture. Each chunk's checked end must fit the declared
length before append; no unbounded growth follows a dishonest decoded stream.
Final decoded length, actual CRC and full Deflate compressed-span consumption
remain shared. Zero-CRC compatibility continues to record the actual CRC. Failure
drops the decoded buffer and compressed capture without returning either value.
Callbacks still abort immediately; compressed-source transport failures remain
distinct from archive and callback failures. Unsupported methods, index bounds,
strict layout, descriptors and ZIP64 remain governed by the existing ZIP owner.

This ZIP token proves physical bytes, not Office semantics or source currency.
OPC still must bind source lineage/version, validate format content and relationships,
reserve capture/decoded/writer memory, and carry those reservations through cache,
plan and publication lifetime. No public semantic API or ordinary cache changes in
0450, and no production caller adopts this new method yet. It is the primitive
needed to combine first decoded materialization with later compressed transfer.

The combined operation retains compressed C plus decoded U payload bytes and
bounded scratch. It does not establish constant-memory streaming, a lower peak,
or allocator high-water behavior. The tests compare exact source-read counts from
independently indexed immutable inputs. The first exploratory comparison reused
the control's layout cache; it remains in history and is not accepted evidence.
The final comparison uses fresh indexes for each alternative. It proves a source
pass is removed in the explicit combined ZIP operation, not a PPTX speedup.

Accepted ADRs were read earlier in full, and their tree is unchanged:
c950b6c8be822561b498d7bbe87c460873dcbf49. Physical ownership, explicit callers,
immutable inputs, bounded resources, cancellation and preservation follow
0002/0003/0005/0006/0010/0011/0024. No dependency, executor, unsafe policy, ambient
network, source-version contract or ordinary facade change is introduced.
