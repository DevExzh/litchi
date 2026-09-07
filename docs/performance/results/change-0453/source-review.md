# Source ownership and semantic review

An internal Shared(PartData)/Owned(Arc<Vec<u8>>) payload distinguishes the
successful initial capture path from the original decoded-copy fallback. Shared
handles retain OPC's original memory and object reservations. No managed
PartData::into_arc escape or physical ZIP type is exposed through PPTX.

The exact payload helper checks source URI, target URI, content type, declared
size and every byte before returning reused_previous=true. Only that explicit
result permits retaining the prior compressed capture during planner rerun.
Fresh PartData loaded under cache bypass may own a different allocation; exact
identity still reuses the prior staged handle. Metadata/byte mismatch prevents
capture reuse. Semantic equality and planned-byte counts remain decoded-byte
properties, independent of the internal ownership variant.

Successful compressed publication does not create a bare decoded Arc. A Memory
refusal at writer authorization or target-name staging triggers owned_arc, which
checks execution before allocation and uses the existing chunked checked copy.
Non-Memory errors fail closed. Full existing decoded destination staging remains
reserved for that copy; a checked additional per-leaf bound covers the inline
size difference between the new enum and the old Arc, including tiny/empty data.
No reduced-admission or universal zero-copy claim is made.

The native-input unit test uses a one-byte cache, proves distinct reread
allocations and exact prior-handle reuse, checks metadata/byte mismatch, retains
managed bytes after package/capture drop, rejects bare Arc escape, copies the
fallback under a reservation, checks cancellation, and releases all byte/object
charges. The public test applies pressure only to an independent source budget,
while destination staging uses its own finite budget. Control Store versus
pressure Deflate proves actual decoded fallback, while every decoded member and
source media read count remain unchanged. Both budgets finish at zero.

Independent read-only review supported the design. Root resolved reviewer
concerns about pressure by demonstrating distinct source/destination budgets,
and added a physical compression discriminator after review correctly caught a
missing assertion. The reviewer also flagged bare size_of; pinned Rust 1.98.1
exposes it in the prelude, as prior strict OPC validation already demonstrated.
Final strict compilation, not that provisional concern, is authoritative.

No production dependency, public API, unsafe implementation, runtime pool or
ambient provider is added. The existing allocator wrapper stays in its isolated
benchmark binary; only its command dispatch and mechanical test comparisons
change. Accepted ADR tree remains c950b6c8be822561b498d7bbe87c460873dcbf49.
