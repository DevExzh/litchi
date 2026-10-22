# 0359r — RTF refusal payload layout audit

- Date: 2026-10-22.
- Status: checkfree; no implementation change is accepted from this experiment.
- Scope: the small unsupported-content refusal route measured in 0359b.

The 0359b probe showed a 1.0328413 after/before ratio for
`unsupported_content`, with identical allocation counts and peak heap. The
remaining cost was audited against the old RTF error payload rather than
waived as noise. The current shared value has two `&'static str` fields and a
`RefusalKind`; a normal layout is 40 bytes at alignment 8. The old RTF-only
unsupported-source payload held one `&'static str` and was 24 bytes including
its enum storage. This makes the extra stores a concrete representation cost,
but the compact representation must also preserve the shared public
constructor, accessors, `Copy` semantics, fixed content-free text and all
error classifications.

## Packed experiment

An isolated build temporarily added `#[repr(C, packed)]` to the public
`litchi_core::edit::Refusal`. No constructor, accessor, formatter, trait,
serde or redaction code was changed. The formatter needed local copies because
Rust correctly rejects references to packed fields. Core tests passed after
that mechanical adjustment. The resulting layout probes were:

| Type | Normal | Packed |
| --- | ---: | ---: |
| `Refusal` size/alignment | 40 / 8 | 33 / 1 |
| `litchi_core::Error` | 48 / 8 | 48 / 8 |
| `litchi_core::sheet::Error` | 40 / 8 | 40 / 8 |
| `litchi_rtf::edit::Error` | 48 / 8 | 40 / 8 |
| `Result<&mut edit::Edit, edit::Error>` | 48 / 8 | 40 / 8 |
| `litchi_rtf::tail_append::Error` | 48 / 8 | 40 / 8 |
| `Result<&mut tail_append::Edit, tail_append::Error>` | 48 / 8 | 40 / 8 |

The eight-byte RTF enum reduction does not carry through the core errors or
the core sheet error. It also changes a public type's alignment to one, so
every by-value copy and every containing public value must tolerate unaligned
fields.

The release refusal-contract benchmark executable grew from 4,139,552 bytes
to 4,152,000 bytes. In `Edit::set_text_underline`, the normal release body is
416 bytes and 111 instructions. The packed body is 528 bytes and 132
instructions. The packed body adds a stack frame and spills, and emits
unaligned stores at offsets `+0x2`, `+0xa`, `+0x12` and `+0x1a`; the normal
body stores the same fields at aligned offsets `+0x8`, `+0x10`, `+0x18` and
`+0x20`. That is a regression on both the refusal path and the method's
successful producer body, regardless of the smaller returned enum.

The complete six-run native qualification was intentionally not repeated
while the host was under load. The static code-size and instruction evidence
already rejects this representation; it would not be responsible to qualify
an implementation that is visibly larger on the hot producer method.

## Decision

The packed experiment is discarded and the shared `Refusal` remains unchanged.
A portable aligned alternative would need to replace the two arbitrary static
strings with an interned descriptor pointer or an index into a fixed table.
That changes the existing `Refusal::new(kind, operation, reason)` semantics,
requires caller-side descriptors or allocation, or loses the public accessors'
exact strings. A format-local compact error variant would violate the shared
`Error::Refused(Refusal)` contract. None is acceptable for this audit.

As a separate RTF-only code-layout experiment, both private constructors were
temporarily changed to `#[cold] #[inline(never)]` functions. This was intended
to keep fixed refusal construction out of successful edit bodies while leaving
the shared value unchanged. The 380-test RTF library suite passed, and a
non-specialized probe confirmed that a valid span still succeeds while an
empty span returns the same `UnsupportedContent` refusal. The experiment made
the exact producer larger: `set_text_underline` grew from 1,248 bytes and 312
instructions to 1,392 bytes and 352 instructions. The refusal branch gained a
call and temporary result storage, and the successful operation path changed
from the inlined vector push sequence to a longer capacity/copy sequence. The
refusal-contract benchmark binary grew from 4,139,552 bytes to 4,165,728
bytes. This experiment was also reverted; the native timing comparison remains
open for a future guarded host window.

The fixed operation and reason strings, `RefusalKind`, display text, error
classification, capability, fidelity and allocation behavior therefore stay
unchanged. The 0359b refusal-route ratio remains an owner follow-up for a
future optimization with a concrete aligned representation; this record does
not qualify it as noise.

The compact result below keeps only reproducible layout and assembly metadata;
the binaries and disassemblies remain outside the repository.
