# ADR 0022: Typed PPTX embedded-font ownership

- Status: Accepted
- Date: 2026-08-03

## Context

The OOXML migration host owned PresentationML embedded-font XML, relationship
validation, inert font resources, and package CRUD in one long module. Its
ordinary structs exposed relationship IDs, part names, content-type strings,
optional resources, and freely mutable field bags. Loading cloned a complete
font program into every face that referenced its part, so many references to one
large physical resource could amplify memory far beyond the unique-part budget.
Name lookup, duplicate validation, and reorder used different case algorithms;
the reorder finished with a panic. Publication invalidated signatures before
fallible cleanup and rewrote exact no-ops.

The feature-gated automatic authoring path was also physically wrong for PPTX.
It reused Word's publisher, XOR-obfuscated the font, emitted `.odttf` with the
Word-only `obfuscatedFont` content type, and ignored the generated `fontKey`.
The checked-in `[MS-OE376]` section 2.1.34 and `[MS-OI29500]` section 2.1.32
state that only WordprocessingML may reference that content type. They also
record PowerPoint's use of the Embedded OpenType/MicroType container under
`application/x-fontdata`, despite the base standard's additional raw
`application/x-font-ttf` profile. `[MS-OE376]` section 2.1.1134 and
`[MS-OI29500]` section 2.1.1100 add PowerPoint's unique-typeface requirement.

The ECMA schema permits zero through four ordered face references. It defines
`pitchFamily` as a closed 18-value domain and `charset` as signed `xsd:byte`.
It does not require distinct relationship IDs per face or a `/ppt/fonts/`
physical location. Strict and Transitional roots require their matching
relationship namespaces and relationship-type families.

## Decision

`litchi-pptx::font` is the sole owner of the bounded embedded-font grammar,
semantic values, and OPC graph service. The migration host deletes its duplicate
owner and long compatibility exports. Its ordinary package and presentation
facades expose `fonts`, `put_fonts`, and `remove_fonts`.

The contextual public vocabulary is `Fonts`, `Font`, `Face`, `Data`, `Style`,
`Format`, `Pitch`, `Family`, `PitchFamily`, `Panose`, `Charset`, `License`,
`Permission`, `Restrictions`, `Key`, and `Conformance`. Relationship IDs,
target part names, content-type strings, and physical provenance remain private. `Font::new`
represents the schema-valid face-less descriptor; `from_face`, `with`, and
`put` are the concise checked authoring path. A face always carries typed
`Style` and an inert `Data` value. A detached font may add, replace, query, and
remove faces without an option bag or native discriminator.

`PitchFamily` composes closed `Pitch` and `Family` enums, making all and only the
18 wire values constructible. `Charset` retains the complete signed-byte wire
domain behind a typed value, and `Panose` retains exactly ten bytes. `License`
validates the compact OpenType `fsType` word and exposes one
mutually exclusive `Permission` plus bitflags `Restrictions`; its fields cannot
be forged. PPTX exposes no obfuscation function.

`Data::powerpoint(Vec<u8>)` adopts and validates an EOT payload for PowerPoint's
`application/x-fontdata` profile. `Data::standard` validates and explicitly
retains the base-standard raw `application/x-font-ttf` profile. Loaded producer
payloads have a separate bounded preservation path so unknown but existing
containers can round-trip without being misrepresented as fresh validated
authoring. Internally each payload is
an `Arc<Vec<u8>>`: cloning `Data` or loading several face references to one part
shares the allocation. The facade returns slices and contextual values rather
than `Arc<RwLock<...>>` type noise. Fresh physical names are collision-checked
and format-aware; loaded noncanonical internal target locations are preserved.

One documented NFD/default-case-fold/NFD key drives ordinary typeface lookup
and every detached mutation. This is a Litchi selector policy, not a claim that
ECMA specifies Unicode normalization. Producer spelling is retained. A
malformed producer collection with caseless-equivalent names remains readable;
semantic selection reports typed ambiguity while checked numeric `Key::Index`
selection remains available for repair. New conflicting names are rejected.
Reorder applies a complete checked permutation by moving values and contains no
unchecked index or panic path.

Package `load`, consuming `put`, and `remove` validate all presentation,
slideshow, and template main content types, including macro-enabled variants.
They require one internal Presentation-owner relationship family matching the
root conformance, matching `r:id` namespace, supported font content type,
relationship-free font parts, bounded resources, and coherent inbound
ownership. Repeated face references may share one relationship ID and target.
Aggregate payload limits count unique physical targets or shared allocations,
not references. Fresh writes set `p:presentation@embedTrueTypeFonts="1"`;
complete removal sets it false. Unknown presentation XML remains byte-preserved
outside the focused list and root attribute.

An exact loaded no-op returns `false` before serialization or unsigning. A real
change validates XML, identities, limits, collisions, shared-resource
conflicts, and the staged round trip first. It then mutates an OPC candidate
whose built-in payloads remain `Arc`-shared, invalidates signatures on that
candidate, and publishes by final assignment. Any error leaves the caller's
package unchanged. The owner never loads, installs, renders, shapes, executes,
or otherwise activates a font program.

Automatic font discovery/subsetting remains optional in the migration host and
does not become a canonical PPTX dependency. The shared helper now prepares
deterministically ordered, owned, un-obfuscated programs with no package or
format knowledge. It surfaces `OS/2.fsType`, rejects restricted and bitmap-only
embedding, honors no-subsetting, and reports discovery/subsetting failures
instead of silently skipping or changing policy. The DOCX adapter alone applies
Word obfuscation; the PPTX adapter wraps one standalone OpenType face in an
uncompressed EOT 1.0 container, builds typed semantic values, and performs one
bulk consuming `font::put`. Proprietary MicroType compression is preserved when
loaded but is never synthesized. The umbrella `fonts` feature forwards to the
OOXML host feature.

The shared discovery request is `Request { family, style }`, where `style` is
the closed four-face enum. `Glyphs` privately wraps its roaring bitmap and can
only be populated from Rust `char`, so surrogates and values above `U+10FFFF`
cannot enter through the public collector. Bold, italic, and bold-italic runs
select their corresponding system face and publish the corresponding DOCX or
PPTX face instead of being mislabeled Regular. Cmap failures propagate; request
and result allocations are fallible and bounded at the adapter seam. A
font-kit memory handle transfers its original `Vec` allocation when its `Arc`
is unique and clones only when another owner remains; path-backed handles keep
their direct owned file read.

OPC save configuration uses `FontEmbedding::{None, Full, Subset}` rather than
two booleans, making `embed = false, subset = true` unrepresentable. Automatic
embedding rejects an opened or partially preserved document whose complete
glyph inventory is unavailable. Empty therefore means a verified empty mutable
model, not a missing scanner. Full embedding currently rejects font collections
until lossless selected-face extraction is implemented; it never publishes a
TTC while silently discarding its face index.

Word's obfuscation key is the compact `FontKey([u8; 16])`; parsing and canonical
braced formatting occur only at the XML boundary. Strict Word charsets with no
legal IANA lexical form are omitted instead of serialized as an invalid value.
The settings patcher is explicitly bounded and accepts UTF-8 only until an
encoding-preserving writer exists, so UTF-16 source XML fails safely rather
than receiving UTF-8 byte splices.

The uncompressed EOT writer follows the version-1 field layout, including the
absence of `Padding5` after `FullName`. It builds only the bounded header, then
reuses the source font `Vec` and performs the one unavoidable in-place shift
required by a contiguous container. It does not retain a second full-font
allocation.

## Consequences

- Embedded-font CRUD has one owner and a short semantic facade; ordinary users
  never manipulate an `rId`, part path, or MIME string.
- Invalid pitch/family combinations, raw charset strings, Word-only PPTX
  obfuscation, and face values without data are not authorable through the
  ordinary API.
- Raw Unicode integers, raw Word font-key strings, contradictory font-save
  booleans, and style-less automatic face requests are likewise absent from the
  ordinary API.
- Shared program bytes remain one allocation across faces, package snapshots,
  and publication. This is a structural memory property, not a throughput or
  cache-performance claim; ADR 0005 still requires measurements.
- A uniquely owned in-memory system-font handle crosses the loader boundary by
  move. Shared handles retain copy-on-ownership-conflict semantics rather than
  invalidating another owner.
- A valid face-less producer entry and repeated-rId graph remain representable.
  New authoring can choose the concise face-bearing constructor.
- `x-font-ttf` is supported for standards preservation, but native PowerPoint
  verification determines what can be claimed for current Microsoft Office.

## Verification

Verification covers Strict and Transitional XML; zero and four faces; all 18
pitch/family values; signed charset boundaries; PANOSE and XML rejection;
Unicode ambiguity with numeric repair; repeated relationship IDs and shared
targets; noncanonical part locations; all six main-part content types; external,
missing, wrong-dialect, wrong-content-type, outbound, orphan, shared, collision,
allocation, and aggregate-limit failures; exact no-op/signature preservation;
unknown XML preservation; failure atomicity; public downstream API compilation;
reference packages from LibreOffice and Apache POI; feature-on/off builds; and
warning-denied Clippy and rustdoc.

Desktop Microsoft PowerPoint for macOS opened
`target/office-verification/pptx-font-crud-generated.pptx` without repair or a
font-license warning, rendered the visible Boldonse text, and reported
`Boldonse` in the selected text box's font control. Computer Use changed `Test`
to `Test Test`, saved a separate copy, closed it, and reopened that copy without
repair. ZIP validation and the canonical reverse reader found the edited text,
one Regular Boldonse face, one inert `application/x-fontdata` part, and the
presentation-owned relationship. The observed copy retained the exact 36,187
byte EOT payload and SHA-256, although the verifier intentionally permits a
structurally valid Office normalization or re-subset. This certifies the tested
Transitional EOT CRUD artifact on that desktop build, not the automatic raw
OpenType-to-EOT wrapper, the standards-only `x-font-ttf` profile, Strict native
behavior, other Office versions, or performance.
