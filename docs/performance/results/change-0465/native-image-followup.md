# Native image follow-up (diagnostic inspection)

This document records a read-only diagnostic inspection of the retained
LibreOffice round trip from change 0464. It is static package evidence only.
No package-level image oracle has been implemented, no production API has been
changed, and this inspection adds no native Office image-acceptance claim.

## Exact artifacts

The native round-trip input was the retained formal capture:

```text
path: docs/performance/results/change-0464/captures/R1/normal-bytes/output.pptx
bytes: 55891
sha256: 550e8d8e8082f43f705f0c0a8bac0f2dc4fde09fb4c0d1989747abb3ca6ffbfb
```

The native round-trip output was:

```text
path: docs/performance/results/change-0464/native-roundtrip-r2/resaved.pptx
bytes: 39830
sha256: d2ca3109448f5a0797f00561beff4eafa84616f7dbf897ff2456667fbf555915
```

The retained supplemental inventory reports image inventory as unavailable on
all three post-round-trip slides. Each error is typed `unsafe_edit` with:

```text
operation: source-backed picture inventory
reason: markup-compatibility elements and attributes are refused
```

## Static package observations

The input package contains three direct slide pictures and three physical PNG
parts:

```text
ppt/media/image1.png          11658 bytes
ppt/media/image1-copy1.png   11658 bytes
ppt/media/image1-copy2.png   11658 bytes
```

Each has the same payload SHA-256:

```text
4d1c50656f3cfc644b2dec9c7af94965b46b4adb27f84b09a285c3470801562b
```

The post-round-trip package has one `ppt/media/image1.png` part with that same
size and SHA. Each of its three slides has one direct `p:pic`, one
`a:blip r:embed="rId1"`, an internal image relationship, and `image/png`
content type. Slides 2 and 3 changed their relationship targets from the
`image1-copy*` names to `image1`, which is package deduplication; the payload
identity is unchanged.

The post package adds one markup-compatibility branch to each slide after
`p:cSld`:

```xml
<mc:AlternateContent>
  <mc:Choice Requires="p14">
    <p:transition spd="slow" p14:dur="2000"/>
  </mc:Choice>
  <mc:Fallback>
    <p:transition spd="slow"/>
  </mc:Fallback>
</mc:AlternateContent>
```

The branch is outside every picture subtree. No `p:pic` is inside an
`mc:Choice` or `mc:Fallback`, and there is no `mc:Ignorable` attribute in the
slide XML. A one-off independent Python ZIP/XML inspection confirmed the
three image counts, internal targets, `image/png` content types, and equal
payload identities above. That inspection was not retained as an oracle or
verification artifact.

## Current production refusal

`SourceSlide::query_images` invokes
`reject_picture_markup_compatibility` on the complete selected slide before
projecting picture descriptors
([source.rs](../../../../crates/litchi-pptx/src/presentation/source.rs:2363)).
The scanner returns `Error::UnsafeEdit` as soon as it sees an
markup-compatibility element or attribute
([source.rs](../../../../crates/litchi-pptx/src/presentation/source.rs:2913)).
The exact operation and reason are the typed values recorded above. This is a
deliberate fail-closed source-backed boundary: selecting an MCE branch would
otherwise project a lossy picture view.

## Future bounded oracle plan

A separate, read-only package oracle could establish image facts for this
artifact without weakening the production boundary. It would need to reject
closed-world violations rather than silently select a branch:

- any `p:pic` outside a direct `p:spTree`, or inside or below an MCE branch;
- MCE content inside a picture subtree;
- malformed or ambiguous `a:blip` relationship attributes;
- external, missing, non-image, or unsupported relationship targets;
- missing media parts, missing content types, or unsupported image content
  types.

Positive coverage would compare the input and post-round-trip packages by
ordered direct-picture count, resolved internal target and content type, and
media payload byte count and SHA-256. Negative coverage should mutate a
picture into an MCE branch, remove or retarget its relationship, mark it
external, change its content type, and mutate its media bytes; each case must
fail closed. Such an oracle could report package-level payload identity only.
It must not convert that result into rendering, application acceptance, or a
source-backed `images()` success claim.
