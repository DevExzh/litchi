# Embedded fonts in binary PowerPoint

`litchi-ppt` treats embedded font programs as inert document data. Reading,
validating, preserving, or replacing a font never installs it, loads it into a
rendering engine, executes it, or performs system font discovery.

## Feature selection

The base crate can inspect and preserve the MS-PPT font records without font
tooling:

```toml
litchi-ppt = { version = "0", default-features = false }
```

Enable `fonts` for bounded EOT 1.0 validation and preparation adapters. More
expensive capabilities remain separate:

```toml
litchi-ppt = { version = "0", default-features = false, features = ["fonts"] }
```

| Feature | Adds |
|---------|------|
| `fonts` | Inert EOT validation and explicit-license embedding adapters |
| `font-subset` | Font subsetting support in `litchi-fonts` |
| `font-discovery` | Explicit system-font discovery in `litchi-fonts` |
| `automatic-fonts` | Discovery and subsetting convenience workflow |

None of these features enables macros, VBA, ActiveX, controls, actions, DDE, or
embedded-code execution.

## Format contract

The base collection is owned by the live document's environment/text-info
records. A `FontCollectionEntry` consists of one required `FontEntityAtom`
followed by at most four `FontEmbedDataBlob` facets in this order:

| Facet instance | Style |
|----------------|-------|
| `0` | Plain |
| `1` | Bold |
| `2` | Italic |
| `3` | Bold italic |

`FontIndexRef` and `FontIndexRef10` are zero-based collection ordinals. They do
not address `FontEntityAtom.recInstance`; duplicate entity instances therefore
do not change reference resolution.

Typeface names occupy exactly 32 UTF-16 code units on the wire. A shorter name
has a null terminator. A valid 32-unit name may consume the complete field and
have no terminator. Malformed UTF-16 is rejected rather than replaced.

The PowerPoint 10 collection and `FontEmbedFlags10Atom` live in the document's
single `___PPT10` programmable-tag payload. Bits 0 and 1 of the flags word are
typed; higher undefined bits are retained as inert data and ignored when
interpreting the two defined settings.

## Resource and transaction policy

Use `Record::parse_with_limits`, package `*_with_limits` constructors, or
`OpenOptions::record_limits` for untrusted input. Limits cover the input stream,
individual record and payload sizes, record count, nesting depth, and aggregate
copied payload bytes. EOT validation has independent input, output, font,
table-count, and name-byte limits.

Exact no-op package transactions must return the original compound-file bytes.
A changed embedded-font transaction must refuse signed or encrypted sources;
it must not silently invalidate a signature, decrypt/re-encrypt a package, or
publish a partial edit. Existing font order is stable: replacement and append
operations cannot make ordinal text references point at another font. Removal
or reordering requires proof that all affected references are understood and
updated; otherwise the operation is refused.

EOT license decisions use the embedded OpenType `OS/2.fsType` value as the
authority. Resolver metadata must match it. Restricted, bitmap-only, ambiguous,
or reserved licenses are rejected, and preview/print permission is never
promoted to editable use.

## Independent conformance vectors

`tests/ppt_embedded_fonts_spec.rs` manually encodes record headers and payloads
without calling the production writer. It covers:

- base and PowerPoint 10 collections;
- all four facet instances;
- a complete 32-unit non-null face name;
- malformed UTF-16 and malformed entity length;
- ordinal lookup with duplicate entity instances;
- ignored PowerPoint 10 flag bits;
- record and EOT input limits;
- malformed EOT size and name fields;
- exact no-op preservation for signed and encrypted packages;
- refusal of changed font publication for signed and encrypted packages.

The EOT vector contains a minimal inert SFNT with an `OS/2` table. Tests never
load, discover, rasterize, or execute the font.
