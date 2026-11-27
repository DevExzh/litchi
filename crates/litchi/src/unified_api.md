# Opening a file

[`open`] is the entry point. It detects the format from the content in
one pass and returns a [`File`]: a flat enum with one variant per format,
plus [`File::Locked`] for an encrypted input and [`File::Signed`] for a
signed one. [`open_bytes`] does the same for an in-memory input, and
[`open_with`] and [`open_bytes_with`] take [`OpenOptions`] for limits and
policy. The common case needs no options and no format knowledge.

```
# #[cfg(feature = "docx")]
# {
let path = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/ooxml/docx/documentProperties.docx"
);
let file = litchi::open(path)?;
assert_eq!(file.format(), Some(litchi::FileFormat::Docx));
assert_eq!(file.family(), Some(litchi::Family::Document));

// Reads every format shares are methods on `File`.
assert_eq!(file.text()?.trim(), "Hello World!");
assert!(file.tolerance_report()?.is_clean());

// Family reads go through the view of the file's family.
let document = file.document()?;
assert_eq!(document.paragraph_count()?, 1);
# }
# Ok::<(), Box<dyn std::error::Error>>(())
```

A view refuses a file of another family, and names the format it found:

```
# #[cfg(all(feature = "docx", feature = "xlsx"))]
# {
# let path = concat!(
#     env!("CARGO_MANIFEST_DIR"),
#     "/../../test-data/ooxml/docx/documentProperties.docx"
# );
let file = litchi::open(path)?;
let error = file.workbook().unwrap_err();
assert_eq!(error.code(), litchi::Code::UnsupportedFormat);
assert_eq!(error.detected(), Some(litchi::FileFormat::Docx));
# }
# Ok::<(), Box<dyn std::error::Error>>(())
```

Format-specific reads live on the handle a variant carries;
[`format`](mod@format) documents each handle:

```
# #[cfg(feature = "pptx")]
# {
use litchi::File;

# let path = concat!(
#     env!("CARGO_MANIFEST_DIR"),
#     "/../../test-data/ooxml/pptx/SmartArt.pptx"
# );
let file = litchi::open(path)?;
if let File::Pptx(pptx) = &file {
    let (width, height) = pptx.slide_size()?;
    assert_eq!((width, height), (12_192_000, 6_858_000)); // EMU, 16:9
    assert_eq!(pptx.slide_catalog()?.len(), 1);
}
# }
# Ok::<(), Box<dyn std::error::Error>>(())
```

With the `markdown` feature a document or presentation converts to
Markdown through `litchi::markdown::ToMarkdown`:

```
# #[cfg(all(feature = "docx", feature = "markdown"))]
# {
use litchi::markdown::ToMarkdown;

# let path = concat!(
#     env!("CARGO_MANIFEST_DIR"),
#     "/../../test-data/ooxml/docx/documentProperties.docx"
# );
let file = litchi::open(path)?;
assert_eq!(file.to_markdown()?, "Hello World\\!\n\n");
# }
# Ok::<(), Box<dyn std::error::Error>>(())
```

# Encrypted and signed files

An encrypted input opens as [`File::Locked`]. Its reads refuse with
[`Code::DocumentProtected`] and a suggestion, until `Locked::unlock` (the
`encryption` feature) consumes the handle with the password. The password
never passes through [`OpenOptions`].

```
# #[cfg(all(feature = "doc", feature = "encryption"))]
# {
use litchi::{Code, EncryptionKind, File, FileFormat, Password};

# let path = concat!(
#     env!("CARGO_MANIFEST_DIR"),
#     "/../../test-data/poi/test-data/document/password_password_cryptoapi.doc"
# );
let file = litchi::open(path)?;
let error = file.text().unwrap_err();
assert_eq!(error.code(), Code::DocumentProtected);
assert!(error.suggestion().is_some_and(|hint| hint.contains("Locked::unlock")));

let file = match file {
    File::Locked(locked) => {
        assert_eq!(locked.format(), Some(FileFormat::Doc));
        assert_eq!(locked.encryption(), EncryptionKind::CryptoApi);
        locked.unlock(Password::new("password".to_owned()))?
    },
    file => file,
};
assert!(!file.text()?.is_empty());
# }
# Ok::<(), Box<dyn std::error::Error>>(())
```

A signed OOXML or legacy binary input opens as [`File::Signed`]. It reads
like the file it wraps, and nothing is verified at open. Editing a signed
file requires consuming the signature first with [`Signed::unsign`], so
that no edit strips a signature silently.

```
# #[cfg(feature = "docx")]
# {
use litchi::{File, SignatureOrigin};

# let path = concat!(
#     env!("CARGO_MANIFEST_DIR"),
#     "/../../test-data/poi/test-data/xmldsign/hello-world-signed.docx"
# );
let file = litchi::open(path)?;
assert!(file.is_signed());
assert_eq!(file.text()?.trim(), "Hello world");
if let File::Signed(signed) = file {
    assert_eq!(signed.origin(), SignatureOrigin::Opc);
    let unsigned: File = signed.unsign();
    assert!(!unsigned.is_signed());
}
# }
# Ok::<(), Box<dyn std::error::Error>>(())
```

# Errors

Every failure of the unified API is one [`Error`] with a stable [`Code`]
(OfficeCLI's code list plus Litchi's `unsupported_capability`,
`limit_exceeded` and `cancelled`), a message, an optional suggestion
([`Error::suggestion`]), the format involved and the typed error of the
format crate that failed ([`Error::leaf`], also its
[`source`](std::error::Error::source)).

```
# #[cfg(feature = "docx")]
# {
let error = litchi::open_bytes(b"plain text".to_vec()).unwrap_err();
assert_eq!(error.code(), litchi::Code::UnsupportedFormat);
assert_eq!(error.code().as_str(), "unsupported_format");
assert_eq!(error.detected(), None);
# }
```

# Apple iWork

Pages, Keynote and Numbers packages open through [`open`] as
`File::Pages`, `File::Keynote` and `File::Numbers` (the `pages`, `keynote`
and `numbers` features): paths, queries, views, the leaves' focused writes
and atomic saves, like every other format. An app-authored package
*directory* opens the same way, as its single-file form, and `save()`
writes it back as a directory. The `iwork` feature enables all three
formats.

# Unified operations

Every format reads and writes through [`File`]: DOC, DOCX, RTF, ODT, OTH
and ODM documents; PPT, PPTX and ODP presentations; XLS, XLSX, XLSB and
ODS workbooks; the ODG, ODC, ODI, ODB and ODF formula objects; and Pages,
Keynote and Numbers. Reads are paths ([`File::get`]), queries
([`File::query`]) and the family views. Writes are immutable: `set`,
`add`, `remove`, `move_to` and `swap` each return a new [`File`] and leave
the input unchanged; [`File::edit`] stages several of them as one
transaction whose patch inverts and re-applies exactly; [`batch`] runs a
list of [`BatchItem`]s. A snapshot publishes through `save`, `save_as`,
`to_bytes` and `write_to`, and [`create`] starts a new file.

```
# #[cfg(feature = "docx")]
# {
use litchi::{FileFormat, NewNode, props};

let file = litchi::create(FileFormat::Docx)?;
let added = file.add("/body", NewNode::new("paragraph").prop("text", "Hello"))?;
let file = added.file.set(&added.path, props! { "text" => "Hello, world" })?;
assert_eq!(file.text()?.trim(), "Hello, world");
# }
# Ok::<(), Box<dyn std::error::Error>>(())
```

Not every element supports every verb in every format. [`File::capabilities`]
lists the verbs and properties the current handle supports and the typed
gaps it refuses; an unsupported request fails with
[`Code::UnsupportedCapability`], never with a silent partial write.

# Low-level modules

Each format crate is re-exported under its own module for format-specific
reading and writing: `doc`, `ppt` and `xls` for the legacy binary formats,
`docx`, `pptx`, `xlsx` and `xlsb` for the OOXML owners, `odt`, `ods` and
`odp` for OpenDocument, `pages`, `keynote` and `numbers` for Apple iWork,
and `rtf`. [`common`] holds the shared vocabulary
([`common::ShapeType`], [`common::RGBColor`], [`common::Length`], ...).
Most users need only [`open`] and [`create`], and reach for these modules
when they need a format-specific feature the unified API does not cover.
