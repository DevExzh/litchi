//! Semantic family readers for flat OpenDocument XML documents.
//!
//! Flat documents (`.fodt`, `.fods`, `.fodp`, `.fodg`, `.fodc`, `.fodi`) carry
//! every part of a packaged document as top-level sections of one
//! `office:document` root. This module splits those sections into an
//! in-memory synthetic package and reopens it through the packaged family
//! readers, so flat files expose the same read-only semantic models
//! (`Document`, `Spreadsheet`, `Presentation`, `DrawingDocument`,
//! `ChartDocument`, `ImageDocument`) without duplicating their parsers.
//!
//! The wrappers are read-only: the original flat bytes are kept verbatim and
//! `save`/`to_bytes` always return them byte-identically. All parsing stays
//! inert — no scripts, formulas, or links are evaluated — and inherits the
//! bounded parsing limits of the packaged readers.

use crate::constants;
use crate::core::PackageWriter;
use crate::generic::{FlatOpenDocument, OpenDocumentFamily};
use litchi_core::{Error, Metadata, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::io::Read;
use std::path::Path;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

/// Candidate prefixes for the synthesized part roots, in preference order.
const WRAPPER_PREFIXES: [&str; 3] = ["office", "officeflat", "flatwrapper"];

/// Capacity slack for the XML declaration and synthesized part root tags.
const PART_WRAPPER_OVERHEAD: usize = 64;

/// Maximum wrapper-prefix candidates tried before rejecting the document.
const MAX_WRAPPER_PREFIX_ATTEMPTS: usize = 64;

/// Byte-range slices of one flat document, grouped by package part.
struct FlatSections<'a> {
    /// Serialized `xmlns` declarations copied from the root element.
    namespace_decls: String,
    /// Prefix bound to the office namespace for the synthesized part roots.
    wrapper_prefix: String,
    /// Extra `xmlns` declaration needed when the root did not bind the prefix.
    wrapper_decl: Option<String>,
    /// Sections forming `content.xml`, in document order.
    content: Vec<&'a str>,
    /// Sections forming `styles.xml`, in document order.
    styles: Vec<&'a str>,
    /// Sections forming `meta.xml`, in document order.
    meta: Vec<&'a str>,
    /// Sections forming `settings.xml`, in document order.
    settings: Vec<&'a str>,
}

impl FlatSections<'_> {
    /// Render one synthesized part as a standalone XML document.
    fn render(&self, root_local_name: &str, sections: &[&str]) -> String {
        let mut part = String::with_capacity(
            sections.iter().map(|section| section.len()).sum::<usize>()
                + self.namespace_decls.len()
                + root_local_name.len()
                + self.wrapper_prefix.len()
                + PART_WRAPPER_OVERHEAD,
        );
        part.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<");
        part.push_str(&self.wrapper_prefix);
        part.push(':');
        part.push_str(root_local_name);
        part.push_str(&self.namespace_decls);
        if let Some(decl) = &self.wrapper_decl {
            part.push(' ');
            part.push_str(decl);
        }
        part.push('>');
        for section in sections {
            part.push_str(section);
        }
        part.push_str("</");
        part.push_str(&self.wrapper_prefix);
        part.push(':');
        part.push_str(root_local_name);
        part.push_str(">\n");
        part
    }
}

/// Split the top-level sections of a flat document into package parts.
///
/// Sections are borrowed slices of the flat XML; allocation only happens when
/// the synthesized part documents are rendered.
fn split_flat_sections(xml: &str) -> Result<FlatSections<'_>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut namespace_decls: Vec<(String, String)> = Vec::new();
    let mut root_seen = false;
    let mut content = Vec::new();
    let mut styles = Vec::new();
    let mut meta = Vec::new();
    let mut settings = Vec::new();
    // Top-level section whose start tag was seen but whose end tag is still
    // pending: (route, byte offset of the start tag).
    let mut pending_section: Option<(SectionRoute, usize)> = None;

    loop {
        let event_start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid flat OpenDocument XML: {error}"))
            })?;
        match event {
            Event::Start(ref element) => {
                if depth == 0 {
                    if root_seen {
                        return Err(Error::InvalidFormat(
                            "flat OpenDocument must contain one root element".to_string(),
                        ));
                    }
                    root_seen = true;
                    namespace_decls = root_namespace_decls(element)?;
                } else if depth == 1
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
                {
                    pending_section = Some((
                        SectionRoute::of(element.local_name().as_ref()),
                        event_start,
                    ));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("flat OpenDocument nesting overflow".to_string())
                })?;
            },
            Event::Empty(ref element)
                if depth == 1
                    && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE) =>
            {
                let event_end = reader.buffer_position() as usize;
                SectionRoute::of(element.local_name().as_ref()).push(
                    &xml[event_start..event_end],
                    &mut content,
                    &mut styles,
                    &mut meta,
                    &mut settings,
                );
            },
            Event::End(_) => {
                if depth == 2
                    && let Some((route, section_start)) = pending_section.take()
                {
                    let event_end = reader.buffer_position() as usize;
                    route.push(
                        &xml[section_start..event_end],
                        &mut content,
                        &mut styles,
                        &mut meta,
                        &mut settings,
                    );
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("unexpected flat OpenDocument closing tag".to_string())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen {
        return Err(Error::InvalidFormat(
            "flat OpenDocument has no root element".to_string(),
        ));
    }

    let (wrapper_prefix, wrapper_decl) = wrapper_prefix(&namespace_decls)?;
    let mut serialized = String::new();
    for (key, value) in &namespace_decls {
        serialized.push(' ');
        serialized.push_str(key);
        serialized.push_str("=\"");
        serialized.push_str(value);
        serialized.push('"');
    }

    Ok(FlatSections {
        namespace_decls: serialized,
        wrapper_prefix,
        wrapper_decl,
        content,
        styles,
        meta,
        settings,
    })
}

/// Package parts that receive one top-level flat section.
///
/// `office:font-face-decls` and `office:automatic-styles` appear in both
/// `content.xml` and `styles.xml` of a packaged document, matching the ODF
/// content models of `office:document-content` and `office:document-styles`.
#[derive(Clone, Copy)]
enum SectionRoute {
    Content,
    Styles,
    ContentAndStyles,
    Meta,
    Settings,
}

impl SectionRoute {
    fn of(local_name: &[u8]) -> Self {
        match local_name {
            b"meta" => Self::Meta,
            b"settings" => Self::Settings,
            b"styles" | b"master-styles" => Self::Styles,
            b"font-face-decls" | b"automatic-styles" => Self::ContentAndStyles,
            _ => Self::Content,
        }
    }

    fn push<'a>(
        self,
        section: &'a str,
        content: &mut Vec<&'a str>,
        styles: &mut Vec<&'a str>,
        meta: &mut Vec<&'a str>,
        settings: &mut Vec<&'a str>,
    ) {
        match self {
            Self::Content => content.push(section),
            Self::Styles => styles.push(section),
            Self::ContentAndStyles => {
                content.push(section);
                styles.push(section);
            },
            Self::Meta => meta.push(section),
            Self::Settings => settings.push(section),
        }
    }
}

/// Collect the `xmlns` declarations of the flat root element verbatim.
fn root_namespace_decls(root: &quick_xml::events::BytesStart<'_>) -> Result<Vec<(String, String)>> {
    let mut decls = Vec::new();
    for attribute in root.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid flat OpenDocument root attribute: {error}"))
        })?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            let key = std::str::from_utf8(key).map_err(|_| {
                Error::InvalidFormat("invalid flat OpenDocument namespace name".to_string())
            })?;
            let value = std::str::from_utf8(attribute.value.as_ref()).map_err(|_| {
                Error::InvalidFormat("invalid flat OpenDocument namespace value".to_string())
            })?;
            decls.push((key.to_string(), value.to_string()));
        }
    }
    Ok(decls)
}

/// Choose a prefix for the synthesized part roots, adding a declaration when
/// the flat root did not already bind that prefix to the office namespace.
///
/// When every preferred prefix is bound to a foreign namespace, numbered
/// `flatwrapperN` candidates are tried before giving up with an error; a
/// crafted document must never reach a panic here.
fn wrapper_prefix(decls: &[(String, String)]) -> Result<(String, Option<String>)> {
    let office_uri = std::str::from_utf8(OFFICE_NAMESPACE)
        .expect("office namespace URI is valid UTF-8");
    let candidates = WRAPPER_PREFIXES
        .iter()
        .map(|prefix| (*prefix).to_string())
        .chain((WRAPPER_PREFIXES.len()..).map(|index| format!("flatwrapper{index}")));
    for prefix in candidates.take(MAX_WRAPPER_PREFIX_ATTEMPTS) {
        let key = format!("xmlns:{prefix}");
        match decls.iter().find(|(name, _)| *name == key) {
            Some((_, uri)) if uri.as_bytes() == OFFICE_NAMESPACE => return Ok((prefix, None)),
            Some(_) => continue,
            None => return Ok((prefix, Some(format!("{key}=\"{office_uri}\"")))),
        }
    }
    Err(Error::InvalidFormat(
        "flat OpenDocument root binds every synthesized wrapper prefix candidate".to_string(),
    ))
}

/// Synthesize an in-memory package from a validated flat document.
fn synthesize_package(flat: &FlatOpenDocument) -> Result<Vec<u8>> {
    let sections = split_flat_sections(flat.xml())?;
    let mut writer = PackageWriter::new();
    writer.set_mimetype(flat.mimetype())?;
    writer.add_file(
        constants::ODF_CONTENT,
        sections
            .render("document-content", &sections.content)
            .as_bytes(),
    )?;
    if !sections.styles.is_empty() {
        writer.add_file(
            constants::ODF_STYLES,
            sections.render("document-styles", &sections.styles).as_bytes(),
        )?;
    }
    if !sections.meta.is_empty() {
        writer.add_file(
            constants::ODF_META,
            sections.render("document-meta", &sections.meta).as_bytes(),
        )?;
    }
    if !sections.settings.is_empty() {
        writer.add_file(
            constants::ODF_SETTINGS,
            sections
                .render("document-settings", &sections.settings)
                .as_bytes(),
        )?;
    }
    writer.finish_to_bytes()
}

/// Open a validated flat document of `expected_family` through its packaged
/// family reader.
fn open_packaged<R>(
    bytes: Vec<u8>,
    expected_family: OpenDocumentFamily,
    extension: &'static str,
    reader: impl FnOnce(Vec<u8>) -> Result<R>,
) -> Result<(FlatOpenDocument, R)> {
    let flat = FlatOpenDocument::from_bytes(bytes)?;
    if flat.family() != expected_family {
        return Err(Error::InvalidFormat(format!(
            "not a flat OpenDocument {extension} file: MIME type is '{}'",
            flat.mimetype()
        )));
    }
    let document = reader(synthesize_package(&flat)?)?;
    Ok((flat, document))
}

macro_rules! flat_family_wrapper {
    (
        $(#[$meta:meta])*
        pub struct $name:ident {
            inner: $inner:ty,
            accessor: $accessor:ident,
            mut_accessor: $mut_accessor:ident,
            family: $family:expr,
            extension: $extension:literal,
        }
    ) => {
        $(#[$meta])*
        pub struct $name {
            flat: FlatOpenDocument,
            document: $inner,
        }

        impl $name {
            #[doc = concat!("Open and validate a flat `.", $extension, "` document from a path.")]
            pub fn open(path: impl AsRef<Path>) -> Result<Self> {
                let file = std::fs::File::open(path)?;
                Self::from_reader(file)
            }

            #[doc = concat!("Read and validate a flat `.", $extension, "` document stream.")]
            pub fn from_reader(mut reader: impl Read) -> Result<Self> {
                let mut bytes = Vec::new();
                reader.read_to_end(&mut bytes)?;
                Self::from_bytes(bytes)
            }

            #[doc = concat!("Validate a flat `.", $extension, "` document from owned bytes.")]
            pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
                let (flat, document) =
                    open_packaged(bytes, $family, $extension, <$inner>::from_bytes)?;
                Ok(Self { flat, document })
            }

            #[doc = concat!("Borrow the packaged semantic model.\n\nThe model is read-only through this reference; mutating or rebuilding it does not affect the stored flat bytes.")]
            pub fn $accessor(&self) -> &$inner {
                &self.document
            }

            #[doc = concat!("Borrow the packaged semantic model mutably.\n\nSeveral read APIs parse lazily and need `&mut self`. The wrapper stays read-only: `save` and `to_bytes` always return the original flat bytes.")]
            pub fn $mut_accessor(&mut self) -> &mut $inner {
                &mut self.document
            }

            /// Borrow the format-neutral flat wrapper for settings, forms,
            /// variable declarations, images, and embedded objects.
            pub fn flat_document(&self) -> &FlatOpenDocument {
                &self.flat
            }

            /// Extract common document metadata from the flat `office:meta`
            /// section, or an empty value when the section is absent.
            pub fn metadata(&self) -> Result<Metadata> {
                self.document.metadata()
            }

            /// Return the exact original flat bytes.
            pub fn as_bytes(&self) -> &[u8] {
                self.flat.as_bytes()
            }

            /// Clone the exact original flat bytes.
            pub fn to_bytes(&self) -> Vec<u8> {
                self.flat.to_bytes()
            }

            /// Consume this wrapper and return the exact original flat bytes.
            pub fn into_bytes(self) -> Vec<u8> {
                self.flat.into_bytes()
            }

            /// Save the flat document without reconstructing its XML; the
            /// output is byte-identical to the input.
            pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
                self.flat.save(path)
            }
        }
    };
}

flat_family_wrapper! {
    /// Read-only semantic reader for flat OpenDocument text (`.fodt`).
    ///
    /// Exposes the full packaged [`Document`](crate::Document) read model —
    /// text, paragraphs, tables, styles, metadata — while saving returns the
    /// original flat bytes exactly.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::FlatTextDocument;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let document = FlatTextDocument::open("notes.fodt")?;
    /// println!("{}", document.document().text()?);
    /// # Ok(())
    /// # }
    /// ```
    pub struct FlatTextDocument {
        inner: crate::odt::Document,
        accessor: document,
        mut_accessor: document_mut,
        family: OpenDocumentFamily::Text,
        extension: "fodt",
    }
}

flat_family_wrapper! {
    /// Read-only semantic reader for flat OpenDocument spreadsheets (`.fods`).
    ///
    /// Exposes the full packaged [`Spreadsheet`](crate::Spreadsheet) read
    /// model — sheets, rows, cells, values, named expressions, metadata —
    /// while saving returns the original flat bytes exactly.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::FlatSpreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = FlatSpreadsheet::open("budget.fods")?;
    /// println!("Sheets: {}", spreadsheet.spreadsheet_mut().sheet_count()?);
    /// # Ok(())
    /// # }
    /// ```
    pub struct FlatSpreadsheet {
        inner: crate::ods::Spreadsheet,
        accessor: spreadsheet,
        mut_accessor: spreadsheet_mut,
        family: OpenDocumentFamily::Spreadsheet,
        extension: "fods",
    }
}

flat_family_wrapper! {
    /// Read-only semantic reader for flat OpenDocument presentations (`.fodp`).
    ///
    /// Exposes the full packaged [`Presentation`](crate::Presentation) read
    /// model — slides, shapes, notes, metadata — while saving returns the
    /// original flat bytes exactly.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::FlatPresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let presentation = FlatPresentation::open("talk.fodp")?;
    /// for slide in presentation.presentation().slides()? {
    ///     println!("Slide {}: {}", slide.index() + 1, slide.text()?);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub struct FlatPresentation {
        inner: crate::odp::Presentation,
        accessor: presentation,
        mut_accessor: presentation_mut,
        family: OpenDocumentFamily::Presentation,
        extension: "fodp",
    }
}

flat_family_wrapper! {
    /// Read-only semantic reader for flat OpenDocument drawings (`.fodg`).
    ///
    /// Exposes the full packaged [`DrawingDocument`](crate::DrawingDocument)
    /// read model — pages, layers, shapes, metadata — while saving returns
    /// the original flat bytes exactly.
    pub struct FlatDrawingDocument {
        inner: crate::odg::DrawingDocument,
        accessor: drawing,
        mut_accessor: drawing_mut,
        family: OpenDocumentFamily::Drawing,
        extension: "fodg",
    }
}

flat_family_wrapper! {
    /// Read-only semantic reader for flat OpenDocument charts (`.fodc`).
    ///
    /// Exposes the full packaged [`ChartDocument`](crate::ChartDocument) read
    /// model — the chart element tree, titles, metadata — while saving
    /// returns the original flat bytes exactly.
    pub struct FlatChartDocument {
        inner: crate::odc::ChartDocument,
        accessor: chart,
        mut_accessor: chart_mut,
        family: OpenDocumentFamily::Chart,
        extension: "fodc",
    }
}

flat_family_wrapper! {
    /// Read-only semantic reader for flat OpenDocument images (`.fodi`).
    ///
    /// Exposes the full packaged [`ImageDocument`](crate::ImageDocument) read
    /// model — the image frame, inert payload access, metadata — while
    /// saving returns the original flat bytes exactly.
    pub struct FlatImageDocument {
        inner: crate::odi::ImageDocument,
        accessor: image,
        mut_accessor: image_mut,
        family: OpenDocumentFamily::Image,
        extension: "fodi",
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn flat_text(body: &str) -> String {
        format!(
            r#"<?xml version="1.0"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:mimetype="{}" office:version="1.3">{body}</office:document>"#,
            constants::ODF_TEXT
        )
    }

    #[test]
    fn split_routes_sections_to_packaged_parts() {
        let xml = flat_text(
            "<office:meta/><office:settings/><office:styles/>\
             <office:automatic-styles/><office:master-styles/>\
             <office:body><office:text/></office:body>",
        );
        let sections = split_flat_sections(&xml).unwrap();
        assert_eq!(sections.content.len(), 2); // automatic-styles + body
        assert_eq!(sections.styles.len(), 3); // styles + automatic-styles + master-styles
        assert_eq!(sections.meta.len(), 1);
        assert_eq!(sections.settings.len(), 1);

        let content = sections.render("document-content", &sections.content);
        assert!(content.starts_with("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<office:document-content"));
        assert!(content.contains("xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\""));
        assert!(content.ends_with("</office:document-content>\n"));
        assert!(content.contains("<office:body><office:text/></office:body>"));
    }

    #[test]
    fn wrapper_prefix_falls_back_when_office_prefix_is_rebound() {
        let xml = concat!(
            r#"<?xml version="1.0"?><o:document xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
            r#"xmlns:office="urn:example:foreign" o:mimetype="application/vnd.oasis.opendocument.text">"#,
            r#"<o:body><o:text/></o:body></o:document>"#,
        );
        let sections = split_flat_sections(xml).unwrap();
        assert_eq!(sections.wrapper_prefix, "officeflat");
        let content = sections.render("document-content", &sections.content);
        assert!(content.contains("<officeflat:document-content"));
        assert!(content.contains("xmlns:officeflat=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\""));
    }

    #[test]
    fn wrapper_prefix_generates_numbered_fallbacks_and_never_panics() {
        // All preferred prefixes bound to foreign namespaces.
        let decls: Vec<(String, String)> = WRAPPER_PREFIXES
            .iter()
            .map(|prefix| (format!("xmlns:{prefix}"), "urn:example:foreign".to_string()))
            .collect();
        let (prefix, decl) = wrapper_prefix(&decls).unwrap();
        assert_eq!(prefix, format!("flatwrapper{}", WRAPPER_PREFIXES.len()));
        assert!(decl.unwrap().starts_with(&format!("xmlns:{prefix}=")));

        // Every candidate bound: an error, not a panic.
        let decls: Vec<(String, String)> = (0..MAX_WRAPPER_PREFIX_ATTEMPTS + 4)
            .map(|index| (format!("xmlns:flatwrapper{index}"), "urn:example:foreign".to_string()))
            .chain(WRAPPER_PREFIXES.iter().map(|prefix| {
                (format!("xmlns:{prefix}"), "urn:example:foreign".to_string())
            }))
            .collect();
        assert!(wrapper_prefix(&decls).is_err());
    }

    #[test]
    fn flat_text_document_reads_semantics_and_round_trips() {
        let xml = flat_text(
            "<office:body><office:text><text:h>Title</text:h><text:p>Hello flat world</text:p></office:text></office:body>",
        );
        let document = FlatTextDocument::from_bytes(xml.clone().into_bytes()).unwrap();
        let text = document.document().text().unwrap();
        assert!(text.contains("Title"));
        assert!(text.contains("Hello flat world"));
        assert_eq!(document.to_bytes(), xml.as_bytes());
    }

    #[test]
    fn flat_text_document_rejects_other_families() {
        let xml = format!(
            r#"<?xml version="1.0"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="{}"><office:body><office:spreadsheet/></office:body></office:document>"#,
            constants::ODF_SPREADSHEET
        );
        assert!(FlatTextDocument::from_bytes(xml.into_bytes()).is_err());
    }
}
