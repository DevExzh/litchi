//! Flat (single-XML) `OpenDocument` documents as package parts.
//!
//! A flat document (`.fodt`, `.fods`, `.fodp`, `.fodg`, ...) is one
//! `office:document` element whose direct children are the parts a package
//! stores separately. [`Parts::split`] partitions the root's children into
//! `content.xml` (scripts, font faces, automatic styles, body), `styles.xml`
//! (font faces, styles, automatic styles, master styles), `meta.xml` and
//! `settings.xml`, each wrapped in its package root element with the flat
//! root's namespace declarations and version. [`Parts::to_bytes`] publishes
//! them as an in-memory package, so a family's packaged reader (and editor)
//! can work on a flat document without a second engine.
//!
//! Every child's bytes are copied verbatim; nested namespace declarations,
//! text and inline binary data are kept. Only the root start tag is
//! rewritten, dropping its `office:mimetype` attribute.
//!
//! [`Parts::merge`] maps an edit of those parts back: every top-level child
//! a part's edit changed replaces that child's bytes in the flat document,
//! and every other byte of the flat document is kept. An edit that changes
//! a part's root, adds or removes top-level children, or changes one shared
//! child (font faces, automatic styles) differently in two parts is refused.

use std::io;
use std::ops::Range;
use std::sync::Arc;

use quick_xml::Reader;
use quick_xml::events::Event;
use sha2::{Digest, Sha256};

use crate::core::{PackageWriter, StreamingArchiveLimits};
use crate::{
    Error, LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits, SourceLimits, XmlStreamLimits,
};
use litchi_core::{Error as CoreError, Resource, ResourceLimit};

type Result<T> = std::result::Result<T, Error>;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

/// The top-level children each part holds, in order.
const CONTENT: &[&[u8]] = &[b"scripts", b"font-face-decls", b"automatic-styles", b"body"];
const STYLES: &[&[u8]] = &[
    b"font-face-decls",
    b"styles",
    b"automatic-styles",
    b"master-styles",
];
const META: &[&[u8]] = &[b"meta"];
const SETTINGS: &[&[u8]] = &[b"settings"];

/// The validated policy required by the raw flat-document helpers.
///
/// A flat source has three policy domains. `SourceLimits` admits the original
/// XML buffer (and an independently supplied package), `SemanticLimits` bounds
/// the XML projection, nesting and publication, and `XmlStreamLimits` bounds
/// XML member parsing when the projection is handed to a packaged owner.
/// Keeping this focused carrier separate from [`OpenOptions`] prevents a raw
/// parser from silently selecting a default owner profile. Owners construct
/// it from the options they retained at ingress.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    source: SourceLimits,
    semantic: SemanticLimits,
    xml: XmlStreamLimits,
}

impl Limits {
    /// Build flat limits from already validated physical, semantic and XML
    /// profiles.
    #[must_use]
    pub const fn new(source: SourceLimits, semantic: SemanticLimits, xml: XmlStreamLimits) -> Self {
        Self {
            source,
            semantic,
            xml,
        }
    }

    /// Build flat limits from an already validated owner options envelope.
    #[must_use]
    pub const fn from_options(options: OpenOptions) -> Self {
        Self::new(
            options.source_limits_value(),
            options.semantic_limits_value(),
            options.xml_limits_value(),
        )
    }

    /// The physical policy retained for the source.
    #[must_use]
    pub const fn source_limits_value(self) -> SourceLimits {
        self.source
    }

    /// The semantic policy retained for the projection.
    #[must_use]
    pub const fn semantic_limits_value(self) -> SemanticLimits {
        self.semantic
    }

    /// The XML stream policy retained for the packaged projection.
    #[must_use]
    pub const fn xml_limits_value(self) -> XmlStreamLimits {
        self.xml
    }

    fn check_source_bytes(self, observed: usize) -> Result<()> {
        let observed = u64::try_from(observed).unwrap_or(u64::MAX);
        let maximum = self.source.max_input_bytes();
        if observed > maximum {
            return Err(Error::Limit {
                kind: LimitKind::Source(crate::SourceLimitKind::InputBytes),
                observed,
                maximum,
            });
        }
        Ok(())
    }

    fn check_package_bytes(self, observed: usize) -> Result<()> {
        self.check_source_bytes(observed)
    }

    fn check_content_bytes(self, observed: usize) -> Result<()> {
        self.semantic.check(
            SemanticLimitKind::ContentBytes,
            u64::try_from(observed).unwrap_or(u64::MAX),
        )
    }

    fn check_styles_bytes(self, observed: usize) -> Result<()> {
        self.semantic.check(
            SemanticLimitKind::StylesBytes,
            u64::try_from(observed).unwrap_or(u64::MAX),
        )
    }

    fn check_output_bytes(self, observed: usize) -> Result<()> {
        self.semantic.check(
            SemanticLimitKind::OutputBytes,
            u64::try_from(observed).unwrap_or(u64::MAX),
        )
    }

    fn check_synthetic_package_bytes(self, observed: usize) -> Result<()> {
        // A package produced by `to_bytes` is an internal ZIP projection. It
        // is not admitted against the caller's flat-source ceiling, but the
        // parser still receives an explicit, hard-bounded scratch profile.
        let observed = u64::try_from(observed).unwrap_or(u64::MAX);
        let maximum = SourceLimits::MAXIMUM.max_input_bytes();
        if observed > maximum {
            return Err(Error::Limit {
                kind: LimitKind::Source(crate::SourceLimitKind::InputBytes),
                observed,
                maximum,
            });
        }
        Ok(())
    }

    fn check_depth(self, observed: usize) -> Result<()> {
        self.semantic.check(
            SemanticLimitKind::Depth,
            u64::try_from(observed).unwrap_or(u64::MAX),
        )
    }

    fn check_children(observed: usize, source_bytes: usize) -> Result<()> {
        // This is a raw XML range index, not a semantic unit count. Every
        // direct child contributes at least one source byte, and the source
        // byte ceiling was checked before scanning. Keeping this proof tied
        // to the admitted buffer avoids charging raw XML structure to the
        // owner's paragraphs/tables/sheets/slides budget.
        if observed > source_bytes {
            return Err(Error::Core(CoreError::InvalidFormat(
                "flat ODF child range count exceeds its source buffer".to_string(),
            )));
        }
        Ok(())
    }

    fn writer_limits() -> StreamingArchiveLimits {
        // This writer emits a synthetic ZIP projection. Its bounded archive
        // defaults are deliberately independent of SemanticLimitKind::OutputBytes,
        // which measures the published flat document. The logical XML parts
        // are checked against their semantic dimensions before this writer is
        // entered; ZIP overhead must not consume the flat output budget.
        let defaults = StreamingArchiveLimits::default();
        defaults.with_byte_limits(
            defaults.max_entry_size,
            defaults.max_total_size,
            SourceLimits::MAXIMUM.max_input_bytes(),
        )
    }

    /// Options for parsing the synthetic package emitted by [`Parts::to_bytes`].
    ///
    /// The package is an internal projection of an already admitted flat
    /// source, so its physical source ceiling is the hard scratch ceiling;
    /// the caller's semantic policy remains attached.
    pub const fn projection_options(self) -> OpenOptions {
        OpenOptions::new()
            .source_limits(SourceLimits::MAXIMUM)
            .semantic_limits(self.semantic)
            .xml_limits(self.xml)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::from_options(OpenOptions::new())
    }
}

/// One package part of a flat document.
#[derive(Debug, Clone, PartialEq, Eq)]
struct Part {
    xml: String,
    /// The flat ranges of the children the part holds, in part order.
    sources: Vec<Range<usize>>,
}

/// The parts of one flat document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Parts {
    mimetype: String,
    content: Part,
    styles: Part,
    meta: Option<Part>,
    settings: Option<Part>,
    /// The flat document's length, which [`Parts::merge`] checks.
    length: usize,
    /// A content proof for the exact source whose ranges are retained below.
    /// Keeping the digest avoids retaining a second full flat source while
    /// still rejecting a same-length foreign source at merge time.
    source_digest: [u8; 32],
    /// The greatest XML nesting depth admitted while splitting the source.
    /// Output paths reuse this proof when a caller supplies a tighter policy.
    depth: usize,
    /// Where the flat root's content starts (after its start tag), for a
    /// part an edit creates.
    root_inner: usize,
}

/// The root start tag and the top-level children of one XML document.
struct TopLevel<'a> {
    root: &'a str,
    root_end: usize,
    children: Vec<(&'a [u8], Range<usize>)>,
    max_depth: usize,
}

/// A fallibly growing sink for the synthetic ZIP projection.
///
/// The archive writer enforces its transport limits before each write; this
/// sink keeps the final in-memory copy bounded and makes allocation failure a
/// typed I/O error rather than an infallible `Vec` growth.
struct BoundedBytes {
    bytes: Vec<u8>,
    limit: usize,
}

impl BoundedBytes {
    fn new(limit: usize) -> Self {
        Self {
            bytes: Vec::new(),
            limit,
        }
    }

    fn into_inner(self) -> Vec<u8> {
        self.bytes
    }
}

impl io::Write for BoundedBytes {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let length = self
            .bytes
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| io::Error::other("flat ODF synthetic ZIP size overflow"))?;
        if length > self.limit {
            return Err(io::Error::other(
                "flat ODF synthetic ZIP exceeds its scratch limit",
            ));
        }
        if length > self.bytes.capacity() {
            let doubled = self.bytes.capacity().saturating_mul(2).max(1);
            let capacity = doubled.max(length).min(self.limit);
            self.bytes
                // `try_reserve_exact` takes additional bytes relative to
                // `len`, not relative to the existing capacity.  Reserving
                // from capacity can leave the next extend infallibly growing
                // beyond this bounded sink's limit when there is slack.
                .try_reserve_exact(capacity.saturating_sub(self.bytes.len()))
                .map_err(|error| io::Error::other(format!("flat ODF ZIP allocation: {error}")))?;
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// Scan `xml`'s root start tag and its direct children.
fn top_level(xml: &str, limits: Limits) -> Result<TopLevel<'_>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut max_depth = 0usize;
    let mut root: Option<(&str, usize)> = None;
    let mut children = Vec::new();
    let mut open_child: Option<(usize, &[u8])> = None;
    loop {
        let start = offset(reader.buffer_position());
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("flat ODF is not well-formed XML: {error}")))?;
        let end = offset(reader.buffer_position());
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    root = Some((&xml[start..end], end));
                } else if depth == 1 {
                    open_child = Some((start, local(xml, start, element.name().as_ref().len())));
                }
                let next_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::Core(CoreError::InvalidFormat(
                        "flat ODF XML nesting depth overflow".to_string(),
                    ))
                })?;
                limits.check_depth(next_depth)?;
                max_depth = max_depth.max(next_depth);
                depth = next_depth;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    return Err(invalid("the flat ODF root element is empty"));
                }
                let next_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::Core(CoreError::InvalidFormat(
                        "flat ODF XML nesting depth overflow".to_string(),
                    ))
                })?;
                limits.check_depth(next_depth)?;
                max_depth = max_depth.max(next_depth);
                if depth == 1 {
                    Limits::check_children(children.len().saturating_add(1), xml.len())?;
                    children
                        .try_reserve(1)
                        .map_err(|source| allocation("flat ODF child index", source))?;
                    children.push((local(xml, start, element.name().as_ref().len()), start..end));
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1
                    && let Some((child_start, name)) = open_child.take()
                {
                    Limits::check_children(children.len().saturating_add(1), xml.len())?;
                    children
                        .try_reserve(1)
                        .map_err(|source| allocation("flat ODF child index", source))?;
                    children.push((name, child_start..end));
                }
            },
            Event::Eof => break,
            Event::DocType(_) => {
                return Err(invalid("flat ODF with a document type declaration"));
            },
            _ => {},
        }
    }
    let (root, root_end) = root.ok_or_else(|| invalid("flat ODF has no root element"))?;
    Ok(TopLevel {
        root,
        root_end,
        children,
        max_depth,
    })
}

impl Parts {
    /// Split a flat document under the caller-selected raw policy.
    ///
    /// # Errors
    ///
    /// Returns an error when `xml` is not UTF-8, not well-formed, has no
    /// `office:document` root in the `OpenDocument` office namespace, or
    /// declares no `office:mimetype`.
    pub fn split(xml: &[u8], limits: &Limits) -> Result<Self> {
        limits.check_source_bytes(xml.len())?;
        limits.check_content_bytes(xml.len())?;
        let text = std::str::from_utf8(xml)
            .map_err(|error| invalid(format!("flat ODF is not UTF-8: {error}")))?;
        let bom = if text.starts_with('\u{feff}') { 3 } else { 0 };
        let body = &text[bom..];
        let scanned = top_level(body, *limits)?;
        let mut reader = Reader::from_str(scanned.root);
        let Ok(Event::Start(element)) = reader.read_event() else {
            return Err(invalid("the flat ODF root start tag is malformed"));
        };
        let (prefix, attributes, mimetype) = root_tag(scanned.root, &element)?;
        let part = |root_local: &str, locals: &[&[u8]]| -> Result<Part> {
            let name = qualified(prefix, root_local)?;
            let maximum = match root_local {
                "document-styles" => SemanticLimitKind::StylesBytes,
                _ => SemanticLimitKind::ContentBytes,
            };
            let header = r#"<?xml version="1.0" encoding="UTF-8"?>"#;
            let closing = 3 + name.len();
            let child_bytes = scanned
                .children
                .iter()
                .filter(|child| locals.contains(&child.0))
                .try_fold(0usize, |total, (_, range)| {
                    total.checked_add(range.len()).ok_or_else(|| {
                        Error::Core(CoreError::InvalidFormat(
                            "flat ODF part length overflow".to_string(),
                        ))
                    })
                })?;
            let length = header
                .len()
                .checked_add(1 + name.len())
                .and_then(|length| length.checked_add(attributes.len()))
                .and_then(|length| length.checked_add(1))
                .and_then(|length| length.checked_add(child_bytes))
                .and_then(|length| length.checked_add(closing))
                .ok_or_else(|| {
                    Error::Core(CoreError::InvalidFormat(
                        "flat ODF part length overflow".to_string(),
                    ))
                })?;
            limits.semantic.check_usize(maximum, length)?;
            let mut xml = String::new();
            xml.try_reserve(length)
                .map_err(|source| allocation("flat ODF part XML", source))?;
            xml.push_str(header);
            xml.push('<');
            xml.push_str(&name);
            xml.push_str(&attributes);
            xml.push('>');
            let source_count = scanned
                .children
                .iter()
                .filter(|child| locals.contains(&child.0))
                .count();
            let mut sources = Vec::new();
            sources
                .try_reserve_exact(source_count)
                .map_err(|source| allocation("flat ODF part child ranges", source))?;
            for wanted in locals {
                for (_, range) in scanned.children.iter().filter(|child| child.0 == *wanted) {
                    xml.push_str(&body[range.clone()]);
                    sources.push(range.start + bom..range.end + bom);
                }
            }
            xml.push_str("</");
            xml.push_str(&name);
            xml.push('>');
            Ok(Part { xml, sources })
        };
        let has = |wanted: &[u8]| scanned.children.iter().any(|child| child.0 == wanted);
        Ok(Self {
            content: part("document-content", CONTENT)?,
            styles: part("document-styles", STYLES)?,
            meta: has(b"meta")
                .then(|| part("document-meta", META))
                .transpose()?,
            settings: has(b"settings")
                .then(|| part("document-settings", SETTINGS))
                .transpose()?,
            mimetype,
            length: xml.len(),
            source_digest: Sha256::digest(xml).into(),
            depth: scanned.max_depth,
            root_inner: scanned.root_end + bom,
        })
    }

    /// The document's declared media type (`office:mimetype`).
    #[must_use]
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// The `content.xml` part.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        &self.content.xml
    }

    /// The `styles.xml` part.
    #[must_use]
    pub fn styles_xml(&self) -> &str {
        &self.styles.xml
    }

    /// The `meta.xml` part, when the document has `office:meta`.
    #[must_use]
    pub fn meta_xml(&self) -> Option<&str> {
        self.meta.as_ref().map(|part| part.xml.as_str())
    }

    /// The `settings.xml` part, when the document has `office:settings`.
    #[must_use]
    pub fn settings_xml(&self) -> Option<&str> {
        self.settings.as_ref().map(|part| part.xml.as_str())
    }

    /// Publish the parts as an in-memory package under the selected raw
    /// policy in `limits`. The logical XML parts are checked against their
    /// semantic dimensions; ZIP framing uses a separate bounded scratch
    /// profile and does not consume the published flat-document output
    /// ceiling.
    ///
    /// # Errors
    ///
    /// Returns an error when the package writer refuses a part.
    pub fn to_bytes(&self, limits: &Limits) -> Result<Vec<u8>> {
        limits.check_source_bytes(self.length)?;
        limits.check_content_bytes(self.length)?;
        limits.check_depth(self.depth)?;
        limits
            .semantic
            .check_usize(SemanticLimitKind::ContentBytes, self.content.xml.len())?;
        limits.check_styles_bytes(self.styles.xml.len())?;
        if let Some(meta) = &self.meta {
            limits
                .semantic
                .check_usize(SemanticLimitKind::ContentBytes, meta.xml.len())?;
        }
        if let Some(settings) = &self.settings {
            limits
                .semantic
                .check_usize(SemanticLimitKind::ContentBytes, settings.xml.len())?;
        }
        let scratch_limit =
            usize::try_from(SourceLimits::MAXIMUM.max_input_bytes()).unwrap_or(usize::MAX);
        let writer =
            PackageWriter::with_writer(BoundedBytes::new(scratch_limit), Limits::writer_limits());
        let mut writer = writer;
        writer.set_mimetype(&self.mimetype)?;
        // The parts carry the producer's bytes (indentation included), so
        // they publish as exact source rather than freshly authored XML.
        for (path, part) in [
            ("content.xml", Some(self.content_xml())),
            ("styles.xml", Some(self.styles_xml())),
            ("meta.xml", self.meta_xml()),
            ("settings.xml", self.settings_xml()),
        ] {
            if let Some(part) = part {
                writer.write_exact_source_file(path, part.as_bytes(), "text/xml")?;
            }
        }
        let bytes = writer.finish_to_writer().map_err(writer_error)?;
        Ok(bytes.into_inner())
    }

    /// Republish `flat` (the document these parts were split from) with an
    /// edit of the parts: each top-level child an edited part changed
    /// replaces that child's bytes; every other flat byte is kept. A part
    /// passed as `None` is unchanged. An edit that creates `meta.xml` for a
    /// flat document without `office:meta` inserts its children first in
    /// the flat root.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when `flat` is not the split source,
    /// including a same-length source with different bytes, and
    /// a typed refusal when an edited part's root start tag changed,
    /// it gained or lost top-level children, a part disappeared, or two
    /// parts changed one shared child differently.
    pub fn merge(
        &self,
        flat: &[u8],
        content: Option<&str>,
        styles: Option<&str>,
        meta: Option<&str>,
        settings: Option<&str>,
        limits: &Limits,
    ) -> Result<Vec<u8>> {
        limits.check_source_bytes(flat.len())?;
        limits.check_content_bytes(flat.len())?;
        if flat.len() != self.length || Sha256::digest(flat).as_slice() != self.source_digest {
            return Err(Error::patch_conflict(
                "odf-common.flat.merge",
                "the flat document is not the one these parts were split from",
            ));
        }
        limits.check_depth(self.depth)?;
        let mut replacements: Vec<(Range<usize>, String)> = Vec::new();
        for (kind, before, after) in [
            (
                SemanticLimitKind::ContentBytes,
                Some(&self.content),
                content,
            ),
            (SemanticLimitKind::StylesBytes, Some(&self.styles), styles),
            (SemanticLimitKind::ContentBytes, self.meta.as_ref(), meta),
            (
                SemanticLimitKind::ContentBytes,
                self.settings.as_ref(),
                settings,
            ),
        ] {
            let Some(after) = after else {
                continue;
            };
            limits.semantic.check_usize(kind, after.len())?;
            match before {
                Some(before) => changed_children(before, after, &mut replacements, *limits)?,
                None => {
                    // A part the edit created (`meta.xml` of a flat document
                    // without `office:meta`): its children go first.
                    let scanned = top_level(after, *limits)?;
                    let mut inserted = String::new();
                    for (_, range) in &scanned.children {
                        inserted
                            .try_reserve(range.len())
                            .map_err(|source| allocation("flat ODF inserted children", source))?;
                        inserted.push_str(&after[range.clone()]);
                    }
                    if !inserted.is_empty() {
                        push_replacement(
                            &mut replacements,
                            self.root_inner..self.root_inner,
                            inserted,
                        )?;
                    }
                },
            }
        }
        replacements.sort_by_key(|(range, _)| (range.start, range.end));
        let output_length = replacements
            .iter()
            .try_fold(flat.len(), |length, (range, text)| {
                length
                    .checked_sub(range.end.saturating_sub(range.start))
                    .and_then(|length| length.checked_add(text.len()))
                    .ok_or_else(|| {
                        Error::Core(CoreError::InvalidFormat(
                            "flat ODF merged output length overflow".to_string(),
                        ))
                    })
            })?;
        limits.check_content_bytes(output_length)?;
        limits.check_output_bytes(output_length)?;
        let mut out = Vec::new();
        out.try_reserve_exact(output_length)
            .map_err(|source| allocation("flat ODF merged output", source))?;
        let mut cursor = 0;
        for (range, text) in replacements {
            out.extend_from_slice(&flat[cursor..range.start]);
            out.extend_from_slice(text.as_bytes());
            cursor = range.end;
        }
        out.extend_from_slice(&flat[cursor..]);
        Ok(out)
    }
}

impl Parts {
    /// [`Parts::merge`] with an independently supplied package. The package
    /// is admitted under the source limits in `limits` before any member is
    /// copied. Owners that pass a package they already admitted should use
    /// [`Self::merge_package_projection`], because an internally generated
    /// ZIP is a projection rather than a second flat-source read.
    ///
    /// # Errors
    ///
    /// As [`Parts::merge`]; a typed refusal for an added member; the
    /// package reader's error for an unreadable package.
    pub fn merge_package(&self, flat: &[u8], package: &[u8], limits: &Limits) -> Result<Vec<u8>> {
        limits.check_package_bytes(package.len())?;
        self.merge_package_parts(
            flat,
            package,
            limits,
            OpenOptions::new()
                .source_limits(limits.source)
                .xml_limits(limits.xml),
        )
    }

    /// Republish a package already admitted by an owner into the original
    /// flat source. The package is still parsed and its members are checked,
    /// but its synthetic ZIP size is charged to the semantic projection
    /// rather than the original flat source ceiling.
    pub fn merge_package_projection(
        &self,
        flat: &[u8],
        package: &[u8],
        limits: &Limits,
    ) -> Result<Vec<u8>> {
        limits.check_synthetic_package_bytes(package.len())?;
        self.merge_package_parts(flat, package, limits, limits.projection_options())
    }

    fn merge_package_parts(
        &self,
        flat: &[u8],
        package: &[u8],
        limits: &Limits,
        options: OpenOptions,
    ) -> Result<Vec<u8>> {
        limits.check_source_bytes(flat.len())?;
        let package = copy_bytes(package, "flat ODF package")?;
        let owned = crate::OwnedPackage::from_bytes_with_options(package, options)?;
        for member in owned.files()? {
            if !matches!(
                member.as_str(),
                "mimetype"
                    | "content.xml"
                    | "styles.xml"
                    | "meta.xml"
                    | "settings.xml"
                    | "META-INF/"
                    | "META-INF/manifest.xml"
            ) {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odf-common.flat",
                    "the edit added the package member a member which a flat ODF document cannot hold",
                )));
            }
        }
        let read = |path: &str| -> Result<Option<String>> {
            if !owned.has_file(path)? {
                return Ok(None);
            }
            String::from_utf8(owned.get_file(path)?)
                .map(Some)
                .map_err(|error| invalid(format!("edited {path} is not UTF-8: {error}")))
        };
        let content =
            read("content.xml")?.ok_or_else(|| invalid("the edited package has no content.xml"))?;
        let styles = read("styles.xml")?;
        let meta = read("meta.xml")?;
        let settings = read("settings.xml")?;
        self.merge(
            flat,
            Some(&content),
            styles.as_deref(),
            meta.as_deref(),
            settings.as_deref(),
            limits,
        )
    }
}

/// Record the children `after` changed relative to `before`.
fn changed_children(
    before: &Part,
    after: &str,
    replacements: &mut Vec<(Range<usize>, String)>,
    limits: Limits,
) -> Result<()> {
    let old = top_level(&before.xml, limits)?;
    let new = top_level(after, limits)?;
    if old.root != new.root {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odf-common.flat",
            "the edit changed a flat ODF part's root element",
        )));
    }
    if old.children.len() != new.children.len()
        || old.children.len() != before.sources.len()
        || old
            .children
            .iter()
            .zip(&new.children)
            .any(|(left, right)| left.0 != right.0)
    {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odf-common.flat",
            "the edit added or removed a top-level element of a flat ODF part",
        )));
    }
    for ((old_child, new_child), source) in
        old.children.iter().zip(&new.children).zip(&before.sources)
    {
        let old_text = &before.xml[old_child.1.clone()];
        let new_text = &after[new_child.1.clone()];
        if old_text != new_text {
            let mut replacement = String::new();
            replacement
                .try_reserve(new_text.len())
                .map_err(|source| allocation("flat ODF replacement", source))?;
            replacement.push_str(new_text);
            push_replacement(replacements, source.clone(), replacement)?;
        }
    }
    Ok(())
}

/// Add one replacement; one shared child changed two ways is refused.
fn push_replacement(
    replacements: &mut Vec<(Range<usize>, String)>,
    range: Range<usize>,
    text: String,
) -> Result<()> {
    if let Some(existing) = replacements.iter().find(|(existing, _)| *existing == range) {
        if existing.1 == text {
            return Ok(());
        }
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odf-common.flat",
            "the edit changed one shared flat ODF element differently in two parts",
        )));
    }
    replacements
        .try_reserve(1)
        .map_err(|source| allocation("flat ODF replacements", source))?;
    replacements.push((range, text));
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Core(CoreError::InvalidFormat(message.into()))
}

fn copy_bytes(bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(bytes.len())
        .map_err(|source| allocation(resource, source))?;
    owned.extend_from_slice(bytes);
    Ok(owned)
}

fn writer_error(error: crate::core::PackageWriterError) -> Error {
    if let Some(limit) = error.limit()
        && matches!(
            limit.resource(),
            crate::core::PackageWriterLimitResource::OutputBytes
        )
    {
        return Error::Core(CoreError::ResourceLimit(ResourceLimit {
            resource: Resource::OutputBytes,
            observed: limit.actual(),
            limit: limit.maximum(),
            scope: Arc::from("ODF synthetic ZIP scratch"),
        }));
    }
    Error::from(error.into_core_error())
}

fn offset(position: u64) -> usize {
    usize::try_from(position).unwrap_or(usize::MAX)
}

/// The local name of the element whose tag starts at `start` with a
/// qualified name of `length` bytes.
fn local(xml: &str, start: usize, length: usize) -> &[u8] {
    let name = &xml.as_bytes()[start + 1..start + 1 + length];
    name.iter()
        .position(|byte| *byte == b':')
        .map_or(name, |colon| &name[colon + 1..])
}

fn qualified(prefix: &str, local: &str) -> Result<String> {
    let length = prefix
        .len()
        .checked_add(usize::from(!prefix.is_empty()))
        .and_then(|length| length.checked_add(local.len()))
        .ok_or_else(|| {
            Error::Core(CoreError::InvalidFormat(
                "flat ODF qualified name length overflow".to_string(),
            ))
        })?;
    let mut name = String::new();
    name.try_reserve(length)
        .map_err(|source| allocation("flat ODF qualified name", source))?;
    name.push_str(prefix);
    if !prefix.is_empty() {
        name.push(':');
    }
    name.push_str(local);
    Ok(name)
}

/// The root's prefix, its attribute text without `office:mimetype`, and
/// the media type it declares.
fn root_tag<'a>(
    tag: &'a str,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<(&'a str, String, String)> {
    let name_length = element.name().as_ref().len();
    let name = &tag[1..=name_length];
    let (prefix, local) = name.split_once(':').unwrap_or(("", name));
    if local != "document" {
        return Err(invalid("the flat ODF root is not office:document"));
    }
    let mimetype_key = qualified(prefix, "mimetype")?;
    let mut declared = false;
    let mut mimetype = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("flat ODF root attribute: {error}")))?;
        let key = attribute.key.as_ref();
        let binds_prefix = if prefix.is_empty() {
            key == b"xmlns"
        } else {
            key.strip_prefix(b"xmlns:") == Some(prefix.as_bytes())
        };
        if binds_prefix && attribute.value.as_ref() == OFFICE {
            declared = true;
        }
        if key == mimetype_key.as_bytes() {
            // Media types are ASCII; a reference in one is not one.
            let value = String::from_utf8_lossy(attribute.value.as_ref());
            let mut owned = String::new();
            owned
                .try_reserve(value.len())
                .map_err(|source| allocation("flat ODF MIME type", source))?;
            owned.push_str(&value);
            mimetype = Some(owned);
        }
    }
    if !declared {
        return Err(invalid(
            "the flat ODF root is not in the OpenDocument office namespace",
        ));
    }
    let mimetype = mimetype.ok_or_else(|| invalid("the flat ODF root declares no mimetype"))?;
    let attribute_text = &tag[1 + name_length..tag.len() - 1];
    let mut attributes = String::new();
    attributes
        .try_reserve(attribute_text.len())
        .map_err(|source| allocation("flat ODF root attributes", source))?;
    attributes.push_str(attribute_text);
    // Drop `prefix:mimetype="…"` (either quote) from the attribute text.
    let needle = format!("{mimetype_key}=");
    if let Some(at) = attributes.find(&needle) {
        let value_start = at + needle.len();
        if let Some(quote) = attributes[value_start..].chars().next()
            && let Some(close) = attributes[value_start + 1..].find(quote)
        {
            attributes.replace_range(at..value_start + 1 + close + 1, "");
        }
    }
    Ok((prefix, attributes, mimetype))
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Core(CoreError::Allocation { resource, source })
}

#[cfg(test)]
mod tests {
    use super::{Limits, Parts};
    use crate::Error;
    use crate::XmlStreamLimits;
    use crate::limits::{
        LimitKind, SemanticLimitKind, SemanticLimits, SourceLimitKind, SourceLimits,
    };
    use std::io::Write as _;

    const FLAT: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n",
        "<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
        "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.3\" ",
        "office:mimetype=\"application/vnd.oasis.opendocument.text\">\n",
        " <office:meta><dc:title xmlns:dc=\"http://purl.org/dc/elements/1.1/\">T</dc:title></office:meta>\n",
        " <office:styles><style:style xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" style:name=\"S\"/></office:styles>\n",
        " <office:automatic-styles/>\n",
        " <office:body>\n  <office:text><text:p>Hello</text:p></office:text>\n </office:body>\n",
        "</office:document>\n",
    );

    fn limits() -> Limits {
        Limits::default()
    }

    #[test]
    fn splits_children_into_package_parts_keeping_their_bytes() {
        let limits = limits();
        let parts = Parts::split(FLAT.as_bytes(), &limits).unwrap();
        assert_eq!(parts.mimetype(), "application/vnd.oasis.opendocument.text");
        let root = " xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.3\" ";
        assert_eq!(
            parts.content_xml(),
            format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content{root}><office:automatic-styles/><office:body>\n  <office:text><text:p>Hello</text:p></office:text>\n </office:body></office:document-content>"
            )
        );
        assert!(parts.styles_xml().contains("<style:style xmlns:style="));
        assert!(parts.styles_xml().contains("<office:document-styles"));
        assert!(parts.meta_xml().unwrap().contains("<dc:title"));
        assert_eq!(parts.settings_xml(), None);
        let package = parts.to_bytes(&limits).unwrap();
        assert_eq!(&package[..2], b"PK");
    }

    #[test]
    fn refuses_what_is_not_a_flat_document() {
        let limits = limits();
        assert!(Parts::split(b"<root/>", &limits).is_err());
        assert!(Parts::split(b"<office:document xmlns:office=\"x\"/>", &limits).is_err());
        assert!(
            Parts::split(
                b"<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><office:body/></office:document>",
                &limits
            )
            .is_err(),
            "no mimetype"
        );
        assert!(Parts::split(b"<office:document><unclosed>", &limits).is_err());
        assert!(Parts::split(&[0xff, 0xfe], &limits).is_err());
    }

    #[test]
    fn merges_changed_children_back_and_keeps_every_other_byte() {
        let limits = limits();
        let parts = Parts::split(FLAT.as_bytes(), &limits).unwrap();
        let content = parts.content_xml().replace("Hello", "Bye");
        let merged = parts
            .merge(FLAT.as_bytes(), Some(&content), None, None, None, &limits)
            .unwrap();
        assert_eq!(
            String::from_utf8(merged).unwrap(),
            FLAT.replace("Hello", "Bye")
        );
        // An unchanged edit is the flat document itself.
        let same = parts
            .merge(
                FLAT.as_bytes(),
                Some(parts.content_xml()),
                Some(parts.styles_xml()),
                parts.meta_xml(),
                None,
                &limits,
            )
            .unwrap();
        assert_eq!(same, FLAT.as_bytes());
    }

    #[test]
    fn refuses_root_and_structure_changes_and_shared_conflicts() {
        let limits = limits();
        let parts = Parts::split(FLAT.as_bytes(), &limits).unwrap();
        let rooted = parts
            .content_xml()
            .replace("office:version=\"1.3\"", "office:version=\"1.4\"");
        assert!(matches!(
            parts.merge(FLAT.as_bytes(), Some(&rooted), None, None, None, &limits),
            Err(
                Error::Core(litchi_core::Error::Unsupported(_) | litchi_core::Error::Refused(_))
                    | Error::Refused(_)
            )
        ));
        let added = parts.content_xml().replace(
            "<office:automatic-styles/>",
            "<office:automatic-styles/><office:scripts/>",
        );
        assert!(
            parts
                .merge(FLAT.as_bytes(), Some(&added), None, None, None, &limits)
                .is_err()
        );
        let one = parts.content_xml().replace(
            "<office:automatic-styles/>",
            "<office:automatic-styles><a/></office:automatic-styles>",
        );
        let two = parts.styles_xml().replace(
            "<office:automatic-styles/>",
            "<office:automatic-styles><b/></office:automatic-styles>",
        );
        assert!(
            parts
                .merge(FLAT.as_bytes(), Some(&one), Some(&two), None, None, &limits)
                .is_err()
        );
        assert!(
            parts
                .merge(b"short", Some(&one), None, None, None, &limits)
                .is_err()
        );

        let foreign = FLAT.replace("Hello", "Jello");
        assert_eq!(foreign.len(), FLAT.len());
        assert!(matches!(
            parts.merge(
                foreign.as_bytes(),
                Some(parts.content_xml()),
                None,
                None,
                None,
                &limits,
            ),
            Err(Error::PatchConflict { .. })
        ));
    }

    #[test]
    fn merges_an_edited_package() {
        let limits = limits();
        let parts = Parts::split(FLAT.as_bytes(), &limits).unwrap();
        let package = parts.to_bytes(&limits).unwrap();
        assert_eq!(
            parts
                .merge_package(FLAT.as_bytes(), &package, &limits)
                .unwrap(),
            FLAT.as_bytes()
        );
    }

    #[test]
    fn a_created_meta_part_goes_first() {
        let flat = FLAT.replace(
            " <office:meta><dc:title xmlns:dc=\"http://purl.org/dc/elements/1.1/\">T</dc:title></office:meta>\n",
            "",
        );
        let limits = limits();
        let parts = Parts::split(flat.as_bytes(), &limits).unwrap();
        assert_eq!(parts.meta_xml(), None);
        let meta = "<?xml version=\"1.0\"?><office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><office:meta><x/></office:meta></office:document-meta>";
        let merged = String::from_utf8(
            parts
                .merge(flat.as_bytes(), None, None, Some(meta), None, &limits)
                .unwrap(),
        )
        .unwrap();
        assert!(merged.contains("application/vnd.oasis.opendocument.text\"><office:meta><x/></office:meta>\n <office:styles>"));
    }

    #[test]
    fn flat_source_and_semantic_limits_are_typed_at_boundaries() {
        let source = SourceLimits::builder()
            .max_input_bytes(FLAT.len() as u64)
            .unwrap()
            .build()
            .unwrap();
        let semantic = SemanticLimits::builder()
            .max_content_bytes(FLAT.len() as u64)
            .unwrap()
            .build()
            .unwrap();
        let limits = Limits::new(source, semantic, XmlStreamLimits::default());
        assert!(Parts::split(FLAT.as_bytes(), &limits).is_ok());

        let one_under_source = SourceLimits::builder()
            .max_input_bytes((FLAT.len() - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        let error = Parts::split(
            FLAT.as_bytes(),
            &Limits::new(one_under_source, semantic, XmlStreamLimits::default()),
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Source(SourceLimitKind::InputBytes),
                observed,
                maximum,
            } if observed == FLAT.len() as u64 && maximum == (FLAT.len() - 1) as u64
        ));

        let one_under_content = SemanticLimits::builder()
            .max_content_bytes((FLAT.len() - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        let error = Parts::split(
            FLAT.as_bytes(),
            &Limits::new(source, one_under_content, XmlStreamLimits::default()),
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Semantic(SemanticLimitKind::ContentBytes),
                observed,
                maximum,
            } if observed == FLAT.len() as u64 && maximum == (FLAT.len() - 1) as u64
        ));
    }

    #[test]
    fn generated_package_uses_output_policy_not_flat_input_policy() {
        let source = SourceLimits::builder()
            .max_input_bytes(FLAT.len() as u64)
            .unwrap()
            .build()
            .unwrap();
        let limits = Limits::new(source, SemanticLimits::DEFAULT, XmlStreamLimits::default());
        let parts = Parts::split(FLAT.as_bytes(), &limits).unwrap();
        let package = parts.to_bytes(&limits).unwrap();
        assert!(package.len() > FLAT.len());

        let error = parts
            .merge_package(FLAT.as_bytes(), &package, &limits)
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Source(SourceLimitKind::InputBytes),
                ..
            }
        ));
        assert_eq!(
            parts
                .merge_package_projection(FLAT.as_bytes(), &package, &limits)
                .unwrap(),
            FLAT.as_bytes()
        );
    }

    #[test]
    fn projection_retains_xml_policy_and_uses_scratch_source_ceiling() {
        let xml = XmlStreamLimits::new(8 * 1024 * 1024, 128, 200_000, 512, 1024 * 1024)
            .expect("finite selected XML policy");
        let semantic = SemanticLimits::DEFAULT;
        let limits = Limits::new(SourceLimits::DEFAULT, semantic, xml);

        let projected = limits.projection_options();
        // The generated ZIP is an internal projection, so it intentionally
        // uses the hard scratch source ceiling rather than the flat source's
        // 512 MiB admission bound. The selected XML policy must nevertheless
        // cross the projection boundary unchanged.
        assert_eq!(projected.source_limits_value(), SourceLimits::MAXIMUM);
        assert_eq!(projected.semantic_limits_value(), semantic);
        assert_eq!(projected.xml_limits_value(), xml);

        let parts = Parts::split(FLAT.as_bytes(), &limits).expect("flat source");
        let package = parts.to_bytes(&limits).expect("synthetic package");
        assert_eq!(
            parts
                .merge_package_projection(FLAT.as_bytes(), &package, &limits)
                .expect("projection remains admissible"),
            FLAT.as_bytes()
        );
    }

    #[test]
    fn output_and_depth_limits_fail_before_publication() {
        let unrestricted = Limits::default();
        let parts = Parts::split(FLAT.as_bytes(), &unrestricted).unwrap();
        let package = parts.to_bytes(&unrestricted).unwrap();
        let output = SemanticLimits::builder()
            .max_output_bytes((FLAT.len() - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        let limits = Limits::new(SourceLimits::DEFAULT, output, XmlStreamLimits::default());
        let parts = Parts::split(FLAT.as_bytes(), &limits).unwrap();
        // ZIP framing is an internal projection and does not consume the
        // published flat-document output budget.
        assert!(!package.is_empty());
        assert!(!parts.to_bytes(&limits).unwrap().is_empty());
        assert!(matches!(
            parts.merge(
                FLAT.as_bytes(),
                Some(parts.content_xml()),
                Some(parts.styles_xml()),
                parts.meta_xml(),
                parts.settings_xml(),
                &limits,
            ),
            Err(Error::Limit {
                kind: LimitKind::Semantic(SemanticLimitKind::OutputBytes),
                ..
            })
        ));

        let depth = SemanticLimits::builder()
            .max_depth(2)
            .unwrap()
            .build()
            .unwrap();
        let error = Parts::split(
            FLAT.as_bytes(),
            &Limits::new(SourceLimits::DEFAULT, depth, XmlStreamLimits::default()),
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Semantic(SemanticLimitKind::Depth),
                ..
            }
        ));

        let tight = Limits::new(
            SourceLimits::DEFAULT,
            SemanticLimits::builder()
                .max_depth(2)
                .unwrap()
                .build()
                .unwrap(),
            XmlStreamLimits::default(),
        );
        // The source was admitted under the loose profile.  Every output
        // route must still recheck its retained proof under the recipient's
        // tighter policy, including unchanged edits.
        assert!(parts.to_bytes(&tight).is_err());
        assert!(
            parts
                .merge(FLAT.as_bytes(), None, None, None, None, &tight)
                .is_err()
        );
    }

    #[test]
    fn empty_elements_are_checked_at_their_attempted_depth() {
        let flat = concat!(
            "<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
            "office:mimetype=\"application/vnd.oasis.opendocument.text\">",
            "<office:body><office:text/></office:body>",
            "</office:document>",
        );
        let limits = Limits::new(
            SourceLimits::DEFAULT,
            SemanticLimits::builder()
                .max_depth(2)
                .unwrap()
                .build()
                .unwrap(),
            XmlStreamLimits::default(),
        );
        assert!(matches!(
            Parts::split(flat.as_bytes(), &limits),
            Err(Error::Limit {
                kind: LimitKind::Semantic(SemanticLimitKind::Depth),
                ..
            })
        ));
    }

    #[test]
    fn bounded_zip_sink_reserves_from_length_when_capacity_has_slack() {
        let mut sink = super::BoundedBytes {
            bytes: Vec::with_capacity(64),
            limit: 256,
        };
        sink.write_all(b"seed").unwrap();
        let payload = vec![b'x'; 196];
        sink.write_all(&payload).unwrap();
        assert_eq!(sink.bytes.len(), 200);
        assert!(sink.bytes.capacity() <= sink.limit);
    }
}
