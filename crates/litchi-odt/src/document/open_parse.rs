//! Fused single-tokenization open parse for ODT `content.xml`.
//!
//! [`crate::document::SourceBackedDocument`] opening historically tokenized
//! `content.xml` twice: once for the package-content structure validation
//! ([`litchi_odf_common::core::validate_content_document_part`]) and once for
//! the automatic-styles scan ([`StyleRegistry::from_xml`]).  This module runs
//! both over one shared quick-xml [`Reader`] event stream, with the
//! namespace-binding maintenance `NsReader` would perform replicated by the
//! hand-rolled [`BindingTracker`] (change 0224).
//!
//! Observable behavior is identical to the sequential passes:
//!
//! - The validation handler replicates every check of
//!   `validate_content_document_part` for the ODT root and family with
//!   identical error messages, applied to the same events in the same order;
//!   its mid-stream errors abort the scan and return immediately, exactly
//!   where the historical first pass early-returned before `styles.xml` was
//!   even fetched.
//! - The style handler replicates `StyleRegistry::from_xml` byte-exactly,
//!   including its literal raw-qualified-name `style:style` match (no
//!   namespace resolution), its raw undecoded attribute values, and its
//!   error messages; its first error is recorded and later events no longer
//!   mutate its state, matching the historical early return.
//! - A tokenization failure surfaces with the validation mapping
//!   (`invalid ODT content.xml: …`): both historical scans tokenize the same
//!   bytes with the same quick-xml core from the start, and validation ran
//!   first and read to EOF, so it reported every read failure before the
//!   style scan ever started.  The fused reader keeps the validation
//!   configuration (`check_end_names` and `check_comments`), which is a
//!   superset of the historical style-scan checks, so no tokenization error
//!   can be introduced or masked relative to history.
//!
//! Error precedence after a clean scan: the validation end-of-stream check
//! runs first (still ahead of any `styles.xml` work), then the caller parses
//! `styles.xml`, then [`OpenParse::finish`] surfaces the recorded
//! content-styles error, and finally `try_extend` runs — the historical
//! call order.
//!
//! Both facades run this fused parse: the source-backed
//! [`crate::document::SourceBackedDocument`] and, since change 0222, the
//! owned [`crate::document::Document`] open path. The standalone
//! `validate_content_document_part` stays byte-identical in
//! litchi-odf-common, and the historical sequential sequence survives as
//! cfg(test) equivalence oracles (the `sequential` helper in this module's
//! tests and the owned-path oracle in `super::package`'s tests).

use litchi_core::{Error, Result};
use quick_xml::events::{BytesRef, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::Reader;

use super::source::{CONTENT_ROOT, FAMILY_NAME};
use crate::binding_tracker::BindingTracker;
use crate::elements::element::Element;
use crate::elements::style::{Style, StyleRegistry};
use crate::elements::text::{Kind, TEXT_NAMESPACE, TextBlockKindHandler};

// Parity with the `validate_content_document_part` pre-scan limit in
// litchi-odf-common (`MAX_CONTENT_BYTES` there is crate-private).
const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

/// Validate the ODT content root and build its visible text-block catalog in
/// one borrowing `Reader` pass.
///
/// Validation is driven before the catalog handler for every event. A catalog
/// error is deferred while tokenization and validation continue to EOF so a
/// later reader or validation/finish error preserves the historical
/// validate-then-scan precedence.
pub(crate) fn run_catalog(content_xml: &str) -> Result<Vec<Kind>> {
    if content_xml.len() > MAX_CONTENT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{FAMILY_NAME} content.xml exceeds the family limit"
        )));
    }
    let mut validate = ValidateHandler::new()?;
    let mut catalog = TextBlockKindHandler::new();
    let mut catalog_error = None;

    let mut reader = Reader::from_str(content_xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().check_comments = true;
    let mut tracker = BindingTracker::new();
    let mut pending_pop = false;
    loop {
        if pending_pop {
            tracker.pop();
            pending_pop = false;
        }
        let event = match reader.read_event() {
            Ok(event) => event,
            Err(error) => {
                return Err(Error::InvalidFormat(format!(
                    "invalid {FAMILY_NAME} content.xml: {error}"
                )));
            },
        };
        match &event {
            Event::Start(element) => {
                tracker.push(element).map_err(|error| {
                    Error::InvalidFormat(format!("invalid {FAMILY_NAME} content.xml: {error}"))
                })?;
            },
            Event::Empty(element) => {
                tracker.push(element).map_err(|error| {
                    Error::InvalidFormat(format!("invalid {FAMILY_NAME} content.xml: {error}"))
                })?;
                pending_pop = true;
            },
            Event::End(_) => pending_pop = true,
            _ => {},
        }

        let (office, text_namespace) = match &event {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, _) = tracker.resolve_element(element.name());
                match namespace {
                    ResolveResult::Bound(Namespace(uri)) => (
                        validate.depth <= 2 && uri == OFFICE_NAMESPACE,
                        uri == TEXT_NAMESPACE,
                    ),
                    _ => (false, false),
                }
            },
            _ => (false, false),
        };

        let is_eof = matches!(&event, Event::Eof);
        validate.on_event(office, &event)?;
        if catalog_error.is_none()
            && let Err(error) = catalog.on_event(text_namespace, &event)
        {
            catalog_error = Some(error);
        }
        if is_eof {
            break;
        }
    }

    validate.finish()?;
    if let Some(error) = catalog_error {
        return Err(error);
    }
    catalog.finish()
}

/// The two-handler fused open parse over `content.xml`.
pub(crate) struct OpenParse {
    registry: StyleRegistry,
    style_error: Option<Error>,
}

impl OpenParse {
    /// Tokenize `content_xml` once, driving the content-validation and
    /// automatic-styles handlers over the shared event stream.
    ///
    /// A validation error (mid-stream, read failure, or end-of-stream) is
    /// returned immediately, before the caller fetches `styles.xml`, exactly
    /// where the historical first pass early-returned.  A content-styles
    /// error is only recorded; [`OpenParse::finish`] reports it after the
    /// `styles.xml` parse, matching the historical pass order.
    pub(crate) fn run(content_xml: &str) -> Result<Self> {
        if content_xml.len() > MAX_CONTENT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "{FAMILY_NAME} content.xml exceeds the family limit"
            )));
        }
        let mut validate = ValidateHandler::new()?;
        let mut styles = StyleHandler::default();
        let mut style_error = None;

        let mut reader = Reader::from_str(content_xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        let mut tracker = BindingTracker::new();
        let mut pending_pop = false;
        loop {
            // Borrowing read: events borrow `content_xml` directly, avoiding
            // the per-event buffer copy of the `_into` API. The tokenization
            // error stream is identical to the historical validation-first
            // read (which tokenized the same bytes and read to EOF): the
            // plain `Reader` here is the same tokenizer `NsReader` wraps,
            // with the same configuration.
            //
            // Binding maintenance replicates `NsReader::read_event_impl`'s
            // ordering: a deferred pop of the previous `End`/`Empty` scope
            // runs before the read, and the push for a `Start`/`Empty` runs
            // before the event is classified, so a namespace error preempts
            // the validation error for the same element, exactly where
            // `read_event` returned `Err` historically.
            if pending_pop {
                tracker.pop();
                pending_pop = false;
            }
            let event = match reader.read_event() {
                Ok(event) => event,
                Err(error) => {
                    return Err(Error::InvalidFormat(format!(
                        "invalid {FAMILY_NAME} content.xml: {error}"
                    )));
                },
            };
            match &event {
                Event::Start(element) => {
                    // `NamespaceError`'s `Display` is what
                    // `quick_xml::Error::Namespace` forwards to, so this
                    // message is byte-identical to the `NsReader` failure.
                    tracker.push(element).map_err(|error| {
                        Error::InvalidFormat(format!("invalid {FAMILY_NAME} content.xml: {error}"))
                    })?;
                },
                Event::Empty(element) => {
                    tracker.push(element).map_err(|error| {
                        Error::InvalidFormat(format!("invalid {FAMILY_NAME} content.xml: {error}"))
                    })?;
                    pending_pop = true;
                },
                Event::End(_) => pending_pop = true,
                _ => {},
            }
            // The ValidateHandler consumes the resolved namespace only in its
            // Start/Empty arms at depth <= 2 (root, `office:body`,
            // `office:forms`/family element — the same arms as the standalone
            // validator); the StyleHandler matches raw qualified names and
            // never resolves. Bindings declared at depth >= 3 scope just their
            // own subtree, so no consumed resolution can change. Resolve only
            // where the result is observable, exactly where the historical
            // validation loop classified it. The tracker still maintains
            // bindings (and its byte-exact error stream) for every event at
            // any depth, as `NsReader` did.
            let office = match &event {
                Event::Start(element) | Event::Empty(element) if validate.depth <= 2 => {
                    let (namespace, _) = tracker.resolve_element(element.name());
                    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE)
                },
                _ => false,
            };
            let is_eof = matches!(event, Event::Eof);
            // Validation errors early-returned historically, discarding any
            // style state; abort the scan at the same event.
            validate.on_event(office, &event)?;
            if style_error.is_none()
                && let Err(error) = styles.on_event(&event)
            {
                style_error = Some(error);
            }
            if is_eof {
                break;
            }
        }
        // The historical validation pass completed (including its
        // end-of-stream checks) before the style scan started, so a
        // structural failure here still beats any recorded style error.
        validate.finish()?;

        Ok(Self {
            registry: styles.registry,
            style_error,
        })
    }

    /// Surface the recorded content-styles error or return the registry the
    /// historical `StyleRegistry::from_xml` pass would have built.
    pub(crate) fn finish(self) -> Result<StyleRegistry> {
        if let Some(error) = self.style_error {
            return Err(error);
        }
        Ok(self.registry)
    }
}

/// Streaming event handler replicating `validate_content_document_part` for
/// the ODT root (`office:document-content`), body (`office:body`), and family
/// element (`office:text`, derived from [`CONTENT_ROOT`]).
///
/// Every check, error message, and error position matches the standalone
/// pass, which keeps its own inline loop for the other callers.
#[derive(Debug)]
struct ValidateHandler {
    expected_local: String,
    depth: usize,
    root_closed: bool,
    body_seen: bool,
    expected_seen: bool,
    in_body: bool,
    declaration_seen: bool,
    first_event: bool,
}

impl ValidateHandler {
    fn new() -> Result<Self> {
        // Same marker derivation as the standalone pass; unreachable for the
        // fixed ODT root but kept at the same position.
        let expected_local = CONTENT_ROOT
            .strip_prefix("<office:")
            .and_then(|marker| {
                marker
                    .split(|character: char| !character.is_ascii_alphanumeric() && character != '-')
                    .next()
            })
            .filter(|local| !local.is_empty())
            .ok_or_else(|| {
                Error::InvalidFormat(format!("{FAMILY_NAME} content.xml has no expected body"))
            })?;
        Ok(Self {
            expected_local: expected_local.to_string(),
            depth: 0,
            root_closed: false,
            body_seen: false,
            expected_seen: false,
            in_body: false,
            declaration_seen: false,
            first_event: true,
        })
    }

    fn on_event(&mut self, office: bool, event: &Event<'_>) -> Result<()> {
        match event {
            Event::Start(element) => {
                if self.root_closed {
                    return Err(Error::InvalidFormat(format!(
                        "{FAMILY_NAME} content.xml has content after its root"
                    )));
                }
                let local = element.local_name();
                match self.depth {
                    0 if office && local.as_ref() == b"document-content" => self.depth = 1,
                    0 => {
                        return Err(Error::InvalidFormat(format!(
                            "{FAMILY_NAME} content.xml has the wrong root"
                        )));
                    },
                    1 if office && local.as_ref() == b"body" && !self.body_seen => {
                        self.body_seen = true;
                        self.in_body = true;
                        self.depth = 2;
                    },
                    1 if office && local.as_ref() == b"body" => {
                        return Err(Error::InvalidFormat(format!(
                            "{FAMILY_NAME} content.xml has duplicate office:body"
                        )));
                    },
                    1 => {
                        self.depth = self.depth.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "{FAMILY_NAME} content.xml nesting overflows"
                            ))
                        })?;
                    },
                    2 if self.in_body
                        && office
                        && local.as_ref() == b"forms"
                        && !self.expected_seen =>
                    {
                        self.depth = 3;
                    },
                    2 if self.in_body
                        && office
                        && local.as_ref() == self.expected_local.as_bytes()
                        && !self.expected_seen =>
                    {
                        self.expected_seen = true;
                        self.depth = 3;
                    },
                    2 if self.in_body => {
                        return Err(Error::InvalidFormat(format!(
                            "{FAMILY_NAME} content.xml has the wrong office body"
                        )));
                    },
                    _ => {
                        self.depth = self.depth.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "{FAMILY_NAME} content.xml nesting overflows"
                            ))
                        })?;
                    },
                }
            },
            Event::Empty(element) => {
                if self.root_closed || self.depth == 0 {
                    return Err(Error::InvalidFormat(format!(
                        "{FAMILY_NAME} content.xml has an invalid empty root"
                    )));
                }
                let local = element.local_name();
                if self.depth == 1 {
                    if office && local.as_ref() == b"body" && !self.body_seen {
                        self.body_seen = true;
                    } else if office && local.as_ref() == b"body" {
                        return Err(Error::InvalidFormat(format!(
                            "{FAMILY_NAME} content.xml has duplicate office:body"
                        )));
                    }
                } else if self.in_body
                    && self.depth == 2
                    && office
                    && local.as_ref() == b"forms"
                    && !self.expected_seen
                {
                    // `office:forms` may precede the family body.
                } else if self.in_body
                    && self.depth == 2
                    && office
                    && local.as_ref() == self.expected_local.as_bytes()
                    && !self.expected_seen
                {
                    self.expected_seen = true;
                } else if self.in_body && self.depth == 2 {
                    return Err(Error::InvalidFormat(format!(
                        "{FAMILY_NAME} content.xml has the wrong office body"
                    )));
                }
            },
            Event::End(_) => {
                self.depth = self.depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat(format!("{FAMILY_NAME} content.xml has an unexpected end"))
                })?;
                if self.depth == 0 {
                    self.root_closed = true;
                } else if self.in_body && self.depth == 1 {
                    self.in_body = false;
                }
            },
            Event::Text(text) => {
                if (self.depth == 0 || self.root_closed)
                    && !text.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(Error::InvalidFormat(format!(
                        "{FAMILY_NAME} content.xml has unexpected text outside its root"
                    )));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if self.depth == 0 || self.root_closed => {
                return Err(Error::InvalidFormat(format!(
                    "{FAMILY_NAME} content.xml has content outside its root"
                )));
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(format!(
                    "{FAMILY_NAME} content.xml must not contain a doctype"
                )));
            },
            Event::GeneralRef(reference) if !valid_xml_reference(reference) => {
                return Err(Error::InvalidFormat(format!(
                    "{FAMILY_NAME} content.xml has an invalid character or entity reference"
                )));
            },
            Event::Decl(_)
                if self.declaration_seen
                    || !self.first_event
                    || self.depth != 0
                    || self.root_closed =>
            {
                return Err(Error::InvalidFormat(format!(
                    "{FAMILY_NAME} content.xml has an XML declaration outside its prologue"
                )));
            },
            Event::Decl(_) => self.declaration_seen = true,
            Event::Eof => {},
            Event::Comment(_) | Event::PI(_) | Event::CData(_) | Event::GeneralRef(_) => {},
        }
        self.first_event = false;
        Ok(())
    }

    fn finish(self) -> Result<()> {
        if !self.root_closed || self.depth != 0 || !self.body_seen || !self.expected_seen {
            return Err(Error::InvalidFormat(format!(
                "{FAMILY_NAME} content.xml has no complete expected body"
            )));
        }
        Ok(())
    }
}

// Byte-identical replication of the crate-private reference predicate in
// litchi-odf-common (`validation::valid_xml_reference`), which the standalone
// validation pass applies to general entity and character references.
fn valid_xml_reference(reference: &BytesRef<'_>) -> bool {
    let bytes: &[u8] = reference;
    matches!(bytes, b"amp" | b"lt" | b"gt" | b"apos" | b"quot")
        || (reference.is_char_ref()
            && reference
                .resolve_char_ref()
                .ok()
                .flatten()
                .is_some_and(xml10_character))
}

fn xml10_character(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

/// Streaming event handler replicating [`StyleRegistry::from_xml`]: the raw
/// qualified-name `style:style` match is intentionally literal and
/// prefix-sensitive, attribute values stay undecoded, and the error messages
/// are identical.
#[derive(Debug, Default)]
struct StyleHandler {
    registry: StyleRegistry,
}

impl StyleHandler {
    fn on_event(&mut self, event: &Event<'_>) -> Result<()> {
        if let Event::Start(element) = event
            && element.name().as_ref() == b"style:style"
        {
            let mut style_element = Element::try_new("style:style")?;
            for attribute in element.attributes() {
                let attribute = attribute.map_err(|error| {
                    Error::XmlError(format!("invalid ODT style attribute: {error}"))
                })?;
                let key = std::str::from_utf8(attribute.key.as_ref()).map_err(|error| {
                    Error::XmlError(format!(
                        "invalid UTF-8 in ODT style attribute name: {error}"
                    ))
                })?;
                let value = std::str::from_utf8(attribute.value.as_ref()).map_err(|error| {
                    Error::XmlError(format!(
                        "invalid UTF-8 in ODT style attribute value: {error}"
                    ))
                })?;
                style_element.try_set_attribute(key, value, "ODT style registry attribute")?;
            }
            self.registry
                .try_add_style(Style::from_element(style_element)?)?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::{CONTENT_ROOT, FAMILY_NAME, OpenParse};
    use crate::elements::style::StyleElements;
    use litchi_core::Result;
    use std::path::{Path, PathBuf};

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    fn document(body: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:style="{STYLE}" xmlns:text="{TEXT}" office:version="1.4">{body}</office:document-content>"#
        )
    }

    /// Deterministic behavioral projection: `StyleRegistry` is not
    /// `PartialEq` and its `Debug` order follows the hash map, so compare
    /// sorted style names with their extracted attributes and properties.
    fn project(registry: &crate::elements::style::StyleRegistry) -> String {
        let mut names: Vec<&str> = registry.styles.keys().map(String::as_str).collect();
        names.sort_unstable();
        let mut projected = String::new();
        for name in names {
            let style = &registry.styles[name];
            projected.push_str(&format!(
                "{name}|{:?}|{:?}|{}\n",
                style.family(),
                style.parent_style_name(),
                {
                    let properties = style.properties();
                    format!(
                        "{:?}|{:?}|{:?}|{:?}",
                        properties.text, properties.paragraph, properties.table, properties.graphic
                    )
                }
            ));
        }
        projected
    }

    /// The historical sequential passes with their exact error precedence:
    /// content validation first, then the automatic-styles scan.
    fn sequential(xml: &str) -> Result<String> {
        litchi_odf_common::core::validate_content_document_part(xml, CONTENT_ROOT, FAMILY_NAME)?;
        let registry = StyleElements::parse_styles(xml)?;
        Ok(project(&registry))
    }

    /// The fused open parse with the same comparison projection.
    fn fused(xml: &str) -> Result<String> {
        let registry = OpenParse::run(xml)?.finish()?;
        Ok(project(&registry))
    }

    /// Pre-0221 oracle: the buffered, fully-resolved fused open-parse loop,
    /// retained verbatim to cross-check the borrowing, depth-gated
    /// [`OpenParse::run`] on the synthetic edge cases and the fixture corpus.
    /// Since change 0224 it doubles as the `NsReader` differential oracle for
    /// the plain-`Reader` + [`BindingTracker`] driving loop: it still performs
    /// quick-xml's own binding maintenance (and resolves every event), so any
    /// divergence in the tracker's push/pop/error-stream replication shows up
    /// here.
    fn run_buffered_oracle(content_xml: &str) -> Result<OpenParse> {
        use litchi_core::Error;
        use quick_xml::events::Event;
        use quick_xml::name::{Namespace, ResolveResult};
        use quick_xml::reader::NsReader;

        if content_xml.len() > super::MAX_CONTENT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "{FAMILY_NAME} content.xml exceeds the family limit"
            )));
        }
        let mut validate = super::ValidateHandler::new()?;
        let mut styles = super::StyleHandler::default();
        let mut style_error = None;

        let mut reader = NsReader::from_str(content_xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = match reader.read_resolved_event_into(&mut buffer) {
                Ok(resolved) => resolved,
                Err(error) => {
                    return Err(Error::InvalidFormat(format!(
                        "invalid {FAMILY_NAME} content.xml: {error}"
                    )));
                },
            };
            let office = matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == super::OFFICE_NAMESPACE);
            let is_eof = matches!(event, Event::Eof);
            validate.on_event(office, &event)?;
            if style_error.is_none()
                && let Err(error) = styles.on_event(&event)
            {
                style_error = Some(error);
            }
            buffer.clear();
            if is_eof {
                break;
            }
        }
        validate.finish()?;

        Ok(OpenParse {
            registry: styles.registry,
            style_error,
        })
    }

    /// The pre-0221 fused loop with the same comparison projection.
    fn fused_oracle(xml: &str) -> Result<String> {
        let registry = run_buffered_oracle(xml)?.finish()?;
        Ok(project(&registry))
    }

    fn assert_gated_equivalent(label: &str, xml: &str) {
        let expected = fused_oracle(xml).map_err(|error| error.to_string());
        let actual = fused(xml).map_err(|error| error.to_string());
        assert_eq!(
            expected, actual,
            "{label}: gated and buffered-resolved fused runs disagree"
        );
    }

    fn assert_equivalent(label: &str, xml: &str) {
        let expected = sequential(xml).map_err(|error| error.to_string());
        let actual = fused(xml).map_err(|error| error.to_string());
        assert_eq!(expected, actual, "{label}: fused and sequential disagree");
    }

    fn corpus_files() -> Vec<PathBuf> {
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let mut files = Vec::new();
        collect_odt(&root.join("test-data"), &mut files);
        files.sort();
        files
    }

    fn collect_odt(directory: &Path, files: &mut Vec<PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_odt(&path, files);
            } else if path
                .extension()
                .is_some_and(|extension| extension == "odt" || extension == "fodt")
            {
                files.push(path);
            }
        }
    }

    fn content_xml(path: &Path) -> Option<String> {
        let bytes = std::fs::read(path).ok()?;
        if path
            .extension()
            .is_some_and(|extension| extension == "fodt")
        {
            return String::from_utf8(bytes).ok();
        }
        let reader = soapberry_zip::office::ArchiveReader::new(&bytes).ok()?;
        let entry = reader.read("content.xml").ok()?;
        String::from_utf8(entry).ok()
    }

    #[test]
    fn fused_parse_matches_sequential_passes_on_odt_corpus() {
        let files = corpus_files();
        assert!(!files.is_empty(), "no .odt corpus fixtures discovered");
        let mut compared = 0usize;
        for path in &files {
            let Some(xml) = content_xml(path) else {
                continue;
            };
            assert_equivalent(&path.display().to_string(), &xml);
            compared += 1;
        }
        assert!(compared > 0, "no .odt corpus fixtures yielded content.xml");
    }

    #[test]
    fn synthetic_documents_match_the_sequential_passes() {
        let body_with_style = concat!(
            r#"<office:automatic-styles><style:style style:name="Body" style:family="paragraph"></style:style></office:automatic-styles>"#,
            r#"<office:body><office:text><text:p>hi</text:p></office:text></office:body>"#,
        );
        let cases: Vec<(&str, String)> = vec![
            ("minimal-with-style", document(body_with_style)),
            // Attribute values stay raw and undecoded in both passes.
            (
                "entity-in-style-name",
                document(concat!(
                    r#"<office:automatic-styles><style:style style:name="a&amp;b" style:family="paragraph"></style:style></office:automatic-styles>"#,
                    r#"<office:body><office:text/></office:body>"#,
                )),
            ),
            (
                "forms-before-family-body",
                document(r#"<office:body><office:forms/><office:text/></office:body>"#),
            ),
            (
                "wrong-root",
                document("").replace("document-content", "document"),
            ),
            (
                "incomplete-body",
                format!(
                    r#"<office:document-content xmlns:office="{OFFICE}"></office:document-content>"#
                ),
            ),
            (
                "empty-root",
                format!(r#"<office:document-content xmlns:office="{OFFICE}"/>"#),
            ),
            (
                "wrong-family-body",
                document(r#"<office:body><office:spreadsheet/></office:body>"#),
            ),
            (
                "duplicate-body",
                document(r#"<office:body><office:text/></office:body><office:body/>"#),
            ),
            (
                "doctype",
                format!(
                    r#"<!DOCTYPE office:document-content>{}"#,
                    document(r#"<office:body><office:text/></office:body>"#)
                ),
            ),
            (
                "trailing-text",
                format!(
                    "{}tail",
                    document(r#"<office:body><office:text/></office:body>"#)
                ),
            ),
            (
                "content-after-root",
                format!(
                    "{}<office:document-content/>",
                    document(r#"<office:body><office:text/></office:body>"#)
                ),
            ),
            (
                "late-declaration",
                document(r#"<?xml version="1.1"?><office:body><office:text/></office:body>"#),
            ),
            (
                "invalid-character-reference",
                document(
                    r#"<office:body><office:text><text:p>&#xD800;</text:p></office:text></office:body>"#,
                ),
            ),
            (
                "mismatched-end-tag",
                format!(
                    r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:text></office:body></office:document-content>"#
                ),
            ),
            (
                "duplicate-style-attribute",
                document(concat!(
                    r#"<office:automatic-styles><style:style style:name="a" style:name="b"></style:style></office:automatic-styles>"#,
                    r#"<office:body><office:text/></office:body>"#,
                )),
            ),
            // A style error recorded before a later tokenization failure
            // still loses to the validation-side read error, matching the
            // historical order in which validation tokenized first.
            (
                "style-error-before-malformed-xml",
                format!(
                    r#"<office:document-content xmlns:office="{OFFICE}" xmlns:style="{STYLE}"><office:automatic-styles><style:style style:name="a" style:name="b"></style:style></office:automatic-styles><office:body><office:text></office:body></office:document-content>"#
                ),
            ),
        ];
        for (label, xml) in &cases {
            assert_equivalent(label, xml);
        }
    }

    #[test]
    fn style_matching_is_literal_and_prefix_sensitive() {
        // A custom prefix bound to the style namespace is NOT collected: the
        // historical scan matches the raw bytes `style:style`.
        let custom_prefix = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:s="{STYLE}"><office:automatic-styles><s:style s:name="Ghost" s:family="paragraph"></s:style></office:automatic-styles><office:body><office:text/></office:body></office:document-content>"#
        );
        assert_equivalent("custom-style-prefix-not-collected", &custom_prefix);
        let registry = OpenParse::run(&custom_prefix)
            .and_then(OpenParse::finish)
            .expect("custom-prefix document parses");
        assert!(
            registry.get_style("Ghost").is_none(),
            "literal byte matching must not collect a custom prefix"
        );

        // The canonical prefix bound to a FOREIGN namespace IS collected:
        // the match is on bytes, not on the resolved namespace.
        let foreign_binding = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:style="urn:foreign"><office:automatic-styles><style:style style:name="Kept" style:family="paragraph"></style:style></office:automatic-styles><office:body><office:text/></office:body></office:document-content>"#
        );
        assert_equivalent("canonical-prefix-foreign-uri-collected", &foreign_binding);
        let registry = OpenParse::run(&foreign_binding)
            .and_then(OpenParse::finish)
            .expect("foreign-binding document parses");
        assert!(
            registry.get_style("Kept").is_some(),
            "literal byte matching must collect the canonical prefix"
        );
    }

    #[test]
    fn gated_run_matches_buffered_oracle_on_synthetic_edge_cases() {
        let minimal = document(r#"<office:body><office:text/></office:body>"#);
        let cases: Vec<(&str, String)> = vec![
            ("minimal", minimal.clone()),
            (
                "rebinding-on-body-hides-it",
                minimal.replace("<office:body>", r#"<office:body xmlns:office="urn:evil">"#),
            ),
            (
                "rebinding-on-empty-body-hides-it",
                document(r#"<office:body xmlns:office="urn:evil"/><office:text/>"#),
            ),
            (
                "rebinding-on-family-element-rejected",
                minimal.replace(
                    "<office:text/>",
                    r#"<office:text xmlns:office="urn:evil"/>"#,
                ),
            ),
            (
                "deep-rebinding-accepted",
                document(
                    r#"<office:body><office:text><text:p xmlns:office="urn:evil"><office:annotation/></text:p></office:text></office:body>"#,
                ),
            ),
            (
                "rebind-and-restore-nested",
                document(
                    r#"<office:body><office:text><text:p xmlns:office="urn:evil"><office:annotation/></text:p><text:p><office:annotation/></text:p></office:text></office:body>"#,
                ),
            ),
            (
                "aliased-prefix-everywhere",
                format!(
                    r#"<x:document-content xmlns:x="{OFFICE}"><x:body><x:text/></x:body></x:document-content>"#
                ),
            ),
            (
                "default-namespace-accepted",
                format!(
                    r#"<document-content xmlns="{OFFICE}"><body><text/></body></document-content>"#
                ),
            ),
            (
                "unknown-prefix-at-root",
                minimal.replace(
                    &format!(
                        r#" xmlns:office="{OFFICE}" xmlns:style="{STYLE}" xmlns:text="{TEXT}""#
                    ),
                    "",
                ),
            ),
            (
                "unknown-prefix-deep-accepted",
                document(
                    r#"<office:body><office:text><text:p><weird:thing/></text:p></office:text></office:body>"#,
                ),
            ),
            (
                "second-prefix-bound-to-office",
                format!(
                    r#"<office:document-content xmlns:office="{OFFICE}" xmlns:o2="{OFFICE}"><o2:body><o2:text/></o2:body></office:document-content>"#
                ),
            ),
            (
                "forms-empty-before-family",
                document(r#"<office:body><office:forms/><office:text/></office:body>"#),
            ),
            (
                "forms-start-end-before-family",
                document(
                    r#"<office:body><office:forms></office:forms><office:text/></office:body>"#,
                ),
            ),
            (
                "mismatched-end-tag-deep",
                document(
                    r#"<office:body><office:text><text:p></text:q></office:text></office:body>"#,
                ),
            ),
            (
                "mismatched-end-tag-at-body",
                format!(
                    r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:text></office:body></office:document-content>"#
                ),
            ),
            (
                "comment-pi-cdata-interleaved",
                format!(
                    r#"<?xml version="1.0"?><!--prologue--><?pi data?>{}"#,
                    document(
                        r#"<office:body><office:text><text:p><![CDATA[x]]><!--in--><?p i?>&amp;</text:p></office:text></office:body>"#,
                    )
                ),
            ),
            (
                "comment-and-pi-after-root-accepted",
                format!("{}<!--tail--><?pi data?>", minimal),
            ),
            (
                "double-hyphen-comment",
                document(
                    r#"<office:body><office:text><!-- a -- b --></office:text></office:body>"#,
                ),
            ),
            ("cdata-outside-root", format!("<![CDATA[x]]>{minimal}")),
            ("truncated-at-open-tag", format!("{}<text:p", minimal)),
            (
                "unclosed-body-at-eof",
                format!(r#"<office:document-content xmlns:office="{OFFICE}"><office:body>"#),
            ),
            (
                "truncated-inside-automatic-styles",
                format!(
                    r#"<office:document-content xmlns:office="{OFFICE}" xmlns:style="{STYLE}"><office:automatic-styles><style:style style:name="a""#
                ),
            ),
            ("missing-body", document("")),
            ("body-without-family", document(r#"<office:body/>"#)),
            (
                "wrong-root",
                minimal.replace("document-content", "document"),
            ),
            (
                "empty-root",
                format!(r#"<office:document-content xmlns:office="{OFFICE}"/>"#),
            ),
            (
                "style-collected-under-deep-rebinding",
                document(concat!(
                    r#"<office:automatic-styles><x xmlns:style="urn:evil"><style:style style:name="Deep" style:family="paragraph"/></x></office:automatic-styles>"#,
                    r#"<office:body><office:text/></office:body>"#,
                )),
            ),
            (
                "style-error-before-mismatched-end",
                format!(
                    r#"<office:document-content xmlns:office="{OFFICE}" xmlns:style="{STYLE}"><office:automatic-styles><style:style style:name="a" style:name="b"></style:style></office:automatic-styles><office:body><office:text><text:p></text:q></office:text></office:body></office:document-content>"#
                ),
            ),
        ];
        for (label, xml) in &cases {
            assert_gated_equivalent(label, xml);
        }
    }

    /// Change 0224 differential battery: every byte-exactness obligation of
    /// the [`BindingTracker`] contract (push scan, silent break, declaration
    /// limit, reserved-prefix errors in attribute order, unbinding, deferred
    /// pop) is exercised against the `NsReader` oracle, including errors
    /// raised by binding maintenance at depth >= 3.
    #[test]
    fn tracker_matches_nsreader_on_namespace_edge_cases() {
        const XML_URI: &str = "http://www.w3.org/XML/1998/namespace";
        const XMLNS_URI: &str = "http://www.w3.org/2000/xmlns/";
        let minimal = document(r#"<office:body><office:text/></office:body>"#);
        let declarations = |count: usize| {
            (0..count)
                .map(|index| format!(r#"xmlns:d{index}="urn:d""#))
                .collect::<Vec<_>>()
                .join(" ")
        };

        let cases: Vec<(String, String)> = vec![
            // The malformed-attribute scan break keeps bindings declared
            // before the break (root stays bound) and never reaches bindings
            // declared after it.
            (
                "malformed-attribute-after-xmlns-on-root".to_string(),
                minimal.replace(r#" office:version="1.4""#, " bad="),
            ),
            (
                "malformed-attribute-before-xmlns-on-root".to_string(),
                format!(
                    r#"<office:document-content broken= xmlns:office="{OFFICE}"><office:body><office:text/></office:body></office:document-content>"#
                ),
            ),
            (
                "empty-attribute-key-breaks-scan".to_string(),
                minimal.replace(&format!(r#"xmlns:text="{TEXT}""#), r#"="x""#),
            ),
            // The 256-declaration limit fires identically at any depth.
            (
                "declarations-at-limit-accepted".to_string(),
                document(&format!(
                    r#"<office:body><office:text><text:p {}>x</text:p></office:text></office:body>"#,
                    declarations(256)
                )),
            ),
            (
                "declarations-over-limit-deep".to_string(),
                document(&format!(
                    r#"<office:body><office:text><text:p {}>x</text:p></office:text></office:body>"#,
                    declarations(257)
                )),
            ),
            (
                "declarations-over-limit-on-root".to_string(),
                minimal.replace(
                    r#"<office:document-content"#,
                    &format!("<office:document-content {}", declarations(257)),
                ),
            ),
            // Reserved-prefix rules, each firing on the root and deep inside
            // the body (depth >= 3 maintenance must not be gated).
            (
                "xml-prefix-foreign-uri-on-root".to_string(),
                minimal.replace(
                    &format!(r#"xmlns:text="{TEXT}""#),
                    r#"xmlns:xml="urn:wrong""#,
                ),
            ),
            (
                "xml-prefix-foreign-uri-deep".to_string(),
                document(
                    r#"<office:body><office:text><text:p xmlns:xml="urn:wrong">x</text:p></office:text></office:body>"#,
                ),
            ),
            (
                "xml-prefix-reserved-uri-accepted".to_string(),
                minimal.replace(
                    &format!(r#"xmlns:text="{TEXT}""#),
                    &format!(r#"xmlns:xml="{XML_URI}""#),
                ),
            ),
            (
                "xmlns-prefix-declared-on-root".to_string(),
                minimal.replace(&format!(r#"xmlns:text="{TEXT}""#), r#"xmlns:xmlns="urn:x""#),
            ),
            (
                "prefix-bound-to-xml-uri".to_string(),
                minimal.replace(
                    &format!(r#"xmlns:text="{TEXT}""#),
                    &format!(r#"xmlns:foo="{XML_URI}""#),
                ),
            ),
            (
                "prefix-bound-to-xmlns-uri-deep".to_string(),
                document(&format!(
                    r#"<office:body><office:text><text:p xmlns:foo="{XMLNS_URI}">x</text:p></office:text></office:body>"#
                )),
            ),
            // The first failing declaration wins, in attribute order.
            (
                "reserved-error-ordering-xmlns-uri-first".to_string(),
                minimal.replace(
                    &format!(r#"xmlns:text="{TEXT}""#),
                    &format!(r#"xmlns:foo="{XMLNS_URI}" xmlns:xml="urn:wrong""#),
                ),
            ),
            (
                "reserved-error-ordering-xml-first".to_string(),
                minimal.replace(
                    &format!(r#"xmlns:text="{TEXT}""#),
                    &format!(r#"xmlns:xml="urn:wrong" xmlns:foo="{XMLNS_URI}""#),
                ),
            ),
            // Unbinding: an emptied prefix resolves `Unknown` (rejects at the
            // consumed depth), an emptied default resolves `Unbound`.
            (
                "prefix-unbound-on-body-rejected".to_string(),
                minimal.replace("<office:body>", r#"<office:body xmlns:office="">"#),
            ),
            (
                "prefix-unbound-deep-accepted".to_string(),
                document(
                    r#"<office:body><office:text><text:p xmlns:text=""><text:span/></text:p></office:text></office:body>"#,
                ),
            ),
            (
                "default-namespace-unbound-at-body".to_string(),
                format!(
                    r#"<document-content xmlns="{OFFICE}"><body xmlns=""><text/></body></document-content>"#
                ),
            ),
            (
                "default-namespace-rebound-after-unbinding".to_string(),
                format!(
                    r#"<document-content xmlns="{OFFICE}"><body xmlns=""><x xmlns="{OFFICE}"></x><text/></body></document-content>"#
                ),
            ),
            // Duplicate declarations: no duplicate check (`with_checks(false)`),
            // the last binding wins on resolution.
            (
                "duplicate-xmlns-last-wins-accepted".to_string(),
                minimal.replace(
                    &format!(r#"xmlns:office="{OFFICE}""#),
                    &format!(r#"xmlns:office="urn:wrong" xmlns:office="{OFFICE}""#),
                ),
            ),
            (
                "duplicate-xmlns-last-wins-rejected".to_string(),
                minimal.replace(
                    &format!(r#"xmlns:office="{OFFICE}""#),
                    &format!(r#"xmlns:office="{OFFICE}" xmlns:office="urn:wrong""#),
                ),
            ),
            // Declaration values stay raw and undecoded: an entity reference
            // in the URI is never expanded, so the binding cannot match.
            (
                "entity-in-xmlns-value-stays-raw".to_string(),
                minimal.replace(
                    &format!(r#"xmlns:office="{OFFICE}""#),
                    &format!(r#"xmlns:office="{OFFICE}&amp;""#),
                ),
            ),
            // Single quotes and whitespace variants around declarations.
            (
                "single-quoted-xmlns-value".to_string(),
                minimal.replace(
                    &format!(r#"xmlns:office="{OFFICE}""#),
                    &format!("xmlns:office='{OFFICE}'"),
                ),
            ),
            (
                "spaced-xmlns-declaration".to_string(),
                minimal.replace(
                    &format!(r#"xmlns:office="{OFFICE}""#),
                    &format!(r#"xmlns:office = "{OFFICE}""#),
                ),
            ),
            // Element names in the `xmlns` series are not declarations; the
            // prefilter must fall through to the full scan harmlessly.
            (
                "xmlns-named-element-deep".to_string(),
                document(
                    r#"<office:body><office:text><text:p><xmlnsfoo xmlns2="v"/></text:p></office:text></office:body>"#,
                ),
            ),
        ];
        // Asserting only cross-implementation agreement would let a toothless
        // case pass (e.g. if an "error" input were silently accepted by both
        // sides); pin the expected outcome of every case as well.
        let expected_outcomes: &[(&str, bool)] = &[
            ("malformed-attribute-after-xmlns-on-root", true),
            ("malformed-attribute-before-xmlns-on-root", false),
            ("empty-attribute-key-breaks-scan", true),
            ("declarations-at-limit-accepted", true),
            ("declarations-over-limit-deep", false),
            ("declarations-over-limit-on-root", false),
            ("xml-prefix-foreign-uri-on-root", false),
            ("xml-prefix-foreign-uri-deep", false),
            ("xml-prefix-reserved-uri-accepted", true),
            ("xmlns-prefix-declared-on-root", false),
            ("prefix-bound-to-xml-uri", false),
            ("prefix-bound-to-xmlns-uri-deep", false),
            ("reserved-error-ordering-xmlns-uri-first", false),
            ("reserved-error-ordering-xml-first", false),
            ("prefix-unbound-on-body-rejected", false),
            ("prefix-unbound-deep-accepted", true),
            ("default-namespace-unbound-at-body", false),
            ("default-namespace-rebound-after-unbinding", false),
            ("duplicate-xmlns-last-wins-accepted", true),
            ("duplicate-xmlns-last-wins-rejected", false),
            ("entity-in-xmlns-value-stays-raw", false),
            ("single-quoted-xmlns-value", true),
            ("spaced-xmlns-declaration", true),
            ("xmlns-named-element-deep", true),
        ];
        assert_eq!(
            cases.len(),
            expected_outcomes.len(),
            "every battery case needs a pinned outcome"
        );
        for ((label, xml), (_, expect_ok)) in cases.iter().zip(expected_outcomes) {
            assert_gated_equivalent(label, xml);
            assert_eq!(
                fused(xml).is_ok(),
                *expect_ok,
                "{label}: unexpected accept/reject outcome"
            );
        }
    }

    /// Per-event element-name resolutions from the new `Reader` + tracker
    /// driving loop, recorded for every `Start`/`Empty`/`End` event at every
    /// depth (not only the consumed depth <= 2 ones).
    fn tracker_resolutions(xml: &str) -> std::result::Result<Vec<String>, String> {
        use quick_xml::reader::Reader;

        let mut reader = Reader::from_str(xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        let mut tracker = crate::binding_tracker::BindingTracker::new();
        let mut pending_pop = false;
        let mut resolutions = Vec::new();
        loop {
            if pending_pop {
                tracker.pop();
                pending_pop = false;
            }
            let event = reader
                .read_event()
                .map_err(|error| format!("invalid ODT content.xml: {error}"))?;
            match &event {
                quick_xml::events::Event::Start(element) => {
                    if let Err(error) = tracker.push(element) {
                        return Err(format!("invalid ODT content.xml: {error}"));
                    }
                    resolutions.push(format!(
                        "start:{:?}",
                        tracker.resolve_element(element.name()).0
                    ));
                },
                quick_xml::events::Event::Empty(element) => {
                    if let Err(error) = tracker.push(element) {
                        return Err(format!("invalid ODT content.xml: {error}"));
                    }
                    resolutions.push(format!(
                        "empty:{:?}",
                        tracker.resolve_element(element.name()).0
                    ));
                    pending_pop = true;
                },
                quick_xml::events::Event::End(element) => {
                    // The scope pop is deferred to the next read, so the end
                    // tag still resolves in its own scope.
                    resolutions.push(format!(
                        "end:{:?}",
                        tracker.resolve_element(element.name()).0
                    ));
                    pending_pop = true;
                },
                quick_xml::events::Event::Eof => break,
                _ => {},
            }
        }
        Ok(resolutions)
    }

    /// The same resolutions from `NsReader`'s fully-resolved event stream
    /// (`read_resolved_event_into` resolves every `Start`/`Empty`/`End` name).
    fn nsreader_resolutions(xml: &str) -> std::result::Result<Vec<String>, String> {
        use quick_xml::reader::NsReader;

        let mut reader = NsReader::from_str(xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        let mut buffer = Vec::new();
        let mut resolutions = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| format!("invalid ODT content.xml: {error}"))?;
            match &event {
                quick_xml::events::Event::Start(_) => {
                    resolutions.push(format!("start:{namespace:?}"));
                },
                quick_xml::events::Event::Empty(_) => {
                    resolutions.push(format!("empty:{namespace:?}"));
                },
                quick_xml::events::Event::End(_) => {
                    resolutions.push(format!("end:{namespace:?}"));
                },
                quick_xml::events::Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }
        Ok(resolutions)
    }

    fn assert_resolutions_equivalent(label: &str, xml: &str) {
        assert_eq!(
            tracker_resolutions(xml),
            nsreader_resolutions(xml),
            "{label}: tracker and NsReader per-event resolutions disagree"
        );
    }

    /// Resolution fidelity at every depth: rebinding, unbinding, reserved
    /// prefixes, default namespaces, and error preemption, plus the whole
    /// `BindingTracker`-relevant synthetic battery.
    #[test]
    fn tracker_matches_nsreader_resolutions_on_synthetic_cases() {
        let cases: Vec<(&str, String)> = vec![
            (
                "minimal",
                document(r#"<office:body><office:text/></office:body>"#),
            ),
            (
                "rebinding-nested",
                document(
                    r#"<office:body><office:text><text:p xmlns:office="urn:evil"><office:annotation/></text:p><text:p><office:annotation/></text:p></office:text></office:body>"#,
                ),
            ),
            (
                "rebinding-on-empty",
                document(r#"<office:body xmlns:office="urn:evil"/><office:text/>"#),
            ),
            (
                "unbinding-and-restore",
                document(
                    r#"<office:body><office:text><text:p xmlns:text=""><text:span/><x xmlns:text="urn:other"><text:span/></x></text:p></office:text></office:body>"#,
                ),
            ),
            (
                "default-namespace-mixed-with-prefixed",
                format!(
                    r#"<document-content xmlns="{OFFICE}" xmlns:office="urn:other"><body><text/><office:note/></body></document-content>"#
                ),
            ),
            (
                "reserved-xml-prefix-element-names",
                document(
                    r#"<office:body><office:text><text:p xml:space="preserve"><xml:thing/><xmlns:other/></text:p></office:text></office:body>"#,
                ),
            ),
            (
                "unknown-prefixes-at-all-depths",
                document(
                    r#"<office:body><office:text><text:p><a:b><c:d/></a:b></text:p></office:text></office:body>"#,
                ),
            ),
            (
                "duplicate-declarations-last-wins",
                document(
                    r#"<office:body><office:text><text:p xmlns:text="urn:first" xmlns:text="urn:second"><text:span/></text:p></office:text></office:body>"#,
                ),
            ),
            (
                "error-stops-stream",
                document(
                    r#"<office:body><office:text><text:p xmlns:xml="urn:wrong">x</text:p><text:p/></office:text></office:body>"#,
                ),
            ),
        ];
        for (label, xml) in &cases {
            assert_resolutions_equivalent(label, xml);
        }
    }

    #[test]
    fn tracker_matches_nsreader_resolutions_on_odt_corpus() {
        let files = corpus_files();
        assert!(!files.is_empty(), "no .odt corpus fixtures discovered");
        let mut compared = 0usize;
        for path in &files {
            let Some(xml) = content_xml(path) else {
                continue;
            };
            assert_resolutions_equivalent(&path.display().to_string(), &xml);
            compared += 1;
        }
        assert!(compared > 0, "no .odt corpus fixtures yielded content.xml");
    }

    #[test]
    fn gated_run_matches_buffered_oracle_on_odt_corpus() {
        let files = corpus_files();
        assert!(!files.is_empty(), "no .odt corpus fixtures discovered");
        let mut compared = 0usize;
        for path in &files {
            let Some(xml) = content_xml(path) else {
                continue;
            };
            assert_gated_equivalent(&path.display().to_string(), &xml);
            compared += 1;
        }
        assert!(compared > 0, "no .odt corpus fixtures yielded content.xml");
    }
}
