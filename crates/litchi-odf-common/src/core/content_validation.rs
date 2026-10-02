//! Shared state machine for validating an ODF `content.xml` document root.

use litchi_core::{Error, Result};
use quick_xml::events::Event;

const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;
const MAX_CONTENT_DEPTH: usize = 4096;

/// Streaming validator for an ODF document root, body, and family element.
///
/// The caller supplies the resolved-office-namespace result for each
/// `Start`/`Empty` event.  [`Self::needs_office_namespace`] indicates when
/// that resolution is observable; namespace binding maintenance must still
/// run for every element so deep namespace errors retain quick-xml's order.
#[derive(Debug)]
pub struct ContentDocumentValidator {
    family_name: String,
    expected_local: String,
    depth: usize,
    root_closed: bool,
    body_seen: bool,
    expected_seen: bool,
    in_body: bool,
    declaration_seen: bool,
    first_event: bool,
}

impl ContentDocumentValidator {
    /// Reject a declared materialized `content.xml` size above the common
    /// family limit before the package entry is read into memory.
    ///
    /// The size is optional because callers may be validating detached XML
    /// without package metadata.  Package-backed callers should pass the
    /// value returned by `SourceBackedPackage::member_materialized_size`,
    /// which represents plaintext bytes for encrypted entries.
    pub fn check_materialized_size(size: Option<u64>, family_name: &str) -> Result<()> {
        if size.is_some_and(|size| size > MAX_CONTENT_BYTES as u64) {
            return Err(Error::InvalidFormat(format!(
                "{family_name} content.xml exceeds the family limit"
            )));
        }
        Ok(())
    }

    /// Create a validator and apply the bounded `content.xml` size check.
    ///
    /// # Errors
    ///
    /// Returns the same family error as the historical standalone validator
    /// when the XML exceeds 256 MiB or the body marker is not an
    /// `office:`-qualified element marker.
    pub fn new(content_xml: &str, body_marker: &str, family_name: &str) -> Result<Self> {
        Self::check_materialized_size(Some(content_xml.len() as u64), family_name)?;
        let expected_local = body_marker
            .strip_prefix("<office:")
            .and_then(|marker| {
                marker
                    .split(|character: char| !character.is_ascii_alphanumeric() && character != '-')
                    .next()
            })
            .filter(|local| !local.is_empty())
            .ok_or_else(|| {
                Error::InvalidFormat(format!("{family_name} content.xml has no expected body"))
            })?;
        Ok(Self {
            family_name: family_name.to_owned(),
            expected_local: expected_local.to_owned(),
            depth: 0,
            root_closed: false,
            body_seen: false,
            expected_seen: false,
            in_body: false,
            declaration_seen: false,
            first_event: true,
        })
    }

    /// Return whether the current event's office namespace is observable.
    ///
    /// The standalone validator only consumes resolved names for the root,
    /// `office:body`, and the first-level body children.  Deep elements still
    /// require tracker maintenance, but their resolved namespace is not read.
    #[must_use]
    pub fn needs_office_namespace(&self) -> bool {
        self.depth <= 2
    }

    /// Consume one tokenized event using its resolved office-namespace flag.
    ///
    /// The flag is used only for `Start` and `Empty` events at the depth
    /// reported by [`Self::needs_office_namespace`].  All structural checks,
    /// error strings, and event ordering match the historical validator.
    pub fn on_event(&mut self, office: bool, event: &Event<'_>) -> Result<()> {
        let family_name = &self.family_name;
        match event {
            Event::Start(element) => {
                if self.root_closed {
                    return Err(Error::InvalidFormat(format!(
                        "{family_name} content.xml has content after its root"
                    )));
                }
                if self.depth >= MAX_CONTENT_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "{family_name} content.xml nesting exceeds maximum depth of {MAX_CONTENT_DEPTH}"
                    )));
                }
                let local = element.local_name();
                match self.depth {
                    0 if office && local.as_ref() == b"document-content" => self.depth = 1,
                    0 => {
                        return Err(Error::InvalidFormat(format!(
                            "{family_name} content.xml has the wrong root"
                        )));
                    },
                    1 if office && local.as_ref() == b"body" && !self.body_seen => {
                        self.body_seen = true;
                        self.in_body = true;
                        self.depth = 2;
                    },
                    1 if office && local.as_ref() == b"body" => {
                        return Err(Error::InvalidFormat(format!(
                            "{family_name} content.xml has duplicate office:body"
                        )));
                    },
                    1 => {
                        self.depth = self.depth.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "{family_name} content.xml nesting overflows"
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
                            "{family_name} content.xml has the wrong office body"
                        )));
                    },
                    _ => {
                        self.depth = self.depth.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "{family_name} content.xml nesting overflows"
                            ))
                        })?;
                    },
                }
            },
            Event::Empty(element) => {
                if self.root_closed || self.depth == 0 {
                    return Err(Error::InvalidFormat(format!(
                        "{family_name} content.xml has an invalid empty root"
                    )));
                }
                let local = element.local_name();
                if self.depth == 1 {
                    if office && local.as_ref() == b"body" && !self.body_seen {
                        self.body_seen = true;
                    } else if office && local.as_ref() == b"body" {
                        return Err(Error::InvalidFormat(format!(
                            "{family_name} content.xml has duplicate office:body"
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
                        "{family_name} content.xml has the wrong office body"
                    )));
                }
            },
            Event::End(_) => {
                self.depth = self.depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat(format!("{family_name} content.xml has an unexpected end"))
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
                        "{family_name} content.xml has unexpected text outside its root"
                    )));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if self.depth == 0 || self.root_closed => {
                return Err(Error::InvalidFormat(format!(
                    "{family_name} content.xml has content outside its root"
                )));
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(format!(
                    "{family_name} content.xml must not contain a doctype"
                )));
            },
            Event::GeneralRef(reference) if !crate::validation::valid_xml_reference(reference) => {
                return Err(Error::InvalidFormat(format!(
                    "{family_name} content.xml has an invalid character or entity reference"
                )));
            },
            Event::Decl(_)
                if self.declaration_seen
                    || !self.first_event
                    || self.depth != 0
                    || self.root_closed =>
            {
                return Err(Error::InvalidFormat(format!(
                    "{family_name} content.xml has an XML declaration outside its prologue"
                )));
            },
            Event::Decl(_) => self.declaration_seen = true,
            Event::Eof => {},
            Event::Comment(_) | Event::PI(_) | Event::CData(_) | Event::GeneralRef(_) => {},
        }
        self.first_event = false;
        Ok(())
    }

    /// Finish validation after the reader has delivered `Event::Eof`.
    pub fn finish(self) -> Result<()> {
        if !self.root_closed || self.depth != 0 || !self.body_seen || !self.expected_seen {
            return Err(Error::InvalidFormat(format!(
                "{} content.xml has no complete expected body",
                self.family_name
            )));
        }
        Ok(())
    }
}

/// Byte span of a document type declaration in the prologue of an XML part,
/// or `None` when the prologue has none.
///
/// Only the prologue is tokenized: scanning stops at the root element. A
/// prologue that does not tokenize is left alone for the part's validator to
/// reject with its own error.
fn prologue_doctype_span(xml: &[u8]) -> Option<(usize, usize)> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut buffer = Vec::new();
    loop {
        let before = usize::try_from(reader.buffer_position()).ok()?;
        let event = reader.read_event_into(&mut buffer).ok()?;
        match event {
            Event::DocType(_) => {
                let after = usize::try_from(reader.buffer_position()).ok()?;
                let span = xml.get(before..after)?;
                // Defensive: only a complete `<!...>` declaration is replaced.
                return (span.starts_with(b"<!") && span.ends_with(b">"))
                    .then_some((before, after));
            },
            Event::Start(_) | Event::Empty(_) | Event::End(_) | Event::Eof => return None,
            Event::Decl(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
}

/// Neutralize a document type declaration in the prologue of an ODF
/// `content.xml` (LENIENCY_AUDIT F9.11), returning whether one was found.
///
/// ODF 1.2 documents are not DTD-validated and XML lets a non-validating
/// processor skip the declaration; LibreOffice's SAX import ignores it, so
/// old generator pipelines (DocBook-to-ODF stylesheets, OpenOffice.org 1.x
/// style prologues) open there. The declaration's bytes are overwritten with
/// the same number of spaces — legal prologue whitespace — so every later
/// byte offset is unchanged and no downstream parser ever sees a DTD.
///
/// The declaration is never processed: no external subset is fetched and no
/// internal-subset entity is declared, so a reference to an entity it would
/// have declared is still rejected by content validation. A declaration
/// after the root element is not touched and stays an error.
pub fn neutralize_prologue_doctype_bytes(xml: &mut [u8]) -> bool {
    let Some((start, end)) = prologue_doctype_span(xml) else {
        return false;
    };
    if let Some(span) = xml.get_mut(start..end) {
        span.fill(b' ');
        true
    } else {
        false
    }
}

/// [`neutralize_prologue_doctype_bytes`] for an already decoded part.
pub fn neutralize_prologue_doctype(xml: &mut String) -> bool {
    let Some((start, end)) = prologue_doctype_span(xml.as_bytes()) else {
        return false;
    };
    // The span starts at `<` and ends after `>`, both ASCII, so both ends
    // are character boundaries.
    if xml.is_char_boundary(start) && xml.is_char_boundary(end) {
        xml.replace_range(start..end, &" ".repeat(end - start));
        true
    } else {
        false
    }
}

#[cfg(test)]
mod doctype_tests {
    use super::{neutralize_prologue_doctype, neutralize_prologue_doctype_bytes};

    const ROOT: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text/></office:body></office:document-content>"#;

    #[test]
    fn a_prologue_doctype_is_blanked_in_place() {
        let doctype = r#"<!DOCTYPE office:document-content [ <!ENTITY e "a > b"> <!-- ] > --> ]>"#;
        let original = format!("<?xml version=\"1.0\"?>\n<!-- lead -->{doctype}{ROOT}");
        let mut xml = original.clone();
        assert!(neutralize_prologue_doctype(&mut xml));
        assert_eq!(xml.len(), original.len(), "byte offsets are preserved");
        assert!(!xml.contains("DOCTYPE"));
        assert!(xml.ends_with(ROOT));
        assert_eq!(
            xml.replace(' ', ""),
            original.replace(doctype, "").replace(' ', "")
        );

        let mut bytes = original.clone().into_bytes();
        assert!(neutralize_prologue_doctype_bytes(&mut bytes));
        assert_eq!(bytes, xml.into_bytes());
    }

    #[test]
    fn documents_without_a_prologue_doctype_are_untouched() {
        for source in [
            ROOT.to_string(),
            format!("<?xml version=\"1.0\"?>{ROOT}"),
            // A DOCTYPE after the root element is not a prologue declaration.
            format!("{ROOT}<!DOCTYPE late>"),
            // Unreadable XML is left for the part's validator to reject.
            "<?xml version=\"1.0\"?><!DOCTYPE unterminated".to_string(),
        ] {
            let mut xml = source.clone();
            assert!(!neutralize_prologue_doctype(&mut xml), "{source}");
            assert_eq!(xml, source);
        }
    }
}
