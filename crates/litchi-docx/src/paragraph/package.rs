//! Package relationship integration for `WordprocessingML` paragraphs.

use crate::error::Result;
use crate::hyperlink::Hyperlink;
use litchi_opc::rel::Relationships;
use quick_xml::events::Event;
use quick_xml::{Reader, XmlVersion};

use super::model::{Inline, InlineHyperlink, Paragraph, Run};

impl Paragraph {
    /// Return ordered direct paragraph children with hyperlinks resolved.
    ///
    /// Unlike [`Self::inlines`], direct `<w:hyperlink>` children become typed
    /// [`Inline::Hyperlink`] values. The relationship identifier remains
    /// package-internal; callers receive only the resolved link target and the
    /// hyperlink's typed direct runs. All other unknown children retain their
    /// exact fallback representation.
    ///
    /// # Errors
    ///
    /// Returns an error when paragraph XML is malformed or a hyperlink cannot
    /// be resolved safely.
    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "non-exhaustive public Inline values must pass through unchanged"
    )]
    pub fn inlines_with_relationships(&self, rels: &Relationships) -> Result<Vec<Inline>> {
        self.inlines()?
            .into_iter()
            .map(|candidate| match candidate {
                Inline::Unknown(opaque) if opaque.is_word_hyperlink() => Ok(Inline::Hyperlink(
                    Box::new(parse_hyperlink(opaque.xml_bytes(), rels)?),
                )),
                other => Ok(other),
            })
            .collect()
    }

    /// Get all hyperlinks in this paragraph.
    ///
    /// Returns a vector of `Hyperlink` objects representing all hyperlinks
    /// found in this paragraph. Requires relationships to resolve external URLs.
    ///
    /// # Arguments
    ///
    /// * `rels` - Relationships for resolving relationship IDs to URLs
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let para = doc.paragraph(0)?.unwrap();
    /// let hyperlinks = para.hyperlinks(&main_part.rels())?;
    /// for link in hyperlinks {
    ///     println!("Link: {} -> {:?}", link.text(), link.url());
    /// }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error when the paragraph XML or hyperlink attributes are
    /// malformed.
    pub fn hyperlinks(&self, rels: &Relationships) -> Result<Vec<Hyperlink>> {
        Hyperlink::extract_from_paragraph(self.xml_bytes(), rels)
    }
}

fn parse_hyperlink(xml: &[u8], rels: &Relationships) -> Result<InlineHyperlink> {
    let mut links = Hyperlink::extract_from_paragraph(xml, rels)?;
    if links.len() != 1 {
        return Err(crate::Error::InvalidFormat(format!(
            "direct hyperlink child resolved to {} hyperlinks",
            links.len()
        )));
    }
    let link = links.remove(0);
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut capture = None::<(usize, usize)>;
    let mut runs = Vec::new();
    let mut target_frame = None;
    let mut document_location = None;
    let mut has_unmodeled_content = false;
    let mut saw_root = false;

    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|_conversion_error| {
            crate::Error::InvalidFormat("hyperlink child offset does not fit usize".into())
        })?;
        let event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
            .into_owned();
        let end = usize::try_from(reader.buffer_position()).map_err(|_conversion_error| {
            crate::Error::InvalidFormat("hyperlink child offset does not fit usize".into())
        })?;
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("hyperlink nesting is too deep".into())
                })?;
                if depth == 1 {
                    if saw_root || element.local_name().as_ref() != b"hyperlink" {
                        return Err(crate::Error::InvalidFormat(
                            "invalid direct hyperlink fragment root".into(),
                        ));
                    }
                    saw_root = true;
                    for attribute_result in element.attributes() {
                        let attribute = attribute_result
                            .map_err(|error| crate::Error::Xml(error.to_string()))?;
                        let value = attribute
                            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                            .map_err(|error| crate::Error::Xml(error.to_string()))?
                            .into_owned();
                        match attribute.key.local_name().as_ref() {
                            b"tgtFrame" => set_once(&mut target_frame, value, "w:tgtFrame")?,
                            b"docLocation" => {
                                set_once(&mut document_location, value, "w:docLocation")?;
                            },
                            _ => {},
                        }
                    }
                } else if depth == 2 {
                    if element.local_name().as_ref() == b"r" {
                        capture = Some((start, depth));
                    } else {
                        has_unmodeled_content = true;
                    }
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("hyperlink nesting is too deep".into())
                })?;
                if child_depth == 2 {
                    if element.local_name().as_ref() == b"r" {
                        runs.push(Run::new(xml[start..end].to_vec()));
                    } else {
                        has_unmodeled_content = true;
                    }
                }
            },
            Event::End(_) => {
                if let Some((run_start, run_depth)) = capture
                    && depth == run_depth
                {
                    runs.push(Run::new(xml[run_start..end].to_vec()));
                    capture = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("invalid hyperlink nesting".into())
                })?;
            },
            Event::Text(text) if depth == 1 => {
                if text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) {
                    has_unmodeled_content = true;
                }
            },
            Event::CData(text) if depth == 1 => {
                if text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) {
                    has_unmodeled_content = true;
                }
            },
            Event::Eof if depth != 0 || capture.is_some() => {
                return Err(crate::Error::InvalidFormat(
                    "unterminated hyperlink XML".into(),
                ));
            },
            Event::Eof => break,
            Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)
            | Event::Text(_)
            | Event::CData(_) => {},
        }
    }
    if !saw_root {
        return Err(crate::Error::InvalidFormat(
            "hyperlink XML has no root".into(),
        ));
    }
    Ok(InlineHyperlink::new(
        link,
        runs,
        target_frame,
        document_location,
        has_unmodeled_content,
    ))
}

fn set_once(slot: &mut Option<String>, value: String, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(crate::Error::InvalidFormat(format!(
            "hyperlink has duplicate {name}"
        )));
    }
    Ok(())
}
