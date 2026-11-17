#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Paragraph content adapters for drawings, revisions, and Office Math.

use crate::drawing::{LegacyAnchor, Object, parse, parse_legacy_with_context};
use crate::error::{Error, Result};
use crate::image::{InlineImage, parse_inline_images};
use crate::math::OfficeMath;
use crate::revision::{Revision, parse_revisions_with_context};
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::{is_omml_name, omml_equation_xml};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, PrefixDeclaration};
use quick_xml::reader::NsReader;
use smallvec::SmallVec;

use super::super::model::Paragraph;

impl Paragraph {
    /// Extract all OMML equations from this paragraph.
    ///
    /// Inline equations are `<m:oMath>` elements outside an `<m:oMathPara>`.
    /// Word writes them beside ordinary runs; older run-nested equations are
    /// also read. The returned strings preserve the source XML exactly.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn omml_equations(&self) -> Result<Vec<String>> {
        self.raw_office_math(false)
    }

    /// Extract inline Office Math equations as validated typed fragments.
    ///
    /// Inline equations are `<m:oMath>` elements outside math paragraphs. Their
    /// exact raw XML remains available through [`Self::omml_equations`]; this
    /// method turns each fragment into a validated [`OfficeMath`] value
    /// suitable for reuse with the mutable writer. Inherited namespaces are
    /// declared locally, preserving strict and aliased source bindings.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn inline_office_math(&self) -> Result<Vec<OfficeMath>> {
        self.typed_office_math(false)
    }

    /// Extract all inline and floating images from this paragraph.
    ///
    /// Returns a vector of `InlineImage` objects found in `<w:drawing>`
    /// elements within this paragraph, covering both `<wp:inline>` pictures
    /// and floating `<wp:anchor>` pictures. Floating placement metadata is
    /// not retained.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for para in document.paragraphs()? {
    ///     for image in para.images()? {
    ///         println!("Image: {} ({}x{} pixels)",
    ///             image.name(),
    ///             image.width_px(),
    ///             image.height_px()
    ///         );
    ///     }
    /// }
    /// ```
    #[inline]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn images(&self) -> Result<SmallVec<[InlineImage; 4]>> {
        let _parser_admission = self.parser_admission()?;
        parse_inline_images(self.xml_bytes())
    }

    /// The charts drawn in this paragraph (`w:drawing` elements holding a
    /// `c:chart` or an extended `cx:chart`), in document order.
    ///
    /// ```
    /// use litchi_docx::Package;
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("No chart here");
    /// let package = builder.build()?;
    /// let paragraphs = package.document()?.paragraphs()?;
    /// assert!(paragraphs[0].charts()?.is_empty());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for malformed paragraph XML.
    pub fn charts(&self) -> Result<SmallVec<[crate::chart::InlineChart; 2]>> {
        let _parser_admission = self.parser_admission()?;
        crate::chart::inline::parse_inline_charts(self.xml_bytes())
    }

    /// Extract all drawing objects (shapes, text boxes) from this paragraph.
    ///
    /// Returns a vector of [`Object`] values found in `<w:drawing>` elements
    /// within this paragraph. This includes shapes, text boxes, and other `DrawingML` objects.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for para in document.paragraphs()? {
    ///     for drawing in para.drawing_objects()? {
    ///         println!("Shape: {} (type: {:?})",
    ///             drawing.name(),
    ///             drawing.preset()
    ///         );
    ///         if !drawing.text().is_empty() {
    ///             println!("  Text: {}", drawing.text());
    ///         }
    ///     }
    /// }
    /// ```
    #[inline]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn drawing_objects(&self) -> Result<SmallVec<[Object; 4]>> {
        let _parser_admission = self.parser_admission()?;
        parse(self.xml_bytes())
    }

    /// Extract legacy Word object and picture anchors from this paragraph.
    ///
    /// The returned entries expose only the typed `[MS-DOCX]` 2.2.6
    /// `w14:anchorId` metadata and whether the source element was `w:object`
    /// or `w:pict`. VML, OLE, image payloads, and layout are intentionally
    /// inert and are not interpreted.
    #[inline]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn legacy_anchors(&self) -> Result<SmallVec<[LegacyAnchor; 4]>> {
        let _parser_admission = self.parser_admission()?;
        let namespaces = self.xml_data.inherited_namespaces()?;
        parse_legacy_with_context(self.xml_bytes(), &namespaces)
    }

    /// Extract all tracked changes (revisions) from this paragraph.
    ///
    /// Returns a vector of `Revision` objects representing all tracked changes
    /// (insertions, deletions, moves) within this paragraph.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// for para in document.paragraphs()? {
    ///     for revision in para.revisions()? {
    ///         println!("Revision by {}: {} - {}",
    ///             revision.author().unwrap_or("Unattributed"),
    ///             revision.revision_type(),
    ///             revision.text()
    ///         );
    ///     }
    /// }
    /// ```
    #[inline]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn revisions(&self) -> Result<SmallVec<[Revision; 4]>> {
        let _parser_admission = self.parser_admission()?;
        let namespaces = self.xml_data.inherited_namespaces()?;
        parse_revisions_with_context(self.xml_bytes(), &namespaces)
    }

    /// Read revisions in source order using explicit finite resource limits.
    /// Nested annotations retain their own metadata and affected text.
    ///
    /// # Errors
    /// Returns an error for malformed revision XML or an exceeded resource limit.
    pub(crate) fn revisions_bounded(
        &self,
        limits: crate::revision::Limits,
    ) -> Result<SmallVec<[Revision; 4]>> {
        let _parser_admission = self.parser_admission()?;
        let namespaces = self.xml_data.inherited_namespaces()?;
        crate::revision::parse_revisions_with_limits(self.xml_bytes(), &namespaces, limits)
    }

    /// Extract paragraph-level OMML equations.
    ///
    /// Returns the original XML of equations enclosed by `<m:oMathPara>`
    /// containers, in document order. A bare paragraph-child `<m:oMath>` is
    /// inline math and is available through [`Self::omml_equations`].
    ///
    /// # Example
    /// ```ignore
    /// let para = document.paragraphs()?[0];
    /// let display_equations = para.paragraph_level_equations()?;
    /// for equation in display_equations {
    ///     println!("Display equation: {equation}");
    /// }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn paragraph_level_equations(&self) -> Result<Vec<String>> {
        self.raw_office_math(true)
    }

    /// Extract display Office Math equations as validated typed fragments.
    ///
    /// Display equations are `<m:oMath>` elements enclosed by an
    /// `<m:oMathPara>` math paragraph. The result is flattened
    /// into document order; use [`Self::paragraph_level_equations`] when the
    /// original XML strings are required. Typed fragments declare their
    /// inherited source namespaces locally.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn display_office_math(&self) -> Result<Vec<OfficeMath>> {
        self.typed_office_math(true)
    }

    /// The paragraph's math paragraphs (`m:oMathPara`), each with the
    /// equations it holds, as their original XML in document order.
    ///
    /// A display-equation paragraph holds one; a paragraph that also holds
    /// text, a bookmark or inline math may hold several.
    ///
    /// ```
    /// use litchi_docx::Package;
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("No math here");
    /// let package = builder.build()?;
    /// assert!(package.document()?.paragraphs()?[0].math_paragraphs()?.is_empty());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for malformed paragraph XML or a scan beyond the
    /// Office Math bounds.
    pub fn math_paragraphs(&self) -> Result<Vec<String>> {
        const MAX_DEPTH: usize = 128;
        const MAX_NODES: usize = 1_000_000;

        let _parser_admission = self.parser_admission()?;
        let xml = self.xml_bytes();
        let mut reader = NsReader::from_reader(xml);
        for (prefix, namespace) in self.xml_data.inherited_namespaces()?.iter() {
            let prefix = prefix
                .as_deref()
                .map_or(PrefixDeclaration::Default, PrefixDeclaration::Named);
            reader
                .resolver_mut()
                .add(prefix, Namespace(namespace))
                .map_err(|_| Error::Xml("invalid inherited Office Math namespace".into()))?;
        }
        let origin = ReaderOrigin::of(xml);
        let mut depth = 0usize;
        let mut capture = None::<(usize, usize)>;
        let mut found = Vec::new();
        let mut nodes = 0usize;
        loop {
            let start = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| Error::InvalidFormat("Office Math offset exceeds usize".into()))?;
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let end = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| Error::InvalidFormat("Office Math offset exceeds usize".into()))?;
            let (namespace, event) = reader.resolver().resolve_event(event);
            match event {
                Event::Start(element) => {
                    nodes += 1;
                    if nodes > MAX_NODES || depth >= MAX_DEPTH {
                        return Err(Error::InvalidFormat(
                            "Office Math scan limit exceeded".into(),
                        ));
                    }
                    depth += 1;
                    if capture.is_none() && is_omml_name(&namespace, element.name(), b"oMathPara") {
                        capture = Some((start, depth));
                    }
                },
                Event::Empty(element) => {
                    nodes += 1;
                    if capture.is_none() && is_omml_name(&namespace, element.name(), b"oMathPara") {
                        found.push(omml_equation_xml_range(xml, start, end)?);
                    }
                },
                Event::End(_) => {
                    if let Some((from, at)) = capture
                        && at == depth
                    {
                        found.push(omml_equation_xml_range(xml, from, end)?);
                        capture = None;
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| Error::Xml("invalid Office Math nesting".into()))?;
                },
                Event::Eof => {
                    if depth != 0 || capture.is_some() {
                        return Err(Error::Xml("unterminated Office Math XML".into()));
                    }
                    return Ok(found);
                },
                _ => {},
            }
        }
    }

    fn raw_office_math(&self, display: bool) -> Result<Vec<String>> {
        self.office_math_ranges(display)?
            .into_iter()
            .map(|(start, length)| Ok(omml_equation_xml(self.xml_bytes(), start, length)?))
            .collect()
    }

    fn typed_office_math(&self, display: bool) -> Result<Vec<OfficeMath>> {
        let ranges = self.office_math_ranges(display)?;
        if ranges.is_empty() {
            return Ok(Vec::new());
        }
        let source = self.xml_data.xml_ref()?;
        ranges
            .into_iter()
            .map(|(start, length)| {
                let fragment = source.subrange(start, length).ok_or_else(|| {
                    Error::InvalidFormat("Office Math range exceeds its paragraph".into())
                })?;
                let xml = fragment.self_contained_bytes()?;
                let text =
                    std::str::from_utf8(&xml).map_err(|error| Error::Xml(error.to_string()))?;
                OfficeMath::from_xml(text)
            })
            .collect()
    }

    /// Classify by the math-paragraph ancestor, not by a Word run envelope.
    /// Namespace context includes bindings inherited from the owning part.
    fn office_math_ranges(&self, display: bool) -> Result<Vec<(u32, u32)>> {
        const MAX_DEPTH: usize = 128;
        const MAX_NODES: usize = 1_000_000;

        let _parser_admission = self.parser_admission()?;
        let xml = self.xml_bytes();
        let mut reader = NsReader::from_reader(xml);
        for (prefix, namespace) in self.xml_data.inherited_namespaces()?.iter() {
            let prefix = prefix
                .as_deref()
                .map_or(PrefixDeclaration::Default, PrefixDeclaration::Named);
            reader
                .resolver_mut()
                .add(prefix, Namespace(namespace))
                .map_err(|_| Error::Xml("invalid inherited Office Math namespace".into()))?;
        }
        let origin = ReaderOrigin::of(xml);
        let mut ancestors = Vec::new();
        let mut capture = None;
        let mut ranges = Vec::new();
        let mut nodes = 0usize;
        loop {
            let start = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| Error::InvalidFormat("Office Math offset exceeds usize".into()))?;
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let end = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| Error::InvalidFormat("Office Math offset exceeds usize".into()))?;
            let (namespace, event) = reader.resolver().resolve_event(event);
            if matches!(event, Event::Start(_) | Event::Empty(_)) {
                nodes += 1;
                if nodes > MAX_NODES || ancestors.len() >= MAX_DEPTH {
                    return Err(Error::InvalidFormat(
                        "Office Math scan limit exceeded".into(),
                    ));
                }
            }
            let completed = match event {
                Event::Start(element) => {
                    let in_display = ancestors.contains(&true);
                    ancestors.push(is_omml_name(&namespace, element.name(), b"oMathPara"));
                    if capture.is_none() && is_omml_name(&namespace, element.name(), b"oMath") {
                        capture = Some((start, ancestors.len(), in_display));
                    }
                    None
                },
                Event::Empty(element)
                    if capture.is_none() && is_omml_name(&namespace, element.name(), b"oMath") =>
                {
                    Some((start, ancestors.contains(&true)))
                },
                Event::End(_) => {
                    let completed = if capture.is_some_and(|(_, depth, _)| depth == ancestors.len())
                    {
                        capture.take().map(|(start, _, display)| (start, display))
                    } else {
                        None
                    };
                    ancestors
                        .pop()
                        .ok_or_else(|| Error::Xml("invalid Office Math nesting".into()))?;
                    completed
                },
                Event::Eof => {
                    if !ancestors.is_empty() || capture.is_some() {
                        return Err(Error::Xml("unterminated Office Math XML".into()));
                    }
                    break;
                },
                _ => None,
            };
            if let Some((start, _)) = completed.filter(|(_, value)| *value == display) {
                let length = end
                    .checked_sub(start)
                    .and_then(|length| u32::try_from(length).ok())
                    .ok_or_else(|| Error::InvalidFormat("Office Math length exceeds u32".into()))?;
                let start = u32::try_from(start)
                    .map_err(|_| Error::InvalidFormat("Office Math offset exceeds u32".into()))?;
                ranges.push((start, length));
            }
        }
        Ok(ranges)
    }
}

/// The UTF-8 text of `xml[start..end]`.
fn omml_equation_xml_range(xml: &[u8], start: usize, end: usize) -> Result<String> {
    let bytes = xml
        .get(start..end)
        .ok_or_else(|| Error::InvalidFormat("Office Math range exceeds its paragraph".into()))?;
    Ok(std::str::from_utf8(bytes)
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned())
}

/// A piece of a paragraph's content in document order: text, an inline
/// equation or a math paragraph (their original OMML).
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum TextPiece {
    /// Text of the paragraph's `w:t`, `w:tab` (`\t`), `w:br` and `w:cr`
    /// (`\n`) elements.
    Text(String),
    /// An inline equation (`m:oMath`).
    Equation(String),
    /// A math paragraph (`m:oMathPara`).
    MathParagraph(String),
}

impl Paragraph {
    /// The paragraph's text and equations in document order: what a
    /// consumer that renders equations in place reads. Deleted text and
    /// field instructions are not text; text inside an equation is part of
    /// the equation.
    ///
    /// ```
    /// use litchi_docx::{Package, TextPiece};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Plain");
    /// let package = builder.build()?;
    /// let pieces = package.document()?.paragraphs()?[0].text_and_equations()?;
    /// assert_eq!(pieces, [TextPiece::Text("Plain".into())]);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for malformed paragraph XML or a scan beyond the
    /// Office Math bounds.
    pub fn text_and_equations(&self) -> Result<Vec<TextPiece>> {
        const MAX_DEPTH: usize = 128;
        const MAX_NODES: usize = 1_000_000;

        let _parser_admission = self.parser_admission()?;
        let xml = self.xml_bytes();
        let mut reader = NsReader::from_reader(xml);
        reader.config_mut().trim_text(false);
        for (prefix, namespace) in self.xml_data.inherited_namespaces()?.iter() {
            let prefix = prefix
                .as_deref()
                .map_or(PrefixDeclaration::Default, PrefixDeclaration::Named);
            reader
                .resolver_mut()
                .add(prefix, Namespace(namespace))
                .map_err(|_| Error::Xml("invalid inherited Office Math namespace".into()))?;
        }
        let origin = ReaderOrigin::of(xml);
        let mut depth = 0usize;
        // The open equation: its start, its depth and whether it is a math
        // paragraph.
        let mut capture = None::<(usize, usize, bool)>;
        // The depth of an open `w:t`.
        let mut text_at = None::<usize>;
        let mut pieces: Vec<TextPiece> = Vec::new();
        let push_text = |pieces: &mut Vec<TextPiece>, text: &str| {
            if let Some(TextPiece::Text(last)) = pieces.last_mut() {
                last.push_str(text);
            } else if !text.is_empty() {
                pieces.push(TextPiece::Text(text.to_owned()));
            }
        };
        let mut nodes = 0usize;
        loop {
            let start = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| Error::InvalidFormat("Office Math offset exceeds usize".into()))?;
            let event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let end = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| Error::InvalidFormat("Office Math offset exceeds usize".into()))?;
            let (namespace, event) = reader.resolver().resolve_event(event);
            let word = matches!(
                namespace,
                quick_xml::name::ResolveResult::Bound(Namespace(
                    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        | b"http://purl.oclc.org/ooxml/wordprocessingml/main"
                ))
            );
            match event {
                Event::Start(element) => {
                    nodes += 1;
                    if nodes > MAX_NODES || depth >= MAX_DEPTH {
                        return Err(Error::InvalidFormat(
                            "Office Math scan limit exceeded".into(),
                        ));
                    }
                    depth += 1;
                    if capture.is_none() {
                        if is_omml_name(&namespace, element.name(), b"oMath") {
                            capture = Some((start, depth, false));
                        } else if is_omml_name(&namespace, element.name(), b"oMathPara") {
                            capture = Some((start, depth, true));
                        } else if word && element.local_name().as_ref() == b"t" {
                            text_at = Some(depth);
                        }
                    }
                },
                Event::Empty(element) => {
                    nodes += 1;
                    if capture.is_some() {
                        continue;
                    }
                    if is_omml_name(&namespace, element.name(), b"oMath") {
                        pieces.push(TextPiece::Equation(omml_equation_xml_range(
                            xml, start, end,
                        )?));
                    } else if is_omml_name(&namespace, element.name(), b"oMathPara") {
                        pieces.push(TextPiece::MathParagraph(omml_equation_xml_range(
                            xml, start, end,
                        )?));
                    } else if word {
                        match element.local_name().as_ref() {
                            b"tab" => push_text(&mut pieces, "\t"),
                            b"br" | b"cr" => push_text(&mut pieces, "\n"),
                            _ => {},
                        }
                    }
                },
                Event::Text(text) if text_at == Some(depth) && capture.is_none() => {
                    let text = text
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    push_text(&mut pieces, &text);
                },
                Event::GeneralRef(reference) if text_at == Some(depth) && capture.is_none() => {
                    let character = litchi_ooxml_common::xml::decode_xml_reference(&reference)?;
                    push_text(&mut pieces, &character);
                },
                Event::End(_) => {
                    if let Some((from, at, display)) = capture
                        && at == depth
                    {
                        let omml = omml_equation_xml_range(xml, from, end)?;
                        pieces.push(if display {
                            TextPiece::MathParagraph(omml)
                        } else {
                            TextPiece::Equation(omml)
                        });
                        capture = None;
                    }
                    if text_at == Some(depth) {
                        text_at = None;
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| Error::Xml("invalid Office Math nesting".into()))?;
                },
                Event::Eof => {
                    if depth != 0 || capture.is_some() {
                        return Err(Error::Xml("unterminated Office Math XML".into()));
                    }
                    return Ok(pieces);
                },
                _ => {},
            }
        }
    }
}

impl Paragraph {
    /// The paragraph's OLE equation objects (MathType and Equation Editor
    /// 3.0), in document order: every `w:object` whose `o:OLEObject` names
    /// an equation ProgID ([`crate::object::is_equation_prog_id`]).
    ///
    /// ```
    /// use litchi_docx::Package;
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("No objects");
    /// let package = builder.build()?;
    /// assert!(package.document()?.paragraphs()?[0].equation_objects()?.is_empty());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for malformed paragraph XML or a scan beyond its
    /// bounds.
    pub fn equation_objects(&self) -> Result<Vec<crate::object::EquationObject>> {
        let _parser_admission = self.parser_admission()?;
        Ok(crate::object::equation_spans(self.xml_bytes())?
            .iter()
            .map(crate::object::EquationSpan::object)
            .collect())
    }
}
