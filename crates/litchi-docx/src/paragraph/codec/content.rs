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
use litchi_ooxml_common::xml::{is_omml_name, omml_formula_xml};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, PrefixDeclaration};
use quick_xml::reader::NsReader;
use smallvec::SmallVec;

use super::super::model::Paragraph;

impl Paragraph {
    /// Extract all OMML formulas from this paragraph.
    ///
    /// Inline equations are `<m:oMath>` elements outside an `<m:oMathPara>`.
    /// Word writes them beside ordinary runs; older run-nested equations are
    /// also read. The returned strings preserve the source XML exactly.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn omml_formulas(&self) -> Result<Vec<String>> {
        self.raw_office_math(false)
    }

    /// Extract inline Office Math equations as validated typed fragments.
    ///
    /// Inline equations are `<m:oMath>` elements outside math paragraphs. Their
    /// exact raw XML remains available through [`Self::omml_formulas`]; this
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

    /// Extract paragraph-level OMML formulas.
    ///
    /// Returns the original XML of equations enclosed by `<m:oMathPara>`
    /// containers, in document order. A bare paragraph-child `<m:oMath>` is
    /// inline math and is available through [`Self::omml_formulas`].
    ///
    /// # Example
    /// ```ignore
    /// let para = document.paragraphs()?[0];
    /// let display_formulas = para.paragraph_level_formulas()?;
    /// for formula in display_formulas {
    ///     println!("Display formula: {}", formula);
    /// }
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn paragraph_level_formulas(&self) -> Result<Vec<String>> {
        self.raw_office_math(true)
    }

    /// Extract display Office Math equations as validated typed fragments.
    ///
    /// Display equations are `<m:oMath>` elements enclosed by an
    /// `<m:oMathPara>` math paragraph. The result is flattened
    /// into document order; use [`Self::paragraph_level_formulas`] when the
    /// original XML strings are required. Typed fragments declare their
    /// inherited source namespaces locally.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn display_office_math(&self) -> Result<Vec<OfficeMath>> {
        self.typed_office_math(true)
    }

    fn raw_office_math(&self, display: bool) -> Result<Vec<String>> {
        self.office_math_ranges(display)?
            .into_iter()
            .map(|(start, length)| Ok(omml_formula_xml(self.xml_bytes(), start, length)?))
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
