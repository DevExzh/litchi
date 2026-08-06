//! Paragraph content adapters for drawings, revisions, and Office Math.

use crate::drawing::{LegacyAnchor, Object, parse, parse_legacy};
use crate::error::{Error, Result};
use crate::image::{InlineImage, parse_inline_images};
use crate::math::OfficeMath;
use crate::namespace::scan_word_element_ranges;
use crate::revision::{Revision, parse_revisions};
use litchi_ooxml_common::xml::{omml_formula_xml, scan_omml_formula_ranges};
use smallvec::SmallVec;

use super::super::model::Paragraph;

impl Paragraph {
    /// Extract all OMML formulas from this paragraph.
    ///
    /// Returns a vector of OMML formula strings found in any run within this paragraph.
    /// This extracts inline formulas (formulas within runs).
    pub fn omml_formulas(&self) -> Result<Vec<String>> {
        let mut run_ranges = Vec::new();
        scan_word_element_ranges(self.xml_bytes(), &[b"r".as_slice()], |_, start, length| {
            let start = start as usize;
            let end = start
                .checked_add(length as usize)
                .ok_or_else(|| Error::InvalidFormat("Word run range overflows".to_string()))?;
            run_ranges.push((start, end));
            Ok(())
        })?;

        let mut formulas = Vec::new();
        scan_omml_formula_ranges(self.xml_bytes(), |start, length| {
            let formula_start = start as usize;
            let formula_end = formula_start
                .checked_add(length as usize)
                .ok_or_else(|| Error::InvalidFormat("OMML formula range overflows".to_string()))?;
            let is_inline = run_ranges
                .iter()
                .any(|&(run_start, run_end)| formula_start >= run_start && formula_end <= run_end);
            if is_inline {
                formulas.push(omml_formula_xml(self.xml_bytes(), start, length)?);
            }
            Ok::<(), Error>(())
        })?;
        Ok(formulas)
    }

    /// Extract inline Office Math equations as validated typed fragments.
    ///
    /// Inline equations are `<m:oMath>` elements nested in Word runs.  Their
    /// exact raw XML remains available through [`Self::omml_formulas`]; this
    /// method turns each fragment into a validated [`OfficeMath`] value
    /// suitable for reuse with the mutable writer.
    pub fn inline_office_math(&self) -> Result<Vec<OfficeMath>> {
        self.omml_formulas()?
            .into_iter()
            .map(OfficeMath::from_xml)
            .collect()
    }

    /// Extract all inline images from this paragraph.
    ///
    /// Returns a vector of `InlineImage` objects found in `<w:drawing>` elements
    /// within this paragraph.
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
    pub fn images(&self) -> Result<SmallVec<[InlineImage; 4]>> {
        parse_inline_images(self.xml_bytes())
    }

    /// Extract all drawing objects (shapes, text boxes) from this paragraph.
    ///
    /// Returns a vector of [`Object`] values found in `<w:drawing>` elements
    /// within this paragraph. This includes shapes, text boxes, and other DrawingML objects.
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
    pub fn drawing_objects(&self) -> Result<SmallVec<[Object; 4]>> {
        parse(self.xml_bytes())
    }

    /// Extract legacy Word object and picture anchors from this paragraph.
    ///
    /// The returned entries expose only the typed `[MS-DOCX]` 2.2.6
    /// `w14:anchorId` metadata and whether the source element was `w:object`
    /// or `w:pict`. VML, OLE, image payloads, and layout are intentionally
    /// inert and are not interpreted.
    #[inline]
    pub fn legacy_anchors(&self) -> Result<SmallVec<[LegacyAnchor; 4]>> {
        parse_legacy(self.xml_bytes())
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
    ///             revision.author(),
    ///             revision.revision_type(),
    ///             revision.text()
    ///         );
    ///     }
    /// }
    /// ```
    #[inline]
    pub fn revisions(&self) -> Result<SmallVec<[Revision; 4]>> {
        parse_revisions(self.xml_bytes())
    }

    /// Extract paragraph-level OMML formulas.
    ///
    /// Returns a vector of OMML formula strings that are direct children of the paragraph
    /// (display math), not nested within runs. These are block-level formulas.
    ///
    /// # Example
    /// ```ignore
    /// let para = document.paragraphs()?[0];
    /// let display_formulas = para.paragraph_level_formulas()?;
    /// for formula in display_formulas {
    ///     println!("Display formula: {}", formula);
    /// }
    /// ```
    pub fn paragraph_level_formulas(&self) -> Result<Vec<String>> {
        let mut run_ranges = Vec::new();
        scan_word_element_ranges(self.xml_bytes(), &[b"r".as_slice()], |_, start, length| {
            let start = start as usize;
            let end = start
                .checked_add(length as usize)
                .ok_or_else(|| Error::InvalidFormat("Word run range overflows".to_string()))?;
            run_ranges.push((start, end));
            Ok(())
        })?;

        let mut formulas = Vec::new();
        scan_omml_formula_ranges(self.xml_bytes(), |start, length| {
            let formula_start = start as usize;
            let formula_end = formula_start
                .checked_add(length as usize)
                .ok_or_else(|| Error::InvalidFormat("OMML formula range overflows".to_string()))?;
            let is_inline = run_ranges
                .iter()
                .any(|&(run_start, run_end)| formula_start >= run_start && formula_end <= run_end);
            if !is_inline {
                formulas.push(omml_formula_xml(self.xml_bytes(), start, length)?);
            }
            Ok::<(), Error>(())
        })?;
        Ok(formulas)
    }

    /// Extract display Office Math equations as validated typed fragments.
    ///
    /// Display equations are `<m:oMath>` elements outside Word runs, normally
    /// enclosed by an `<m:oMathPara>` math paragraph.  The result is flattened
    /// into document order; use [`Self::paragraph_level_formulas`] when the
    /// original XML strings are required.
    pub fn display_office_math(&self) -> Result<Vec<OfficeMath>> {
        self.paragraph_level_formulas()?
            .into_iter()
            .map(OfficeMath::from_xml)
            .collect()
    }
}
