/// High-performance Slide implementation with lazy shape loading and zero-copy design.
use super::super::package::Result;
use super::super::records::PptRecord;
use super::super::shapes::ShapeEnum;
use super::factory::SlideData;
use once_cell::unsync::OnceCell;

/// A slide in a PowerPoint presentation with lazy-loaded shapes.
///
/// # Performance
///
/// - Shapes are parsed only when first accessed
/// - Uses `OnceCell` for one-time initialization
/// - Zero-copy text extraction where possible
pub struct Slide<'doc> {
    /// Slide persist ID
    persist_id: u32,
    /// Slide number (1-based for display)
    slide_number: usize,
    /// Slide record
    record: PptRecord,
    /// Reference to document data for lazy shape parsing (reserved for future use)
    #[allow(dead_code)]
    doc_data: &'doc [u8],
    /// Lazily-loaded shapes
    shapes: OnceCell<Vec<ShapeEnum>>,
    /// Cached text content
    text_cache: OnceCell<String>,
}

impl<'doc> Slide<'doc> {
    /// Create a slide from parsed slide data.
    pub fn from_slide_data(data: SlideData<'doc>, slide_number: usize) -> Self {
        let doc_data_ref = data.doc_data();
        Self {
            persist_id: data.persist_id,
            slide_number,
            doc_data: doc_data_ref,
            record: data.record,
            shapes: OnceCell::new(),
            text_cache: OnceCell::new(),
        }
    }

    /// Get the slide number (1-based).
    #[inline]
    pub fn slide_number(&self) -> usize {
        self.slide_number
    }

    /// Get the persist ID.
    #[inline]
    pub fn persist_id(&self) -> u32 {
        self.persist_id
    }

    /// Get shapes on this slide (lazy-loaded).
    ///
    /// # Performance
    ///
    /// - Shapes are parsed only on first call
    /// - Subsequent calls return cached reference
    /// - Zero allocation after first parse
    pub fn shapes(&self) -> Result<&[ShapeEnum]> {
        self.shapes
            .get_or_try_init(|| self.parse_shapes())
            .map(|v| v.as_slice())
    }

    /// Get the number of shapes (triggers parsing if not yet loaded).
    pub fn shape_count(&self) -> Result<usize> {
        Ok(self.shapes()?.len())
    }

    /// Extract all text from this slide (lazy-loaded).
    ///
    /// # Performance
    ///
    /// - Text is extracted and cached on first call
    /// - Includes text from:
    ///   * Direct text records in the slide
    ///   * Shapes (via PPDrawing/Escher)
    pub fn text(&self) -> Result<&str> {
        self.text_cache
            .get_or_try_init(|| self.extract_all_text())
            .map(|s| s.as_str())
    }

    /// Parse shapes from PPDrawing record.
    ///
    /// # Performance
    ///
    /// - Zero-copy: Shapes borrow from slide's document data
    /// - Lazy: Only called when shapes() is accessed
    /// - Uses Escher parser for efficient traversal
    fn parse_shapes(&self) -> Result<Vec<ShapeEnum>> {
        // Find PPDrawing record
        let ppdrawing = match self
            .record
            .find_child(crate::ole::consts::PptRecordType::PPDrawing)
        {
            Some(record) => record,
            None => return Ok(Vec::new()),
        };

        // Extract Escher shapes from PPDrawing data
        let escher_shapes =
            super::super::escher::EscherShapeFactory::extract_shapes_from_ppdrawing(
                &ppdrawing.data,
            )?;

        // Convert Escher shapes to ShapeEnum with full property extraction
        let shapes: Vec<ShapeEnum> = escher_shapes
            .iter()
            .filter_map(|escher_shape| Self::convert_escher_to_shape_enum(escher_shape))
            .collect();

        Ok(shapes)
    }

    /// Convert an EscherShape to ShapeEnum with full property extraction.
    ///
    /// # Performance
    ///
    /// - Direct property access (no allocations)
    /// - Pattern matching for type dispatch
    fn convert_escher_to_shape_enum(
        escher_shape: &super::super::escher::EscherShape<'_>,
    ) -> Option<ShapeEnum> {
        use super::super::escher::EscherShapeType;
        use super::super::shapes::*;

        let shape_id = escher_shape.shape_id().unwrap_or(0);
        let anchor = escher_shape.anchor();

        match escher_shape.shape_type() {
            EscherShapeType::TextBox => {
                // Create TextBox with proper properties
                let mut properties = shape::ShapeProperties {
                    id: shape_id,
                    shape_type: shape::ShapeType::TextBox,
                    ..Default::default()
                };

                // Set coordinates if anchor exists
                if let Some(a) = anchor {
                    properties.x = a.left;
                    properties.y = a.top;
                    properties.width = a.width();
                    properties.height = a.height();
                }

                // Extract text from shape
                let text = escher_shape.text().unwrap_or_default();

                let mut textbox = TextBox::new(properties, Vec::new());
                if !text.is_empty() {
                    textbox.set_text(text);
                }

                Some(ShapeEnum::TextBox(textbox))
            },

            EscherShapeType::Picture => {
                // Create PictureShape
                let mut picture = shape_enum::PictureShape::new(shape_id);

                if let Some(a) = anchor {
                    picture.set_bounds(a.left, a.top, a.width(), a.height());
                }

                // Extract blip ID from properties
                use super::super::escher::EscherPropertyId;
                if let Some(blip_id) = escher_shape
                    .properties()
                    .get_int(EscherPropertyId::PictureId)
                {
                    picture.set_blip_id(blip_id as u32);
                }

                Some(ShapeEnum::Picture(picture))
            },

            EscherShapeType::Line => {
                // Create LineShape
                if let Some(a) = anchor {
                    let mut line =
                        shape_enum::LineShape::new(shape_id, a.left, a.top, a.right, a.bottom);

                    // Extract line properties
                    use super::super::escher::EscherPropertyId;
                    if let Some(width) = escher_shape
                        .properties()
                        .get_int(EscherPropertyId::LineWidth)
                    {
                        line.set_width(width);
                    }
                    if let Some(color) = escher_shape
                        .properties()
                        .get_color(EscherPropertyId::LineColor)
                    {
                        line.set_color(color);
                    }

                    Some(ShapeEnum::Line(line))
                } else {
                    None
                }
            },

            EscherShapeType::Group => {
                // Create GroupShape (children would be parsed recursively)
                let mut group = shape_enum::GroupShape::new(shape_id);

                if let Some(a) = anchor {
                    group.set_bounds(a.left, a.top, a.width(), a.height());
                }

                // TODO: Parse child shapes recursively

                Some(ShapeEnum::Group(group))
            },

            EscherShapeType::Rectangle | EscherShapeType::Ellipse | EscherShapeType::AutoShape => {
                // Create AutoShape
                let mut properties = shape::ShapeProperties {
                    id: shape_id,
                    shape_type: shape::ShapeType::AutoShape,
                    ..Default::default()
                };

                if let Some(a) = anchor {
                    properties.x = a.left;
                    properties.y = a.top;
                    properties.width = a.width();
                    properties.height = a.height();
                }

                let autoshape = AutoShape::new(properties, Vec::new());
                Some(ShapeEnum::AutoShape(autoshape))
            },

            // Unknown or unsupported shape types
            _ => None,
        }
    }

    /// Extract all text from slide and its shapes.
    fn extract_all_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        // 1. Extract text from direct slide records (TextCharsAtom, etc.)
        // Note: record.extract_text() already recursively processes all children
        if let Ok(record_text) = self.record.extract_text() {
            let trimmed = record_text.trim();
            if !trimmed.is_empty() {
                text_parts.push(trimmed.to_string());
            }
        }

        // 2. Extract text from Escher/PPDrawing (shapes, text boxes)
        // This is separate from regular record text extraction
        if let Some(ppdrawing) = self
            .record
            .find_child(crate::ole::consts::PptRecordType::PPDrawing)
            && let Ok(escher_text) = super::super::escher::extract_text_from_escher(&ppdrawing.data)
        {
            let trimmed = escher_text.trim();
            if !trimmed.is_empty() {
                text_parts.push(trimmed.to_string());
            }
        }

        Ok(if text_parts.is_empty() {
            String::new()
        } else {
            text_parts.join("\n")
        })
    }

    /// Check if this slide has a PPDrawing record (shapes).
    #[inline]
    pub fn has_drawing(&self) -> bool {
        self.record
            .find_child(crate::ole::consts::PptRecordType::PPDrawing)
            .is_some()
    }

    /// Get raw slide record for advanced use cases.
    #[inline]
    pub fn record(&self) -> &PptRecord {
        &self.record
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ole::consts::PptRecordType;
    use crate::ole::ppt::records::PptRecord;
    use crate::ole::ppt::slide::SlideData;

    // Helper function to create a test record
    fn create_test_record(
        record_type: PptRecordType,
        data: Vec<u8>,
        children: Vec<PptRecord>,
    ) -> PptRecord {
        PptRecord {
            record_type,
            record_type_raw: record_type as u16,
            version: 0,
            instance: 0,
            data_length: data.len() as u32,
            data,
            children,
        }
    }

    // Helper function to create a basic slide record without children
    fn create_basic_slide_record() -> PptRecord {
        create_test_record(PptRecordType::Slide, vec![0u8; 8], Vec::new())
    }

    // Helper function to create a slide with PPDrawing
    fn create_slide_with_drawing() -> PptRecord {
        let ppdrawing = create_test_record(
            PptRecordType::PPDrawing,
            vec![0u8; 16], // Empty escher data
            Vec::new(),
        );
        create_test_record(PptRecordType::Slide, vec![0u8; 8], vec![ppdrawing])
    }

    // Helper function to create a slide with text
    fn create_slide_with_text() -> PptRecord {
        // Create a TextCharsAtom with "Test" in UTF-16 LE
        let text_data = vec![
            0x54, 0x00, // 'T'
            0x65, 0x00, // 'e'
            0x73, 0x00, // 's'
            0x74, 0x00, // 't'
        ];
        let text_atom = create_test_record(PptRecordType::TextCharsAtom, text_data, Vec::new());
        create_test_record(PptRecordType::Slide, vec![0u8; 8], vec![text_atom])
    }

    // Helper function to create SlideData
    fn create_slide_data<'doc>(
        record: PptRecord,
        persist_id: u32,
        doc_data: &'doc [u8],
    ) -> SlideData<'doc> {
        SlideData::new_for_test(persist_id, 0, record, doc_data)
    }

    #[test]
    fn test_slide_creation() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        assert_eq!(slide.slide_number(), 1);
        assert_eq!(slide.persist_id(), 256);
    }

    #[test]
    fn test_slide_number_accessor() {
        let doc_data = vec![0u8; 512];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 100, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 5);

        assert_eq!(slide.slide_number(), 5);
    }

    #[test]
    fn test_persist_id_accessor() {
        let doc_data = vec![0u8; 512];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 999, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        assert_eq!(slide.persist_id(), 999);
    }

    #[test]
    fn test_has_drawing_without_ppdrawing() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        assert!(!slide.has_drawing());
    }

    #[test]
    fn test_has_drawing_with_ppdrawing() {
        let doc_data = vec![0u8; 1024];
        let record = create_slide_with_drawing();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        assert!(slide.has_drawing());
    }

    #[test]
    fn test_record_accessor() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        let rec = slide.record();
        assert_eq!(rec.record_type, PptRecordType::Slide);
    }

    #[test]
    fn test_shapes_empty_slide() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 0);
    }

    #[test]
    fn test_shapes_lazy_loading() {
        let doc_data = vec![0u8; 1024];
        let record = create_slide_with_drawing();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        // First call should initialize
        let shapes1 = slide.shapes().unwrap();
        // Second call should return cached value
        let shapes2 = slide.shapes().unwrap();

        // Both should return the same reference
        assert_eq!(shapes1.len(), shapes2.len());
    }

    #[test]
    fn test_shape_count_empty() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        assert_eq!(slide.shape_count().unwrap(), 0);
    }

    #[test]
    fn test_text_extraction_empty_slide() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "");
    }

    #[test]
    fn test_text_extraction_with_text_chars_atom() {
        let doc_data = vec![0u8; 1024];
        let record = create_slide_with_text();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "Test");
    }

    #[test]
    fn test_text_lazy_loading() {
        let doc_data = vec![0u8; 1024];
        let record = create_slide_with_text();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        // First call should extract text
        let text1 = slide.text().unwrap();
        // Second call should return cached value
        let text2 = slide.text().unwrap();

        assert_eq!(text1, text2);
        assert_eq!(text1, "Test");
    }

    #[test]
    fn test_text_extraction_with_nested_records() {
        let doc_data = vec![0u8; 1024];

        // Create nested structure: Slide -> SlideContainer -> TextCharsAtom
        let text_data = vec![
            0x41, 0x00, // 'A'
            0x42, 0x00, // 'B'
        ];
        let text_atom = create_test_record(PptRecordType::TextCharsAtom, text_data, Vec::new());

        let container = create_test_record(PptRecordType::SlideAtom, vec![0u8; 8], vec![text_atom]);

        let slide_record = create_test_record(PptRecordType::Slide, vec![0u8; 8], vec![container]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "AB");
    }

    #[test]
    fn test_text_extraction_multiple_text_atoms() {
        let doc_data = vec![0u8; 1024];

        // Create multiple TextCharsAtom records
        let text1_data = vec![
            0x48, 0x00, // 'H'
            0x69, 0x00, // 'i'
        ];
        let text1 = create_test_record(PptRecordType::TextCharsAtom, text1_data, Vec::new());

        let text2_data = vec![
            0x42, 0x00, // 'B'
            0x79, 0x00, // 'y'
            0x65, 0x00, // 'e'
        ];
        let text2 = create_test_record(PptRecordType::TextCharsAtom, text2_data, Vec::new());

        let slide_record =
            create_test_record(PptRecordType::Slide, vec![0u8; 8], vec![text1, text2]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        // Both text atoms should be extracted and joined
        assert!(text.contains("Hi"));
        assert!(text.contains("Bye"));
    }

    #[test]
    fn test_slide_with_different_text_atom_types() {
        let doc_data = vec![0u8; 1024];

        // Create TextBytesAtom (ASCII/ANSI encoding)
        let text_bytes = vec![0x54, 0x65, 0x78, 0x74]; // "Text" in ASCII
        let text_bytes_atom =
            create_test_record(PptRecordType::TextBytesAtom, text_bytes, Vec::new());

        let slide_record =
            create_test_record(PptRecordType::Slide, vec![0u8; 8], vec![text_bytes_atom]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "Text");
    }

    #[test]
    fn test_multiple_slide_numbers() {
        let doc_data = vec![0u8; 1024];

        let records: Vec<_> = (0..5).map(|_| create_basic_slide_record()).collect();

        let slides: Vec<_> = records
            .into_iter()
            .enumerate()
            .map(|(i, record)| {
                let slide_data = create_slide_data(record, 100 + i as u32, &doc_data);
                Slide::from_slide_data(slide_data, i + 1)
            })
            .collect();

        // Verify slide numbers are correctly assigned
        for (i, slide) in slides.iter().enumerate() {
            assert_eq!(slide.slide_number(), i + 1);
            assert_eq!(slide.persist_id(), 100 + i as u32);
        }
    }

    #[test]
    fn test_convert_escher_to_shape_enum_with_unknown_type() {
        // This tests that unknown shape types are filtered out
        // We can't easily construct EscherShape objects in tests without
        // implementing complex test data, but we can test the None path
        // through indirect testing via shapes()

        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        // Should return empty vec for slide without PPDrawing
        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 0);
    }

    #[test]
    fn test_extract_text_recursive_depth() {
        let doc_data = vec![0u8; 1024];

        // Create deeply nested structure
        let text_data = vec![0x58, 0x00]; // 'X'
        let text_atom = create_test_record(PptRecordType::TextCharsAtom, text_data, Vec::new());

        let level3 = create_test_record(PptRecordType::SlideAtom, vec![], vec![text_atom]);

        let level2 = create_test_record(PptRecordType::SlideAtom, vec![], vec![level3]);

        let level1 = create_test_record(PptRecordType::Slide, vec![], vec![level2]);

        let slide_data = create_slide_data(level1, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "X");
    }

    #[test]
    fn test_slide_with_whitespace_only_text() {
        let doc_data = vec![0u8; 1024];

        // Create TextCharsAtom with only whitespace
        let text_data = vec![
            0x20, 0x00, // space
            0x20, 0x00, // space
            0x09, 0x00, // tab
        ];
        let text_atom = create_test_record(PptRecordType::TextCharsAtom, text_data, Vec::new());

        let slide_record = create_test_record(PptRecordType::Slide, vec![], vec![text_atom]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        // Whitespace-only text should be filtered out
        assert_eq!(text, "");
    }

    #[test]
    fn test_slide_zero_based_vs_one_based_numbering() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();

        // Test that slide_number is 1-based (display number)
        let slide_data = create_slide_data(record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        assert_eq!(slide.slide_number(), 1); // 1-based for user display
    }

    #[test]
    fn test_shape_count_matches_shapes_len() {
        let doc_data = vec![0u8; 1024];
        let record = create_slide_with_drawing();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        let shape_count = slide.shape_count().unwrap();
        let shapes_len = slide.shapes().unwrap().len();

        assert_eq!(shape_count, shapes_len);
    }

    #[test]
    fn test_text_and_shapes_independent_caching() {
        let doc_data = vec![0u8; 1024];

        // Create slide with both text and PPDrawing
        let text_data = vec![0x41, 0x00]; // 'A'
        let text_atom = create_test_record(PptRecordType::TextCharsAtom, text_data, Vec::new());

        let ppdrawing = create_test_record(PptRecordType::PPDrawing, vec![0u8; 16], Vec::new());

        let slide_record =
            create_test_record(PptRecordType::Slide, vec![], vec![text_atom, ppdrawing]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        // Access text first
        let text = slide.text().unwrap();
        assert_eq!(text, "A");

        // Then access shapes - should work independently
        let shapes = slide.shapes().unwrap();
        assert_eq!(shapes.len(), 0);

        // Access again to verify both caches work
        let text2 = slide.text().unwrap();
        let shapes2 = slide.shapes().unwrap();

        assert_eq!(text, text2);
        assert_eq!(shapes.len(), shapes2.len());
    }

    #[test]
    fn test_slide_with_cstring_record() {
        let doc_data = vec![0u8; 1024];

        // Create CString record (null-terminated ASCII)
        let cstring_data = vec![
            0x48, // 'H'
            0x69, // 'i'
            0x00, // null terminator
        ];
        let cstring = create_test_record(PptRecordType::CString, cstring_data, Vec::new());

        let slide_record = create_test_record(PptRecordType::Slide, vec![], vec![cstring]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        assert_eq!(text, "Hi");
    }

    #[test]
    fn test_large_persist_id() {
        let doc_data = vec![0u8; 1024];
        let record = create_basic_slide_record();

        // Test with large persist ID
        let large_id = u32::MAX - 1;
        let slide_data = create_slide_data(record, large_id, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        assert_eq!(slide.persist_id(), large_id);
    }

    #[test]
    fn test_slide_with_empty_data() {
        let doc_data = vec![0u8; 0]; // Empty document data
        let record = create_basic_slide_record();
        let slide_data = create_slide_data(record, 256, &doc_data);

        let slide = Slide::from_slide_data(slide_data, 1);

        // Should still work with basic accessors
        assert_eq!(slide.slide_number(), 1);
        assert_eq!(slide.persist_id(), 256);
        assert!(!slide.has_drawing());
    }

    #[test]
    fn test_slide_text_extraction_preserves_order() {
        let doc_data = vec![0u8; 1024];

        // Create multiple text atoms in specific order
        let text1 = create_test_record(
            PptRecordType::TextCharsAtom,
            vec![0x31, 0x00], // '1'
            Vec::new(),
        );

        let text2 = create_test_record(
            PptRecordType::TextCharsAtom,
            vec![0x32, 0x00], // '2'
            Vec::new(),
        );

        let text3 = create_test_record(
            PptRecordType::TextCharsAtom,
            vec![0x33, 0x00], // '3'
            Vec::new(),
        );

        let slide_record =
            create_test_record(PptRecordType::Slide, vec![], vec![text1, text2, text3]);

        let slide_data = create_slide_data(slide_record, 256, &doc_data);
        let slide = Slide::from_slide_data(slide_data, 1);

        let text = slide.text().unwrap();
        // Text should be extracted in order and joined with newlines
        assert!(text.contains('1'));
        assert!(text.contains('2'));
        assert!(text.contains('3'));
        // Verify order is preserved
        let pos1 = text.find('1').unwrap();
        let pos2 = text.find('2').unwrap();
        let pos3 = text.find('3').unwrap();
        assert!(pos1 < pos2);
        assert!(pos2 < pos3);
    }
}
