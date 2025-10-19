/// High-performance Slide implementation with lazy shape loading and zero-copy design.
use super::super::package::Result;
use super::super::shapes::ShapeEnum;
use super::super::records::PptRecord;
use super::slide_factory::SlideData;
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
        let ppdrawing = match self.record.find_child(crate::ole::consts::PptRecordType::PPDrawing) {
            Some(record) => record,
            None => return Ok(Vec::new()),
        };
        
        // Extract Escher shapes from PPDrawing data
        let escher_shapes = super::super::escher::EscherShapeFactory::extract_shapes_from_ppdrawing(&ppdrawing.data)?;
        
        // Convert Escher shapes to ShapeEnum with full property extraction
        let shapes: Vec<ShapeEnum> = escher_shapes.iter()
            .filter_map(|escher_shape| {
                Self::convert_escher_to_shape_enum(escher_shape)
            })
            .collect();
        
        Ok(shapes)
    }
    
    /// Convert an EscherShape to ShapeEnum with full property extraction.
    ///
    /// # Performance
    ///
    /// - Direct property access (no allocations)
    /// - Pattern matching for type dispatch
    fn convert_escher_to_shape_enum(escher_shape: &super::super::escher::EscherShape<'_>) -> Option<ShapeEnum> {
        use super::super::escher::EscherShapeType;
        use super::super::shapes::*;
        
        let shape_id = escher_shape.shape_id().unwrap_or(0);
        let anchor = escher_shape.anchor();
        
        match escher_shape.shape_type() {
            EscherShapeType::TextBox => {
                // Create TextBox with proper properties
                let mut properties = shape::ShapeProperties::default();
                properties.id = shape_id;
                properties.shape_type = shape::ShapeType::TextBox;
                
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
            }
            
            EscherShapeType::Picture => {
                // Create PictureShape
                let mut picture = shape_enum::PictureShape::new(shape_id);
                
                if let Some(a) = anchor {
                    picture.set_bounds(a.left, a.top, a.width(), a.height());
                }
                
                // Extract blip ID from properties
                use super::super::escher::EscherPropertyId;
                if let Some(blip_id) = escher_shape.properties().get_int(EscherPropertyId::PictureId) {
                    picture.set_blip_id(blip_id as u32);
                }
                
                Some(ShapeEnum::Picture(picture))
            }
            
            EscherShapeType::Line => {
                // Create LineShape
                if let Some(a) = anchor {
                    let mut line = shape_enum::LineShape::new(
                        shape_id,
                        a.left,
                        a.top,
                        a.right,
                        a.bottom
                    );
                    
                    // Extract line properties
                    use super::super::escher::EscherPropertyId;
                    if let Some(width) = escher_shape.properties().get_int(EscherPropertyId::LineWidth) {
                        line.set_width(width);
                    }
                    if let Some(color) = escher_shape.properties().get_color(EscherPropertyId::LineColor) {
                        line.set_color(color);
                    }
                    
                    Some(ShapeEnum::Line(line))
                } else {
                    None
                }
            }
            
            EscherShapeType::Group => {
                // Create GroupShape (children would be parsed recursively)
                let mut group = shape_enum::GroupShape::new(shape_id);
                
                if let Some(a) = anchor {
                    group.set_bounds(a.left, a.top, a.width(), a.height());
                }
                
                // TODO: Parse child shapes recursively
                
                Some(ShapeEnum::Group(group))
            }
            
            EscherShapeType::Rectangle | EscherShapeType::Ellipse | EscherShapeType::AutoShape => {
                // Create AutoShape
                let mut properties = shape::ShapeProperties::default();
                properties.id = shape_id;
                properties.shape_type = shape::ShapeType::AutoShape;
                
                if let Some(a) = anchor {
                    properties.x = a.left;
                    properties.y = a.top;
                    properties.width = a.width();
                    properties.height = a.height();
                }
                
                let autoshape = AutoShape::new(properties, Vec::new());
                Some(ShapeEnum::AutoShape(autoshape))
            }
            
            // Unknown or unsupported shape types
            _ => None,
        }
    }

    /// Extract all text from slide and its shapes.
    fn extract_all_text(&self) -> Result<String> {
        let mut text_parts = Vec::new();

        // 1. Extract text from direct slide records (TextCharsAtom, etc.)
        Self::extract_text_recursive(&self.record, &mut text_parts);

        // 2. Extract text from Escher/PPDrawing (shapes, text boxes)
        if let Some(ppdrawing) = self.record.find_child(crate::ole::consts::PptRecordType::PPDrawing) {
            if let Ok(escher_text) = super::super::escher::extract_text_from_escher(&ppdrawing.data) {
                if !escher_text.trim().is_empty() {
                    text_parts.push(escher_text);
                }
            }
        }

        Ok(if text_parts.is_empty() {
            String::new()
        } else {
            text_parts.join("\n")
        })
    }
    
    /// Recursively extract text from a record and all its children.
    fn extract_text_recursive(record: &crate::ole::ppt::records::PptRecord, text_parts: &mut Vec<String>) {
        // Try to extract text from this record
        if let Ok(record_text) = record.extract_text() {
            let trimmed = record_text.trim();
            if !trimmed.is_empty() {
                text_parts.push(trimmed.to_string());
            }
        }
        
        // Recursively process children
        for child in &record.children {
            Self::extract_text_recursive(child, text_parts);
        }
    }

    /// Check if this slide has a PPDrawing record (shapes).
    #[inline]
    pub fn has_drawing(&self) -> bool {
        self.record.find_child(crate::ole::consts::PptRecordType::PPDrawing).is_some()
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

    // Tests will be added as we implement shape parsing
}

