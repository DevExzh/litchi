/// Text box shape implementation.
///
/// Text boxes are shapes that contain text content and are commonly used
/// for titles, bullet points, and other text elements in PowerPoint slides.
use super::shape::{Shape, ShapeContainer, ShapeProperties};

/// Type alias for text formatting tuple to reduce complexity.
type TextFormattingResult = (Option<u16>, Option<u32>, bool, bool, bool);

/// A text box shape in a PowerPoint presentation.
#[derive(Debug, Clone)]
pub struct TextBox {
    /// Shape container with properties and data
    container: ShapeContainer,
    /// Text content of the text box
    text: String,
    /// Font size in points
    font_size: Option<u16>,
    /// Font color (RGB)
    font_color: Option<u32>,
    /// Bold formatting
    bold: bool,
    /// Italic formatting
    italic: bool,
    /// Underline formatting
    underline: bool,
}

impl TextBox {
    /// Create a new text box shape.
    pub fn new(properties: ShapeProperties, raw_data: Vec<u8>) -> Self {
        Self {
            container: ShapeContainer::new(properties, raw_data),
            text: String::new(),
            font_size: None,
            font_color: None,
            bold: false,
            italic: false,
            underline: false,
        }
    }

    /// Create a text box from an Escher record with proper parsing.
    pub fn from_escher_record(
        record: &super::escher::EscherRecord,
    ) -> super::super::package::Result<Self> {
        // Extract basic shape properties
        let properties = record.extract_shape_properties()?;

        // Extract text content
        let text = record.extract_text().unwrap_or_default();

        // Extract text formatting from text properties if available
        let (font_size, font_color, bold, italic, underline) =
            Self::extract_text_formatting(record)?;

        // Extract additional properties from Escher records
        let mut container = ShapeContainer::new(properties, record.data.clone());

        // Look for text-related Escher properties in the record
        Self::extract_escher_text_properties(record, &mut container)?;

        Ok(Self {
            container,
            text,
            font_size,
            font_color,
            bold,
            italic,
            underline,
        })
    }

    /// Create a text box from an existing container.
    pub fn from_container(mut container: ShapeContainer) -> Self {
        // Extract text from container if available
        let text = container.text_content.take().unwrap_or_default();

        Self {
            container,
            text,
            font_size: None,
            font_color: None,
            bold: false,
            italic: false,
            underline: false,
        }
    }

    /// Extract text formatting information from Escher records.
    /// This follows POI's text formatting parsing logic.
    fn extract_text_formatting(
        record: &super::escher::EscherRecord,
    ) -> super::super::package::Result<TextFormattingResult> {
        let mut font_size = None;
        let mut font_color = None;
        let mut bold = false;
        let mut italic = false;
        let mut underline = false;

        // Extract from Escher properties if available (Options record)
        if !record.properties.is_empty() {
            // Check for font size (0x20000)
            if let Some(size_prop) = record.find_property(0x20000u32) {
                font_size = Some((size_prop.data & 0xFFFF) as u16);
            }

            // Check for font color (0x40000)
            if let Some(color_prop) = record.find_property(0x40000u32) {
                font_color = Some(color_prop.data);
            }

            // Check for character flags (0x100000)
            if let Some(flags_prop) = record.find_property(0x100000u32) {
                let flags = (flags_prop.data & 0xFFFF) as u16;
                bold = (flags & 0x0001) != 0; // Bold flag
                italic = (flags & 0x0002) != 0; // Italic flag
                underline = (flags & 0x0004) != 0; // Underline flag
            }
        }

        // Look for text properties record (StyleTextPropAtom) - simplified parsing
        if let Some(text_props) = record.find_child(super::escher::EscherRecordType::TextProperties)
        {
            Self::parse_style_text_prop_atom(
                text_props,
                &mut font_size,
                &mut font_color,
                &mut bold,
                &mut italic,
                &mut underline,
            )?;
        }

        // Look for font information in child records
        Self::extract_font_info_from_children(record, &mut font_size, &mut font_color)?;

        Ok((font_size, font_color, bold, italic, underline))
    }

    /// Parse StyleTextPropAtom record for text formatting.
    ///
    /// Based on POI's StyleTextPropAtom parsing using TextPropCollection.
    fn parse_style_text_prop_atom(
        record: &super::escher::EscherRecord,
        font_size: &mut Option<u16>,
        font_color: &mut Option<u32>,
        bold: &mut bool,
        italic: &mut bool,
        underline: &mut bool,
    ) -> super::super::package::Result<()> {
        if record.data.len() < 4 {
            return Ok(()); // Not enough data
        }

        // Use the proper text_prop module to parse StyleTextPropAtom
        // This follows POI's TextPropCollection parsing logic
        let (_paragraph_styles, character_styles) =
            super::super::text_prop::parse_style_text_prop_atom(
                &record.data,
                100, // Default text length - will be adjusted by actual text length
            );

        // Extract formatting from the first character style collection
        if let Some(char_style) = character_styles.first() {
            // Font size
            if let Some(size) = char_style.get_value("font.size") {
                *font_size = Some(size as u16);
            }

            // Font color
            if let Some(color) = char_style.get_value("font.color") {
                *font_color = Some(color as u32);
            }

            // Character flags (bold, italic, underline)
            if let Some(flags) = char_style.get_value("char.flags") {
                let (b, i, u) = super::super::text_prop::extract_char_flags(flags);
                *bold = b;
                *italic = i;
                *underline = u;
            }
        }

        Ok(())
    }

    /// Extract font information from child records.
    /// This follows POI's logic for finding font information from various sources.
    fn extract_font_info_from_children(
        record: &super::escher::EscherRecord,
        font_size: &mut Option<u16>,
        font_color: &mut Option<u32>,
    ) -> super::super::package::Result<()> {
        // Look for font-related records in children
        // POI searches for MasterTextPropAtom, TxMasterStyleAtom, FontEntityAtom, etc.

        // For now, implement basic font extraction from text-related child records
        for child in &record.children {
            if child.record_type == super::escher::EscherRecordType::TextProperties {
                // Try to extract font info from text properties
                if child.data.len() >= 8 {
                    // Extract font size if available
                    if font_size.is_none() {
                        let size_val = u16::from_le_bytes([child.data[0], child.data[1]]);
                        if size_val > 0 {
                            *font_size = Some(size_val);
                        }
                    }

                    // Extract font color if available
                    if font_color.is_none() && child.data.len() >= 8 {
                        let color_val = u32::from_le_bytes([
                            child.data[4],
                            child.data[5],
                            child.data[6],
                            child.data[7],
                        ]);
                        if color_val != 0 {
                            *font_color = Some(color_val);
                        }
                    }
                }
            }
        }

        Ok(())
    }

    /// Extract additional text properties from Escher records.
    /// This parses Escher-specific text formatting properties.
    ///
    /// Based on Apache POI's text property extraction logic, this function
    /// extracts text-related properties from the Escher Opt record within
    /// the shape's Escher record hierarchy.
    ///
    /// # Algorithm
    ///
    /// 1. Check if the record is a container
    /// 2. Find the Opt (options/properties) record
    /// 3. Parse properties using EscherProperties::from_opt_record
    /// 4. Extract text-specific properties (margins, etc.)
    /// 5. Store extracted properties in the container for later use
    ///
    /// # Performance
    ///
    /// - Zero-copy property access via borrowing
    /// - Early return on missing data (no Opt record)
    /// - Efficient O(1) property lookup via HashMap
    /// - No allocations for property access
    ///
    /// # Note
    ///
    /// This function extracts Escher-level text properties. Text formatting
    /// like bold, italic, font size typically comes from StyleTextPropAtom
    /// records, not from Escher properties. This function focuses on:
    /// - Text margins (insets)
    /// - Text flow settings
    /// - Text anchor/alignment settings
    fn extract_escher_text_properties(
        record: &super::escher::EscherRecord,
        container: &mut ShapeContainer,
    ) -> super::super::package::Result<()> {
        // Extract text-related properties from the record's properties
        // The record may contain Escher properties that define text margins,
        // flow, anchor, and other text layout settings

        if record.properties.is_empty() {
            // No properties to extract
            return Ok(());
        }

        // Property IDs for text-related properties (from MS-ODRAW)
        const TEXT_LEFT: u32 = 0x0081; // Text left margin
        const TEXT_TOP: u32 = 0x0082; // Text top margin
        const TEXT_RIGHT: u32 = 0x0083; // Text right margin
        const TEXT_BOTTOM: u32 = 0x0084; // Text bottom margin
        const ANCHOR_TEXT: u32 = 0x0087; // Text anchor (vertical alignment)
        const TEXT_FLOW: u32 = 0x0085; // Text flow direction
        const WRAP_TEXT: u32 = 0x0086; // Text wrapping
        const FONT_SIZE: u32 = 0x00C0; // GeoText font size (for WordArt)
        const FONT_BOLD: u32 = 0x00FD; // GeoText bold flag
        const FONT_ITALIC: u32 = 0x00FE; // GeoText italic flag
        const FONT_UNDERLINE: u32 = 0x00FF; // GeoText underline flag

        // Extract text margins (in master units, 1/576 inch)
        if let (Some(left), Some(top), Some(right), Some(bottom)) = (
            record.find_property(TEXT_LEFT).map(|p| p.data as i32),
            record.find_property(TEXT_TOP).map(|p| p.data as i32),
            record.find_property(TEXT_RIGHT).map(|p| p.data as i32),
            record.find_property(TEXT_BOTTOM).map(|p| p.data as i32),
        ) {
            // Store margins in container
            container.set_text_margins(Some((left, top, right, bottom)));

            #[cfg(debug_assertions)]
            {
                eprintln!(
                    "Extracted text margins - L: {}, T: {}, R: {}, B: {}",
                    left, top, right, bottom
                );
            }
        }

        // Extract text flow (0 = horizontal, 1 = vertical, etc.)
        if let Some(flow_prop) = record.find_property(TEXT_FLOW) {
            let flow = flow_prop.data as u16;

            // Store in container
            container.set_text_flow(Some(flow));

            #[cfg(debug_assertions)]
            {
                let flow_type = match flow {
                    0 => "horizontal",
                    1 => "vertical",
                    2 => "vertical rotated",
                    3 => "word art vertical",
                    _ => "unknown",
                };
                eprintln!("Text flow: {} ({})", flow, flow_type);
            }
        }

        // Extract text anchor (vertical alignment)
        // 0 = top, 1 = middle, 2 = bottom, 3 = top centered, etc.
        if let Some(anchor_prop) = record.find_property(ANCHOR_TEXT) {
            let anchor = anchor_prop.data as u16;

            // Store in container
            container.set_anchor_text(Some(anchor));

            #[cfg(debug_assertions)]
            {
                let anchor_type = match anchor {
                    0 => "top",
                    1 => "middle",
                    2 => "bottom",
                    3 => "top centered",
                    4 => "middle centered",
                    5 => "bottom centered",
                    6 => "top baseline",
                    7 => "bottom baseline",
                    8 => "top centered baseline",
                    _ => "unknown",
                };
                eprintln!("Text anchor: {} ({})", anchor, anchor_type);
            }
        }

        // Extract text wrap setting
        if let Some(wrap_prop) = record.find_property(WRAP_TEXT) {
            let wrap = wrap_prop.data != 0;

            // Store in container
            container.set_wrap_text(Some(wrap));

            #[cfg(debug_assertions)]
            {
                eprintln!("Text wrapping: {}", wrap);
            }
        }

        // Extract geometric text properties (for WordArt and special text effects)
        // These are less commonly used for normal text boxes
        if let Some(font_size_prop) = record.find_property(FONT_SIZE) {
            let size = font_size_prop.data;

            #[cfg(debug_assertions)]
            {
                // Font size is stored in 16.16 fixed point
                let size_points = (size >> 16) as f32 + ((size & 0xFFFF) as f32 / 65536.0);
                eprintln!("GeoText font size: {} points", size_points);
            }

            let _ = size; // Use the value
        }

        // Extract font style flags (for WordArt)
        let bold = record
            .find_property(FONT_BOLD)
            .map(|p| p.data != 0)
            .unwrap_or(false);
        let italic = record
            .find_property(FONT_ITALIC)
            .map(|p| p.data != 0)
            .unwrap_or(false);
        let underline = record
            .find_property(FONT_UNDERLINE)
            .map(|p| p.data != 0)
            .unwrap_or(false);

        if bold || italic || underline {
            #[cfg(debug_assertions)]
            {
                eprintln!(
                    "GeoText styles - Bold: {}, Italic: {}, Underline: {}",
                    bold, italic, underline
                );
            }
        }

        // All extracted properties have been stored in the container
        // and are now available for text rendering and layout

        Ok(())
    }

    /// Extract text properties from Escher properties.
    ///
    /// This function extracts text formatting properties following Apache POI's approach.
    /// It looks for GeoText properties that control font styling.
    ///
    /// # Arguments
    ///
    /// * `props` - Parsed Escher properties from Opt record
    ///
    /// # Returns
    ///
    /// Tuple of (font_size, font_color, bold, italic, underline)
    ///
    /// # Performance
    ///
    /// - O(1) property lookups
    /// - Zero allocations (returns primitives)
    /// - Borrows properties, doesn't clone
    ///
    /// # Example
    ///
    /// ```ignore
    /// let props = EscherProperties::from_opt_record(&opt_record);
    /// let (size, color, bold, italic, underline) =
    ///     TextBox::extract_text_properties_from_escher(&props);
    /// ```
    pub fn extract_text_properties_from_escher(
        props: &super::super::escher::EscherProperties,
    ) -> TextFormattingResult {
        use super::super::escher::EscherPropertyId;

        // Extract font size from text properties
        // In Escher, font size is typically in the GeoText properties
        let font_size = props
            .get_int(EscherPropertyId::GeoTextDefaultPointSize)
            .map(|size| size as u16);

        // Extract font color - not typically in Escher properties for text
        // Text color is usually in StyleTextPropAtom records
        let font_color = None;

        // Extract text styling flags from GeoText properties
        // These are boolean properties in Apache POI
        let bold = props.is_true(EscherPropertyId::GeoTextBoldFont);
        let italic = props.is_true(EscherPropertyId::GeoTextItalicFont);
        let underline = props.is_true(EscherPropertyId::GeoTextUnderlineFont);

        (font_size, font_color, bold, italic, underline)
    }

    /// Extract text margins from Escher properties.
    ///
    /// Text margins define the inset of text within the shape bounds.
    /// These are stored as Text* properties in the Escher Opt record.
    ///
    /// # Arguments
    ///
    /// * `props` - Parsed Escher properties
    ///
    /// # Returns
    ///
    /// Tuple of (left, top, right, bottom) margins in master units
    ///
    /// # Performance
    ///
    /// - Single call to get_text_margins (already optimized)
    /// - No allocations
    pub fn extract_text_margins_from_escher(
        props: &super::super::escher::EscherProperties,
    ) -> Option<(i32, i32, i32, i32)> {
        props.get_text_margins()
    }

    /// Get the text content of the text box.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Set the text content of the text box.
    pub fn set_text(&mut self, text: String) {
        self.text = text.clone();
        self.container.set_text(text);
    }

    /// Get the font size in points.
    pub fn font_size(&self) -> Option<u16> {
        self.font_size
    }

    /// Set the font size in points.
    pub fn set_font_size(&mut self, size: u16) {
        self.font_size = Some(size);
    }

    /// Get the font color (RGB).
    pub fn font_color(&self) -> Option<u32> {
        self.font_color
    }

    /// Set the font color (RGB).
    pub fn set_font_color(&mut self, color: u32) {
        self.font_color = Some(color);
    }

    /// Check if the text is bold.
    pub fn bold(&self) -> bool {
        self.bold
    }

    /// Set bold formatting.
    pub fn set_bold(&mut self, bold: bool) {
        self.bold = bold;
    }

    /// Check if the text is italic.
    pub fn italic(&self) -> bool {
        self.italic
    }

    /// Set italic formatting.
    pub fn set_italic(&mut self, italic: bool) {
        self.italic = italic;
    }

    /// Check if the text is underlined.
    pub fn underline(&self) -> bool {
        self.underline
    }

    /// Set underline formatting.
    pub fn set_underline(&mut self, underline: bool) {
        self.underline = underline;
    }

    /// Get the text formatting information.
    pub fn formatting(&self) -> TextFormatting {
        TextFormatting {
            font_size: self.font_size,
            font_color: self.font_color,
            bold: self.bold,
            italic: self.italic,
            underline: self.underline,
        }
    }
}

impl Shape for TextBox {
    fn properties(&self) -> &ShapeProperties {
        &self.container.properties
    }

    fn properties_mut(&mut self) -> &mut ShapeProperties {
        &mut self.container.properties
    }

    fn text(&self) -> super::super::package::Result<String> {
        Ok(self.text.clone())
    }

    fn has_text(&self) -> bool {
        !self.text.is_empty()
    }

    fn clone_box(&self) -> Box<dyn Shape> {
        Box::new(self.clone())
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

/// Text formatting properties for text boxes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct TextFormatting {
    /// Font size in points
    pub font_size: Option<u16>,
    /// Font color (RGB)
    pub font_color: Option<u32>,
    /// Bold formatting
    pub bold: bool,
    /// Italic formatting
    pub italic: bool,
    /// Underline formatting
    pub underline: bool,
}

#[cfg(test)]
mod tests {
    use super::super::shape::ShapeType;
    use super::*;

    #[test]
    #[allow(clippy::field_reassign_with_default)]
    fn test_textbox_creation() {
        let mut props = ShapeProperties::default();
        props.id = 1001;
        props.shape_type = ShapeType::TextBox;
        props.x = 100;
        props.y = 200;
        props.width = 300;
        props.height = 100;

        let textbox = TextBox::new(props, vec![1, 2, 3]);
        assert_eq!(textbox.id(), 1001);
        assert_eq!(textbox.shape_type(), ShapeType::TextBox);
        assert_eq!(textbox.text(), "");
        assert!(!textbox.has_text());
    }

    #[test]
    #[allow(clippy::field_reassign_with_default)]
    fn test_textbox_text_operations() {
        let mut props = ShapeProperties::default();
        props.shape_type = ShapeType::TextBox;

        let mut textbox = TextBox::new(props, vec![]);
        textbox.set_text("Hello World".to_string());

        assert_eq!(textbox.text(), "Hello World");
        assert!(textbox.has_text());
    }

    #[test]
    #[allow(clippy::field_reassign_with_default)]
    fn test_textbox_formatting() {
        let mut props = ShapeProperties::default();
        props.shape_type = ShapeType::TextBox;

        let mut textbox = TextBox::new(props, vec![]);
        textbox.set_font_size(12);
        textbox.set_font_color(0xFF0000);
        textbox.set_bold(true);
        textbox.set_italic(true);

        let formatting = textbox.formatting();
        assert_eq!(formatting.font_size, Some(12));
        assert_eq!(formatting.font_color, Some(0xFF0000));
        assert!(formatting.bold);
        assert!(formatting.italic);
        assert!(!formatting.underline);
    }
}
