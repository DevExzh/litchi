/// Text box shape implementation.
///
/// Text boxes are shapes that contain text content and are commonly used
/// for titles, bullet points, and other text elements in PowerPoint slides.
use super::shape::{Shape, ShapeProperties, ShapeContainer};

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
    pub fn from_escher_record(record: &super::escher::EscherRecord) -> super::super::package::Result<Self> {
        // Extract basic shape properties
        let properties = record.extract_shape_properties()?;

        // Extract text content
        let text = record.extract_text().unwrap_or_default();

        // Extract text formatting from text properties if available
        let (font_size, font_color, bold, italic, underline) = Self::extract_text_formatting(record)?;

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
    fn extract_text_formatting(record: &super::escher::EscherRecord) -> super::super::package::Result<TextFormattingResult> {
        let mut font_size = None;
        let mut font_color = None;
        let mut bold = false;
        let mut italic = false;
        let mut underline = false;

        // Extract from Escher properties if available (Options record)
        if !record.properties.is_empty() {
            let _prop_values = record.extract_property_values();

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
                bold = (flags & 0x0001) != 0;      // Bold flag
                italic = (flags & 0x0002) != 0;    // Italic flag
                underline = (flags & 0x0004) != 0; // Underline flag
            }
        }

        // Look for text properties record (StyleTextPropAtom) - simplified parsing
        if let Some(text_props) = record.find_child(super::escher::EscherRecordType::TextProperties) {
            Self::parse_style_text_prop_atom(text_props, &mut font_size, &mut font_color, &mut bold, &mut italic, &mut underline)?;
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
        let (_paragraph_styles, character_styles) = super::super::text_prop::parse_style_text_prop_atom(
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
                        let color_val = u32::from_le_bytes([child.data[4], child.data[5], child.data[6], child.data[7]]);
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
    fn extract_escher_text_properties(_record: &super::escher::EscherRecord, _container: &mut ShapeContainer) -> super::super::package::Result<()> {
        // In a full implementation, this would parse Escher text properties
        // such as font size, color, alignment, etc. from the Escher record data
        // POI does this through EscherOptRecord and related property parsing

        Ok(())
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
    use super::*;
    use super::super::shape::ShapeType;

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
