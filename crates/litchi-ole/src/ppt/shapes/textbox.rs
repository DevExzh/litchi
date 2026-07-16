/// Text box shape implementation.
///
/// Text boxes are shapes that contain text content and are commonly used
/// for titles, bullet points, and other text elements in PowerPoint slides.
use super::shape::{Shape, ShapeContainer, ShapeProperties};
use crate::ppt::text_run::{TextRun, TextRunFormatting};

/// Type alias for text formatting tuple to reduce complexity.
type TextFormattingResult = (Option<u16>, Option<u32>, bool, bool, bool);

/// A text box shape in a PowerPoint presentation.
///
/// Uses lifetime parameter `'a` to enable zero-copy parsing when the shape
/// data can be borrowed from a larger buffer.
#[derive(Debug, Clone)]
pub struct TextBox<'a> {
    /// Shape container with properties and data
    container: ShapeContainer<'a>,
    /// Text content of the text box
    text: String,
    /// Character-formatting runs from the embedded `StyleTextPropAtom`
    runs: Vec<TextRun>,
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

impl<'a> TextBox<'a> {
    /// Create a new text box shape with owned data.
    pub fn new(properties: ShapeProperties, raw_data: Vec<u8>) -> Self {
        Self {
            container: ShapeContainer::new(properties, raw_data),
            text: String::new(),
            runs: Vec::new(),
            font_size: None,
            font_color: None,
            bold: false,
            italic: false,
            underline: false,
        }
    }

    /// Create a text box from an Escher record with zero-copy parsing.
    pub fn from_escher_record(
        record: &'a super::escher::EscherRecord<'a>,
    ) -> super::super::package::Result<Self> {
        // Extract basic shape properties
        let properties = record.extract_shape_properties()?;

        let (text, runs) = if let Some(textbox_record) =
            Self::find_descendant(record, super::escher::EscherRecordType::ClientTextbox)
        {
            let wrapper = crate::ppt::EscherTextboxWrapper::new(textbox_record.data.to_vec())?;
            (wrapper.text().to_string(), wrapper.runs().to_vec())
        } else {
            (String::new(), Vec::new())
        };
        let (font_size, font_color, bold, italic, underline) = Self::formatting_from_runs(&runs);

        // Extract additional properties from Escher records with zero-copy
        let mut container = ShapeContainer::new_borrowed(properties, &record.data);

        // Look for text-related Escher properties in the record
        Self::extract_escher_text_properties(record, &mut container)?;

        Ok(Self {
            container,
            text,
            runs,
            font_size,
            font_color,
            bold,
            italic,
            underline,
        })
    }

    /// Create a text box from an existing container.
    pub fn from_container(mut container: ShapeContainer<'a>) -> Self {
        // Extract text from container if available
        let text = container.text_content.take().unwrap_or_default();
        let runs = if text.is_empty() {
            Vec::new()
        } else {
            vec![TextRun::new(text.clone(), 0)]
        };

        Self {
            container,
            text,
            runs,
            font_size: None,
            font_color: None,
            bold: false,
            italic: false,
            underline: false,
        }
    }

    fn find_descendant<'record>(
        record: &'record super::escher::EscherRecord<'a>,
        record_type: super::escher::EscherRecordType,
    ) -> Option<&'record super::escher::EscherRecord<'a>> {
        if record.record_type == record_type {
            return Some(record);
        }
        record
            .children
            .iter()
            .find_map(|child| Self::find_descendant(child, record_type))
    }

    fn formatting_from_runs(runs: &[TextRun]) -> TextFormattingResult {
        let formatting = runs
            .first()
            .map(|run| &run.formatting)
            .cloned()
            .unwrap_or_default();
        (
            formatting.font_size,
            formatting.font_color,
            formatting.bold,
            formatting.italic,
            formatting.underline,
        )
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
        record: &super::escher::EscherRecord<'a>,
        container: &mut ShapeContainer<'a>,
    ) -> super::super::package::Result<()> {
        let Some(options) = Self::find_descendant(record, super::escher::EscherRecordType::Opt)
            .or_else(|| {
                Self::find_descendant(record, super::escher::EscherRecordType::SecondaryOpt)
            })
            .or_else(|| {
                Self::find_descendant(record, super::escher::EscherRecordType::TertiaryOpt)
            })
        else {
            return Ok(());
        };

        // Property IDs for text-related properties (from MS-ODRAW)
        const TEXT_LEFT: u32 = 0x0081; // Text left margin
        const TEXT_TOP: u32 = 0x0082; // Text top margin
        const TEXT_RIGHT: u32 = 0x0083; // Text right margin
        const TEXT_BOTTOM: u32 = 0x0084; // Text bottom margin
        const WRAP_TEXT: u32 = 0x0085; // MSOWRAPMODE
        const ANCHOR_TEXT: u32 = 0x0087; // Text anchor (vertical alignment)
        const TEXT_FLOW: u32 = 0x0088; // Text flow direction

        // Extract text margins in EMUs.
        if let (Some(left), Some(top), Some(right), Some(bottom)) = (
            options.find_property(TEXT_LEFT).map(|p| p.data as i32),
            options.find_property(TEXT_TOP).map(|p| p.data as i32),
            options.find_property(TEXT_RIGHT).map(|p| p.data as i32),
            options.find_property(TEXT_BOTTOM).map(|p| p.data as i32),
        ) {
            container.set_text_margins(Some((left, top, right, bottom)));
        }

        if let Some(flow_prop) = options.find_property(TEXT_FLOW) {
            container.set_text_flow(Some(flow_prop.data as u16));
        }

        // Extract text anchor (vertical alignment)
        // 0 = top, 1 = middle, 2 = bottom, 3 = top centered, etc.
        if let Some(anchor_prop) = options.find_property(ANCHOR_TEXT) {
            container.set_anchor_text(Some(anchor_prop.data as u16));
        }

        if let Some(wrap_prop) = options.find_property(WRAP_TEXT) {
            container.set_wrap_text(Some(wrap_prop.data as u16));
        }

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

    /// Get the character-formatting runs in document order.
    pub fn runs(&self) -> &[TextRun] {
        &self.runs
    }

    /// Set the text content of the text box.
    pub fn set_text(&mut self, text: String) {
        self.text = text.clone();
        self.runs = if text.is_empty() {
            Vec::new()
        } else {
            vec![TextRun::with_formatting(
                text.clone(),
                0,
                self.current_run_formatting(),
            )]
        };
        self.container.set_text(text);
    }

    /// Get the font size in points.
    pub fn font_size(&self) -> Option<u16> {
        self.font_size
    }

    /// Set the font size in points.
    pub fn set_font_size(&mut self, size: u16) {
        self.font_size = Some(size);
        for run in &mut self.runs {
            run.formatting.font_size = Some(size);
        }
    }

    /// Get the font color (RGB).
    pub fn font_color(&self) -> Option<u32> {
        self.font_color
    }

    /// Set the font color (RGB).
    pub fn set_font_color(&mut self, color: u32) {
        self.font_color = Some(color);
        for run in &mut self.runs {
            run.formatting.font_color = Some(color);
        }
    }

    /// Check if the text is bold.
    pub fn bold(&self) -> bool {
        self.bold
    }

    /// Set bold formatting.
    pub fn set_bold(&mut self, bold: bool) {
        self.bold = bold;
        for run in &mut self.runs {
            run.formatting.bold = bold;
        }
    }

    /// Check if the text is italic.
    pub fn italic(&self) -> bool {
        self.italic
    }

    /// Set italic formatting.
    pub fn set_italic(&mut self, italic: bool) {
        self.italic = italic;
        for run in &mut self.runs {
            run.formatting.italic = italic;
        }
    }

    /// Check if the text is underlined.
    pub fn underline(&self) -> bool {
        self.underline
    }

    /// Set underline formatting.
    pub fn set_underline(&mut self, underline: bool) {
        self.underline = underline;
        for run in &mut self.runs {
            run.formatting.underline = underline;
        }
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

    /// Get text inset values `(left, top, right, bottom)` in EMUs.
    pub fn text_margins(&self) -> Option<(i32, i32, i32, i32)> {
        self.container.text_margins()
    }

    /// Get the raw `MSOTXFL` text-flow value.
    pub fn text_flow(&self) -> Option<u16> {
        self.container.text_flow
    }

    /// Get the raw `MSOWRAPMODE` wrapping value.
    pub fn wrap_mode(&self) -> Option<u16> {
        self.container.wrap_text
    }

    /// Whether the wrapping mode allows wrapping within the shape.
    pub fn word_wrap_enabled(&self) -> Option<bool> {
        self.container.word_wrap_enabled()
    }

    /// Get the raw `MSOANCHOR` vertical text anchor value.
    pub fn text_anchor(&self) -> Option<u16> {
        self.container.anchor_text
    }

    fn current_run_formatting(&self) -> TextRunFormatting {
        TextRunFormatting {
            font_size: self.font_size,
            font_color: self.font_color,
            bold: self.bold,
            italic: self.italic,
            underline: self.underline,
            font_name: None,
        }
    }
}

impl<'a> Shape for TextBox<'a>
where
    'a: 'static,
{
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

    fn push_record(target: &mut Vec<u8>, version: u16, instance: u16, kind: u16, data: &[u8]) {
        target.extend_from_slice(&(version | (instance << 4)).to_le_bytes());
        target.extend_from_slice(&kind.to_le_bytes());
        target.extend_from_slice(&(data.len() as u32).to_le_bytes());
        target.extend_from_slice(data);
    }

    fn formatted_textbox_record() -> Vec<u8> {
        let mut ppt_records = Vec::new();
        push_record(&mut ppt_records, 0, 0, 3999, &4u32.to_le_bytes());
        push_record(&mut ppt_records, 0, 0, 4008, b"abcd");

        let mut style = Vec::new();
        style.extend_from_slice(&5u32.to_le_bytes());
        style.extend_from_slice(&0i16.to_le_bytes());
        style.extend_from_slice(&0u32.to_le_bytes());
        style.extend_from_slice(&2u32.to_le_bytes());
        style.extend_from_slice(&0x0006_0001u32.to_le_bytes());
        style.extend_from_slice(&1i16.to_le_bytes());
        style.extend_from_slice(&18i16.to_le_bytes());
        style.extend_from_slice(&0x0011_2233i32.to_le_bytes());
        style.extend_from_slice(&3u32.to_le_bytes());
        style.extend_from_slice(&0x0006_0002u32.to_le_bytes());
        style.extend_from_slice(&2i16.to_le_bytes());
        style.extend_from_slice(&24i16.to_le_bytes());
        style.extend_from_slice(&0x0044_5566i32.to_le_bytes());
        push_record(&mut ppt_records, 0, 0, 4001, &style);

        let properties = [
            (0x0081u16, 91_440u32),
            (0x0082u16, 45_720u32),
            (0x0083u16, 91_440u32),
            (0x0084u16, 45_720u32),
            (0x0085u16, 0u32),
            (0x0087u16, 2u32),
            (0x0088u16, 1u32),
        ];
        let mut opt = Vec::new();
        for (id, value) in properties {
            opt.extend_from_slice(&id.to_le_bytes());
            opt.extend_from_slice(&value.to_le_bytes());
        }

        let mut shape = Vec::new();
        let mut sp = Vec::new();
        sp.extend_from_slice(&1001u32.to_le_bytes());
        sp.extend_from_slice(&0x0A00u32.to_le_bytes());
        push_record(&mut shape, 2, 24, 0xF00A, &sp);
        push_record(&mut shape, 3, properties.len() as u16, 0xF00B, &opt);
        push_record(&mut shape, 0xF, 0, 0xF00D, &ppt_records);

        let mut result = Vec::new();
        push_record(&mut result, 0xF, 0, 0xF004, &shape);
        result
    }

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

    #[test]
    fn parses_embedded_style_runs_and_nested_text_options() {
        let data = formatted_textbox_record();
        let (record, consumed) = super::super::escher::EscherRecord::parse(&data, 0).unwrap();
        assert_eq!(consumed, data.len());

        let textbox = TextBox::from_escher_record(&record).unwrap();

        assert_eq!(textbox.text(), "abcd");
        assert_eq!(textbox.runs().len(), 2);
        assert_eq!(textbox.runs()[0].text, "ab");
        assert_eq!(textbox.runs()[0].formatting.font_size, Some(18));
        assert_eq!(textbox.runs()[0].formatting.font_color, Some(0x0011_2233));
        assert!(textbox.runs()[0].formatting.bold);
        assert!(!textbox.runs()[0].formatting.italic);
        assert_eq!(textbox.runs()[1].text, "cd");
        assert_eq!(textbox.runs()[1].formatting.font_size, Some(24));
        assert_eq!(textbox.runs()[1].formatting.font_color, Some(0x0044_5566));
        assert!(!textbox.runs()[1].formatting.bold);
        assert!(textbox.runs()[1].formatting.italic);
        assert_eq!(textbox.formatting().font_size, Some(18));
        assert!(textbox.bold());
        assert_eq!(
            textbox.text_margins(),
            Some((91_440, 45_720, 91_440, 45_720))
        );
        assert_eq!(textbox.wrap_mode(), Some(0));
        assert_eq!(textbox.word_wrap_enabled(), Some(true));
        assert_eq!(textbox.text_anchor(), Some(2));
        assert_eq!(textbox.text_flow(), Some(1));
    }

    #[test]
    fn setters_keep_character_runs_consistent() {
        let mut textbox = TextBox::new(ShapeProperties::default(), Vec::new());
        textbox.set_bold(true);
        textbox.set_font_size(20);
        textbox.set_text("hello".to_string());
        textbox.set_italic(true);

        assert_eq!(textbox.runs().len(), 1);
        assert_eq!(textbox.runs()[0].text, "hello");
        assert_eq!(textbox.runs()[0].formatting.font_size, Some(20));
        assert!(textbox.runs()[0].formatting.bold);
        assert!(textbox.runs()[0].formatting.italic);
    }
}
