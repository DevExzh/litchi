/// Text box shape implementation.
///
/// Text boxes are shapes that contain text content and are commonly used
/// for titles, bullet points, and other text elements in PowerPoint slides.
use super::shape::{Shape, ShapeContainer, ShapeProperties};
use crate::odraw::ShapeExt as _;
use crate::text_run::{ParagraphRun, ParagraphRunFormatting, TextRun, TextRunFormatting};
use crate::{TextRuler, TextStyleExtension9, TextStyleExtension10, TextStyleExtension11};
use litchi_odraw::shape::Shape as OdrawShape;

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
    /// Paragraph-formatting runs from the embedded `StyleTextPropAtom`
    paragraph_runs: Vec<ParagraphRun>,
    /// Textbox-specific ruler overrides
    text_ruler: Option<TextRuler>,
    /// Header/footer metacharacter placeholders in the textbox
    metachars: Vec<crate::text_metachar::TextMetachar>,
    /// Outline text references tying the textbox to outline text bodies
    outline_text_refs: Vec<crate::text_si_exception::OutlineTextRef>,
    /// PowerPoint 9 picture-bullet and automatic-numbering extensions
    text_style_extension9: Option<TextStyleExtension9>,
    /// PowerPoint 10 alternate-script font extensions
    text_style_extension10: Option<TextStyleExtension10>,
    /// PowerPoint 11 smart-tag extensions
    text_style_extension11: Option<TextStyleExtension11>,
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
            paragraph_runs: Vec::new(),
            text_ruler: None,
            metachars: Vec::new(),
            outline_text_refs: Vec::new(),
            text_style_extension9: None,
            text_style_extension10: None,
            text_style_extension11: None,
            font_size: None,
            font_color: None,
            bold: false,
            italic: false,
            underline: false,
        }
    }

    /// Creates an owned text box from a typed, borrowed OfficeArt shape.
    pub(crate) fn from_odraw(
        properties: ShapeProperties,
        shape: &OdrawShape<'_>,
    ) -> super::super::package::Result<TextBox<'static>> {
        let wrapper = crate::odraw::textbox(shape)?
            .map(|record| crate::EscherTextboxWrapper::new(record.data().to_vec()))
            .transpose()?;
        let (text, runs, paragraph_runs, text_ruler, metachars, outline_text_refs) = match wrapper {
            Some(wrapper) => (
                wrapper.text().to_owned(),
                wrapper.runs().to_vec(),
                wrapper.paragraph_runs().to_vec(),
                wrapper.text_ruler().cloned(),
                wrapper.metachars().to_vec(),
                wrapper.outline_text_refs().to_vec(),
            ),
            None => (
                String::new(),
                Vec::new(),
                Vec::new(),
                None,
                Vec::new(),
                Vec::new(),
            ),
        };
        let (font_size, font_color, bold, italic, underline) = Self::formatting_from_runs(&runs);
        let tags = shape.programmable_tags()?;
        let text_style_extension9 = tags.as_ref().and_then(|tags| tags.powerpoint9()).cloned();
        let text_style_extension10 = tags.as_ref().and_then(|tags| tags.powerpoint10()).cloned();
        let text_style_extension11 = tags.as_ref().and_then(|tags| tags.powerpoint11()).cloned();

        let mut container = ShapeContainer::new(properties, Vec::new());
        Self::extract_odraw_text_properties(shape, &mut container)?;

        Ok(TextBox {
            container,
            text,
            runs,
            paragraph_runs,
            text_ruler,
            metachars,
            outline_text_refs,
            text_style_extension9,
            text_style_extension10,
            text_style_extension11,
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
        let paragraph_runs = if text.is_empty() {
            Vec::new()
        } else {
            vec![ParagraphRun::with_formatting(
                text.clone(),
                0,
                ParagraphRunFormatting::default(),
            )]
        };

        Self {
            container,
            text,
            runs,
            paragraph_runs,
            text_ruler: None,
            metachars: Vec::new(),
            outline_text_refs: Vec::new(),
            text_style_extension9: None,
            text_style_extension10: None,
            text_style_extension11: None,
            font_size: None,
            font_color: None,
            bold: false,
            italic: false,
            underline: false,
        }
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

    /// Extract additional text properties from typed OfficeArt options.
    ///
    /// Based on Apache POI's text property extraction logic, this function
    /// extracts text-related properties from the Escher Opt record within
    /// the shape's Escher record hierarchy.
    ///
    /// # Note
    ///
    /// This function extracts Escher-level text properties. Text formatting
    /// like bold, italic, font size typically comes from StyleTextPropAtom
    /// records, not from Escher properties. This function focuses on:
    /// - Text margins (insets)
    /// - Text flow settings
    /// - Text anchor/alignment settings
    fn extract_odraw_text_properties<'container>(
        shape: &OdrawShape<'_>,
        container: &mut ShapeContainer<'container>,
    ) -> super::super::package::Result<()> {
        use litchi_odraw::{
            RecordKind,
            prop::{Id, Props},
        };

        let secondary = shape
            .meta()
            .find(RecordKind::SecondaryOpt)?
            .map(|record| Props::parse(&record))
            .transpose()?;
        let tertiary = shape
            .meta()
            .find(RecordKind::TertiaryOpt)?
            .map(|record| Props::parse(&record))
            .transpose()?;
        let property = |id| {
            shape
                .props()
                .get_int(id)
                .or_else(|| secondary.as_ref().and_then(|props| props.get_int(id)))
                .or_else(|| tertiary.as_ref().and_then(|props| props.get_int(id)))
        };
        let boolean = |id| {
            shape
                .props()
                .get_bool(id)
                .or_else(|| secondary.as_ref().and_then(|props| props.get_bool(id)))
                .or_else(|| tertiary.as_ref().and_then(|props| props.get_bool(id)))
        };

        container.text_id = property(Id::TextId);
        let margin_value = |id, name| -> super::super::package::Result<Option<i32>> {
            const MAX_TEXT_MARGIN: i32 = 0x0132_F540;
            property(id)
                .map(|value| {
                    if !(0..=MAX_TEXT_MARGIN).contains(&value) {
                        Err(super::super::package::Error::Corrupted(format!(
                            "OfficeArt {name} margin exceeds the MS-ODRAW limit"
                        )))
                    } else {
                        Ok(value)
                    }
                })
                .transpose()
        };
        container.text_left = margin_value(Id::TextLeft, "left")?;
        container.text_top = margin_value(Id::TextTop, "top")?;
        container.text_right = margin_value(Id::TextRight, "right")?;
        container.text_bottom = margin_value(Id::TextBottom, "bottom")?;
        container.id_of_next_shape = property(Id::IdOfNextShape).map(|value| value as u32);

        let enum_value = |id, name| -> super::super::package::Result<Option<u16>> {
            property(id)
                .map(|value| {
                    u16::try_from(value).map_err(|_| {
                        super::super::package::Error::Corrupted(format!(
                            "OfficeArt {name} value does not fit in 16 bits"
                        ))
                    })
                })
                .transpose()
        };
        container.wrap_text = enum_value(Id::WrapText, "WrapText")?;
        container.anchor_text = enum_value(Id::AnchorText, "anchorText")?;
        container.text_flow = enum_value(Id::TextFlow, "txflTextFlow")?;
        container.font_rotation = enum_value(Id::FontRotation, "cdirFont")?;
        container.text_direction = enum_value(Id::TextDirection, "txdir")?;

        container.select_text = boolean(Id::SelectText);
        container.auto_text_margin = boolean(Id::AutoTextMargin);
        container.size_shape_to_fit_text = boolean(Id::FitShapeToText);

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
    /// let props = litchi_odraw::prop::Props::parse(&opt_record)?;
    /// let (size, color, bold, italic, underline) = TextBox::format_from_props(&props);
    /// ```
    pub fn format_from_props(props: &litchi_odraw::prop::Props<'_>) -> TextFormattingResult {
        use litchi_odraw::prop::Id;

        // Extract font size from text properties
        // In Escher, font size is typically in the GeoText properties
        let font_size = props
            .get_int(Id::GeoTextDefaultPointSize)
            .map(|size| size as u16);

        // Extract font color - not typically in Escher properties for text
        // Text color is usually in StyleTextPropAtom records
        let font_color = None;

        // Extract text styling flags from GeoText properties
        // These are boolean properties in Apache POI
        let bold = props.is_true(Id::GeoTextBoldFont);
        let italic = props.is_true(Id::GeoTextItalicFont);
        let underline = props.is_true(Id::GeoTextUnderlineFont);

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
    /// Tuple of (left, top, right, bottom) margins in EMUs
    ///
    /// # Performance
    ///
    /// - Single call to get_text_margins (already optimized)
    /// - No allocations
    pub fn margins_from_props(
        props: &litchi_odraw::prop::Props<'_>,
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

    /// Get the paragraph-formatting runs in document order.
    pub fn paragraph_runs(&self) -> &[ParagraphRun] {
        &self.paragraph_runs
    }

    /// Get textbox-specific tab, margin, and indent overrides.
    pub fn text_ruler(&self) -> Option<&TextRuler> {
        self.text_ruler.as_ref()
    }

    /// Header/footer metacharacter placeholders in this text box, in record
    /// order (MS-PPT 2.9.47-2.9.52). Placeholders are never substituted,
    /// formatted, or laid out.
    pub fn metachars(&self) -> &[crate::text_metachar::TextMetachar] {
        &self.metachars
    }

    /// Outline text references (`OutlineTextRefAtom`, MS-PPT 2.9.78) tying
    /// this text box to outline text bodies.
    pub fn outline_text_refs(&self) -> &[crate::text_si_exception::OutlineTextRef] {
        &self.outline_text_refs
    }

    /// Get PowerPoint 9 picture-bullet and automatic-numbering extensions.
    pub fn text_style_extension9(&self) -> Option<&TextStyleExtension9> {
        self.text_style_extension9.as_ref()
    }

    /// Get PowerPoint 10 alternate-script font extensions.
    pub fn text_style_extension10(&self) -> Option<&TextStyleExtension10> {
        self.text_style_extension10.as_ref()
    }

    /// Get PowerPoint 11 smart-tag extensions.
    pub fn text_style_extension11(&self) -> Option<&TextStyleExtension11> {
        self.text_style_extension11.as_ref()
    }

    /// Set the text content of the text box.
    pub fn set_text(&mut self, text: String) -> Result<(), super::shape::MutationError> {
        self.container
            .ensure_mutable(super::shape::Mutation::Text)?;
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
        self.paragraph_runs = if text.is_empty() {
            Vec::new()
        } else {
            vec![ParagraphRun::with_formatting(
                text.clone(),
                0,
                ParagraphRunFormatting::default(),
            )]
        };
        self.container.set_decoded_text(text);
        Ok(())
    }

    pub(crate) fn mark_source_bound(&mut self) {
        self.container.mark_source_bound();
    }

    /// Get the font size in points.
    pub fn font_size(&self) -> Option<u16> {
        self.font_size
    }

    /// Set the font size in points.
    pub fn set_font_size(&mut self, size: u16) -> Result<(), super::shape::MutationError> {
        self.container
            .ensure_mutable(super::shape::Mutation::Formatting)?;
        self.font_size = Some(size);
        for run in &mut self.runs {
            run.formatting.font_size = Some(size);
        }
        Ok(())
    }

    /// Get the font color (RGB).
    pub fn font_color(&self) -> Option<u32> {
        self.font_color
    }

    /// Set the font color (RGB).
    pub fn set_font_color(&mut self, color: u32) -> Result<(), super::shape::MutationError> {
        self.container
            .ensure_mutable(super::shape::Mutation::Formatting)?;
        let color = color & 0x00FF_FFFF;
        self.font_color = Some(color);
        let red = (color >> 16) & 0xFF;
        let green = (color >> 8) & 0xFF;
        let blue = color & 0xFF;
        let raw = red | (green << 8) | (blue << 16) | 0xFE00_0000;
        for run in &mut self.runs {
            run.formatting.font_color = Some(color);
            run.formatting.font_color_raw = Some(raw);
            run.formatting.font_scheme_color = None;
        }
        Ok(())
    }

    /// Get the raw `ColorIndexStruct` value of the first text run.
    pub fn font_color_raw(&self) -> Option<u32> {
        self.runs
            .first()
            .and_then(|run| run.formatting.font_color_raw)
    }

    /// Get the PowerPoint color-scheme index of the first text run.
    pub fn font_scheme_color(&self) -> Option<u8> {
        self.runs
            .first()
            .and_then(|run| run.formatting.font_scheme_color)
    }

    /// Get the zero-based font reference of the first text run.
    pub fn font_index(&self) -> Option<u16> {
        self.runs.first().and_then(|run| run.formatting.font_index)
    }

    /// Check if the text is bold.
    pub fn bold(&self) -> bool {
        self.bold
    }

    /// Set bold formatting.
    pub fn set_bold(&mut self, bold: bool) -> Result<(), super::shape::MutationError> {
        self.container
            .ensure_mutable(super::shape::Mutation::Formatting)?;
        self.bold = bold;
        for run in &mut self.runs {
            run.formatting.bold = bold;
            run.formatting.bold_explicit = Some(bold);
        }
        Ok(())
    }

    /// Check if the text is italic.
    pub fn italic(&self) -> bool {
        self.italic
    }

    /// Set italic formatting.
    pub fn set_italic(&mut self, italic: bool) -> Result<(), super::shape::MutationError> {
        self.container
            .ensure_mutable(super::shape::Mutation::Formatting)?;
        self.italic = italic;
        for run in &mut self.runs {
            run.formatting.italic = italic;
            run.formatting.italic_explicit = Some(italic);
        }
        Ok(())
    }

    /// Check if the text is underlined.
    pub fn underline(&self) -> bool {
        self.underline
    }

    /// Set underline formatting.
    pub fn set_underline(&mut self, underline: bool) -> Result<(), super::shape::MutationError> {
        self.container
            .ensure_mutable(super::shape::Mutation::Formatting)?;
        self.underline = underline;
        for run in &mut self.runs {
            run.formatting.underline = underline;
            run.formatting.underline_explicit = Some(underline);
        }
        Ok(())
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

    /// Get text insets in EMUs with MS-ODRAW defaults applied.
    pub fn effective_text_margins(&self) -> (i32, i32, i32, i32) {
        self.container.effective_text_margins()
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

    /// Get the text identifier stored in the OfficeArt options.
    pub fn text_id(&self) -> Option<i32> {
        self.container.text_id
    }

    /// Get the raw `MSOCDIR` font-rotation value.
    pub fn font_rotation(&self) -> Option<u16> {
        self.container.font_rotation
    }

    /// Get the next shape ID in a linked-textbox sequence.
    pub fn next_shape_id(&self) -> Option<u32> {
        self.container.id_of_next_shape
    }

    /// Get the raw `MSOTXDIR` text-direction value.
    pub fn text_direction(&self) -> Option<u16> {
        self.container.text_direction
    }

    /// Whether one click on the text area enters text editing mode.
    pub fn single_click_selects_text(&self) -> Option<bool> {
        self.container.select_text
    }

    /// Whether the shape uses automatic default text margins.
    pub fn automatic_text_margins(&self) -> Option<bool> {
        self.container.auto_text_margin
    }

    /// Whether the shape dimensions should be adjusted to fit the text.
    pub fn size_shape_to_fit_text(&self) -> Option<bool> {
        self.container.size_shape_to_fit_text
    }

    fn current_run_formatting(&self) -> TextRunFormatting {
        TextRunFormatting {
            font_size: self.font_size,
            font_color: self.font_color,
            font_color_raw: self.font_color.map(|color| {
                let red = (color >> 16) & 0xFF;
                let green = (color >> 8) & 0xFF;
                let blue = color & 0xFF;
                red | (green << 8) | (blue << 16) | 0xFE00_0000
            }),
            font_scheme_color: None,
            bold: self.bold,
            bold_explicit: Some(self.bold),
            italic: self.italic,
            italic_explicit: Some(self.italic),
            underline: self.underline,
            underline_explicit: Some(self.underline),
            ..TextRunFormatting::default()
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

    fn properties_mut(&mut self) -> Result<&mut ShapeProperties, super::shape::MutationError> {
        self.container.properties_mut_checked()
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

    fn formatted_textbox_record(wrap_mode: u32, left_margin: u32) -> Vec<u8> {
        let mut ppt_records = Vec::new();
        push_record(&mut ppt_records, 0, 0, 3999, &4u32.to_le_bytes());
        push_record(&mut ppt_records, 0, 0, 4008, b"abcd");
        let mut ruler = Vec::new();
        ruler.extend_from_slice(&0x010Fu32.to_le_bytes());
        ruler.extend_from_slice(&1i16.to_le_bytes());
        ruler.extend_from_slice(&144i16.to_le_bytes());
        ruler.extend_from_slice(&1u16.to_le_bytes());
        ruler.extend_from_slice(&720i16.to_le_bytes());
        ruler.extend_from_slice(&2u16.to_le_bytes());
        ruler.extend_from_slice(&100i16.to_le_bytes());
        ruler.extend_from_slice(&(-50i16).to_le_bytes());
        push_record(&mut ppt_records, 0, 0, 4006, &ruler);

        let mut style = Vec::new();
        style.extend_from_slice(&5u32.to_le_bytes());
        style.extend_from_slice(&0i16.to_le_bytes());
        style.extend_from_slice(&0u32.to_le_bytes());
        style.extend_from_slice(&2u32.to_le_bytes());
        style.extend_from_slice(&0x0006_0001u32.to_le_bytes());
        style.extend_from_slice(&1i16.to_le_bytes());
        style.extend_from_slice(&18i16.to_le_bytes());
        style.extend_from_slice(&(0xFE33_2211u32 as i32).to_le_bytes());
        style.extend_from_slice(&3u32.to_le_bytes());
        style.extend_from_slice(&0x0006_0002u32.to_le_bytes());
        style.extend_from_slice(&2i16.to_le_bytes());
        style.extend_from_slice(&24i16.to_le_bytes());
        style.extend_from_slice(&0x0400_0000i32.to_le_bytes());
        push_record(&mut ppt_records, 0, 0, 4001, &style);

        let properties = [
            (0x0080u16, 0xCAFE_BABEu32),
            (0x0081u16, left_margin),
            (0x0082u16, 45_720u32),
            (0x0083u16, 91_440u32),
            (0x0084u16, 45_720u32),
            (0x0085u16, wrap_mode),
            (0x0087u16, 2u32),
            (0x0088u16, 1u32),
            (0x0089u16, 3u32),
            (0x008Au16, 2002u32),
        ];
        let mut opt = Vec::new();
        for (id, value) in properties {
            opt.extend_from_slice(&id.to_le_bytes());
            opt.extend_from_slice(&value.to_le_bytes());
        }
        let secondary_properties = [(0x008Bu16, 1u32), (0x00BFu16, 0x001A_0012u32)];
        let mut secondary_opt = Vec::new();
        for (id, value) in secondary_properties {
            secondary_opt.extend_from_slice(&id.to_le_bytes());
            secondary_opt.extend_from_slice(&value.to_le_bytes());
        }

        let mut shape = Vec::new();
        let mut sp = Vec::new();
        sp.extend_from_slice(&1001u32.to_le_bytes());
        sp.extend_from_slice(&0x0800u32.to_le_bytes());
        push_record(&mut shape, 2, 24, 0xF00A, &sp);
        push_record(&mut shape, 3, properties.len() as u16, 0xF00B, &opt);
        push_record(
            &mut shape,
            3,
            secondary_properties.len() as u16,
            0xF121,
            &secondary_opt,
        );
        push_record(&mut shape, 0xF, 0, 0xF00D, &ppt_records);

        let mut result = Vec::new();
        push_record(&mut result, 0xF, 0, 0xF004, &shape);
        result
    }

    fn parse_formatted_textbox(data: &[u8]) -> crate::package::Result<TextBox<'static>> {
        let shapes = litchi_odraw::shape::parse(data)?;
        let shape = shapes.first().ok_or_else(|| {
            crate::package::Error::Corrupted("test drawing contains no shape".to_owned())
        })?;
        let properties = ShapeProperties {
            id: shape.id(),
            shape_type: ShapeType::TextBox,
            ..Default::default()
        };
        TextBox::from_odraw(properties, shape)
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
        assert!(textbox.set_text("Hello World".to_string()).is_ok());

        assert_eq!(textbox.text(), "Hello World");
        assert!(textbox.has_text());
        assert_eq!(textbox.paragraph_runs().len(), 1);
        assert_eq!(textbox.paragraph_runs()[0].text, "Hello World");
    }

    #[test]
    #[allow(clippy::field_reassign_with_default)]
    fn test_textbox_formatting() {
        let mut props = ShapeProperties::default();
        props.shape_type = ShapeType::TextBox;

        let mut textbox = TextBox::new(props, vec![]);
        assert!(textbox.set_font_size(12).is_ok());
        assert!(textbox.set_font_color(0xFF0000).is_ok());
        assert!(textbox.set_bold(true).is_ok());
        assert!(textbox.set_italic(true).is_ok());

        let formatting = textbox.formatting();
        assert_eq!(formatting.font_size, Some(12));
        assert_eq!(formatting.font_color, Some(0xFF0000));
        assert!(formatting.bold);
        assert!(formatting.italic);
        assert!(!formatting.underline);
    }

    #[test]
    fn parses_embedded_style_runs_and_nested_text_options() {
        let data = formatted_textbox_record(0, 91_440);
        let textbox = parse_formatted_textbox(&data).unwrap();

        assert_eq!(textbox.text(), "abcd");
        assert_eq!(textbox.runs().len(), 2);
        assert_eq!(textbox.paragraph_runs().len(), 1);
        assert_eq!(textbox.paragraph_runs()[0].text, "abcd");
        assert_eq!(textbox.paragraph_runs()[0].formatting.property_mask, 0);
        let ruler = textbox.text_ruler().unwrap();
        assert_eq!(ruler.level_count, Some(1));
        assert_eq!(ruler.default_tab_size, Some(144));
        assert_eq!(ruler.tab_stops[0].position, 720);
        assert_eq!(ruler.tab_stops[0].alignment, 2);
        assert_eq!(ruler.levels[0].left_margin, Some(100));
        assert_eq!(ruler.levels[0].indent, Some(-50));
        assert_eq!(textbox.runs()[0].text, "ab");
        assert_eq!(textbox.runs()[0].formatting.font_size, Some(18));
        assert_eq!(textbox.runs()[0].formatting.font_color, Some(0x0011_2233));
        assert_eq!(
            textbox.runs()[0].formatting.font_color_raw,
            Some(0xFE33_2211)
        );
        assert_eq!(textbox.runs()[0].formatting.font_scheme_color, None);
        assert!(textbox.runs()[0].formatting.bold);
        assert!(!textbox.runs()[0].formatting.italic);
        assert_eq!(textbox.runs()[1].text, "cd");
        assert_eq!(textbox.runs()[1].formatting.font_size, Some(24));
        assert_eq!(textbox.runs()[1].formatting.font_color, None);
        assert_eq!(textbox.runs()[1].formatting.font_scheme_color, Some(4));
        assert!(!textbox.runs()[1].formatting.bold);
        assert!(textbox.runs()[1].formatting.italic);
        assert_eq!(textbox.formatting().font_size, Some(18));
        assert_eq!(textbox.font_color_raw(), Some(0xFE33_2211));
        assert_eq!(textbox.font_scheme_color(), None);
        assert!(textbox.bold());
        assert_eq!(
            textbox.text_margins(),
            Some((91_440, 45_720, 91_440, 45_720))
        );
        assert_eq!(
            textbox.effective_text_margins(),
            (91_440, 45_720, 91_440, 45_720)
        );
        assert_eq!(textbox.wrap_mode(), Some(0));
        assert_eq!(textbox.word_wrap_enabled(), Some(true));
        assert_eq!(textbox.text_anchor(), Some(2));
        assert_eq!(textbox.text_flow(), Some(1));
        assert_eq!(textbox.text_id(), Some(0xCAFE_BABEu32 as i32));
        assert_eq!(textbox.font_rotation(), Some(3));
        assert_eq!(textbox.next_shape_id(), Some(2002));
        assert_eq!(textbox.text_direction(), Some(1));
        assert_eq!(textbox.single_click_selects_text(), Some(true));
        assert_eq!(textbox.automatic_text_margins(), Some(false));
        assert_eq!(textbox.size_shape_to_fit_text(), Some(true));
    }

    #[test]
    fn setters_keep_character_runs_consistent() {
        let mut textbox = TextBox::new(ShapeProperties::default(), Vec::new());
        assert!(textbox.set_bold(true).is_ok());
        assert!(textbox.set_font_size(20).is_ok());
        assert!(textbox.set_font_color(0x12_34_56).is_ok());
        assert!(textbox.set_text("hello".to_string()).is_ok());
        assert!(textbox.set_italic(true).is_ok());

        assert_eq!(textbox.runs().len(), 1);
        assert_eq!(textbox.runs()[0].text, "hello");
        assert_eq!(textbox.runs()[0].formatting.font_size, Some(20));
        assert_eq!(textbox.runs()[0].formatting.font_color, Some(0x12_34_56));
        assert_eq!(textbox.font_color_raw(), Some(0xFE56_3412));
        assert!(textbox.runs()[0].formatting.bold);
        assert!(textbox.runs()[0].formatting.italic);
    }

    #[test]
    fn rejects_out_of_range_text_enum_values() {
        let data = formatted_textbox_record(u32::from(u16::MAX) + 1, 91_440);
        let error = parse_formatted_textbox(&data).unwrap_err();
        assert!(error.to_string().contains("WrapText value"));
    }

    #[test]
    fn rejects_out_of_range_text_margins() {
        let data = formatted_textbox_record(0, 0x0132_F541);
        let error = parse_formatted_textbox(&data).unwrap_err();
        assert!(error.to_string().contains("left margin"));
    }

    #[test]
    fn source_bound_textbox_setters_are_atomic() {
        let mut textbox = TextBox::new(ShapeProperties::default(), Vec::new());
        assert!(textbox.set_text("before".to_owned()).is_ok());
        let before_text = textbox.text().to_owned();
        let before_formatting = textbox.formatting();
        let before_run = textbox.runs().first().map(|run| {
            (
                run.text.clone(),
                run.formatting.font_size,
                run.formatting.font_color,
                run.formatting.font_color_raw,
                run.formatting.bold,
                run.formatting.italic,
                run.formatting.underline,
            )
        });
        textbox.mark_source_bound();

        assert_eq!(
            textbox.set_text("changed".to_owned()),
            Err(super::super::shape::MutationError::SourceBound {
                mutation: super::super::shape::Mutation::Text,
            })
        );
        for result in [
            textbox.set_font_size(24),
            textbox.set_font_color(0x12_34_56),
            textbox.set_bold(true),
            textbox.set_italic(true),
            textbox.set_underline(true),
        ] {
            assert_eq!(
                result,
                Err(super::super::shape::MutationError::SourceBound {
                    mutation: super::super::shape::Mutation::Formatting,
                })
            );
        }
        assert_eq!(textbox.text(), before_text);
        assert_eq!(textbox.formatting(), before_formatting);
        assert_eq!(
            textbox.runs().first().map(|run| {
                (
                    run.text.clone(),
                    run.formatting.font_size,
                    run.formatting.font_color,
                    run.formatting.font_color_raw,
                    run.formatting.bold,
                    run.formatting.italic,
                    run.formatting.underline,
                )
            }),
            before_run
        );
    }
}
