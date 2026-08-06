use super::super::prelude::*;

impl Document {
    /// Get typed, inert `MACROBUTTON` fields in story and source order.
    ///
    /// Returned values expose only stored macro or command names, button text,
    /// cached results, and field state. This method never resolves, loads,
    /// invokes, or otherwise executes a macro or command.
    pub fn macro_button_fields(&self) -> Result<Vec<MacroButtonField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::macro_button).collect())
    }

    /// Get the number of typed, inert `MACROBUTTON` fields.
    pub fn macro_button_field_count(&self) -> Result<usize> {
        Ok(self.macro_button_fields()?.len())
    }

    /// Get typed, inert `ADDIN`, `CONTROL`, and `HTMLCONTROL` fields in story and source order.
    ///
    /// Returned values expose only stored kind, instruction, cached-result, and
    /// field-state metadata. This method never loads an add-in, instantiates a
    /// control, invokes code, executes script, renders content, or accesses an
    /// external resource.
    pub fn active_content_fields(&self) -> Result<Vec<ActiveContentField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::active_content_field)
            .collect())
    }

    /// Get the number of typed, inert add-in and control fields.
    pub fn active_content_field_count(&self) -> Result<usize> {
        Ok(self.active_content_fields()?.len())
    }

    /// Get typed, inert `PRINT` fields in story and source order.
    ///
    /// Returned values expose only stored printer-instruction text, cached
    /// results, and field state. This method never interprets control codes,
    /// opens a printer, sends output, changes print settings, or refreshes a
    /// field.
    pub fn print_fields(&self) -> Result<Vec<PrintField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::print_field).collect())
    }

    /// Get the number of typed, inert `PRINT` fields.
    pub fn print_field_count(&self) -> Result<usize> {
        Ok(self.print_fields()?.len())
    }

    /// Get typed, inert `EMBED` fields in story and source order.
    ///
    /// Returned values expose only stored opaque object instructions, cached
    /// results, and field state. This method never loads, inspects,
    /// deserializes, activates, renders, or executes an embedded object,
    /// accesses an external resource, or refreshes a field.
    pub fn embed_fields(&self) -> Result<Vec<EmbedField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::embed_field).collect())
    }

    /// Get the number of typed, inert `EMBED` fields.
    pub fn embed_field_count(&self) -> Result<usize> {
        Ok(self.embed_fields()?.len())
    }

    /// Get typed, inert `BARCODE` fields in story and source order.
    ///
    /// Returned values expose only stored opaque barcode instructions, cached
    /// results, and field state. This method never parses or validates barcode
    /// data or symbology, generates or renders a barcode, accesses an external
    /// resource, or refreshes a field.
    pub fn barcode_fields(&self) -> Result<Vec<BarcodeField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::barcode_field).collect())
    }

    /// Get the number of typed, inert `BARCODE` fields.
    pub fn barcode_field_count(&self) -> Result<usize> {
        Ok(self.barcode_fields()?.len())
    }

    /// Get typed, inert `BIDIOUTLINE` fields in story and source order.
    ///
    /// Returned values expose only stored opaque instructions, cached results,
    /// and field state. This method never reads right-to-left language,
    /// paragraph outline, or layout state; chooses a numbering system;
    /// calculates a result; or refreshes a field.
    pub fn bidi_outline_fields(&self) -> Result<Vec<BidiOutlineField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::bidi_outline_field)
            .collect())
    }

    /// Get the number of typed, inert `BIDIOUTLINE` fields.
    pub fn bidi_outline_field_count(&self) -> Result<usize> {
        Ok(self.bidi_outline_fields()?.len())
    }

    /// Get typed, inert `SHAPE` drawing-canvas anchor fields in story and source order.
    ///
    /// Returned values expose only stored opaque instructions, cached results,
    /// and field state. This method never locates, links, loads, positions,
    /// lays out, or renders a drawing or canvas, or refreshes a field.
    pub fn shape_fields(&self) -> Result<Vec<ShapeField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::shape_field).collect())
    }

    /// Get the number of typed, inert `SHAPE` drawing-canvas anchor fields.
    pub fn shape_field_count(&self) -> Result<usize> {
        Ok(self.shape_fields()?.len())
    }

    /// Get typed, inert legacy form-code fields in story and source order.
    ///
    /// Returned values expose only stored text/checkbox/drop-down kind, opaque
    /// instructions, cached results, field state, and — when the field's
    /// `NilPICFAndBinData` could be located in the Data stream and parsed —
    /// the stored `FFData` form state. This method never fills a form, changes
    /// a selection or checkbox state, invokes entry or exit macros, or
    /// refreshes a field.
    pub fn legacy_form_fields(&self) -> Result<Vec<LegacyFormField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::legacy_form_field)
            .map(|field| self.attach_form_field_data(field))
            .collect())
    }

    /// Attach the stored `FFData` form state to a legacy form-code field.
    ///
    /// The picture character (U+0001) inside the field instruction carries
    /// `sprmCData` and `sprmCPicLocation`, pointing at the field's
    /// `NilPICFAndBinData` in the Data stream (MS-DOC 2.9.158). Invalid or
    /// absent binary data MUST be ignored, so failures leave the field's
    /// `form_data` as `None` rather than failing the whole listing.
    fn attach_form_field_data(&self, mut field: LegacyFormField) -> LegacyFormField {
        field.set_form_data(self.parse_form_field_data(field.field()));
        field
    }

    /// Locate and parse the `FFData` of one legacy form-code field.
    fn parse_form_field_data(&self, field: &Field) -> Option<FormFieldData> {
        let data_stream = self.data_stream.as_deref()?;
        let chp_table = self.chp_bin_table.as_ref()?;
        let (story_start, _story_end) = self.field_story_range_if_present(field.story)?;
        let (code_start, code_end) = field.code_range();
        let instruction = self
            .field_story_text(field.story, code_start, code_end)
            .ok()?;
        let base_cp = story_start.checked_add(code_start)?;
        // CPs count UTF-16 code units, so scan the instruction by code unit.
        for (unit_index, unit) in instruction.encode_utf16().enumerate() {
            if unit != 0x0001 {
                continue;
            }
            let picture_cp = base_cp.checked_add(u32::try_from(unit_index).ok()?)?;
            let picture_end = picture_cp.checked_add(1)?;
            for run in chp_table.runs_in_range(picture_cp, picture_end) {
                let properties = &run.properties;
                if !properties.is_data {
                    continue;
                }
                let Some(offset) = properties.pic_offset else {
                    continue;
                };
                if let Ok(data) = FormFieldData::parse_at(data_stream, offset) {
                    return Some(data);
                }
            }
        }
        None
    }

    /// Get the number of typed, inert legacy form-code fields.
    pub fn legacy_form_field_count(&self) -> Result<usize> {
        Ok(self.legacy_form_fields()?.len())
    }
}
