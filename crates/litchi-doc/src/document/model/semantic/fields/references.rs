use super::super::prelude::*;

impl Document {
    /// Get typed, inert bookmark-reference fields in story and source order.
    ///
    /// Returned values expose only stored categories, bookmark names, options,
    /// switches, cached results, and field state. This method never looks up a
    /// bookmark, reads a referenced range, resolves a page or note number,
    /// creates a link, calculates a relative position, or refreshes a field.
    pub fn reference_fields(&self) -> Result<Vec<ReferenceField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::reference_field)
            .collect())
    }

    /// Get the number of typed, inert bookmark-reference fields.
    pub fn reference_field_count(&self) -> Result<usize> {
        Ok(self.reference_fields()?.len())
    }

    /// Get typed, inert `SET` fields in story and source order.
    ///
    /// Returned values expose only stored target names, opaque expressions,
    /// cached results, and field state. This method never evaluates an
    /// expression, looks up or changes a bookmark, changes document state, or
    /// refreshes a field.
    pub fn set_fields(&self) -> Result<Vec<SetField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::set_field).collect())
    }

    /// Get the number of typed, inert `SET` fields.
    pub fn set_field_count(&self) -> Result<usize> {
        Ok(self.set_fields()?.len())
    }

    /// Get typed, inert `=` formula fields in story and source order.
    ///
    /// Returned values expose only stored optional formulas, cached results,
    /// and field state. This method never parses or evaluates a formula, reads
    /// table cells or bookmarks, resolves field values, or refreshes a field.
    pub fn formula_fields(&self) -> Result<Vec<FormulaField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::formula_field).collect())
    }

    /// Get the number of typed, inert `=` formula fields.
    pub fn formula_field_count(&self) -> Result<usize> {
        Ok(self.formula_fields()?.len())
    }

    /// Get typed, inert `EQ` equation fields in story and source order.
    ///
    /// Returned values expose stored opaque expressions, cached results, and
    /// field state only. This method never parses, calculates, formats,
    /// renders, or refreshes an equation.
    pub fn equations(&self) -> Result<Vec<EquationField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::equation_field)
            .collect())
    }

    /// Get the number of typed, inert `EQ` fields.
    pub fn equation_count(&self) -> Result<usize> {
        Ok(self.equations()?.len())
    }

    /// Get typed, inert `HYPERLINK` fields in story and source order.
    ///
    /// Returned values expose stored targets, options, cached results, and
    /// field state only. This method never opens, resolves, follows, activates,
    /// or refreshes a link.
    pub fn hyperlink_fields(&self) -> Result<Vec<HyperlinkField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::hyperlink_field)
            .collect())
    }

    /// Get the number of typed, inert `HYPERLINK` fields.
    pub fn hyperlink_field_count(&self) -> Result<usize> {
        Ok(self.hyperlink_fields()?.len())
    }

    /// Get typed, inert `QUOTE` fields in story and source order.
    ///
    /// Returned values expose only stored text arguments, switches, cached
    /// results, and field state. This method never interprets character codes,
    /// expands nested fields, inserts text, or refreshes a field.
    pub fn quote_fields(&self) -> Result<Vec<QuoteField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::quote_field).collect())
    }

    /// Get the number of typed, inert `QUOTE` fields.
    pub fn quote_field_count(&self) -> Result<usize> {
        Ok(self.quote_fields()?.len())
    }

    /// Get typed, inert `SYMBOL` fields in story and source order.
    ///
    /// Returned values expose only stored character arguments, switches, cached
    /// results, and field state. This method never maps a character code, looks
    /// up a font, inserts a glyph, changes formatting or layout, or refreshes a
    /// field.
    pub fn symbol_fields(&self) -> Result<Vec<SymbolField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::symbol_field).collect())
    }

    /// Get the number of typed, inert `SYMBOL` fields.
    pub fn symbol_field_count(&self) -> Result<usize> {
        Ok(self.symbol_fields()?.len())
    }

    /// Get typed, inert legacy automatic-numbering fields in story and source order.
    ///
    /// Returned values expose only stored kinds, switches, cached results, and
    /// field state. This method never calculates paragraph numbers, reads
    /// heading or style state, changes paragraphs or layout, or refreshes a
    /// field.
    pub fn auto_number_fields(&self) -> Result<Vec<AutoNumberField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::auto_number_field)
            .collect())
    }

    /// Get the number of typed, inert legacy automatic-numbering fields.
    pub fn auto_number_field_count(&self) -> Result<usize> {
        Ok(self.auto_number_fields()?.len())
    }

    /// Get typed, inert `LISTNUM` fields in story and source order.
    ///
    /// Returned values expose only stored optional list names, switches, cached
    /// results, and field state. This method never looks up a list, determines a
    /// level or start value, calculates a number, changes layout, or refreshes
    /// a field.
    pub fn list_number_fields(&self) -> Result<Vec<ListNumberField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::list_number_field)
            .collect())
    }

    /// Get the number of typed, inert `LISTNUM` fields.
    pub fn list_number_field_count(&self) -> Result<usize> {
        Ok(self.list_number_fields()?.len())
    }

    /// Get typed, inert `SEQ` fields in story and source order.
    ///
    /// Returned values expose only stored identifiers, optional bookmark names,
    /// opaque tails, cached results, and field state. This method never looks
    /// up a bookmark, increments or resets a sequence, calculates a number, or
    /// refreshes a field.
    pub fn sequence_fields(&self) -> Result<Vec<SequenceField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::sequence_field)
            .collect())
    }

    /// Get the number of typed, inert `SEQ` fields.
    pub fn sequence_field_count(&self) -> Result<usize> {
        Ok(self.sequence_fields()?.len())
    }

    /// Get typed, inert `STYLEREF` fields in story and source order.
    ///
    /// Returned values expose only stored style names, options, switches, cached
    /// results, and field state. This method never looks up styled text, searches
    /// document stories, calculates paragraph numbers or relative positions,
    /// resolves page layout, or refreshes a field.
    pub fn style_reference_fields(&self) -> Result<Vec<StyleReferenceField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::style_reference_field)
            .collect())
    }

    /// Get the number of typed, inert `STYLEREF` fields.
    pub fn style_reference_field_count(&self) -> Result<usize> {
        Ok(self.style_reference_fields()?.len())
    }

    /// Get typed, inert `GLOSSARY` and `AUTOTEXT` fields in story and source order.
    ///
    /// Returned values expose only stored category, entry name, switches,
    /// cached results, and field state. This method never looks up a building
    /// block, reads a template, inserts content, changes bookmarks, opens a
    /// resource, or refreshes a field.
    pub fn auto_text_fields(&self) -> Result<Vec<AutoTextField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::auto_text_field)
            .collect())
    }

    /// Get the number of typed, inert `GLOSSARY` and `AUTOTEXT` fields.
    pub fn auto_text_field_count(&self) -> Result<usize> {
        Ok(self.auto_text_fields()?.len())
    }

    /// Get typed, inert `AUTOTEXTLIST` fields in story and source order.
    ///
    /// Returned values expose only stored display text, style/tip options,
    /// switches, cached results, and field state. This method never shows a
    /// selection UI, looks up a building block, reads a template, inserts
    /// content, or refreshes a field.
    pub fn auto_text_list_fields(&self) -> Result<Vec<AutoTextListField>> {
        let fields = self.fields()?;
        Ok(fields
            .iter()
            .filter_map(FieldText::auto_text_list_field)
            .collect())
    }

    /// Get the number of typed, inert `AUTOTEXTLIST` fields.
    pub fn auto_text_list_field_count(&self) -> Result<usize> {
        Ok(self.auto_text_list_fields()?.len())
    }

    /// Get typed, inert `GOTOBUTTON` fields in story and source order.
    ///
    /// Returned values expose only stored destinations, button text, cached
    /// results, and field state. This method never resolves a destination,
    /// changes the insertion point, activates a jump, or refreshes a field.
    pub fn go_to_button_fields(&self) -> Result<Vec<GoToButtonField>> {
        let fields = self.fields()?;
        Ok(fields.iter().filter_map(FieldText::go_to_button).collect())
    }

    /// Get the number of typed, inert `GOTOBUTTON` fields.
    pub fn go_to_button_field_count(&self) -> Result<usize> {
        Ok(self.go_to_button_fields()?.len())
    }
}
