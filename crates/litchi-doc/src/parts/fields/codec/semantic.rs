//! Typed, inert field views derived from parsed instructions.

use super::super::model::{
    ActiveContentField, ActiveContentFieldKind, AdvanceField, AutoNumberField, AutoNumberFieldKind,
    AutoTextField, AutoTextFieldKind, AutoTextListField, BarcodeField, BidiOutlineField,
    CompareField, DdeField, DdeFieldKind, DocumentContextField, DocumentContextFieldKind,
    DocumentInformationField, DocumentInformationFieldKind, DocumentPropertyField,
    DocumentVariableField, EmbedField, EquationField, ExternalIncludeField, FieldText, FieldType,
    FormulaField, GoToButtonField, HyperlinkField, IfField, IncludeFieldKind, IndexField,
    InfoField, LegacyFormField, LegacyFormFieldKind, LinkField, ListNumberField, MacroButtonField,
    MailMergeConditionalControlField, MailMergeConditionalControlKind, MailMergeCounterField,
    MailMergeCounterKind, MailMergeDataField, MailMergeNextField, MailMergeRecipientField,
    MailMergeRecipientFieldKind, MergeField, PrintField, PromptField, PromptFieldKind, QuoteField,
    ReferenceField, ReferenceFieldKind, SequenceField, SetField, ShapeField, StyleReferenceField,
    SymbolField, TableOfAuthoritiesField, TableOfContentsField, UserIdentityField,
    UserIdentityFieldKind,
};
use super::parser::{
    is_mail_merge_next_instruction, parse_advance_field_adjustments, parse_auto_number_field_parts,
    parse_auto_text_field_parts, parse_auto_text_list_field_parts,
    parse_barcode_field_instructions, parse_bidi_outline_field_instructions,
    parse_compare_field_comparison, parse_dde_field_parts, parse_document_context_field_parts,
    parse_document_information_field_parts, parse_document_property_field_parts,
    parse_document_variable_field_parts, parse_embed_field_instructions,
    parse_equation_field_expression, parse_external_include_field_parts,
    parse_formula_field_formula, parse_go_to_button_parts, parse_hyperlink_field_parts,
    parse_if_field_expression, parse_index_field_parts, parse_info_field_parts,
    parse_legacy_form_field_instructions, parse_link_field_parts, parse_list_number_field_parts,
    parse_macro_button_parts, parse_mail_merge_conditional_control_parts,
    parse_mail_merge_counter_kind, parse_mail_merge_data_field_parts,
    parse_mail_merge_recipient_field_parts, parse_merge_field_parts,
    parse_print_field_instructions, parse_prompt_field_parts, parse_quote_field_parts,
    parse_reference_field_parts, parse_sequence_field_parts, parse_set_field_parts,
    parse_shape_field_instructions, parse_style_reference_field_parts, parse_symbol_field_parts,
    parse_table_of_authorities_field_parts, parse_table_of_contents_field_parts,
    parse_user_identity_field_parts,
};
impl FieldText {
    /// Return inert typed metadata when this is a well-formed `MACROBUTTON`
    /// field.
    ///
    /// The macro or command name and button text are parsed only from stored
    /// field text. Neither is resolved, loaded, invoked, or executed.
    /// Malformed instructions remain available through this generic type and
    /// return `None` here.
    #[must_use]
    pub fn macro_button(&self) -> Option<MacroButtonField> {
        if self.field.field_type != FieldType::MacroButton {
            return None;
        }
        let (macro_name, display_text) = parse_macro_button_parts(&self.instruction)?;
        Some(MacroButtonField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            macro_name,
            display_text,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `GOTOBUTTON`
    /// field.
    ///
    /// The destination and button text are parsed only from stored field text.
    /// Neither is resolved, navigated to, or activated. Malformed instructions
    /// remain available through this generic type and return `None` here.
    #[must_use]
    pub fn go_to_button(&self) -> Option<GoToButtonField> {
        if self.field.field_type != FieldType::GoToButton {
            return None;
        }
        let (target, button_text) = parse_go_to_button_parts(&self.instruction)?;
        Some(GoToButtonField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            target,
            button_text,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a native add-in or control field.
    ///
    /// The stored instruction and cached result are never interpreted to load
    /// an add-in, instantiate a control, invoke code, execute script, render
    /// content, or access an external resource.
    #[must_use]
    pub fn active_content_field(&self) -> Option<ActiveContentField> {
        let kind = match self.field.field_type {
            FieldType::AddIn => ActiveContentFieldKind::AddIn,
            FieldType::Control => ActiveContentFieldKind::OcxControl,
            FieldType::HtmlControl => ActiveContentFieldKind::HtmlControl,
            _ => return None,
        };
        Some(ActiveContentField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a native `PRINT` field.
    ///
    /// Stored printer-instruction text, cached results, and field-marker state
    /// are opaque metadata only. This method never interprets printer-control
    /// codes, opens a printer, sends output, changes print settings, or
    /// refreshes a field. Malformed instructions remain available through this
    /// generic type and return `None` here.
    #[must_use]
    pub fn print_field(&self) -> Option<PrintField> {
        if self.field.field_type != FieldType::Print {
            return None;
        }
        let printer_instructions = parse_print_field_instructions(&self.instruction)?;
        Some(PrintField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            printer_instructions,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a native `EMBED` field.
    ///
    /// Stored opaque object instructions, cached results, and field-marker
    /// state remain metadata only. This method never loads, inspects,
    /// deserializes, activates, renders, or executes an embedded object,
    /// accesses an external resource, or refreshes a field. Malformed
    /// instructions remain available through this generic type and return
    /// `None` here.
    #[must_use]
    pub fn embed_field(&self) -> Option<EmbedField> {
        if self.field.field_type != FieldType::EmbeddedObject {
            return None;
        }
        let object_instructions = parse_embed_field_instructions(&self.instruction)?;
        Some(EmbedField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            object_instructions,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a native `BARCODE` field.
    ///
    /// Stored opaque barcode instructions, cached results, and field-marker
    /// state remain metadata only. This method never parses or validates
    /// barcode data or symbology, generates or renders a barcode, accesses an
    /// external resource, or refreshes a field. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn barcode_field(&self) -> Option<BarcodeField> {
        if self.field.field_type != FieldType::BarCode {
            return None;
        }
        let barcode_instructions = parse_barcode_field_instructions(&self.instruction)?;
        Some(BarcodeField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            barcode_instructions,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a native `BIDIOUTLINE` field.
    ///
    /// Stored opaque instructions, cached results, and field-marker state
    /// remain metadata only. This method never reads right-to-left language,
    /// paragraph outline, or layout state; chooses a numbering system;
    /// calculates a result; or refreshes a field. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn bidi_outline_field(&self) -> Option<BidiOutlineField> {
        if self.field.field_type != FieldType::BidiOutline {
            return None;
        }
        let opaque_instructions = parse_bidi_outline_field_instructions(&self.instruction)?;
        Some(BidiOutlineField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            opaque_instructions,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a native `SHAPE` field.
    ///
    /// Stored opaque instructions, cached results, and field-marker state
    /// remain metadata only. This method never locates, links, loads,
    /// positions, lays out, or renders a drawing or canvas, or refreshes a
    /// field. Malformed instructions remain available through this generic type
    /// and return `None` here.
    #[must_use]
    pub fn shape_field(&self) -> Option<ShapeField> {
        if self.field.field_type != FieldType::Shape {
            return None;
        }
        let opaque_instructions = parse_shape_field_instructions(&self.instruction)?;
        Some(ShapeField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            opaque_instructions,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a native legacy form-code field.
    ///
    /// Stored kind, opaque instructions, cached results, and field-marker state
    /// remain metadata only. This method never reads associated form
    /// properties, fills a form, changes a selection or checkbox state, invokes
    /// entry or exit macros, or refreshes a field. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn legacy_form_field(&self) -> Option<LegacyFormField> {
        let (kind, keyword) = match self.field.field_type {
            FieldType::FormText => (LegacyFormFieldKind::Text, "FORMTEXT"),
            FieldType::FormCheckbox => (LegacyFormFieldKind::CheckBox, "FORMCHECKBOX"),
            FieldType::FormDropdown => (LegacyFormFieldKind::DropDown, "FORMDROPDOWN"),
            _ => return None,
        };
        let opaque_instructions = parse_legacy_form_field_instructions(&self.instruction, keyword)?;
        Some(LegacyFormField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            opaque_instructions,
            cached_result: self.result.clone(),
            form_data: None,
        })
    }

    /// Return inert typed metadata when this is a well-formed bookmark-reference field.
    ///
    /// Stored bookmark names, options, and cached results are never used to
    /// look up a bookmark, read a referenced range, resolve a page or note
    /// number, create a link, calculate relative position, or refresh a field.
    /// Malformed instructions remain available through this generic type and
    /// return `None` here.
    #[must_use]
    pub fn reference_field(&self) -> Option<ReferenceField> {
        let kind = match self.field.field_type {
            FieldType::Ref => ReferenceFieldKind::Reference,
            FieldType::RefWithoutKeyword => ReferenceFieldKind::ReferenceWithoutKeyword,
            FieldType::PageRef => ReferenceFieldKind::PageReference,
            FieldType::FootnoteRef => ReferenceFieldKind::FootnoteReference,
            FieldType::NoteRef => ReferenceFieldKind::NoteReference,
            _ => return None,
        };
        let parts = parse_reference_field_parts(&self.instruction, kind)?;
        Some(ReferenceField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            bookmark: parts.bookmark,
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `SET` field.
    ///
    /// Stored target names, expressions, and cached results are never used to
    /// evaluate an expression, look up or change a bookmark, change document
    /// state, or refresh a field. Malformed instructions remain available
    /// through this generic type and return `None` here.
    #[must_use]
    pub fn set_field(&self) -> Option<SetField> {
        if self.field.field_type != FieldType::Set {
            return None;
        }
        let (target_name, expression) = parse_set_field_parts(&self.instruction)?;
        Some(SetField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            target_name,
            expression,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `=` formula field.
    ///
    /// Stored formulas and cached results are never used to parse or evaluate
    /// a formula, read table cells or bookmarks, resolve field values, or
    /// refresh a field. Malformed instructions remain available through this
    /// generic type and return `None` here.
    #[must_use]
    pub fn formula_field(&self) -> Option<FormulaField> {
        if self.field.field_type != FieldType::Formula {
            return None;
        }
        let formula = parse_formula_field_formula(&self.instruction)?;
        Some(FormulaField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            formula,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `EQ` field.
    ///
    /// Stored equation syntax and cached results are never parsed, calculated,
    /// formatted, rendered, or refreshed. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn equation_field(&self) -> Option<EquationField> {
        if self.field.field_type != FieldType::Equation {
            return None;
        }
        let expression = parse_equation_field_expression(&self.instruction)?;
        Some(EquationField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            expression,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `HYPERLINK` field.
    ///
    /// Stored targets, options, and cached results are never opened, resolved,
    /// followed, activated, or refreshed. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn hyperlink_field(&self) -> Option<HyperlinkField> {
        if self.field.field_type != FieldType::Hyperlink {
            return None;
        }
        let parts = parse_hyperlink_field_parts(&self.instruction)?;
        Some(HyperlinkField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            external_target: parts.external_target,
            bookmark: parts.bookmark,
            screen_tip: parts.screen_tip,
            target_frame: parts.target_frame,
            appends_image_map_coordinates: parts.appends_image_map_coordinates,
            opens_new_window: parts.opens_new_window,
            unknown_switches: parts.unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `QUOTE` field.
    ///
    /// Stored text, switches, and cached results are never used to interpret
    /// character codes, expand nested fields, insert text, or refresh a field.
    /// Malformed instructions remain available through this generic type and
    /// return `None` here.
    #[must_use]
    pub fn quote_field(&self) -> Option<QuoteField> {
        if self.field.field_type != FieldType::Quote {
            return None;
        }
        let (text, switches) = parse_quote_field_parts(&self.instruction)?;
        Some(QuoteField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            text,
            switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `SYMBOL` field.
    ///
    /// Stored character arguments, switches, and cached results are never used
    /// to map a character code, look up a font, insert a glyph, change
    /// formatting or layout, or refresh a field. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn symbol_field(&self) -> Option<SymbolField> {
        if self.field.field_type != FieldType::Symbol {
            return None;
        }
        let (character_argument, switches) = parse_symbol_field_parts(&self.instruction)?;
        Some(SymbolField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            character_argument,
            switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a legacy automatic-numbering field.
    ///
    /// Stored kinds, switches, and cached results are never used to calculate
    /// paragraph numbers, read heading or style state, change paragraphs or
    /// layout, or refresh a field. Malformed instructions remain available
    /// through this generic type and return `None` here.
    #[must_use]
    pub fn auto_number_field(&self) -> Option<AutoNumberField> {
        let kind = AutoNumberFieldKind::from_field_type(self.field.field_type)?;
        let (instruction_kind, switches) = parse_auto_number_field_parts(&self.instruction)?;
        if instruction_kind != kind {
            return None;
        }
        Some(AutoNumberField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `LISTNUM` field.
    ///
    /// Stored optional list names, switches, and cached results are never used
    /// to look up a list, determine a level or start value, calculate a number,
    /// change layout, or refresh a field. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn list_number_field(&self) -> Option<ListNumberField> {
        if self.field.field_type != FieldType::ListNumber {
            return None;
        }
        let (list_name, switches) = parse_list_number_field_parts(&self.instruction)?;
        Some(ListNumberField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            list_name,
            switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `SEQ` field.
    ///
    /// Stored identifiers, bookmark names, tails, and cached results are never
    /// used to look up a bookmark, increment or reset a sequence, calculate a
    /// number, or refresh a field. Malformed instructions remain available
    /// through this generic type and return `None` here.
    #[must_use]
    pub fn sequence_field(&self) -> Option<SequenceField> {
        if self.field.field_type != FieldType::Sequence {
            return None;
        }
        let (identifier, bookmark, tail) = parse_sequence_field_parts(&self.instruction)?;
        Some(SequenceField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            identifier,
            bookmark,
            tail,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `STYLEREF` field.
    ///
    /// Stored style names, options, switches, and cached results are never used
    /// to look up styled text, search document stories, calculate paragraph
    /// numbers or relative positions, resolve page layout, or refresh a field.
    /// Malformed instructions remain available through this generic type and
    /// return `None` here.
    #[must_use]
    pub fn style_reference_field(&self) -> Option<StyleReferenceField> {
        if self.field.field_type != FieldType::StyleRef {
            return None;
        }
        let parts = parse_style_reference_field_parts(&self.instruction)?;
        Some(StyleReferenceField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            style_name: parts.style_name,
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `TOC` field.
    ///
    /// Stored configuration and cached results are never used to scan entries,
    /// read bookmarks, resolve links, calculate page numbers, regenerate a
    /// table of contents, or refresh a field. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn table_of_contents(&self) -> Option<TableOfContentsField> {
        if self.field.field_type != FieldType::TableOfContents {
            return None;
        }
        let parts = parse_table_of_contents_field_parts(&self.instruction)?;
        Some(TableOfContentsField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `TOA` field.
    ///
    /// Stored configuration and cached results are never used to find
    /// citations, scan hidden text, read bookmarks, calculate page numbers,
    /// paginate, regenerate a table of authorities, or refresh a field.
    /// Malformed instructions remain available through this generic type and
    /// return `None` here.
    #[must_use]
    pub fn table_of_authorities(&self) -> Option<TableOfAuthoritiesField> {
        if self.field.field_type != FieldType::TableOfAuthorities {
            return None;
        }
        let parts = parse_table_of_authorities_field_parts(&self.instruction)?;
        Some(TableOfAuthoritiesField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `INDEX` field.
    ///
    /// Stored configuration and cached results are never used to scan index
    /// markers, read bookmarks, calculate page numbers, sort entries,
    /// paginate, generate an index, or refresh a field. Malformed instructions
    /// remain available through this generic type and return `None` here.
    #[must_use]
    pub fn index(&self) -> Option<IndexField> {
        if self.field.field_type != FieldType::Index {
            return None;
        }
        let parts = parse_index_field_parts(&self.instruction)?;
        Some(IndexField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `GLOSSARY` or
    /// `AUTOTEXT` field.
    ///
    /// Stored entry names, switches, and cached results are never used to look
    /// up a building block, read a template, insert content, change bookmarks,
    /// open a resource, or refresh a field. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn auto_text_field(&self) -> Option<AutoTextField> {
        let kind = match self.field.field_type {
            FieldType::Glossary => AutoTextFieldKind::Glossary,
            FieldType::AutoText => AutoTextFieldKind::AutoText,
            _ => return None,
        };
        let parts = parse_auto_text_field_parts(&self.instruction)?;
        Some(AutoTextField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            entry_name: parts.entry_name,
            unknown_switches: parts.unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `AUTOTEXTLIST` field.
    ///
    /// Stored display text, style/tip options, and cached results are never
    /// used to show a selection UI, look up a building block, read a template,
    /// insert content, or refresh a field. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn auto_text_list_field(&self) -> Option<AutoTextListField> {
        if self.field.field_type != FieldType::AutoTextList {
            return None;
        }
        let parts = parse_auto_text_list_field_parts(&self.instruction)?;
        Some(AutoTextListField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            display_text: parts.display_text,
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `MERGEFIELD` field.
    ///
    /// The stored data-column name, switches, and cached result are never
    /// resolved against a data source, merged into the document, or refreshed.
    /// Malformed instructions remain available through this generic type and
    /// return `None` here.
    #[must_use]
    pub fn merge_field(&self) -> Option<MergeField> {
        if self.field.field_type != FieldType::MergeField {
            return None;
        }
        let (field_name, switches) = parse_merge_field_parts(&self.instruction)?;
        Some(MergeField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            field_name,
            switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `DATA` mail-merge
    /// source field.
    ///
    /// Stored data-source, header-source, and switch data are never used to
    /// open, read, connect to, resolve, or modify a source. This method never
    /// selects a record, performs a merge, or refreshes a field. Malformed
    /// instructions remain available through this generic type and return
    /// `None` here.
    #[must_use]
    pub fn mail_merge_data(&self) -> Option<MailMergeDataField> {
        if self.field.field_type != FieldType::Data {
            return None;
        }
        let (data_source, header_source, switches) =
            parse_mail_merge_data_field_parts(&self.instruction)?;
        Some(MailMergeDataField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            data_source,
            header_source,
            switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `DOCVARIABLE`
    /// field.
    ///
    /// The stored variable name, switches, and cached result are never resolved
    /// against document variables or refreshed. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn document_variable(&self) -> Option<DocumentVariableField> {
        if self.field.field_type != FieldType::DocumentVariable {
            return None;
        }
        let (variable_name, unknown_switches) =
            parse_document_variable_field_parts(&self.instruction)?;
        Some(DocumentVariableField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            variable_name,
            unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `DOCPROPERTY`
    /// field.
    ///
    /// The stored property name, switches, and cached result are never resolved
    /// against document properties or refreshed. Malformed instructions remain
    /// available through this generic type and return `None` here.
    #[must_use]
    pub fn document_property(&self) -> Option<DocumentPropertyField> {
        if self.field.field_type != FieldType::DocumentProperty {
            return None;
        }
        let (property_name, switches) = parse_document_property_field_parts(&self.instruction)?;
        Some(DocumentPropertyField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            property_name,
            switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a native `INFO` field.
    ///
    /// Stored property selectors, optional replacement values, switches, cached
    /// results, and field-marker state remain metadata only. This method never
    /// reads, resolves, modifies, or writes document or template properties, or
    /// refreshes a field. The native field type permits recognition of both
    /// explicit and keyword-omitted instruction forms. Malformed instructions
    /// remain available through this generic type and return `None` here.
    #[must_use]
    pub fn info_field(&self) -> Option<InfoField> {
        if self.field.field_type != FieldType::Info {
            return None;
        }
        let (information_type, new_value, switches) = parse_info_field_parts(&self.instruction)?;
        Some(InfoField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            information_type,
            new_value,
            switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed built-in
    /// document-information field.
    ///
    /// Built-in document-information fields retain only their native category,
    /// stored switches, cached result, and field state. This method never reads
    /// document properties or host identity data, calculates dates, revisions,
    /// or statistics, resolves a value, or refreshes a field. Malformed
    /// instructions and mismatched native field types remain available through
    /// this generic type and return `None` here.
    #[must_use]
    pub fn document_information(&self) -> Option<DocumentInformationField> {
        let native_kind = DocumentInformationFieldKind::from_field_type(self.field.field_type)?;
        let (kind, switches) = parse_document_information_field_parts(&self.instruction)?;
        if kind != native_kind {
            return None;
        }
        Some(DocumentInformationField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed built-in
    /// document-context or runtime field.
    ///
    /// `FILENAME`, `TEMPLATE`, `DATE`, `TIME`, `PAGE`, `FILESIZE`, `SECTION`, and
    /// `SECTIONPAGES` retain only their native category, stored switches, cached result,
    /// and field state. This method never reads a document path, attached
    /// template, host filesystem state or file size, current clock, or page
    /// and section layout, resolves a value, or refreshes a field. Malformed
    /// instructions and mismatched native field types remain available through
    /// this generic type and return `None` here.
    #[must_use]
    pub fn document_context(&self) -> Option<DocumentContextField> {
        let native_kind = DocumentContextFieldKind::from_field_type(self.field.field_type)?;
        let (kind, switches) = parse_document_context_field_parts(&self.instruction)?;
        if kind != native_kind {
            return None;
        }
        Some(DocumentContextField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `DDE` or
    /// `DDEAUTO` field.
    ///
    /// Stored application, source, item, and switch data are never used to
    /// launch an application, initiate a DDE conversation, open a source,
    /// request data, refresh a field, convert content, or execute code.
    /// Malformed instructions remain available through this generic type and
    /// return `None` here.
    #[must_use]
    pub fn dde_link(&self) -> Option<DdeField> {
        let parts = parse_dde_field_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, parts.kind),
            (FieldType::Dde, DdeFieldKind::Dde) | (FieldType::DdeAuto, DdeFieldKind::DdeAuto)
        ) {
            return None;
        }
        Some(DdeField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind: parts.kind,
            application: parts.application,
            source: parts.source,
            item: parts.item,
            automatic_updates: parts.automatic_updates,
            representation: parts.representation,
            omit_graphic_data: parts.omit_graphic_data,
            unknown_switches: parts.unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `LINK` field.
    ///
    /// Stored application, source, item, and switch data are never used to
    /// activate an OLE server, launch an application, open a source, request
    /// data, refresh a field, convert content, or execute code. Malformed
    /// instructions remain available through this generic type and return
    /// `None` here.
    #[must_use]
    pub fn link_field(&self) -> Option<LinkField> {
        if self.field.field_type != FieldType::Link {
            return None;
        }
        let parts = parse_link_field_parts(&self.instruction)?;
        Some(LinkField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            application_type: parts.application_type,
            source: parts.source,
            item: parts.item,
            automatic_updates: parts.automatic_updates,
            result_options: parts.result_options,
            formatting_modes: parts.formatting_modes,
            switches: parts.switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed external-include
    /// field.
    ///
    /// Stored source, bookmark, converter, and XML-option data are never used
    /// to open, resolve, import, fetch, refresh, transform, convert, evaluate,
    /// or execute source content. Malformed instructions remain available
    /// through this generic type and return `None` here.
    #[must_use]
    pub fn external_include(&self) -> Option<ExternalIncludeField> {
        let parts = parse_external_include_field_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, parts.kind),
            (
                FieldType::Include | FieldType::IncludeText,
                IncludeFieldKind::Text
            ) | (
                FieldType::Import | FieldType::IncludePicture,
                IncludeFieldKind::Picture
            )
        ) {
            return None;
        }
        Some(ExternalIncludeField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind: parts.kind,
            source: parts.source,
            bookmark: parts.bookmark,
            suppress_nested_field_updates: parts.suppress_nested_field_updates,
            omit_picture_data: parts.omit_picture_data,
            options: parts.options,
            unknown_switches: parts.unknown_switches,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed mail-merge counter.
    ///
    /// The stored kind and cached result are never used to select or count
    /// records, open a data source, perform a merge, or refresh a field
    /// result. Malformed instructions remain available through this generic
    /// type and return `None` here.
    #[must_use]
    pub fn mail_merge_counter(&self) -> Option<MailMergeCounterField> {
        let kind = parse_mail_merge_counter_kind(&self.instruction)?;
        if !matches!(
            (self.field.field_type, kind),
            (FieldType::MergeRecord, MailMergeCounterKind::Record)
                | (FieldType::MergeSequence, MailMergeCounterKind::Sequence)
        ) {
            return None;
        }
        Some(MailMergeCounterField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `NEXT` field.
    ///
    /// Cached text and field state are never used to advance a record, open a
    /// data source, perform a merge, or refresh a field result. Malformed
    /// instructions remain available through this generic type and return
    /// `None` here.
    #[must_use]
    pub fn mail_merge_next(&self) -> Option<MailMergeNextField> {
        if self.field.field_type != FieldType::Next
            || !is_mail_merge_next_instruction(&self.instruction)
        {
            return None;
        }
        Some(MailMergeNextField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            cached_result: self.result.clone(),
        })
    }

    /// Return inert metadata when this is a `NEXTIF` or `SKIPIF` field with a comparison.
    ///
    /// The stored comparison and cached result are never parsed or evaluated.
    /// This method never changes record selection, opens a data source,
    /// performs a merge, or refreshes a field result. Instructions without a
    /// comparison remain available through this generic type and return `None`
    /// here.
    #[must_use]
    pub fn mail_merge_conditional_control(&self) -> Option<MailMergeConditionalControlField> {
        let (kind, comparison) = parse_mail_merge_conditional_control_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, kind),
            (FieldType::NextIf, MailMergeConditionalControlKind::NextIf)
                | (FieldType::SkipIf, MailMergeConditionalControlKind::SkipIf)
        ) {
            return None;
        }
        Some(MailMergeConditionalControlField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            comparison,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert metadata when this is an `IF` field with an expression.
    ///
    /// The stored expression and cached result are never parsed or evaluated.
    /// This method never resolves field values or refreshes a field result.
    /// Instructions without an expression remain available through this generic
    /// type and return `None` here.
    #[must_use]
    pub fn if_field(&self) -> Option<IfField> {
        if self.field.field_type != FieldType::If {
            return None;
        }
        let expression = parse_if_field_expression(&self.instruction)?;
        Some(IfField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            expression,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert metadata when this is a `COMPARE` field with a comparison.
    ///
    /// The stored comparison and cached result are never parsed or evaluated.
    /// This method never resolves nested field values or refreshes a field
    /// result. Instructions without a comparison remain available through this
    /// generic type and return `None` here.
    #[must_use]
    pub fn compare_field(&self) -> Option<CompareField> {
        if self.field.field_type != FieldType::Compare {
            return None;
        }
        let comparison = parse_compare_field_comparison(&self.instruction)?;
        Some(CompareField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            comparison,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `ASK` or `FILLIN` field.
    ///
    /// Stored prompt, bookmark, default-response, and cached-result data are
    /// never used to display a prompt, capture a response, create or update a
    /// bookmark, perform a merge, or refresh a field. Malformed instructions
    /// remain available through this generic type and return `None` here.
    #[must_use]
    pub fn prompt_field(&self) -> Option<PromptField> {
        let (kind, bookmark, prompt, default_response, prompts_once_per_mail_merge) =
            parse_prompt_field_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, kind),
            (FieldType::Ask, PromptFieldKind::Ask) | (FieldType::FillIn, PromptFieldKind::FillIn)
        ) {
            return None;
        }
        Some(PromptField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            bookmark,
            prompt,
            default_response,
            prompts_once_per_mail_merge,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed user-identity
    /// field.
    ///
    /// Stored override, formatting, and cached-result data are never used to
    /// read or modify a host user's identity, apply formatting, or refresh a
    /// field. Malformed instructions remain available through this generic type
    /// and return `None` here.
    #[must_use]
    pub fn user_identity_field(&self) -> Option<UserIdentityField> {
        let (kind, override_value, formatting) =
            parse_user_identity_field_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, kind),
            (FieldType::UserAddress, UserIdentityFieldKind::Address)
                | (FieldType::UserInitials, UserIdentityFieldKind::Initials)
                | (FieldType::UserName, UserIdentityFieldKind::Name)
        ) {
            return None;
        }
        Some(UserIdentityField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            override_value,
            formatting,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert typed metadata when this is a well-formed `ADVANCE` field.
    ///
    /// Stored point adjustments and cached-result data are never used to move
    /// text, change layout, reflow content, or refresh a field. Malformed
    /// instructions remain available through this generic type and return
    /// `None` here.
    #[must_use]
    pub fn advance_field(&self) -> Option<AdvanceField> {
        if self.field.field_type != FieldType::Advance {
            return None;
        }
        let adjustments = parse_advance_field_adjustments(&self.instruction)?;
        Some(AdvanceField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            adjustments,
            cached_result: self.result.clone(),
        })
    }

    /// Return inert metadata when this is a well-formed `ADDRESSBLOCK` or
    /// `GREETINGLINE` field.
    ///
    /// Stored layout, locale, country, fallback, and cached-result data are
    /// never used to open a data source, select a record, perform a merge,
    /// expand placeholders, generate text, or refresh a field. Malformed
    /// instructions remain generic fields and return `None` here.
    #[must_use]
    pub fn mail_merge_recipient_field(&self) -> Option<MailMergeRecipientField> {
        let (
            kind,
            country_inclusion,
            formats_using_recipient_country,
            excluded_countries,
            format_template,
            language,
            greeting_fallback_text,
            unknown_switches,
        ) = parse_mail_merge_recipient_field_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, kind),
            (
                FieldType::AddressBlock,
                MailMergeRecipientFieldKind::AddressBlock
            ) | (
                FieldType::GreetingLine,
                MailMergeRecipientFieldKind::GreetingLine
            )
        ) {
            return None;
        }
        Some(MailMergeRecipientField {
            field: self.field.clone(),
            instruction: self.instruction.clone(),
            kind,
            country_inclusion,
            formats_using_recipient_country,
            excluded_countries,
            format_template,
            language,
            greeting_fallback_text,
            unknown_switches,
            cached_result: self.result.clone(),
        })
    }
}
