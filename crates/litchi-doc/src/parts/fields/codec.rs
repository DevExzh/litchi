//! Field instruction and PLCF wire codecs.

use super::model::corrupted;
use super::model::*;
use crate::package::Result;

impl FieldText {
    pub(crate) fn from_field<F>(field: &Field, mut text_at_range: F) -> Result<Self>
    where
        F: FnMut(u32, u32) -> Result<String>,
    {
        let instruction_start = field
            .start_cp
            .checked_add(1)
            .ok_or_else(|| corrupted("field instruction start overflows"))?;
        let instruction_end = field.separator_cp.unwrap_or(field.end_cp);
        if instruction_start > instruction_end {
            return Err(corrupted(
                "field instruction range has its start after its end",
            ));
        }
        let instruction = text_at_range(instruction_start, instruction_end)?;
        let result = match field.separator_cp {
            Some(separator) => {
                let start = separator
                    .checked_add(1)
                    .ok_or_else(|| corrupted("field result start overflows"))?;
                if start > field.end_cp {
                    return Err(corrupted("field result range has its start after its end"));
                }
                Some(text_at_range(start, field.end_cp)?)
            },
            None => None,
        };

        Ok(Self {
            field: field.clone(),
            instruction,
            result,
        })
    }

    /// Return inert typed metadata when this is a well-formed `MACROBUTTON`
    /// field.
    ///
    /// The macro or command name and button text are parsed only from stored
    /// field text. Neither is resolved, loaded, invoked, or executed.
    /// Malformed instructions remain available through this generic type and
    /// return `None` here.
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
    pub fn external_include(&self) -> Option<ExternalIncludeField> {
        let parts = parse_external_include_field_parts(&self.instruction)?;
        if !matches!(
            (self.field.field_type, parts.kind),
            (FieldType::Include, IncludeFieldKind::Text)
                | (FieldType::IncludeText, IncludeFieldKind::Text)
                | (FieldType::Import, IncludeFieldKind::Picture)
                | (FieldType::IncludePicture, IncludeFieldKind::Picture)
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

pub(super) const MAX_MACRO_BUTTON_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_GO_TO_BUTTON_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_MERGE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_MERGE_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_MAIL_MERGE_DATA_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_MAIL_MERGE_DATA_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_TABLE_OF_CONTENTS_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_TABLE_OF_CONTENTS_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_TABLE_OF_AUTHORITIES_ENTRY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_TABLE_OF_AUTHORITIES_ENTRY_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_INDEX_ENTRY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_INDEX_ENTRY_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_REFERENCED_DOCUMENT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_REFERENCED_DOCUMENT_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_PRIVATE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_TABLE_OF_AUTHORITIES_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_TABLE_OF_AUTHORITIES_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_INDEX_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_INDEX_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_REFERENCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_REFERENCE_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_SET_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_FORMULA_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_EQUATION_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_HYPERLINK_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_QUOTE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_QUOTE_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_PRINT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_EMBED_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_BARCODE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_BIDI_OUTLINE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_SHAPE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_LEGACY_FORM_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_SYMBOL_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_SYMBOL_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_AUTO_NUMBER_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_AUTO_NUMBER_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_LIST_NUMBER_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_LIST_NUMBER_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_STYLE_REFERENCE_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_AUTO_TEXT_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_AUTO_TEXT_LIST_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_DOCUMENT_VARIABLE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_DOCUMENT_VARIABLE_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_DOCUMENT_PROPERTY_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_INFO_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_INFO_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_DOCUMENT_INFORMATION_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_DOCUMENT_CONTEXT_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_DDE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_DDE_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_LINK_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_LINK_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_EXTERNAL_INCLUDE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_EXTERNAL_INCLUDE_FIELD_SWITCHES: usize = 64;
pub(super) const MAX_MAIL_MERGE_COUNTER_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_MAIL_MERGE_NEXT_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_MAIL_MERGE_CONDITIONAL_CONTROL_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_IF_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_COMPARE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_PROMPT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_USER_IDENTITY_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_ADVANCE_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_ADVANCE_FIELD_ADJUSTMENTS: usize = 64;
pub(super) const MAX_MAIL_MERGE_RECIPIENT_FIELD_INSTRUCTION_BYTES: usize = 64 * 1024;
pub(super) const MAX_MAIL_MERGE_RECIPIENT_FIELD_SWITCHES: usize = 64;

struct DdeParts {
    kind: DdeFieldKind,
    application: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    representation: Option<DdeRepresentation>,
    omit_graphic_data: bool,
    unknown_switches: Vec<MergeFieldSwitch>,
}

struct HyperlinkParts {
    external_target: Option<String>,
    bookmark: Option<String>,
    screen_tip: Option<String>,
    target_frame: Option<String>,
    appends_image_map_coordinates: bool,
    opens_new_window: bool,
    unknown_switches: Vec<MergeFieldSwitch>,
}

struct LinkParts {
    application_type: String,
    source: String,
    item: Option<String>,
    automatic_updates: bool,
    result_options: Vec<LinkResultOption>,
    formatting_modes: Vec<LinkFormatting>,
    switches: Vec<MergeFieldSwitch>,
}

struct ExternalIncludeParts {
    kind: IncludeFieldKind,
    source: String,
    bookmark: Option<String>,
    suppress_nested_field_updates: bool,
    omit_picture_data: bool,
    options: Vec<ExternalIncludeOption>,
    unknown_switches: Vec<MergeFieldSwitch>,
}

struct TableOfContentsParts {
    options: Vec<TableOfContentsOption>,
    unknown_switches: Vec<MergeFieldSwitch>,
}

pub(super) struct TableOfContentsEntryParts {
    pub(super) entry: String,
    pub(super) options: Vec<TableOfContentsEntryOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(super) struct TableOfAuthoritiesEntryParts {
    pub(super) options: Vec<TableOfAuthoritiesEntryOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(super) struct IndexEntryParts {
    pub(super) entry: String,
    pub(super) options: Vec<IndexEntryOption>,
    pub(super) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(super) struct ReferencedDocumentParts {
    pub(super) source: String,
    pub(super) relative_path: bool,
    pub(super) switches: Vec<MergeFieldSwitch>,
}

struct TableOfAuthoritiesParts {
    options: Vec<TableOfAuthoritiesOption>,
    unknown_switches: Vec<MergeFieldSwitch>,
}

struct IndexParts {
    options: Vec<IndexOption>,
    unknown_switches: Vec<MergeFieldSwitch>,
}

struct ReferenceParts {
    bookmark: String,
    options: Vec<ReferenceFieldOption>,
    unknown_switches: Vec<MergeFieldSwitch>,
}

struct StyleReferenceParts {
    style_name: String,
    options: Vec<StyleReferenceFieldOption>,
    unknown_switches: Vec<MergeFieldSwitch>,
}

struct AutoTextParts {
    entry_name: String,
    unknown_switches: Vec<MergeFieldSwitch>,
}

struct AutoTextListParts {
    display_text: Option<String>,
    options: Vec<AutoTextListOption>,
    unknown_switches: Vec<MergeFieldSwitch>,
}

fn parse_macro_button_parts(instruction: &str) -> Option<(String, String)> {
    if instruction.len() > MAX_MACRO_BUTTON_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("MACROBUTTON") {
        return None;
    }

    let macro_name = next_field_argument(instruction, &mut position).ok()??;
    if macro_name.is_empty() {
        return None;
    }
    let display_text = next_field_argument(instruction, &mut position).ok()??;
    if display_text.is_empty() {
        return None;
    }
    if next_field_argument(instruction, &mut position)
        .ok()?
        .is_some()
    {
        return None;
    }

    Some((macro_name, display_text))
}

fn parse_go_to_button_parts(instruction: &str) -> Option<(String, String)> {
    if instruction.len() > MAX_GO_TO_BUTTON_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("GOTOBUTTON") {
        return None;
    }

    let target = next_field_argument(instruction, &mut position).ok()??;
    if target.is_empty() {
        return None;
    }
    let button_text = next_field_argument(instruction, &mut position).ok()??;
    if button_text.is_empty() {
        return None;
    }
    if next_field_argument(instruction, &mut position)
        .ok()?
        .is_some()
    {
        return None;
    }

    Some((target, button_text))
}

fn parse_merge_field_parts(instruction: &str) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_MERGE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("MERGEFIELD") {
        return None;
    }

    let field_name = next_field_argument(instruction, &mut position).ok()??;
    if field_name.is_empty() {
        return None;
    }

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_MERGE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((field_name, switches))
}

fn parse_mail_merge_data_field_parts(
    instruction: &str,
) -> Option<(String, Option<String>, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_MAIL_MERGE_DATA_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("DATA") {
        return None;
    }

    let data_source = next_field_argument(instruction, &mut position).ok()??;
    if data_source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let header_source = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => {
            let source = next_field_argument(instruction, &mut position).ok()??;
            if source.is_empty() {
                return None;
            }
            Some(source)
        },
    };

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_MAIL_MERGE_DATA_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((data_source, header_source, switches))
}

fn parse_table_of_contents_field_parts(instruction: &str) -> Option<TableOfContentsParts> {
    if instruction.len() > MAX_TABLE_OF_CONTENTS_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("TOC") {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_TABLE_OF_CONTENTS_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'a' => options.push(TableOfContentsOption::CaptionWithoutLabel(
                argument.clone()?,
            )),
            'b' => options.push(TableOfContentsOption::Bookmark(argument.clone()?)),
            'c' => options.push(TableOfContentsOption::CaptionSequence(argument.clone()?)),
            'd' => options.push(TableOfContentsOption::SequencePageSeparator(
                argument.clone()?,
            )),
            'f' => options.push(TableOfContentsOption::TableEntryIdentifier(
                argument.clone()?,
            )),
            'h' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsOption::Hyperlinks);
            },
            'l' => options.push(TableOfContentsOption::TableEntryLevels(argument.clone()?)),
            'n' => options.push(TableOfContentsOption::OmitPageNumbers(argument)),
            'o' => options.push(TableOfContentsOption::HeadingStyleRange(argument)),
            'p' => options.push(TableOfContentsOption::EntryPageNumberSeparator(
                argument.clone()?,
            )),
            's' => options.push(TableOfContentsOption::SequenceIdentifier(argument.clone()?)),
            't' => options.push(TableOfContentsOption::StyleMappings(argument.clone()?)),
            'u' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsOption::OutlineLevels);
            },
            'w' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsOption::PreserveTabs);
            },
            'x' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsOption::PreserveNewlines);
            },
            'z' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsOption::HidePageNumbersInWebLayout);
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(TableOfContentsParts {
        options,
        unknown_switches,
    })
}

pub(super) fn parse_table_of_contents_entry_field_parts(
    instruction: &str,
) -> Option<TableOfContentsEntryParts> {
    if instruction.len() > MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("TC") {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    if matches!(
        peek_field_character(instruction, position),
        None | Some('\\')
    ) {
        return None;
    }
    let entry = next_field_argument(instruction, &mut position).ok()??;
    if entry.is_empty() {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_TABLE_OF_CONTENTS_ENTRY_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'f' => options.push(TableOfContentsEntryOption::ListIdentifier(argument?)),
            'l' => options.push(TableOfContentsEntryOption::Level(argument?)),
            'n' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfContentsEntryOption::OmitPageNumber);
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(TableOfContentsEntryParts {
        entry,
        options,
        unknown_switches,
    })
}

pub(super) fn parse_table_of_authorities_entry_field_parts(
    instruction: &str,
) -> Option<TableOfAuthoritiesEntryParts> {
    if instruction.len() > MAX_TABLE_OF_AUTHORITIES_ENTRY_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("TA") {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len()
                >= MAX_TABLE_OF_AUTHORITIES_ENTRY_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'b' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfAuthoritiesEntryOption::BoldPageNumber);
            },
            'c' => options.push(TableOfAuthoritiesEntryOption::Category(argument?)),
            'i' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfAuthoritiesEntryOption::ItalicPageNumber);
            },
            'l' => options.push(TableOfAuthoritiesEntryOption::LongCitation(argument?)),
            'r' => options.push(TableOfAuthoritiesEntryOption::PageRangeBookmark(argument?)),
            's' => options.push(TableOfAuthoritiesEntryOption::ShortCitation(argument?)),
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(TableOfAuthoritiesEntryParts {
        options,
        unknown_switches,
    })
}

pub(super) fn parse_index_entry_field_parts(instruction: &str) -> Option<IndexEntryParts> {
    if instruction.len() > MAX_INDEX_ENTRY_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("XE") {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    if matches!(
        peek_field_character(instruction, position),
        None | Some('\\')
    ) {
        return None;
    }
    let entry = next_field_argument(instruction, &mut position).ok()??;
    if entry.is_empty() {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_INDEX_ENTRY_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'b' => {
                if argument.is_some() {
                    return None;
                }
                options.push(IndexEntryOption::BoldPageNumber);
            },
            'f' => options.push(IndexEntryOption::EntryType(argument?)),
            'i' => {
                if argument.is_some() {
                    return None;
                }
                options.push(IndexEntryOption::ItalicPageNumber);
            },
            'r' => options.push(IndexEntryOption::PageRangeBookmark(argument?)),
            't' => options.push(IndexEntryOption::CrossReference(argument?)),
            'y' => options.push(IndexEntryOption::Yomi(argument?)),
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(IndexEntryParts {
        entry,
        options,
        unknown_switches,
    })
}

pub(super) fn parse_referenced_document_field_parts(
    instruction: &str,
) -> Option<ReferencedDocumentParts> {
    if instruction.len() > MAX_REFERENCED_DOCUMENT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("RD") {
        return None;
    }

    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    let mut relative_path = false;
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_REFERENCED_DOCUMENT_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        if name == 'f' {
            if relative_path || argument.is_some() {
                return None;
            }
            relative_path = true;
        }
        switches.push(MergeFieldSwitch { name, argument });
    }

    Some(ReferencedDocumentParts {
        source,
        relative_path,
        switches,
    })
}

pub(super) fn private_field_opaque_instructions(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_PRIVATE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let instruction = instruction.trim_start();
    let keyword = instruction.get(.."PRIVATE".len())?;
    if !keyword.eq_ignore_ascii_case("PRIVATE") {
        return None;
    }
    let remainder = instruction.get("PRIVATE".len()..)?;
    match remainder.chars().next() {
        None | Some('"') | Some('\\') => Some(remainder.trim().to_string()),
        Some(character) if character.is_whitespace() => Some(remainder.trim().to_string()),
        Some(_) => None,
    }
}

fn parse_table_of_authorities_field_parts(instruction: &str) -> Option<TableOfAuthoritiesParts> {
    if instruction.len() > MAX_TABLE_OF_AUTHORITIES_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("TOA") {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_TABLE_OF_AUTHORITIES_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'b' => options.push(TableOfAuthoritiesOption::Bookmark(argument.clone()?)),
            'c' => options.push(TableOfAuthoritiesOption::Category(argument.clone()?)),
            'd' => options.push(TableOfAuthoritiesOption::SequencePageSeparator(
                argument.clone()?,
            )),
            'e' => options.push(TableOfAuthoritiesOption::EntryPageNumberSeparator(
                argument.clone()?,
            )),
            'f' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::EntryFormatting);
            },
            'g' => options.push(TableOfAuthoritiesOption::PageRangeSeparator(
                argument.clone()?,
            )),
            'h' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::CategoryHeadings);
            },
            'l' => options.push(TableOfAuthoritiesOption::PageReferenceSeparator(
                argument.clone()?,
            )),
            'p' => {
                if argument.is_some() {
                    return None;
                }
                options.push(TableOfAuthoritiesOption::UsePassim);
            },
            's' => options.push(TableOfAuthoritiesOption::SequenceIdentifier(
                argument.clone()?,
            )),
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(TableOfAuthoritiesParts {
        options,
        unknown_switches,
    })
}

fn parse_reference_field_parts(
    instruction: &str,
    kind: ReferenceFieldKind,
) -> Option<ReferenceParts> {
    if instruction.len() > MAX_REFERENCE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let bookmark = if kind == ReferenceFieldKind::ReferenceWithoutKeyword {
        next_field_argument(instruction, &mut position).ok()??
    } else {
        let keyword = next_field_argument(instruction, &mut position).ok()??;
        let keyword_matches = match kind {
            ReferenceFieldKind::Reference => keyword.eq_ignore_ascii_case("REF"),
            ReferenceFieldKind::PageReference => keyword.eq_ignore_ascii_case("PAGEREF"),
            ReferenceFieldKind::FootnoteReference | ReferenceFieldKind::NoteReference => {
                keyword.eq_ignore_ascii_case("FTNREF") || keyword.eq_ignore_ascii_case("NOTEREF")
            },
            ReferenceFieldKind::ReferenceWithoutKeyword => false,
        };
        if !keyword_matches {
            return None;
        }
        next_field_argument(instruction, &mut position).ok()??
    };
    if bookmark.is_empty() {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let is_ref = matches!(
        kind,
        ReferenceFieldKind::Reference | ReferenceFieldKind::ReferenceWithoutKeyword
    );
    let is_note_reference = matches!(
        kind,
        ReferenceFieldKind::FootnoteReference | ReferenceFieldKind::NoteReference
    );
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_REFERENCE_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'd' if is_ref => {
                options.push(ReferenceFieldOption::SequencePageSeparator(
                    argument.clone()?,
                ));
            },
            'f' if is_ref => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::ReferencedNoteContent);
            },
            'f' if is_note_reference => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::NoteMarkFormatting);
            },
            'h' => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::Hyperlink);
            },
            'n' if is_ref => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::ParagraphNumberWithoutContext);
            },
            'p' => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::RelativePosition);
            },
            'r' if is_ref => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::ParagraphNumberRelativeContext);
            },
            't' if is_ref => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::SuppressNonNumberText);
            },
            'w' if is_ref => {
                if argument.is_some() {
                    return None;
                }
                options.push(ReferenceFieldOption::ParagraphNumberFullContext);
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(ReferenceParts {
        bookmark,
        options,
        unknown_switches,
    })
}

fn parse_set_field_parts(instruction: &str) -> Option<(String, String)> {
    if instruction.len() > MAX_SET_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("SET") {
        return None;
    }

    let target_name = next_field_argument(instruction, &mut position).ok()??;
    if target_name.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let expression = instruction.get(position..)?;
    if expression.trim().is_empty() {
        return None;
    }

    Some((target_name, expression.to_string()))
}

fn parse_formula_field_formula(instruction: &str) -> Option<Option<String>> {
    if instruction.len() > MAX_FORMULA_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let formula = instruction.trim().strip_prefix('=')?.trim();
    Some((!formula.is_empty()).then_some(formula.to_string()))
}

fn parse_equation_field_expression(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_EQUATION_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("EQ") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

fn parse_hyperlink_field_parts(instruction: &str) -> Option<HyperlinkParts> {
    if instruction.len() > MAX_HYPERLINK_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("HYPERLINK") {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let external_target = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => {
            let target = next_field_argument(instruction, &mut position).ok()??;
            if target.is_empty() {
                return None;
            }
            Some(target)
        },
    };

    let mut bookmark = None;
    let mut screen_tip = None;
    let mut target_frame = None;
    let mut appends_image_map_coordinates = false;
    let mut opens_new_window = false;
    let mut unknown_switches = Vec::new();
    let mut switch_count = 0;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switch_count >= MAX_HYPERLINK_FIELD_SWITCHES {
            return None;
        }
        switch_count += 1;

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };

        let slot = match name {
            'l' => &mut bookmark,
            'o' => &mut screen_tip,
            't' => &mut target_frame,
            'm' => {
                if appends_image_map_coordinates || argument.is_some() {
                    return None;
                }
                appends_image_map_coordinates = true;
                continue;
            },
            'n' => {
                if opens_new_window || argument.is_some() {
                    return None;
                }
                opens_new_window = true;
                continue;
            },
            _ => {
                unknown_switches.push(MergeFieldSwitch { name, argument });
                continue;
            },
        };
        let value = argument?;
        if value.is_empty() || slot.replace(value).is_some() {
            return None;
        }
    }

    if external_target.is_none() && bookmark.is_none() {
        return None;
    }

    Some(HyperlinkParts {
        external_target,
        bookmark,
        screen_tip,
        target_frame,
        appends_image_map_coordinates,
        opens_new_window,
        unknown_switches,
    })
}

fn parse_print_field_instructions(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_PRINT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("PRINT") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

fn parse_embed_field_instructions(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_EMBED_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("EMBED") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

fn parse_barcode_field_instructions(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_BARCODE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("BARCODE") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

fn parse_bidi_outline_field_instructions(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_BIDI_OUTLINE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("BIDIOUTLINE") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

fn parse_shape_field_instructions(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_SHAPE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("SHAPE") {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

fn parse_legacy_form_field_instructions(
    instruction: &str,
    expected_keyword: &str,
) -> Option<String> {
    if instruction.len() > MAX_LEGACY_FORM_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case(expected_keyword) {
        return None;
    }
    Some(instruction.get(position..)?.trim().to_string())
}

fn parse_quote_field_parts(instruction: &str) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_QUOTE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("QUOTE") {
        return None;
    }

    let text = next_field_argument(instruction, &mut position).ok()??;
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_QUOTE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    Some((text, switches))
}

fn parse_symbol_field_parts(instruction: &str) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_SYMBOL_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("SYMBOL") {
        return None;
    }

    let character_argument = next_field_argument(instruction, &mut position).ok()??;
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_SYMBOL_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    Some((character_argument, switches))
}

fn parse_auto_number_field_parts(
    instruction: &str,
) -> Option<(AutoNumberFieldKind, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_AUTO_NUMBER_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = AutoNumberFieldKind::from_keyword(&keyword)?;
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_AUTO_NUMBER_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    Some((kind, switches))
}

fn parse_list_number_field_parts(
    instruction: &str,
) -> Option<(Option<String>, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_LIST_NUMBER_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("LISTNUM") {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let list_name = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_LIST_NUMBER_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    Some((list_name, switches))
}

fn parse_sequence_field_parts(instruction: &str) -> Option<(String, Option<String>, String)> {
    if instruction.len() > MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("SEQ") {
        return None;
    }

    let identifier = next_field_argument(instruction, &mut position).ok()??;
    if identifier.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let bookmark = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => {
            let bookmark = next_field_argument(instruction, &mut position).ok()??;
            if bookmark.is_empty() {
                return None;
            }
            Some(bookmark)
        },
    };

    skip_field_whitespace(instruction, &mut position);
    let tail = instruction.get(position..)?.trim().to_string();
    Some((identifier, bookmark, tail))
}

fn parse_style_reference_field_parts(instruction: &str) -> Option<StyleReferenceParts> {
    if instruction.len() > MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("STYLEREF") {
        return None;
    }

    let style_name = next_field_argument(instruction, &mut position).ok()??;
    if style_name.is_empty() {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_STYLE_REFERENCE_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'l' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::FollowingText);
            },
            'n' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::ParagraphNumber);
            },
            'p' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::RelativePosition);
            },
            'r' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::ParagraphNumberRelativeContext);
            },
            't' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::SuppressNonNumberText);
            },
            'w' => {
                if argument.is_some() {
                    return None;
                }
                options.push(StyleReferenceFieldOption::ParagraphNumberFullContext);
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(StyleReferenceParts {
        style_name,
        options,
        unknown_switches,
    })
}

fn parse_index_field_parts(instruction: &str) -> Option<IndexParts> {
    if instruction.len() > MAX_INDEX_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("INDEX") {
        return None;
    }

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || options.len() + unknown_switches.len() >= MAX_INDEX_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'b' => options.push(IndexOption::Bookmark(argument.clone()?)),
            'c' => options.push(IndexOption::Columns(argument.clone()?)),
            'd' => options.push(IndexOption::SequencePageSeparator(argument.clone()?)),
            'e' => options.push(IndexOption::EntryPageNumberSeparator(argument.clone()?)),
            'f' => options.push(IndexOption::EntryType(argument.clone()?)),
            'g' => options.push(IndexOption::PageRangeSeparator(argument.clone()?)),
            'h' => options.push(IndexOption::Heading(argument.clone()?)),
            'k' => options.push(IndexOption::CrossReferenceSeparator(argument.clone()?)),
            'l' => options.push(IndexOption::PageNumberSeparator(argument.clone()?)),
            'o' => options.push(IndexOption::EastAsianSortOrder(argument.clone()?)),
            'p' => options.push(IndexOption::LetterRange(argument.clone()?)),
            'r' => {
                if argument.is_some() {
                    return None;
                }
                options.push(IndexOption::RunIn);
            },
            's' => options.push(IndexOption::SequenceIdentifier(argument.clone()?)),
            'y' => {
                if argument.is_some() {
                    return None;
                }
                options.push(IndexOption::UseYomi);
            },
            'z' => options.push(IndexOption::LanguageId(argument.clone()?)),
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(IndexParts {
        options,
        unknown_switches,
    })
}

fn parse_auto_text_field_parts(instruction: &str) -> Option<AutoTextParts> {
    if instruction.len() > MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("GLOSSARY") && !keyword.eq_ignore_ascii_case("AUTOTEXT") {
        return None;
    }
    let entry_name = next_field_argument(instruction, &mut position).ok()??;
    if entry_name.is_empty() {
        return None;
    }

    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || unknown_switches.len() >= MAX_AUTO_TEXT_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        unknown_switches.push(MergeFieldSwitch { name, argument });
    }

    Some(AutoTextParts {
        entry_name,
        unknown_switches,
    })
}

fn parse_auto_text_list_field_parts(instruction: &str) -> Option<AutoTextListParts> {
    if instruction.len() > MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("AUTOTEXTLIST") {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let display_text = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\'
            || options.len() + unknown_switches.len() >= MAX_AUTO_TEXT_LIST_FIELD_SWITCHES
        {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            's' => options.push(AutoTextListOption::Style(argument.clone()?)),
            't' => options.push(AutoTextListOption::Tip(argument.clone()?)),
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(AutoTextListParts {
        display_text,
        options,
        unknown_switches,
    })
}

fn parse_document_variable_field_parts(
    instruction: &str,
) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_DOCUMENT_VARIABLE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("DOCVARIABLE") {
        return None;
    }

    let variable_name = next_field_argument(instruction, &mut position).ok()??;
    if variable_name.is_empty() {
        return None;
    }

    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || unknown_switches.len() >= MAX_DOCUMENT_VARIABLE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        unknown_switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((variable_name, unknown_switches))
}

fn parse_document_property_field_parts(
    instruction: &str,
) -> Option<(String, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_DOCUMENT_PROPERTY_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("DOCPROPERTY") {
        return None;
    }

    let property_name = next_field_argument(instruction, &mut position).ok()??;
    if property_name.is_empty() {
        return None;
    }

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_DOCUMENT_PROPERTY_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((property_name, switches))
}

fn parse_info_field_parts(
    instruction: &str,
) -> Option<(String, Option<String>, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_INFO_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let first_argument = next_field_argument(instruction, &mut position).ok()??;
    let information_type = if first_argument.eq_ignore_ascii_case("INFO") {
        next_field_argument(instruction, &mut position).ok()??
    } else {
        first_argument
    };
    if information_type.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let new_value = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_INFO_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((information_type, new_value, switches))
}

fn parse_document_information_field_parts(
    instruction: &str,
) -> Option<(DocumentInformationFieldKind, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_DOCUMENT_INFORMATION_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = DocumentInformationFieldKind::from_keyword(&keyword)?;

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_DOCUMENT_INFORMATION_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((kind, switches))
}

fn parse_document_context_field_parts(
    instruction: &str,
) -> Option<(DocumentContextFieldKind, Vec<MergeFieldSwitch>)> {
    if instruction.len() > MAX_DOCUMENT_CONTEXT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = DocumentContextFieldKind::from_keyword(&keyword)?;

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_DOCUMENT_CONTEXT_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch {
            name: name.to_ascii_lowercase(),
            argument,
        });
    }

    Some((kind, switches))
}

fn parse_dde_field_parts(instruction: &str) -> Option<DdeParts> {
    if instruction.len() > MAX_DDE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("DDE") {
        DdeFieldKind::Dde
    } else if keyword.eq_ignore_ascii_case("DDEAUTO") {
        DdeFieldKind::DdeAuto
    } else {
        return None;
    };

    let application = next_field_argument(instruction, &mut position).ok()??;
    if application.is_empty() {
        return None;
    }
    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let item = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut automatic_updates = kind == DdeFieldKind::DdeAuto;
    let mut saw_automatic_update = false;
    let mut representation = None;
    let mut omit_graphic_data = false;
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || unknown_switches.len() >= MAX_DDE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'a' if kind == DdeFieldKind::Dde => {
                if saw_automatic_update || argument.is_some() {
                    return None;
                }
                automatic_updates = true;
                saw_automatic_update = true;
            },
            'a' => return None,
            'd' => {
                if representation.is_some() || omit_graphic_data || argument.is_some() {
                    return None;
                }
                omit_graphic_data = true;
            },
            'b' | 'h' | 'p' | 'r' | 't' | 'u' => {
                if representation.is_some() || omit_graphic_data || argument.is_some() {
                    return None;
                }
                representation = Some(match name {
                    'b' => DdeRepresentation::Bitmap,
                    'h' => DdeRepresentation::Html,
                    'p' => DdeRepresentation::Picture,
                    'r' => DdeRepresentation::RichText,
                    't' => DdeRepresentation::Text,
                    'u' => DdeRepresentation::UnicodeText,
                    _ => unreachable!("DDE representation switch was matched above"),
                });
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(DdeParts {
        kind,
        application,
        source,
        item,
        automatic_updates,
        representation,
        omit_graphic_data,
        unknown_switches,
    })
}

fn parse_link_field_parts(instruction: &str) -> Option<LinkParts> {
    if instruction.len() > MAX_LINK_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("LINK") {
        return None;
    }

    let application_type = next_field_argument(instruction, &mut position).ok()??;
    if application_type.is_empty() {
        return None;
    }
    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let item = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_LINK_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    let mut automatic_updates = false;
    let mut result_options = Vec::new();
    let mut formatting_modes = Vec::new();
    for switch in &switches {
        match switch.name {
            'a' => {
                if switch.argument.is_some() {
                    return None;
                }
                automatic_updates = true;
            },
            'f' => {
                let value = switch.argument.as_deref()?.parse::<i64>().ok()?;
                formatting_modes.push(match value {
                    0 => LinkFormatting::Source,
                    2 => LinkFormatting::Destination,
                    4 => LinkFormatting::SpreadsheetSource,
                    5 => LinkFormatting::SpreadsheetDestination,
                    other => LinkFormatting::Unsupported(other),
                });
            },
            'b' | 'd' | 'h' | 'p' | 'r' | 't' | 'u' => {
                if switch.argument.is_some() {
                    return None;
                }
                result_options.push(match switch.name {
                    'b' => LinkResultOption::Bitmap,
                    'd' => LinkResultOption::OmitGraphicData,
                    'h' => LinkResultOption::Html,
                    'p' => LinkResultOption::Picture,
                    'r' => LinkResultOption::RichText,
                    't' => LinkResultOption::Text,
                    'u' => LinkResultOption::UnicodeText,
                    _ => unreachable!("LINK result switch was matched above"),
                });
            },
            _ => {},
        }
    }

    Some(LinkParts {
        application_type,
        source,
        item,
        automatic_updates,
        result_options,
        formatting_modes,
        switches,
    })
}

fn parse_external_include_field_parts(instruction: &str) -> Option<ExternalIncludeParts> {
    if instruction.len() > MAX_EXTERNAL_INCLUDE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind =
        if keyword.eq_ignore_ascii_case("INCLUDETEXT") || keyword.eq_ignore_ascii_case("INCLUDE") {
            IncludeFieldKind::Text
        } else if keyword.eq_ignore_ascii_case("INCLUDEPICTURE")
            || keyword.eq_ignore_ascii_case("IMPORT")
        {
            IncludeFieldKind::Picture
        } else {
            return None;
        };

    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let bookmark = match (kind, peek_field_character(instruction, position)) {
        (IncludeFieldKind::Text, None | Some('\\')) => None,
        (IncludeFieldKind::Text, Some(_)) => {
            Some(next_field_argument(instruction, &mut position).ok()??)
        },
        (IncludeFieldKind::Picture, _) => None,
    };

    let mut suppress_nested_field_updates = false;
    let mut omit_picture_data = false;
    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut switch_count = 0;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switch_count >= MAX_EXTERNAL_INCLUDE_FIELD_SWITCHES {
            return None;
        }
        switch_count += 1;

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };

        match (kind, name) {
            (_, 'c') => options.push(ExternalIncludeOption::Converter(argument?)),
            (IncludeFieldKind::Text, 'e') => {
                options.push(ExternalIncludeOption::Encoding(argument?));
            },
            (IncludeFieldKind::Text, 'm') => {
                options.push(ExternalIncludeOption::MimeType(argument?));
            },
            (IncludeFieldKind::Text, 'n') => {
                options.push(ExternalIncludeOption::NamespaceMapping(argument?));
            },
            (IncludeFieldKind::Text, 't') => {
                options.push(ExternalIncludeOption::Xslt(argument?));
            },
            (IncludeFieldKind::Text, 'x') => {
                options.push(ExternalIncludeOption::XPath(argument?));
            },
            (IncludeFieldKind::Text, '!') => {
                if suppress_nested_field_updates || argument.is_some() {
                    return None;
                }
                suppress_nested_field_updates = true;
            },
            (IncludeFieldKind::Picture, 'd') => {
                if omit_picture_data || argument.is_some() {
                    return None;
                }
                omit_picture_data = true;
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(ExternalIncludeParts {
        kind,
        source,
        bookmark,
        suppress_nested_field_updates,
        omit_picture_data,
        options,
        unknown_switches,
    })
}

fn parse_mail_merge_counter_kind(instruction: &str) -> Option<MailMergeCounterKind> {
    if instruction.len() > MAX_MAIL_MERGE_COUNTER_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("MERGEREC") {
        MailMergeCounterKind::Record
    } else if keyword.eq_ignore_ascii_case("MERGESEQ") {
        MailMergeCounterKind::Sequence
    } else {
        return None;
    };
    if next_field_argument(instruction, &mut position)
        .ok()?
        .is_some()
    {
        return None;
    }

    Some(kind)
}

fn is_mail_merge_next_instruction(instruction: &str) -> bool {
    if instruction.len() > MAX_MAIL_MERGE_NEXT_INSTRUCTION_BYTES {
        return false;
    }

    let mut position = 0;
    let Ok(Some(keyword)) = next_field_argument(instruction, &mut position) else {
        return false;
    };
    keyword.eq_ignore_ascii_case("NEXT")
        && matches!(next_field_argument(instruction, &mut position), Ok(None))
}

fn parse_mail_merge_conditional_control_parts(
    instruction: &str,
) -> Option<(MailMergeConditionalControlKind, String)> {
    if instruction.len() > MAX_MAIL_MERGE_CONDITIONAL_CONTROL_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("NEXTIF") {
        MailMergeConditionalControlKind::NextIf
    } else if keyword.eq_ignore_ascii_case("SKIPIF") {
        MailMergeConditionalControlKind::SkipIf
    } else {
        return None;
    };
    let comparison = instruction.get(position..)?.trim();
    (!comparison.is_empty()).then_some((kind, comparison.to_string()))
}

fn parse_if_field_expression(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_IF_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("IF") {
        return None;
    }
    let expression = instruction.get(position..)?.trim();
    (!expression.is_empty()).then_some(expression.to_string())
}

fn parse_compare_field_comparison(instruction: &str) -> Option<String> {
    if instruction.len() > MAX_COMPARE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("COMPARE") {
        return None;
    }
    let comparison = instruction.get(position..)?.trim();
    (!comparison.is_empty()).then_some(comparison.to_string())
}

#[allow(clippy::type_complexity)]
fn parse_prompt_field_parts(
    instruction: &str,
) -> Option<(
    PromptFieldKind,
    Option<String>,
    Option<String>,
    Option<String>,
    bool,
)> {
    if instruction.len() > MAX_PROMPT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("ASK") {
        PromptFieldKind::Ask
    } else if keyword.eq_ignore_ascii_case("FILLIN") {
        PromptFieldKind::FillIn
    } else {
        return None;
    };

    let (bookmark, prompt) = match kind {
        PromptFieldKind::Ask => {
            let bookmark = next_field_argument(instruction, &mut position).ok()??;
            if bookmark.is_empty() {
                return None;
            }
            let prompt = next_field_argument(instruction, &mut position).ok()??;
            (Some(bookmark), Some(prompt))
        },
        PromptFieldKind::FillIn => {
            skip_field_whitespace(instruction, &mut position);
            let prompt = match peek_field_character(instruction, position) {
                None | Some('\\') => None,
                Some(_) => next_field_argument(instruction, &mut position).ok()?,
            };
            (None, prompt)
        },
    };

    let mut default_response = None;
    let mut prompts_once_per_mail_merge = false;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        match name.to_ascii_lowercase() {
            'd' => {
                if default_response.is_some() {
                    return None;
                }
                default_response = Some(next_field_argument(instruction, &mut position).ok()??);
            },
            'o' => {
                if prompts_once_per_mail_merge {
                    return None;
                }
                skip_field_whitespace(instruction, &mut position);
                if !matches!(
                    peek_field_character(instruction, position),
                    None | Some('\\')
                ) {
                    return None;
                }
                prompts_once_per_mail_merge = true;
            },
            _ => return None,
        }
    }

    Some((
        kind,
        bookmark,
        prompt,
        default_response,
        prompts_once_per_mail_merge,
    ))
}

fn parse_user_identity_field_parts(
    instruction: &str,
) -> Option<(
    UserIdentityFieldKind,
    Option<String>,
    Option<UserIdentityFormatting>,
)> {
    if instruction.len() > MAX_USER_IDENTITY_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("USERADDRESS") {
        UserIdentityFieldKind::Address
    } else if keyword.eq_ignore_ascii_case("USERINITIALS") {
        UserIdentityFieldKind::Initials
    } else if keyword.eq_ignore_ascii_case("USERNAME") {
        UserIdentityFieldKind::Name
    } else {
        return None;
    };

    skip_field_whitespace(instruction, &mut position);
    let override_value = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut formatting = None;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' {
            return None;
        }
        let name = next_field_character(instruction, &mut position)?;
        if name != '*' || formatting.is_some() {
            return None;
        }
        let value = next_field_argument(instruction, &mut position).ok()??;
        formatting = Some(if value.eq_ignore_ascii_case("Caps") {
            UserIdentityFormatting::Caps
        } else if value.eq_ignore_ascii_case("FirstCap") {
            UserIdentityFormatting::FirstCap
        } else if value.eq_ignore_ascii_case("Lower") {
            UserIdentityFormatting::Lower
        } else if value.eq_ignore_ascii_case("Upper") {
            UserIdentityFormatting::Upper
        } else {
            return None;
        });
    }

    Some((kind, override_value, formatting))
}

fn parse_advance_field_adjustments(instruction: &str) -> Option<Vec<AdvanceFieldAdjustment>> {
    if instruction.len() > MAX_ADVANCE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("ADVANCE") {
        return None;
    }

    let mut adjustments = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' {
            return None;
        }
        let name = next_field_character(instruction, &mut position)?;
        let operation = match name.to_ascii_lowercase() {
            'd' => AdvanceFieldOperation::Down,
            'l' => AdvanceFieldOperation::Left,
            'r' => AdvanceFieldOperation::Right,
            'u' => AdvanceFieldOperation::Up,
            'x' => AdvanceFieldOperation::HorizontalPosition,
            'y' => AdvanceFieldOperation::VerticalPosition,
            _ => return None,
        };
        if adjustments.len() >= MAX_ADVANCE_FIELD_ADJUSTMENTS {
            return None;
        }
        let points = next_field_argument(instruction, &mut position)
            .ok()??
            .parse::<i64>()
            .ok()?;
        adjustments.push(AdvanceFieldAdjustment { operation, points });
    }

    Some(adjustments)
}

#[allow(clippy::type_complexity)]
fn parse_mail_merge_recipient_field_parts(
    instruction: &str,
) -> Option<(
    MailMergeRecipientFieldKind,
    Option<AddressBlockCountryInclusion>,
    bool,
    Vec<String>,
    Option<String>,
    Option<String>,
    Option<String>,
    Vec<MergeFieldSwitch>,
)> {
    if instruction.len() > MAX_MAIL_MERGE_RECIPIENT_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("ADDRESSBLOCK") {
        MailMergeRecipientFieldKind::AddressBlock
    } else if keyword.eq_ignore_ascii_case("GREETINGLINE") {
        MailMergeRecipientFieldKind::GreetingLine
    } else {
        return None;
    };

    let mut country_inclusion = None;
    let mut formats_using_recipient_country = false;
    let mut excluded_countries = Vec::new();
    let mut format_template = None;
    let mut language = None;
    let mut greeting_fallback_text = None;
    let mut unknown_switches = Vec::new();
    let mut switch_count = 0;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switch_count >= MAX_MAIL_MERGE_RECIPIENT_FIELD_SWITCHES {
            return None;
        }
        switch_count += 1;

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        match (kind, name.to_ascii_lowercase()) {
            (MailMergeRecipientFieldKind::AddressBlock, 'c') => {
                if country_inclusion.is_some() {
                    return None;
                }
                let value = next_field_argument(instruction, &mut position).ok()??;
                country_inclusion = Some(match value.as_str() {
                    "0" => AddressBlockCountryInclusion::Omit,
                    "1" => AddressBlockCountryInclusion::Always,
                    "2" => AddressBlockCountryInclusion::UnlessExcluded,
                    _ => return None,
                });
            },
            (MailMergeRecipientFieldKind::AddressBlock, 'd') => {
                if formats_using_recipient_country {
                    return None;
                }
                skip_field_whitespace(instruction, &mut position);
                if !matches!(
                    peek_field_character(instruction, position),
                    None | Some('\\')
                ) {
                    return None;
                }
                formats_using_recipient_country = true;
            },
            (MailMergeRecipientFieldKind::AddressBlock, 'e') => {
                excluded_countries.push(next_field_argument(instruction, &mut position).ok()??);
            },
            (_, 'f') => {
                if format_template.is_some() {
                    return None;
                }
                format_template = Some(next_field_argument(instruction, &mut position).ok()??);
            },
            (_, 'l') => {
                if language.is_some() {
                    return None;
                }
                language = Some(next_field_argument(instruction, &mut position).ok()??);
            },
            (MailMergeRecipientFieldKind::GreetingLine, 'c' | 'e') => {
                if greeting_fallback_text.is_some() {
                    return None;
                }
                greeting_fallback_text =
                    Some(next_field_argument(instruction, &mut position).ok()??);
            },
            _ => {
                skip_field_whitespace(instruction, &mut position);
                let argument = match peek_field_character(instruction, position) {
                    None | Some('\\') => None,
                    Some(_) => next_field_argument(instruction, &mut position).ok()?,
                };
                unknown_switches.push(MergeFieldSwitch {
                    name: name.to_ascii_lowercase(),
                    argument,
                });
            },
        }
    }

    Some((
        kind,
        country_inclusion,
        formats_using_recipient_country,
        excluded_countries,
        format_template,
        language,
        greeting_fallback_text,
        unknown_switches,
    ))
}

fn next_field_argument(
    input: &str,
    position: &mut usize,
) -> std::result::Result<Option<String>, ()> {
    skip_field_whitespace(input, position);
    let Some(first) = next_field_character(input, position) else {
        return Ok(None);
    };

    if first != '"' {
        *position -= first.len_utf8();
        let mut value = String::new();
        while let Some(character) = next_field_character(input, position) {
            if character.is_whitespace() || character == '"' {
                *position -= character.len_utf8();
                break;
            }
            if character == '\\' {
                let escaped = next_field_character(input, position).ok_or(())?;
                if !matches!(escaped, '"' | '\\') {
                    return Err(());
                }
                value.push(escaped);
            } else {
                value.push(character);
            }
        }
        return Ok(Some(value));
    }

    let mut value = String::new();
    loop {
        let character = next_field_character(input, position).ok_or(())?;
        match character {
            '"' => return Ok(Some(value)),
            '\\' => {
                let escaped = next_field_character(input, position).ok_or(())?;
                if !matches!(escaped, '"' | '\\') {
                    return Err(());
                }
                value.push(escaped);
            },
            _ => value.push(character),
        }
    }
}

fn skip_field_whitespace(input: &str, position: &mut usize) {
    while let Some(character) = input.get(*position..).and_then(|rest| rest.chars().next()) {
        if !character.is_whitespace() {
            break;
        }
        *position += character.len_utf8();
    }
}

fn next_field_character(input: &str, position: &mut usize) -> Option<char> {
    let character = input.get(*position..)?.chars().next()?;
    *position += character.len_utf8();
    Some(character)
}

fn peek_field_character(input: &str, position: usize) -> Option<char> {
    input.get(position..)?.chars().next()
}
