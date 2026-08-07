//! Dynamic text-field variants and their semantic XML projection.

#![allow(
    clippy::wildcard_imports,
    reason = "semantic field owners share the stable model facade namespace"
)]
use super::*;
/// Typed conditional and placeholder text metadata in document order.
///
/// Formula strings are retained verbatim and are never evaluated. `display_text`
/// is the cached text stored by the document producer.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum DynamicTextField {
    Placeholder {
        placeholder_type: PlaceholderType,
        description: Option<String>,
        display_text: String,
    },
    ConditionalText {
        condition: String,
        value_if_true: String,
        value_if_false: String,
        current_value: Option<bool>,
        display_text: String,
    },
    HiddenText {
        condition: String,
        string_value: String,
        is_hidden: Option<bool>,
        display_text: String,
    },
    HiddenParagraph {
        condition: String,
        is_hidden: Option<bool>,
        display_text: String,
    },
    /// An inert calculated sequence field from ODF 1.2 section 7.4.11.
    Sequence {
        name: String,
        formula: Option<String>,
        number_format: Option<SequenceNumberFormat>,
        reference_name: Option<String>,
        display_text: String,
    },
    /// A cached reference to a named sequence value.
    SequenceReference {
        reference_name: String,
        reference_format: Option<SequenceReferenceFormat>,
        display_text: String,
    },
    VariableSet {
        name: String,
        formula: Option<String>,
        value: CalculatedFieldValue,
        display: Option<VariableSetDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    VariableGet {
        name: String,
        display: Option<FormulaFieldDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    Expression {
        formula: Option<String>,
        value: Option<CalculatedFieldValue>,
        display: Option<FormulaFieldDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    VariableInput {
        name: String,
        description: Option<String>,
        value_type: FieldValueType,
        display: Option<VariableSetDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    UserFieldGet {
        name: String,
        display: Option<UserFieldDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    UserFieldInput {
        name: String,
        description: Option<String>,
        data_style_name: Option<String>,
        display_text: String,
    },
    TextInput {
        description: Option<String>,
        display_text: String,
    },
    /// An inert drop-down input field with stored choice metadata.
    ///
    /// The labels and cached selected text are retained exactly as document
    /// metadata. No selection interface is shown and no label is selected,
    /// changed, or resolved by this API.
    DropDown {
        name: String,
        labels: Vec<DropDownLabel>,
        display_text: String,
    },
    /// An inert inline script declaration.
    ///
    /// Linked targets and embedded payloads are retained as document metadata
    /// only. This API never opens, resolves, or executes either form.
    Script {
        /// Optional inert external script reference.
        href: Option<String>,
        /// Optional producer-supplied script-language identifier.
        language: Option<String>,
        /// The stored inline script payload, if any.
        content: String,
    },
    /// An inert table-cell formula display field.
    TableFormula {
        formula: Option<String>,
        display: Option<FormulaFieldDisplay>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Cached, non-calculating measurement field text.
    Measure {
        kind: MeasureKind,
        display_text: String,
    },
    Reference {
        reference_name: Option<String>,
        reference_format: Option<CrossReferenceFormat>,
        display_text: String,
    },
    BookmarkReference {
        reference_name: Option<String>,
        reference_format: Option<CrossReferenceFormat>,
        display_text: String,
    },
    NoteReference {
        reference_name: Option<String>,
        note_class: NoteReferenceClass,
        reference_format: Option<NoteReferenceFormat>,
        display_text: String,
    },
    DocumentStatistic {
        kind: StatisticKind,
        number_format: Option<SequenceNumberFormat>,
        display_text: String,
    },
    /// Current, previous, or next page number with inert cached presentation.
    PageNumber {
        number_format: Option<SequenceNumberFormat>,
        fixed: Option<bool>,
        page_adjust: Option<i64>,
        select_page: Option<PageSelection>,
        display_text: String,
    },
    /// Current date or an explicitly fixed date/date-time value.
    Date {
        value: Option<FieldDateValue>,
        adjustment: Option<FieldDuration>,
        fixed: Option<bool>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Current time or an explicitly fixed time/date-time value.
    Time {
        value: Option<FieldTimeValue>,
        adjustment: Option<FieldDuration>,
        fixed: Option<bool>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Previous/next page continuation reminder.
    PageContinuation {
        select_page: PageContinuationSelection,
        string_value: Option<String>,
        display_text: String,
    },
    /// Set or disable the document's single alternative page variable.
    PageVariableSet {
        active: Option<bool>,
        page_adjust: Option<i64>,
        display_text: String,
    },
    /// Display the current alternative page-variable value.
    PageVariableGet {
        number_format: Option<SequenceNumberFormat>,
        display_text: String,
    },
    /// Cached filename presentation; never reads a host path or document location.
    FileName {
        display: Option<FileNameDisplay>,
        fixed: Option<bool>,
        display_text: String,
    },
    /// Cached template presentation; never opens or locates a template resource.
    TemplateName {
        display: Option<TemplateNameDisplay>,
        display_text: String,
    },
    /// Cached active spreadsheet sheet label; never resolves live sheet state.
    SheetName { display_text: String },
    /// Cached chapter presentation; never resolves or updates the document outline.
    Chapter {
        display: Option<ChapterDisplay>,
        outline_level: Option<NonNegativeInteger>,
        display_text: String,
    },
    /// Cached presentation and optional fixed value of a metadata field.
    DocumentMetadata {
        kind: MetadataFieldKind,
        value: Option<MetadataFieldValue>,
        fixed: Option<bool>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Fixed or live cached string metadata such as title or creator.
    ///
    /// Author fields retain stored text only and never read or modify host
    /// identity data.
    DocumentIdentity {
        kind: IdentityFieldKind,
        fixed: Option<bool>,
        display_text: String,
    },
    /// Cached subsequent-author identity/contact data.
    ///
    /// These fields never read or modify host identity or contact data, even when
    /// `text:fixed` is omitted or false.
    Sender {
        kind: SenderFieldKind,
        fixed: Option<bool>,
        display_text: String,
    },
    /// Named custom document metadata with inert cached typed attributes.
    UserDefinedMetadata {
        name: String,
        values: UserDefinedMetadataValues,
        fixed: Option<bool>,
        data_style_name: Option<String>,
        display_text: String,
    },
    /// Cached data from a named DDE declaration; never refreshed or connected.
    DdeConnection {
        connection_name: String,
        display_text: String,
    },
    /// RDF-backed metadata field with namespace-resolved inert inline content.
    MetaField {
        xml_id: String,
        data_style_name: Option<String>,
        content: MetaFieldContent,
    },
}

impl DynamicTextField {
    /// The cached text present in the ODF file, without evaluating any formula.
    pub fn display_text(&self) -> &str {
        match self {
            Self::Placeholder { display_text, .. }
            | Self::ConditionalText { display_text, .. }
            | Self::HiddenText { display_text, .. }
            | Self::HiddenParagraph { display_text, .. }
            | Self::Sequence { display_text, .. }
            | Self::SequenceReference { display_text, .. }
            | Self::VariableSet { display_text, .. }
            | Self::VariableGet { display_text, .. }
            | Self::Expression { display_text, .. }
            | Self::VariableInput { display_text, .. }
            | Self::UserFieldGet { display_text, .. }
            | Self::UserFieldInput { display_text, .. }
            | Self::TextInput { display_text, .. }
            | Self::DropDown { display_text, .. }
            | Self::TableFormula { display_text, .. }
            | Self::Measure { display_text, .. }
            | Self::Reference { display_text, .. }
            | Self::BookmarkReference { display_text, .. }
            | Self::NoteReference { display_text, .. }
            | Self::DocumentStatistic { display_text, .. }
            | Self::DdeConnection { display_text, .. }
            | Self::PageNumber { display_text, .. }
            | Self::Date { display_text, .. }
            | Self::Time { display_text, .. }
            | Self::PageContinuation { display_text, .. }
            | Self::PageVariableSet { display_text, .. }
            | Self::PageVariableGet { display_text, .. }
            | Self::FileName { display_text, .. }
            | Self::TemplateName { display_text, .. }
            | Self::SheetName { display_text, .. }
            | Self::Chapter { display_text, .. }
            | Self::DocumentMetadata { display_text, .. }
            | Self::DocumentIdentity { display_text, .. }
            | Self::Sender { display_text, .. }
            | Self::UserDefinedMetadata { display_text, .. } => display_text,
            Self::Script { content, .. } => content,
            Self::MetaField { content, .. } => content.display_text(),
        }
    }

    /// Effective `text:active` value for a page-variable setter.
    ///
    /// ODF and `LibreOffice` default the omitted attribute to `true`.
    pub fn effective_page_variable_active(&self) -> Option<bool> {
        match self {
            Self::PageVariableSet { active, .. } => Some(active.unwrap_or(true)),
            _ => None,
        }
    }

    /// Effective page adjustment for a page-variable setter.
    ///
    /// The standard default for an omitted adjustment is zero.
    pub fn effective_page_variable_adjustment(&self) -> Option<i64> {
        match self {
            Self::PageVariableSet { page_adjust, .. } => Some(page_adjust.unwrap_or(0)),
            _ => None,
        }
    }

    /// Validate this field for safe ODF XML serialization.
    ///
    /// Conditions remain opaque strings: validation never parses or evaluates a
    /// formula. It only enforces required values, bounded allocation sizes, and
    /// XML 1.0 character validity.
    pub fn validate(&self) -> Result<()> {
        let mut aggregate = 0usize;
        match self {
            Self::Placeholder {
                description,
                display_text,
                ..
            } => {
                validate_dynamic_value(
                    "text:description",
                    description.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "placeholder display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::ConditionalText {
                condition,
                value_if_true,
                value_if_false,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:condition", Some(condition), true, &mut aggregate)?;
                validate_dynamic_value(
                    "text:string-value-if-true",
                    Some(value_if_true),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "text:string-value-if-false",
                    Some(value_if_false),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "conditional display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::HiddenText {
                condition,
                string_value,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:condition", Some(condition), true, &mut aggregate)?;
                validate_dynamic_value(
                    "text:string-value",
                    Some(string_value),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "hidden text display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::HiddenParagraph {
                condition,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:condition", Some(condition), true, &mut aggregate)?;
                validate_dynamic_value(
                    "hidden paragraph display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::DdeConnection {
                connection_name,
                display_text,
            } => {
                validate_dynamic_value(
                    "text:connection-name",
                    Some(connection_name),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "DDE cached text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Sequence {
                name,
                formula,
                number_format,
                reference_name,
                display_text,
            } => {
                validate_dynamic_value("text:name", Some(name), false, &mut aggregate)?;
                validate_dynamic_value("text:formula", formula.as_deref(), false, &mut aggregate)?;
                if let Some(number_format) = number_format {
                    number_format.validate()?;
                    aggregate = aggregate
                        .checked_add(number_format.format().len())
                        .ok_or_else(|| {
                            Error::InvalidFormat("dynamic field size overflow".to_string())
                        })?;
                }
                validate_dynamic_value(
                    "text:ref-name",
                    reference_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "sequence display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
                if aggregate > MAX_DYNAMIC_FIELD_AGGREGATE {
                    return Err(Error::InvalidFormat(format!(
                        "dynamic field exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} aggregate bytes"
                    )));
                }
            },
            Self::SequenceReference {
                reference_name,
                display_text,
                ..
            } => {
                validate_dynamic_value(
                    "text:ref-name",
                    Some(reference_name),
                    true,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "sequence reference display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::VariableSet {
                name,
                formula,
                value,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:name", Some(name), true, &mut aggregate)?;
                validate_dynamic_value("text:formula", formula.as_deref(), false, &mut aggregate)?;
                value.validate(&mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "variable-set display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::VariableGet {
                name,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:name", Some(name), true, &mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "variable-get display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Expression {
                formula,
                value,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:formula", formula.as_deref(), false, &mut aggregate)?;
                if let Some(value) = value {
                    value.validate(&mut aggregate)?;
                }
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "expression display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::VariableInput {
                name,
                description,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:name", Some(name), true, &mut aggregate)?;
                validate_dynamic_value(
                    "text:description",
                    description.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "variable-input display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::UserFieldGet {
                name,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:name", Some(name), true, &mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "user-field-get display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::UserFieldInput {
                name,
                description,
                data_style_name,
                display_text,
            } => {
                validate_dynamic_value("text:name", Some(name), true, &mut aggregate)?;
                validate_dynamic_value(
                    "text:description",
                    description.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "user-field-input display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::TextInput {
                description,
                display_text,
            } => {
                validate_dynamic_value(
                    "text:description",
                    description.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "text-input display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::DropDown {
                name,
                labels,
                display_text,
            } => {
                validate_dynamic_value("text:name", Some(name), false, &mut aggregate)?;
                if labels.len() > MAX_DROP_DOWN_LABELS {
                    return Err(Error::InvalidFormat(format!(
                        "text:drop-down exceeds {MAX_DROP_DOWN_LABELS} labels"
                    )));
                }
                for label in labels {
                    validate_dynamic_value(
                        "text:label text:value",
                        label.value.as_deref(),
                        false,
                        &mut aggregate,
                    )?;
                }
                validate_dynamic_value(
                    "drop-down display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Script {
                href,
                language,
                content,
            } => {
                validate_dynamic_value("xlink:href", href.as_deref(), false, &mut aggregate)?;
                validate_dynamic_value(
                    "script:language",
                    language.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "inline script content",
                    Some(content),
                    false,
                    &mut aggregate,
                )?;
                if href.is_some() && !content.is_empty() {
                    return Err(Error::InvalidFormat(
                        "text:script cannot combine xlink:href with inline content".to_string(),
                    ));
                }
            },
            Self::TableFormula {
                formula,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:formula", formula.as_deref(), false, &mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "table-formula display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Measure { display_text, .. } => {
                validate_dynamic_value(
                    "measure display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Reference {
                reference_name,
                display_text,
                ..
            }
            | Self::BookmarkReference {
                reference_name,
                display_text,
                ..
            }
            | Self::NoteReference {
                reference_name,
                display_text,
                ..
            } => {
                validate_dynamic_value(
                    "text:ref-name",
                    reference_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "reference display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::DocumentStatistic {
                number_format,
                display_text,
                ..
            } => {
                if let Some(number_format) = number_format {
                    number_format.validate()?;
                    aggregate = aggregate
                        .checked_add(number_format.format().len())
                        .ok_or_else(|| {
                            Error::InvalidFormat("dynamic field size overflow".to_string())
                        })?;
                }
                validate_dynamic_value(
                    "statistic display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
                if aggregate > MAX_DYNAMIC_FIELD_AGGREGATE {
                    return Err(Error::InvalidFormat(format!(
                        "dynamic field exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} aggregate bytes"
                    )));
                }
            },
            Self::PageNumber {
                number_format,
                display_text,
                ..
            } => {
                if let Some(number_format) = number_format {
                    number_format.validate()?;
                    aggregate = aggregate
                        .checked_add(number_format.format().len())
                        .ok_or_else(|| {
                            Error::InvalidFormat("dynamic field size overflow".to_string())
                        })?;
                }
                validate_dynamic_value(
                    "page-number display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Date {
                value,
                adjustment,
                data_style_name,
                display_text,
                ..
            } => {
                if let Some(value) = value {
                    value.validate(&mut aggregate)?;
                }
                if let Some(adjustment) = adjustment {
                    adjustment.validate("text:date-adjust", &mut aggregate)?;
                }
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "date display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Time {
                value,
                adjustment,
                data_style_name,
                display_text,
                ..
            } => {
                if let Some(value) = value {
                    value.validate(&mut aggregate)?;
                }
                if let Some(adjustment) = adjustment {
                    adjustment.validate("text:time-adjust", &mut aggregate)?;
                }
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "time display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::PageContinuation {
                string_value,
                display_text,
                ..
            } => {
                validate_dynamic_value(
                    "text:string-value",
                    string_value.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "page-continuation display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::PageVariableSet { display_text, .. } => {
                validate_dynamic_value(
                    "page-variable-set display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::PageVariableGet {
                number_format,
                display_text,
            } => {
                if let Some(number_format) = number_format {
                    number_format.validate()?;
                    aggregate = aggregate
                        .checked_add(number_format.format().len())
                        .ok_or_else(|| {
                            Error::InvalidFormat("dynamic field size overflow".to_string())
                        })?;
                }
                validate_dynamic_value(
                    "page-variable-get display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::FileName { display_text, .. } => {
                validate_dynamic_value(
                    "file-name display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::TemplateName { display_text, .. } => {
                validate_dynamic_value(
                    "template-name display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::SheetName { display_text } => {
                validate_dynamic_value(
                    "sheet-name display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Chapter {
                outline_level,
                display_text,
                ..
            } => {
                validate_dynamic_value(
                    "text:outline-level",
                    outline_level.as_ref().map(NonNegativeInteger::as_str),
                    true,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "chapter display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::DocumentMetadata {
                kind,
                value,
                data_style_name,
                display_text,
                ..
            } => {
                validate_document_metadata_value(*kind, value.as_ref(), &mut aggregate)?;
                if !kind.permits_data_style() && data_style_name.is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "{} does not permit style:data-style-name",
                        kind.element_name()
                    )));
                }
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "document metadata display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::DocumentIdentity { display_text, .. } => {
                validate_dynamic_value(
                    "document identity display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::Sender { display_text, .. } => {
                validate_dynamic_value(
                    "sender display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::UserDefinedMetadata {
                name,
                values,
                data_style_name,
                display_text,
                ..
            } => {
                validate_dynamic_value("text:name", Some(name), false, &mut aggregate)?;
                values.validate(&mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                validate_dynamic_value(
                    "user-defined metadata display text",
                    Some(display_text),
                    false,
                    &mut aggregate,
                )?;
            },
            Self::MetaField {
                xml_id,
                data_style_name,
                content,
            } => {
                validate_xml_id(xml_id)?;
                validate_dynamic_value("xml:id", Some(xml_id), true, &mut aggregate)?;
                validate_dynamic_value(
                    "style:data-style-name",
                    data_style_name.as_deref(),
                    false,
                    &mut aggregate,
                )?;
                let rebuilt = MetaFieldContent::new(content.nodes().to_vec())?;
                if &rebuilt != content {
                    return Err(Error::InvalidFormat(
                        "inconsistent text:meta-field content cache".to_string(),
                    ));
                }
            },
        }
        Ok(())
    }

    /// Serialize one self-contained ODF field element.
    ///
    /// The returned fragment declares the `text` namespace locally, so it is
    /// namespace-correct regardless of the prefixes used by its destination
    /// document. Formula attributes are emitted verbatim after XML escaping and
    /// are never executed.
    pub fn to_xml_fragment(&self) -> Result<String> {
        if let Self::Script {
            href,
            language,
            content,
        } = self
        {
            self.validate()?;
            let mut xml = String::from("<text:script xmlns:text=\"");
            xml.push_str(TEXT_DATABASE_NAMESPACE);
            xml.push('"');
            if let Some(href) = href {
                xml.push_str(" xmlns:xlink=\"");
                xml.push_str(XLINK_NAMESPACE);
                xml.push_str("\" xlink:type=\"simple\" xlink:href=\"");
                push_xml_attribute(&mut xml, href);
                xml.push('"');
            }
            if let Some(language) = language {
                xml.push_str(" xmlns:script=\"");
                xml.push_str(SCRIPT_NAMESPACE);
                xml.push_str("\" script:language=\"");
                push_xml_attribute(&mut xml, language);
                xml.push('"');
            }
            if content.is_empty() {
                xml.push_str("/>");
            } else {
                xml.push('>');
                push_xml_text(&mut xml, content);
                xml.push_str("</text:script>");
            }
            return Ok(xml);
        }
        if let Self::DropDown {
            name,
            labels,
            display_text,
        } = self
        {
            self.validate()?;
            let mut xml = String::from("<text:drop-down xmlns:text=\"");
            xml.push_str(TEXT_DATABASE_NAMESPACE);
            xml.push_str("\" text:name=\"");
            push_xml_attribute(&mut xml, name);
            xml.push('\"');
            if labels.is_empty() && display_text.is_empty() {
                xml.push_str("/>");
                return Ok(xml);
            }
            xml.push('>');
            for label in labels {
                xml.push_str("<text:label");
                if let Some(value) = label.value.as_deref() {
                    xml.push_str(" text:value=\"");
                    push_xml_attribute(&mut xml, value);
                    xml.push('\"');
                }
                if let Some(current_selected) = label.current_selected {
                    xml.push_str(" text:current-selected=\"");
                    xml.push_str(if current_selected { "true" } else { "false" });
                    xml.push('\"');
                }
                xml.push_str("/>");
            }
            push_xml_text(&mut xml, display_text);
            xml.push_str("</text:drop-down>");
            return Ok(xml);
        }
        if let Self::MetaField {
            xml_id,
            data_style_name,
            content,
        } = self
        {
            self.validate()?;
            let mut xml = String::new();
            xml.push_str("<text:meta-field xmlns:text=\"");
            xml.push_str(TEXT_DATABASE_NAMESPACE);
            xml.push_str("\" xml:id=\"");
            push_xml_attribute(&mut xml, xml_id);
            xml.push('"');
            if let Some(data_style_name) = data_style_name {
                xml.push_str(" xmlns:style=\"");
                xml.push_str(STYLE_NAMESPACE);
                xml.push_str("\" style:data-style-name=\"");
                push_xml_attribute(&mut xml, data_style_name);
                xml.push('"');
            }
            xml.push('>');
            content.write_xml(&mut xml);
            xml.push_str("</text:meta-field>");
            return Ok(xml);
        }
        Ok(self.to_element()?.to_xml_string())
    }

    pub(crate) fn to_element(&self) -> Result<Element> {
        self.validate()?;
        let mut element = match self {
            Self::Placeholder { .. } => Element::new("text:placeholder"),
            Self::ConditionalText { .. } => Element::new("text:conditional-text"),
            Self::HiddenText { .. } => Element::new("text:hidden-text"),
            Self::HiddenParagraph { .. } => Element::new("text:hidden-paragraph"),
            Self::DdeConnection { .. } => Element::new("text:dde-connection"),
            Self::Sequence { .. } => Element::new("text:sequence"),
            Self::SequenceReference { .. } => Element::new("text:sequence-ref"),
            Self::VariableSet { .. } => Element::new("text:variable-set"),
            Self::VariableGet { .. } => Element::new("text:variable-get"),
            Self::Expression { .. } => Element::new("text:expression"),
            Self::VariableInput { .. } => Element::new("text:variable-input"),
            Self::UserFieldGet { .. } => Element::new("text:user-field-get"),
            Self::UserFieldInput { .. } => Element::new("text:user-field-input"),
            Self::TextInput { .. } => Element::new("text:text-input"),
            Self::DropDown { .. } => unreachable!("drop-down uses nested-label serializer"),
            Self::Script { .. } => unreachable!("script uses a namespace-aware serializer"),
            Self::TableFormula { .. } => Element::new("text:table-formula"),
            Self::Measure { .. } => Element::new("text:measure"),
            Self::Reference { .. } => Element::new("text:reference-ref"),
            Self::BookmarkReference { .. } => Element::new("text:bookmark-ref"),
            Self::NoteReference { .. } => Element::new("text:note-ref"),
            Self::DocumentStatistic { kind, .. } => Element::new(kind.element_name()),
            Self::PageNumber { .. } => Element::new("text:page-number"),
            Self::Date { .. } => Element::new("text:date"),
            Self::Time { .. } => Element::new("text:time"),
            Self::PageContinuation { .. } => Element::new("text:page-continuation"),
            Self::PageVariableSet { .. } => Element::new("text:page-variable-set"),
            Self::PageVariableGet { .. } => Element::new("text:page-variable-get"),
            Self::FileName { .. } => Element::new("text:file-name"),
            Self::TemplateName { .. } => Element::new("text:template-name"),
            Self::SheetName { .. } => Element::new("text:sheet-name"),
            Self::Chapter { .. } => Element::new("text:chapter"),
            Self::DocumentMetadata { kind, .. } => Element::new(kind.element_name()),
            Self::DocumentIdentity { kind, .. } => Element::new(kind.element_name()),
            Self::Sender { kind, .. } => Element::new(kind.element_name()),
            Self::UserDefinedMetadata { .. } => Element::new("text:user-defined"),
            Self::MetaField { .. } => unreachable!("meta-field uses ordered mixed serializer"),
        };
        element.set_attribute("xmlns:text", TEXT_DATABASE_NAMESPACE);

        match self {
            Self::Placeholder {
                placeholder_type,
                description,
                display_text,
            } => {
                element.set_attribute("text:placeholder-type", placeholder_type.as_str());
                if let Some(description) = description {
                    element.set_attribute("text:description", description);
                }
                element.set_text(display_text);
            },
            Self::ConditionalText {
                condition,
                value_if_true,
                value_if_false,
                current_value,
                display_text,
            } => {
                element.set_attribute("text:condition", condition);
                element.set_attribute("text:string-value-if-true", value_if_true);
                element.set_attribute("text:string-value-if-false", value_if_false);
                if let Some(current_value) = current_value {
                    element.set_attribute(
                        "text:current-value",
                        if *current_value { "true" } else { "false" },
                    );
                }
                element.set_text(display_text);
            },
            Self::HiddenText {
                condition,
                string_value,
                is_hidden,
                display_text,
            } => {
                element.set_attribute("text:condition", condition);
                element.set_attribute("text:string-value", string_value);
                if let Some(is_hidden) = is_hidden {
                    element
                        .set_attribute("text:is-hidden", if *is_hidden { "true" } else { "false" });
                }
                element.set_text(display_text);
            },
            Self::HiddenParagraph {
                condition,
                is_hidden,
                display_text,
            } => {
                element.set_attribute("text:condition", condition);
                if let Some(is_hidden) = is_hidden {
                    element
                        .set_attribute("text:is-hidden", if *is_hidden { "true" } else { "false" });
                }
                element.set_text(display_text);
            },
            Self::DdeConnection {
                connection_name,
                display_text,
            } => {
                element.set_attribute("text:connection-name", connection_name);
                element.set_text(display_text);
            },
            Self::Sequence {
                name,
                formula,
                number_format,
                reference_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(formula) = formula {
                    element.set_attribute("text:formula", formula);
                }
                if let Some(number_format) = number_format {
                    element.set_attribute(
                        "xmlns:style",
                        "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
                    );
                    element.set_attribute("style:num-format", number_format.format());
                    if let Some(letter_sync) = number_format.letter_sync() {
                        element.set_attribute(
                            "style:num-letter-sync",
                            if letter_sync { "true" } else { "false" },
                        );
                    }
                }
                if let Some(reference_name) = reference_name {
                    element.set_attribute("text:ref-name", reference_name);
                }
                element.set_text(display_text);
            },
            Self::SequenceReference {
                reference_name,
                reference_format,
                display_text,
            } => {
                element.set_attribute("text:ref-name", reference_name);
                if let Some(reference_format) = reference_format {
                    element.set_attribute("text:reference-format", reference_format.as_str());
                }
                element.set_text(display_text);
            },
            Self::VariableSet {
                name,
                formula,
                value,
                display,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(formula) = formula {
                    element.set_attribute("text:formula", formula);
                }
                value.write_attributes(&mut element);
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::VariableGet {
                name,
                display,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::Expression {
                formula,
                value,
                display,
                data_style_name,
                display_text,
            } => {
                if let Some(formula) = formula {
                    element.set_attribute("text:formula", formula);
                }
                if let Some(value) = value {
                    value.write_attributes(&mut element);
                }
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::VariableInput {
                name,
                description,
                value_type,
                display,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(description) = description {
                    element.set_attribute("text:description", description);
                }
                element.set_attribute("xmlns:office", OFFICE_NAMESPACE);
                element.set_attribute("office:value-type", value_type.as_str());
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::UserFieldGet {
                name,
                display,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::UserFieldInput {
                name,
                description,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                if let Some(description) = description {
                    element.set_attribute("text:description", description);
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::TextInput {
                description,
                display_text,
            } => {
                if let Some(description) = description {
                    element.set_attribute("text:description", description);
                }
                element.set_text(display_text);
            },
            Self::DropDown { .. } => unreachable!("drop-down uses nested-label serializer"),
            Self::Script { .. } => unreachable!("script uses a namespace-aware serializer"),
            Self::TableFormula {
                formula,
                display,
                data_style_name,
                display_text,
            } => {
                if let Some(formula) = formula {
                    element.set_attribute("text:formula", formula);
                }
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::Measure { kind, display_text } => {
                element.set_attribute("text:kind", kind.as_str());
                element.set_text(display_text);
            },
            Self::Reference {
                reference_name,
                reference_format,
                display_text,
            }
            | Self::BookmarkReference {
                reference_name,
                reference_format,
                display_text,
            } => {
                if let Some(reference_name) = reference_name {
                    element.set_attribute("text:ref-name", reference_name);
                }
                if let Some(reference_format) = reference_format {
                    element.set_attribute("text:reference-format", reference_format.as_str());
                }
                element.set_text(display_text);
            },
            Self::NoteReference {
                reference_name,
                note_class,
                reference_format,
                display_text,
            } => {
                if let Some(reference_name) = reference_name {
                    element.set_attribute("text:ref-name", reference_name);
                }
                element.set_attribute("text:note-class", note_class.as_str());
                if let Some(reference_format) = reference_format {
                    element.set_attribute("text:reference-format", reference_format.as_str());
                }
                element.set_text(display_text);
            },
            Self::DocumentStatistic {
                number_format,
                display_text,
                ..
            } => {
                if let Some(number_format) = number_format {
                    element.set_attribute("xmlns:style", STYLE_NAMESPACE);
                    element.set_attribute("style:num-format", number_format.format());
                    if let Some(letter_sync) = number_format.letter_sync() {
                        element.set_attribute(
                            "style:num-letter-sync",
                            if letter_sync { "true" } else { "false" },
                        );
                    }
                }
                element.set_text(display_text);
            },
            Self::PageNumber {
                number_format,
                fixed,
                page_adjust,
                select_page,
                display_text,
            } => {
                if let Some(number_format) = number_format {
                    element.set_attribute("xmlns:style", STYLE_NAMESPACE);
                    element.set_attribute("style:num-format", number_format.format());
                    if let Some(letter_sync) = number_format.letter_sync() {
                        element.set_attribute(
                            "style:num-letter-sync",
                            if letter_sync { "true" } else { "false" },
                        );
                    }
                }
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                if let Some(page_adjust) = page_adjust {
                    element.set_attribute("text:page-adjust", &page_adjust.to_string());
                }
                if let Some(select_page) = select_page {
                    element.set_attribute("text:select-page", select_page.as_str());
                }
                element.set_text(display_text);
            },
            Self::Date {
                value,
                adjustment,
                fixed,
                data_style_name,
                display_text,
            } => {
                if let Some(value) = value {
                    element.set_attribute("text:date-value", value.as_str());
                }
                if let Some(adjustment) = adjustment {
                    element.set_attribute("text:date-adjust", adjustment.as_str());
                }
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::Time {
                value,
                adjustment,
                fixed,
                data_style_name,
                display_text,
            } => {
                if let Some(value) = value {
                    element.set_attribute("text:time-value", value.as_str());
                }
                if let Some(adjustment) = adjustment {
                    element.set_attribute("text:time-adjust", adjustment.as_str());
                }
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::PageContinuation {
                select_page,
                string_value,
                display_text,
            } => {
                element.set_attribute("text:select-page", select_page.as_str());
                if let Some(string_value) = string_value {
                    element.set_attribute("text:string-value", string_value);
                }
                element.set_text(display_text);
            },
            Self::PageVariableSet {
                active,
                page_adjust,
                display_text,
            } => {
                if let Some(active) = active {
                    element.set_attribute("text:active", if *active { "true" } else { "false" });
                }
                if let Some(page_adjust) = page_adjust {
                    element.set_attribute("text:page-adjust", &page_adjust.to_string());
                }
                element.set_text(display_text);
            },
            Self::PageVariableGet {
                number_format,
                display_text,
            } => {
                if let Some(number_format) = number_format {
                    element.set_attribute("xmlns:style", STYLE_NAMESPACE);
                    element.set_attribute("style:num-format", number_format.format());
                    if let Some(letter_sync) = number_format.letter_sync() {
                        element.set_attribute(
                            "style:num-letter-sync",
                            if letter_sync { "true" } else { "false" },
                        );
                    }
                }
                element.set_text(display_text);
            },
            Self::FileName {
                display,
                fixed,
                display_text,
            } => {
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                element.set_text(display_text);
            },
            Self::TemplateName {
                display,
                display_text,
            } => {
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                element.set_text(display_text);
            },
            Self::SheetName { display_text } => {
                element.set_text(display_text);
            },
            Self::Chapter {
                display,
                outline_level,
                display_text,
            } => {
                if let Some(display) = display {
                    element.set_attribute("text:display", display.as_str());
                }
                if let Some(outline_level) = outline_level {
                    element.set_attribute("text:outline-level", outline_level.as_str());
                }
                element.set_text(display_text);
            },
            Self::DocumentMetadata {
                kind,
                value,
                fixed,
                data_style_name,
                display_text,
            } => {
                if let Some(value) = value {
                    match value {
                        MetadataFieldValue::Date(value) => {
                            element.set_attribute("text:date-value", value.as_str());
                        },
                        MetadataFieldValue::Time(value) => {
                            element.set_attribute("text:time-value", value.as_str());
                        },
                        MetadataFieldValue::Duration(value) => {
                            element.set_attribute("text:duration", value.as_str());
                        },
                    }
                }
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                if kind.permits_data_style() {
                    set_data_style(&mut element, data_style_name.as_deref());
                }
                element.set_text(display_text);
            },
            Self::DocumentIdentity {
                fixed,
                display_text,
                ..
            } => {
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                element.set_text(display_text);
            },
            Self::Sender {
                fixed,
                display_text,
                ..
            } => {
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                element.set_text(display_text);
            },
            Self::UserDefinedMetadata {
                name,
                values,
                fixed,
                data_style_name,
                display_text,
            } => {
                element.set_attribute("text:name", name);
                values.write_attributes(&mut element);
                if let Some(fixed) = fixed {
                    element.set_attribute("text:fixed", if *fixed { "true" } else { "false" });
                }
                set_data_style(&mut element, data_style_name.as_deref());
                element.set_text(display_text);
            },
            Self::MetaField { .. } => unreachable!("meta-field uses ordered mixed serializer"),
        }
        Ok(element)
    }
}
