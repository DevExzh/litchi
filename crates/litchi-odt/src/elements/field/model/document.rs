//! Document field handles and element-level field parsing.

#![allow(
    clippy::wildcard_imports,
    reason = "semantic field owners share the stable model facade namespace"
)]
use super::*;
/// Represents a text field in the document
#[derive(Debug, Clone)]
pub struct Field {
    element: Element,
}

impl Field {
    /// Create a new field from an element
    pub fn from_element(element: Element) -> Result<Self> {
        let tag = element.tag_name();
        if !Self::is_field_tag(tag) {
            return Err(Error::InvalidFormat(format!(
                "Element {} is not a field",
                tag
            )));
        }
        Ok(Self { element })
    }

    /// Check if a tag name represents a field
    pub fn is_field_tag(tag: &str) -> bool {
        matches!(
            tag,
            "text:page-number"
                | "text:page-count"
                | "text:page-continuation"
                | "text:page-variable-set"
                | "text:page-variable-get"
                | "text:date"
                | "text:time"
                | "text:file-name"
                | "text:template-name"
                | "text:sheet-name"
                | "text:author-name"
                | "text:author-initials"
                | "text:sender-firstname"
                | "text:sender-lastname"
                | "text:sender-initials"
                | "text:sender-title"
                | "text:sender-position"
                | "text:sender-email"
                | "text:sender-phone-private"
                | "text:sender-fax"
                | "text:sender-company"
                | "text:sender-phone-work"
                | "text:sender-street"
                | "text:sender-city"
                | "text:sender-postal-code"
                | "text:sender-country"
                | "text:sender-state-or-province"
                | "text:chapter"
                | "text:title"
                | "text:subject"
                | "text:keywords"
                | "text:description"
                | "text:user-defined"
                | "text:creator"
                | "text:initial-creator"
                | "text:printed-by"
                | "text:creation-date"
                | "text:creation-time"
                | "text:modification-date"
                | "text:modification-time"
                | "text:print-date"
                | "text:print-time"
                | "text:editing-duration"
                | "text:editing-cycles"
                | "text:reference-ref"
                | "text:sequence-ref"
                | "text:bookmark-ref"
                | "text:note-ref"
                | "text:variable-set"
                | "text:variable-get"
                | "text:variable-input"
                | "text:user-field-get"
                | "text:user-field-input"
                | "text:sequence"
                | "text:expression"
                | "text:text-input"
                | "text:drop-down"
                | "text:script"
                | "text:placeholder"
                | "text:conditional-text"
                | "text:hidden-text"
                | "text:hidden-paragraph"
                | "text:measure"
                | "text:meta-field"
                | "text:dde-connection"
                | "text:table-formula"
                | "text:database-display"
                | "text:database-next"
                | "text:database-row-select"
                | "text:database-row-number"
                | "text:database-name"
                | "text:word-count"
                | "text:paragraph-count"
                | "text:character-count"
                | "text:table-count"
                | "text:image-count"
                | "text:object-count"
        )
    }

    /// Get the field type
    pub fn field_type(&self) -> &str {
        self.element.tag_name()
    }

    /// Get the field value (text content)
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }

    /// Get the field display format
    pub fn format(&self) -> Option<&str> {
        self.element
            .get_attribute("style:data-style-name")
            .or_else(|| self.element.get_attribute("number:style"))
    }

    /// Get the field name (for named fields like variables or user fields)
    pub fn name(&self) -> Option<&str> {
        self.element
            .get_attribute("text:name")
            .or_else(|| self.element.get_attribute("text:variable-name"))
    }

    /// Get reference target (for reference fields)
    pub fn reference_target(&self) -> Option<&str> {
        self.element
            .get_attribute("text:ref-name")
            .or_else(|| self.element.get_attribute("text:reference-name"))
    }

    /// Convert a conditional-content field to its strict typed representation.
    ///
    /// Returns `Ok(None)` for other field kinds. Conditions remain inert strings.
    pub fn dynamic_text_field(&self) -> Result<Option<DynamicTextField>> {
        let text = || self.value();
        let result = match self.field_type() {
            "text:placeholder" => DynamicTextField::Placeholder {
                placeholder_type: PlaceholderType::parse(required_field_attribute(
                    self,
                    "text:placeholder-type",
                )?)?,
                description: self
                    .element
                    .get_attribute("text:description")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:conditional-text" => DynamicTextField::ConditionalText {
                condition: required_field_attribute(self, "text:condition")?.to_owned(),
                value_if_true: required_field_attribute(self, "text:string-value-if-true")?
                    .to_owned(),
                value_if_false: required_field_attribute(self, "text:string-value-if-false")?
                    .to_owned(),
                current_value: optional_field_bool(self, "text:current-value")?,
                display_text: text(),
            },
            "text:hidden-text" => DynamicTextField::HiddenText {
                condition: required_field_attribute(self, "text:condition")?.to_owned(),
                string_value: required_field_attribute(self, "text:string-value")?.to_owned(),
                is_hidden: optional_field_bool(self, "text:is-hidden")?,
                display_text: text(),
            },
            "text:hidden-paragraph" => DynamicTextField::HiddenParagraph {
                condition: required_field_attribute(self, "text:condition")?.to_owned(),
                is_hidden: optional_field_bool(self, "text:is-hidden")?,
                display_text: text(),
            },
            "text:script" => {
                reject_unknown_field_attributes(
                    self,
                    &["xlink:type", "xlink:href", "script:language"],
                )?;
                let href = match (
                    self.element.get_attribute("xlink:type"),
                    self.element.get_attribute("xlink:href"),
                ) {
                    (None, None) => None,
                    (Some("simple"), Some(href)) => Some(href.to_owned()),
                    (Some("simple"), None) => {
                        return Err(Error::InvalidFormat(
                            "text:script xlink:type requires xlink:href".to_string(),
                        ));
                    },
                    (None, Some(_)) => {
                        return Err(Error::InvalidFormat(
                            "text:script xlink:href requires xlink:type='simple'".to_string(),
                        ));
                    },
                    (Some(kind), Some(_)) => {
                        return Err(Error::InvalidFormat(format!(
                            "text:script xlink:type must be 'simple', got '{kind}'"
                        )));
                    },
                    (Some(kind), None) => {
                        return Err(Error::InvalidFormat(format!(
                            "text:script xlink:type must be 'simple', got '{kind}'"
                        )));
                    },
                };
                let result = DynamicTextField::Script {
                    href,
                    language: self
                        .element
                        .get_attribute("script:language")
                        .map(str::to_owned),
                    content: text(),
                };
                result.validate()?;
                result
            },
            "text:dde-connection" => DynamicTextField::DdeConnection {
                connection_name: required_field_attribute(self, "text:connection-name")?.to_owned(),
                display_text: text(),
            },
            "text:sequence" => {
                let format = self.element.get_attribute("style:num-format");
                let letter_sync = optional_field_bool(self, "style:num-letter-sync")?;
                let number_format = match (format, letter_sync) {
                    (Some(format), letter_sync) => {
                        Some(SequenceNumberFormat::new(format, letter_sync)?)
                    },
                    (None, Some(_)) => {
                        return Err(Error::InvalidFormat(
                            "style:num-letter-sync requires style:num-format".to_string(),
                        ));
                    },
                    (None, None) => None,
                };
                DynamicTextField::Sequence {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    formula: self
                        .element
                        .get_attribute("text:formula")
                        .map(str::to_owned),
                    number_format,
                    reference_name: self
                        .element
                        .get_attribute("text:ref-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:sequence-ref" => DynamicTextField::SequenceReference {
                reference_name: required_field_attribute(self, "text:ref-name")?.to_owned(),
                reference_format: self
                    .element
                    .get_attribute("text:reference-format")
                    .map(SequenceReferenceFormat::parse)
                    .transpose()?,
                display_text: text(),
            },
            "text:variable-set" => DynamicTextField::VariableSet {
                name: required_field_attribute(self, "text:name")?.to_owned(),
                formula: self
                    .element
                    .get_attribute("text:formula")
                    .map(str::to_owned),
                value: parse_calculated_value(self, true)?.expect("required calculated value"),
                display: self
                    .element
                    .get_attribute("text:display")
                    .map(VariableSetDisplay::parse)
                    .transpose()?,
                data_style_name: self
                    .element
                    .get_attribute("style:data-style-name")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:variable-get" => {
                reject_calculated_value_attributes(self)?;
                DynamicTextField::VariableGet {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(FormulaFieldDisplay::parse)
                        .transpose()?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:expression" => DynamicTextField::Expression {
                formula: self
                    .element
                    .get_attribute("text:formula")
                    .map(str::to_owned),
                value: parse_calculated_value(self, false)?,
                display: self
                    .element
                    .get_attribute("text:display")
                    .map(FormulaFieldDisplay::parse)
                    .transpose()?,
                data_style_name: self
                    .element
                    .get_attribute("style:data-style-name")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:variable-input" => DynamicTextField::VariableInput {
                name: required_field_attribute(self, "text:name")?.to_owned(),
                description: self
                    .element
                    .get_attribute("text:description")
                    .map(str::to_owned),
                value_type: parse_value_type_only(self)?,
                display: self
                    .element
                    .get_attribute("text:display")
                    .map(VariableSetDisplay::parse)
                    .transpose()?,
                data_style_name: self
                    .element
                    .get_attribute("style:data-style-name")
                    .map(str::to_owned),
                display_text: text(),
            },
            "text:user-field-get" => {
                reject_calculated_value_attributes(self)?;
                DynamicTextField::UserFieldGet {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(UserFieldDisplay::parse)
                        .transpose()?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:user-field-input" => {
                reject_calculated_value_attributes(self)?;
                DynamicTextField::UserFieldInput {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    description: self
                        .element
                        .get_attribute("text:description")
                        .map(str::to_owned),
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:text-input" => {
                reject_calculated_value_attributes(self)?;
                DynamicTextField::TextInput {
                    description: self
                        .element
                        .get_attribute("text:description")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:table-formula" => {
                reject_calculated_value_attributes(self)?;
                DynamicTextField::TableFormula {
                    formula: self
                        .element
                        .get_attribute("text:formula")
                        .map(str::to_owned),
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(FormulaFieldDisplay::parse)
                        .transpose()?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:measure" => {
                reject_unknown_field_attributes(self, &["text:kind"])?;
                DynamicTextField::Measure {
                    kind: MeasureKind::parse(required_field_attribute(self, "text:kind")?)?,
                    display_text: text(),
                }
            },
            "text:reference-ref" => {
                reject_unknown_field_attributes(self, &["text:ref-name", "text:reference-format"])?;
                DynamicTextField::Reference {
                    reference_name: self
                        .element
                        .get_attribute("text:ref-name")
                        .map(str::to_owned),
                    reference_format: self
                        .element
                        .get_attribute("text:reference-format")
                        .map(CrossReferenceFormat::parse)
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:bookmark-ref" => {
                reject_unknown_field_attributes(self, &["text:ref-name", "text:reference-format"])?;
                DynamicTextField::BookmarkReference {
                    reference_name: self
                        .element
                        .get_attribute("text:ref-name")
                        .map(str::to_owned),
                    reference_format: self
                        .element
                        .get_attribute("text:reference-format")
                        .map(CrossReferenceFormat::parse)
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:note-ref" => {
                reject_unknown_field_attributes(
                    self,
                    &["text:ref-name", "text:reference-format", "text:note-class"],
                )?;
                DynamicTextField::NoteReference {
                    reference_name: self
                        .element
                        .get_attribute("text:ref-name")
                        .map(str::to_owned),
                    note_class: NoteReferenceClass::parse(required_field_attribute(
                        self,
                        "text:note-class",
                    )?)?,
                    reference_format: self
                        .element
                        .get_attribute("text:reference-format")
                        .map(NoteReferenceFormat::parse)
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:page-count"
            | "text:paragraph-count"
            | "text:word-count"
            | "text:character-count"
            | "text:table-count"
            | "text:image-count"
            | "text:object-count" => {
                reject_unknown_field_attributes(
                    self,
                    &["style:num-format", "style:num-letter-sync"],
                )?;
                let kind = match self.field_type() {
                    "text:page-count" => StatisticKind::Page,
                    "text:paragraph-count" => StatisticKind::Paragraph,
                    "text:word-count" => StatisticKind::Word,
                    "text:character-count" => StatisticKind::Character,
                    "text:table-count" => StatisticKind::Table,
                    "text:image-count" => StatisticKind::Image,
                    "text:object-count" => StatisticKind::Object,
                    _ => unreachable!(),
                };
                DynamicTextField::DocumentStatistic {
                    kind,
                    number_format: parse_common_number_format(self)?,
                    display_text: text(),
                }
            },
            "text:page-number" => {
                reject_unknown_field_attributes(
                    self,
                    &[
                        "style:num-format",
                        "style:num-letter-sync",
                        "text:fixed",
                        "text:page-adjust",
                        "text:select-page",
                    ],
                )?;
                DynamicTextField::PageNumber {
                    number_format: parse_common_number_format(self)?,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    page_adjust: self
                        .element
                        .get_attribute("text:page-adjust")
                        .map(|value| {
                            value.parse::<i64>().map_err(|_| {
                                Error::InvalidFormat(format!(
                                    "invalid text:page-adjust integer '{value}'"
                                ))
                            })
                        })
                        .transpose()?,
                    select_page: self
                        .element
                        .get_attribute("text:select-page")
                        .map(PageSelection::parse)
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:date" => {
                reject_unknown_field_attributes(
                    self,
                    &[
                        "text:date-value",
                        "text:date-adjust",
                        "text:fixed",
                        "style:data-style-name",
                    ],
                )?;
                DynamicTextField::Date {
                    value: self
                        .element
                        .get_attribute("text:date-value")
                        .map(FieldDateValue::new)
                        .transpose()?,
                    adjustment: self
                        .element
                        .get_attribute("text:date-adjust")
                        .map(FieldDuration::new)
                        .transpose()?,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:time" => {
                reject_unknown_field_attributes(
                    self,
                    &[
                        "text:time-value",
                        "text:time-adjust",
                        "text:fixed",
                        "style:data-style-name",
                    ],
                )?;
                DynamicTextField::Time {
                    value: self
                        .element
                        .get_attribute("text:time-value")
                        .map(FieldTimeValue::new)
                        .transpose()?,
                    adjustment: self
                        .element
                        .get_attribute("text:time-adjust")
                        .map(FieldDuration::new)
                        .transpose()?,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:page-continuation" => {
                reject_unknown_field_attributes(self, &["text:select-page", "text:string-value"])?;
                DynamicTextField::PageContinuation {
                    select_page: PageContinuationSelection::parse(required_field_attribute(
                        self,
                        "text:select-page",
                    )?)?,
                    string_value: self
                        .element
                        .get_attribute("text:string-value")
                        .map(str::to_owned),
                    display_text: text(),
                }
            },
            "text:page-variable-set" => {
                reject_unknown_field_attributes(self, &["text:active", "text:page-adjust"])?;
                DynamicTextField::PageVariableSet {
                    active: optional_field_bool(self, "text:active")?,
                    page_adjust: self
                        .element
                        .get_attribute("text:page-adjust")
                        .map(|value| {
                            value.parse::<i64>().map_err(|_| {
                                Error::InvalidFormat(format!(
                                    "invalid page-variable text:page-adjust integer '{value}'"
                                ))
                            })
                        })
                        .transpose()?,
                    display_text: text(),
                }
            },
            "text:page-variable-get" => {
                reject_unknown_field_attributes(
                    self,
                    &["style:num-format", "style:num-letter-sync"],
                )?;
                DynamicTextField::PageVariableGet {
                    number_format: parse_common_number_format(self)?,
                    display_text: text(),
                }
            },
            "text:file-name" => {
                reject_unknown_field_attributes(self, &["text:display", "text:fixed"])?;
                let result = DynamicTextField::FileName {
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(FileNameDisplay::parse)
                        .transpose()?,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            "text:template-name" => {
                reject_unknown_field_attributes(self, &["text:display"])?;
                let result = DynamicTextField::TemplateName {
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(TemplateNameDisplay::parse)
                        .transpose()?,
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            "text:sheet-name" => {
                reject_unknown_field_attributes(self, &[])?;
                let result = DynamicTextField::SheetName {
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            "text:chapter" => {
                reject_unknown_field_attributes(self, &["text:display", "text:outline-level"])?;
                let result = DynamicTextField::Chapter {
                    display: self
                        .element
                        .get_attribute("text:display")
                        .map(ChapterDisplay::parse)
                        .transpose()?,
                    outline_level: self
                        .element
                        .get_attribute("text:outline-level")
                        .map(NonNegativeInteger::new)
                        .transpose()?,
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            "text:creation-date"
            | "text:creation-time"
            | "text:print-date"
            | "text:print-time"
            | "text:editing-cycles"
            | "text:editing-duration"
            | "text:modification-date"
            | "text:modification-time" => {
                let kind = match self.field_type() {
                    "text:creation-date" => MetadataFieldKind::CreationDate,
                    "text:creation-time" => MetadataFieldKind::CreationTime,
                    "text:print-date" => MetadataFieldKind::PrintDate,
                    "text:print-time" => MetadataFieldKind::PrintTime,
                    "text:editing-cycles" => MetadataFieldKind::EditingCycles,
                    "text:editing-duration" => MetadataFieldKind::EditingDuration,
                    "text:modification-date" => MetadataFieldKind::ModificationDate,
                    "text:modification-time" => MetadataFieldKind::ModificationTime,
                    _ => unreachable!(),
                };
                let allowed = match kind {
                    MetadataFieldKind::CreationDate
                    | MetadataFieldKind::PrintDate
                    | MetadataFieldKind::ModificationDate => {
                        &["text:fixed", "style:data-style-name", "text:date-value"][..]
                    },
                    MetadataFieldKind::CreationTime
                    | MetadataFieldKind::PrintTime
                    | MetadataFieldKind::ModificationTime => {
                        &["text:fixed", "style:data-style-name", "text:time-value"][..]
                    },
                    MetadataFieldKind::EditingDuration => {
                        &["text:fixed", "style:data-style-name", "text:duration"][..]
                    },
                    MetadataFieldKind::EditingCycles => &["text:fixed"][..],
                };
                reject_unknown_field_attributes(self, allowed)?;
                let value = match kind {
                    MetadataFieldKind::CreationDate
                    | MetadataFieldKind::PrintDate
                    | MetadataFieldKind::ModificationDate => self
                        .element
                        .get_attribute("text:date-value")
                        .map(FieldDateValue::new)
                        .transpose()?
                        .map(MetadataFieldValue::Date),
                    MetadataFieldKind::CreationTime
                    | MetadataFieldKind::PrintTime
                    | MetadataFieldKind::ModificationTime => self
                        .element
                        .get_attribute("text:time-value")
                        .map(FieldTimeValue::new)
                        .transpose()?
                        .map(MetadataFieldValue::Time),
                    MetadataFieldKind::EditingDuration => self
                        .element
                        .get_attribute("text:duration")
                        .map(FieldDuration::new)
                        .transpose()?
                        .map(MetadataFieldValue::Duration),
                    MetadataFieldKind::EditingCycles => None,
                };
                let result = DynamicTextField::DocumentMetadata {
                    kind,
                    value,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            "text:initial-creator"
            | "text:description"
            | "text:printed-by"
            | "text:title"
            | "text:subject"
            | "text:keywords"
            | "text:creator"
            | "text:author-name"
            | "text:author-initials" => {
                reject_unknown_field_attributes(self, &["text:fixed"])?;
                let kind = match self.field_type() {
                    "text:initial-creator" => IdentityFieldKind::InitialCreator,
                    "text:description" => IdentityFieldKind::Description,
                    "text:printed-by" => IdentityFieldKind::PrintedBy,
                    "text:title" => IdentityFieldKind::Title,
                    "text:subject" => IdentityFieldKind::Subject,
                    "text:keywords" => IdentityFieldKind::Keywords,
                    "text:creator" => IdentityFieldKind::Creator,
                    "text:author-name" => IdentityFieldKind::AuthorName,
                    "text:author-initials" => IdentityFieldKind::AuthorInitials,
                    _ => unreachable!(),
                };
                DynamicTextField::DocumentIdentity {
                    kind,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    display_text: text(),
                }
            },
            "text:sender-firstname"
            | "text:sender-lastname"
            | "text:sender-initials"
            | "text:sender-title"
            | "text:sender-position"
            | "text:sender-email"
            | "text:sender-phone-private"
            | "text:sender-fax"
            | "text:sender-company"
            | "text:sender-phone-work"
            | "text:sender-street"
            | "text:sender-city"
            | "text:sender-postal-code"
            | "text:sender-country"
            | "text:sender-state-or-province" => {
                reject_unknown_field_attributes(self, &["text:fixed"])?;
                let kind = match self.field_type() {
                    "text:sender-firstname" => SenderFieldKind::FirstName,
                    "text:sender-lastname" => SenderFieldKind::LastName,
                    "text:sender-initials" => SenderFieldKind::Initials,
                    "text:sender-title" => SenderFieldKind::Title,
                    "text:sender-position" => SenderFieldKind::Position,
                    "text:sender-email" => SenderFieldKind::Email,
                    "text:sender-phone-private" => SenderFieldKind::PrivatePhone,
                    "text:sender-fax" => SenderFieldKind::Fax,
                    "text:sender-company" => SenderFieldKind::Company,
                    "text:sender-phone-work" => SenderFieldKind::WorkPhone,
                    "text:sender-street" => SenderFieldKind::Street,
                    "text:sender-city" => SenderFieldKind::City,
                    "text:sender-postal-code" => SenderFieldKind::PostalCode,
                    "text:sender-country" => SenderFieldKind::Country,
                    "text:sender-state-or-province" => SenderFieldKind::StateOrProvince,
                    _ => unreachable!(),
                };
                let result = DynamicTextField::Sender {
                    kind,
                    fixed: optional_field_bool(self, "text:fixed")?,
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            "text:user-defined" => {
                reject_unknown_field_attributes(
                    self,
                    &[
                        "text:name",
                        "text:fixed",
                        "style:data-style-name",
                        "office:value",
                        "office:date-value",
                        "office:time-value",
                        "office:boolean-value",
                        "office:string-value",
                    ],
                )?;
                let result = DynamicTextField::UserDefinedMetadata {
                    name: required_field_attribute(self, "text:name")?.to_owned(),
                    values: UserDefinedMetadataValues {
                        number: self
                            .element
                            .get_attribute("office:value")
                            .map(str::to_owned),
                        date: self
                            .element
                            .get_attribute("office:date-value")
                            .map(FieldDateValue::new)
                            .transpose()?,
                        time: self
                            .element
                            .get_attribute("office:time-value")
                            .map(FieldDuration::new)
                            .transpose()?,
                        boolean: optional_field_bool(self, "office:boolean-value")?,
                        string: self
                            .element
                            .get_attribute("office:string-value")
                            .map(str::to_owned),
                    },
                    fixed: optional_field_bool(self, "text:fixed")?,
                    data_style_name: self
                        .element
                        .get_attribute("style:data-style-name")
                        .map(str::to_owned),
                    display_text: text(),
                };
                result.validate()?;
                result
            },
            _ => return Ok(None),
        };
        Ok(Some(result))
    }
}

pub(crate) fn validate_xml_schema_date(value: &str) -> Result<()> {
    let core = strip_xml_schema_timezone(value)?;
    validate_xml_schema_date_core(core)
        .map_err(|_| Error::InvalidFormat(format!("invalid XML Schema date '{value}'")))
}

pub(crate) fn validate_xml_schema_date_time(value: &str) -> Result<()> {
    let (date, time) = value
        .split_once('T')
        .ok_or_else(|| Error::InvalidFormat(format!("invalid XML Schema dateTime '{value}'")))?;
    if time.contains('T')
        || validate_xml_schema_date_core(date).is_err()
        || validate_xml_schema_time(time).is_err()
    {
        return Err(Error::InvalidFormat(format!(
            "invalid XML Schema dateTime '{value}'"
        )));
    }
    Ok(())
}

pub(crate) fn validate_xml_schema_time(value: &str) -> Result<()> {
    let core = strip_xml_schema_timezone(value)?;
    let mut parts = core.split(':');
    let hour = parse_two_digits(parts.next(), "hour")?;
    let minute = parse_two_digits(parts.next(), "minute")?;
    let second_lexical = parts
        .next()
        .ok_or_else(|| Error::InvalidFormat(format!("invalid XML Schema time '{value}'")))?;
    if parts.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "invalid XML Schema time '{value}'"
        )));
    }
    let (seconds, fraction) = match second_lexical.split_once('.') {
        Some((seconds, fraction))
            if !fraction.is_empty() && fraction.bytes().all(|b| b.is_ascii_digit()) =>
        {
            (seconds, Some(fraction))
        },
        Some(_) => {
            return Err(Error::InvalidFormat(format!(
                "invalid XML Schema time '{value}'"
            )));
        },
        None => (second_lexical, None),
    };
    let second = if seconds.len() == 2 && seconds.bytes().all(|b| b.is_ascii_digit()) {
        seconds.parse::<u8>().unwrap_or(u8::MAX)
    } else {
        u8::MAX
    };
    let midnight_24 = hour == 24
        && minute == 0
        && second == 0
        && fraction.is_none_or(|value| value.bytes().all(|b| b == b'0'));
    if minute > 59 || second > 59 || (hour > 23 && !midnight_24) {
        return Err(Error::InvalidFormat(format!(
            "invalid XML Schema time '{value}'"
        )));
    }
    Ok(())
}

fn validate_xml_schema_date_core(value: &str) -> Result<()> {
    let unsigned = value.strip_prefix('-').unwrap_or(value);
    let mut parts = unsigned.split('-');
    let year = parts.next().unwrap_or_default();
    let month = parse_two_digits(parts.next(), "month")?;
    let day = parse_two_digits(parts.next(), "day")?;
    if parts.next().is_some()
        || year.len() < 4
        || !year.bytes().all(|b| b.is_ascii_digit())
        || year.bytes().all(|b| b == b'0')
    {
        return Err(Error::InvalidFormat("invalid XML Schema date".to_string()));
    }
    let leap =
        decimal_mod(year, 400) == 0 || (decimal_mod(year, 4) == 0 && decimal_mod(year, 100) != 0);
    let max_day = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => 0,
    };
    if day == 0 || day > max_day {
        return Err(Error::InvalidFormat("invalid XML Schema date".to_string()));
    }
    Ok(())
}

fn strip_xml_schema_timezone(value: &str) -> Result<&str> {
    if let Some(core) = value.strip_suffix('Z') {
        return if core.is_empty() {
            Err(Error::InvalidFormat(
                "empty XML Schema temporal value".to_string(),
            ))
        } else {
            Ok(core)
        };
    }
    let bytes = value.as_bytes();
    if bytes.len() >= 6
        && matches!(bytes[bytes.len() - 6], b'+' | b'-')
        && bytes[bytes.len() - 3] == b':'
    {
        let timezone = &value[value.len() - 5..];
        let hour = parse_two_digits(Some(&timezone[..2]), "timezone hour")?;
        let minute = parse_two_digits(Some(&timezone[3..]), "timezone minute")?;
        if hour > 14 || minute > 59 || (hour == 14 && minute != 0) {
            return Err(Error::InvalidFormat(format!(
                "invalid XML Schema timezone in '{value}'"
            )));
        }
        return Ok(&value[..value.len() - 6]);
    }
    Ok(value)
}

fn parse_two_digits(value: Option<&str>, component: &str) -> Result<u8> {
    let value = value
        .ok_or_else(|| Error::InvalidFormat(format!("missing XML Schema temporal {component}")))?;
    if value.len() != 2 || !value.bytes().all(|b| b.is_ascii_digit()) {
        return Err(Error::InvalidFormat(format!(
            "invalid XML Schema temporal {component}"
        )));
    }
    value
        .parse::<u8>()
        .map_err(|_| Error::InvalidFormat(format!("invalid XML Schema temporal {component}")))
}

fn decimal_mod(value: &str, modulus: u16) -> u16 {
    value.bytes().fold(0u16, |remainder, digit| {
        (remainder * 10 + u16::from(digit - b'0')) % modulus
    })
}

const CALCULATED_VALUE_ATTRIBUTES: [&str; 7] = [
    "office:value",
    "office:currency",
    "office:date-value",
    "office:time-value",
    "office:boolean-value",
    "office:string-value",
    "office:value-type",
];

fn reject_calculated_value_attributes(field: &Field) -> Result<()> {
    if let Some(name) = CALCULATED_VALUE_ATTRIBUTES
        .iter()
        .find(|name| field.element.get_attribute(name).is_some())
    {
        return Err(Error::InvalidFormat(format!(
            "{} does not permit {name}",
            field.field_type()
        )));
    }
    Ok(())
}

fn parse_calculated_value(field: &Field, required: bool) -> Result<Option<CalculatedFieldValue>> {
    let Some(value_type) = field.element.get_attribute("office:value-type") else {
        if CALCULATED_VALUE_ATTRIBUTES[..6]
            .iter()
            .any(|name| field.element.get_attribute(name).is_some())
        {
            return Err(Error::InvalidFormat(
                "cached field value attributes require office:value-type".to_string(),
            ));
        }
        return if required {
            Err(Error::InvalidFormat(format!(
                "{} requires office:value-type and its matching value",
                field.field_type()
            )))
        } else {
            Ok(None)
        };
    };
    let attr = |name| field.element.get_attribute(name);
    let required_attr = |name| {
        attr(name).ok_or_else(|| {
            Error::InvalidFormat(format!("office:value-type '{value_type}' requires {name}"))
        })
    };
    let value = match value_type {
        "float" => CalculatedFieldValue::Float(required_attr("office:value")?.to_owned()),
        "percentage" => CalculatedFieldValue::Percentage(required_attr("office:value")?.to_owned()),
        "currency" => CalculatedFieldValue::Currency {
            value: required_attr("office:value")?.to_owned(),
            currency: attr("office:currency").map(str::to_owned),
        },
        "date" => CalculatedFieldValue::Date(required_attr("office:date-value")?.to_owned()),
        "time" => CalculatedFieldValue::Time(required_attr("office:time-value")?.to_owned()),
        "boolean" => CalculatedFieldValue::Boolean(
            optional_field_bool(field, "office:boolean-value")?.ok_or_else(|| {
                Error::InvalidFormat(
                    "office:value-type 'boolean' requires office:boolean-value".to_string(),
                )
            })?,
        ),
        "string" => CalculatedFieldValue::String(attr("office:string-value").map(str::to_owned)),
        _ => {
            return Err(Error::InvalidFormat(format!(
                "invalid calculated field office:value-type '{value_type}'"
            )));
        },
    };
    let allowed: &[&str] = match value_type {
        "float" | "percentage" => &["office:value-type", "office:value"],
        "currency" => &["office:value-type", "office:value", "office:currency"],
        "date" => &["office:value-type", "office:date-value"],
        "time" => &["office:value-type", "office:time-value"],
        "boolean" => &["office:value-type", "office:boolean-value"],
        "string" => &["office:value-type", "office:string-value"],
        _ => unreachable!(),
    };
    if let Some(extra) = CALCULATED_VALUE_ATTRIBUTES
        .iter()
        .find(|name| !allowed.contains(name) && attr(name).is_some())
    {
        return Err(Error::InvalidFormat(format!(
            "office:value-type '{value_type}' does not permit {extra}"
        )));
    }
    let mut aggregate = 0usize;
    value.validate(&mut aggregate)?;
    Ok(Some(value))
}

fn parse_value_type_only(field: &Field) -> Result<FieldValueType> {
    let value_type = FieldValueType::parse(required_field_attribute(field, "office:value-type")?)?;
    if let Some(extra) = CALCULATED_VALUE_ATTRIBUTES[..6]
        .iter()
        .find(|name| field.element.get_attribute(name).is_some())
    {
        return Err(Error::InvalidFormat(format!(
            "text:variable-input permits office:value-type but not {extra}"
        )));
    }
    Ok(value_type)
}

fn parse_common_number_format(field: &Field) -> Result<Option<SequenceNumberFormat>> {
    let format = field.element.get_attribute("style:num-format");
    let letter_sync = optional_field_bool(field, "style:num-letter-sync")?;
    match (format, letter_sync) {
        (Some(format), letter_sync) => Ok(Some(SequenceNumberFormat::new(format, letter_sync)?)),
        (None, Some(_)) => Err(Error::InvalidFormat(
            "style:num-letter-sync requires style:num-format".to_string(),
        )),
        (None, None) => Ok(None),
    }
}

fn reject_unknown_field_attributes(field: &Field, allowed: &[&str]) -> Result<()> {
    if let Some(name) = field
        .element
        .attributes()
        .keys()
        .find(|name| !allowed.contains(&name.as_str()))
    {
        return Err(Error::InvalidFormat(format!(
            "{} does not permit attribute {name}",
            field.field_type()
        )));
    }
    Ok(())
}

fn required_field_attribute<'a>(field: &'a Field, name: &str) -> Result<&'a str> {
    field
        .element
        .get_attribute(name)
        .ok_or_else(|| Error::InvalidFormat(format!("{} requires {name}", field.field_type())))
}

fn optional_field_bool(field: &Field, name: &str) -> Result<Option<bool>> {
    field
        .element
        .get_attribute(name)
        .map(|value| match value {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid {name} boolean '{value}'"
            ))),
        })
        .transpose()
}

/// Represents a page number field
#[derive(Debug, Clone)]
#[allow(dead_code)] // Library API for document creation
pub struct PageNumberField {
    element: Element,
}

#[allow(dead_code)] // Library API for document creation
impl PageNumberField {
    /// Create a new page number field
    pub fn new() -> Self {
        Self {
            element: Element::new("text:page-number"),
        }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:page-number" {
            return Err(Error::InvalidFormat(
                "Element is not a page number field".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the current page number value
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }
}

impl Default for PageNumberField {
    fn default() -> Self {
        Self::new()
    }
}

/// Represents a date field
#[derive(Debug, Clone)]
#[allow(dead_code)] // Library API for document creation
pub struct DateField {
    element: Element,
}

#[allow(dead_code)] // Library API for document creation
impl DateField {
    /// Create a new date field
    pub fn new() -> Self {
        Self {
            element: Element::new("text:date"),
        }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        if element.tag_name() != "text:date" {
            return Err(Error::InvalidFormat(
                "Element is not a date field".to_string(),
            ));
        }
        Ok(Self { element })
    }

    /// Get the date value
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }

    /// Get the fixed date (if any)
    pub fn fixed_date(&self) -> Option<&str> {
        self.element.get_attribute("text:date-value")
    }

    /// Get whether this date is fixed
    pub fn is_fixed(&self) -> bool {
        self.element
            .get_bool_attribute("text:fixed")
            .unwrap_or(false)
    }
}

impl Default for DateField {
    fn default() -> Self {
        Self::new()
    }
}

/// Represents a reference field
#[derive(Debug, Clone)]
#[allow(dead_code)] // Library API for document creation
pub struct ReferenceField {
    element: Element,
}

#[allow(dead_code)] // Library API for document creation
impl ReferenceField {
    /// Create a new reference field
    pub fn new(ref_name: &str) -> Self {
        let mut element = Element::new("text:reference-ref");
        element.set_attribute("text:ref-name", ref_name);
        Self { element }
    }

    /// Create from element
    pub fn from_element(element: Element) -> Result<Self> {
        let tag = element.tag_name();
        if !matches!(
            tag,
            "text:reference-ref" | "text:bookmark-ref" | "text:sequence-ref"
        ) {
            return Err(Error::InvalidFormat(format!(
                "Element {} is not a reference field",
                tag
            )));
        }
        Ok(Self { element })
    }

    /// Get the reference name
    pub fn ref_name(&self) -> Option<&str> {
        self.element.get_attribute("text:ref-name")
    }

    /// Get the reference format
    pub fn ref_format(&self) -> Option<&str> {
        self.element.get_attribute("text:reference-format")
    }

    /// Get the reference value
    pub fn value(&self) -> String {
        self.element.get_text_recursive()
    }
}
