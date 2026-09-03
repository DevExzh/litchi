#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::struct_excessive_bools,
    reason = "the public model preserves independent OOXML flags"
)]
use super::{
    Appearance, BindingFlavor, Calendar, DataBinding, FORMATTING_ALLOWED_NAMESPACE,
    FormattingAllowed, Kind, Limits, Lock, STORE_ITEM_CHECKSUM_NAMESPACE, SdtColor,
    validate_data_binding_values,
};
use crate::error::{Error, Result};
use litchi_core::xml::escape_xml;
use litchi_ooxml_common::properties::time::DateTime;
use std::fmt::Write as _;

const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const WORD_2010_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const WORD_2012_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const MAX_SDT_ID: u32 = i32::MAX as u32;

/// Borrowed semantic properties used to author one w:sdtPr.
#[derive(Debug, Clone, Copy)]
pub struct AuthoringView<'a> {
    id: u32,
    kind: Kind,
    lock: Lock,
    tag: Option<&'a str>,
    title: Option<&'a str>,
    temporary: bool,
    showing_placeholder: bool,
    placeholder_doc_part: Option<&'a str>,
    tab_index: Option<u32>,
    data_binding: Option<&'a DataBinding>,
    list_items: &'a [(String, String)],
    checked: Option<bool>,
    date_format: Option<&'a str>,
    date_value: Option<&'a str>,
    date_calendar: Option<&'a Calendar>,
    repeating_section_title: Option<&'a str>,
    formatting_allowed: Option<FormattingAllowed>,
    appearance: Option<Appearance>,
    color: Option<SdtColor>,
    web_extension_linked: Option<bool>,
    web_extension_created: Option<bool>,
}

impl<'a> AuthoringView<'a> {
    /// Start a canonical property view with no optional properties.
    #[must_use]
    pub const fn new(id: u32, kind: Kind, lock: Lock) -> Self {
        Self {
            id,
            kind,
            lock,
            tag: None,
            title: None,
            temporary: false,
            showing_placeholder: false,
            placeholder_doc_part: None,
            tab_index: None,
            data_binding: None,
            list_items: &[],
            checked: None,
            date_format: None,
            date_value: None,
            date_calendar: None,
            repeating_section_title: None,
            formatting_allowed: None,
            appearance: None,
            color: None,
            web_extension_linked: None,
            web_extension_created: None,
        }
    }

    #[must_use]
    pub const fn tag(mut self, value: Option<&'a str>) -> Self {
        self.tag = value;
        self
    }

    #[must_use]
    pub const fn title(mut self, value: Option<&'a str>) -> Self {
        self.title = value;
        self
    }

    #[must_use]
    pub const fn temporary(mut self, value: bool) -> Self {
        self.temporary = value;
        self
    }

    #[must_use]
    pub const fn showing_placeholder(mut self, value: bool) -> Self {
        self.showing_placeholder = value;
        self
    }

    #[must_use]
    pub const fn placeholder_doc_part(mut self, value: Option<&'a str>) -> Self {
        self.placeholder_doc_part = value;
        self
    }

    #[must_use]
    pub const fn tab_index(mut self, value: Option<u32>) -> Self {
        self.tab_index = value;
        self
    }

    #[must_use]
    pub const fn data_binding(mut self, value: Option<&'a DataBinding>) -> Self {
        self.data_binding = value;
        self
    }

    #[must_use]
    pub const fn list_items(mut self, value: &'a [(String, String)]) -> Self {
        self.list_items = value;
        self
    }

    #[must_use]
    pub const fn checked(mut self, value: Option<bool>) -> Self {
        self.checked = value;
        self
    }

    #[must_use]
    pub const fn date_format(mut self, value: Option<&'a str>) -> Self {
        self.date_format = value;
        self
    }

    #[must_use]
    pub const fn date_value(mut self, value: Option<&'a str>) -> Self {
        self.date_value = value;
        self
    }

    /// Set the calendar used by a date content control.
    #[must_use]
    pub const fn date_calendar(mut self, value: Option<&'a Calendar>) -> Self {
        self.date_calendar = value;
        self
    }

    #[must_use]
    pub const fn repeating_section_title(mut self, value: Option<&'a str>) -> Self {
        self.repeating_section_title = value;
        self
    }

    #[must_use]
    pub const fn formatting_allowed(mut self, value: Option<FormattingAllowed>) -> Self {
        self.formatting_allowed = value;
        self
    }

    /// Set Word 2012 visual treatment metadata.
    #[must_use]
    pub const fn appearance(mut self, value: Option<Appearance>) -> Self {
        self.appearance = value;
        self
    }

    /// Set Word 2012 visual base-color metadata.
    #[must_use]
    pub const fn color(mut self, value: Option<SdtColor>) -> Self {
        self.color = value;
        self
    }

    /// Set the inert Word 2012 web-extension linked marker.
    #[must_use]
    pub const fn web_extension_linked(mut self, value: Option<bool>) -> Self {
        self.web_extension_linked = value;
        self
    }

    /// Set the inert Word 2012 web-extension-created marker.
    #[must_use]
    pub const fn web_extension_created(mut self, value: Option<bool>) -> Self {
        self.web_extension_created = value;
        self
    }
}

/// Extension namespaces required by authored content-control properties.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct NamespaceRequirements {
    word_2010: bool,
    word_2012: bool,
    checksum: bool,
    formatting_allowed: bool,
}

impl NamespaceRequirements {
    #[must_use]
    pub const fn word_2010(self) -> bool {
        self.word_2010
    }

    #[must_use]
    pub const fn word_2012(self) -> bool {
        self.word_2012
    }

    #[must_use]
    pub const fn checksum(self) -> bool {
        self.checksum
    }

    #[must_use]
    pub const fn formatting_allowed(self) -> bool {
        self.formatting_allowed
    }

    #[must_use]
    pub const fn needs_mce(self) -> bool {
        self.word_2010 || self.word_2012 || self.checksum || self.formatting_allowed
    }
}

/// Canonical content-control properties and their namespace requirements.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuthoredProperties {
    xml: String,
    namespace_requirements: NamespaceRequirements,
}

impl AuthoredProperties {
    #[must_use]
    pub fn xml(&self) -> &str {
        &self.xml
    }

    #[must_use]
    pub fn into_xml(self) -> String {
        self.xml
    }

    #[must_use]
    pub const fn namespace_requirements(&self) -> NamespaceRequirements {
        self.namespace_requirements
    }
}

/// Author one complete, canonical w:sdtPr element.
///
/// # Errors
///
/// Returns an error if the operation cannot be completed.
pub fn write_sdt_pr(view: &AuthoringView<'_>, limits: &Limits) -> Result<AuthoredProperties> {
    limits.validate()?;
    if view.id > MAX_SDT_ID {
        return Err(Error::Invalid(
            "content control id exceeds the Int32Value maximum".into(),
        ));
    }
    validate_view(view, limits)?;
    let requirements = requirements(view);
    let metadata_bytes = metadata_bytes(view)?;
    if metadata_bytes > limits.max_metadata_bytes {
        return Err(metadata_limit());
    }
    let estimate = output_estimate(view, metadata_bytes)?;
    if estimate > limits.max_output_bytes {
        return Err(output_limit());
    }

    let mut xml = String::new();
    xml.try_reserve(estimate)
        .map_err(|_source_error| output_limit())?;
    xml.push_str("<w:sdtPr");
    if requirements.needs_mce() {
        write!(&mut xml, r#" xmlns:mc="{MC_NAMESPACE}""#)?;
    }
    if requirements.word_2010 {
        write!(&mut xml, r#" xmlns:w14="{WORD_2010_NAMESPACE}""#)?;
    }
    if requirements.word_2012 {
        write!(&mut xml, r#" xmlns:w15="{WORD_2012_NAMESPACE}""#)?;
    }
    if requirements.checksum {
        write!(
            &mut xml,
            r#" xmlns:w16sdtdh="{STORE_ITEM_CHECKSUM_NAMESPACE}""#
        )?;
    }
    if requirements.formatting_allowed {
        write!(
            &mut xml,
            r#" xmlns:w16sdtfl="{FORMATTING_ALLOWED_NAMESPACE}""#
        )?;
    }
    if requirements.needs_mce() {
        xml.push_str(r#" mc:Ignorable=""#);
        let mut separator = "";
        for (needed, prefix) in [
            (requirements.word_2010, "w14"),
            (requirements.word_2012, "w15"),
            (requirements.checksum, "w16sdtdh"),
            (requirements.formatting_allowed, "w16sdtfl"),
        ] {
            if needed {
                xml.push_str(separator);
                xml.push_str(prefix);
                separator = " ";
            }
        }
        xml.push('"');
    }
    xml.push('>');

    write_optional_value(&mut xml, "alias", view.title)?;
    write_lock(&mut xml, view.lock, view.formatting_allowed)?;
    if let Some(placeholder) = view.placeholder_doc_part {
        write!(
            &mut xml,
            r#"<w:placeholder><w:docPart w:val="{}"/></w:placeholder>"#,
            escape_xml(placeholder)
        )?;
    }
    if view.showing_placeholder {
        xml.push_str("<w:showingPlcHdr/>");
    }
    write_binding(&mut xml, view.data_binding)?;
    if view.temporary {
        xml.push_str("<w:temporary/>");
    }
    write!(&mut xml, r#"<w:id w:val="{}"/>"#, view.id)?;
    write_optional_value(&mut xml, "tag", view.tag)?;
    if let Some(tab_index) = view.tab_index {
        write!(&mut xml, r#"<w:tabIndex w:val="{tab_index}"/>"#)?;
    }
    write_kind(&mut xml, view)?;
    xml.push_str("</w:sdtPr>");

    if xml.len() > limits.max_output_bytes {
        return Err(output_limit());
    }
    Ok(AuthoredProperties {
        xml,
        namespace_requirements: requirements,
    })
}

#[cfg(test)]
mod id_range_tests {
    use super::*;

    #[test]
    fn writes_the_maximum_int32_content_control_id() {
        let view = AuthoringView::new(i32::MAX as u32, Kind::RichText, Lock::Unlocked);

        let authored = write_sdt_pr(&view, &Limits::default()).unwrap();

        assert!(authored.xml().contains(r#"<w:id w:val="2147483647"/>"#));
    }

    #[test]
    fn rejects_content_control_id_above_the_int32_range() {
        let view = AuthoringView::new(i32::MAX as u32 + 1, Kind::RichText, Lock::Unlocked);

        let error = write_sdt_pr(&view, &Limits::default()).unwrap_err();

        assert!(matches!(error, Error::Invalid(message) if message.contains("Int32Value")));
    }
}

fn validate_view(view: &AuthoringView<'_>, limits: &Limits) -> Result<()> {
    if view.formatting_allowed.is_some() && !view.lock.locks_content() {
        return Err(Error::InvalidFormat(
            "formattingAllowed requires contentLocked or sdtContentLocked".into(),
        ));
    }
    if !view.list_items.is_empty() && !matches!(view.kind, Kind::ComboBox | Kind::Dropdown) {
        return Err(Error::InvalidFormat(
            "list items require a combo-box or drop-down content control".into(),
        ));
    }
    if view.checked.is_some() && view.kind != Kind::Checkbox {
        return Err(Error::InvalidFormat(
            "checked state requires a checkbox content control".into(),
        ));
    }
    if (view.date_format.is_some() || view.date_value.is_some() || view.date_calendar.is_some())
        && view.kind != Kind::Date
    {
        return Err(Error::InvalidFormat(
            "date properties require a date content control".into(),
        ));
    }
    if view.repeating_section_title.is_some() && view.kind != Kind::RepeatingSection {
        return Err(Error::InvalidFormat(
            "section title requires a repeating-section content control".into(),
        ));
    }
    if view.kind == Kind::EntityPicker {
        return Err(Error::InvalidFormat(
            "EntityPicker authoring requires package-owned Custom XML metadata".into(),
        ));
    }
    if view.list_items.len() > limits.max_list_items
        || view.list_items.len() > limits.max_list_items_per_control
    {
        return Err(Error::InvalidFormat(
            "content-control list-item count exceeds configured limit".into(),
        ));
    }
    for (label, value) in [
        ("content-control tag", view.tag),
        ("content-control alias", view.title),
        (
            "placeholder document-part reference",
            view.placeholder_doc_part,
        ),
        ("date format", view.date_format),
        ("date fullDate", view.date_value),
        ("date calendar", view.date_calendar.map(Calendar::as_str)),
        ("repeating-section title", view.repeating_section_title),
    ] {
        if let Some(value) = value {
            validate_xml_scalars(value, label)?;
        }
    }
    for (display, value) in view.list_items {
        validate_xml_scalars(display, "list-item display text")?;
        validate_xml_scalars(value, "list-item value")?;
    }
    if let Some(value) = view.date_value {
        DateTime::new(value.to_owned()).map_err(|_source_error| {
            Error::InvalidFormat("date fullDate is not a valid bounded xsd:dateTime".into())
        })?;
    }
    if let Some(binding) = view.data_binding {
        validate_xml_scalars(binding.xpath(), "data-binding XPath")?;
        validate_xml_scalars(binding.store_item_id(), "data-binding store item ID")?;
        if let Some(value) = binding.prefix_mappings() {
            validate_xml_scalars(value, "data-binding prefix mappings")?;
        }
        validate_data_binding_values(
            binding.xpath(),
            binding.store_item_id(),
            binding.prefix_mappings(),
        )?;
    }
    Ok(())
}

fn validate_xml_scalars(value: &str, label: &str) -> Result<()> {
    if value.chars().all(|character| {
        matches!(
            character,
            '\u{9}'
                | '\u{A}'
                | '\u{D}'
                | '\u{20}'..='\u{D7FF}'
                | '\u{E000}'..='\u{FFFD}'
                | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "{label} contains a character forbidden by XML 1.0"
        )))
    }
}

fn requirements(view: &AuthoringView<'_>) -> NamespaceRequirements {
    NamespaceRequirements {
        word_2010: matches!(view.kind, Kind::Checkbox | Kind::EntityPicker)
            || matches!(view.date_calendar, Some(Calendar::Umalqura)),
        word_2012: matches!(view.kind, Kind::RepeatingSection | Kind::RepeatingItem)
            || view
                .data_binding
                .is_some_and(|binding| binding.flavor() == BindingFlavor::Word2012)
            || view.appearance.is_some()
            || view.color.is_some()
            || view.web_extension_linked.is_some()
            || view.web_extension_created.is_some(),
        checksum: view
            .data_binding
            .and_then(DataBinding::checksum_value)
            .is_some(),
        formatting_allowed: view.formatting_allowed.is_some(),
    }
}

fn write_optional_value(xml: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    if let Some(value) = value {
        write!(xml, r#"<w:{name} w:val="{}"/>"#, escape_xml(value))?;
    }
    Ok(())
}

fn write_kind(xml: &mut String, view: &AuthoringView<'_>) -> Result<()> {
    if let Some(appearance) = view.appearance {
        write!(
            xml,
            r#"<w15:appearance w15:val="{}"/>"#,
            appearance.as_str()
        )?;
    }
    if let Some(color) = view.color {
        match color {
            SdtColor::Auto => xml.push_str(r#"<w15:color w:val="auto"/>"#),
            SdtColor::Rgb([red, green, blue]) => {
                write!(
                    xml,
                    r#"<w15:color w:val="{red:02X}{green:02X}{blue:02X}"/>"#
                )?;
            },
        }
    }
    if let Some(value) = view.web_extension_linked {
        write!(
            xml,
            r#"<w15:webExtensionLinked w:val="{}"/>"#,
            u8::from(value)
        )?;
    }
    if let Some(value) = view.web_extension_created {
        write!(
            xml,
            r#"<w15:webExtensionCreated w:val="{}"/>"#,
            u8::from(value)
        )?;
    }
    match view.kind {
        Kind::ComboBox | Kind::Dropdown => {
            let name = view.kind.as_str();
            write!(xml, "<w:{name}>")?;
            for (display, value) in view.list_items {
                write!(
                    xml,
                    r#"<w:listItem w:displayText="{}" w:value="{}"/>"#,
                    escape_xml(display),
                    escape_xml(value)
                )?;
            }
            write!(xml, "</w:{name}>")?;
        },
        Kind::Date => {
            xml.push_str("<w:date");
            if let Some(value) = view.date_value {
                write!(xml, r#" w:fullDate="{}""#, escape_xml(value))?;
            }
            if view.date_format.is_none() && view.date_calendar.is_none() {
                xml.push_str("/>");
                return Ok(());
            }
            xml.push('>');
            if let Some(format) = view.date_format {
                write!(xml, r#"<w:dateFormat w:val="{}"/>"#, escape_xml(format))?;
            }
            if let Some(calendar) = view.date_calendar {
                write_calendar(xml, calendar)?;
            }
            xml.push_str("</w:date>");
        },
        Kind::Checkbox => {
            xml.push_str("<w14:checkbox>");
            write!(
                xml,
                r#"<w14:checked w14:val="{}"/>"#,
                u8::from(view.checked.unwrap_or(false))
            )?;
            xml.push_str("</w14:checkbox>");
        },
        Kind::EntityPicker => xml.push_str("<w14:entityPicker/>"),
        Kind::RepeatingSection => {
            xml.push_str("<w15:repeatingSection>");
            if let Some(title) = view.repeating_section_title {
                write!(xml, r#"<w15:sectionTitle w:val="{}"/>"#, escape_xml(title))?;
            }
            xml.push_str("</w15:repeatingSection>");
        },
        Kind::RepeatingItem => xml.push_str("<w15:repeatingSectionItem/>"),
        kind @ (Kind::RichText
        | Kind::Text
        | Kind::Picture
        | Kind::Citation
        | Kind::Equation
        | Kind::Group
        | Kind::DocPartList
        | Kind::DocPart
        | Kind::Bibliography) => write!(xml, "<w:{}/>", kind.as_str())?,
    }
    Ok(())
}

fn write_calendar(xml: &mut String, calendar: &Calendar) -> Result<()> {
    if matches!(calendar, Calendar::Umalqura) {
        xml.push_str("<mc:AlternateContent><mc:Choice Requires=\"w14\"><w:calendar w:val=\"umalqura\"/></mc:Choice><mc:Fallback><w:calendar w:val=\"hijri\"/></mc:Fallback></mc:AlternateContent>");
    } else {
        write!(
            xml,
            r#"<w:calendar w:val="{}"/>"#,
            escape_xml(calendar.as_str())
        )?;
    }
    Ok(())
}

fn write_binding(xml: &mut String, binding: Option<&DataBinding>) -> Result<()> {
    let Some(binding) = binding else {
        return Ok(());
    };
    let prefix = match binding.flavor() {
        BindingFlavor::Core => "w",
        BindingFlavor::Word2012 => "w15",
    };
    write!(xml, "<{prefix}:dataBinding")?;
    if let Some(prefix_mappings) = binding.prefix_mappings() {
        write!(
            xml,
            r#" w:prefixMappings="{}""#,
            escape_xml(prefix_mappings)
        )?;
    }
    write!(
        xml,
        r#" w:xpath="{}" w:storeItemID="{}""#,
        escape_xml(binding.xpath()),
        escape_xml(binding.store_item_id())
    )?;
    if let Some(checksum) = binding.checksum() {
        write!(
            xml,
            r#" w16sdtdh:storeItemChecksum="{}""#,
            checksum.to_base64()
        )?;
    } else if binding.checksum_value().is_some() {
        return Err(Error::InvalidFormat(
            "malformed storeItemChecksum cannot be canonically authored".into(),
        ));
    }
    xml.push_str("/>");
    Ok(())
}

fn write_lock(
    xml: &mut String,
    lock: Lock,
    formatting_allowed: Option<FormattingAllowed>,
) -> Result<()> {
    if lock == Lock::Unlocked {
        return Ok(());
    }
    write!(xml, r#"<w:lock w:val="{}""#, lock.as_str())?;
    if let Some(allowed) = formatting_allowed {
        write!(
            xml,
            r#" w16sdtfl:formattingAllowed="{}""#,
            u8::from(allowed.as_bool())
        )?;
    }
    xml.push_str("/>");
    Ok(())
}

fn metadata_bytes(view: &AuthoringView<'_>) -> Result<usize> {
    let mut bytes = 0usize;
    for value in [
        view.tag,
        view.title,
        view.placeholder_doc_part,
        view.date_format,
        view.date_value,
        view.date_calendar.map(Calendar::as_str),
        view.repeating_section_title,
        view.data_binding.map(DataBinding::xpath),
        view.data_binding.map(DataBinding::store_item_id),
        view.data_binding.and_then(DataBinding::prefix_mappings),
    ]
    .into_iter()
    .flatten()
    {
        bytes = bytes.checked_add(value.len()).ok_or_else(metadata_limit)?;
    }
    for (display, value) in view.list_items {
        bytes = bytes
            .checked_add(display.len())
            .and_then(|size| size.checked_add(value.len()))
            .ok_or_else(metadata_limit)?;
    }
    if let Some(checksum) = view.data_binding.and_then(DataBinding::checksum_value) {
        bytes = bytes
            .checked_add(checksum.lexical().len())
            .ok_or_else(metadata_limit)?;
    }
    Ok(bytes)
}

fn output_estimate(view: &AuthoringView<'_>, metadata_bytes: usize) -> Result<usize> {
    let escaped = metadata_bytes.checked_mul(6).ok_or_else(output_limit)?;
    let item_markup = view
        .list_items
        .len()
        .checked_mul(64)
        .ok_or_else(output_limit)?;
    1024usize
        .checked_add(escaped)
        .and_then(|size| size.checked_add(item_markup))
        .ok_or_else(output_limit)
}

fn output_limit() -> Error {
    Error::InvalidFormat("content-control authored output exceeds configured limit".into())
}

fn metadata_limit() -> Error {
    Error::InvalidFormat("content-control authored metadata exceeds configured limit".into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn word_2012_namespace_and_token_are_authored_once() {
        let view = AuthoringView::new(1, Kind::RepeatingSection, Lock::Unlocked)
            .repeating_section_title(Some("Items"));
        let authored = write_sdt_pr(&view, &Limits::default()).unwrap();
        assert!(authored.namespace_requirements().word_2012());
        assert_eq!(authored.xml().matches("xmlns:w15=").count(), 1);
        assert_eq!(authored.xml().matches(r#"mc:Ignorable="w15""#).count(), 1);
    }

    #[test]
    fn formatting_invariant_and_output_limit_are_enforced() {
        let invalid = AuthoringView::new(1, Kind::Text, Lock::Unlocked)
            .formatting_allowed(Some(FormattingAllowed::Allowed));
        assert!(write_sdt_pr(&invalid, &Limits::default()).is_err());

        let limits = Limits {
            max_output_bytes: 16,
            ..Limits::default()
        };
        let view = AuthoringView::new(1, Kind::Text, Lock::Unlocked);
        assert!(write_sdt_pr(&view, &limits).is_err());

        let limits = Limits {
            max_metadata_bytes: 1,
            ..Limits::default()
        };
        let view = AuthoringView::new(1, Kind::Text, Lock::Unlocked).tag(Some("ab"));
        assert!(write_sdt_pr(&view, &limits).is_err());
    }

    #[test]
    fn word_2012_binding_preserves_its_normative_vocabulary() {
        let binding =
            DataBinding::word_2012("/root/value", "{00112233-4455-6677-8899-AABBCCDDEEFF}")
                .unwrap();
        let view = AuthoringView::new(1, Kind::Text, Lock::Unlocked).data_binding(Some(&binding));
        let authored = write_sdt_pr(&view, &Limits::default()).unwrap();
        assert!(authored.xml().contains("<w15:dataBinding"));
        assert!(authored.xml().contains(r#"mc:Ignorable="w15""#));
        assert!(!authored.xml().contains("<w:dataBinding"));
    }

    #[test]
    fn umalqura_calendar_is_authored_with_its_required_hijri_fallback() {
        let calendar = Calendar::Umalqura;
        let view = AuthoringView::new(1, Kind::Date, Lock::Unlocked).date_calendar(Some(&calendar));
        let authored = write_sdt_pr(&view, &Limits::default()).unwrap();
        assert_eq!(
            authored.xml(),
            r#"<w:sdtPr xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:id w:val="1"/><w:date><mc:AlternateContent><mc:Choice Requires="w14"><w:calendar w:val="umalqura"/></mc:Choice><mc:Fallback><w:calendar w:val="hijri"/></mc:Fallback></mc:AlternateContent></w:date></w:sdtPr>"#
        );
        assert!(authored.namespace_requirements().word_2010());
    }

    #[test]
    fn children_follow_ct_sdt_pr_schema_order() {
        let binding =
            DataBinding::new("/root/value", "{00112233-4455-6677-8899-AABBCCDDEEFF}").unwrap();
        let view = AuthoringView::new(7, Kind::Text, Lock::ContentLocked)
            .title(Some("Alias"))
            .placeholder_doc_part(Some("Placeholder"))
            .showing_placeholder(true)
            .data_binding(Some(&binding))
            .temporary(true)
            .tag(Some("Tag"));
        let authored = write_sdt_pr(&view, &Limits::default()).unwrap();
        assert_eq!(
            authored.xml(),
            concat!(
                r#"<w:sdtPr>"#,
                r#"<w:alias w:val="Alias"/>"#,
                r#"<w:lock w:val="contentLocked"/>"#,
                r#"<w:placeholder><w:docPart w:val="Placeholder"/></w:placeholder>"#,
                r#"<w:showingPlcHdr/>"#,
                r#"<w:dataBinding w:xpath="/root/value" w:storeItemID="{00112233-4455-6677-8899-AABBCCDDEEFF}"/>"#,
                r#"<w:temporary/>"#,
                r#"<w:id w:val="7"/>"#,
                r#"<w:tag w:val="Tag"/>"#,
                r#"<w:text/>"#,
                r#"</w:sdtPr>"#
            )
        );
    }

    #[test]
    fn invalid_xml_scalar_and_full_date_are_refused() {
        let invalid_scalar =
            AuthoringView::new(1, Kind::Text, Lock::Unlocked).tag(Some("bad\u{1}value"));
        assert!(write_sdt_pr(&invalid_scalar, &Limits::default()).is_err());

        let invalid_date = AuthoringView::new(1, Kind::Date, Lock::Unlocked)
            .date_value(Some("2026-02-30T04:05:06Z"));
        assert!(write_sdt_pr(&invalid_date, &Limits::default()).is_err());

        let valid_date = AuthoringView::new(1, Kind::Date, Lock::Unlocked)
            .date_value(Some("2026-08-08T04:05:06Z"));
        assert!(write_sdt_pr(&valid_date, &Limits::default()).is_ok());
    }

    #[test]
    fn entity_picker_and_list_item_limits_are_refused() {
        let binding =
            DataBinding::new("/root/value", "{00112233-4455-6677-8899-AABBCCDDEEFF}").unwrap();
        let entity =
            AuthoringView::new(1, Kind::EntityPicker, Lock::Unlocked).data_binding(Some(&binding));
        assert!(write_sdt_pr(&entity, &Limits::default()).is_err());

        let items = vec![
            ("one".to_owned(), "1".to_owned()),
            ("two".to_owned(), "2".to_owned()),
        ];
        let view = AuthoringView::new(1, Kind::Dropdown, Lock::Unlocked).list_items(&items);
        let limits = Limits {
            max_list_items_per_control: 1,
            ..Limits::default()
        };
        assert!(write_sdt_pr(&view, &limits).is_err());

        let limits = Limits {
            max_list_items: 1,
            ..Limits::default()
        };
        assert!(write_sdt_pr(&view, &limits).is_err());
    }
}
