#![expect(
    clippy::module_name_repetitions,
    reason = "public names retain established OOXML facade terminology"
)]
/// Content control writer support for DOCX documents.
///
/// Content controls are structured document regions that can contain text, dates,
/// drop-down lists, and other content types. They're commonly used in templates
/// and forms.
use crate::content_control::{AuthoringView, Kind, Limits, Lock, write_sdt_pr};
use crate::error::{Error, Result};

pub use crate::content_control::{Checksum, DataBinding, FormattingAllowed};

/// A mutable content control in a Word document.
///
/// Content controls provide structured editing regions with validation,
/// data binding, and user interface enhancements.
#[derive(Debug, Clone)]
pub struct MutableContentControl {
    /// Control ID (unique within document)
    id: u32,
    /// Control tag (for programmatic identification)
    tag: Option<String>,
    /// Control title (displayed to user)
    title: Option<String>,
    /// Content control type
    control_type: ContentControlType,
    /// Whether the control can be deleted
    allow_delete: bool,
    /// Whether the content can be edited
    allow_edit: bool,
    /// Placeholder text
    placeholder: Option<String>,
    /// Explicit glossary document-part name used by w:placeholder.
    placeholder_doc_part: Option<String>,
    /// Optional fullDate lexical value for a date picker.
    date_value: Option<String>,
    /// Optional checked custom XML binding.
    data_binding: Option<DataBinding>,
    /// Word 2024 formatting exception for a content-locked control.
    formatting_allowed: Option<FormattingAllowed>,
}

/// Type of content control.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ContentControlType {
    /// Rich text content control (can contain formatted text and paragraphs)
    RichText,
    /// Plain text content control (text only, no formatting)
    PlainText,
    /// Drop-down list content control
    DropDownList {
        /// List items (display text, value)
        items: Vec<(String, String)>,
    },
    /// Date picker content control
    DatePicker {
        /// Date format string
        format: String,
    },
    /// Checkbox content control
    Checkbox {
        /// Checked state
        checked: bool,
    },
}

impl MutableContentControl {
    /// Create a new rich text content control.
    ///
    /// # Arguments
    ///
    /// * `id` - Unique control ID
    /// * `tag` - Optional tag for programmatic identification
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let control = MutableContentControl::rich_text(1, Some("CustomerName"));
    /// ```
    #[must_use]
    pub fn rich_text(id: u32, tag: Option<&str>) -> Self {
        Self {
            id,
            tag: tag.map(ToString::to_string),
            title: None,
            control_type: ContentControlType::RichText,
            allow_delete: true,
            allow_edit: true,
            placeholder: None,
            placeholder_doc_part: None,
            date_value: None,
            data_binding: None,
            formatting_allowed: None,
        }
    }

    /// Create a new plain text content control.
    ///
    /// # Arguments
    ///
    /// * `id` - Unique control ID
    /// * `tag` - Optional tag for programmatic identification
    #[must_use]
    pub fn plain_text(id: u32, tag: Option<&str>) -> Self {
        Self {
            id,
            tag: tag.map(ToString::to_string),
            title: None,
            control_type: ContentControlType::PlainText,
            allow_delete: true,
            allow_edit: true,
            placeholder: None,
            placeholder_doc_part: None,
            date_value: None,
            data_binding: None,
            formatting_allowed: None,
        }
    }

    /// Create a new drop-down list content control.
    ///
    /// # Arguments
    ///
    /// * `id` - Unique control ID
    /// * `tag` - Optional tag for programmatic identification
    /// * `items` - List items (display text, value)
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let items = vec![
    ///     ("Red".to_string(), "red".to_string()),
    ///     ("Green".to_string(), "green".to_string()),
    ///     ("Blue".to_string(), "blue".to_string()),
    /// ];
    /// let control = MutableContentControl::dropdown(1, Some("Color"), items);
    /// ```
    #[must_use]
    pub fn dropdown(id: u32, tag: Option<&str>, items: Vec<(String, String)>) -> Self {
        Self {
            id,
            tag: tag.map(ToString::to_string),
            title: None,
            control_type: ContentControlType::DropDownList { items },
            allow_delete: true,
            allow_edit: true,
            placeholder: None,
            placeholder_doc_part: None,
            date_value: None,
            data_binding: None,
            formatting_allowed: None,
        }
    }

    /// Create a new date picker content control.
    ///
    /// # Arguments
    ///
    /// * `id` - Unique control ID
    /// * `tag` - Optional tag for programmatic identification
    /// * `format` - Date format string (e.g., "MM/dd/yyyy")
    pub fn date_picker(id: u32, tag: Option<&str>, format: impl Into<String>) -> Self {
        Self {
            id,
            tag: tag.map(ToString::to_string),
            title: None,
            control_type: ContentControlType::DatePicker {
                format: format.into(),
            },
            allow_delete: true,
            allow_edit: true,
            placeholder: None,
            placeholder_doc_part: None,
            date_value: None,
            data_binding: None,
            formatting_allowed: None,
        }
    }

    /// Create a new checkbox content control.
    ///
    /// # Arguments
    ///
    /// * `id` - Unique control ID
    /// * `tag` - Optional tag for programmatic identification
    /// * `checked` - Initial checked state
    #[must_use]
    pub fn checkbox(id: u32, tag: Option<&str>, checked: bool) -> Self {
        Self {
            id,
            tag: tag.map(ToString::to_string),
            title: None,
            control_type: ContentControlType::Checkbox { checked },
            allow_delete: true,
            allow_edit: true,
            placeholder: None,
            placeholder_doc_part: None,
            date_value: None,
            data_binding: None,
            formatting_allowed: None,
        }
    }

    /// Get the control ID.
    #[inline]
    #[must_use]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the control tag.
    #[inline]
    #[must_use]
    pub fn tag(&self) -> Option<&str> {
        self.tag.as_deref()
    }

    /// Set the control tag.
    pub fn set_tag(&mut self, tag: Option<String>) {
        self.tag = tag;
    }

    /// Get the control title.
    #[inline]
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Set the control title.
    pub fn set_title(&mut self, title: Option<String>) {
        self.title = title;
    }

    /// Set whether the control can be deleted.
    pub fn set_allow_delete(&mut self, allow: bool) {
        self.allow_delete = allow;
    }

    /// Set whether the content can be edited.
    pub fn set_allow_edit(&mut self, allow: bool) {
        self.allow_edit = allow;
        if allow {
            self.formatting_allowed = None;
        }
    }

    /// Set legacy literal placeholder text.
    ///
    /// `WordprocessingML` placeholder properties reference glossary document
    /// parts rather than containing literal text. A control retaining this
    /// legacy value is rejected during checked serialization.
    pub fn set_placeholder(&mut self, placeholder: Option<String>) {
        self.placeholder = placeholder;
        if self.placeholder.is_some() {
            self.placeholder_doc_part = None;
        }
    }

    /// Set an explicit glossary document-part name for w:placeholder.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_placeholder_doc_part(&mut self, name: Option<String>) -> Result<&mut Self> {
        self.placeholder_doc_part = name;
        self.placeholder = None;
        Ok(self)
    }

    /// Set an explicit glossary document-part name using builder syntax.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn with_placeholder_doc_part(mut self, name: impl Into<String>) -> Result<Self> {
        self.set_placeholder_doc_part(Some(name.into()))?;
        Ok(self)
    }

    /// Set the optional bounded xsd:dateTime fullDate value of a date picker.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_date_value(&mut self, value: Option<String>) -> Result<&mut Self> {
        if value.is_some() && !matches!(self.control_type, ContentControlType::DatePicker { .. }) {
            return Err(Error::InvalidFormat(
                "fullDate requires a date-picker content control".into(),
            ));
        }
        self.date_value = value;
        Ok(self)
    }

    /// Return the custom XML data binding, when present.
    #[inline]
    #[must_use]
    pub fn data_binding(&self) -> Option<&DataBinding> {
        self.data_binding.as_ref()
    }

    /// Replace the custom XML data binding with a validated semantic value.
    pub fn set_data_binding(&mut self, binding: Option<DataBinding>) -> &mut Self {
        self.data_binding = binding;
        self
    }

    /// Construct and install a checked custom XML data binding.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn bind(
        &mut self,
        xpath: impl Into<String>,
        store_item_id: impl Into<String>,
    ) -> Result<&mut Self> {
        self.data_binding = Some(DataBinding::new(xpath, store_item_id)?);
        Ok(self)
    }

    /// Construct and install a checked Word 2012 formatted data binding.
    ///
    /// The `XPath` remains inert lexical metadata and is never executed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn bind_word_2012(
        &mut self,
        xpath: impl Into<String>,
        store_item_id: impl Into<String>,
    ) -> Result<&mut Self> {
        self.data_binding = Some(DataBinding::word_2012(xpath, store_item_id)?);
        Ok(self)
    }

    /// Install a validated custom XML data binding using builder syntax.
    #[must_use]
    pub fn with_data_binding(mut self, binding: DataBinding) -> Self {
        self.data_binding = Some(binding);
        self
    }

    /// Construct and install a checked binding using builder syntax.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn with_binding(
        mut self,
        xpath: impl Into<String>,
        store_item_id: impl Into<String>,
    ) -> Result<Self> {
        self.bind(xpath, store_item_id)?;
        Ok(self)
    }

    /// Construct and install a checked Word 2012 binding using builder syntax.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn with_word_2012_binding(
        mut self,
        xpath: impl Into<String>,
        store_item_id: impl Into<String>,
    ) -> Result<Self> {
        self.bind_word_2012(xpath, store_item_id)?;
        Ok(self)
    }

    /// Return the Word 2024 formatting exception, when present.
    #[inline]
    #[must_use]
    pub fn formatting_allowed(&self) -> Option<FormattingAllowed> {
        self.formatting_allowed
    }

    /// Set the Word 2024 formatting exception.
    ///
    /// The extension is meaningful only for contentLocked and
    /// sdtContentLocked, so setting it on an editable control is rejected.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_formatting_allowed(
        &mut self,
        allowed: Option<FormattingAllowed>,
    ) -> Result<&mut Self> {
        if allowed.is_some() && self.allow_edit {
            return Err(Error::InvalidFormat(
                "formattingAllowed requires contentLocked or sdtContentLocked".into(),
            ));
        }
        self.formatting_allowed = allowed;
        Ok(self)
    }

    /// Set the Word 2024 formatting exception using builder syntax.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn with_formatting_allowed(mut self, allowed: FormattingAllowed) -> Result<Self> {
        self.set_formatting_allowed(Some(allowed))?;
        Ok(self)
    }

    /// Get the content control type.
    #[must_use]
    pub fn control_type(&self) -> &ContentControlType {
        &self.control_type
    }

    /// Generate XML for this content control (start tag).
    ///
    /// Content controls wrap around content, so this generates the opening tags.
    #[allow(
        dead_code,
        reason = "writer helper is retained for package integration"
    )]
    pub(crate) fn to_xml_start(&self) -> Result<String> {
        if self.placeholder.is_some() {
            return Err(Error::InvalidFormat(
                "literal placeholder text is not a WordprocessingML placeholder; use a glossary document-part reference".into(),
            ));
        }
        let kind = match self.control_type {
            ContentControlType::RichText => Kind::RichText,
            ContentControlType::PlainText => Kind::Text,
            ContentControlType::DropDownList { .. } => Kind::Dropdown,
            ContentControlType::DatePicker { .. } => Kind::Date,
            ContentControlType::Checkbox { .. } => Kind::Checkbox,
        };
        let lock = match (self.allow_delete, self.allow_edit) {
            (true, true) => Lock::Unlocked,
            (false, true) => Lock::SdtLocked,
            (true, false) => Lock::ContentLocked,
            (false, false) => Lock::SdtContentLocked,
        };
        let mut view = AuthoringView::new(self.id, kind, lock)
            .tag(self.tag.as_deref())
            .title(self.title.as_deref())
            .placeholder_doc_part(self.placeholder_doc_part.as_deref())
            .data_binding(self.data_binding.as_ref())
            .date_value(self.date_value.as_deref())
            .formatting_allowed(self.formatting_allowed);
        match &self.control_type {
            ContentControlType::DropDownList { items } => view = view.list_items(items),
            ContentControlType::DatePicker { format } => {
                view = view.date_format(Some(format));
            },
            ContentControlType::Checkbox { checked } => view = view.checked(Some(*checked)),
            ContentControlType::RichText | ContentControlType::PlainText => {},
        }

        let properties = write_sdt_pr(&view, &Limits::default())?;
        let mut xml = properties.into_xml();
        xml.try_reserve("<w:sdt><w:sdtContent>".len())
            .map_err(|_source_error| {
                Error::InvalidFormat("content-control output allocation failed".into())
            })?;
        xml.insert_str(0, "<w:sdt>");
        xml.push_str("<w:sdtContent>");
        Ok(xml)
    }

    /// Generate XML for content control end tag.
    #[allow(
        dead_code,
        reason = "writer helper is retained for package integration"
    )]
    pub(crate) fn to_xml_end() -> &'static str {
        "</w:sdtContent></w:sdt>"
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_rich_text_control() {
        let control = MutableContentControl::rich_text(1, Some("TestTag"));
        assert_eq!(control.id(), 1);
        assert_eq!(control.tag(), Some("TestTag"));
        assert!(matches!(
            control.control_type(),
            ContentControlType::RichText
        ));
    }

    #[test]
    fn test_plain_text_control() {
        let control = MutableContentControl::plain_text(2, None);
        assert_eq!(control.id(), 2);
        assert_eq!(control.tag(), None);
        assert!(matches!(
            control.control_type(),
            ContentControlType::PlainText
        ));
    }

    #[test]
    fn test_dropdown_control() {
        let items = vec![
            ("Option 1".to_string(), "opt1".to_string()),
            ("Option 2".to_string(), "opt2".to_string()),
        ];
        let control = MutableContentControl::dropdown(3, Some("Dropdown"), items.clone());
        assert_eq!(control.id(), 3);

        if let ContentControlType::DropDownList { items: ctrl_items } = control.control_type() {
            assert_eq!(ctrl_items.len(), 2);
            assert_eq!(ctrl_items[0].0, "Option 1");
        } else {
            panic!("Wrong control type");
        }
    }

    #[test]
    fn test_date_picker_control() {
        let control = MutableContentControl::date_picker(4, None, "MM/dd/yyyy");
        assert_eq!(control.id(), 4);

        if let ContentControlType::DatePicker { format } = control.control_type() {
            assert_eq!(format, "MM/dd/yyyy");
        } else {
            panic!("Wrong control type");
        }
    }

    #[test]
    fn test_checkbox_control() {
        let control = MutableContentControl::checkbox(5, Some("Check"), true);
        assert_eq!(control.id(), 5);

        if let ContentControlType::Checkbox { checked } = control.control_type() {
            assert!(*checked);
        } else {
            panic!("Wrong control type");
        }
    }

    #[test]
    fn test_xml_generation() {
        let mut control = MutableContentControl::rich_text(1, Some("MyTag"));
        control.set_title(Some("My Control".to_string()));

        let xml = control.to_xml_start().unwrap();
        // Debug output
        eprintln!("Generated XML: {}", xml);
        assert!(xml.contains(r#"w:val="1""#));
        assert!(xml.contains(r#"w:val="MyTag""#));
        assert!(xml.contains(r#"w:val="My Control""#));
        assert!(xml.contains("<w:richText/>"));
        assert!(xml.contains("<w:sdtContent>"));
    }

    #[test]
    fn test_lock_properties() {
        let cases = [
            (true, true, None),
            (false, true, Some("sdtLocked")),
            (true, false, Some("contentLocked")),
            (false, false, Some("sdtContentLocked")),
        ];
        for (allow_delete, allow_edit, expected) in cases {
            let mut control = MutableContentControl::plain_text(1, None);
            control.set_allow_delete(allow_delete);
            control.set_allow_edit(allow_edit);
            let xml = control.to_xml_start().unwrap();
            assert_eq!(
                xml.matches("<w:lock ").count(),
                usize::from(expected.is_some())
            );
            if let Some(expected) = expected {
                assert!(xml.contains(&format!(r#"w:val="{expected}""#)));
            }
        }
    }

    #[test]
    fn generated_checkbox_round_trips_through_reader() {
        let control = MutableContentControl::checkbox(5, Some("Check & Verify"), true);
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{}{}</w:document>"#,
            control.to_xml_start().unwrap(),
            MutableContentControl::to_xml_end()
        );
        let parsed =
            crate::content_control::ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].id(), 5);
        assert_eq!(parsed[0].tag(), Some("Check & Verify"));
        assert_eq!(parsed[0].kind(), Kind::Checkbox);
        assert_eq!(parsed[0].checked(), Some(true));
    }

    #[test]
    fn checked_data_binding_checksum_is_canonical_and_reopens() {
        let checksum = Checksum::from_bytes([0x4d, 0x3c, 0x2b, 0x1a]);
        let binding = DataBinding::new("/x:root/x:value", "{00112233-4455-6677-8899-AABBCCDDEEFF}")
            .unwrap()
            .with_prefix_mappings("xmlns:x='urn:example'")
            .unwrap()
            .with_checksum(checksum.clone());
        let control =
            MutableContentControl::plain_text(7, Some("bound")).with_data_binding(binding);

        let start = control.to_xml_start().unwrap();
        assert!(start.contains(r#"w16sdtdh:storeItemChecksum="TTwrGg==""#));
        assert!(start.contains(&format!(
            r#"xmlns:w16sdtdh="{}""#,
            crate::content_control::STORE_ITEM_CHECKSUM_NAMESPACE
        )));
        assert_eq!(start.matches("w16sdtdh").count(), 3);
        assert!(start.contains(r#"mc:Ignorable="w16sdtdh""#));

        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{}{}</w:document>"#,
            start,
            MutableContentControl::to_xml_end()
        );
        let reopened =
            crate::content_control::ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        let reopened_binding = reopened[0].data_binding().unwrap();
        assert_eq!(reopened_binding.xpath(), "/x:root/x:value");
        assert_eq!(
            reopened_binding.checksum().map(Checksum::as_bytes),
            Some(checksum.as_bytes())
        );
    }

    #[test]
    fn formatting_allowed_is_checked_and_reopens() {
        let mut invalid = MutableContentControl::plain_text(8, None);
        assert!(
            invalid
                .set_formatting_allowed(Some(FormattingAllowed::Allowed))
                .is_err()
        );

        let mut control = MutableContentControl::plain_text(9, None);
        control.set_allow_edit(false);
        control
            .set_formatting_allowed(Some(FormattingAllowed::Allowed))
            .unwrap();
        let start = control.to_xml_start().unwrap();
        assert!(
            start.contains(r#"<w:lock w:val="contentLocked" w16sdtfl:formattingAllowed="1"/>"#)
        );
        assert!(start.contains(r#"mc:Ignorable="w16sdtfl""#));

        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{}{}</w:document>"#,
            start,
            MutableContentControl::to_xml_end()
        );
        let reopened =
            crate::content_control::ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        assert_eq!(
            reopened[0].formatting_allowed(),
            Some(FormattingAllowed::Allowed)
        );

        control.set_allow_edit(true);
        assert_eq!(control.formatting_allowed(), None);
    }

    #[test]
    fn extension_ignorable_tokens_are_unique_and_stable() {
        let binding = DataBinding::new("/root/value", "{00112233-4455-6677-8899-AABBCCDDEEFF}")
            .unwrap()
            .with_checksum(Checksum::from_bytes([0, 0, 0, 0]));
        let mut control =
            MutableContentControl::checkbox(10, None, true).with_data_binding(binding);
        control.set_allow_edit(false);
        control
            .set_formatting_allowed(Some(FormattingAllowed::Disallowed))
            .unwrap();

        let xml = control.to_xml_start().unwrap();
        assert_eq!(
            xml.matches(r#"mc:Ignorable="w14 w16sdtdh w16sdtfl""#)
                .count(),
            1
        );
        for token in ["w14", "w16sdtdh", "w16sdtfl"] {
            assert_eq!(
                xml.matches(&format!("xmlns:{token}=")).count(),
                1,
                "duplicate namespace declaration for {token}"
            );
        }
    }

    #[test]
    fn word_2012_binding_without_checksum_reopens_with_its_flavor() {
        let control = MutableContentControl::plain_text(11, Some("formatted"))
            .with_word_2012_binding("/root/value", "{00112233-4455-6677-8899-AABBCCDDEEFF}")
            .unwrap();
        let start = control.to_xml_start().unwrap();
        assert!(start.contains("<w15:dataBinding"));
        assert!(start.contains(r#"mc:Ignorable="w15""#));
        assert!(!start.contains("storeItemChecksum"));

        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{}{}</w:document>"#,
            start,
            MutableContentControl::to_xml_end()
        );
        let reopened =
            crate::content_control::ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        assert_eq!(
            reopened[0].data_binding().unwrap().flavor(),
            crate::content_control::BindingFlavor::Word2012
        );
    }

    #[test]
    fn word_2012_binding_with_checksum_reopens_semantically() {
        let checksum = Checksum::from_bytes([1, 2, 3, 4]);
        let binding =
            DataBinding::word_2012("/root/value", "{00112233-4455-6677-8899-AABBCCDDEEFF}")
                .unwrap()
                .with_checksum(checksum.clone());
        let control =
            MutableContentControl::plain_text(12, Some("formatted")).with_data_binding(binding);
        let start = control.to_xml_start().unwrap();
        assert!(start.contains("<w15:dataBinding"));
        assert!(start.contains(r#"mc:Ignorable="w15 w16sdtdh""#));
        assert!(start.contains(r#"w16sdtdh:storeItemChecksum="AQIDBA==""#));

        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{}{}</w:document>"#,
            start,
            MutableContentControl::to_xml_end()
        );
        let reopened =
            crate::content_control::ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        let binding = reopened[0].data_binding().unwrap();
        assert_eq!(
            binding.flavor(),
            crate::content_control::BindingFlavor::Word2012
        );
        assert_eq!(
            binding.checksum().map(Checksum::as_bytes),
            Some(checksum.as_bytes())
        );
    }

    #[test]
    fn legacy_placeholder_text_refuses_and_explicit_reference_authors() {
        let mut legacy = MutableContentControl::plain_text(13, None);
        legacy.set_placeholder(Some("literal text".to_owned()));
        assert!(legacy.to_xml_start().is_err());

        let explicit = MutableContentControl::plain_text(14, None)
            .with_placeholder_doc_part("DefaultPlaceholder_22675703")
            .unwrap();
        assert!(explicit.to_xml_start().unwrap().contains(
            r#"<w:placeholder><w:docPart w:val="DefaultPlaceholder_22675703"/></w:placeholder>"#
        ));
    }

    #[test]
    fn writer_rejects_invalid_scalars_dates_and_list_quota() {
        let invalid = MutableContentControl::plain_text(15, Some("bad\u{1}tag"));
        assert!(invalid.to_xml_start().is_err());

        let mut date = MutableContentControl::date_picker(16, None, "yyyy-MM-dd");
        date.set_date_value(Some("2026-02-30T04:05:06Z".to_owned()))
            .unwrap();
        assert!(date.to_xml_start().is_err());

        let items = (0..=Limits::default().max_list_items_per_control)
            .map(|index| (index.to_string(), index.to_string()))
            .collect();
        let oversized = MutableContentControl::dropdown(17, None, items);
        assert!(oversized.to_xml_start().is_err());
    }
}
