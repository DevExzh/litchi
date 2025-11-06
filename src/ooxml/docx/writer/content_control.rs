/// Content control writer support for DOCX documents.
///
/// Content controls are structured document regions that can contain text, dates,
/// drop-down lists, and other content types. They're commonly used in templates
/// and forms.
use crate::ooxml::error::Result;
use std::fmt::Write as FmtWrite;

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
    pub fn rich_text(id: u32, tag: Option<&str>) -> Self {
        Self {
            id,
            tag: tag.map(|s| s.to_string()),
            title: None,
            control_type: ContentControlType::RichText,
            allow_delete: true,
            allow_edit: true,
            placeholder: None,
        }
    }

    /// Create a new plain text content control.
    ///
    /// # Arguments
    ///
    /// * `id` - Unique control ID
    /// * `tag` - Optional tag for programmatic identification
    pub fn plain_text(id: u32, tag: Option<&str>) -> Self {
        Self {
            id,
            tag: tag.map(|s| s.to_string()),
            title: None,
            control_type: ContentControlType::PlainText,
            allow_delete: true,
            allow_edit: true,
            placeholder: None,
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
    pub fn dropdown(id: u32, tag: Option<&str>, items: Vec<(String, String)>) -> Self {
        Self {
            id,
            tag: tag.map(|s| s.to_string()),
            title: None,
            control_type: ContentControlType::DropDownList { items },
            allow_delete: true,
            allow_edit: true,
            placeholder: None,
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
            tag: tag.map(|s| s.to_string()),
            title: None,
            control_type: ContentControlType::DatePicker {
                format: format.into(),
            },
            allow_delete: true,
            allow_edit: true,
            placeholder: None,
        }
    }

    /// Create a new checkbox content control.
    ///
    /// # Arguments
    ///
    /// * `id` - Unique control ID
    /// * `tag` - Optional tag for programmatic identification
    /// * `checked` - Initial checked state
    pub fn checkbox(id: u32, tag: Option<&str>, checked: bool) -> Self {
        Self {
            id,
            tag: tag.map(|s| s.to_string()),
            title: None,
            control_type: ContentControlType::Checkbox { checked },
            allow_delete: true,
            allow_edit: true,
            placeholder: None,
        }
    }

    /// Get the control ID.
    #[inline]
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Get the control tag.
    #[inline]
    pub fn tag(&self) -> Option<&str> {
        self.tag.as_deref()
    }

    /// Set the control tag.
    pub fn set_tag(&mut self, tag: Option<String>) {
        self.tag = tag;
    }

    /// Get the control title.
    #[inline]
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
    }

    /// Set placeholder text.
    pub fn set_placeholder(&mut self, placeholder: Option<String>) {
        self.placeholder = placeholder;
    }

    /// Get the content control type.
    pub fn control_type(&self) -> &ContentControlType {
        &self.control_type
    }

    /// Generate XML for this content control (start tag).
    ///
    /// Content controls wrap around content, so this generates the opening tags.
    #[allow(dead_code)]
    pub(crate) fn to_xml_start(&self) -> Result<String> {
        let mut xml = String::with_capacity(512);

        // Start content control properties
        write!(&mut xml, r#"<w:sdt><w:sdtPr>"#)?;

        // Add ID
        write!(&mut xml, r#"<w:id w:val="{}"/>"#, self.id)?;

        // Add tag if present
        if let Some(ref tag) = self.tag {
            write!(&mut xml, r#"<w:tag w:val="{}"/>"#, escape_xml(tag))?;
        }

        // Add title if present
        if let Some(ref title) = self.title {
            write!(&mut xml, r#"<w:alias w:val="{}"/>"#, escape_xml(title))?;
        }

        // Add control type-specific properties
        match &self.control_type {
            ContentControlType::RichText => {
                xml.push_str("<w:richText/>");
            },
            ContentControlType::PlainText => {
                xml.push_str("<w:text/>");
            },
            ContentControlType::DropDownList { items } => {
                xml.push_str("<w:dropDownList>");
                for (display, value) in items.iter() {
                    write!(
                        &mut xml,
                        r#"<w:listItem w:displayText="{}" w:value="{}"/>"#,
                        escape_xml(display),
                        escape_xml(value)
                    )?;
                }
                xml.push_str("</w:dropDownList>");
            },
            ContentControlType::DatePicker { format } => {
                write!(
                    &mut xml,
                    r#"<w:date w:fullDate="2000-01-01T00:00:00Z"><w:dateFormat w:val="{}"/></w:date>"#,
                    escape_xml(format)
                )?;
            },
            ContentControlType::Checkbox { checked } => {
                if *checked {
                    xml.push_str(r#"<w14:checkbox><w14:checked w14:val="1"/></w14:checkbox>"#);
                } else {
                    xml.push_str(r#"<w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>"#);
                }
            },
        }

        // Add placeholder if present
        if let Some(ref placeholder) = self.placeholder {
            write!(
                &mut xml,
                r#"<w:placeholder><w:docPart w:val="{}"/></w:placeholder>"#,
                escape_xml(placeholder)
            )?;
        }

        // Add lock properties
        if !self.allow_delete {
            xml.push_str("<w:lock w:val=\"sdtContentLocked\"/>");
        }
        if !self.allow_edit {
            xml.push_str("<w:lock w:val=\"contentLocked\"/>");
        }

        xml.push_str("</w:sdtPr><w:sdtContent>");

        Ok(xml)
    }

    /// Generate XML for content control end tag.
    #[allow(dead_code)]
    pub(crate) fn to_xml_end() -> &'static str {
        "</w:sdtContent></w:sdt>"
    }
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
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
        let mut control = MutableContentControl::plain_text(1, None);
        control.set_allow_delete(false);
        control.set_allow_edit(false);

        let xml = control.to_xml_start().unwrap();
        assert!(xml.contains("w:lock"));
    }
}
