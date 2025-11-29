/// Content control support for Word documents.
///
/// Content controls are structured regions in a document that can contain
/// specific types of content (text, dates, lists, etc.).
use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

/// A content control in a Word document.
///
/// Content controls provide structured content regions that can be
/// bound to data or restricted to specific content types.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::docx::Package;
///
/// let pkg = Package::open("document.docx")?;
/// let doc = pkg.document()?;
///
/// for control in doc.content_controls()? {
///     if let Some(tag) = control.tag() {
///         println!("Control {}: {}", control.id(), tag);
///     }
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone)]
pub struct ContentControl {
    /// Control ID
    id: u32,
    /// Control tag (optional identifier)
    tag: Option<String>,
    /// Control title
    title: Option<String>,
    /// Control type (text, date, comboBox, etc.)
    control_type: Option<String>,
    /// Whether the control can be deleted
    lock_delete: bool,
    /// Whether the content can be edited
    lock_content: bool,
}

impl ContentControl {
    /// Create a new ContentControl.
    pub fn new(
        id: u32,
        tag: Option<String>,
        title: Option<String>,
        control_type: Option<String>,
        lock_delete: bool,
        lock_content: bool,
    ) -> Self {
        Self {
            id,
            tag,
            title,
            control_type,
            lock_delete,
            lock_content,
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

    /// Get the control title.
    #[inline]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Get the control type.
    #[inline]
    pub fn control_type(&self) -> Option<&str> {
        self.control_type.as_deref()
    }

    /// Check if the control is locked for deletion.
    #[inline]
    pub fn is_lock_delete(&self) -> bool {
        self.lock_delete
    }

    /// Check if the content is locked for editing.
    #[inline]
    pub fn is_lock_content(&self) -> bool {
        self.lock_content
    }

    /// Extract content controls from document XML bytes.
    pub(crate) fn extract_from_document(doc_xml: &[u8]) -> Result<Vec<ContentControl>> {
        let mut reader = Reader::from_reader(doc_xml);
        reader.config_mut().trim_text(true);

        let mut controls = Vec::new();
        let mut in_sdt_pr = false;
        let mut current_id: Option<u32> = None;
        let mut current_tag: Option<String> = None;
        let mut current_title: Option<String> = None;
        let mut current_type: Option<String> = None;
        let mut current_lock_delete = false;
        let mut current_lock_content = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    match e.local_name().as_ref() {
                        b"sdtPr" => {
                            // Content control properties start
                            in_sdt_pr = true;
                            current_id = None;
                            current_tag = None;
                            current_title = None;
                            current_type = None;
                            current_lock_delete = false;
                            current_lock_content = false;
                        },
                        b"id" if in_sdt_pr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    let id_str = String::from_utf8_lossy(&attr.value);
                                    current_id = atoi_simd::parse::<u32>(id_str.as_bytes()).ok();
                                }
                            }
                        },
                        b"tag" if in_sdt_pr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    current_tag =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        },
                        b"alias" if in_sdt_pr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    current_title =
                                        Some(String::from_utf8_lossy(&attr.value).into_owned());
                                }
                            }
                        },
                        b"text" if in_sdt_pr => {
                            current_type = Some("text".to_string());
                        },
                        b"date" if in_sdt_pr => {
                            current_type = Some("date".to_string());
                        },
                        b"comboBox" if in_sdt_pr => {
                            current_type = Some("comboBox".to_string());
                        },
                        b"dropDownList" if in_sdt_pr => {
                            current_type = Some("dropDownList".to_string());
                        },
                        b"picture" if in_sdt_pr => {
                            current_type = Some("picture".to_string());
                        },
                        b"lock" if in_sdt_pr => {
                            for attr in e.attributes().flatten() {
                                if attr.key.local_name().as_ref() == b"val" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    match val.as_ref() {
                                        "sdtLocked" => current_lock_delete = true,
                                        "contentLocked" => current_lock_content = true,
                                        "sdtContentLocked" => {
                                            current_lock_delete = true;
                                            current_lock_content = true;
                                        },
                                        _ => {},
                                    }
                                }
                            }
                        },
                        _ => {},
                    }
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"sdtPr" {
                        // End of content control properties
                        if let Some(id) = current_id {
                            controls.push(ContentControl::new(
                                id,
                                current_tag.clone(),
                                current_title.clone(),
                                current_type.clone(),
                                current_lock_delete,
                                current_lock_content,
                            ));
                        }
                        in_sdt_pr = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(controls)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_content_control_creation() {
        let control = ContentControl::new(
            1,
            Some("field1".to_string()),
            Some("My Field".to_string()),
            Some("text".to_string()),
            false,
            false,
        );

        assert_eq!(control.id(), 1);
        assert_eq!(control.tag(), Some("field1"));
        assert_eq!(control.title(), Some("My Field"));
        assert_eq!(control.control_type(), Some("text"));
        assert!(!control.is_lock_delete());
        assert!(!control.is_lock_content());
    }
}
