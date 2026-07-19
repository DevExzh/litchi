/// Content control support for Word documents.
///
/// Content controls are structured regions in a document that can contain
/// specific types of content (text, dates, lists, etc.).
use crate::docx::namespace::{is_wordprocessing_namespace, word_attribute_value};
use crate::error::{OoxmlError, Result};
use crate::custom_xml_data::is_st_guid;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const WORD_2010_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/word/2010/wordml";
const WORD_2012_NAMESPACE: &[u8] = b"http://schemas.microsoft.com/office/word/2012/wordml";

/// A content control in a Word document.
///
/// Content controls provide structured content regions that can be
/// bound to data or restricted to specific content types.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::docx::Package;
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
    /// Whether Word should remove the control after its contents are edited.
    temporary: bool,
    /// Whether the control is currently displaying placeholder content.
    showing_placeholder: bool,
    /// Building-block name used for the placeholder.
    placeholder: Option<String>,
    /// Keyboard tab order.
    tab_index: Option<u32>,
    /// XPath of the custom XML data binding.
    data_binding_xpath: Option<String>,
    /// Custom XML data-store item identifier.
    data_binding_store_item_id: Option<String>,
    /// Namespace prefix mappings used by the data-binding XPath.
    data_binding_prefix_mappings: Option<String>,
    /// Display text and values declared by combo-box or drop-down controls.
    list_items: Vec<(String, String)>,
    /// Checked state for a Word 2010 checkbox control.
    checked: Option<bool>,
    /// Display format for a date control.
    date_format: Option<String>,
    /// ISO date value stored on a date control.
    date_value: Option<String>,
    /// Title of a Word 2012 repeating section.
    repeating_section_title: Option<String>,
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
            temporary: false,
            showing_placeholder: false,
            placeholder: None,
            tab_index: None,
            data_binding_xpath: None,
            data_binding_store_item_id: None,
            data_binding_prefix_mappings: None,
            list_items: Vec::new(),
            checked: None,
            date_format: None,
            date_value: None,
            repeating_section_title: None,
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

    /// Check whether Word should remove the control after it is edited.
    #[inline]
    pub fn is_temporary(&self) -> bool {
        self.temporary
    }

    /// Check whether the control is displaying placeholder content.
    #[inline]
    pub fn is_showing_placeholder(&self) -> bool {
        self.showing_placeholder
    }

    /// Get the building-block name used for placeholder content.
    #[inline]
    pub fn placeholder(&self) -> Option<&str> {
        self.placeholder.as_deref()
    }

    /// Get the keyboard tab order.
    #[inline]
    pub fn tab_index(&self) -> Option<u32> {
        self.tab_index
    }

    /// Get the XPath of the custom XML data binding.
    #[inline]
    pub fn data_binding_xpath(&self) -> Option<&str> {
        self.data_binding_xpath.as_deref()
    }

    /// Get the custom XML data-store item identifier.
    #[inline]
    pub fn data_binding_store_item_id(&self) -> Option<&str> {
        self.data_binding_store_item_id.as_deref()
    }

    /// Get namespace prefix mappings used by the data-binding XPath.
    #[inline]
    pub fn data_binding_prefix_mappings(&self) -> Option<&str> {
        self.data_binding_prefix_mappings.as_deref()
    }

    /// Validate binding metadata without evaluating the XPath or resolving URIs.
    pub fn validate_data_binding(&self) -> Result<()> {
        match (
            self.data_binding_xpath.as_deref(),
            self.data_binding_store_item_id.as_deref(),
        ) {
            (None, None) => Ok(()),
            (Some(xpath), Some(store_item_id)) => validate_data_binding_values(
                xpath,
                store_item_id,
                self.data_binding_prefix_mappings.as_deref(),
            ),
            _ => Err(OoxmlError::InvalidFormat(
                "content-control data binding is incomplete".to_string(),
            )),
        }
    }

    /// Get the display text and values declared by a list control.
    #[inline]
    pub fn list_items(&self) -> &[(String, String)] {
        &self.list_items
    }

    /// Get the checked state of a checkbox control.
    #[inline]
    pub fn checked(&self) -> Option<bool> {
        self.checked
    }

    /// Get the display format of a date control.
    #[inline]
    pub fn date_format(&self) -> Option<&str> {
        self.date_format.as_deref()
    }

    /// Get the ISO date value stored on a date control.
    #[inline]
    pub fn date_value(&self) -> Option<&str> {
        self.date_value.as_deref()
    }

    /// Get the title of a repeating-section control.
    #[inline]
    pub fn repeating_section_title(&self) -> Option<&str> {
        self.repeating_section_title.as_deref()
    }

    /// Extract content controls from document XML bytes.
    pub(crate) fn extract_from_document(doc_xml: &[u8]) -> Result<Vec<ContentControl>> {
        let mut reader = NsReader::from_reader(doc_xml);
        let mut controls = Vec::new();
        let mut pending = None;
        let mut depth = 0usize;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "content-control XML nesting is too deep".to_string(),
                        )
                    })?;
                    if is_word_element(&namespace, &element, b"sdtPr") {
                        if pending.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "nested content-control properties are invalid".to_string(),
                            ));
                        }
                        pending = Some(PendingContentControl::new(depth));
                    } else if let Some(control) = pending.as_mut() {
                        parse_property(
                            &namespace, &element, decoder, &resolver, depth, false, control,
                        )?;
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "content-control XML nesting is too deep".to_string(),
                        )
                    })?;
                    if is_word_element(&namespace, &element, b"sdtPr") {
                        if pending.is_some() {
                            return Err(OoxmlError::InvalidFormat(
                                "nested content-control properties are invalid".to_string(),
                            ));
                        }
                        // An empty sdtPr has no ID and cannot be represented by this legacy API.
                    } else if let Some(control) = pending.as_mut() {
                        parse_property(
                            &namespace,
                            &element,
                            decoder,
                            &resolver,
                            child_depth,
                            true,
                            control,
                        )?;
                    }
                },
                Event::End(element) => {
                    if pending.as_ref().is_some_and(|control| {
                        control.depth == depth
                            && is_wordprocessing_namespace(&namespace)
                            && element.local_name().as_ref() == b"sdtPr"
                    }) {
                        let control = pending.take().ok_or_else(|| {
                            OoxmlError::InvalidFormat(
                                "missing content-control properties".to_string(),
                            )
                        })?;
                        if let Some(control) = control.finish()? {
                            if controls
                                .iter()
                                .any(|existing: &ContentControl| existing.id == control.id)
                            {
                                return Err(OoxmlError::InvalidFormat(format!(
                                    "duplicate content-control ID {}",
                                    control.id
                                )));
                            }
                            controls.push(control);
                        }
                    } else if let Some(control) = pending.as_mut()
                        && control
                            .context
                            .is_some_and(|(_, context_depth)| context_depth == depth)
                    {
                        control.context = None;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid content-control XML nesting".to_string())
                    })?;
                },
                Event::Eof if pending.is_some() => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated content-control properties".to_string(),
                    ));
                },
                Event::Eof if depth != 0 => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated document XML".to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        Ok(controls)
    }
}

/// Validate the lexical form of an SDT data binding without executing XPath.
pub fn validate_data_binding_values(
    xpath: &str,
    store_item_id: &str,
    prefix_mappings: Option<&str>,
) -> Result<()> {
    const MAX_BINDING_BYTES: usize = 64 * 1024;
    if xpath.is_empty()
        || xpath.len() > MAX_BINDING_BYTES
        || xpath.chars().any(|character| character == '\0' || character.is_control() && !character.is_whitespace())
    {
        return Err(OoxmlError::InvalidFormat(
            "content-control XPath is empty or exceeds lexical limits".to_string(),
        ));
    }
    if !is_st_guid(store_item_id) {
        return Err(OoxmlError::InvalidFormat(format!(
            "content-control storeItemID '{store_item_id}' is not ST_Guid"
        )));
    }
    let Some(mut remaining) = prefix_mappings else {
        return Ok(());
    };
    if remaining.len() > MAX_BINDING_BYTES {
        return Err(OoxmlError::InvalidFormat(
            "content-control prefixMappings exceeds lexical limits".to_string(),
        ));
    }
    let mut prefixes = std::collections::HashSet::new();
    while !remaining.trim_start().is_empty() {
        remaining = remaining.trim_start();
        let after_xmlns = remaining.strip_prefix("xmlns").ok_or_else(|| {
            OoxmlError::InvalidFormat("prefixMappings requires xmlns declarations".to_string())
        })?;
        let (prefix, after_prefix) = if let Some(after_colon) = after_xmlns.strip_prefix(':') {
            let end = after_colon.find('=').ok_or_else(|| {
                OoxmlError::InvalidFormat("prefixMappings declaration has no '='".to_string())
            })?;
            let prefix = &after_colon[..end];
            if prefix.is_empty()
                || !prefix.bytes().enumerate().all(|(index, byte)| {
                    byte.is_ascii_alphabetic()
                        || byte == b'_'
                        || index > 0 && (byte.is_ascii_digit() || matches!(byte, b'.' | b'-'))
                })
            {
                return Err(OoxmlError::InvalidFormat(
                    "prefixMappings contains an invalid namespace prefix".to_string(),
                ));
            }
            (prefix, &after_colon[end..])
        } else {
            ("", after_xmlns)
        };
        if !prefixes.insert(prefix.to_string()) {
            return Err(OoxmlError::InvalidFormat(
                "prefixMappings contains a duplicate namespace prefix".to_string(),
            ));
        }
        let after_equals = after_prefix.strip_prefix('=').ok_or_else(|| {
            OoxmlError::InvalidFormat("prefixMappings declaration has no '='".to_string())
        })?;
        let quote = after_equals.as_bytes().first().copied().ok_or_else(|| {
            OoxmlError::InvalidFormat("prefixMappings declaration has no URI".to_string())
        })?;
        if !matches!(quote, b'\'' | b'"') {
            return Err(OoxmlError::InvalidFormat(
                "prefixMappings URI must be quoted".to_string(),
            ));
        }
        let quoted = &after_equals[1..];
        let end = quoted.find(quote as char).ok_or_else(|| {
            OoxmlError::InvalidFormat("prefixMappings URI quote is not closed".to_string())
        })?;
        if quoted[..end].is_empty() {
            return Err(OoxmlError::InvalidFormat(
                "prefixMappings namespace URI must not be empty".to_string(),
            ));
        }
        remaining = &quoted[end + 1..];
        if !remaining.is_empty()
            && !remaining
                .as_bytes()
                .first()
                .is_some_and(u8::is_ascii_whitespace)
        {
            return Err(OoxmlError::InvalidFormat(
                "prefixMappings declarations must be whitespace separated".to_string(),
            ));
        }
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum PropertyContext {
    Placeholder,
    Date,
    List,
    Checkbox,
    RepeatingSection,
}

struct PendingContentControl {
    depth: usize,
    id: Option<u32>,
    tag: Option<String>,
    title: Option<String>,
    control_type: Option<String>,
    lock: Option<(bool, bool)>,
    temporary: Option<bool>,
    showing_placeholder: Option<bool>,
    placeholder: Option<String>,
    placeholder_seen: bool,
    tab_index: Option<u32>,
    data_binding_xpath: Option<String>,
    data_binding_store_item_id: Option<String>,
    data_binding_prefix_mappings: Option<String>,
    list_items: Vec<(String, String)>,
    checked: Option<bool>,
    date_format: Option<String>,
    date_value: Option<String>,
    repeating_section_title: Option<String>,
    context: Option<(PropertyContext, usize)>,
}

impl PendingContentControl {
    fn new(depth: usize) -> Self {
        Self {
            depth,
            id: None,
            tag: None,
            title: None,
            control_type: None,
            lock: None,
            temporary: None,
            showing_placeholder: None,
            placeholder: None,
            placeholder_seen: false,
            tab_index: None,
            data_binding_xpath: None,
            data_binding_store_item_id: None,
            data_binding_prefix_mappings: None,
            list_items: Vec::new(),
            checked: None,
            date_format: None,
            date_value: None,
            repeating_section_title: None,
            context: None,
        }
    }

    fn finish(self) -> Result<Option<ContentControl>> {
        let Some(id) = self.id else {
            return Ok(None);
        };
        let (lock_delete, lock_content) = self.lock.unwrap_or((false, false));
        Ok(Some(ContentControl {
            id,
            tag: self.tag,
            title: self.title,
            control_type: Some(self.control_type.unwrap_or_else(|| "richText".to_string())),
            lock_delete,
            lock_content,
            temporary: self.temporary.unwrap_or(false),
            showing_placeholder: self.showing_placeholder.unwrap_or(false),
            placeholder: self.placeholder,
            tab_index: self.tab_index,
            data_binding_xpath: self.data_binding_xpath,
            data_binding_store_item_id: self.data_binding_store_item_id,
            data_binding_prefix_mappings: self.data_binding_prefix_mappings,
            list_items: self.list_items,
            checked: self.checked,
            date_format: self.date_format,
            date_value: self.date_value,
            repeating_section_title: self.repeating_section_title,
        }))
    }
}

fn parse_property(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    empty: bool,
    control: &mut PendingContentControl,
) -> Result<()> {
    if depth == control.depth + 1 {
        parse_direct_property(namespace, element, decoder, resolver, depth, empty, control)
    } else if depth == control.depth + 2 {
        parse_nested_property(namespace, element, decoder, resolver, control)
    } else {
        Ok(())
    }
}

fn parse_direct_property(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    empty: bool,
    control: &mut PendingContentControl,
) -> Result<()> {
    let name = element.local_name();
    if is_wordprocessing_namespace(namespace) {
        match name.as_ref() {
            b"id" => {
                set_once(
                    &mut control.id,
                    required_u32(element, b"val", decoder, resolver, "content-control ID")?,
                    "content-control ID",
                )?;
            },
            b"tag" => {
                let value = required_word_attribute(element, b"val", decoder, resolver, "tag")?;
                set_once(&mut control.tag, value, "content-control tag")?;
            },
            b"alias" => {
                let value = required_word_attribute(element, b"val", decoder, resolver, "alias")?;
                set_once(&mut control.title, value, "content-control alias")?;
            },
            b"lock" => parse_lock(element, decoder, resolver, control)?,
            b"temporary" => {
                let value = parse_on_off(element, decoder, resolver)?;
                set_once(
                    &mut control.temporary,
                    value,
                    "content-control temporary property",
                )?;
            },
            b"showingPlcHdr" => {
                let value = parse_on_off(element, decoder, resolver)?;
                set_once(
                    &mut control.showing_placeholder,
                    value,
                    "content-control placeholder-display property",
                )?;
            },
            b"tabIndex" => {
                let value = required_u32(element, b"val", decoder, resolver, "tab index")?;
                set_once(&mut control.tab_index, value, "content-control tab index")?;
            },
            b"dataBinding" => parse_data_binding(element, decoder, resolver, control)?,
            b"placeholder" => {
                if control.placeholder_seen {
                    return Err(OoxmlError::InvalidFormat(
                        "duplicate content-control placeholder".to_string(),
                    ));
                }
                control.placeholder_seen = true;
                set_context(control, PropertyContext::Placeholder, depth, empty);
            },
            b"date" => {
                set_control_type(control, "date")?;
                control.date_value = word_attribute_value(element, b"fullDate", decoder, resolver)?;
                set_context(control, PropertyContext::Date, depth, empty);
            },
            b"comboBox" => {
                set_control_type(control, "comboBox")?;
                set_context(control, PropertyContext::List, depth, empty);
            },
            b"dropDownList" => {
                set_control_type(control, "dropDownList")?;
                set_context(control, PropertyContext::List, depth, empty);
            },
            b"text" | b"picture" | b"citation" | b"equation" | b"group" | b"docPartList"
            | b"docPartObj" | b"bibliography" | b"richText" => {
                set_control_type(
                    control,
                    std::str::from_utf8(name.as_ref()).map_err(|error| {
                        OoxmlError::InvalidFormat(format!("invalid content-control type: {error}"))
                    })?,
                )?;
            },
            _ => {},
        }
    } else if is_extension_namespace(namespace, WORD_2010_NAMESPACE) {
        match name.as_ref() {
            b"checkbox" => {
                set_control_type(control, "checkbox")?;
                set_context(control, PropertyContext::Checkbox, depth, empty);
            },
            b"entityPicker" => set_control_type(control, "entityPicker")?,
            _ => {},
        }
    } else if is_extension_namespace(namespace, WORD_2012_NAMESPACE) {
        match name.as_ref() {
            b"repeatingSection" => {
                set_control_type(control, "repeatingSection")?;
                set_context(control, PropertyContext::RepeatingSection, depth, empty);
            },
            b"repeatingSectionItem" => set_control_type(control, "repeatingSectionItem")?,
            _ => {},
        }
    }
    Ok(())
}

fn parse_nested_property(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    control: &mut PendingContentControl,
) -> Result<()> {
    let Some((context, _)) = control.context else {
        return Ok(());
    };
    let name = element.local_name();
    match context {
        PropertyContext::Placeholder if is_word_element(namespace, element, b"docPart") => {
            let value = required_word_attribute(element, b"val", decoder, resolver, "placeholder")?;
            set_once(
                &mut control.placeholder,
                value,
                "content-control placeholder",
            )?;
        },
        PropertyContext::Date if is_word_element(namespace, element, b"dateFormat") => {
            let value = required_word_attribute(element, b"val", decoder, resolver, "date format")?;
            set_once(
                &mut control.date_format,
                value,
                "content-control date format",
            )?;
        },
        PropertyContext::List if is_word_element(namespace, element, b"listItem") => {
            let value = required_word_attribute(element, b"value", decoder, resolver, "list item")?;
            let display = word_attribute_value(element, b"displayText", decoder, resolver)?
                .unwrap_or_else(|| value.clone());
            control.list_items.push((display, value));
        },
        PropertyContext::Checkbox
            if is_extension_namespace(namespace, WORD_2010_NAMESPACE)
                && name.as_ref() == b"checked" =>
        {
            let value = extension_attribute_value(
                element,
                b"val",
                WORD_2010_NAMESPACE,
                b"w14",
                decoder,
                resolver,
            )?;
            set_once(
                &mut control.checked,
                value.as_deref().map_or(Ok(true), parse_on_off_value)?,
                "checkbox state",
            )?;
        },
        PropertyContext::RepeatingSection
            if is_extension_namespace(namespace, WORD_2012_NAMESPACE)
                && name.as_ref() == b"sectionTitle" =>
        {
            let value =
                required_word_attribute(element, b"val", decoder, resolver, "section title")?;
            set_once(
                &mut control.repeating_section_title,
                value,
                "repeating-section title",
            )?;
        },
        _ => {},
    }
    Ok(())
}

fn parse_lock(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    control: &mut PendingContentControl,
) -> Result<()> {
    if control.lock.is_some() {
        return Err(OoxmlError::InvalidFormat(
            "duplicate content-control lock".to_string(),
        ));
    }
    let value = required_word_attribute(element, b"val", decoder, resolver, "lock")?;
    control.lock = Some(match value.as_str() {
        "unlocked" => (false, false),
        "sdtLocked" => (true, false),
        "contentLocked" => (false, true),
        "sdtContentLocked" => (true, true),
        _ => {
            return Err(OoxmlError::InvalidFormat(format!(
                "invalid content-control lock value '{value}'"
            )));
        },
    });
    Ok(())
}

fn parse_data_binding(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    control: &mut PendingContentControl,
) -> Result<()> {
    if control.data_binding_xpath.is_some() || control.data_binding_store_item_id.is_some() {
        return Err(OoxmlError::InvalidFormat(
            "duplicate content-control data binding".to_string(),
        ));
    }
    control.data_binding_xpath = Some(required_word_attribute(
        element,
        b"xpath",
        decoder,
        resolver,
        "data-binding XPath",
    )?);
    control.data_binding_store_item_id = Some(required_word_attribute(
        element,
        b"storeItemID",
        decoder,
        resolver,
        "data-binding store item ID",
    )?);
    control.data_binding_prefix_mappings =
        word_attribute_value(element, b"prefixMappings", decoder, resolver)?;
    Ok(())
}

fn set_context(
    control: &mut PendingContentControl,
    context: PropertyContext,
    depth: usize,
    empty: bool,
) {
    if !empty {
        control.context = Some((context, depth));
    }
}

fn set_control_type(control: &mut PendingContentControl, value: &str) -> Result<()> {
    if let Some(existing) = &control.control_type {
        return Err(OoxmlError::InvalidFormat(format!(
            "multiple content-control types '{existing}' and '{value}'"
        )));
    }
    control.control_type = Some(value.to_string());
    Ok(())
}

fn set_once<T>(target: &mut Option<T>, value: T, description: &str) -> Result<()> {
    if target.is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "duplicate {description}"
        )));
    }
    *target = Some(value);
    Ok(())
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<u32> {
    let value = required_word_attribute(element, name, decoder, resolver, description)?;
    value
        .parse::<u32>()
        .map_err(|_| OoxmlError::InvalidFormat(format!("invalid {description} value '{value}'")))
}

fn required_word_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<String> {
    word_attribute_value(element, name, decoder, resolver)?
        .ok_or_else(|| OoxmlError::InvalidFormat(format!("missing {description} attribute")))
}

fn parse_on_off(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<bool> {
    match word_attribute_value(element, b"val", decoder, resolver)? {
        Some(value) => parse_on_off_value(&value),
        None => Ok(true),
    }
}

fn parse_on_off_value(value: &str) -> Result<bool> {
    match value {
        "true" | "1" | "on" => Ok(true),
        "false" | "0" | "off" => Ok(false),
        _ => Err(OoxmlError::InvalidFormat(format!(
            "invalid on/off value '{value}'"
        ))),
    }
}

fn extension_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    namespace: &[u8],
    conventional_prefix: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (resolved, _) = resolver.resolve_attribute(attribute.key);
        let matches = matches!(resolved, ResolveResult::Bound(Namespace(value)) if value == namespace)
            || matches!(resolved, ResolveResult::Unknown(prefix) if prefix.as_slice() == conventional_prefix);
        if !matches {
            continue;
        }
        if value.is_some() {
            return Err(OoxmlError::InvalidFormat(format!(
                "duplicate extension attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

fn is_word_element(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, name: &[u8]) -> bool {
    is_wordprocessing_namespace(namespace) && element.local_name().as_ref() == name
}

fn is_extension_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

#[cfg(test)]
mod tests {
    use super::*;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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

    #[test]
    fn extracts_namespaced_properties_and_decodes_values() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}"><w:body><w:sdt><w:sdtPr>
                <w:id w:val="42"/><w:tag w:val="customer&amp;id"/>
                <w:alias w:val="Customer &amp; address"/><w:tabIndex w:val="7"/>
                <w:temporary/><w:showingPlcHdr w:val="on"/>
                <w:placeholder><w:docPart w:val="DefaultPlaceholder_1"/></w:placeholder>
                <w:dataBinding w:prefixMappings="xmlns:x='urn:test&amp;more'"
                    w:xpath="/x:root/x:name" w:storeItemID="{{ABC}}"/>
                <w:dropDownList><w:listItem w:displayText="A &amp; B" w:value="ab"/>
                    <w:listItem w:value="fallback"/></w:dropDownList>
                <w:lock w:val="sdtContentLocked"/>
            </w:sdtPr><w:sdtContent/></w:sdt></w:body></w:document>"#
        );
        let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        let control = &controls[0];
        assert_eq!(control.id(), 42);
        assert_eq!(control.tag(), Some("customer&id"));
        assert_eq!(control.title(), Some("Customer & address"));
        assert_eq!(control.control_type(), Some("dropDownList"));
        assert_eq!(control.tab_index(), Some(7));
        assert!(control.is_temporary());
        assert!(control.is_showing_placeholder());
        assert_eq!(control.placeholder(), Some("DefaultPlaceholder_1"));
        assert_eq!(control.data_binding_xpath(), Some("/x:root/x:name"));
        assert_eq!(control.data_binding_store_item_id(), Some("{ABC}"));
        assert_eq!(
            control.data_binding_prefix_mappings(),
            Some("xmlns:x='urn:test&more'")
        );
        assert_eq!(
            control.list_items(),
            &[
                ("A & B".into(), "ab".into()),
                ("fallback".into(), "fallback".into())
            ]
        );
        assert!(control.is_lock_delete());
        assert!(control.is_lock_content());
    }

    #[test]
    fn accepts_strict_and_aliased_word_namespaces() {
        let xml = r#"<x:document xmlns:x="http://purl.oclc.org/ooxml/wordprocessingml/main">
            <x:sdtPr><x:id x:val="1"/><x:text/></x:sdtPr>
            <x:sdtPr><x:id x:val="2"/></x:sdtPr></x:document>"#;
        let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        assert_eq!(controls.len(), 2);
        assert_eq!(controls[0].control_type(), Some("text"));
        assert_eq!(controls[1].control_type(), Some("richText"));
    }

    #[test]
    fn extracts_checkbox_date_and_repeating_section_metadata() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}"
                xmlns:c="http://schemas.microsoft.com/office/word/2010/wordml"
                xmlns:r="http://schemas.microsoft.com/office/word/2012/wordml">
                <w:sdtPr><w:id w:val="1"/><c:checkbox><c:checked/></c:checkbox></w:sdtPr>
                <w:sdtPr><w:id w:val="2"/><w:date w:fullDate="2026-07-14T00:00:00Z">
                    <w:dateFormat w:val="yyyy-MM-dd"/></w:date></w:sdtPr>
                <w:sdtPr><w:id w:val="3"/><r:repeatingSection>
                    <r:sectionTitle w:val="People"/></r:repeatingSection></w:sdtPr>
            </w:document>"#
        );
        let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        assert_eq!(controls[0].control_type(), Some("checkbox"));
        assert_eq!(controls[0].checked(), Some(true));
        assert_eq!(controls[1].control_type(), Some("date"));
        assert_eq!(controls[1].date_value(), Some("2026-07-14T00:00:00Z"));
        assert_eq!(controls[1].date_format(), Some("yyyy-MM-dd"));
        assert_eq!(controls[2].control_type(), Some("repeatingSection"));
        assert_eq!(controls[2].repeating_section_title(), Some("People"));
    }

    #[test]
    fn ignores_foreign_lookalikes_and_idless_controls() {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:f="urn:foreign">
                <f:sdtPr><f:id f:val="8"/><f:text/></f:sdtPr>
                <w:sdtPr><f:id f:val="9"/><f:tag f:val="spoof"/></w:sdtPr>
                <w:sdtPr/><w:sdtPr><w:id w:val="10"/><f:text/></w:sdtPr>
            </w:document>"#
        );
        let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].id(), 10);
        assert_eq!(controls[0].control_type(), Some("richText"));
    }

    #[test]
    fn recognizes_all_standard_type_markers() {
        let types = [
            "text",
            "picture",
            "citation",
            "equation",
            "group",
            "docPartList",
            "docPartObj",
            "bibliography",
            "richText",
            "comboBox",
        ];
        for (index, control_type) in types.iter().enumerate() {
            let xml = format!(
                r#"<w:sdtPr xmlns:w="{W}"><w:id w:val="{}"/><w:{control_type}/></w:sdtPr>"#,
                index + 1
            );
            let controls = ContentControl::extract_from_document(xml.as_bytes()).unwrap();
            assert_eq!(controls[0].control_type(), Some(*control_type));
        }
    }

    #[test]
    fn rejects_invalid_or_duplicate_properties() {
        let invalid = [
            r#"<w:sdtPr xmlns:w="W"><w:id/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="x"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:id w:val="2"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:lock w:val="invalid"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:text/><w:picture/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:temporary w:val="maybe"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:temporary/><w:temporary w:val="0"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:lock w:val="unlocked"/><w:lock w:val="unlocked"/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:placeholder/><w:placeholder/></w:sdtPr>"#,
            r#"<w:sdtPr xmlns:w="W"><w:id w:val="1"/><w:dataBinding w:xpath="/x"/></w:sdtPr>"#,
        ];
        for xml in invalid {
            let xml = xml.replace("xmlns:w=\"W\"", &format!("xmlns:w=\"{W}\""));
            assert!(ContentControl::extract_from_document(xml.as_bytes()).is_err());
        }

        let duplicate = format!(
            r#"<w:document xmlns:w="{W}"><w:sdtPr><w:id w:val="1"/></w:sdtPr>
                <w:sdtPr><w:id w:val="1"/></w:sdtPr></w:document>"#
        );
        assert!(ContentControl::extract_from_document(duplicate.as_bytes()).is_err());
    }

    #[test]
    fn rejects_truncated_properties() {
        let xml = format!(r#"<w:document xmlns:w="{W}"><w:sdtPr><w:id w:val="1"/>"#);
        assert!(ContentControl::extract_from_document(xml.as_bytes()).is_err());
    }
}
