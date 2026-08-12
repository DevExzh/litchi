#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
use std::borrow::Cow;
use std::ops::Range;

use super::model::{MAX_DOCUMENT_VARIABLE_DEPTH, MAX_DOCUMENT_VARIABLE_XML_BYTES, Variables};
use crate::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

/// Transitional `WordprocessingML` namespace.
const TRANSITIONAL_WORD_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// Strict `WordprocessingML` namespace.
const STRICT_WORD_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SettingsDialect {
    Transitional,
    Strict,
}

#[derive(Clone, Copy, Debug)]
pub(crate) struct SourcePolicy {
    pub(crate) dialect: SettingsDialect,
    pub(crate) protected: bool,
}

impl Variables {
    /// Parse a bounded Word settings XML payload.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        parse_variables(xml)
    }

    /// Serialize a standalone transitional `w:docVars` element.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        const OPEN: &str =
            r#"<w:docVars xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#;
        const CLOSE: &str = "</w:docVars>";
        let capacity = checked_len_add(
            checked_len_add(OPEN.len(), document_variable_entries_len(self, "w")?)?,
            CLOSE.len(),
        )?;
        ensure_output_limit(capacity)?;
        let mut xml = String::new();
        xml.try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "standalone document-variable XML",
                source,
            })?;
        xml.push_str(OPEN);
        self.write_entries(&mut xml, "w");
        xml.push_str(CLOSE);
        debug_assert_eq!(xml.len(), capacity);
        Ok(xml)
    }

    /// Append `docVar` children using the requested XML prefix.
    pub fn write_entries(&self, xml: &mut String, prefix: &str) {
        for (name, value) in self.iter() {
            xml.push('<');
            if !prefix.is_empty() {
                xml.push_str(prefix);
                xml.push(':');
            }
            xml.push_str("docVar");
            xml.push(' ');
            if !prefix.is_empty() {
                xml.push_str(prefix);
                xml.push(':');
            }
            xml.push_str("name=\"");
            escape_attribute(xml, name);
            xml.push_str("\" ");
            if !prefix.is_empty() {
                xml.push_str(prefix);
                xml.push(':');
            }
            xml.push_str("val=\"");
            escape_attribute(xml, value);
            xml.push_str("\"/>");
        }
    }
}

/// Parse document variables from a bounded Word settings XML payload.
///
/// # Errors
///
/// Returns an error if the operation cannot be completed.
pub fn parse_variables(xml: &[u8]) -> Result<Variables> {
    if xml.len() > MAX_DOCUMENT_VARIABLE_XML_BYTES {
        return Err(invalid(format!(
            "settings XML exceeds the {MAX_DOCUMENT_VARIABLE_XML_BYTES} byte document-variable limit"
        )));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut variables = Variables::new();
    let mut names = Vec::new();
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut saw_doc_vars = false;
    let mut doc_vars_depth = None;
    let mut open_doc_var_depth = None;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("document-variable XML nesting overflow"))?;
                if depth > MAX_DOCUMENT_VARIABLE_DEPTH {
                    return Err(invalid(format!(
                        "document-variable XML exceeds depth {MAX_DOCUMENT_VARIABLE_DEPTH}"
                    )));
                }
                if depth == 1 {
                    validate_settings_root(&namespace, &element, saw_root)?;
                    saw_root = true;
                } else if depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"docVars"
                {
                    begin_doc_vars(&mut saw_doc_vars, &mut doc_vars_depth, depth)?;
                } else if depth == 3
                    && doc_vars_depth == Some(2)
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"docVar"
                {
                    push_parsed_variable(&mut variables, &mut names, &element, decoder, &resolver)?;
                    open_doc_var_depth = Some(depth);
                } else if is_wordprocessing_namespace(&namespace)
                    && matches!(element.local_name().as_ref(), b"docVars" | b"docVar")
                {
                    return Err(invalid("misplaced or nested document-variable element"));
                } else if doc_vars_depth.is_some() && is_wordprocessing_namespace(&namespace) {
                    return Err(invalid("unexpected WordprocessingML child in docVars"));
                }
            },
            Event::Empty(element) => {
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("document-variable XML nesting overflow"))?;
                if child_depth > MAX_DOCUMENT_VARIABLE_DEPTH {
                    return Err(invalid(format!(
                        "document-variable XML exceeds depth {MAX_DOCUMENT_VARIABLE_DEPTH}"
                    )));
                }
                if child_depth == 1 {
                    validate_settings_root(&namespace, &element, saw_root)?;
                    saw_root = true;
                } else if child_depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"docVars"
                {
                    begin_doc_vars(&mut saw_doc_vars, &mut doc_vars_depth, child_depth)?;
                    doc_vars_depth = None;
                } else if child_depth == 3
                    && doc_vars_depth == Some(2)
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"docVar"
                {
                    push_parsed_variable(&mut variables, &mut names, &element, decoder, &resolver)?;
                } else if is_wordprocessing_namespace(&namespace)
                    && matches!(element.local_name().as_ref(), b"docVars" | b"docVar")
                {
                    return Err(invalid("misplaced or nested document-variable element"));
                } else if doc_vars_depth.is_some() && is_wordprocessing_namespace(&namespace) {
                    return Err(invalid("unexpected WordprocessingML child in docVars"));
                }
            },
            Event::End(_) => {
                if open_doc_var_depth == Some(depth) {
                    open_doc_var_depth = None;
                }
                if doc_vars_depth == Some(depth) {
                    doc_vars_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid document-variable XML nesting"))?;
            },
            Event::Text(text)
                if (open_doc_var_depth.is_some() || doc_vars_depth == Some(depth))
                    && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(invalid("document-variable elements cannot contain text"));
            },
            Event::CData(text)
                if (open_doc_var_depth.is_some() || doc_vars_depth == Some(depth))
                    && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace()) =>
            {
                return Err(invalid("document-variable elements cannot contain text"));
            },
            Event::Eof if depth != 0 => {
                return Err(invalid("unterminated document-variable settings XML"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !saw_root {
        return Err(invalid("settings part has no settings root"));
    }
    names.sort_unstable();
    if let Some(duplicate) = names.windows(2).find(|pair| pair[0] == pair[1]) {
        return Err(invalid(format!(
            "duplicate document variable name {:?}",
            duplicate[0]
        )));
    }
    Ok(variables)
}

pub(crate) fn inspect_source_policy(xml: &[u8]) -> Result<SourcePolicy> {
    if xml.len() > MAX_DOCUMENT_VARIABLE_XML_BYTES {
        return Err(invalid(format!(
            "settings XML exceeds the {MAX_DOCUMENT_VARIABLE_XML_BYTES} byte document-variable limit"
        )));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut dialect = None;
    let mut document_protection = None;
    let mut write_protection = None;
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("settings XML nesting overflow"))?;
                if depth > MAX_DOCUMENT_VARIABLE_DEPTH {
                    return Err(invalid(format!(
                        "settings XML exceeds depth {MAX_DOCUMENT_VARIABLE_DEPTH}"
                    )));
                }
                if depth == 1 {
                    dialect = Some(settings_root_dialect(&namespace, &element)?);
                } else if depth == 2 && is_wordprocessing_namespace(&namespace) {
                    inspect_protection_element(
                        &element,
                        decoder,
                        &resolver,
                        &mut document_protection,
                        &mut write_protection,
                    )?;
                }
            },
            Event::Empty(element) => {
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("settings XML nesting overflow"))?;
                if child_depth > MAX_DOCUMENT_VARIABLE_DEPTH {
                    return Err(invalid(format!(
                        "settings XML exceeds depth {MAX_DOCUMENT_VARIABLE_DEPTH}"
                    )));
                }
                if child_depth == 1 {
                    dialect = Some(settings_root_dialect(&namespace, &element)?);
                } else if child_depth == 2 && is_wordprocessing_namespace(&namespace) {
                    inspect_protection_element(
                        &element,
                        decoder,
                        &resolver,
                        &mut document_protection,
                        &mut write_protection,
                    )?;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid settings XML nesting"))?;
            },
            Event::Eof if depth != 0 => return Err(invalid("unterminated settings XML")),
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(SourcePolicy {
        dialect: dialect.ok_or_else(|| invalid("settings part has no settings root"))?,
        protected: document_protection.unwrap_or(false) || write_protection.unwrap_or(false),
    })
}

fn settings_root_dialect(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> Result<SettingsDialect> {
    if element.local_name().as_ref() != b"settings" {
        return Err(invalid("document variables require one Word settings root"));
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == TRANSITIONAL_WORD_NAMESPACE => {
            Ok(SettingsDialect::Transitional)
        },
        ResolveResult::Bound(Namespace(value)) if *value == STRICT_WORD_NAMESPACE => {
            Ok(SettingsDialect::Strict)
        },
        _ => Err(invalid("document variables require one Word settings root")),
    }
}

fn inspect_protection_element(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    document: &mut Option<bool>,
    write: &mut Option<bool>,
) -> Result<()> {
    let (slot, attribute, label) = match element.local_name().as_ref() {
        b"documentProtection" => (document, b"enforcement".as_slice(), "documentProtection"),
        b"writeProtection" => (write, b"val".as_slice(), "writeProtection"),
        _ => return Ok(()),
    };
    if slot.is_some() {
        return Err(invalid(format!("duplicate {label} setting")));
    }
    let value = word_attribute_value(element, attribute, decoder, resolver)?;
    *slot = Some(match value.as_deref() {
        None | Some("true" | "1" | "on") => true,
        Some("false" | "0" | "off") => false,
        Some(_) => return Err(invalid(format!("invalid {label} on/off value"))),
    });
    Ok(())
}

/// Rewrite only the typed `docVars` owner in a complete settings source.
pub(super) fn rewrite(xml: &[u8], before: &Variables, after: &Variables) -> Result<Vec<u8>> {
    let current = parse_variables(xml)?;
    if current != *before {
        return Err(invalid("document-variable source is stale"));
    }
    after.validate()?;

    let layout = scan_settings_layout(xml)?;
    let replacement_len = if after.is_empty() {
        0
    } else {
        document_variables_element_len(&layout, after)?
    };
    let output_len = rewritten_settings_len(xml, &layout, replacement_len)?;
    ensure_output_limit(output_len)?;
    let replacement = if replacement_len == 0 {
        Vec::new()
    } else {
        document_variables_element_fallible(&layout, after, replacement_len)?.into_bytes()
    };

    let output = if let Some(range) = layout.doc_vars_range {
        replace_range(xml, range, &replacement)?
    } else if let Some(range) = layout.root_empty_range {
        expand_empty_root(xml, range, &layout.root_qname, &replacement)?
    } else if replacement.is_empty() {
        clone_bytes(xml, "document-variable no-op XML")?
    } else {
        let offset = layout
            .doc_vars_insert_at
            .or(layout.root_end)
            .ok_or_else(|| invalid("settings root has no document-variable insertion point"))?;
        insert_fragment(xml, offset, &replacement)?
    };
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

/// Refuse a source-backed rewrite when replacing `docVars` would discard
/// unmodeled content inside the selected owner.
pub(crate) fn ensure_source_backed_rewrite_safe(xml: &[u8], before: &Variables) -> Result<()> {
    let layout = scan_settings_layout(xml)?;
    let Some(range) = layout.doc_vars_range.as_ref() else {
        return Ok(());
    };
    let capacity = document_variables_element_len(&layout, before)?;
    let canonical = document_variables_element_fallible(&layout, before, capacity)?;
    if xml[range.clone()] != *canonical.as_bytes() {
        return Err(Error::DocumentVariablesPreservation(
            "the selected docVars owner is not in the exact preservation-safe form",
        ));
    }
    Ok(())
}

#[derive(Debug, Default)]
struct SettingsLayout {
    root_qname: Vec<u8>,
    word_prefix: Option<Vec<u8>>,
    strict: bool,
    root_empty_range: Option<Range<usize>>,
    root_end: Option<usize>,
    doc_vars_range: Option<Range<usize>>,
    doc_vars_insert_at: Option<usize>,
}

fn scan_settings_layout(xml: &[u8]) -> Result<SettingsLayout> {
    if xml.len() > MAX_DOCUMENT_VARIABLE_XML_BYTES {
        return Err(invalid(format!(
            "settings XML exceeds the {MAX_DOCUMENT_VARIABLE_XML_BYTES} byte document-variable limit"
        )));
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut layout = SettingsLayout::default();
    let mut depth = 0usize;
    let mut doc_vars_start = None;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_source_error| invalid("document-variable XML offset is too large"))?;
        let event = reader
            .read_event()
            .map_err(|error| xml_error(error.to_string()))?
            .into_owned();
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_source_error| invalid("document-variable XML offset is too large"))?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("document-variable XML nesting overflow"))?;
                if depth > MAX_DOCUMENT_VARIABLE_DEPTH {
                    return Err(invalid(format!(
                        "document-variable XML exceeds depth {MAX_DOCUMENT_VARIABLE_DEPTH}"
                    )));
                }
                if depth == 1 {
                    capture_settings_root(&mut layout, &namespace, &element)?;
                } else if depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"docVars"
                {
                    if doc_vars_start.is_some() || layout.doc_vars_range.is_some() {
                        return Err(invalid("duplicate docVars container"));
                    }
                    doc_vars_start = Some(event_start);
                }
                if depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && layout.doc_vars_insert_at.is_none()
                    && is_after_doc_vars(element.local_name().as_ref())
                {
                    layout.doc_vars_insert_at = Some(event_start);
                }
            },
            Event::Empty(element) => {
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("document-variable XML nesting overflow"))?;
                if child_depth > MAX_DOCUMENT_VARIABLE_DEPTH {
                    return Err(invalid(format!(
                        "document-variable XML exceeds depth {MAX_DOCUMENT_VARIABLE_DEPTH}"
                    )));
                }
                if child_depth == 1 {
                    capture_settings_root(&mut layout, &namespace, &element)?;
                    layout.root_empty_range = Some(event_start..event_end);
                } else if child_depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"docVars"
                {
                    if layout.doc_vars_range.is_some() || doc_vars_start.is_some() {
                        return Err(invalid("duplicate docVars container"));
                    }
                    layout.doc_vars_range = Some(event_start..event_end);
                }
                if child_depth == 2
                    && is_wordprocessing_namespace(&namespace)
                    && layout.doc_vars_insert_at.is_none()
                    && is_after_doc_vars(element.local_name().as_ref())
                {
                    layout.doc_vars_insert_at = Some(event_start);
                }
            },
            Event::End(_) => {
                if depth == 2
                    && let Some(start) = doc_vars_start.take()
                {
                    layout.doc_vars_range = Some(start..event_end);
                }
                if depth == 1 {
                    layout.root_end = Some(event_start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("invalid document-variable XML nesting"))?;
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }

    if depth != 0 {
        return Err(invalid("unterminated document-variable settings XML"));
    }
    if layout.root_qname.is_empty() {
        return Err(invalid("settings root is missing"));
    }
    Ok(layout)
}

fn capture_settings_root(
    layout: &mut SettingsLayout,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> Result<()> {
    if !layout.root_qname.is_empty()
        || !is_wordprocessing_namespace(namespace)
        || element.local_name().as_ref() != b"settings"
    {
        return Err(invalid("document variables require one Word settings root"));
    }
    layout.root_qname = element.name().as_ref().to_vec();
    layout.word_prefix = element
        .name()
        .prefix()
        .map(|prefix| prefix.into_inner().to_vec());
    layout.strict = matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == STRICT_WORD_NAMESPACE
    );
    Ok(())
}

fn is_after_doc_vars(local_name: &[u8]) -> bool {
    matches!(
        local_name,
        b"rsids"
            | b"uiCompat97To2003"
            | b"attachedSchema"
            | b"themeFontLang"
            | b"clrSchemeMapping"
            | b"doNotIncludeSubdocsInStats"
            | b"doNotAutoCompressPictures"
            | b"forceUpgrade"
            | b"mathPr"
            | b"captions"
            | b"readModeInkLockDown"
            | b"smartTagType"
            | b"schemaLibrary"
            | b"shapeDefaults"
            | b"doNotEmbedSmartTags"
            | b"decimalSymbol"
            | b"listSeparator"
    )
}

fn document_variables_element_len(layout: &SettingsLayout, variables: &Variables) -> Result<usize> {
    let prefix = match layout.word_prefix.as_deref() {
        Some(prefix) => std::str::from_utf8(prefix)
            .map(Cow::Borrowed)
            .map_err(|error| xml_error(format!("invalid settings prefix: {error}")))?,
        None => Cow::Borrowed("w"),
    };
    let namespace = if layout.strict {
        STRICT_WORD_NAMESPACE
    } else {
        TRANSITIONAL_WORD_NAMESPACE
    };
    let mut capacity = checked_len_add(1, prefix.len())?;
    capacity = checked_len_add(capacity, ":docVars".len())?;
    if layout.word_prefix.is_none() {
        capacity = checked_len_add(capacity, " xmlns:".len())?;
        capacity = checked_len_add(capacity, prefix.len())?;
        capacity = checked_len_add(capacity, "=\"".len())?;
        capacity = checked_len_add(capacity, namespace.len())?;
        capacity = checked_len_add(capacity, 1)?;
    }
    capacity = checked_len_add(capacity, 1)?;
    for (name, value) in variables.iter() {
        capacity = checked_len_add(capacity, 1 + prefix.len() + ":docVar ".len())?;
        capacity = checked_len_add(capacity, prefix.len() + ":name=\"".len())?;
        capacity = checked_len_add(capacity, escaped_attribute_len(name)?)?;
        capacity = checked_len_add(capacity, "\" ".len())?;
        capacity = checked_len_add(capacity, prefix.len() + ":val=\"".len())?;
        capacity = checked_len_add(capacity, escaped_attribute_len(value)?)?;
        capacity = checked_len_add(capacity, "\"/>".len())?;
    }
    capacity = checked_len_add(capacity, 2 + prefix.len() + ":docVars>".len())?;
    if capacity > MAX_DOCUMENT_VARIABLE_XML_BYTES {
        return Err(invalid(format!(
            "encoded document variables exceed the {MAX_DOCUMENT_VARIABLE_XML_BYTES} byte limit"
        )));
    }
    Ok(capacity)
}

fn document_variables_element_fallible(
    layout: &SettingsLayout,
    variables: &Variables,
    capacity: usize,
) -> Result<String> {
    let prefix = match layout.word_prefix.as_deref() {
        Some(prefix) => std::str::from_utf8(prefix)
            .map(Cow::Borrowed)
            .map_err(|error| xml_error(format!("invalid settings prefix: {error}")))?,
        None => Cow::Borrowed("w"),
    };
    let namespace = if layout.strict {
        STRICT_WORD_NAMESPACE
    } else {
        TRANSITIONAL_WORD_NAMESPACE
    };

    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "encoded document-variable owner",
            source,
        })?;
    output.push('<');
    output.push_str(&prefix);
    output.push_str(":docVars");
    if layout.word_prefix.is_none() {
        output.push_str(" xmlns:");
        output.push_str(&prefix);
        output.push_str("=\"");
        output.push_str(&String::from_utf8_lossy(namespace));
        output.push('"');
    }
    output.push('>');
    for (name, value) in variables.iter() {
        output.push('<');
        output.push_str(&prefix);
        output.push_str(":docVar ");
        output.push_str(&prefix);
        output.push_str(":name=\"");
        escape_attribute(&mut output, name);
        output.push_str("\" ");
        output.push_str(&prefix);
        output.push_str(":val=\"");
        escape_attribute(&mut output, value);
        output.push_str("\"/>");
    }
    output.push_str("</");
    output.push_str(&prefix);
    output.push_str(":docVars>");
    debug_assert_eq!(output.len(), capacity);
    Ok(output)
}

fn document_variable_entries_len(variables: &Variables, prefix: &str) -> Result<usize> {
    let qualified = if prefix.is_empty() {
        0
    } else {
        prefix.len() + 1
    };
    variables
        .iter()
        .try_fold(0usize, |mut capacity, (name, value)| {
            capacity = checked_len_add(capacity, 1 + qualified + "docVar ".len())?;
            capacity = checked_len_add(capacity, qualified + "name=\"".len())?;
            capacity = checked_len_add(capacity, escaped_attribute_len(name)?)?;
            capacity = checked_len_add(capacity, "\" ".len())?;
            capacity = checked_len_add(capacity, qualified + "val=\"".len())?;
            capacity = checked_len_add(capacity, escaped_attribute_len(value)?)?;
            checked_len_add(capacity, "\"/>".len())
        })
}

fn rewritten_settings_len(
    source: &[u8],
    layout: &SettingsLayout,
    replacement_len: usize,
) -> Result<usize> {
    if let Some(range) = &layout.doc_vars_range {
        source
            .len()
            .checked_sub(range.len())
            .and_then(|length| length.checked_add(replacement_len))
    } else if layout.root_empty_range.is_some() {
        source
            .len()
            .checked_add(replacement_len)
            .and_then(|length| length.checked_add(layout.root_qname.len()))
            .and_then(|length| length.checked_add(2))
    } else {
        source.len().checked_add(replacement_len)
    }
    .ok_or_else(|| invalid("rewritten settings XML length overflows usize"))
}

fn replace_range(source: &[u8], range: Range<usize>, replacement: &[u8]) -> Result<Vec<u8>> {
    let capacity = source
        .len()
        .checked_sub(range.len())
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| invalid("rewritten settings XML length overflows usize"))?;
    ensure_output_limit(capacity)?;
    let mut output = reserved_bytes(capacity, "rewritten document-variable settings XML")?;
    output.extend_from_slice(&source[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[range.end..]);
    Ok(output)
}

fn expand_empty_root(
    source: &[u8],
    range: Range<usize>,
    root_qname: &[u8],
    replacement: &[u8],
) -> Result<Vec<u8>> {
    let root = &source[range.clone()];
    let slash = root
        .windows(2)
        .rposition(|window| window == b"/>")
        .ok_or_else(|| invalid("invalid empty settings root"))?;
    let capacity = source
        .len()
        .checked_add(replacement.len())
        .and_then(|length| length.checked_add(root_qname.len()))
        .and_then(|length| length.checked_add(2))
        .ok_or_else(|| invalid("rewritten settings XML length overflows usize"))?;
    ensure_output_limit(capacity)?;
    let mut output = reserved_bytes(capacity, "expanded document-variable settings XML")?;
    output.extend_from_slice(&source[..range.start]);
    output.extend_from_slice(&root[..slash]);
    output.push(b'>');
    output.extend_from_slice(replacement);
    output.extend_from_slice(b"</");
    output.extend_from_slice(root_qname);
    output.push(b'>');
    output.extend_from_slice(&source[range.end..]);
    Ok(output)
}

fn insert_fragment(source: &[u8], offset: usize, replacement: &[u8]) -> Result<Vec<u8>> {
    let capacity = source
        .len()
        .checked_add(replacement.len())
        .ok_or_else(|| invalid("rewritten settings XML length overflows usize"))?;
    ensure_output_limit(capacity)?;
    let mut output = reserved_bytes(capacity, "inserted document-variable settings XML")?;
    output.extend_from_slice(&source[..offset]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[offset..]);
    Ok(output)
}

fn checked_len_add(left: usize, right: usize) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| invalid("encoded document-variable length overflows usize"))
}

fn escaped_attribute_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let encoded = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        checked_len_add(length, encoded)
    })
}

fn ensure_output_limit(length: usize) -> Result<()> {
    if length > MAX_DOCUMENT_VARIABLE_XML_BYTES {
        return Err(invalid(format!(
            "rewritten settings XML exceeds the {MAX_DOCUMENT_VARIABLE_XML_BYTES} byte document-variable limit"
        )));
    }
    Ok(())
}

fn reserved_bytes(capacity: usize, resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation { resource, source })?;
    Ok(output)
}

fn clone_bytes(source: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = reserved_bytes(source.len(), resource)?;
    output.extend_from_slice(source);
    Ok(output)
}

fn validate_settings_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    saw_root: bool,
) -> Result<()> {
    if saw_root
        || !is_wordprocessing_namespace(namespace)
        || element.local_name().as_ref() != b"settings"
    {
        return Err(invalid("document variables require one Word settings root"));
    }
    Ok(())
}

#[cfg(test)]
mod allocation_boundary_tests {
    use super::*;

    #[test]
    fn empty_root_expansion_accepts_exact_limit_and_rejects_one_byte_over() {
        let exact_source_len = MAX_DOCUMENT_VARIABLE_XML_BYTES - 3;
        let mut exact = vec![b'x'; exact_source_len];
        exact[..2].copy_from_slice(b"/>");
        let output = expand_empty_root(&exact, 0..2, b"r", &[]).unwrap();
        assert_eq!(output.len(), MAX_DOCUMENT_VARIABLE_XML_BYTES);

        let mut over = vec![b'x'; exact_source_len + 1];
        over[..2].copy_from_slice(b"/>");
        assert!(expand_empty_root(&over, 0..2, b"r", &[]).is_err());
    }
}

fn begin_doc_vars(
    saw_doc_vars: &mut bool,
    doc_vars_depth: &mut Option<usize>,
    depth: usize,
) -> Result<()> {
    if std::mem::replace(saw_doc_vars, true) {
        return Err(invalid("duplicate docVars container"));
    }
    *doc_vars_depth = Some(depth);
    Ok(())
}

fn push_parsed_variable(
    variables: &mut Variables,
    names: &mut Vec<String>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let name = word_attribute_value(element, b"name", decoder, resolver)?
        .ok_or_else(|| invalid("document variable name attribute is required"))?;
    let value = word_attribute_value(element, b"val", decoder, resolver)?
        .ok_or_else(|| invalid("document variable val attribute is required"))?;
    let mut indexed_name = String::new();
    indexed_name
        .try_reserve_exact(name.len())
        .map_err(|source| Error::Allocation {
            resource: "document-variable duplicate-name index",
            source,
        })?;
    indexed_name.push_str(&name);
    names.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "document-variable duplicate-name index",
        source,
    })?;
    names.push(indexed_name);
    variables.push_parsed(name, value)
}

fn is_wordprocessing_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == TRANSITIONAL_WORD_NAMESPACE || *value == STRICT_WORD_NAMESPACE
    )
}

fn word_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| xml_error(error.to_string()))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let is_word_attribute = is_wordprocessing_namespace(&namespace)
            || matches!(namespace, ResolveResult::Unbound)
            || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"w");
        if !is_word_attribute {
            continue;
        }
        if value.is_some() {
            return Err(invalid(format!(
                "duplicate Word attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| xml_error(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            _ => output.push(character),
        }
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(message: impl Into<String>) -> Error {
    Error::Xml(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn serializes_and_parses_escaped_values() {
        let mut variables = Variables::new();
        variables
            .insert("Company & Team", "A < B && C \"yes\"")
            .unwrap();
        variables.insert("empty", "").unwrap();
        let xml = variables.to_xml().unwrap();
        let reparsed = parse_variables(
            format!(
                r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docVars>{}</w:docVars></w:settings>"#,
                xml.trim_start_matches("<w:docVars xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">")
                    .trim_end_matches("</w:docVars>")
            )
            .as_bytes(),
        )
        .unwrap();
        assert_eq!(reparsed.get("Company & Team"), Some("A < B && C \"yes\""));
        assert_eq!(reparsed.get("empty"), Some(""));
    }

    #[test]
    fn rejects_missing_duplicate_nested_and_oversized_input() {
        let wrap = |content: &str| {
            format!(
                r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{content}</w:settings>"#
            )
        };
        for content in [
            r#"<w:docVars><w:docVar w:val="x"/></w:docVars>"#,
            r#"<w:docVars><w:docVar w:name="x"/></w:docVars>"#,
            r#"<w:docVars/><w:docVars/>"#,
            r#"<w:docVars><w:docVar w:name="x" w:val="1"/><w:docVar w:name="x" w:val="2"/></w:docVars>"#,
            r#"<w:docVars><w:docVar w:name="x" w:val="1"><w:docVar w:name="y" w:val="2"/></w:docVar></w:docVars>"#,
            r#"<w:docVar w:name="x" w:val="1"/>"#,
        ] {
            assert!(parse_variables(wrap(content).as_bytes()).is_err());
        }
        let oversized = vec![b' '; MAX_DOCUMENT_VARIABLE_XML_BYTES + 1];
        assert!(parse_variables(&oversized).is_err());
    }

    #[test]
    fn source_backed_rewrite_refuses_selected_namespace_and_entity_syntax() {
        let namespace = format!(
            r#"<w:settings xmlns:w="{}"><w:docVars xmlns:x="urn:selected"><w:docVar w:name="name" w:val="value"/></w:docVars></w:settings>"#,
            String::from_utf8_lossy(TRANSITIONAL_WORD_NAMESPACE)
        );
        let variables = parse_variables(namespace.as_bytes()).unwrap();
        assert!(ensure_source_backed_rewrite_safe(namespace.as_bytes(), &variables).is_err());

        let entity = format!(
            r#"<w:settings xmlns:w="{}"><w:docVars><w:docVar w:name="name" w:val="value"/>&opaque;</w:docVars></w:settings>"#,
            String::from_utf8_lossy(TRANSITIONAL_WORD_NAMESPACE)
        );
        let variables = parse_variables(entity.as_bytes()).unwrap();
        assert!(ensure_source_backed_rewrite_safe(entity.as_bytes(), &variables).is_err());
    }
}
