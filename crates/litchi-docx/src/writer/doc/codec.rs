#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
use crate::error::{Error, Result};
use quick_xml::Reader;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use std::fmt::Write as FmtWrite;

use super::model::{MutableDocument, Protection};

impl MutableDocument {
    pub(super) fn write_document_prefix(&self, xml: &mut String) {
        if let Some(prefix) = &self.preserved_prefix {
            xml.push_str(prefix);
        } else {
            xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
            xml.push_str(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>"#);
        }
    }

    pub(super) fn write_document_suffix(&self, xml: &mut String) {
        if let Some(suffix) = &self.preserved_suffix {
            xml.push_str(suffix);
        } else {
            xml.push_str("</w:body></w:document>");
        }
    }

    /// Generate section properties XML including header/footer/footnote/endnote references.
    pub(super) fn generate_section_properties(
        &self,
        xml: &mut String,
        rel_mapper: &super::super::relmap::RelationshipMapper,
    ) -> Result<()> {
        self.section.write_xml(xml, Some(rel_mapper))
    }
}

pub(super) fn compact_changed_document_xml(source: &str) -> Result<String> {
    let mut reader = Reader::from_str(source);
    reader.config_mut().trim_text(false);
    let mut output = String::with_capacity(source.len());
    let mut preserve_space = Vec::new();
    let mut pending_whitespace = String::new();
    let mut text_run_has_content = false;
    let mut roots = 0_usize;

    loop {
        match reader.read_event().map_err(Error::from)? {
            Event::Start(element) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                if preserve_space.is_empty() {
                    roots = roots.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("document XML root count overflowed".into())
                    })?;
                }
                let inherited = preserve_space.last().copied().unwrap_or(false);
                preserve_space.push(element_preserves_space(
                    &element,
                    reader.decoder(),
                    inherited,
                )?);
                write_compact_start(&mut output, &element, reader.decoder(), false)?;
            },
            Event::Empty(element) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                if preserve_space.is_empty() {
                    roots = roots.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("document XML root count overflowed".into())
                    })?;
                }
                let inherited = preserve_space.last().copied().unwrap_or(false);
                let _ = element_preserves_space(&element, reader.decoder(), inherited)?;
                write_compact_start(&mut output, &element, reader.decoder(), true)?;
            },
            Event::End(element) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                preserve_space.pop().ok_or_else(|| {
                    Error::InvalidFormat("unexpected document XML end element".into())
                })?;
                output.push_str("</");
                push_utf8(&mut output, element.name().as_ref())?;
                output.push('>');
            },
            Event::Text(text) => {
                let bytes = text.as_ref();
                if preserve_space.is_empty() {
                    if bytes.iter().all(u8::is_ascii_whitespace) {
                        continue;
                    }
                    return Err(Error::InvalidFormat(
                        "character data outside the document XML root".into(),
                    ));
                }
                if bytes.iter().all(u8::is_ascii_whitespace)
                    && !preserve_space.last().copied().unwrap_or(false)
                {
                    if text_run_has_content {
                        push_utf8(&mut output, bytes)?;
                    } else {
                        push_utf8(&mut pending_whitespace, bytes)?;
                    }
                } else {
                    output.push_str(&pending_whitespace);
                    pending_whitespace.clear();
                    push_utf8(&mut output, bytes)?;
                    text_run_has_content = true;
                }
            },
            Event::CData(data) => {
                if preserve_space.is_empty() {
                    return Err(Error::InvalidFormat(
                        "CDATA outside the document XML root".into(),
                    ));
                }
                output.push_str(&pending_whitespace);
                pending_whitespace.clear();
                output.push_str("<![CDATA[");
                push_utf8(&mut output, data.as_ref())?;
                output.push_str("]]>");
                text_run_has_content = true;
            },
            Event::GeneralRef(reference) => {
                if preserve_space.is_empty() {
                    return Err(Error::InvalidFormat(
                        "entity reference outside the document XML root".into(),
                    ));
                }
                output.push_str(&pending_whitespace);
                pending_whitespace.clear();
                output.push('&');
                push_utf8(&mut output, reference.as_ref())?;
                output.push(';');
                text_run_has_content = true;
            },
            Event::Decl(declaration) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                output.push_str("<?");
                push_utf8(&mut output, declaration.as_ref())?;
                output.push_str("?>");
            },
            Event::PI(instruction) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                output.push_str("<?");
                push_utf8(&mut output, instruction.as_ref())?;
                output.push_str("?>");
            },
            Event::Comment(comment) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                output.push_str("<!--");
                push_utf8(&mut output, comment.as_ref())?;
                output.push_str("-->");
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "document XML type declarations are not publishable".into(),
                ));
            },
            Event::Eof => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                break;
            },
        }
    }
    if !preserve_space.is_empty() || roots != 1 {
        return Err(Error::InvalidFormat(
            "document XML must contain exactly one closed root".into(),
        ));
    }
    Ok(output)
}

fn finish_compact_text_run(pending_whitespace: &mut String, has_content: &mut bool) {
    pending_whitespace.clear();
    *has_content = false;
}

fn element_preserves_space(
    element: &BytesStart<'_>,
    decoder: Decoder,
    inherited: bool,
) -> Result<bool> {
    let mut preserve = inherited;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() != b"xml:space" {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(Error::from)?;
        preserve = match value.as_ref() {
            "preserve" => true,
            "default" => false,
            _ => {
                return Err(Error::InvalidFormat(
                    "xml:space must be 'default' or 'preserve'".into(),
                ));
            },
        };
    }
    Ok(preserve)
}

fn write_compact_start(
    output: &mut String,
    element: &BytesStart<'_>,
    decoder: Decoder,
    empty: bool,
) -> Result<()> {
    output.push('<');
    push_utf8(output, element.name().as_ref())?;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        output.push(' ');
        push_utf8(output, attribute.key.as_ref())?;
        output.push_str("=\"");
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(Error::from)?;
        output.push_str(quick_xml::escape::escape(value.as_ref()).as_ref());
        output.push('"');
    }
    output.push_str(if empty { "/>" } else { ">" });
    Ok(())
}

fn push_utf8(output: &mut String, bytes: &[u8]) -> Result<()> {
    output.push_str(std::str::from_utf8(bytes).map_err(|error| Error::Xml(error.to_string()))?);
    Ok(())
}

pub(super) fn write_document_protection(
    xml: &mut String,
    protection: &Protection,
    prefix: &str,
    local_namespace: Option<&str>,
) -> Result<()> {
    write!(xml, "<{prefix}:documentProtection").map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(namespace) = local_namespace {
        write!(
            xml,
            " xmlns:{prefix}=\"{}\"",
            litchi_core::xml::escape_xml(namespace)
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    write!(
        xml,
        " {prefix}:edit=\"{}\" {prefix}:enforcement=\"1\"",
        protection.protection_type.to_xml()
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    if let Some(hash) = &protection.password_hash {
        write!(
            xml,
            " {prefix}:hash=\"{}\"",
            litchi_core::xml::escape_xml(hash)
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(salt) = &protection.salt {
        write!(
            xml,
            " {prefix}:salt=\"{}\"",
            litchi_core::xml::escape_xml(salt)
        )
        .map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(())
}

pub(super) fn patch_document_protection(
    existing: &[u8],
    protection: Option<&Protection>,
) -> Result<Vec<u8>> {
    use crate::namespace::scan_word_element_ranges;

    let existing = std::str::from_utf8(existing).map_err(|_source_error| {
        Error::InvalidFormat("settings.xml must be UTF-8 to modify document protection".into())
    })?;
    let mut ranges = Vec::new();
    scan_word_element_ranges(
        existing.as_bytes(),
        &[b"documentProtection"],
        |_, start, len| {
            let start = usize::try_from(start).map_err(|_source_error| {
                Error::InvalidFormat("settings protection offset does not fit usize".into())
            })?;
            let len = usize::try_from(len).map_err(|_source_error| {
                Error::InvalidFormat("settings protection length does not fit usize".into())
            })?;
            let end = start.checked_add(len).ok_or_else(|| {
                Error::InvalidFormat("settings protection range overflows usize".into())
            })?;
            ranges.push((start, end));
            Ok(())
        },
    )?;
    if ranges.len() > 1 {
        return Err(Error::InvalidFormat(
            "settings.xml contains duplicate documentProtection elements".into(),
        ));
    }

    let root = locate_settings_root(existing.as_bytes())?;
    let (root_name, root_namespace) = root.name_and_namespace();
    let (prefix, local_namespace) = match root_name.split_once(':') {
        Some((prefix, _)) => (prefix, None),
        None => ("w", Some(root_namespace)),
    };
    let mut replacement = String::new();
    if let Some(protection) = protection {
        write_document_protection(&mut replacement, protection, prefix, local_namespace)?;
    }

    if let Some((start, end)) = ranges.first().copied() {
        let mut output = String::with_capacity(existing.len() - (end - start) + replacement.len());
        output.push_str(&existing[..start]);
        output.push_str(&replacement);
        output.push_str(&existing[end..]);
        return Ok(output.into_bytes());
    }
    if replacement.is_empty() {
        return Ok(existing.as_bytes().to_vec());
    }

    match root {
        SettingsRoot::Paired { close_offset, .. } => {
            let mut output = String::with_capacity(existing.len() + replacement.len());
            output.push_str(&existing[..close_offset]);
            output.push_str(&replacement);
            output.push_str(&existing[close_offset..]);
            Ok(output.into_bytes())
        },
        SettingsRoot::Empty { end, name, .. } => {
            let empty_close = end
                .checked_sub(2)
                .ok_or_else(|| Error::InvalidFormat("invalid empty settings root range".into()))?;
            if existing.as_bytes().get(empty_close..end) != Some(b"/>") {
                return Err(Error::InvalidFormat(
                    "invalid empty settings root syntax".into(),
                ));
            }
            let mut output =
                String::with_capacity(existing.len() + replacement.len() + name.len() + 4);
            output.push_str(&existing[..empty_close]);
            output.push('>');
            output.push_str(&replacement);
            output.push_str("</");
            output.push_str(&name);
            output.push('>');
            output.push_str(&existing[end..]);
            Ok(output.into_bytes())
        },
    }
}

enum SettingsRoot {
    Paired {
        close_offset: usize,
        name: String,
        namespace: String,
    },
    Empty {
        end: usize,
        name: String,
        namespace: String,
    },
}

impl SettingsRoot {
    fn name_and_namespace(&self) -> (&str, &str) {
        match self {
            Self::Paired {
                name, namespace, ..
            }
            | Self::Empty {
                name, namespace, ..
            } => (name, namespace),
        }
    }
}

fn locate_settings_root(xml: &[u8]) -> Result<SettingsRoot> {
    use crate::namespace::is_wordprocessing_namespace;
    use quick_xml::events::Event;
    use quick_xml::reader::NsReader;

    enum RootEvent {
        Start(Option<(String, String)>),
        Empty(Option<(String, String)>),
        End(bool),
        Eof,
        Other,
    }

    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut root_info = None;
    loop {
        let event_start = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("settings root offset does not fit usize".into())
        })?;
        let event = {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            match event {
                Event::Start(element) => RootEvent::Start(settings_root_info(
                    &namespace,
                    element.name().as_ref(),
                    element.local_name().as_ref(),
                )),
                Event::Empty(element) => RootEvent::Empty(settings_root_info(
                    &namespace,
                    element.name().as_ref(),
                    element.local_name().as_ref(),
                )),
                Event::End(element) => RootEvent::End(
                    is_wordprocessing_namespace(&namespace)
                        && element.local_name().as_ref() == b"settings",
                ),
                Event::Eof => RootEvent::Eof,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => RootEvent::Other,
            }
        };
        let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            Error::InvalidFormat("settings root offset does not fit usize".into())
        })?;

        match event {
            RootEvent::Start(info) if depth == 0 => {
                if saw_root || info.is_none() {
                    return Err(Error::InvalidFormat(
                        "settings.xml has an invalid or trailing root".into(),
                    ));
                }
                saw_root = true;
                root_info = info;
                depth = 1;
            },
            RootEvent::Start(_) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("settings XML nesting is too deep".into())
                })?;
            },
            RootEvent::Empty(info) if depth == 0 => {
                if saw_root || info.is_none() {
                    return Err(Error::InvalidFormat(
                        "settings.xml has an invalid or trailing root".into(),
                    ));
                }
                let (name, namespace) = info.ok_or_else(|| {
                    Error::InvalidFormat("empty settings root has no name".into())
                })?;
                return Ok(SettingsRoot::Empty {
                    end: event_end,
                    name,
                    namespace,
                });
            },
            RootEvent::End(is_root) => {
                if depth == 1 {
                    if !is_root {
                        return Err(Error::InvalidFormat(
                            "settings.xml has an invalid root closing element".into(),
                        ));
                    }
                    let (name, namespace) = root_info.take().ok_or_else(|| {
                        Error::InvalidFormat("settings root metadata is missing".into())
                    })?;
                    return Ok(SettingsRoot::Paired {
                        close_offset: event_start,
                        name,
                        namespace,
                    });
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid settings XML nesting".into()))?;
            },
            RootEvent::Eof => {
                return Err(Error::InvalidFormat(
                    "settings.xml has no complete settings root".into(),
                ));
            },
            RootEvent::Empty(_) | RootEvent::Other => {},
        }
    }
}

fn settings_root_info(
    namespace: &quick_xml::name::ResolveResult<'_>,
    qualified_name: &[u8],
    local_name: &[u8],
) -> Option<(String, String)> {
    use quick_xml::name::{Namespace, ResolveResult};

    if local_name != b"settings" || !crate::namespace::is_wordprocessing_namespace(namespace) {
        return None;
    }
    let ResolveResult::Bound(Namespace(namespace)) = namespace else {
        return None;
    };
    Some((
        String::from_utf8_lossy(qualified_name).into_owned(),
        String::from_utf8_lossy(namespace).into_owned(),
    ))
}
