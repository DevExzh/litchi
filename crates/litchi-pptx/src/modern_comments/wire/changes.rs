use super::super::semantic::changes::{Change, Changes, Metadata, Replies, Reply};
use super::super::semantic::extensions::OpaqueXml;
use super::super::{A, AC, P, PC2, PC226};
use super::monikers::{parse_monikers, write_monikers};
use super::xml::{
    Fragment, attr, attribute, close, only_attributes, open, resolve_namespace, scan,
};
use crate::{Error, Result};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn parse_comment_changes(xml: &[u8]) -> Result<Changes> {
    let scan = scan(xml, "comment change")?;
    if scan.root.namespace != PC226 || scan.root.local != "cmChg" {
        return Err(invalid("comment change root must be pc226:cmChg"));
    }
    only_attributes(&scan.root.attributes, &["chg"], "comment change")?;
    let changes = parse_comment_bits(
        attribute(&scan.root.attributes, "chg", true)?
            .ok_or_else(|| invalid("comment change requires chg"))?,
    )?;
    let mut metadata = None;
    let mut monikers = None;
    let mut replies = Vec::new();
    let mut extension = None;
    let mut order = 0u8;
    for child in &scan.children {
        if child.namespace == AC && child.local == "chgData" {
            if order != 0 || metadata.is_some() {
                return Err(invalid(
                    "comment change chgData is duplicated or out of order",
                ));
            }
            metadata = Some(parse_metadata(child)?);
            order = 1;
        } else if child.namespace == PC2 && child.local == "cmMkLst" {
            if order > 1 || monikers.is_some() {
                return Err(invalid(
                    "comment change cmMkLst is duplicated or out of order",
                ));
            }
            monikers = Some(parse_monikers(&child.xml)?);
            order = 2;
        } else if child.namespace == PC226 && child.local == "cmRplyChg" {
            if order > 2 {
                return Err(invalid("comment reply change is out of order"));
            }
            replies.push(parse_reply_changes(child)?);
            order = 2;
        } else if child.namespace == P && child.local == "extLst" {
            if order > 2 || extension.is_some() {
                return Err(invalid(
                    "comment change extLst is duplicated or out of order",
                ));
            }
            extension = Some(OpaqueXml::new(child.xml.clone())?);
            order = 3;
        } else {
            return Err(invalid("unexpected comment change child"));
        }
    }
    let value = Changes {
        changes,
        metadata,
        monikers: monikers.ok_or_else(|| invalid("comment change requires cmMkLst"))?,
        reply_changes: replies,
        extension_xml: extension,
        namespace_declarations: Vec::new(),
    };
    value.validate()?;
    Ok(value)
}

fn parse_reply_changes(fragment: &Fragment) -> Result<Replies> {
    only_attributes(&fragment.attributes, &["chg"], "comment reply change")?;
    let changes = parse_reply_bits(
        attribute(&fragment.attributes, "chg", true)?
            .ok_or_else(|| invalid("comment reply change requires chg"))?,
    )?;
    let scan = scan(&fragment.xml, "comment reply change")?;
    let mut metadata = None;
    let mut monikers = None;
    let mut extension = None;
    let mut order = 0u8;
    for child in &scan.children {
        if child.namespace == AC && child.local == "chgData" {
            if order != 0 || metadata.is_some() {
                return Err(invalid(
                    "reply change chgData is duplicated or out of order",
                ));
            }
            metadata = Some(parse_metadata(child)?);
            order = 1;
        } else if child.namespace == PC2 && child.local == "cmRplyMkLst" {
            if order > 1 || monikers.is_some() {
                return Err(invalid(
                    "reply change moniker list is duplicated or out of order",
                ));
            }
            monikers = Some(parse_monikers(&child.xml)?);
            order = 2;
        } else if child.namespace == P && child.local == "extLst" {
            if order > 2 || extension.is_some() {
                return Err(invalid("reply change extLst is duplicated or out of order"));
            }
            extension = Some(OpaqueXml::new(child.xml.clone())?);
            order = 3;
        } else {
            return Err(invalid("unexpected comment reply change child"));
        }
    }
    let value = Replies {
        changes,
        metadata,
        monikers: monikers.ok_or_else(|| invalid("reply change requires cmRplyMkLst"))?,
        extension_xml: extension,
        namespace_declarations: Vec::new(),
    };
    value.validate()?;
    Ok(value)
}

// The schema has two distinct token lists. Keeping the parsers separate avoids
// accepting a reply-only bit in a comment command or vice versa.
fn parse_comment_bits(value: &str) -> Result<Vec<Change>> {
    let mut output = Vec::new();
    for token in value.split_whitespace() {
        let bit = Change::parse(token)?;
        if output.contains(&bit) {
            return Err(invalid("duplicate comment change bit"));
        }
        output.push(bit);
    }
    if output.is_empty() {
        return Err(invalid("comment change bit list is empty"));
    }
    Ok(output)
}

fn parse_reply_bits(value: &str) -> Result<Vec<Reply>> {
    let mut output = Vec::new();
    for token in value.split_whitespace() {
        let bit = Reply::parse(token)?;
        if output.contains(&bit) {
            return Err(invalid("duplicate reply change bit"));
        }
        output.push(bit);
    }
    if output.is_empty() {
        return Err(invalid("reply change bit list is empty"));
    }
    Ok(output)
}

fn parse_metadata(fragment: &Fragment) -> Result<Metadata> {
    let mut value = Metadata::default();
    for (key, text) in &fragment.attributes {
        match key.as_str() {
            "name" => value.name = Some(text.clone()),
            "userId" => value.user_id = Some(text.clone()),
            "providerId" => value.provider_id = Some(text.clone()),
            "clId" => value.client_id = Some(text.clone()),
            "email" => value.email = Some(text.clone()),
            "dt" => value.date_time = Some(text.clone()),
            "v" => {
                value.version = Some(
                    text.parse()
                        .map_err(|_err| invalid("invalid change version"))?,
                );
            },
            "id" => value.change_id = Some(text.clone()),
            "actId" => {
                value.action_id = Some(
                    text.parse()
                        .map_err(|_err| invalid("invalid change action ID"))?,
                );
            },
            other => {
                return Err(invalid(format!(
                    "unexpected change metadata attribute '{other}'"
                )));
            },
        }
    }
    let scan = scan(&fragment.xml, "change metadata")?;
    for child in &scan.children {
        if child.namespace == A && child.local == "extLst" && value.extension_xml.is_none() {
            value.extension_xml = Some(OpaqueXml::new(child.xml.clone())?);
        } else {
            return Err(invalid("unexpected change metadata child"));
        }
    }
    value.validate()?;
    Ok(value)
}

pub(super) fn write_comment_changes(value: &Changes) -> Result<Vec<u8>> {
    value.validate()?;
    let mut out = Vec::new();
    open(&mut out, "pc226", "cmChg");
    attr(&mut out, "chg", &comment_bits(&value.changes));
    out.extend_from_slice(
        b" xmlns:pc226=\"http://schemas.microsoft.com/office/powerpoint/2022/06/main/command\"",
    );
    out.extend_from_slice(
        b" xmlns:pc2=\"http://schemas.microsoft.com/office/powerpoint/2019/9/main/command\"",
    );
    out.extend_from_slice(
        b" xmlns:ac=\"http://schemas.microsoft.com/office/drawing/2013/main/command\"",
    );
    out.extend_from_slice(b" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
    out.extend_from_slice(
        b" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"",
    );
    out.push(b'>');
    if let Some(metadata) = &value.metadata {
        write_metadata(&mut out, metadata);
    }
    out.extend_from_slice(&write_monikers(&value.monikers)?);
    for reply in &value.reply_changes {
        write_reply_changes(&mut out, reply)?;
    }
    if let Some(extension) = &value.extension_xml {
        out.extend_from_slice(extension.as_bytes());
    }
    close(&mut out, "pc226", "cmChg");
    Ok(out)
}

fn write_reply_changes(out: &mut Vec<u8>, value: &Replies) -> Result<()> {
    open(out, "pc226", "cmRplyChg");
    attr(out, "chg", &reply_bits(&value.changes));
    out.push(b'>');
    if let Some(metadata) = &value.metadata {
        write_metadata(out, metadata);
    }
    out.extend_from_slice(&write_monikers(&value.monikers)?);
    if let Some(extension) = &value.extension_xml {
        out.extend_from_slice(extension.as_bytes());
    }
    close(out, "pc226", "cmRplyChg");
    Ok(())
}

fn write_metadata(out: &mut Vec<u8>, value: &Metadata) {
    open(out, "ac", "chgData");
    for (name, text) in [
        ("name", value.name.as_deref()),
        ("userId", value.user_id.as_deref()),
        ("providerId", value.provider_id.as_deref()),
        ("clId", value.client_id.as_deref()),
        ("email", value.email.as_deref()),
        ("dt", value.date_time.as_deref()),
        ("id", value.change_id.as_deref()),
    ] {
        if let Some(text) = text {
            attr(out, name, text);
        }
    }
    if let Some(version) = value.version {
        attr(out, "v", &version.to_string());
    }
    if let Some(action_id) = value.action_id {
        attr(out, "actId", &action_id.to_string());
    }
    if let Some(extension) = &value.extension_xml {
        out.push(b'>');
        out.extend_from_slice(extension.as_bytes());
        close(out, "ac", "chgData");
    } else {
        out.extend_from_slice(b"/>");
    }
}

fn comment_bits(value: &[Change]) -> String {
    value
        .iter()
        .map(|value| value.token())
        .collect::<Vec<_>>()
        .join(" ")
}

fn reply_bits(value: &[Reply]) -> String {
    value
        .iter()
        .map(|value| value.token())
        .collect::<Vec<_>>()
        .join(" ")
}

pub(crate) fn collect_change_commands(xml: &[u8]) -> Result<Vec<Changes>> {
    locate_commands(xml)?
        .into_iter()
        .map(|(_, _, value)| Ok(value))
        .collect()
}

pub(crate) fn replace_change_commands(xml: &[u8], replacements: &[Changes]) -> Result<Vec<u8>> {
    let locations = locate_commands(xml)?;
    if locations.len() != replacements.len() {
        return Err(invalid(
            "comment change command count changed during mutation",
        ));
    }
    if locations.is_empty() {
        return Ok(xml.to_vec());
    }
    let mut output = Vec::with_capacity(xml.len());
    let mut cursor = 0usize;
    for ((start, end, _), replacement) in locations.into_iter().zip(replacements) {
        output.extend_from_slice(&xml[cursor..start]);
        output.extend_from_slice(&write_comment_changes(replacement)?);
        cursor = end;
    }
    output.extend_from_slice(&xml[cursor..]);
    if output.len() > super::super::MAX_BYTES {
        return Err(invalid("rewritten comment change descriptor is too large"));
    }
    Ok(output)
}

fn locate_commands(xml: &[u8]) -> Result<Vec<(usize, usize, Changes)>> {
    #[derive(Debug)]
    struct Open {
        namespace: String,
        local: String,
        start: usize,
    }
    if xml.len() > super::super::MAX_BYTES {
        return Err(invalid("comment change descriptor is too large"));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut output = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(super::xml::xml_error)?;
        let namespace = resolve_namespace(resolved)?;
        match event {
            Event::Start(element) => stack.push(Open {
                namespace,
                local: String::from_utf8(element.local_name().as_ref().to_vec())
                    .map_err(super::xml::xml_error)?,
                start,
            }),
            Event::Empty(element) => {
                let local = String::from_utf8(element.local_name().as_ref().to_vec())
                    .map_err(super::xml::xml_error)?;
                if namespace == PC226 && local == "cmChg" {
                    let end = reader.buffer_position() as usize;
                    output.push((start, end, parse_comment_changes(&xml[start..end])?));
                }
            },
            Event::End(element) => {
                let open = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected change closing element"))?;
                let local = String::from_utf8(element.local_name().as_ref().to_vec())
                    .map_err(super::xml::xml_error)?;
                if open.namespace != namespace || open.local != local {
                    return Err(invalid("mismatched change command element"));
                }
                if open.namespace == PC226 && open.local == "cmChg" {
                    let end = reader.buffer_position() as usize;
                    output.push((
                        open.start,
                        end,
                        parse_comment_changes(&xml[open.start..end])?,
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(invalid("hostile construct in comment change descriptor"));
            },
            Event::Text(_) | Event::CData(_) | Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated comment change descriptor"));
    }
    Ok(output)
}
