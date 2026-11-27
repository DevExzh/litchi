use super::super::semantic::changes::{Change, Changes, Metadata, Replies, Reply};
use super::super::semantic::extensions::OpaqueXml;
use super::super::{A, AC, P, PC2, PC226};
use super::monikers::{parse_monikers, write_monikers};
use super::xml::{
    Fragment, attr, attribute, close, only_attributes, open, resolve_namespace, scan,
};
use crate::{Error, Result};
use litchi_core::xml::ReaderOrigin;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

/// The last child group a `pc226:cmChg` accepted:
///
/// ```text
/// Start --chgData--> Metadata --cmMkLst--> MonikersOrReplies --extLst--> Extensions
/// Start --cmMkLst--> MonikersOrReplies
/// Start | Metadata | MonikersOrReplies --cmRplyChg--> MonikersOrReplies
/// Start | Metadata --extLst--> Extensions
/// ```
///
/// `chgData` is accepted in `Start` only. `cmMkLst` and `extLst` are
/// accepted at most once and are refused once the stage has reached or
/// passed their own group. `cmRplyChg` is repeatable and is refused only in
/// `Extensions`. A `cmRplyChg` that arrives before `cmMkLst` is accepted and
/// moves to `MonikersOrReplies`, so a later `cmMkLst` is refused as out of
/// order; without one the element fails after its last child because
/// `cmMkLst` is required.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum CommentChangeStage {
    Start,
    Metadata,
    MonikersOrReplies,
    Extensions,
}

/// The last child a `pc226:cmRplyChg` accepted, in schema order:
///
/// ```text
/// Start --chgData--> Metadata --cmRplyMkLst--> Monikers --extLst--> Extensions
/// Start --cmRplyMkLst--> Monikers
/// Start | Metadata --extLst--> Extensions
/// ```
///
/// `chgData` is accepted in `Start` only. `cmRplyMkLst` and `extLst` are
/// accepted at most once and are refused once the stage has reached or
/// passed them. `cmRplyMkLst` is required, which is checked after the last
/// child.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum ReplyChangeStage {
    Start,
    Metadata,
    Monikers,
    Extensions,
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
    let mut stage = CommentChangeStage::Start;
    for child in &scan.children {
        if child.namespace == AC && child.local == "chgData" {
            if stage != CommentChangeStage::Start {
                return Err(invalid(
                    "comment change chgData is duplicated or out of order",
                ));
            }
            metadata = Some(parse_metadata(child)?);
            stage = CommentChangeStage::Metadata;
        } else if child.namespace == PC2 && child.local == "cmMkLst" {
            if stage >= CommentChangeStage::MonikersOrReplies {
                return Err(invalid(
                    "comment change cmMkLst is duplicated or out of order",
                ));
            }
            monikers = Some(parse_monikers(&child.xml)?);
            stage = CommentChangeStage::MonikersOrReplies;
        } else if child.namespace == PC226 && child.local == "cmRplyChg" {
            if stage >= CommentChangeStage::Extensions {
                return Err(invalid("comment reply change is out of order"));
            }
            replies.push(parse_reply_changes(child)?);
            stage = CommentChangeStage::MonikersOrReplies;
        } else if child.namespace == P && child.local == "extLst" {
            if stage >= CommentChangeStage::Extensions {
                return Err(invalid(
                    "comment change extLst is duplicated or out of order",
                ));
            }
            extension = Some(OpaqueXml::new(child.xml.clone())?);
            stage = CommentChangeStage::Extensions;
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
    let mut stage = ReplyChangeStage::Start;
    for child in &scan.children {
        if child.namespace == AC && child.local == "chgData" {
            if stage != ReplyChangeStage::Start {
                return Err(invalid(
                    "reply change chgData is duplicated or out of order",
                ));
            }
            metadata = Some(parse_metadata(child)?);
            stage = ReplyChangeStage::Metadata;
        } else if child.namespace == PC2 && child.local == "cmRplyMkLst" {
            if stage >= ReplyChangeStage::Monikers {
                return Err(invalid(
                    "reply change moniker list is duplicated or out of order",
                ));
            }
            monikers = Some(parse_monikers(&child.xml)?);
            stage = ReplyChangeStage::Monikers;
        } else if child.namespace == P && child.local == "extLst" {
            if stage >= ReplyChangeStage::Extensions {
                return Err(invalid("reply change extLst is duplicated or out of order"));
            }
            extension = Some(OpaqueXml::new(child.xml.clone())?);
            stage = ReplyChangeStage::Extensions;
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
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut output = Vec::new();
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("XML source position exceeds usize"))?;
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
                    let end = origin
                        .offset(reader.buffer_position())
                        .ok_or_else(|| invalid("XML source position exceeds usize"))?;
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
                    let end = origin
                        .offset(reader.buffer_position())
                        .ok_or_else(|| invalid("XML source position exceeds usize"))?;
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

#[cfg(test)]
mod tests {
    use super::*;

    /// Every child sequence of at most `longest` children drawn from `letters`.
    fn sequences(letters: &str, longest: usize) -> Vec<String> {
        let mut all = vec![String::new()];
        let mut previous = all.clone();
        for _ in 0..longest {
            let mut next = Vec::new();
            for prefix in &previous {
                for letter in letters.chars() {
                    next.push(format!("{prefix}{letter}"));
                }
            }
            all.extend(next.iter().cloned());
            previous = next;
        }
        all
    }

    /// Checks every row and requires the table to hold the empty parent, each
    /// child alone and every ordered pair (duplicates included) of `letters`.
    fn assert_matrix(rows: &[(&str, &str)], letters: &str, outcome: impl Fn(&str) -> String) {
        for sequence in sequences(letters, 2) {
            assert!(
                rows.iter().any(|(key, _)| *key == sequence),
                "missing row {sequence:?}"
            );
        }
        let differing: Vec<String> = rows
            .iter()
            .filter_map(|(key, expected)| {
                let actual = outcome(key);
                (actual != *expected).then(|| format!("({key:?}, {actual:?}),"))
            })
            .collect();
        assert!(
            differing.is_empty(),
            "rows differ from the table:\n{}",
            differing.join("\n")
        );
    }

    const GUID: &str = "{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}";

    fn outcome_text<T>(result: Result<T>, describe: impl Fn(&T) -> String) -> String {
        match result {
            Ok(value) => describe(&value),
            Err(Error::Invalid(message)) => format!("Invalid: {message}"),
            Err(other) => format!("{other:?}"),
        }
    }

    fn comment_monikers() -> String {
        format!(r#"<pc2:cmMkLst xmlns:pc2="{PC2}"><pc2:cmMk id="{GUID}"/></pc2:cmMkLst>"#)
    }

    fn reply_monikers() -> String {
        format!(
            r#"<pc2:cmRplyMkLst xmlns:pc2="{PC2}"><pc2:cmMk id="{GUID}"/><pc2:cmRplyMk id="{GUID}"/></pc2:cmRplyMkLst>"#
        )
    }

    fn reply_change(children: &str) -> String {
        format!(r#"<pc226:cmRplyChg xmlns:pc226="{PC226}" chg="mod">{children}</pc226:cmRplyChg>"#)
    }

    fn comment_change(children: &str) -> Result<Changes> {
        parse_comment_changes(
            format!(r#"<pc226:cmChg xmlns:pc226="{PC226}" chg="mod">{children}</pc226:cmChg>"#)
                .as_bytes(),
        )
    }

    /// Child-order table for `pc226:cmChg`. Key letters: `C` `ac:chgData`, `M`
    /// `pc2:cmMkLst`, `R` `pc226:cmRplyChg`, `E` `p:extLst`, `X` an unknown
    /// child.
    const COMMENT_CHANGE_CHILD_ORDER: &[(&str, &str)] = &[
        ("", "Invalid: comment change requires cmMkLst"),
        ("C", "Invalid: comment change requires cmMkLst"),
        ("M", "ok metadata=false replies=0 extension=false"),
        ("R", "Invalid: comment change requires cmMkLst"),
        ("E", "Invalid: comment change requires cmMkLst"),
        (
            "CC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        ("CM", "ok metadata=true replies=0 extension=false"),
        ("CR", "Invalid: comment change requires cmMkLst"),
        ("CE", "Invalid: comment change requires cmMkLst"),
        (
            "MC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "MM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        ("MR", "ok metadata=false replies=1 extension=false"),
        ("ME", "ok metadata=false replies=0 extension=true"),
        (
            "RC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "RM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        ("RR", "Invalid: comment change requires cmMkLst"),
        ("RE", "Invalid: comment change requires cmMkLst"),
        (
            "EC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "EM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        ("ER", "Invalid: comment reply change is out of order"),
        (
            "EE",
            "Invalid: comment change extLst is duplicated or out of order",
        ),
        (
            "CCC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "CCM",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "CCR",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "CCE",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "CMC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "CMM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        ("CMR", "ok metadata=true replies=1 extension=false"),
        ("CME", "ok metadata=true replies=0 extension=true"),
        (
            "CRC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "CRM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        ("CRR", "Invalid: comment change requires cmMkLst"),
        ("CRE", "Invalid: comment change requires cmMkLst"),
        (
            "CEC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "CEM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        ("CER", "Invalid: comment reply change is out of order"),
        (
            "CEE",
            "Invalid: comment change extLst is duplicated or out of order",
        ),
        (
            "MCC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "MCM",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "MCR",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "MCE",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "MMC",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        (
            "MMM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        (
            "MMR",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        (
            "MME",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        (
            "MRC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "MRM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        ("MRR", "ok metadata=false replies=2 extension=false"),
        ("MRE", "ok metadata=false replies=1 extension=true"),
        (
            "MEC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "MEM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        ("MER", "Invalid: comment reply change is out of order"),
        (
            "MEE",
            "Invalid: comment change extLst is duplicated or out of order",
        ),
        (
            "RCC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "RCM",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "RCR",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "RCE",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "RMC",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        (
            "RMM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        (
            "RMR",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        (
            "RME",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        (
            "RRC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "RRM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        ("RRR", "Invalid: comment change requires cmMkLst"),
        ("RRE", "Invalid: comment change requires cmMkLst"),
        (
            "REC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "REM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        ("RER", "Invalid: comment reply change is out of order"),
        (
            "REE",
            "Invalid: comment change extLst is duplicated or out of order",
        ),
        (
            "ECC",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "ECM",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "ECR",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "ECE",
            "Invalid: comment change chgData is duplicated or out of order",
        ),
        (
            "EMC",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        (
            "EMM",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        (
            "EMR",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        (
            "EME",
            "Invalid: comment change cmMkLst is duplicated or out of order",
        ),
        ("ERC", "Invalid: comment reply change is out of order"),
        ("ERM", "Invalid: comment reply change is out of order"),
        ("ERR", "Invalid: comment reply change is out of order"),
        ("ERE", "Invalid: comment reply change is out of order"),
        (
            "EEC",
            "Invalid: comment change extLst is duplicated or out of order",
        ),
        (
            "EEM",
            "Invalid: comment change extLst is duplicated or out of order",
        ),
        (
            "EER",
            "Invalid: comment change extLst is duplicated or out of order",
        ),
        (
            "EEE",
            "Invalid: comment change extLst is duplicated or out of order",
        ),
        ("CMRE", "ok metadata=true replies=1 extension=true"),
        ("CMRRE", "ok metadata=true replies=2 extension=true"),
        ("MRRRR", "ok metadata=false replies=4 extension=false"),
        ("X", "Invalid: unexpected comment change child"),
        ("MX", "Invalid: unexpected comment change child"),
        ("XM", "Invalid: unexpected comment change child"),
        ("MEX", "Invalid: unexpected comment change child"),
    ];

    fn comment_change_outcome(key: &str) -> String {
        let mut children = String::new();
        for letter in key.chars() {
            children.push_str(&match letter {
                'C' => format!(r#"<ac:chgData xmlns:ac="{AC}" name="Ada"/>"#),
                'M' => comment_monikers(),
                'R' => reply_change(&reply_monikers()),
                'E' => format!(r#"<p:extLst xmlns:p="{P}"/>"#),
                'X' => r#"<q:other xmlns:q="urn:example:other"/>"#.to_owned(),
                other => panic!("unknown key letter {other:?}"),
            });
        }
        outcome_text(comment_change(&children), |value| {
            format!(
                "ok metadata={} replies={} extension={}",
                value.metadata.is_some(),
                value.reply_changes.len(),
                value.extension_xml.is_some()
            )
        })
    }

    #[test]
    fn comment_change_children_keep_their_order_and_duplicate_rules() {
        assert_matrix(COMMENT_CHANGE_CHILD_ORDER, "CMRE", comment_change_outcome);
    }

    /// Child-order table for `pc226:cmRplyChg`. Key letters: `C` `ac:chgData`,
    /// `K` `pc2:cmRplyMkLst`, `E` `p:extLst`, `X` an unknown child.
    const REPLY_CHANGE_CHILD_ORDER: &[(&str, &str)] = &[
        ("", "Invalid: reply change requires cmRplyMkLst"),
        ("C", "Invalid: reply change requires cmRplyMkLst"),
        ("K", "ok metadata=false extension=false"),
        ("E", "Invalid: reply change requires cmRplyMkLst"),
        (
            "CC",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        ("CK", "ok metadata=true extension=false"),
        ("CE", "Invalid: reply change requires cmRplyMkLst"),
        (
            "KC",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "KK",
            "Invalid: reply change moniker list is duplicated or out of order",
        ),
        ("KE", "ok metadata=false extension=true"),
        (
            "EC",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "EK",
            "Invalid: reply change moniker list is duplicated or out of order",
        ),
        (
            "EE",
            "Invalid: reply change extLst is duplicated or out of order",
        ),
        (
            "CCC",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "CCK",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "CCE",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "CKC",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "CKK",
            "Invalid: reply change moniker list is duplicated or out of order",
        ),
        ("CKE", "ok metadata=true extension=true"),
        (
            "CEC",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "CEK",
            "Invalid: reply change moniker list is duplicated or out of order",
        ),
        (
            "CEE",
            "Invalid: reply change extLst is duplicated or out of order",
        ),
        (
            "KCC",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "KCK",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "KCE",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "KKC",
            "Invalid: reply change moniker list is duplicated or out of order",
        ),
        (
            "KKK",
            "Invalid: reply change moniker list is duplicated or out of order",
        ),
        (
            "KKE",
            "Invalid: reply change moniker list is duplicated or out of order",
        ),
        (
            "KEC",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "KEK",
            "Invalid: reply change moniker list is duplicated or out of order",
        ),
        (
            "KEE",
            "Invalid: reply change extLst is duplicated or out of order",
        ),
        (
            "ECC",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "ECK",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "ECE",
            "Invalid: reply change chgData is duplicated or out of order",
        ),
        (
            "EKC",
            "Invalid: reply change moniker list is duplicated or out of order",
        ),
        (
            "EKK",
            "Invalid: reply change moniker list is duplicated or out of order",
        ),
        (
            "EKE",
            "Invalid: reply change moniker list is duplicated or out of order",
        ),
        (
            "EEC",
            "Invalid: reply change extLst is duplicated or out of order",
        ),
        (
            "EEK",
            "Invalid: reply change extLst is duplicated or out of order",
        ),
        (
            "EEE",
            "Invalid: reply change extLst is duplicated or out of order",
        ),
        ("X", "Invalid: unexpected comment reply change child"),
        ("KX", "Invalid: unexpected comment reply change child"),
        ("XK", "Invalid: unexpected comment reply change child"),
        ("KEX", "Invalid: unexpected comment reply change child"),
    ];

    fn reply_change_outcome(key: &str) -> String {
        let mut children = String::new();
        for letter in key.chars() {
            children.push_str(&match letter {
                'C' => format!(r#"<ac:chgData xmlns:ac="{AC}" name="Ada"/>"#),
                'K' => reply_monikers(),
                'E' => format!(r#"<p:extLst xmlns:p="{P}"/>"#),
                'X' => r#"<q:other xmlns:q="urn:example:other"/>"#.to_owned(),
                other => panic!("unknown key letter {other:?}"),
            });
        }
        let parent = format!("{}{}", comment_monikers(), reply_change(&children));
        outcome_text(comment_change(&parent), |value| {
            let reply = value.reply_changes.first().unwrap();
            format!(
                "ok metadata={} extension={}",
                reply.metadata.is_some(),
                reply.extension_xml.is_some()
            )
        })
    }

    #[test]
    fn reply_change_children_keep_their_order_and_duplicate_rules() {
        assert_matrix(REPLY_CHANGE_CHILD_ORDER, "CKE", reply_change_outcome);
    }
}
