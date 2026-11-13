use super::super::semantic::extensions::OpaqueXml;
use super::super::semantic::tasks::{
    Action, Anchor, Assign, Details, Event, History, Schedule, Title, Undo, User,
};
use super::super::{P, P228};
use super::xml::scan_with_context;
use super::xml::{Fragment, attr, attribute, close, no_attributes, only_attributes, open, scan};
use crate::{Error, Result};
use std::str::FromStr;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn parse_task_details(xml: &[u8]) -> Result<Details> {
    let scan = scan(xml, "task details")?;
    if scan.root.namespace != P228 || scan.root.local != "taskDetails" {
        return Err(invalid("task details root must be p228:taskDetails"));
    }
    no_attributes(&scan.root.attributes, "task details")?;
    let mut history = None;
    let mut extension = None;
    for child in &scan.children {
        match (child.namespace.as_str(), child.local.as_str()) {
            (P228, "history") if history.is_none() => {
                history = Some(parse_history(&child.xml, &scan.namespaces)?);
            },
            (P, "extLst") if extension.is_none() => {
                extension = Some(OpaqueXml::new(child.xml.clone())?);
            },
            _ => return Err(invalid("unexpected task details child")),
        }
    }
    let value = Details {
        history: history.ok_or_else(|| invalid("task details requires history"))?,
        extension_xml: extension,
        namespace_declarations: scan.namespaces,
    };
    value.validate()?;
    Ok(value)
}

fn parse_history(
    xml: &[u8],
    context: &[super::super::model::NamespaceDeclaration],
) -> Result<History> {
    let scan = scan_with_context(xml, "task history", context)?;
    if scan.root.namespace != P228 || scan.root.local != "history" {
        return Err(invalid("task history root must be p228:history"));
    }
    no_attributes(&scan.root.attributes, "task history")?;
    let mut events = Vec::with_capacity(scan.children.len());
    for child in &scan.children {
        if child.namespace != P228 || child.local != "event" {
            return Err(invalid("task history permits only p228:event children"));
        }
        events.push(parse_event(child, &scan.namespaces)?);
    }
    let value = History { events };
    value.validate()?;
    Ok(value)
}

/// The last child a `p228:event` accepted, in schema order:
///
/// ```text
/// Start --atrbtn--> Attribution --anchr--> Anchor --action--> Action --extLst--> Extensions
/// Start --anchr--> Anchor
/// Start | Attribution --action--> Action
/// Start | Attribution | Anchor --extLst--> Extensions
/// ```
///
/// An action is any `p228` child other than `atrbtn` and `anchr`. `atrbtn`
/// is accepted in `Start` only; `anchr`, the action and `extLst` are each
/// accepted at most once and are refused once the stage has reached or
/// passed them. A child that arrives before `atrbtn` is accepted, so a later
/// `atrbtn` is refused as out of order; without one the event fails after
/// its last child because `atrbtn` is required.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum EventStage {
    Start,
    Attribution,
    Anchor,
    Action,
    Extensions,
}

fn parse_event(
    fragment: &Fragment,
    context: &[super::super::model::NamespaceDeclaration],
) -> Result<Event> {
    only_attributes(&fragment.attributes, &["time", "id"], "task history event")?;
    let time = attribute(&fragment.attributes, "time", true)?
        .ok_or_else(|| invalid("task history event requires time"))?
        .to_owned();
    let id = attribute(&fragment.attributes, "id", true)?
        .ok_or_else(|| invalid("task history event requires id"))?
        .to_owned();
    let mut attributed_by = None;
    let mut anchor = None;
    let mut action = None;
    let mut extension = None;
    let mut stage = EventStage::Start;
    let event_scan = scan_with_context(&fragment.xml, "task history event", context)?;
    for child in &event_scan.children {
        if child.namespace == P228 && child.local == "atrbtn" {
            if stage != EventStage::Start {
                return Err(invalid("task event atrbtn is duplicated or out of order"));
            }
            attributed_by = Some(parse_user(child)?);
            stage = EventStage::Attribution;
        } else if child.namespace == P228 && child.local == "anchr" {
            if stage >= EventStage::Anchor {
                return Err(invalid("task event anchr is duplicated or out of order"));
            }
            anchor = Some(parse_anchor(child, &event_scan.namespaces)?);
            stage = EventStage::Anchor;
        } else if child.namespace == P && child.local == "extLst" {
            if stage >= EventStage::Extensions {
                return Err(invalid("task event extLst is duplicated or out of order"));
            }
            extension = Some(OpaqueXml::new(child.xml.clone())?);
            stage = EventStage::Extensions;
        } else if child.namespace == P228 {
            if stage >= EventStage::Action {
                return Err(invalid("task event action is duplicated or out of order"));
            }
            action = Some(parse_action(child)?);
            stage = EventStage::Action;
        } else {
            return Err(invalid("unexpected task history event child"));
        }
    }
    let value = Event {
        time,
        id,
        attributed_by: attributed_by.ok_or_else(|| invalid("task event requires atrbtn"))?,
        anchor,
        action,
        extension_xml: extension,
        namespace_declarations: context.to_vec(),
    };
    value.validate()?;
    Ok(value)
}

fn parse_user(fragment: &Fragment) -> Result<User> {
    only_attributes(&fragment.attributes, &["authorId"], "task user")?;
    let author_id = attribute(&fragment.attributes, "authorId", true)?
        .ok_or_else(|| invalid("task user requires authorId"))?
        .to_owned();
    Ok(User { author_id })
}

fn parse_anchor(
    fragment: &Fragment,
    context: &[super::super::model::NamespaceDeclaration],
) -> Result<Anchor> {
    only_attributes(&fragment.attributes, &[], "task anchor")?;
    no_attributes(&fragment.attributes, "task anchor")?;
    let scan = scan_with_context(&fragment.xml, "task anchor", context)?;
    let mut comment_id = None;
    let mut extension = None;
    for child in &scan.children {
        match (child.namespace.as_str(), child.local.as_str()) {
            (P228, "comment") if comment_id.is_none() => {
                if !child.attributes.iter().all(|(key, _)| key == "id") {
                    return Err(invalid("unexpected task anchor comment attribute"));
                }
                comment_id = Some(
                    attribute(&child.attributes, "id", true)?
                        .ok_or_else(|| invalid("task anchor comment requires id"))?
                        .to_owned(),
                );
            },
            (P, "extLst") if extension.is_none() => {
                extension = Some(OpaqueXml::new(child.xml.clone())?);
            },
            _ => return Err(invalid("unexpected task anchor child")),
        }
    }
    Ok(Anchor {
        comment_id: comment_id.ok_or_else(|| invalid("task anchor requires comment"))?,
        extension_xml: extension,
        namespace_declarations: context.to_vec(),
    })
}

fn parse_action(fragment: &Fragment) -> Result<Action> {
    match fragment.local.as_str() {
        "asgn" => {
            only_attributes(&fragment.attributes, &["authorId"], "task assignment")?;
            Ok(Action::Assign(Assign {
                author_id: attribute(&fragment.attributes, "authorId", true)?
                    .ok_or_else(|| invalid("task assignment requires authorId"))?
                    .to_owned(),
            }))
        },
        "add" => {
            no_attributes(&fragment.attributes, "task add")?;
            Ok(Action::Add)
        },
        "title" => {
            only_attributes(&fragment.attributes, &["val"], "task title")?;
            Ok(Action::Title(Title {
                value: attribute(&fragment.attributes, "val", true)?
                    .ok_or_else(|| invalid("task title requires val"))?
                    .to_owned(),
            }))
        },
        "date" => {
            only_attributes(&fragment.attributes, &["stDt", "endDt"], "task schedule")?;
            Ok(Action::Schedule(Schedule {
                start_date: attribute(&fragment.attributes, "stDt", false)?.map(str::to_owned),
                end_date: attribute(&fragment.attributes, "endDt", false)?.map(str::to_owned),
            }))
        },
        "pcntCmplt" => {
            only_attributes(&fragment.attributes, &["val"], "task progress")?;
            let value = attribute(&fragment.attributes, "val", true)?
                .ok_or_else(|| invalid("task progress requires val"))?;
            Ok(Action::Progress(super::super::model::Progress::from_str(
                value,
            )?))
        },
        "unasgnAll" => {
            no_attributes(&fragment.attributes, "task unassign-all")?;
            Ok(Action::UnassignAll)
        },
        "undo" => {
            only_attributes(&fragment.attributes, &["id"], "task undo")?;
            Ok(Action::Undo(Undo {
                event_id: attribute(&fragment.attributes, "id", true)?
                    .ok_or_else(|| invalid("task undo requires id"))?
                    .to_owned(),
            }))
        },
        "unknown" => {
            no_attributes(&fragment.attributes, "unknown task event")?;
            Ok(Action::Unknown(OpaqueXml::new(fragment.xml.clone())?))
        },
        _ => Ok(Action::Unknown(OpaqueXml::new(fragment.xml.clone())?)),
    }
}

pub(super) fn write_task_details(value: &Details) -> Result<Vec<u8>> {
    value.validate()?;
    let mut out = Vec::new();
    open(&mut out, "p228", "taskDetails");
    out.extend_from_slice(
        b" xmlns:p228=\"http://schemas.microsoft.com/office/powerpoint/2022/08/main\"",
    );
    out.extend_from_slice(
        b" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"",
    );
    for declaration in &value.namespace_declarations {
        if declaration.uri != P228
            && declaration.uri != P
            && declaration.prefix != "p228"
            && declaration.prefix != "p"
        {
            super::xml::namespaces(&mut out, std::slice::from_ref(declaration));
        }
    }
    out.push(b'>');
    write_history(&mut out, &value.history);
    if let Some(extension) = &value.extension_xml {
        out.extend_from_slice(extension.as_bytes());
    }
    close(&mut out, "p228", "taskDetails");
    Ok(out)
}

fn write_history(out: &mut Vec<u8>, value: &History) {
    open(out, "p228", "history");
    if value.events.is_empty() {
        out.extend_from_slice(b"/>");
        return;
    }
    out.push(b'>');
    for event in &value.events {
        write_event(out, event);
    }
    close(out, "p228", "history");
}

fn write_event(out: &mut Vec<u8>, value: &Event) {
    open(out, "p228", "event");
    attr(out, "time", &value.time);
    attr(out, "id", &value.id);
    out.push(b'>');
    write_user(out, "atrbtn", &value.attributed_by.author_id);
    if let Some(anchor) = &value.anchor {
        write_anchor(out, anchor);
    }
    if let Some(action) = &value.action {
        write_action(out, action);
    }
    if let Some(extension) = &value.extension_xml {
        out.extend_from_slice(extension.as_bytes());
    }
    close(out, "p228", "event");
}

fn write_user(out: &mut Vec<u8>, local: &str, author_id: &str) {
    open(out, "p228", local);
    attr(out, "authorId", author_id);
    out.extend_from_slice(b"/>");
}

fn write_anchor(out: &mut Vec<u8>, value: &Anchor) {
    open(out, "p228", "anchr");
    out.push(b'>');
    open(out, "p228", "comment");
    attr(out, "id", &value.comment_id);
    out.extend_from_slice(b"/>");
    if let Some(extension) = &value.extension_xml {
        out.extend_from_slice(extension.as_bytes());
    }
    close(out, "p228", "anchr");
}

fn write_action(out: &mut Vec<u8>, value: &Action) {
    match value {
        Action::Assign(value) => write_user(out, "asgn", &value.author_id),
        Action::Add => empty(out, "add"),
        Action::Title(value) => {
            open(out, "p228", "title");
            attr(out, "val", &value.value);
            out.extend_from_slice(b"/>");
        },
        Action::Schedule(value) => {
            open(out, "p228", "date");
            if let Some(date) = &value.start_date {
                attr(out, "stDt", date);
            }
            if let Some(date) = &value.end_date {
                attr(out, "endDt", date);
            }
            out.extend_from_slice(b"/>");
        },
        Action::Progress(value) => {
            open(out, "p228", "pcntCmplt");
            attr(out, "val", &value.thousandths().to_string());
            out.extend_from_slice(b"/>");
        },
        Action::UnassignAll => empty(out, "unasgnAll"),
        Action::Undo(value) => {
            open(out, "p228", "undo");
            attr(out, "id", &value.event_id);
            out.extend_from_slice(b"/>");
        },
        Action::Unknown(value) => out.extend_from_slice(value.as_bytes()),
    }
}

fn empty(out: &mut Vec<u8>, local: &str) {
    open(out, "p228", local);
    out.extend_from_slice(b"/>");
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]
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

    /// Child-order table for `p228:event`. Key letters: `A` `p228:atrbtn`, `N`
    /// `p228:anchr`, `T` the `p228:add` action, `L` the `p228:title` action, `U`
    /// a `p228` element this crate does not know (kept as an unknown action), `E`
    /// `p:extLst`, `X` a child from an unknown namespace.
    const EVENT_CHILD_ORDER: &[(&str, &str)] = &[
        ("", "Invalid: task event requires atrbtn"),
        ("A", "ok anchor=false action=none extension=false"),
        ("N", "Invalid: task event requires atrbtn"),
        ("T", "Invalid: task event requires atrbtn"),
        ("E", "Invalid: task event requires atrbtn"),
        (
            "AA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        ("AN", "ok anchor=true action=none extension=false"),
        ("AT", "ok anchor=false action=add extension=false"),
        ("AE", "ok anchor=false action=none extension=true"),
        (
            "NA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "NN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        ("NT", "Invalid: task event requires atrbtn"),
        ("NE", "Invalid: task event requires atrbtn"),
        (
            "TA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "TN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "TT",
            "Invalid: task event action is duplicated or out of order",
        ),
        ("TE", "Invalid: task event requires atrbtn"),
        (
            "EA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "EN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "ET",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "EE",
            "Invalid: task event extLst is duplicated or out of order",
        ),
        (
            "AAA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "AAN",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "AAT",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "AAE",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "ANA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "ANN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        ("ANT", "ok anchor=true action=add extension=false"),
        ("ANE", "ok anchor=true action=none extension=true"),
        (
            "ATA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "ATN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "ATT",
            "Invalid: task event action is duplicated or out of order",
        ),
        ("ATE", "ok anchor=false action=add extension=true"),
        (
            "AEA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "AEN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "AET",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "AEE",
            "Invalid: task event extLst is duplicated or out of order",
        ),
        (
            "NAA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "NAN",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "NAT",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "NAE",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "NNA",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "NNN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "NNT",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "NNE",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "NTA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "NTN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "NTT",
            "Invalid: task event action is duplicated or out of order",
        ),
        ("NTE", "Invalid: task event requires atrbtn"),
        (
            "NEA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "NEN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "NET",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "NEE",
            "Invalid: task event extLst is duplicated or out of order",
        ),
        (
            "TAA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "TAN",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "TAT",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "TAE",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "TNA",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "TNN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "TNT",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "TNE",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "TTA",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "TTN",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "TTT",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "TTE",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "TEA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "TEN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "TET",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "TEE",
            "Invalid: task event extLst is duplicated or out of order",
        ),
        (
            "EAA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "EAN",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "EAT",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "EAE",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        (
            "ENA",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "ENN",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "ENT",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "ENE",
            "Invalid: task event anchr is duplicated or out of order",
        ),
        (
            "ETA",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "ETN",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "ETT",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "ETE",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "EEA",
            "Invalid: task event extLst is duplicated or out of order",
        ),
        (
            "EEN",
            "Invalid: task event extLst is duplicated or out of order",
        ),
        (
            "EET",
            "Invalid: task event extLst is duplicated or out of order",
        ),
        (
            "EEE",
            "Invalid: task event extLst is duplicated or out of order",
        ),
        ("ANTE", "ok anchor=true action=add extension=true"),
        ("L", "Invalid: task event requires atrbtn"),
        ("U", "Invalid: task event requires atrbtn"),
        ("AL", "ok anchor=false action=title extension=false"),
        ("AU", "ok anchor=false action=unknown extension=false"),
        (
            "ATL",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "ALT",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "AUT",
            "Invalid: task event action is duplicated or out of order",
        ),
        (
            "ATU",
            "Invalid: task event action is duplicated or out of order",
        ),
        ("AUE", "ok anchor=false action=unknown extension=true"),
        (
            "UA",
            "Invalid: task event atrbtn is duplicated or out of order",
        ),
        ("X", "Invalid: unexpected task history event child"),
        ("AX", "Invalid: unexpected task history event child"),
        ("XA", "Invalid: unexpected task history event child"),
        ("AEX", "Invalid: unexpected task history event child"),
    ];

    fn event_outcome(key: &str) -> String {
        let mut children = String::new();
        for letter in key.chars() {
            children.push_str(&match letter {
                'A' => format!(r#"<p228:atrbtn authorId="{GUID}"/>"#),
                'N' => format!(r#"<p228:anchr><p228:comment id="{GUID}"/></p228:anchr>"#),
                'T' => "<p228:add/>".to_owned(),
                'L' => r#"<p228:title val="Review"/>"#.to_owned(),
                'U' => "<p228:future/>".to_owned(),
                'E' => "<p:extLst/>".to_owned(),
                'X' => r#"<q:other xmlns:q="urn:example:other"/>"#.to_owned(),
                other => panic!("unknown key letter {other:?}"),
            });
        }
        let xml = format!(
            r#"<p228:taskDetails xmlns:p228="{P228}" xmlns:p="{P}"><p228:history><p228:event time="2024-01-02T03:04:05Z" id="{GUID}">{children}</p228:event></p228:history></p228:taskDetails>"#
        );
        match parse_task_details(xml.as_bytes()) {
            Ok(value) => {
                let event = value.history.events.first().unwrap();
                let action = match &event.action {
                    None => "none",
                    Some(Action::Add) => "add",
                    Some(Action::Title(_)) => "title",
                    Some(Action::Unknown(_)) => "unknown",
                    Some(_) => "other",
                };
                format!(
                    "ok anchor={} action={action} extension={}",
                    event.anchor.is_some(),
                    event.extension_xml.is_some()
                )
            },
            Err(Error::Invalid(message)) => format!("Invalid: {message}"),
            Err(other) => format!("{other:?}"),
        }
    }

    #[test]
    fn event_children_keep_their_order_and_duplicate_rules() {
        assert_matrix(EVENT_CHILD_ORDER, "ANTE", event_outcome);
    }
}
