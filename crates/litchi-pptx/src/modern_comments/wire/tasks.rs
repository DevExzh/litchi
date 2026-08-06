use super::super::semantic::{
    OpaqueXml, TaskAction, TaskAnchor, TaskAssign, TaskDetails, TaskEvent, TaskHistory,
    TaskSchedule, TaskTitle, TaskUndo, TaskUser,
};
use super::super::{P, P228};
use super::xml::scan_with_context;
use super::xml::{Fragment, attr, attribute, close, no_attributes, only_attributes, open, scan};
use crate::{Error, Result};
use std::str::FromStr;

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn parse_task_details(xml: &[u8]) -> Result<TaskDetails> {
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
                extension = Some(OpaqueXml::new(child.xml.clone())?)
            },
            _ => return Err(invalid("unexpected task details child")),
        }
    }
    let value = TaskDetails {
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
) -> Result<TaskHistory> {
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
    let value = TaskHistory { events };
    value.validate()?;
    Ok(value)
}

fn parse_event(
    fragment: &Fragment,
    context: &[super::super::model::NamespaceDeclaration],
) -> Result<TaskEvent> {
    only_attributes(&fragment.attributes, &["time", "id"], "task history event")?;
    let time = attribute(&fragment.attributes, "time", true)?
        .unwrap()
        .to_owned();
    let id = attribute(&fragment.attributes, "id", true)?
        .unwrap()
        .to_owned();
    let mut attributed_by = None;
    let mut anchor = None;
    let mut action = None;
    let mut extension = None;
    let mut order = 0u8;
    let event_scan = scan_with_context(&fragment.xml, "task history event", context)?;
    for child in &event_scan.children {
        if child.namespace == P228 && child.local == "atrbtn" {
            if order != 0 || attributed_by.is_some() {
                return Err(invalid("task event atrbtn is duplicated or out of order"));
            }
            attributed_by = Some(parse_user(child)?);
            order = 1;
        } else if child.namespace == P228 && child.local == "anchr" {
            if order > 1 || anchor.is_some() {
                return Err(invalid("task event anchr is duplicated or out of order"));
            }
            anchor = Some(parse_anchor(child, &event_scan.namespaces)?);
            order = 1;
        } else if child.namespace == P && child.local == "extLst" {
            if order > 2 || extension.is_some() {
                return Err(invalid("task event extLst is duplicated or out of order"));
            }
            extension = Some(OpaqueXml::new(child.xml.clone())?);
            order = 3;
        } else if child.namespace == P228 {
            if action.is_some() || order == 3 {
                return Err(invalid("task event action is duplicated or out of order"));
            }
            action = Some(parse_action(child)?);
            order = 2;
        } else {
            return Err(invalid("unexpected task history event child"));
        }
    }
    let value = TaskEvent {
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

fn parse_user(fragment: &Fragment) -> Result<TaskUser> {
    only_attributes(&fragment.attributes, &["authorId"], "task user")?;
    let author_id = attribute(&fragment.attributes, "authorId", true)?
        .unwrap()
        .to_owned();
    Ok(TaskUser { author_id })
}

fn parse_anchor(
    fragment: &Fragment,
    context: &[super::super::model::NamespaceDeclaration],
) -> Result<TaskAnchor> {
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
                        .unwrap()
                        .to_owned(),
                );
            },
            (P, "extLst") if extension.is_none() => {
                extension = Some(OpaqueXml::new(child.xml.clone())?);
            },
            _ => return Err(invalid("unexpected task anchor child")),
        }
    }
    Ok(TaskAnchor {
        comment_id: comment_id.ok_or_else(|| invalid("task anchor requires comment"))?,
        extension_xml: extension,
        namespace_declarations: context.to_vec(),
    })
}

fn parse_action(fragment: &Fragment) -> Result<TaskAction> {
    match fragment.local.as_str() {
        "asgn" => {
            only_attributes(&fragment.attributes, &["authorId"], "task assignment")?;
            Ok(TaskAction::Assign(TaskAssign {
                author_id: attribute(&fragment.attributes, "authorId", true)?
                    .unwrap()
                    .to_owned(),
            }))
        },
        "add" => {
            no_attributes(&fragment.attributes, "task add")?;
            Ok(TaskAction::Add)
        },
        "title" => {
            only_attributes(&fragment.attributes, &["val"], "task title")?;
            Ok(TaskAction::Title(TaskTitle {
                value: attribute(&fragment.attributes, "val", true)?
                    .unwrap()
                    .to_owned(),
            }))
        },
        "date" => {
            only_attributes(&fragment.attributes, &["stDt", "endDt"], "task schedule")?;
            Ok(TaskAction::Schedule(TaskSchedule {
                start_date: attribute(&fragment.attributes, "stDt", false)?.map(str::to_owned),
                end_date: attribute(&fragment.attributes, "endDt", false)?.map(str::to_owned),
            }))
        },
        "pcntCmplt" => {
            only_attributes(&fragment.attributes, &["val"], "task progress")?;
            let value = attribute(&fragment.attributes, "val", true)?.unwrap();
            Ok(TaskAction::Progress(
                super::super::model::Progress::from_str(value)?,
            ))
        },
        "unasgnAll" => {
            no_attributes(&fragment.attributes, "task unassign-all")?;
            Ok(TaskAction::UnassignAll)
        },
        "undo" => {
            only_attributes(&fragment.attributes, &["id"], "task undo")?;
            Ok(TaskAction::Undo(TaskUndo {
                event_id: attribute(&fragment.attributes, "id", true)?
                    .unwrap()
                    .to_owned(),
            }))
        },
        "unknown" => {
            no_attributes(&fragment.attributes, "unknown task event")?;
            Ok(TaskAction::Unknown(OpaqueXml::new(fragment.xml.clone())?))
        },
        _ => Ok(TaskAction::Unknown(OpaqueXml::new(fragment.xml.clone())?)),
    }
}

pub(super) fn write_task_details(value: &TaskDetails) -> Result<Vec<u8>> {
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

fn write_history(out: &mut Vec<u8>, value: &TaskHistory) {
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

fn write_event(out: &mut Vec<u8>, value: &TaskEvent) {
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

fn write_anchor(out: &mut Vec<u8>, value: &TaskAnchor) {
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

fn write_action(out: &mut Vec<u8>, value: &TaskAction) {
    match value {
        TaskAction::Assign(value) => write_user(out, "asgn", &value.author_id),
        TaskAction::Add => empty(out, "add"),
        TaskAction::Title(value) => {
            open(out, "p228", "title");
            attr(out, "val", &value.value);
            out.extend_from_slice(b"/>");
        },
        TaskAction::Schedule(value) => {
            open(out, "p228", "date");
            if let Some(date) = &value.start_date {
                attr(out, "stDt", date);
            }
            if let Some(date) = &value.end_date {
                attr(out, "endDt", date);
            }
            out.extend_from_slice(b"/>");
        },
        TaskAction::Progress(value) => {
            open(out, "p228", "pcntCmplt");
            attr(out, "val", &value.thousandths().to_string());
            out.extend_from_slice(b"/>");
        },
        TaskAction::UnassignAll => empty(out, "unasgnAll"),
        TaskAction::Undo(value) => {
            open(out, "p228", "undo");
            attr(out, "id", &value.event_id);
            out.extend_from_slice(b"/>");
        },
        TaskAction::Unknown(value) => out.extend_from_slice(value.as_bytes()),
    }
}

fn empty(out: &mut Vec<u8>, local: &str) {
    open(out, "p228", local);
    out.extend_from_slice(b"/>");
}
