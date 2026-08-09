//! Row and cell snapshot writer.

use std::collections::BTreeMap;

use litchi_core::xml::escape_xml;
use litchi_sheet::{Cell as Address, Row};

use super::super::super::wire::{sibling_name, write_close, write_tag};
use super::super::model::{CellSlot, RowSlot, SheetData, Span, Tag};
use crate::cell::{Content, Value};
use crate::error::{Result, invalid};
use crate::outline::Outline;
use crate::raw::strings::encode_spreadsheet_text;
use crate::raw::worksheet::edit::model::{
    Action, DescentEffect, HeightEffect, Payload, RowAction, StyleEffect,
};

pub(crate) fn write_sheet_data(
    output: &mut Vec<u8>,
    source: &[u8],
    data: &SheetData,
    cells: BTreeMap<Address, Action>,
    rows: BTreeMap<Row, RowAction>,
    descent_name: &str,
) -> Result<()> {
    let mut by_row = BTreeMap::<u32, RowEdits>::new();
    for (address, action) in cells {
        by_row
            .entry(address.row().get() + 1)
            .or_default()
            .cells
            .insert(address, action);
    }
    for (row, action) in rows {
        by_row.entry(row.get() + 1).or_default().row = Some(action);
    }

    if data.empty {
        write_tag(output, &data.tag, false, &[], &[]);
        for (number, edits) in by_row {
            write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
        }
        write_close(output, &data.tag.name);
        return Ok(());
    }

    output.extend_from_slice(&source[data.span.start..data.tag_end]);
    let mut cursor = data.tag_end;
    let mut pending = by_row.into_iter().peekable();
    for row in &data.rows {
        output.extend_from_slice(&source[cursor..row.span.start]);
        while pending
            .peek()
            .is_some_and(|(number, _)| *number < row.number)
        {
            if let Some((number, edits)) = pending.next() {
                write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
            }
        }
        if pending
            .peek()
            .is_some_and(|(number, _)| *number == row.number)
        {
            let (_, edits) = pending
                .next()
                .ok_or_else(|| invalid("worksheet row edit ordering was lost"))?;
            write_row(output, source, row, &edits, descent_name)?;
        } else {
            output.extend_from_slice(&source[row.span.start..row.span.end]);
        }
        cursor = row.span.end;
    }
    output.extend_from_slice(&source[cursor..data.close_start]);
    for (number, edits) in pending {
        write_new_row(output, &data.tag.name, number, &edits, descent_name)?;
    }
    output.extend_from_slice(&source[data.close_start..data.span.end]);
    Ok(())
}

#[derive(Debug, Default)]
struct RowEdits {
    cells: BTreeMap<Address, Action>,
    row: Option<RowAction>,
}

fn write_row(
    output: &mut Vec<u8>,
    source: &[u8],
    row: &RowSlot,
    edits: &RowEdits,
    descent_name: &str,
) -> Result<()> {
    let actions = &edits.cells;
    let membership_changed = actions.iter().any(|(address, action)| {
        let exists = row
            .cells
            .binary_search_by_key(address, |cell| cell.address)
            .is_ok();
        (!exists && action.creates_missing()) || (exists && matches!(action, Action::Remove))
    });

    if row.empty {
        let creates_cell = actions.values().any(Action::creates_missing);
        let mut removed = Vec::new();
        let mut appended = Vec::new();
        if creates_cell {
            removed.extend(["spans", "r"]);
            appended.push(("r", row.number.to_string()));
        }
        if let Some(action) = edits.row {
            row_effect_attributes(
                action,
                row.descent_attribute.as_deref().unwrap_or(descent_name),
                &mut removed,
                &mut appended,
            );
        }
        write_tag(output, &row.tag, !creates_cell, &removed, &appended);
        if !creates_cell {
            return Ok(());
        }
        for (address, action) in actions {
            write_new_action(output, &row.tag.name, *address, action)?;
        }
        write_close(output, &row.tag.name);
        return Ok(());
    }

    if membership_changed || edits.row.is_some() {
        let mut removed = Vec::new();
        let mut appended = Vec::new();
        if membership_changed {
            removed.push("spans");
        }
        if let Some(action) = edits.row {
            row_effect_attributes(
                action,
                row.descent_attribute.as_deref().unwrap_or(descent_name),
                &mut removed,
                &mut appended,
            );
        }
        write_tag(output, &row.tag, false, &removed, &appended);
    } else {
        output.extend_from_slice(&source[row.span.start..row.tag_end]);
    }
    let mut cursor = row.tag_end;
    let mut pending = actions.iter().peekable();
    for cell in &row.cells {
        output.extend_from_slice(&source[cursor..cell.span.start]);
        while pending
            .peek()
            .is_some_and(|(address, _)| **address < cell.address)
        {
            let (address, action) = pending
                .next()
                .ok_or_else(|| invalid("worksheet cell edit ordering was lost"))?;
            write_new_action(output, &row.tag.name, *address, action)?;
        }
        if pending
            .peek()
            .is_some_and(|(address, _)| **address == cell.address)
        {
            let (_, action) = pending
                .next()
                .ok_or_else(|| invalid("worksheet cell edit ordering was lost"))?;
            match action {
                Action::Update { .. } => write_cell(output, source, cell, action)?,
                Action::Remove => {},
            }
        } else {
            output.extend_from_slice(&source[cell.span.start..cell.span.end]);
        }
        cursor = cell.span.end;
    }
    output.extend_from_slice(&source[cursor..row.close_start]);
    for (address, action) in pending {
        write_new_action(output, &row.tag.name, *address, action)?;
    }
    output.extend_from_slice(&source[row.close_start..row.span.end]);
    Ok(())
}

fn write_new_row(
    output: &mut Vec<u8>,
    sheet_data_name: &str,
    number: u32,
    edits: &RowEdits,
    descent_name: &str,
) -> Result<()> {
    let creates_cell = edits.cells.values().any(Action::creates_missing);
    let materializes = edits.row.is_some_and(RowAction::materializes);
    if !creates_cell && !materializes {
        return Ok(());
    }
    let name = sibling_name(sheet_data_name, "row");
    let tag = Tag {
        name: name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut appended = vec![("r", number.to_string())];
    let mut removed = Vec::new();
    if let Some(action) = edits.row {
        row_effect_attributes(action, descent_name, &mut removed, &mut appended);
    }
    write_tag(output, &tag, !creates_cell, &removed, &appended);
    if !creates_cell {
        return Ok(());
    }
    for (address, action) in &edits.cells {
        write_new_action(output, &name, *address, action)?;
    }
    write_close(output, &name);
    Ok(())
}

fn row_effect_attributes<'a>(
    action: RowAction,
    descent_name: &'a str,
    removed: &mut Vec<&'a str>,
    appended: &mut Vec<(&'a str, String)>,
) {
    if let Some(hidden) = action.hidden {
        removed.push("hidden");
        if hidden {
            appended.push(("hidden", "1".to_owned()));
        }
    }
    if let Some(height) = action.height {
        removed.extend(["ht", "customHeight"]);
        if let HeightEffect::Set(height) = height {
            appended.push(("ht", height.get().to_string()));
            appended.push(("customHeight", "1".to_owned()));
        }
    }
    if let Some(descent) = action.descent {
        removed.push(descent_name);
        if let DescentEffect::Set(value) = descent {
            appended.push((descent_name, value.get().to_string()));
        }
    }
    if let Some(style) = action.style {
        removed.extend(["s", "customFormat"]);
        if let StyleEffect::Set(key) = style {
            appended.push(("s", key.to_string()));
            appended.push(("customFormat", "1".to_owned()));
        }
    }
    if let Some(outline) = action.outline {
        removed.push("outlineLevel");
        if outline != Outline::NONE {
            appended.push(("outlineLevel", outline.get().to_string()));
        }
    }
    for (value, name) in [
        (action.collapsed, "collapsed"),
        (action.thick_top, "thickTop"),
        (action.thick_bottom, "thickBot"),
        (action.phonetic, "ph"),
    ] {
        if let Some(value) = value {
            removed.push(name);
            if value {
                appended.push((name, "1".to_owned()));
            }
        }
    }
}

fn write_cell(output: &mut Vec<u8>, source: &[u8], cell: &CellSlot, action: &Action) -> Result<()> {
    let Action::Update { payload, style } = action else {
        return Err(invalid("cannot rewrite a removed cell"));
    };
    let content = payload
        .as_ref()
        .filter(|payload| matches!(payload, Payload::Set(_) | Payload::SharedString { .. }));
    let cell_type = content.and_then(payload_type);
    let mut removed = vec!["r"];
    if payload.is_some() {
        removed.push("t");
    }
    if style.is_some() {
        removed.push("s");
    }
    let mut appended = vec![("r", cell.address.a1())];
    if let Some(cell_type) = cell_type {
        appended.push(("t", cell_type.to_owned()));
    }
    if let Some(StyleEffect::Set(key)) = style {
        appended.push(("s", key.to_string()));
    }
    let remains_empty = cell.empty && payload.is_none();
    write_tag(output, &cell.tag, remains_empty, &removed, &appended);
    if remains_empty {
        return Ok(());
    }
    if let Some(content) = content {
        write_payload(output, &cell.tag.name, content)?;
    }
    if !cell.empty {
        if payload.is_some() {
            copy_without(
                output,
                source,
                cell.tag_end,
                cell.close_start,
                &cell.primary,
            );
        } else {
            output.extend_from_slice(&source[cell.tag_end..cell.close_start]);
        }
    }
    write_close(output, &cell.tag.name);
    Ok(())
}

fn write_new_action(
    output: &mut Vec<u8>,
    row_name: &str,
    address: Address,
    action: &Action,
) -> Result<()> {
    let Action::Update { payload, style } = action else {
        return Ok(());
    };
    if !action.creates_missing() {
        return Ok(());
    }
    let content = payload
        .as_ref()
        .filter(|payload| matches!(payload, Payload::Set(_) | Payload::SharedString { .. }));
    let name = sibling_name(row_name, "c");
    let tag = Tag {
        name: name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut appended = vec![("r", address.a1())];
    if let Some(cell_type) = content.and_then(payload_type) {
        appended.push(("t", cell_type.to_owned()));
    }
    if let Some(StyleEffect::Set(key)) = style {
        appended.push(("s", key.to_string()));
    }
    let empty = content.is_none();
    write_tag(output, &tag, empty, &[], &appended);
    if let Some(content) = content {
        write_payload(output, &name, content)?;
        write_close(output, &name);
    }
    Ok(())
}

fn content_type(content: &Content) -> Option<&'static str> {
    match content {
        Content::Value(Value::Bool(_)) => Some("b"),
        Content::Value(Value::Text(_)) => Some("inlineStr"),
        Content::Value(Value::Date(_)) => Some("d"),
        Content::Value(Value::Error(_)) => Some("e"),
        Content::Value(Value::Number(_)) | Content::Formula(_) => None,
    }
}

fn payload_type(payload: &Payload) -> Option<&'static str> {
    match payload {
        Payload::Set(content) => content_type(content),
        Payload::SharedString { .. } => Some("s"),
        Payload::Clear | Payload::ClearIfPresent => None,
    }
}

fn write_payload(output: &mut Vec<u8>, cell_name: &str, payload: &Payload) -> Result<()> {
    match payload {
        Payload::Set(content) => write_content(output, cell_name, content),
        Payload::SharedString { index, .. } => {
            write_text_element(output, cell_name, "v", &index.to_string());
            Ok(())
        },
        Payload::Clear | Payload::ClearIfPresent => Ok(()),
    }
}

fn write_content(output: &mut Vec<u8>, cell_name: &str, content: &Content) -> Result<()> {
    match content {
        Content::Value(Value::Bool(value)) => {
            write_text_element(output, cell_name, "v", if *value { "1" } else { "0" });
        },
        Content::Value(Value::Number(value)) => {
            write_text_element(output, cell_name, "v", &escape_xml(value.as_str()));
        },
        Content::Value(Value::Text(value)) => {
            let inline = sibling_name(cell_name, "is");
            let text = sibling_name(cell_name, "t");
            output.extend_from_slice(b"<");
            output.extend_from_slice(inline.as_bytes());
            output.extend_from_slice(b"><");
            output.extend_from_slice(text.as_bytes());
            output.extend_from_slice(b" xml:space=\"preserve\">");
            output.extend_from_slice(escape_xml(&encode_spreadsheet_text(value)).as_bytes());
            output.extend_from_slice(b"</");
            output.extend_from_slice(text.as_bytes());
            output.extend_from_slice(b"></");
            output.extend_from_slice(inline.as_bytes());
            output.extend_from_slice(b">");
        },
        Content::Value(Value::Date(value)) => {
            require_xml_text(value)?;
            write_text_element(output, cell_name, "v", &escape_xml(value));
        },
        Content::Value(Value::Error(value)) => {
            require_xml_text(value.as_str())?;
            write_text_element(output, cell_name, "v", &escape_xml(value.as_str()));
        },
        Content::Formula(formula) => {
            require_xml_text(formula.text())?;
            write_text_element(output, cell_name, "f", &escape_xml(formula.text()));
        },
    }
    Ok(())
}

fn write_text_element(output: &mut Vec<u8>, cell_name: &str, local: &str, value: &str) {
    let name = sibling_name(cell_name, local);
    output.extend_from_slice(b"<");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
    output.extend_from_slice(value.as_bytes());
    output.extend_from_slice(b"</");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
}

fn require_xml_text(value: &str) -> Result<()> {
    if value.chars().all(|character| {
        matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
            || ('\u{20}'..='\u{D7FF}').contains(&character)
            || ('\u{E000}'..='\u{FFFD}').contains(&character)
            || ('\u{10000}'..='\u{10FFFF}').contains(&character)
    }) {
        Ok(())
    } else {
        Err(invalid(
            "cell content contains a character forbidden by XML 1.0",
        ))
    }
}

fn copy_without(output: &mut Vec<u8>, source: &[u8], start: usize, end: usize, removed: &[Span]) {
    let mut cursor = start;
    for span in removed {
        output.extend_from_slice(&source[cursor..span.start]);
        cursor = span.end;
    }
    output.extend_from_slice(&source[cursor..end]);
}
