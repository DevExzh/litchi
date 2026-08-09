#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Glossary package planning and failure-bounded XML assembly.

use super::super::MAX;
use super::super::model::{Catalog, Conformance};
use super::super::{Arc, Error, Result};
use super::semantic::{prepare_opaque_for, write_entry};
use super::validation::{invalid, validate_catalog_fields};
use super::xml::{Node, XmlSink, XmlSize, node_write};

pub(in crate::glossary) struct WritePlan {
    pub(in crate::glossary) background: Option<Node>,
    pub(in crate::glossary) bodies: Vec<Option<Node>>,
    pub(in crate::glossary) producer_entries: Vec<Option<Arc<str>>>,
    pub(in crate::glossary) bytes: usize,
}

pub(in crate::glossary) fn add_sizes(left: [usize; 2], right: [usize; 2]) -> Result<[usize; 2]> {
    Ok([
        left[0]
            .checked_add(right[0])
            .ok_or_else(|| invalid("Transitional glossary size overflow"))?,
        left[1]
            .checked_add(right[1])
            .ok_or_else(|| invalid("Strict glossary size overflow"))?,
    ])
}

pub(in crate::glossary) fn replace_sizes(
    total: [usize; 2],
    old: [usize; 2],
    new: [usize; 2],
) -> Result<[usize; 2]> {
    Ok([
        total[0]
            .checked_sub(old[0])
            .and_then(|value| value.checked_add(new[0]))
            .ok_or_else(|| invalid("Transitional glossary replacement size overflow"))?,
        total[1]
            .checked_sub(old[1])
            .and_then(|value| value.checked_add(new[1]))
            .ok_or_else(|| invalid("Strict glossary replacement size overflow"))?,
    ])
}

pub(in crate::glossary) fn validate_catalog_sizes(
    entry_bytes: [usize; 2],
    background_bytes: [usize; 2],
    entry_count: usize,
) -> Result<()> {
    for conformance in [Conformance::Transitional, Conformance::Strict] {
        let mut size = XmlSize::default();
        write_catalog_open(&mut size, conformance)?;
        if entry_count != 0 {
            size.push_str("<w:docParts></w:docParts>")?;
        }
        size.push_str("</w:glossaryDocument>")?;
        let index = conformance.index();
        let bytes = size
            .bytes
            .checked_add(background_bytes[index])
            .and_then(|bytes| bytes.checked_add(entry_bytes[index]))
            .ok_or_else(|| invalid("glossary catalog size overflow"))?;
        if bytes > MAX {
            return Err(invalid("serialized glossary document exceeds 32 MiB"));
        }
    }
    Ok(())
}

pub(in crate::glossary) fn write_catalog_open<X: XmlSink>(
    x: &mut X,
    conformance: Conformance,
) -> Result<()> {
    x.push_str(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument xmlns:w=""#,
    )?;
    x.push_str(conformance.word())?;
    x.push_str(r#"" xmlns:r=""#)?;
    x.push_str(conformance.relationships())?;
    x.push_str(r#"">"#)
}

pub(in crate::glossary) fn plan_write(
    value: &Catalog,
    conformance: Conformance,
) -> Result<WritePlan> {
    validate_catalog_fields(value)?;
    let background = value
        .background
        .as_deref()
        .map(|xml| prepare_opaque_for(xml, "background", conformance))
        .transpose()?;
    let mut bodies = Vec::new();
    bodies
        .try_reserve_exact(value.entries.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary XML write plan",
            source,
        })?;
    let mut producer_entries = Vec::new();
    producer_entries
        .try_reserve_exact(value.entries.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary producer write plan",
            source,
        })?;
    for entry in &value.entries {
        let producer = entry
            .producer
            .as_deref()
            .filter(|producer| producer.conformance == conformance);
        producer_entries.push(producer.map(|producer| Arc::clone(&producer.xml)));
        bodies.push(if producer.is_some() {
            None
        } else {
            entry
                .body
                .as_deref()
                .map(|xml| prepare_opaque_for(xml, "docPartBody", conformance))
                .transpose()?
        });
    }
    let mut plan = WritePlan {
        background,
        bodies,
        producer_entries,
        bytes: 0,
    };
    let mut size = XmlSize::default();
    emit_catalog(&mut size, value, conformance, &plan)?;
    plan.bytes = size.bytes;
    Ok(plan)
}

pub(in crate::glossary) fn emit_catalog<X: XmlSink>(
    x: &mut X,
    value: &Catalog,
    conformance: Conformance,
    plan: &WritePlan,
) -> Result<()> {
    write_catalog_open(x, conformance)?;
    match (&value.background, &plan.background) {
        (Some(_), Some(background)) => {
            node_write(x, background, conformance == Conformance::Strict)?;
        },
        (None, None) => {},
        _ => return Err(invalid("glossary background write plan is inconsistent")),
    }
    let mut bodies = plan.bodies.iter();
    let mut producer_entries = plan.producer_entries.iter();
    if !value.entries.is_empty() {
        x.push_str("<w:docParts>")?;
        for entry in &value.entries {
            let body = bodies
                .next()
                .ok_or_else(|| invalid("glossary entry write plan is incomplete"))?;
            let producer = producer_entries
                .next()
                .ok_or_else(|| invalid("glossary producer write plan is incomplete"))?;
            if let Some(producer) = producer {
                if body.is_some() {
                    return Err(invalid("glossary producer write plan has a duplicate body"));
                }
                x.push_str(producer)?;
            } else {
                write_entry(x, entry, body.as_ref(), conformance)?;
            }
        }
        x.push_str("</w:docParts>")?;
    }
    if bodies.next().is_some() {
        return Err(invalid("glossary entry write plan has unused bodies"));
    }
    if producer_entries.next().is_some() {
        return Err(invalid("glossary producer write plan has unused entries"));
    }
    x.push_str("</w:glossaryDocument>")?;
    Ok(())
}
