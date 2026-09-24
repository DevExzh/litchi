//! Package-level aggregation for variable declarations.
//!
//! Parsing each XML part remains a codec concern; this boundary owns the
//! document-wide declaration/reference inventory and its cross-part limits.

use super::{Kind, MAX_XML_BYTES, Part, Scope, codec, model::Declarations};
use crate::core::ResolvedReader;
use crate::generic::{FlatMutationBudget, MemoryLease};
use litchi_core::Result;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use std::collections::HashSet;
use std::mem::size_of;

/// Parse content, styles, or flat XML parts as one validated document view.
pub(crate) fn parse_parts(parts: &[(&str, Part)]) -> Result<Declarations> {
    parse_parts_with_optional_budget(parts, None)
}

fn parse_parts_with_optional_budget(
    parts: &[(&str, Part)],
    budget: Option<&FlatMutationBudget>,
) -> Result<Declarations> {
    let total = parts.iter().try_fold(0usize, |size, (xml, _)| {
        size.checked_add(xml.len())
            .ok_or_else(|| codec::invalid("variable declaration XML size overflow"))
    })?;
    if total > MAX_XML_BYTES {
        return Err(codec::invalid("variable declaration XML exceeds 64 MiB"));
    }

    let mut result = Declarations::default();
    let mut names = HashSet::<(Kind, String)>::new();
    let mut containers = HashSet::<(Part, Scope, Kind)>::new();
    let mut uses = HashSet::new();
    let mut all_uses = Vec::<(Kind, String)>::new();
    let mut aggregate = 0usize;
    let mut declaration_count = 0usize;
    for (xml, part) in parts {
        codec::parse_part(
            xml,
            *part,
            &mut result,
            &mut names,
            &mut containers,
            &mut uses,
            &mut all_uses,
            &mut aggregate,
            &mut declaration_count,
            budget,
        )?;
    }
    for (kind, name) in all_uses {
        if let Some(budget) = budget {
            budget.check()?;
            budget.consume_objects(1)?;
        }
        if !names.contains(&(kind, name.clone())) {
            return Err(codec::invalid(format!(
                "ODF {kind:?} variable '{name}' is used without a declaration"
            )));
        }
    }
    if let Some(budget) = budget {
        budget.check()?;
        budget.consume_objects(1)?;
    }
    let dde = crate::dde_connection::parse_dde_connection_parts_with_budget(parts, budget)?;
    result.dde_connections = dde.declarations;
    result.dde_connection_uses = dde.uses;
    if let Some(budget) = budget {
        budget.check()?;
        budget.consume_objects(1)?;
    }
    result.bibliography_configuration =
        crate::bibliography_configuration::parse_bibliography_configuration_parts_with_budget(
            parts, budget,
        )?;
    if let Some(budget) = budget {
        budget.check()?;
        budget.consume_objects(1)?;
    }
    result.auto_mark_files =
        crate::auto_mark_file::parse_auto_mark_file_parts_with_budget(parts, budget)?;
    if let Some(budget) = budget {
        budget.check()?;
        budget.consume_objects(1)?;
    }
    Ok(result)
}

/// Parse declaration parts with a checked projection-memory reservation.
///
/// Variable parsing retains decoded names, attributes, declaration values,
/// scope frames, and cross-reference sets.  Charging the source XML length
/// alone misses those owned structures.  The plan below counts event, name,
/// attribute, text, vector-slot, and set-entry storage from a borrowed event
/// pass before the allocating parser runs.  It remains independent of the
/// final edited-document byte limit.
pub(crate) fn parse_parts_with_budget(
    parts: &[(&str, Part)],
    budget: &FlatMutationBudget,
) -> Result<(Declarations, MemoryLease)> {
    budget.check()?;
    let amount = parser_memory_plan(parts, Some(budget))?;
    let reservation = budget.reserve_bytes(amount, "ODT variable declaration parser projection")?;
    match parse_parts_with_optional_budget(parts, Some(budget)) {
        Ok(parsed) => Ok((parsed, MemoryLease::new(reservation))),
        Err(error) => {
            drop(reservation);
            Err(error)
        },
    }
}

#[derive(Default)]
struct VariableMemoryPlan {
    events: usize,
    max_event: usize,
    elements: usize,
    names: usize,
    attributes: usize,
    attribute_bytes: usize,
    text_bytes: usize,
    dde_declarations: usize,
    dde_uses: usize,
    auto_mark_references: usize,
    bibliography_configurations: usize,
    bibliography_sort_keys: usize,
}

pub(crate) fn parser_memory_plan(
    parts: &[(&str, Part)],
    budget: Option<&FlatMutationBudget>,
) -> Result<usize> {
    let total = parts.iter().try_fold(0usize, |size, (xml, _)| {
        size.checked_add(xml.len())
            .ok_or_else(|| codec::invalid("variable declaration XML size overflow"))
    })?;
    if total > MAX_XML_BYTES {
        return Err(codec::invalid("variable declaration XML exceeds 64 MiB"));
    }
    let mut plan = VariableMemoryPlan::default();
    for (xml, _) in parts {
        if xml.len() > MAX_XML_BYTES {
            return Err(codec::invalid("variable declaration XML exceeds 64 MiB"));
        }
        let mut reader = ResolvedReader::from_xml(xml);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        loop {
            if let Some(budget) = budget {
                budget.event(depth)?;
            }
            let start = reader.buffer_position() as usize;
            let (resolved, event) = reader.read_resolved_event().map_err(|error| {
                codec::invalid(format!("invalid variable declaration XML: {error}"))
            })?;
            let namespace_len = match &resolved {
                ResolveResult::Bound(value) => value.as_ref().len(),
                ResolveResult::Unbound => 0,
                ResolveResult::Unknown(prefix) => {
                    return Err(codec::invalid(format!(
                        "unbound XML namespace prefix '{}'",
                        String::from_utf8_lossy(prefix)
                    )));
                },
            };
            drop(resolved);
            let end = reader.buffer_position() as usize;
            plan.events = plan
                .events
                .checked_add(1)
                .ok_or_else(|| codec::invalid("variable declaration event count overflow"))?;
            plan.max_event = plan.max_event.max(end.saturating_sub(start));
            let is_start = matches!(&event, Event::Start(_));
            match event {
                Event::Start(source) | Event::Empty(source) => {
                    plan.elements = plan
                        .elements
                        .checked_add(1)
                        .ok_or_else(|| codec::invalid("variable declaration element overflow"))?;
                    plan.names = plan
                        .names
                        .checked_add(namespace_len + source.local_name().as_ref().len())
                        .ok_or_else(|| codec::invalid("variable declaration name overflow"))?;
                    match source.local_name().as_ref() {
                        b"dde-connection-decl" => {
                            plan.dde_declarations = plan
                                .dde_declarations
                                .checked_add(1)
                                .ok_or_else(|| codec::invalid("DDE declaration count overflow"))?;
                        },
                        b"dde-connection" => {
                            plan.dde_uses = plan
                                .dde_uses
                                .checked_add(1)
                                .ok_or_else(|| codec::invalid("DDE use count overflow"))?;
                        },
                        b"alphabetical-index-auto-mark-file" => {
                            plan.auto_mark_references = plan
                                .auto_mark_references
                                .checked_add(1)
                                .ok_or_else(|| codec::invalid("auto-mark count overflow"))?;
                        },
                        b"bibliography-configuration" => {
                            plan.bibliography_configurations = plan
                                .bibliography_configurations
                                .checked_add(1)
                                .ok_or_else(|| codec::invalid("bibliography count overflow"))?;
                        },
                        b"sort-key" => {
                            plan.bibliography_sort_keys =
                                plan.bibliography_sort_keys.checked_add(1).ok_or_else(|| {
                                    codec::invalid("bibliography sort-key count overflow")
                                })?;
                        },
                        _ => {},
                    }
                    for attribute in source.attributes() {
                        let attribute = attribute.map_err(|error| {
                            codec::invalid(format!(
                                "invalid variable declaration attribute: {error}"
                            ))
                        })?;
                        plan.attributes = plan.attributes.checked_add(1).ok_or_else(|| {
                            codec::invalid("variable declaration attribute overflow")
                        })?;
                        plan.attribute_bytes = plan
                            .attribute_bytes
                            .checked_add(attribute.key.as_ref().len())
                            .and_then(|value| value.checked_add(attribute.value.len()))
                            .ok_or_else(|| {
                                codec::invalid("variable declaration attribute size overflow")
                            })?;
                    }
                    if is_start {
                        depth = depth
                            .checked_add(1)
                            .ok_or_else(|| codec::invalid("variable declaration depth overflow"))?;
                        if let Some(budget) = budget {
                            budget.observe_depth(depth)?;
                        }
                    } else if let Some(budget) = budget {
                        budget.observe_depth(depth.saturating_add(1))?;
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| codec::invalid("variable declaration stack underflow"))?;
                },
                Event::Text(value) => {
                    plan.text_bytes = plan
                        .text_bytes
                        .checked_add(value.as_ref().len())
                        .ok_or_else(|| codec::invalid("variable declaration text overflow"))?;
                },
                Event::CData(value) => {
                    plan.text_bytes = plan
                        .text_bytes
                        .checked_add(value.as_ref().len())
                        .ok_or_else(|| codec::invalid("variable declaration text overflow"))?;
                },
                Event::DocType(_) => {
                    return Err(codec::invalid("DTDs are not allowed in declaration XML"));
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if depth != 0 {
            return Err(codec::invalid(
                "incomplete variable declaration XML structure",
            ));
        }
    }
    // Each parser vector grows with its previous allocation live until the
    // replacement allocation succeeds. All codecs request exact growth, so
    // charge the destination plus that predecessor for an explicit
    // old-plus-new peak. This is aggregate parser memory, independent of the
    // final document/output byte cap.
    let declaration_slots = plan
        .elements
        .checked_mul(size_of::<(Option<String>, String, Option<String>)>())
        .and_then(|value| value.checked_mul(2))
        .ok_or_else(|| codec::invalid("variable declaration declaration-vector plan overflow"))?;
    let attribute_slots = plan
        .attributes
        .checked_mul(size_of::<(String, String, String)>())
        .and_then(|value| value.checked_mul(2))
        .ok_or_else(|| codec::invalid("variable declaration attribute-vector plan overflow"))?;
    let event_slots = plan
        .events
        .checked_mul(size_of::<usize>())
        .and_then(|value| value.checked_mul(2))
        .ok_or_else(|| codec::invalid("variable declaration event-vector plan overflow"))?;
    let vector_slots = declaration_slots
        .checked_add(attribute_slots)
        .and_then(|value| value.checked_add(event_slots))
        .ok_or_else(|| codec::invalid("variable declaration vector plan overflow"))?;
    let strings = plan
        // The source is visited by the variable codec and by the three inert
        // auxiliary codecs below. Each independently resolves names and owns
        // decoded attribute keys and values while building its projection.
        .names
        .checked_mul(4)
        .and_then(|value| {
            plan.attribute_bytes
                .checked_mul(8)
                .and_then(|amount| value.checked_add(amount))
        })
        .and_then(|value| value.checked_add(plan.text_bytes))
        .ok_or_else(|| codec::invalid("variable declaration string plan overflow"))?;
    let set_slots = plan
        .elements
        .checked_add(plan.attributes)
        .and_then(|value| value.checked_mul(usize::BITS as usize * 4))
        .and_then(|value| value.checked_mul(2))
        .ok_or_else(|| codec::invalid("variable declaration set plan overflow"))?;
    let parser_frames = plan
        .events
        .checked_mul(size_of::<(Option<String>, String, Option<String>)>())
        .and_then(|value| value.checked_mul(8))
        .ok_or_else(|| codec::invalid("variable declaration frame plan overflow"))?;
    let auxiliary_model_slots = plan
        .dde_declarations
        .checked_mul(size_of::<crate::dde_connection::Declaration>())
        .and_then(|value| {
            plan.dde_uses
                .checked_mul(size_of::<crate::dde_connection::Use>())
                .and_then(|amount| value.checked_add(amount))
        })
        .and_then(|value| {
            plan.auto_mark_references
                .checked_mul(size_of::<
                    crate::auto_mark_file::AlphabeticalIndexAutoMarkFile,
                >())
                .and_then(|amount| value.checked_add(amount))
        })
        .and_then(|value| {
            plan.bibliography_configurations
                .checked_mul(size_of::<crate::bibliography_configuration::Configuration>())
                .and_then(|amount| value.checked_add(amount))
        })
        .and_then(|value| {
            plan.bibliography_sort_keys
                .checked_mul(size_of::<crate::bibliography_configuration::SortKey>())
                .and_then(|amount| value.checked_add(amount))
        })
        .and_then(|value| value.checked_mul(2))
        .ok_or_else(|| codec::invalid("auxiliary variable projection plan overflow"))?;
    vector_slots
        .checked_add(strings)
        .and_then(|value| value.checked_add(set_slots))
        .and_then(|value| value.checked_add(parser_frames))
        .and_then(|value| value.checked_add(auxiliary_model_slots))
        .and_then(|value| {
            plan.max_event
                .checked_mul(2)
                .and_then(|amount| value.checked_add(amount))
        })
        .and_then(|value| {
            plan.events
                .checked_mul(size_of::<usize>() * 4)
                .and_then(|amount| value.checked_add(amount))
        })
        .ok_or_else(|| codec::invalid("variable declaration parser memory plan overflow"))
}
