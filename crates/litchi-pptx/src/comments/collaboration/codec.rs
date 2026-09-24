//! Namespace-aware parsing and source splicing for the 2012 comment
//! collaboration extensions.

use std::borrow::Cow;
use std::collections::HashSet;
use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::mce::{
    Capabilities, Limits as MceLimits, NAMESPACE as MCE_NAMESPACE, OffsetLimits, active_offsets,
    process_markup_compatibility,
};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::ParentComment;
use super::{
    NAMESPACE, PRESENCE_EXTENSION_URI, PresenceInfo, THREADING_EXTENSION_URI, ThreadingInfo,
};
use crate::comments::{MAX_DEPTH, MAX_NODES, MAX_PART_BYTES, MAX_STRING_BYTES, PML, STRICT_PML};
use crate::{Error, Result};

pub(crate) const MAX_OFFSETS: usize = MAX_NODES;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Kind {
    Presence,
    Threading,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Owner {
    Author(u32),
    Comment { author_id: u32, index: u32 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Attr {
    pub(crate) key: String,
    pub(crate) value: String,
    pub(crate) value_range: Range<usize>,
    pub(crate) full_range: Range<usize>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Span {
    pub(crate) range: Range<usize>,
    pub(crate) start_end: usize,
    pub(crate) close_start: Option<usize>,
    pub(crate) qname: Vec<u8>,
    pub(crate) empty: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Node {
    pub(crate) span: Span,
    pub(crate) namespace: Vec<u8>,
    pub(crate) local: Vec<u8>,
    pub(crate) attrs: Vec<Attr>,
    pub(crate) parent: Option<usize>,
    pub(crate) children: Vec<usize>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Located {
    pub(crate) kind: Kind,
    pub(crate) owner: Owner,
    pub(crate) root: Span,
    pub(crate) item: Span,
    pub(crate) ext_list: Option<Span>,
    pub(crate) target_extension: Option<Span>,
    pub(crate) target_extensions_all: Vec<Span>,
    pub(crate) payload: Option<Span>,
    pub(crate) parent_comment: Option<Span>,
    pub(crate) payload_node: Option<Node>,
    pub(crate) parent_node: Option<Node>,
    pub(crate) value: Option<Value>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum Value {
    Presence(PresenceInfo),
    Threading(ThreadingInfo),
}

#[derive(Debug, Default)]
struct Scan {
    nodes: Vec<Node>,
    root: Option<usize>,
    processing_instructions: Vec<Range<usize>>,
    non_whitespace_text: Vec<(usize, Range<usize>)>,
}

/// Locate and parse one selected author or comment extension in a complete
/// legacy comment part. The source is never normalized by this operation.
pub(crate) fn locate(source: &[u8], kind: Kind, owner: Owner) -> Result<Located> {
    if source.len() > MAX_PART_BYTES {
        return Err(limit(
            "legacy comment collaboration XML bytes",
            MAX_PART_BYTES,
        ));
    }
    let scan = scan_raw(source)?;
    let root_index = scan
        .root
        .ok_or_else(|| invalid("legacy comment collaboration root is missing"))?;
    let root = &scan.nodes[root_index];
    let expected_root = match kind {
        Kind::Presence => b"cmAuthorLst".as_slice(),
        Kind::Threading => b"cmLst".as_slice(),
    };
    if !pml_namespace(&root.namespace) || root.local.as_slice() != expected_root {
        return Err(invalid(
            "legacy comment collaboration owner has the wrong root",
        ));
    }

    let capabilities = capabilities();
    let mce_limits = mce_limits();
    // The common MCE selector intentionally rejects processing instructions.
    // They are inert in this owner, so remove them only from the temporary
    // MCE input while retaining a mapping back to the original source offsets.
    let mce_source = MceSource::new(source, &scan.processing_instructions)?;
    let processed = process_markup_compatibility(mce_source.bytes(), &capabilities, &mce_limits)
        .map_err(Error::MarkupCompatibility)?
        .xml;
    if processed.len() > MAX_PART_BYTES {
        return Err(limit(
            "MCE-expanded legacy comment collaboration XML bytes",
            MAX_PART_BYTES,
        ));
    }
    // The existing comment parser validates the surrounding legacy comment
    // grammar and required fields. It is intentionally only a validator here;
    // all bytes used for the source-bound owner remain in `source`.
    match kind {
        Kind::Presence => {
            let _ = crate::comments::parse_comment_authors(&processed)?;
        },
        Kind::Threading => {
            let _ = crate::comments::parse_slide_comments(&processed)?;
        },
    }

    let item_indices = item_indices(&scan, root_index, kind, owner)?;
    let offsets = collect_offsets(&scan, &item_indices);
    let active = if offsets.is_empty() {
        HashSet::new()
    } else {
        let limits = OffsetLimits {
            max_source_bytes: MAX_PART_BYTES,
            max_offsets: MAX_OFFSETS,
            max_marked_bytes: MAX_PART_BYTES,
            processing: mce_limits,
        };
        let translated_offsets = mce_source.translate_offsets(&offsets)?;
        let selected = active_offsets(
            mce_source.bytes(),
            &translated_offsets,
            &capabilities,
            &limits,
        )
        .map_err(Error::MarkupCompatibility)?;
        selected
            .into_iter()
            .map(|offset| {
                translated_offsets
                    .binary_search(&offset)
                    .ok()
                    .and_then(|index| offsets.get(index).copied())
                    .ok_or_else(|| invalid("MCE active offset is not source-bound"))
            })
            .collect::<Result<HashSet<_>>>()?
    };
    let active_item = item_indices
        .iter()
        .copied()
        .filter(|index| is_active(&scan.nodes[*index].span, &active))
        .collect::<Vec<_>>();
    if active_item.len() != 1 {
        return Err(invalid(format!(
            "legacy comment collaboration owner has {} effective matching items",
            active_item.len()
        )));
    }
    let item_index = active_item[0];
    let item = &scan.nodes[item_index];

    let ext_lists = direct_non_mce_children(&scan, item_index)
        .filter(|index| {
            let node = &scan.nodes[*index];
            is_active(&node.span, &active)
                && pml_namespace(&node.namespace)
                && node.local.as_slice() == b"extLst"
        })
        .collect::<Vec<_>>();
    if ext_lists.len() > 1 {
        return Err(invalid(
            "legacy comment collaboration owner has duplicate extLst",
        ));
    }
    let ext_list = ext_lists
        .first()
        .map(|index| scan.nodes[*index].span.clone());

    let target_extensions_all = scan
        .nodes
        .iter()
        .enumerate()
        .filter(|(index, _node)| target_extension_owner(&scan, *index, kind) == Some(item_index))
        .map(|(_, node)| node.span.clone())
        .collect::<Vec<_>>();
    let active_targets = scan
        .nodes
        .iter()
        .enumerate()
        .filter(|(index, node)| {
            target_extension_owner(&scan, *index, kind) == Some(item_index)
                && is_active(&node.span, &active)
        })
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    if active_targets.len() > 1 {
        return Err(invalid(
            "legacy comment collaboration owner has duplicate typed extensions",
        ));
    }
    let target_index = active_targets.first().copied();
    let mut payloads = Vec::new();
    if let Some(target_index) = target_index {
        let target = &scan.nodes[target_index];
        for child in direct_non_mce_children(&scan, target_index) {
            let node = &scan.nodes[child];
            if !is_active(&node.span, &active) {
                continue;
            }
            if payload_name(node, kind) {
                payloads.push(child);
            } else {
                return Err(invalid(
                    "recognized legacy comment extension contains an unexpected child",
                ));
            }
        }
        if payloads.len() != 1 {
            return Err(invalid(
                "recognized legacy comment extension requires exactly one typed payload",
            ));
        }
        let payload_index = payloads[0];
        let parent_index = if kind == Kind::Threading {
            let parents = direct_non_mce_children(&scan, payload_index)
                .filter(|index| {
                    let node = &scan.nodes[*index];
                    is_active(&node.span, &active)
                        && p15_namespace(&node.namespace)
                        && node.local.as_slice() == b"parentCm"
                })
                .collect::<Vec<_>>();
            if parents.len() > 1 {
                return Err(invalid(
                    "threadingInfo contains duplicate parentCm elements",
                ));
            }
            for child in direct_non_mce_children(&scan, payload_index) {
                let node = &scan.nodes[child];
                if is_active(&node.span, &active)
                    && !(p15_namespace(&node.namespace) && node.local.as_slice() == b"parentCm")
                {
                    return Err(invalid("threadingInfo contains an unexpected child"));
                }
            }
            parents.first().copied()
        } else {
            if direct_non_mce_children(&scan, payload_index)
                .any(|index| is_active(&scan.nodes[index].span, &active))
            {
                return Err(invalid("presenceInfo must not contain child elements"));
            }
            None
        };
        let payload = &scan.nodes[payload_index];
        let value = match kind {
            Kind::Presence => {
                Value::Presence(parse_presence(payload, payload_index, &scan, &active)?)
            },
            Kind::Threading => Value::Threading(parse_threading(
                payload,
                payload_index,
                parent_index,
                &scan,
                &active,
            )?),
        };
        let parent_comment = parent_index.map(|index| scan.nodes[index].span.clone());
        return Ok(Located {
            kind,
            owner,
            root: root.span.clone(),
            item: item.span.clone(),
            ext_list,
            target_extension: Some(target.span.clone()),
            target_extensions_all,
            payload: Some(payload.span.clone()),
            parent_comment,
            payload_node: Some(payload.clone()),
            parent_node: parent_index.map(|index| scan.nodes[index].clone()),
            value: Some(value),
        });
    }
    Ok(Located {
        kind,
        owner,
        root: root.span.clone(),
        item: item.span.clone(),
        ext_list,
        target_extension: None,
        target_extensions_all,
        payload: None,
        parent_comment: None,
        payload_node: None,
        parent_node: None,
        value: None,
    })
}

/// Source-splice a typed value. A `None` value removes every recognized
/// branch for the selected item so an inactive MCE fallback cannot silently
/// resurrect the supposedly removed owner.
pub(crate) fn rewrite(source: &[u8], located: &Located, value: Option<Value>) -> Result<Vec<u8>> {
    if located.value == value {
        return copy_source(source);
    }
    if let Some(value) = &value {
        validate_value(located.kind, value)?;
    }
    let mut replacements = Vec::new();
    match (&located.value, value) {
        (Some(Value::Presence(before)), Some(Value::Presence(after))) => {
            if *before != after {
                let node = located
                    .payload_node
                    .as_ref()
                    .ok_or_else(|| invalid("presenceInfo payload source is missing"))?;
                let user_attr = node
                    .attrs
                    .iter()
                    .find(|attr| attr.key.as_bytes() == b"userId")
                    .ok_or_else(|| {
                        invalid("required typed collaboration attribute source is missing")
                    })?;
                let provider_attr = node
                    .attrs
                    .iter()
                    .find(|attr| attr.key.as_bytes() == b"providerId")
                    .ok_or_else(|| {
                        invalid("required typed collaboration attribute source is missing")
                    })?;
                let user_len = escaped_len(&after.user_id)?;
                let provider_len = escaped_len(&after.provider_id)?;
                preflight_replacement_sizes(
                    source,
                    &[
                        (&user_attr.value_range, user_len),
                        (&provider_attr.value_range, provider_len),
                    ],
                )?;
                replace_required_attribute(&mut replacements, node, b"userId", &after.user_id)?;
                replace_required_attribute(
                    &mut replacements,
                    node,
                    b"providerId",
                    &after.provider_id,
                )?;
            }
        },
        (Some(Value::Threading(before)), Some(Value::Threading(after))) => {
            let payload = located
                .payload
                .as_ref()
                .ok_or_else(|| invalid("threadingInfo payload source is missing"))?;
            let node = located
                .payload_node
                .as_ref()
                .ok_or_else(|| invalid("threadingInfo payload node source is missing"))?;
            match (&before.parent, &after.parent) {
                (None, Some(after_parent)) if node.span.empty => {
                    let before_bias = before.time_zone_bias.map(|value| value.to_string());
                    let after_bias = after.time_zone_bias.map(|value| value.to_string());
                    let bias_plan = optional_attribute_plan(
                        source,
                        node,
                        b"timeZoneBias",
                        before_bias.as_deref(),
                        after_bias.as_deref(),
                    )?;
                    let opening_len = if let Some(plan) = &bias_plan {
                        empty_opening_length(source, payload)?
                            .checked_sub(plan.range.len())
                            .and_then(|value| value.checked_add(plan.length))
                            .ok_or_else(|| {
                                invalid("self-closing threadingInfo opening size underflow")
                            })?
                    } else {
                        empty_opening_length(source, payload)?
                    };
                    let child_len = parent_length(qname_prefix(&payload.qname), after_parent)?;
                    let close_len = closing_length(&payload.qname)?;
                    let replacement_len = checked_sum(&[opening_len, child_len, close_len])?;
                    ensure_xml_size(replacement_len)?;
                    preflight_replacement_sizes(source, &[(&payload.range, replacement_len)])?;
                    replacements.push(Replacement {
                        range: payload.range.clone(),
                        value: expand_self_closing_payload(
                            source,
                            payload,
                            node,
                            after.time_zone_bias,
                            bias_plan.as_ref(),
                            after_parent,
                            replacement_len,
                        )?,
                    });
                },
                _ => match (&before.parent, &after.parent) {
                    (Some(_), Some(_)) => {
                        let parent_node = located
                            .parent_node
                            .as_ref()
                            .ok_or_else(|| invalid("parentCm node source is missing"))?;
                        let before_parent = before.parent.as_ref().expect("matched Some");
                        let after_parent = after.parent.as_ref().expect("matched Some");
                        let before_bias = before.time_zone_bias.map(|value| value.to_string());
                        let after_bias = after.time_zone_bias.map(|value| value.to_string());
                        let bias_plan = optional_attribute_plan(
                            source,
                            node,
                            b"timeZoneBias",
                            before_bias.as_deref(),
                            after_bias.as_deref(),
                        )?;
                        let before_author = before_parent.author_id.map(|value| value.to_string());
                        let after_author = after_parent.author_id.map(|value| value.to_string());
                        let author_plan = optional_attribute_plan(
                            source,
                            parent_node,
                            b"authorId",
                            before_author.as_deref(),
                            after_author.as_deref(),
                        )?;
                        let before_index = before_parent.index.map(|value| value.to_string());
                        let after_index = after_parent.index.map(|value| value.to_string());
                        let index_plan = optional_attribute_plan(
                            source,
                            parent_node,
                            b"idx",
                            before_index.as_deref(),
                            after_index.as_deref(),
                        )?;
                        preflight_size_plans(source, [bias_plan, author_plan, index_plan])?;
                        replace_optional_i32_attribute(
                            source,
                            &mut replacements,
                            node,
                            b"timeZoneBias",
                            before.time_zone_bias,
                            after.time_zone_bias,
                        )?;
                        replace_optional_u32_attribute(
                            source,
                            &mut replacements,
                            parent_node,
                            b"authorId",
                            before_parent.author_id,
                            after_parent.author_id,
                        )?;
                        replace_optional_u32_attribute(
                            source,
                            &mut replacements,
                            parent_node,
                            b"idx",
                            before_parent.index,
                            after_parent.index,
                        )?;
                    },
                    (Some(_), None) => {
                        let parent_comment = located
                            .parent_comment
                            .as_ref()
                            .ok_or_else(|| invalid("parentCm source is missing"))?;
                        let before_bias = before.time_zone_bias.map(|value| value.to_string());
                        let after_bias = after.time_zone_bias.map(|value| value.to_string());
                        let bias_plan = optional_attribute_plan(
                            source,
                            node,
                            b"timeZoneBias",
                            before_bias.as_deref(),
                            after_bias.as_deref(),
                        )?;
                        let parent_plan = Some(SizePlan {
                            range: parent_comment.range.clone(),
                            length: 0,
                        });
                        preflight_size_plans(source, [bias_plan, parent_plan])?;
                        replace_optional_i32_attribute(
                            source,
                            &mut replacements,
                            node,
                            b"timeZoneBias",
                            before.time_zone_bias,
                            after.time_zone_bias,
                        )?;
                        replacements.push(Replacement {
                            range: parent_comment.range.clone(),
                            value: Vec::new(),
                        });
                    },
                    (None, Some(after_parent)) => {
                        let before_bias = before.time_zone_bias.map(|value| value.to_string());
                        let after_bias = after.time_zone_bias.map(|value| value.to_string());
                        let bias_plan = optional_attribute_plan(
                            source,
                            node,
                            b"timeZoneBias",
                            before_bias.as_deref(),
                            after_bias.as_deref(),
                        )?;
                        let at = payload
                            .close_start
                            .ok_or_else(|| invalid("threadingInfo closing source is missing"))?;
                        let fragment_len =
                            parent_length(qname_prefix(&payload.qname), after_parent)?;
                        let parent_plan = Some(SizePlan {
                            range: at..at,
                            length: fragment_len,
                        });
                        preflight_size_plans(source, [bias_plan, parent_plan])?;
                        replace_optional_i32_attribute(
                            source,
                            &mut replacements,
                            node,
                            b"timeZoneBias",
                            before.time_zone_bias,
                            after.time_zone_bias,
                        )?;
                        let fragment = parent_fragment(qname_prefix(&payload.qname), after_parent)?;
                        insert_child(source, payload, fragment, &mut replacements)?;
                    },
                    (None, None) => {
                        let before_bias = before.time_zone_bias.map(|value| value.to_string());
                        let after_bias = after.time_zone_bias.map(|value| value.to_string());
                        let bias_plan = optional_attribute_plan(
                            source,
                            node,
                            b"timeZoneBias",
                            before_bias.as_deref(),
                            after_bias.as_deref(),
                        )?;
                        preflight_size_plans(source, [bias_plan])?;
                        replace_optional_i32_attribute(
                            source,
                            &mut replacements,
                            node,
                            b"timeZoneBias",
                            before.time_zone_bias,
                            after.time_zone_bias,
                        )?;
                    },
                },
            }
        },
        (Some(_), None) => {
            if located.target_extensions_all.is_empty() {
                return Err(invalid("typed collaboration extension source is missing"));
            }
            let ext_list = located.ext_list.as_ref();
            let remove_ext_list = ext_list.is_some_and(|ext_list| {
                only_whitespace_after_ranges(source, ext_list, &located.target_extensions_all)
            });
            let remove_item = remove_ext_list
                && ext_list.is_some_and(|ext_list| {
                    only_whitespace_after_ranges(
                        source,
                        &located.item,
                        std::slice::from_ref(ext_list),
                    )
                });
            if remove_item {
                replacements.push(Replacement {
                    range: located.item.range.clone(),
                    value: self_closing_opening(source, &located.item)?,
                });
            } else if let Some(ext_list) = ext_list.filter(|_| remove_ext_list) {
                replacements.push(Replacement {
                    range: ext_list.range.clone(),
                    value: Vec::new(),
                });
            } else {
                for range in &located.target_extensions_all {
                    replacements.push(Replacement {
                        range: range.range.clone(),
                        value: Vec::new(),
                    });
                }
            }
        },
        (None, Some(value)) => {
            let ext_list_prefix = located
                .ext_list
                .as_ref()
                .map(|span| qname_prefix(&span.qname))
                .unwrap_or_else(|| qname_prefix(&located.item.qname));
            let ext_qname_len = qualified_length(ext_list_prefix, b"ext")?;
            let payload_size = payload_length(located.kind, &value)?;
            let extension_len =
                extension_fragment_length(located.kind, ext_qname_len, payload_size)?;
            let ext_list_qname_len = located.ext_list.as_ref().map_or_else(
                || qualified_length(ext_list_prefix, b"extLst"),
                |span| Ok(span.qname.len()),
            )?;
            if let Some(ext_list) = &located.ext_list {
                if ext_list.empty {
                    let opening_len = empty_opening_length(source, ext_list)?;
                    let close_len = closing_length(&ext_list.qname)?;
                    let replacement_len = checked_sum(&[opening_len, extension_len, close_len])?;
                    ensure_xml_size(replacement_len)?;
                    preflight_replacement_sizes(source, &[(&ext_list.range, replacement_len)])?;
                } else {
                    let at = ext_list
                        .close_start
                        .ok_or_else(|| invalid("extLst closing source is missing"))?;
                    let insertion = at..at;
                    preflight_replacement_sizes(source, &[(&insertion, extension_len)])?;
                }
            } else {
                let ext_list_len = wrapped_fragment_length(ext_list_qname_len, extension_len)?;
                if located.item.empty {
                    let opening_len = empty_opening_length(source, &located.item)?;
                    let close_len = closing_length(&located.item.qname)?;
                    let replacement_len = checked_sum(&[opening_len, ext_list_len, close_len])?;
                    ensure_xml_size(replacement_len)?;
                    preflight_replacement_sizes(source, &[(&located.item.range, replacement_len)])?;
                } else {
                    let at = located
                        .item
                        .close_start
                        .ok_or_else(|| invalid("comment item closing source is missing"))?;
                    let insertion = at..at;
                    preflight_replacement_sizes(source, &[(&insertion, ext_list_len)])?;
                }
            }
            let ext_qname = qualified(ext_list_prefix, b"ext")?;
            let ext_list_qname: Cow<'_, [u8]> = match located.ext_list.as_ref() {
                Some(span) => Cow::Borrowed(span.qname.as_slice()),
                None => Cow::Owned(qualified(ext_list_prefix, b"extLst")?),
            };
            let extension = extension_fragment_for_value(located.kind, &ext_qname, &value)?;
            if let Some(ext_list) = &located.ext_list {
                if ext_list.empty {
                    let opening = empty_opening(source, ext_list)?;
                    let close = closing(&ext_list.qname)?;
                    let replacement = concat_fragments(&[&opening, &extension, &close])?;
                    replacements.push(Replacement {
                        range: ext_list.range.clone(),
                        value: replacement,
                    });
                } else {
                    let at = ext_list
                        .close_start
                        .ok_or_else(|| invalid("extLst closing source is missing"))?;
                    replacements.push(Replacement {
                        range: at..at,
                        value: extension,
                    });
                }
            } else {
                let ext_list_fragment = wrapped_fragment(&ext_list_qname, &extension)?;
                if located.item.empty {
                    let opening = empty_opening(source, &located.item)?;
                    let close = closing(&located.item.qname)?;
                    let replacement = concat_fragments(&[&opening, &ext_list_fragment, &close])?;
                    replacements.push(Replacement {
                        range: located.item.range.clone(),
                        value: replacement,
                    });
                } else {
                    let at = located
                        .item
                        .close_start
                        .ok_or_else(|| invalid("comment item closing source is missing"))?;
                    replacements.push(Replacement {
                        range: at..at,
                        value: ext_list_fragment,
                    });
                }
            }
        },
        (None, None) => unreachable!("semantic no-op was handled above"),
        (Some(Value::Presence(_)), Some(Value::Threading(_)))
        | (Some(Value::Threading(_)), Some(Value::Presence(_))) => {
            return Err(invalid("typed collaboration value kind changed"));
        },
    }
    apply_replacements(source, replacements)
}

fn scan_raw(source: &[u8]) -> Result<Scan> {
    let mut reader = NsReader::from_reader(source);
    let origin = ReaderOrigin::of(source);
    reader.config_mut().trim_text(false);
    let mut scan = Scan::default();
    let mut stack = Vec::new();
    let mut buffer = Vec::new();
    let mut nodes_seen = 0usize;
    loop {
        let before = position(&reader, origin)?;
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(xml_error)?
            .into_owned();
        let after = position(&reader, origin)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(limit("legacy comment collaboration XML depth", MAX_DEPTH));
                }
                let node = make_node(
                    &element,
                    namespace,
                    &reader,
                    reader.decoder(),
                    before,
                    after,
                    stack.last().copied(),
                )?;
                nodes_seen = nodes_seen.saturating_add(1);
                if nodes_seen > MAX_NODES {
                    return Err(limit("legacy comment collaboration XML nodes", MAX_NODES));
                }
                let index = scan.nodes.len();
                attach_node(&mut scan, node, &mut stack, index)?;
                stack.push(index);
            },
            Event::Empty(element) => {
                let mut node = make_node(
                    &element,
                    namespace,
                    &reader,
                    reader.decoder(),
                    before,
                    after,
                    stack.last().copied(),
                )?;
                node.span.empty = true;
                nodes_seen = nodes_seen.saturating_add(1);
                if nodes_seen > MAX_NODES {
                    return Err(limit("legacy comment collaboration XML nodes", MAX_NODES));
                }
                let index = scan.nodes.len();
                attach_node(&mut scan, node, &mut stack, index)?;
            },
            Event::End(element) => {
                let index = stack
                    .pop()
                    .ok_or_else(|| invalid("legacy comment collaboration has an unmatched end"))?;
                let node = scan
                    .nodes
                    .get_mut(index)
                    .ok_or_else(|| invalid("legacy comment collaboration node index is invalid"))?;
                if node.span.qname.as_slice() != element.name().as_ref() {
                    return Err(invalid(
                        "legacy comment collaboration start/end names differ",
                    ));
                }
                node.span.close_start = Some(before);
                node.span.range.end = after;
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "DTD is forbidden in legacy comment collaboration XML",
                ));
            },
            Event::PI(_) => {
                if scan.processing_instructions.len() >= MAX_NODES {
                    return Err(limit(
                        "legacy comment collaboration processing instructions",
                        MAX_NODES,
                    ));
                }
                scan.processing_instructions.push(before..after);
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                if !decoded.chars().all(char::is_whitespace) {
                    record_non_whitespace_text(&mut scan, stack.last().copied(), before..after)?;
                }
            },
            Event::GeneralRef(reference) => {
                let non_whitespace = reference
                    .resolve_char_ref()
                    .map_err(xml_error)?
                    .is_none_or(|character| !character.is_whitespace());
                if non_whitespace {
                    record_non_whitespace_text(&mut scan, stack.last().copied(), before..after)?;
                }
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::CData(_) => {
                return Err(invalid(
                    "CDATA is forbidden in legacy comment collaboration XML",
                ));
            },
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated legacy comment collaboration XML"));
    }
    Ok(scan)
}

fn record_non_whitespace_text(
    scan: &mut Scan,
    owner: Option<usize>,
    range: Range<usize>,
) -> Result<()> {
    let Some(owner) = owner else {
        return Err(invalid(
            "non-whitespace text is outside legacy comment collaboration XML",
        ));
    };
    if scan.non_whitespace_text.len() >= MAX_NODES {
        return Err(limit("legacy comment collaboration text nodes", MAX_NODES));
    }
    scan.non_whitespace_text.push((owner, range));
    Ok(())
}

struct MceSource<'a> {
    bytes: Cow<'a, [u8]>,
    processing_instructions: Vec<Range<usize>>,
}

impl<'a> MceSource<'a> {
    fn new(source: &'a [u8], processing_instructions: &[Range<usize>]) -> Result<Self> {
        if processing_instructions.is_empty() {
            return Ok(Self {
                bytes: Cow::Borrowed(source),
                processing_instructions: Vec::new(),
            });
        }
        let mut removed = 0usize;
        let mut cursor = 0usize;
        for range in processing_instructions {
            if range.start < cursor || range.end < range.start || range.end > source.len() {
                return Err(invalid("processing-instruction source range is invalid"));
            }
            removed = removed
                .checked_add(range.end - range.start)
                .ok_or_else(|| invalid("processing-instruction source size overflow"))?;
            cursor = range.end;
        }
        let output_len = source
            .len()
            .checked_sub(removed)
            .ok_or_else(|| invalid("processing-instruction source size underflow"))?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(output_len)
            .map_err(|source| Error::Allocation {
                resource: "legacy comment collaboration MCE source",
                source,
            })?;
        cursor = 0;
        for range in processing_instructions {
            bytes.extend_from_slice(&source[cursor..range.start]);
            cursor = range.end;
        }
        bytes.extend_from_slice(&source[cursor..]);
        if bytes.len() != output_len {
            return Err(invalid(
                "legacy comment collaboration MCE source length changed",
            ));
        }
        let mut ranges = Vec::new();
        ranges
            .try_reserve_exact(processing_instructions.len())
            .map_err(|source| Error::Allocation {
                resource: "legacy comment collaboration PI ranges",
                source,
            })?;
        ranges.extend_from_slice(processing_instructions);
        Ok(Self {
            bytes: Cow::Owned(bytes),
            processing_instructions: ranges,
        })
    }

    fn bytes(&self) -> &[u8] {
        self.bytes.as_ref()
    }

    fn translate_offsets(&self, offsets: &[u32]) -> Result<Vec<u32>> {
        if self.processing_instructions.is_empty() {
            let mut result = Vec::new();
            result
                .try_reserve_exact(offsets.len())
                .map_err(|source| Error::Allocation {
                    resource: "legacy comment collaboration MCE offsets",
                    source,
                })?;
            result.extend_from_slice(offsets);
            return Ok(result);
        }
        let mut translated = Vec::new();
        translated
            .try_reserve_exact(offsets.len())
            .map_err(|source| Error::Allocation {
                resource: "legacy comment collaboration MCE offsets",
                source,
            })?;
        let mut range_index = 0usize;
        let mut removed = 0usize;
        for &offset in offsets {
            let offset = usize::try_from(offset)
                .map_err(|_| invalid("legacy comment collaboration offset is too large"))?;
            while let Some(range) = self.processing_instructions.get(range_index) {
                if range.end > offset {
                    break;
                }
                removed = removed
                    .checked_add(range.end - range.start)
                    .ok_or_else(|| invalid("processing-instruction offset overflow"))?;
                range_index += 1;
            }
            if self
                .processing_instructions
                .get(range_index)
                .is_some_and(|range| range.start <= offset && offset < range.end)
            {
                return Err(invalid(
                    "legacy comment collaboration offset falls inside a processing instruction",
                ));
            }
            let translated_offset = offset
                .checked_sub(removed)
                .ok_or_else(|| invalid("legacy comment collaboration offset underflow"))?;
            translated.push(
                u32::try_from(translated_offset)
                    .map_err(|_| invalid("legacy comment collaboration offset is too large"))?,
            );
        }
        Ok(translated)
    }
}

fn attach_node(scan: &mut Scan, node: Node, stack: &mut [usize], index: usize) -> Result<()> {
    if let Some(parent) = stack.last().copied() {
        scan.nodes
            .get_mut(parent)
            .ok_or_else(|| invalid("legacy comment collaboration parent index is invalid"))?
            .children
            .push(index);
    } else if scan.root.replace(index).is_some() {
        return Err(invalid("legacy comment collaboration has multiple roots"));
    }
    scan.nodes.push(node);
    Ok(())
}

fn make_node(
    element: &BytesStart<'_>,
    namespace: ResolveResult<'_>,
    reader: &NsReader<&[u8]>,
    decoder: Decoder,
    before: usize,
    after: usize,
    parent: Option<usize>,
) -> Result<Node> {
    let namespace = match namespace {
        ResolveResult::Bound(Namespace(value)) => value.to_vec(),
        ResolveResult::Unbound => Vec::new(),
        ResolveResult::Unknown(prefix) => {
            return Err(invalid(format!(
                "unbound legacy comment collaboration namespace prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )));
        },
    };
    let raw = element.as_ref();
    let mut attrs = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        let key = String::from_utf8(attribute.key.as_ref().to_vec())
            .map_err(|_| invalid("legacy comment collaboration attribute is not UTF-8"))?;
        if key != "xmlns"
            && !key.starts_with("xmlns:")
            && matches!(
                reader.resolver().resolve_attribute(attribute.key).0,
                ResolveResult::Unknown(_)
            )
        {
            return Err(invalid(
                "legacy comment collaboration attribute uses an unbound namespace prefix",
            ));
        }
        if !seen.insert(key.clone()) {
            return Err(invalid(
                "duplicate XML attribute in legacy comment collaboration",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        if key != "xmlns" && !key.starts_with("xmlns:") {
            let value_range =
                attribute_value_span(raw, attribute.key.as_ref())?.ok_or_else(|| {
                    invalid("legacy comment collaboration attribute value span is missing")
                })?;
            let full_range = attribute_full_span(raw, attribute.key.as_ref())?
                .ok_or_else(|| invalid("legacy comment collaboration attribute span is missing"))?;
            attrs.push(Attr {
                key,
                value,
                value_range: (before + 1 + value_range.start)..(before + 1 + value_range.end),
                full_range: (before + 1 + full_range.start)..(before + 1 + full_range.end),
            });
        }
    }
    Ok(Node {
        span: Span {
            range: before..after,
            start_end: after,
            close_start: None,
            qname: element.name().as_ref().to_vec(),
            empty: false,
        },
        namespace,
        local: element.local_name().as_ref().to_vec(),
        attrs,
        parent,
        children: Vec::new(),
    })
}

fn item_indices(scan: &Scan, root: usize, kind: Kind, owner: Owner) -> Result<Vec<usize>> {
    let expected = match kind {
        Kind::Presence => b"cmAuthor".as_slice(),
        Kind::Threading => b"cm".as_slice(),
    };
    let mut result = Vec::new();
    for (index, node) in scan.nodes.iter().enumerate() {
        if !pml_namespace(&node.namespace) || node.local.as_slice() != expected {
            continue;
        }
        if nearest_non_mce_parent(scan, index) != Some(root) {
            continue;
        }
        let matches = match owner {
            Owner::Author(id) => attr_u32(node, b"id")? == Some(id),
            Owner::Comment {
                author_id,
                index: key,
            } => {
                attr_u32(node, b"authorId")? == Some(author_id)
                    && attr_u32(node, b"idx")? == Some(key)
            },
        };
        if matches {
            result.push(index);
        }
    }
    if result.is_empty() {
        return Err(invalid(
            "requested legacy comment collaboration owner was not found",
        ));
    }
    Ok(result)
}

fn collect_offsets(scan: &Scan, items: &[usize]) -> Vec<u32> {
    let mut offsets = Vec::new();
    offsets.extend(
        items
            .iter()
            .map(|item| scan.nodes[*item].span.range.start as u32),
    );
    let item_set = items.iter().copied().collect::<HashSet<_>>();
    for (index, node) in scan.nodes.iter().enumerate() {
        let mut current = Some(index);
        let mut selected = false;
        while let Some(value) = current {
            if item_set.contains(&value) {
                selected = true;
                break;
            }
            current = scan.nodes[value].parent;
        }
        if !selected {
            continue;
        }
        // Ask the MCE selector about every node in the selected owner, not
        // only recognized schema nodes. Otherwise an active foreign child
        // would be absent from `active` and could be mistaken for an
        // inactive MCE branch by the closed typed grammar.
        offsets.push(node.span.range.start as u32);
    }
    offsets.sort_unstable();
    offsets.dedup();
    offsets
}

fn target_extension_owner(scan: &Scan, index: usize, kind: Kind) -> Option<usize> {
    let node = &scan.nodes[index];
    if !pml_namespace(&node.namespace) || node.local.as_slice() != b"ext" {
        return None;
    }
    let parent = nearest_non_mce_parent(scan, index)?;
    let parent_node = &scan.nodes[parent];
    if !(pml_namespace(&parent_node.namespace) && parent_node.local.as_slice() == b"extLst") {
        return None;
    }
    let item = nearest_non_mce_parent(scan, parent)?;
    let item_node = &scan.nodes[item];
    let expected_item = match kind {
        Kind::Presence => b"cmAuthor".as_slice(),
        Kind::Threading => b"cm".as_slice(),
    };
    if !pml_namespace(&item_node.namespace) || item_node.local.as_slice() != expected_item {
        return None;
    }
    if attr_string(node, b"uri").is_none_or(|value| value != extension_uri(kind)) {
        return None;
    }
    Some(item)
}

fn payload_name(node: &Node, kind: Kind) -> bool {
    p15_namespace(&node.namespace)
        && node.local.as_slice()
            == match kind {
                Kind::Presence => b"presenceInfo".as_slice(),
                Kind::Threading => b"threadingInfo".as_slice(),
            }
}

fn nearest_non_mce_parent(scan: &Scan, index: usize) -> Option<usize> {
    let mut current = scan.nodes[index].parent;
    while let Some(value) = current {
        if !mce_namespace(&scan.nodes[value].namespace) {
            return Some(value);
        }
        current = scan.nodes[value].parent;
    }
    None
}

fn direct_non_mce_children<'a>(scan: &'a Scan, parent: usize) -> impl Iterator<Item = usize> + 'a {
    scan.nodes
        .iter()
        .enumerate()
        .filter(move |(index, _)| nearest_non_mce_parent(scan, *index) == Some(parent))
        .map(|(index, _)| index)
}

fn is_active(span: &Span, active: &HashSet<u32>) -> bool {
    active.contains(&(span.range.start as u32))
}

fn has_active_non_whitespace_text(scan: &Scan, parent: usize, active: &HashSet<u32>) -> bool {
    is_active(&scan.nodes[parent].span, active)
        && scan
            .non_whitespace_text
            .iter()
            .any(|(owner, _range)| *owner == parent)
}

fn parse_presence(
    node: &Node,
    payload_index: usize,
    scan: &Scan,
    active: &HashSet<u32>,
) -> Result<PresenceInfo> {
    ensure_only_attributes(node, &[b"userId", b"providerId"], "presenceInfo")?;
    if has_active_non_whitespace_text(scan, payload_index, active) {
        return Err(invalid("presenceInfo must contain only whitespace text"));
    }
    let user_id = attr_required(node, b"userId", "presenceInfo")?;
    let provider_id = attr_required(node, b"providerId", "presenceInfo")?;
    if direct_non_mce_children(scan, payload_index)
        .any(|index| is_active(&scan.nodes[index].span, active))
    {
        return Err(invalid("presenceInfo must not contain child elements"));
    }
    validate_string(&user_id, "presenceInfo userId")?;
    validate_string(&provider_id, "presenceInfo providerId")?;
    Ok(PresenceInfo::new(user_id, provider_id))
}

fn parse_threading(
    node: &Node,
    payload_index: usize,
    parent_index: Option<usize>,
    scan: &Scan,
    active: &HashSet<u32>,
) -> Result<ThreadingInfo> {
    ensure_only_attributes(node, &[b"timeZoneBias"], "threadingInfo")?;
    if has_active_non_whitespace_text(scan, payload_index, active) {
        return Err(invalid("threadingInfo must contain only whitespace text"));
    }
    let time_zone_bias = attr_i32(node, b"timeZoneBias")?;
    let mut parent = None;
    let mut element_count = 0usize;
    for child in direct_non_mce_children(scan, payload_index) {
        if !is_active(&scan.nodes[child].span, active) {
            continue;
        }
        let child_node = &scan.nodes[child];
        if !p15_namespace(&child_node.namespace) || child_node.local.as_slice() != b"parentCm" {
            continue;
        }
        element_count += 1;
        if Some(child) == parent_index {
            ensure_only_attributes(child_node, &[b"authorId", b"idx"], "parentCm")?;
            if has_active_non_whitespace_text(scan, child, active) {
                return Err(invalid("parentCm must contain only whitespace text"));
            }
            if direct_non_mce_children(scan, child)
                .any(|index| is_active(&scan.nodes[index].span, active))
            {
                return Err(invalid("parentCm must not contain child elements"));
            }
            parent = Some(ParentComment::new(
                attr_u32(child_node, b"authorId")?,
                attr_u32(child_node, b"idx")?,
            ));
        }
    }
    if element_count > 1 {
        return Err(invalid(
            "threadingInfo contains duplicate parentCm elements",
        ));
    }
    Ok(ThreadingInfo::new(time_zone_bias, parent))
}

fn ensure_only_attributes(node: &Node, allowed: &[&[u8]], name: &str) -> Result<()> {
    if let Some(attribute) = node
        .attrs
        .iter()
        .find(|attribute| !allowed.contains(&attribute.key.as_bytes()))
    {
        return Err(invalid(format!(
            "unexpected attribute '{}' on {name}",
            attribute.key
        )));
    }
    Ok(())
}

fn attr_required(node: &Node, key: &[u8], element: &str) -> Result<String> {
    attr_string(node, key)
        .ok_or_else(|| invalid(format!("{element} requires a required attribute")))
}

fn attr_string(node: &Node, key: &[u8]) -> Option<String> {
    node.attrs
        .iter()
        .find(|attr| attr.key.as_bytes() == key)
        .map(|attr| attr.value.clone())
}

fn attr_u32(node: &Node, key: &[u8]) -> Result<Option<u32>> {
    let Some(value) = attr_string(node, key) else {
        return Ok(None);
    };
    value
        .trim()
        .parse()
        .map(Some)
        .map_err(|_| invalid("legacy comment collaboration unsigned integer is invalid"))
}

fn attr_i32(node: &Node, key: &[u8]) -> Result<Option<i32>> {
    let Some(value) = attr_string(node, key) else {
        return Ok(None);
    };
    value
        .trim()
        .parse()
        .map(Some)
        .map_err(|_| invalid("threadingInfo timeZoneBias is not an xsd:int"))
}

pub(crate) fn validate_value(kind: Kind, value: &Value) -> Result<()> {
    match (kind, value) {
        (Kind::Presence, Value::Presence(value)) => {
            validate_string(&value.user_id, "presenceInfo userId")?;
            validate_string(&value.provider_id, "presenceInfo providerId")
        },
        (Kind::Threading, Value::Threading(_)) => Ok(()),
        _ => Err(invalid(
            "typed collaboration value kind does not match owner",
        )),
    }
}

fn validate_string(value: &str, label: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES {
        return Err(invalid(format!("{label} exceeds XML string limits")));
    }
    if let Some(character) = value.chars().find(|&character| !is_xml_10_char(character)) {
        return Err(invalid(format!(
            "{label} contains XML 1.0-forbidden character U+{:04X}",
            u32::from(character)
        )));
    }
    Ok(())
}

fn is_xml_10_char(character: char) -> bool {
    matches!(character, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(
            u32::from(character),
            0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
        )
}

fn replace_required_attribute(
    replacements: &mut Vec<Replacement>,
    node: &Node,
    key: &[u8],
    value: &str,
) -> Result<()> {
    let attr = node
        .attrs
        .iter()
        .find(|attr| attr.key.as_bytes() == key)
        .ok_or_else(|| invalid("required typed collaboration attribute source is missing"))?;
    replacements.push(Replacement {
        range: attr.value_range.clone(),
        value: escaped(value)?,
    });
    Ok(())
}

fn replace_optional_i32_attribute(
    source: &[u8],
    replacements: &mut Vec<Replacement>,
    node: &Node,
    key: &[u8],
    before: Option<i32>,
    after: Option<i32>,
) -> Result<()> {
    replace_optional_attribute(
        source,
        replacements,
        node,
        key,
        before.map(|value| value.to_string()),
        after.map(|value| value.to_string()),
    )
}

fn replace_optional_u32_attribute(
    source: &[u8],
    replacements: &mut Vec<Replacement>,
    node: &Node,
    key: &[u8],
    before: Option<u32>,
    after: Option<u32>,
) -> Result<()> {
    replace_optional_attribute(
        source,
        replacements,
        node,
        key,
        before.map(|value| value.to_string()),
        after.map(|value| value.to_string()),
    )
}

#[derive(Debug)]
struct SizePlan {
    range: Range<usize>,
    length: usize,
}

fn optional_attribute_plan(
    source: &[u8],
    node: &Node,
    key: &[u8],
    before: Option<&str>,
    after: Option<&str>,
) -> Result<Option<SizePlan>> {
    if before == after {
        return Ok(None);
    }
    let existing = node.attrs.iter().find(|attr| attr.key.as_bytes() == key);
    match (existing, after) {
        (Some(attr), Some(value)) => Ok(Some(SizePlan {
            range: attr.value_range.clone(),
            length: escaped_len(value)?,
        })),
        (Some(attr), None) => Ok(Some(SizePlan {
            range: attr.full_range.clone(),
            length: 0,
        })),
        (None, Some(value)) => {
            let at = opening_insert_position(source, &node.span)?;
            let value_len = escaped_len(value)?;
            let fragment_len = checked_sum(&[1, key.len(), 2, value_len, 1])?;
            Ok(Some(SizePlan {
                range: at..at,
                length: fragment_len,
            }))
        },
        (None, None) => Ok(None),
    }
}

fn replace_optional_attribute(
    source: &[u8],
    replacements: &mut Vec<Replacement>,
    node: &Node,
    key: &[u8],
    before: Option<String>,
    after: Option<String>,
) -> Result<()> {
    if before == after {
        return Ok(());
    }
    let existing = node.attrs.iter().find(|attr| attr.key.as_bytes() == key);
    match (existing, after) {
        (Some(attr), Some(value)) => {
            let length = escaped_len(&value)?;
            preflight_replacement_sizes(source, &[(&attr.value_range, length)])?;
            replacements.push(Replacement {
                range: attr.value_range.clone(),
                value: escaped(&value)?,
            });
        },
        (Some(attr), None) => replacements.push(Replacement {
            range: attr.full_range.clone(),
            value: Vec::new(),
        }),
        (None, Some(value)) => {
            let at = opening_insert_position(source, &node.span)?;
            let escaped_len = escaped_len(&value)?;
            let fragment_len = checked_sum(&[1, key.len(), 2, escaped_len, 1])?;
            let insertion = at..at;
            preflight_replacement_sizes(source, &[(&insertion, fragment_len)])?;
            let escaped = escaped(&value)?;
            let mut fragment =
                new_xml_buffer(fragment_len, "legacy comment collaboration attribute")?;
            fragment.push(b' ');
            fragment.extend_from_slice(key);
            fragment.extend_from_slice(b"=\"");
            fragment.extend_from_slice(&escaped);
            fragment.push(b'\"');
            replacements.push(Replacement {
                range: insertion,
                value: fragment,
            });
        },
        (None, None) => {},
    }
    Ok(())
}

fn expand_self_closing_payload(
    source: &[u8],
    payload: &Span,
    node: &Node,
    after_time_zone_bias: Option<i32>,
    bias_plan: Option<&SizePlan>,
    parent: &ParentComment,
    output_len: usize,
) -> Result<Vec<u8>> {
    let raw = source
        .get(payload.range.start..payload.start_end)
        .ok_or_else(|| invalid("self-closing threadingInfo source is missing"))?;
    let opening_len = empty_opening_length_from_raw(raw)?;
    let opening_end = opening_len
        .checked_add(1)
        .ok_or_else(|| invalid("self-closing threadingInfo opening size overflow"))?;
    let base_end = opening_end
        .checked_sub(2)
        .ok_or_else(|| invalid("self-closing threadingInfo opening size underflow"))?;
    let mut output = new_xml_buffer(output_len, "expanded threadingInfo XML")?;
    if let Some(plan) = bias_plan {
        let start = plan
            .range
            .start
            .checked_sub(payload.range.start)
            .ok_or_else(|| invalid("self-closing threadingInfo edit underflow"))?;
        let end = plan
            .range
            .end
            .checked_sub(payload.range.start)
            .ok_or_else(|| invalid("self-closing threadingInfo edit underflow"))?;
        if end > base_end || start > end {
            return Err(invalid(
                "self-closing threadingInfo edit escapes its opening tag",
            ));
        }
        output.extend_from_slice(&raw[..start]);
        let attribute = node
            .attrs
            .iter()
            .find(|attribute| attribute.key.as_bytes() == b"timeZoneBias");
        match (attribute, after_time_zone_bias) {
            (Some(_), Some(value)) => {
                append_signed_decimal(&mut output, value);
            },
            (Some(_), None) => {},
            (None, Some(value)) => {
                output.extend_from_slice(b" timeZoneBias=\"");
                append_signed_decimal(&mut output, value);
                output.push(b'\"');
            },
            (None, None) => {
                return Err(invalid(
                    "self-closing threadingInfo timeZoneBias edit has no target",
                ));
            },
        }
        output.extend_from_slice(&raw[end..base_end]);
    } else {
        output.extend_from_slice(&raw[..base_end]);
    }
    output.push(b'>');
    append_parent_fragment(&mut output, qname_prefix(&payload.qname), parent);
    output.extend_from_slice(b"</");
    output.extend_from_slice(&payload.qname);
    output.push(b'>');
    if output.len() != output_len {
        return Err(invalid(
            "expanded threadingInfo XML length changed during construction",
        ));
    }
    Ok(output)
}

fn insert_child(
    source: &[u8],
    parent: &Span,
    fragment: Vec<u8>,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    if parent.empty {
        let opening = empty_opening(source, parent)?;
        let close = closing(&parent.qname)?;
        let replacement = concat_fragments(&[&opening, &fragment, &close])?;
        replacements.push(Replacement {
            range: parent.range.clone(),
            value: replacement,
        });
    } else {
        let at = parent
            .close_start
            .ok_or_else(|| invalid("parent closing source is missing"))?;
        replacements.push(Replacement {
            range: at..at,
            value: fragment,
        });
    }
    Ok(())
}

fn parent_fragment(prefix: &[u8], value: &ParentComment) -> Result<Vec<u8>> {
    let length = parent_length(prefix, value)?;
    let mut out = new_xml_buffer(length, "parentCm XML")?;
    append_parent_fragment(&mut out, prefix, value);
    Ok(out)
}

fn parent_length(prefix: &[u8], value: &ParentComment) -> Result<usize> {
    checked_sum(&[
        1,
        qualified_length(prefix, b"parentCm")?,
        value.author_id.map_or(0, |value| {
            b" authorId=\"".len() + unsigned_decimal_len(value) + 1
        }),
        value.index.map_or(0, |value| {
            b" idx=\"".len() + unsigned_decimal_len(value) + 1
        }),
        2,
    ])
}

fn append_parent_fragment(out: &mut Vec<u8>, prefix: &[u8], value: &ParentComment) {
    out.push(b'<');
    append_qualified(out, prefix, b"parentCm");
    if let Some(author_id) = value.author_id {
        out.extend_from_slice(b" authorId=\"");
        append_unsigned_decimal(out, author_id);
        out.push(b'\"');
    }
    if let Some(index) = value.index {
        out.extend_from_slice(b" idx=\"");
        append_unsigned_decimal(out, index);
        out.push(b'\"');
    }
    out.extend_from_slice(b"/>");
}

fn payload_length(kind: Kind, value: &Value) -> Result<usize> {
    match (kind, value) {
        (Kind::Presence, Value::Presence(value)) => checked_sum(&[
            b"<p15:presenceInfo xmlns:p15=\"".len(),
            NAMESPACE.len(),
            b"\" userId=\"".len(),
            escaped_len(&value.user_id)?,
            b"\" providerId=\"".len(),
            escaped_len(&value.provider_id)?,
            b"\"/>".len(),
        ]),
        (Kind::Threading, Value::Threading(value)) => {
            let parent = value.parent.as_ref();
            let parent_size = parent
                .map(|parent| parent_length(b"p15", parent))
                .transpose()?
                .unwrap_or(0);
            checked_sum(&[
                b"<p15:threadingInfo xmlns:p15=\"".len(),
                NAMESPACE.len(),
                1,
                value.time_zone_bias.map_or(0, |value| {
                    b" timeZoneBias=\"".len() + signed_decimal_len(value) + 1
                }),
                if parent.is_some() { 1 } else { b"/>".len() },
                parent_size,
                if parent.is_some() {
                    b"</p15:threadingInfo>".len()
                } else {
                    0
                },
            ])
        },
        _ => Err(invalid(
            "typed collaboration payload kind does not match owner",
        )),
    }
}

fn append_payload(out: &mut Vec<u8>, kind: Kind, value: &Value) -> Result<()> {
    match (kind, value) {
        (Kind::Presence, Value::Presence(value)) => {
            out.extend_from_slice(b"<p15:presenceInfo xmlns:p15=\"");
            out.extend_from_slice(NAMESPACE.as_bytes());
            out.extend_from_slice(b"\" userId=\"");
            append_escaped(out, &value.user_id);
            out.extend_from_slice(b"\" providerId=\"");
            append_escaped(out, &value.provider_id);
            out.extend_from_slice(b"\"/>");
        },
        (Kind::Threading, Value::Threading(value)) => {
            out.extend_from_slice(b"<p15:threadingInfo xmlns:p15=\"");
            out.extend_from_slice(NAMESPACE.as_bytes());
            out.push(b'\"');
            if let Some(bias) = value.time_zone_bias {
                out.extend_from_slice(b" timeZoneBias=\"");
                append_signed_decimal(out, bias);
                out.push(b'\"');
            }
            if let Some(parent) = &value.parent {
                out.push(b'>');
                append_parent_fragment(out, b"p15", parent);
                out.extend_from_slice(b"</p15:threadingInfo>");
            } else {
                out.extend_from_slice(b"/>");
            }
        },
        _ => {
            return Err(invalid(
                "typed collaboration payload kind does not match owner",
            ));
        },
    }
    Ok(())
}

fn extension_fragment_for_value(kind: Kind, qname: &[u8], value: &Value) -> Result<Vec<u8>> {
    let payload_size = payload_length(kind, value)?;
    let length = extension_fragment_length(kind, qname.len(), payload_size)?;
    let mut output = new_xml_buffer(length, "legacy comment collaboration extension XML")?;
    output.push(b'<');
    output.extend_from_slice(qname);
    output.extend_from_slice(b" uri=\"");
    output.extend_from_slice(extension_uri(kind).as_bytes());
    output.extend_from_slice(b"\">");
    append_payload(&mut output, kind, value)?;
    output.extend_from_slice(b"</");
    output.extend_from_slice(qname);
    output.push(b'>');
    debug_assert_eq!(output.len(), length);
    Ok(output)
}

fn extension_fragment_length(kind: Kind, qname_len: usize, payload_size: usize) -> Result<usize> {
    checked_sum(&[
        1,
        qname_len,
        b" uri=\"".len(),
        extension_uri(kind).len(),
        b"\">".len(),
        payload_size,
        b"</".len(),
        qname_len,
        1,
    ])
}

fn extension_uri(kind: Kind) -> &'static str {
    match kind {
        Kind::Presence => PRESENCE_EXTENSION_URI,
        Kind::Threading => THREADING_EXTENSION_URI,
    }
}

fn pml_namespace(namespace: &[u8]) -> bool {
    namespace == PML.as_bytes() || namespace == STRICT_PML.as_bytes()
}

fn p15_namespace(namespace: &[u8]) -> bool {
    namespace == NAMESPACE.as_bytes()
}

fn mce_namespace(namespace: &[u8]) -> bool {
    namespace == MCE_NAMESPACE.as_bytes()
}

fn capabilities() -> Capabilities {
    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities.understand_namespace(NAMESPACE);
    capabilities
}

fn mce_limits() -> MceLimits {
    MceLimits {
        max_input_bytes: MAX_PART_BYTES,
        max_output_bytes: MAX_PART_BYTES,
        max_depth: MAX_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    }
}

fn opening_insert_position(source: &[u8], span: &Span) -> Result<usize> {
    let end = span.start_end;
    let raw = source
        .get(span.range.start..end)
        .ok_or_else(|| invalid("legacy comment collaboration opening escapes source"))?;
    let mut index = raw.len();
    while index > 0 && raw[index - 1].is_ascii_whitespace() {
        index -= 1;
    }
    if index == 0 || raw[index - 1] != b'>' {
        return Err(invalid("legacy comment collaboration opening is malformed"));
    }
    if index >= 2 && raw[index - 2] == b'/' {
        Ok(end - 2)
    } else {
        Ok(end - 1)
    }
}

fn empty_opening(source: &[u8], span: &Span) -> Result<Vec<u8>> {
    let raw = source
        .get(span.range.clone())
        .ok_or_else(|| invalid("legacy comment collaboration empty span escapes source"))?;
    let opening_len = empty_opening_length_from_raw(raw)?;
    let mut value = new_xml_buffer(opening_len, "legacy comment collaboration opening")?;
    let end = raw
        .iter()
        .rposition(|byte| !byte.is_ascii_whitespace())
        .map_or(0, |index| index + 1);
    value.extend_from_slice(&raw[..end - 2]);
    value.push(b'>');
    Ok(value)
}

fn empty_opening_length(source: &[u8], span: &Span) -> Result<usize> {
    let raw = source
        .get(span.range.clone())
        .ok_or_else(|| invalid("legacy comment collaboration empty span escapes source"))?;
    empty_opening_length_from_raw(raw)
}

fn empty_opening_length_from_raw(raw: &[u8]) -> Result<usize> {
    let mut end = raw.len();
    while end > 0 && raw[end - 1].is_ascii_whitespace() {
        end -= 1;
    }
    if end < 2 || raw[end - 1] != b'>' || raw[end - 2] != b'/' {
        return Err(invalid(
            "legacy comment collaboration self-closing tag is malformed",
        ));
    }
    end.checked_sub(1)
        .ok_or_else(|| invalid("legacy comment collaboration opening size underflow"))
}

fn self_closing_opening(source: &[u8], span: &Span) -> Result<Vec<u8>> {
    let raw = source
        .get(span.range.start..span.start_end)
        .ok_or_else(|| invalid("legacy comment collaboration opening escapes source"))?;
    let mut end = raw.len();
    while end > 0 && raw[end - 1].is_ascii_whitespace() {
        end -= 1;
    }
    if end == 0 || raw[end - 1] != b'>' {
        return Err(invalid("legacy comment collaboration opening is malformed"));
    }
    let mut body_end = end - 1;
    while body_end > 0 && raw[body_end - 1].is_ascii_whitespace() {
        body_end -= 1;
    }
    let slash = body_end > 0 && raw[body_end - 1] == b'/';
    let length = body_end
        .checked_add(if slash { 1 } else { 2 })
        .ok_or_else(|| invalid("legacy comment collaboration opening size overflow"))?;
    let mut result = new_xml_buffer(length, "legacy comment collaboration opening")?;
    result.extend_from_slice(&raw[..body_end]);
    if !slash {
        result.push(b'/');
    }
    result.push(b'>');
    Ok(result)
}

fn only_whitespace_after_ranges(source: &[u8], container: &Span, ranges: &[Span]) -> bool {
    if container.empty {
        return false;
    }
    let body_start = container.start_end;
    let body_end = container.close_start.unwrap_or(container.range.end);
    if body_start > body_end || body_end > source.len() {
        return false;
    }
    let mut cursor = body_start;
    let mut ordered = ranges
        .iter()
        .filter_map(|span| {
            let start = span.range.start.max(body_start);
            let end = span.range.end.min(body_end);
            (start <= end).then_some(start..end)
        })
        .collect::<Vec<_>>();
    ordered.sort_by_key(|range| range.start);
    for range in ordered {
        if range.start > cursor
            && !source[cursor..range.start]
                .iter()
                .all(u8::is_ascii_whitespace)
        {
            return false;
        }
        cursor = cursor.max(range.end);
    }
    source[cursor..body_end].iter().all(u8::is_ascii_whitespace)
}

fn closing(qname: &[u8]) -> Result<Vec<u8>> {
    let length = closing_length(qname)?;
    let mut result = new_xml_buffer(length, "legacy comment collaboration closing")?;
    result.extend_from_slice(b"</");
    result.extend_from_slice(qname);
    result.push(b'>');
    Ok(result)
}

fn closing_length(qname: &[u8]) -> Result<usize> {
    checked_sum(&[2, qname.len(), 1])
}

fn qname_prefix(qname: &[u8]) -> &[u8] {
    qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |index| &qname[..index])
}

fn qualified(prefix: &[u8], local: &[u8]) -> Result<Vec<u8>> {
    let length = qualified_length(prefix, local)?;
    let mut result = new_xml_buffer(length, "legacy comment collaboration QName")?;
    append_qualified(&mut result, prefix, local);
    Ok(result)
}

fn qualified_length(prefix: &[u8], local: &[u8]) -> Result<usize> {
    checked_sum(&[
        prefix.len(),
        if prefix.is_empty() { 0 } else { 1 },
        local.len(),
    ])
}

fn append_qualified(output: &mut Vec<u8>, prefix: &[u8], local: &[u8]) {
    if !prefix.is_empty() {
        output.extend_from_slice(prefix);
        output.push(b':');
    }
    output.extend_from_slice(local);
}

fn unsigned_decimal_len(value: u32) -> usize {
    let mut value = u64::from(value);
    let mut length = 1;
    while value >= 10 {
        value /= 10;
        length += 1;
    }
    length
}

fn signed_decimal_len(value: i32) -> usize {
    if value < 0 {
        1 + unsigned_decimal_len(value.unsigned_abs())
    } else {
        unsigned_decimal_len(value as u32)
    }
}

fn append_unsigned_decimal(output: &mut Vec<u8>, value: u32) {
    let mut value = u64::from(value);
    let mut digits = [0u8; 10];
    let mut index = digits.len();
    while value >= 10 {
        index -= 1;
        digits[index] = b'0' + (value % 10) as u8;
        value /= 10;
    }
    index -= 1;
    digits[index] = b'0' + value as u8;
    output.extend_from_slice(&digits[index..]);
}

fn append_signed_decimal(output: &mut Vec<u8>, value: i32) {
    if value < 0 {
        output.push(b'-');
        append_unsigned_decimal(output, value.unsigned_abs());
    } else {
        append_unsigned_decimal(output, value as u32);
    }
}

fn escaped(value: &str) -> Result<Vec<u8>> {
    let length = escaped_len(value)?;
    let mut output = new_xml_buffer(length, "legacy comment collaboration escaped XML")?;
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
    Ok(output)
}

fn escaped_len(value: &str) -> Result<usize> {
    let mut length = 0usize;
    for character in value.chars() {
        let addition = match character {
            '&' => b"&amp;".len(),
            '<' => b"&lt;".len(),
            '>' => b"&gt;".len(),
            '"' => b"&quot;".len(),
            '\t' => b"&#x9;".len(),
            '\n' => b"&#xA;".len(),
            '\r' => b"&#xD;".len(),
            _ => character.len_utf8(),
        };
        length = length
            .checked_add(addition)
            .ok_or_else(|| invalid("legacy comment collaboration escaped XML size overflow"))?;
    }
    ensure_xml_size(length)?;
    Ok(length)
}

fn append_escaped(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

fn checked_sum(parts: &[usize]) -> Result<usize> {
    parts.iter().try_fold(0usize, |total, part| {
        total
            .checked_add(*part)
            .ok_or_else(|| invalid("legacy comment collaboration XML size overflow"))
    })
}

fn preflight_replacement_sizes(
    source: &[u8],
    replacements: &[(&Range<usize>, usize)],
) -> Result<()> {
    let mut removed = 0usize;
    let mut added = 0usize;
    for (range, length) in replacements {
        if range.start > range.end || range.end > source.len() {
            return Err(invalid(
                "legacy comment collaboration patch range escapes source",
            ));
        }
        removed = removed
            .checked_add(range.len())
            .ok_or_else(|| invalid("legacy comment collaboration removed size overflow"))?;
        added = added
            .checked_add(*length)
            .ok_or_else(|| invalid("legacy comment collaboration replacement size overflow"))?;
    }
    let output_len = source
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(added))
        .ok_or_else(|| invalid("serialized legacy comment collaboration size overflow"))?;
    ensure_xml_size(output_len)
}

fn preflight_size_plans(
    source: &[u8],
    plans: impl IntoIterator<Item = Option<SizePlan>>,
) -> Result<()> {
    let mut removed = 0usize;
    let mut added = 0usize;
    for plan in plans.into_iter().flatten() {
        if plan.range.start > plan.range.end || plan.range.end > source.len() {
            return Err(invalid(
                "legacy comment collaboration patch range escapes source",
            ));
        }
        removed = removed
            .checked_add(plan.range.len())
            .ok_or_else(|| invalid("legacy comment collaboration removed size overflow"))?;
        added = added
            .checked_add(plan.length)
            .ok_or_else(|| invalid("legacy comment collaboration replacement size overflow"))?;
    }
    let output_len = source
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(added))
        .ok_or_else(|| invalid("serialized legacy comment collaboration size overflow"))?;
    ensure_xml_size(output_len)
}

fn ensure_xml_size(length: usize) -> Result<()> {
    if length > MAX_PART_BYTES {
        return Err(limit(
            "serialized legacy comment collaboration XML bytes",
            MAX_PART_BYTES,
        ));
    }
    Ok(())
}

fn new_xml_buffer(length: usize, resource: &'static str) -> Result<Vec<u8>> {
    ensure_xml_size(length)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation { resource, source })?;
    Ok(output)
}

fn concat_fragments(fragments: &[&[u8]]) -> Result<Vec<u8>> {
    let mut length = 0usize;
    for fragment in fragments {
        length = length
            .checked_add(fragment.len())
            .ok_or_else(|| invalid("legacy comment collaboration XML size overflow"))?;
    }
    let mut output = new_xml_buffer(length, "legacy comment collaboration XML fragment")?;
    for fragment in fragments {
        output.extend_from_slice(fragment);
    }
    Ok(output)
}

fn wrapped_fragment(qname: &[u8], body: &[u8]) -> Result<Vec<u8>> {
    concat_fragments(&[b"<", qname, b">", body, b"</", qname, b">"])
}

fn wrapped_fragment_length(qname_len: usize, body_len: usize) -> Result<usize> {
    checked_sum(&[1, qname_len, 1, body_len, 2, qname_len, 1])
}

#[derive(Debug)]
struct Replacement {
    range: Range<usize>,
    value: Vec<u8>,
}

fn copy_source(source: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "legacy comment collaboration source XML",
            source,
        })?;
    output.extend_from_slice(source);
    Ok(output)
}

fn apply_replacements(source: &[u8], mut replacements: Vec<Replacement>) -> Result<Vec<u8>> {
    replacements.sort_by_key(|replacement| replacement.range.start);
    let mut removed = 0usize;
    let mut added = 0usize;
    let mut cursor = 0usize;
    for replacement in &replacements {
        if replacement.range.start > replacement.range.end
            || replacement.range.end > source.len()
            || replacement.range.start < cursor
        {
            return Err(invalid("legacy comment collaboration patch ranges overlap"));
        }
        removed = removed
            .checked_add(replacement.range.len())
            .ok_or_else(|| invalid("legacy comment collaboration removed size overflow"))?;
        added = added
            .checked_add(replacement.value.len())
            .ok_or_else(|| invalid("legacy comment collaboration replacement size overflow"))?;
        cursor = replacement.range.end;
    }
    let output_len = source
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(added))
        .ok_or_else(|| invalid("serialized legacy comment collaboration size overflow"))?;
    if output_len > MAX_PART_BYTES {
        return Err(limit(
            "serialized legacy comment collaboration XML bytes",
            MAX_PART_BYTES,
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "serialized legacy comment collaboration XML",
            source,
        })?;
    cursor = 0;
    for replacement in replacements {
        output.extend_from_slice(&source[cursor..replacement.range.start]);
        output.extend_from_slice(&replacement.value);
        cursor = replacement.range.end;
    }
    output.extend_from_slice(&source[cursor..]);
    if output.len() != output_len {
        return Err(invalid(
            "serialized legacy comment collaboration length changed",
        ));
    }
    Ok(output)
}

fn attribute_value_span(raw: &[u8], key: &[u8]) -> Result<Option<Range<usize>>> {
    Ok(attribute_ranges(raw, key)?.map(|(_, value)| value))
}

fn attribute_full_span(raw: &[u8], key: &[u8]) -> Result<Option<Range<usize>>> {
    Ok(attribute_ranges(raw, key)?.map(|(full, _)| full))
}

fn attribute_ranges(raw: &[u8], key: &[u8]) -> Result<Option<(Range<usize>, Range<usize>)>> {
    let mut index = 0usize;
    while index < raw.len()
        && raw[index] != b' '
        && raw[index] != b'\t'
        && raw[index] != b'\n'
        && raw[index] != b'\r'
        && raw[index] != b'>'
        && raw[index] != b'/'
    {
        index += 1;
    }
    while index < raw.len() {
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= raw.len() || raw[index] == b'>' || raw[index] == b'/' {
            break;
        }
        let full_start = index;
        while index < raw.len()
            && !raw[index].is_ascii_whitespace()
            && !matches!(raw[index], b'=' | b'>' | b'/')
        {
            index += 1;
        }
        let name = &raw[full_start..index];
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= raw.len() || raw[index] != b'=' {
            return Err(invalid(
                "legacy comment collaboration attribute has no value",
            ));
        }
        index += 1;
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *raw
            .get(index)
            .ok_or_else(|| invalid("legacy comment collaboration attribute value is missing"))?;
        if quote != b'\'' && quote != b'"' {
            return Err(invalid(
                "legacy comment collaboration attribute value is not quoted",
            ));
        }
        index += 1;
        let value_start = index;
        while index < raw.len() && raw[index] != quote {
            index += 1;
        }
        if index >= raw.len() {
            return Err(invalid(
                "legacy comment collaboration attribute value is unterminated",
            ));
        }
        let value_end = index;
        index += 1;
        if name == key {
            return Ok(Some((full_start..index, value_start..value_end)));
        }
    }
    Ok(None)
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("legacy comment collaboration XML offset does not fit usize"))
}

fn limit(resource: &'static str, maximum: usize) -> Error {
    Error::Limit {
        resource,
        limit: maximum,
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
