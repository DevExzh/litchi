//! Byte-preserving rewrites for the two workbook calculation-metadata owners.

use std::borrow::Cow;

use crate::error::{Error, Result, invalid};
use crate::raw::namespace::is_spreadsheetml_name;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::codec::{
    CALC_FEATURES_EXTENSION_URI, CALC_FEATURES_NAMESPACE, Inspection, calc_attribute_index,
    collapse_whitespace, is_namespace_declaration, reject_unsafe_event, xml_error,
};
use super::{Features, Limits, Mode, Properties, ReferenceMode};

const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

#[derive(Clone, Copy, Debug)]
pub(super) struct Span {
    start: usize,
    end: usize,
}

#[derive(Debug)]
struct Element {
    span: Span,
    start_tag_end: usize,
    qname: Vec<u8>,
}

#[derive(Debug)]
struct ExtensionList {
    element: Element,
    start_tag_end: usize,
    close_start: usize,
    empty: bool,
    has_other_content: bool,
}

#[derive(Debug)]
pub(super) struct Layout {
    calc: Option<Element>,
    calc_insert: usize,
    calc_projected: bool,
    ext_list: Option<ExtensionList>,
    target_ext: Option<Element>,
    feature_list: Option<Element>,
    feature_qname: Option<Vec<u8>>,
    features_projected: bool,
    mixed_target: bool,
    root_qname: Vec<u8>,
    root_close: usize,
}

#[derive(Debug)]
enum Kind {
    Root,
    Calc,
    ExtList { other: bool },
    TargetExt { mixed: bool },
    FeatureList,
    Feature,
    Other,
}

#[derive(Debug)]
struct Frame {
    start: usize,
    start_tag_end: usize,
    qname: Vec<u8>,
    kind: Kind,
    projected: bool,
}

pub(super) fn inspect_layout(xml: &[u8], limits: &Limits) -> Result<Layout> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack: Vec<Frame> = Vec::new();
    stack
        .try_reserve(limits.max_depth().min(64))
        .map_err(|source| Error::Allocation {
            resource: "calculation metadata XML stack",
            source,
        })?;
    let mut events = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut root_qname = Vec::new();
    let mut root_close = 0usize;
    let mut calc = None;
    let mut calc_insert = None;
    let mut calc_projected = false;
    let mut ext_list = None;
    let mut target_ext = None;
    let mut feature_list = None;
    let mut feature_qname = None;
    let mut features_projected = false;
    let mut mixed_target = false;
    let mut opaque_bytes = 0usize;
    let mut ext_seen = false;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("XML event count overflow"))?;
        if events > limits.max_events() {
            return Err(invalid("calculation metadata exceeds event count limit"));
        }
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        let end = position(&reader)?;
        reject_unsafe_event(&event)?;
        let resolver = reader.resolver().clone();
        let (_namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                inspect_start(
                    element,
                    false,
                    start,
                    end,
                    decoder,
                    &resolver,
                    limits,
                    &mut stack,
                    &mut root_seen,
                    &mut root_closed,
                    &mut root_qname,
                    &mut calc,
                    &mut calc_insert,
                    &mut calc_projected,
                    &mut ext_seen,
                    &mut ext_list,
                    &mut target_ext,
                    &mut feature_list,
                    &mut feature_qname,
                    &mut features_projected,
                    &mut mixed_target,
                    &mut opaque_bytes,
                )?;
            },
            Event::Empty(element) => {
                inspect_start(
                    element,
                    true,
                    start,
                    end,
                    decoder,
                    &resolver,
                    limits,
                    &mut stack,
                    &mut root_seen,
                    &mut root_closed,
                    &mut root_qname,
                    &mut calc,
                    &mut calc_insert,
                    &mut calc_projected,
                    &mut ext_seen,
                    &mut ext_list,
                    &mut target_ext,
                    &mut feature_list,
                    &mut feature_qname,
                    &mut features_projected,
                    &mut mixed_target,
                    &mut opaque_bytes,
                )?;
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected workbook end element"))?;
                if matches!(frame.kind, Kind::Root) {
                    root_close = start;
                    root_closed = true;
                }
                close_frame(
                    frame,
                    end,
                    start,
                    false,
                    &mut calc,
                    &mut ext_list,
                    &mut target_ext,
                    &mut feature_list,
                    &mut mixed_target,
                    &mut opaque_bytes,
                    limits,
                    stack.last_mut(),
                )?;
            },
            Event::Text(text) => {
                let non_whitespace = !text.decode().map_err(xml_error)?.trim().is_empty();
                if stack.is_empty() && non_whitespace {
                    return Err(invalid("text outside workbook root"));
                }
                if non_whitespace && let Some(frame) = stack.last_mut() {
                    match &mut frame.kind {
                        Kind::TargetExt { mixed } => {
                            *mixed = true;
                            mixed_target = true;
                        },
                        Kind::ExtList { other } => *other = true,
                        _ => {},
                    }
                }
            },
            Event::CData(_) | Event::GeneralRef(_) => {
                if let Some(frame) = stack.last_mut() {
                    match &mut frame.kind {
                        Kind::TargetExt { mixed } => {
                            *mixed = true;
                            mixed_target = true;
                        },
                        Kind::ExtList { other } => *other = true,
                        _ => {},
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("unterminated workbook XML"));
    }
    Ok(Layout {
        calc,
        calc_insert: calc_insert.unwrap_or(root_close),
        calc_projected,
        ext_list,
        target_ext,
        feature_list,
        feature_qname,
        features_projected,
        mixed_target,
        root_qname,
        root_close,
    })
}

#[allow(
    clippy::too_many_arguments,
    reason = "arguments are the complete validated calculation-property attribute set"
)]
fn inspect_start(
    element: BytesStart<'_>,
    empty: bool,
    start: usize,
    end: usize,
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
    limits: &Limits,
    stack: &mut Vec<Frame>,
    root_seen: &mut bool,
    root_closed: &mut bool,
    root_qname: &mut Vec<u8>,
    calc: &mut Option<Element>,
    calc_insert: &mut Option<usize>,
    calc_projected: &mut bool,
    ext_seen: &mut bool,
    ext_list: &mut Option<ExtensionList>,
    target_ext: &mut Option<Element>,
    feature_list: &mut Option<Element>,
    feature_qname: &mut Option<Vec<u8>>,
    features_projected: &mut bool,
    mixed_target: &mut bool,
    opaque_bytes: &mut usize,
) -> Result<()> {
    if *root_closed {
        return Err(invalid("content follows the closed workbook root"));
    }
    let depth = stack.len();
    if depth >= limits.max_depth() {
        return Err(invalid("workbook calculation metadata nesting is too deep"));
    }
    check_attributes(&element, limits)?;
    let (namespace, _) = resolver.resolve_element(element.name());
    let local = element.local_name();
    let core = is_spreadsheetml_name(&namespace, element.name(), local.as_ref());
    let inherited_projection = stack.last().is_some_and(|frame| frame.projected);
    let projected = inherited_projection
        || is_mc(&namespace, local.as_ref(), b"AlternateContent")
        || has_process_content(&element, resolver)?;
    let parent_target = stack
        .last()
        .is_some_and(|frame| matches!(frame.kind, Kind::TargetExt { .. }));
    let parent_features = stack
        .last()
        .is_some_and(|frame| matches!(frame.kind, Kind::FeatureList));
    let parent_ext_list = stack
        .last()
        .is_some_and(|frame| matches!(frame.kind, Kind::ExtList { .. }));
    let qname = element.name().as_ref().to_vec();
    let kind;
    if depth == 0 {
        if *root_seen || !core || local.as_ref() != b"workbook" || empty {
            return Err(invalid(
                "calculation metadata rewriter requires a nonempty workbook root",
            ));
        }
        *root_seen = true;
        root_qname.clone_from(&qname);
        kind = Kind::Root;
    } else if depth == 1 && core && local.as_ref() == b"calcPr" {
        if calc.is_some() {
            return Err(invalid("duplicate physical workbook calcPr element"));
        }
        kind = Kind::Calc;
    } else if depth == 1 && core && local.as_ref() == b"extLst" {
        if *ext_seen {
            return Err(invalid("duplicate physical workbook extLst element"));
        }
        *ext_seen = true;
        kind = Kind::ExtList { other: false };
    } else if parent_ext_list && core && local.as_ref() == b"ext" {
        let (uri, unknown) = raw_extension_attributes(&element, decoder, resolver)?;
        if uri.as_deref() == Some(CALC_FEATURES_EXTENSION_URI) {
            if target_ext.is_some() {
                return Err(invalid("duplicate physical calculation-features extension"));
            }
            *mixed_target |= unknown;
            kind = Kind::TargetExt { mixed: unknown };
        } else {
            kind = Kind::Other;
        }
    } else if parent_target && is_xcalcf(&namespace, local.as_ref(), b"calcFeatures") {
        if feature_list.is_some() {
            return Err(invalid("duplicate physical calcFeatures payload"));
        }
        kind = Kind::FeatureList;
    } else if parent_features && is_xcalcf(&namespace, local.as_ref(), b"feature") {
        if feature_qname.is_none() {
            *feature_qname = Some(qname.clone());
        }
        kind = Kind::Feature;
    } else {
        if parent_target {
            *mixed_target = true;
        }
        kind = Kind::Other;
    }
    if core && local.as_ref() == b"calcPr" && projected {
        *calc_projected = true;
    }
    if projected
        && (matches!(kind, Kind::TargetExt { .. })
            || is_xcalcf(&namespace, local.as_ref(), b"calcFeatures"))
    {
        *features_projected = true;
    }
    if depth == 1 && calc_insert.is_none() && core && follows_calc_pr(local.as_ref()) {
        *calc_insert = Some(start);
    }
    if *ext_seen && depth == 1 && core && local.as_ref() != b"extLst" {
        return Err(invalid("workbook extLst must be the final core child"));
    }
    let frame = Frame {
        start,
        start_tag_end: end,
        qname,
        kind,
        projected,
    };
    if empty {
        close_frame(
            frame,
            end,
            start,
            true,
            calc,
            ext_list,
            target_ext,
            feature_list,
            mixed_target,
            opaque_bytes,
            limits,
            stack.last_mut(),
        )?;
    } else {
        stack.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "calculation metadata XML stack",
            source,
        })?;
        stack.push(frame);
    }
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "arguments are the complete calculation-property rewrite state"
)]
fn close_frame(
    frame: Frame,
    end: usize,
    close_start: usize,
    empty: bool,
    calc: &mut Option<Element>,
    ext_list: &mut Option<ExtensionList>,
    target_ext: &mut Option<Element>,
    feature_list: &mut Option<Element>,
    mixed_target: &mut bool,
    opaque_bytes: &mut usize,
    limits: &Limits,
    parent: Option<&mut Frame>,
) -> Result<()> {
    let span = Span {
        start: frame.start,
        end,
    };
    match frame.kind {
        Kind::Calc => {
            *calc = Some(Element {
                span,
                start_tag_end: frame.start_tag_end,
                qname: frame.qname,
            });
        },
        Kind::ExtList { other } => {
            *ext_list = Some(ExtensionList {
                element: Element {
                    span,
                    start_tag_end: frame.start_tag_end,
                    qname: frame.qname,
                },
                start_tag_end: frame.start_tag_end,
                close_start,
                empty,
                has_other_content: other,
            });
        },
        Kind::TargetExt { mixed } => {
            *mixed_target |= mixed;
            *target_ext = Some(Element {
                span,
                start_tag_end: frame.start_tag_end,
                qname: frame.qname,
            });
        },
        Kind::FeatureList => {
            *feature_list = Some(Element {
                span,
                start_tag_end: frame.start_tag_end,
                qname: frame.qname,
            });
        },
        Kind::Other => {
            if let Some(parent) = parent {
                match &mut parent.kind {
                    Kind::ExtList { other } => {
                        *other = true;
                        *opaque_bytes = opaque_bytes
                            .checked_add(span.end - span.start)
                            .ok_or_else(|| invalid("opaque extension byte count overflow"))?;
                        if *opaque_bytes > limits.max_opaque_bytes() {
                            return Err(invalid("opaque extension bytes exceed limit"));
                        }
                    },
                    Kind::TargetExt { mixed } => {
                        *mixed = true;
                        *mixed_target = true;
                    },
                    _ => {},
                }
            }
        },
        Kind::Root | Kind::Feature => {},
    }
    Ok(())
}

/// Rewrite only the two calculation-metadata owners.
pub(crate) fn rewrite<'a>(
    inspection: &Inspection<'a>,
    properties: Option<&Properties>,
    features: Option<&Features>,
    limits: &Limits,
) -> Result<Cow<'a, [u8]>> {
    let properties_changed = !same_properties(inspection.properties.as_ref(), properties);
    let features_changed = inspection.features.as_ref() != features;
    if !properties_changed && !features_changed {
        return Ok(Cow::Borrowed(inspection.source));
    }
    if properties_changed && inspection.layout.calc_projected {
        return Err(invalid(
            "cannot rewrite calcPr projected through MCE markup",
        ));
    }
    if features_changed && inspection.layout.features_projected {
        return Err(invalid(
            "cannot rewrite calcFeatures projected through MCE markup",
        ));
    }
    if features_changed && inspection.layout.mixed_target {
        return Err(invalid(
            "cannot rewrite mixed opaque calculation-features payload",
        ));
    }

    let mut edits: Vec<Edit> = Vec::new();
    edits.try_reserve(2).map_err(|source| Error::Allocation {
        resource: "calculation metadata rewrite plan",
        source,
    })?;
    if properties_changed {
        let (span, replacement) = match (inspection.layout.calc.as_ref(), properties) {
            (Some(element), Some(value)) => (
                Span {
                    start: element.span.start,
                    end: element.start_tag_end,
                },
                rewrite_properties_start_tag(
                    &inspection.source[element.span.start..element.start_tag_end],
                    value,
                    limits,
                )?,
            ),
            (Some(element), None) => (element.span, Vec::new()),
            (None, Some(value)) => {
                let qname = sibling_name(&inspection.layout.root_qname, b"calcPr");
                (
                    Span {
                        start: inspection.layout.calc_insert,
                        end: inspection.layout.calc_insert,
                    },
                    serialize_properties(&qname, value, limits)?,
                )
            },
            (None, None) => unreachable!(),
        };
        edits.push(Edit {
            span,
            replacement,
            order: 0,
        });
    }
    if features_changed {
        plan_features(inspection, features, limits, &mut edits)?;
    }
    apply_edits(inspection.source, edits, limits)
}

/// Apply the formula-cache invalidation precedence without inferring features.
pub(crate) fn invalidate_formulas<'a>(
    inspection: &Inspection<'a>,
    limits: &Limits,
) -> Result<Cow<'a, [u8]>> {
    let mut properties = inspection.properties.clone().unwrap_or_default();
    properties.set_calculation_id(Some(0));
    properties.set_full_calculation_on_load(Some(true));
    properties.set_force_full_calculation(Some(true));
    properties.set_calculation_completed(Some(false));
    properties.set_calculate_on_save(Some(true));
    rewrite(
        inspection,
        Some(&properties),
        inspection.features.as_ref(),
        limits,
    )
}

fn plan_features(
    inspection: &Inspection<'_>,
    features: Option<&Features>,
    limits: &Limits,
    edits: &mut Vec<Edit>,
) -> Result<()> {
    match (inspection.layout.target_ext.as_ref(), features) {
        (Some(target), None) => {
            let span = if inspection
                .layout
                .ext_list
                .as_ref()
                .is_some_and(|list| !list.has_other_content)
            {
                inspection.layout.ext_list.as_ref().unwrap().element.span
            } else {
                target.span
            };
            edits.push(Edit {
                span,
                replacement: Vec::new(),
                order: 1,
            });
        },
        (Some(_), Some(features)) => {
            let existing =
                inspection.layout.feature_list.as_ref().ok_or_else(|| {
                    invalid("calculation-features extension has no editable payload")
                })?;
            let feature_qname = inspection
                .layout
                .feature_qname
                .as_deref()
                .unwrap_or(existing.qname.as_slice());
            edits.push(Edit {
                span: existing.span,
                replacement: serialize_features(
                    &existing.qname,
                    feature_qname,
                    features,
                    false,
                    limits,
                )?,
                order: 1,
            });
        },
        (None, Some(features)) => {
            let ext_qname = inspection.layout.ext_list.as_ref().map_or_else(
                || sibling_name(&inspection.layout.root_qname, b"ext"),
                |list| sibling_name(&list.element.qname, b"ext"),
            );
            let fragment = serialize_extension(&ext_qname, features, limits)?;
            if let Some(list) = inspection.layout.ext_list.as_ref() {
                if list.empty {
                    let raw = &inspection.source[list.element.span.start..list.element.span.end];
                    let slash = raw
                        .windows(2)
                        .rposition(|value| value == b"/>")
                        .ok_or_else(|| invalid("invalid empty workbook extLst"))?;
                    let mut replacement = Bounded::new(limits.max_output_bytes());
                    replacement.extend(&raw[..slash])?;
                    replacement.push(b'>')?;
                    replacement.extend(&fragment)?;
                    replacement.extend(b"</")?;
                    replacement.extend(&list.element.qname)?;
                    replacement.push(b'>')?;
                    edits.push(Edit {
                        span: list.element.span,
                        replacement: replacement.finish(),
                        order: 1,
                    });
                } else {
                    edits.push(Edit {
                        span: Span {
                            start: list.close_start,
                            end: list.close_start,
                        },
                        replacement: fragment,
                        order: 1,
                    });
                }
            } else {
                let ext_list_qname = sibling_name(&inspection.layout.root_qname, b"extLst");
                let mut replacement = Bounded::new(limits.max_output_bytes());
                replacement.push(b'<')?;
                replacement.extend(&ext_list_qname)?;
                replacement.push(b'>')?;
                replacement.extend(&fragment)?;
                replacement.extend(b"</")?;
                replacement.extend(&ext_list_qname)?;
                replacement.push(b'>')?;
                edits.push(Edit {
                    span: Span {
                        start: inspection.layout.root_close,
                        end: inspection.layout.root_close,
                    },
                    replacement: replacement.finish(),
                    order: 1,
                });
            }
        },
        (None, None) => {},
    }
    Ok(())
}

#[derive(Debug)]
struct Edit {
    span: Span,
    replacement: Vec<u8>,
    order: u8,
}

fn apply_edits<'a>(
    source: &'a [u8],
    mut edits: Vec<Edit>,
    limits: &Limits,
) -> Result<Cow<'a, [u8]>> {
    edits.sort_by_key(|edit| (edit.span.start, edit.order));
    let mut final_len = source.len();
    let mut previous_end = 0usize;
    for edit in &edits {
        if edit.span.start < previous_end
            || edit.span.end < edit.span.start
            || edit.span.end > source.len()
        {
            return Err(invalid(
                "overlapping or invalid calculation metadata rewrite spans",
            ));
        }
        previous_end = edit.span.end;
        final_len = final_len
            .checked_sub(edit.span.end - edit.span.start)
            .and_then(|value| value.checked_add(edit.replacement.len()))
            .ok_or_else(|| invalid("calculation metadata output size overflow"))?;
    }
    if final_len > limits.max_output_bytes() {
        return Err(invalid("calculation metadata output exceeds byte limit"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(final_len)
        .map_err(|source| Error::Allocation {
            resource: "calculation metadata rewrite",
            source,
        })?;
    let mut cursor = 0usize;
    for edit in edits {
        output.extend_from_slice(&source[cursor..edit.span.start]);
        output.extend_from_slice(&edit.replacement);
        cursor = edit.span.end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(Cow::Owned(output))
}

fn rewrite_properties_start_tag(
    raw: &[u8],
    value: &Properties,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let desired = property_values(value);
    let mut seen = [false; 13];
    let mut output = Bounded::new(limits.max_output_bytes());
    let mut position = 1usize;
    while position < raw.len()
        && !raw[position].is_ascii_whitespace()
        && !matches!(raw[position], b'>' | b'/')
    {
        position += 1;
    }
    if position <= 1 || position >= raw.len() {
        return Err(invalid("invalid calcPr start tag"));
    }
    let mut cursor = 0usize;
    let tail_start;
    loop {
        let whitespace_start = position;
        while position < raw.len() && raw[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= raw.len() || matches!(raw[position], b'>' | b'/') {
            tail_start = whitespace_start;
            break;
        }
        let name_start = position;
        while position < raw.len()
            && !raw[position].is_ascii_whitespace()
            && !matches!(raw[position], b'=' | b'>' | b'/')
        {
            position += 1;
        }
        let name_end = position;
        while position < raw.len() && raw[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= raw.len() || raw[position] != b'=' {
            return Err(invalid("invalid calcPr attribute assignment"));
        }
        position += 1;
        while position < raw.len() && raw[position].is_ascii_whitespace() {
            position += 1;
        }
        if position >= raw.len() || !matches!(raw[position], b'\'' | b'"') {
            return Err(invalid("invalid calcPr attribute quote"));
        }
        let quote = raw[position];
        position += 1;
        let value_start = position;
        while position < raw.len() && raw[position] != quote {
            position += 1;
        }
        if position >= raw.len() {
            return Err(invalid("unterminated calcPr attribute value"));
        }
        let value_end = position;
        position += 1;
        let name = &raw[name_start..name_end];
        if !name.contains(&b':')
            && let Some((index, _)) = calc_attribute_index(name)
        {
            if std::mem::replace(&mut seen[index], true) {
                return Err(invalid("duplicate calcPr attribute during rewrite"));
            }
            output.extend(&raw[cursor..whitespace_start])?;
            if let Some(replacement) = desired[index].as_deref() {
                output.extend(&raw[whitespace_start..value_start])?;
                escape_attribute(&mut output, replacement)?;
                cursor = value_end;
            } else {
                cursor = position;
            }
        }
    }
    output.extend(&raw[cursor..tail_start])?;
    for (index, replacement) in desired.iter().enumerate() {
        if !seen[index]
            && let Some(replacement) = replacement.as_deref()
        {
            attr(&mut output, calc_attribute_name(index), replacement)?;
        }
    }
    output.extend(&raw[tail_start..])?;
    Ok(output.finish())
}

fn property_values(value: &Properties) -> [Option<String>; 13] {
    let specified = value.specified();
    [
        specified.calculation_id().map(|value| value.to_string()),
        specified
            .calculation_mode()
            .map(|value| mode(value).to_owned()),
        specified.full_calculation_on_load().map(bool_lexical),
        specified
            .reference_mode()
            .map(|value| reference_mode(value).to_owned()),
        specified.iterative_calculation().map(bool_lexical),
        specified.iteration_count().map(|value| value.to_string()),
        specified.iteration_delta().map(double_lexical),
        specified.full_precision().map(bool_lexical),
        specified.calculation_completed().map(bool_lexical),
        specified.calculate_on_save().map(bool_lexical),
        specified.concurrent_calculation().map(bool_lexical),
        specified
            .concurrent_manual_count()
            .map(|value| value.to_string()),
        specified.force_full_calculation().map(bool_lexical),
    ]
}

fn bool_lexical(value: bool) -> String {
    if value { "true" } else { "false" }.to_owned()
}

fn double_lexical(value: f64) -> String {
    if value.is_nan() {
        "NaN".to_owned()
    } else if value == f64::INFINITY {
        "INF".to_owned()
    } else if value == f64::NEG_INFINITY {
        "-INF".to_owned()
    } else if value == 0.0 && value.is_sign_negative() {
        "-0".to_owned()
    } else {
        value.to_string()
    }
}

fn calc_attribute_name(index: usize) -> &'static [u8] {
    match index {
        0 => b"calcId",
        1 => b"calcMode",
        2 => b"fullCalcOnLoad",
        3 => b"refMode",
        4 => b"iterate",
        5 => b"iterateCount",
        6 => b"iterateDelta",
        7 => b"fullPrecision",
        8 => b"calcCompleted",
        9 => b"calcOnSave",
        10 => b"concurrentCalc",
        11 => b"concurrentManualCount",
        12 => b"forceFullCalc",
        _ => unreachable!(),
    }
}

fn serialize_properties(qname: &[u8], value: &Properties, limits: &Limits) -> Result<Vec<u8>> {
    let mut output = Bounded::new(limits.max_output_bytes());
    output.push(b'<')?;
    output.extend(qname)?;
    let specified = value.specified();
    if let Some(value) = specified.calculation_id() {
        attr(&mut output, b"calcId", &value.to_string())?;
    }
    if let Some(value) = specified.calculation_mode() {
        attr(&mut output, b"calcMode", mode(value))?;
    }
    if let Some(value) = specified.full_calculation_on_load() {
        bool_attr(&mut output, b"fullCalcOnLoad", value)?;
    }
    if let Some(value) = specified.reference_mode() {
        attr(&mut output, b"refMode", reference_mode(value))?;
    }
    if let Some(value) = specified.iterative_calculation() {
        bool_attr(&mut output, b"iterate", value)?;
    }
    if let Some(value) = specified.iteration_count() {
        attr(&mut output, b"iterateCount", &value.to_string())?;
    }
    if let Some(value) = specified.iteration_delta() {
        attr(&mut output, b"iterateDelta", &double_lexical(value))?;
    }
    if let Some(value) = specified.full_precision() {
        bool_attr(&mut output, b"fullPrecision", value)?;
    }
    if let Some(value) = specified.calculation_completed() {
        bool_attr(&mut output, b"calcCompleted", value)?;
    }
    if let Some(value) = specified.calculate_on_save() {
        bool_attr(&mut output, b"calcOnSave", value)?;
    }
    if let Some(value) = specified.concurrent_calculation() {
        bool_attr(&mut output, b"concurrentCalc", value)?;
    }
    if let Some(value) = specified.concurrent_manual_count() {
        attr(&mut output, b"concurrentManualCount", &value.to_string())?;
    }
    if let Some(value) = specified.force_full_calculation() {
        bool_attr(&mut output, b"forceFullCalc", value)?;
    }
    output.extend(b"/>")?;
    Ok(output.finish())
}

fn serialize_extension(ext_qname: &[u8], features: &Features, limits: &Limits) -> Result<Vec<u8>> {
    let mut output = Bounded::new(limits.max_output_bytes());
    output.push(b'<')?;
    output.extend(ext_qname)?;
    attr(&mut output, b"uri", CALC_FEATURES_EXTENSION_URI)?;
    output.push(b'>')?;
    write_features(
        &mut output,
        b"xcalcf:calcFeatures",
        b"xcalcf:feature",
        features,
        true,
        limits,
    )?;
    output.extend(b"</")?;
    output.extend(ext_qname)?;
    output.push(b'>')?;
    Ok(output.finish())
}

fn serialize_features(
    list_qname: &[u8],
    feature_qname: &[u8],
    features: &Features,
    declare_namespace: bool,
    limits: &Limits,
) -> Result<Vec<u8>> {
    let mut output = Bounded::new(limits.max_output_bytes());
    write_features(
        &mut output,
        list_qname,
        feature_qname,
        features,
        declare_namespace,
        limits,
    )?;
    Ok(output.finish())
}

fn write_features(
    output: &mut Bounded,
    list_qname: &[u8],
    feature_qname: &[u8],
    features: &Features,
    declare_namespace: bool,
    limits: &Limits,
) -> Result<()> {
    if features.as_slice().is_empty() || features.len() > limits.max_features() {
        return Err(invalid(
            "calculation features are empty or exceed count limit",
        ));
    }
    let mut total = 0usize;
    output.push(b'<')?;
    output.extend(list_qname)?;
    if declare_namespace {
        output.extend(b" xmlns:xcalcf=\"")?;
        output.extend(CALC_FEATURES_NAMESPACE)?;
        output.push(b'"')?;
    }
    output.push(b'>')?;
    for feature in features {
        let len = feature.as_str().len();
        if len > limits.max_feature_name_bytes() {
            return Err(invalid("calculation feature name exceeds byte limit"));
        }
        total = total
            .checked_add(len)
            .ok_or_else(|| invalid("feature byte count overflow"))?;
        if total > limits.max_feature_names_bytes() {
            return Err(invalid(
                "calculation feature names exceed aggregate byte limit",
            ));
        }
        output.push(b'<')?;
        output.extend(feature_qname)?;
        output.extend(b" name=\"")?;
        escape_attribute(output, feature.as_str())?;
        output.extend(b"\"/>")?;
    }
    output.extend(b"</")?;
    output.extend(list_qname)?;
    output.push(b'>')?;
    Ok(())
}

struct Bounded {
    bytes: Vec<u8>,
    max: usize,
}

impl Bounded {
    fn new(max: usize) -> Self {
        Self {
            bytes: Vec::new(),
            max,
        }
    }
    fn reserve(&mut self, additional: usize) -> Result<()> {
        let wanted = self
            .bytes
            .len()
            .checked_add(additional)
            .ok_or_else(|| invalid("serialized calculation metadata size overflow"))?;
        if wanted > self.max {
            return Err(invalid(
                "serialized calculation metadata exceeds output limit",
            ));
        }
        if additional > self.bytes.capacity().saturating_sub(self.bytes.len()) {
            self.bytes
                .try_reserve(additional)
                .map_err(|source| Error::Allocation {
                    resource: "serialized calculation metadata",
                    source,
                })?;
        }
        Ok(())
    }
    fn extend(&mut self, value: &[u8]) -> Result<()> {
        self.reserve(value.len())?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }
    fn push(&mut self, value: u8) -> Result<()> {
        self.reserve(1)?;
        self.bytes.push(value);
        Ok(())
    }
    fn finish(self) -> Vec<u8> {
        self.bytes
    }
}

fn attr(output: &mut Bounded, name: &[u8], value: &str) -> Result<()> {
    output.push(b' ')?;
    output.extend(name)?;
    output.extend(b"=\"")?;
    escape_attribute(output, value)?;
    output.push(b'"')
}

fn bool_attr(output: &mut Bounded, name: &[u8], value: bool) -> Result<()> {
    attr(output, name, if value { "true" } else { "false" })
}

fn escape_attribute(output: &mut Bounded, value: &str) -> Result<()> {
    let mut rest = value.as_bytes();
    while let Some(index) = rest
        .iter()
        .position(|byte| matches!(byte, b'&' | b'<' | b'"' | b'\r' | b'\n' | b'\t'))
    {
        output.extend(&rest[..index])?;
        output.extend(match rest[index] {
            b'&' => b"&amp;",
            b'<' => b"&lt;",
            b'"' => b"&quot;",
            b'\r' => b"&#xD;",
            b'\n' => b"&#xA;",
            b'\t' => b"&#x9;",
            _ => unreachable!(),
        })?;
        rest = &rest[index + 1..];
    }
    output.extend(rest)
}

fn same_properties(left: Option<&Properties>, right: Option<&Properties>) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => left.same_specification(right),
        (None, None) => true,
        _ => false,
    }
}

fn mode(value: Mode) -> &'static str {
    match value {
        Mode::Manual => "manual",
        Mode::Automatic => "auto",
        Mode::AutomaticExceptTables => "autoNoTable",
    }
}

fn reference_mode(value: ReferenceMode) -> &'static str {
    match value {
        ReferenceMode::A1 => "A1",
        ReferenceMode::R1C1 => "R1C1",
    }
}

fn sibling_name(name: &[u8], local: &[u8]) -> Vec<u8> {
    match name.iter().rposition(|byte| *byte == b':') {
        Some(index) => {
            let mut result = Vec::with_capacity(index + 1 + local.len());
            result.extend_from_slice(&name[..=index]);
            result.extend_from_slice(local);
            result
        },
        None => local.to_vec(),
    }
}

fn follows_calc_pr(local: &[u8]) -> bool {
    matches!(
        local,
        b"oleSize"
            | b"customWorkbookViews"
            | b"pivotCaches"
            | b"smartTagPr"
            | b"smartTagTypes"
            | b"webPublishing"
            | b"fileRecoveryPr"
            | b"webPublishObjects"
            | b"extLst"
    )
}

fn raw_extension_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<(Option<String>, bool)> {
    let mut uri = None;
    let mut unknown = false;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Unbound) && local.as_ref() == b"uri" {
            if uri.is_some() {
                return Err(invalid("duplicate ext uri attribute"));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(xml_error)?;
            uri = Some(collapse_whitespace(&value).into_owned());
        } else {
            unknown = true;
        }
    }
    Ok((uri, unknown))
}

fn has_process_content(
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<bool> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if local.as_ref() == b"ProcessContent"
            && matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == MCE_NAMESPACE)
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn check_attributes(element: &BytesStart<'_>, limits: &Limits) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.attributes().with_checks(false) {
        attribute.map_err(xml_error)?;
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("attribute count overflow"))?;
        if count > limits.max_attributes() {
            return Err(invalid("calculation metadata exceeds attribute limit"));
        }
    }
    Ok(())
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("XML position does not fit usize"))
}

fn is_xcalcf(namespace: &ResolveResult<'_>, local: &[u8], expected: &[u8]) -> bool {
    local == expected
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == CALC_FEATURES_NAMESPACE)
}

fn is_mc(namespace: &ResolveResult<'_>, local: &[u8], expected: &[u8]) -> bool {
    local == expected
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE_NAMESPACE)
}
