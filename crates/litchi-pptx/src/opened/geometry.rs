//! Byte-preserving geometry edits for ordinary PresentationML shapes.
//!
//! The shape reader intentionally tolerates a malformed transform by dropping
//! only that shape's bounds.  An authoring operation has a stricter contract:
//! it must identify one direct `p:spPr/a:xfrm`, prove all four coordinate
//! attributes, and change only their lexical value spans.  This module keeps
//! that distinction explicit.  It never serializes a shape or reconstructs a
//! namespace context from a guessed prefix.

use std::collections::HashSet;
use std::ops::Range;

use litchi_core::Position;
use litchi_core::xml::ReaderOrigin;
use litchi_drawingml::coordinate::{Coordinate, Extent};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

const PML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const DML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";

/// The four raw value spans belonging to one direct DrawingML transform.
#[derive(Debug)]
struct TransformGeometry {
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    off_seen: bool,
    ext_seen: bool,
    x_span: Range<usize>,
    y_span: Range<usize>,
    width_span: Range<usize>,
    height_span: Range<usize>,
}

/// The direct geometry surface observed below one native `p:sp`.
#[derive(Debug)]
struct ShapeGeometry {
    start: usize,
    end: Option<usize>,
    shape_properties: usize,
    transforms: usize,
    transform: Option<TransformGeometry>,
    malformed: bool,
    /// Whether the transform is a graphic frame's own `p:xfrm`.
    frame: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ElementKind {
    Other,
    Shape,
    ShapeProperties,
    Transform,
    Offset,
    Extent,
}

#[derive(Debug)]
struct ElementFrame {
    kind: ElementKind,
    shape: Option<usize>,
    /// `a:off` and `a:ext` are empty schema elements.  A non-empty form is
    /// retained lexically only when it contains no child event at all;
    /// otherwise the known transform is malformed and cannot be authored.
    leaf: bool,
}

#[derive(Debug)]
struct Replacement {
    span: Range<usize>,
    value: Vec<u8>,
}

/// Rewrite complete local bounds of one pre-order ordinary shape.
///
/// The returned bytes retain every source byte except the four unqualified
/// numeric attribute values in the selected shape's direct transform.  An
/// exact semantic no-op returns `None`.  MCE-selected owners and non-native
/// shape kinds are refused before any output allocation.
#[cfg(test)]
pub(super) fn rewrite_shape_bounds(
    xml: &[u8],
    position: Position,
    bounds: (i64, i64, i64, i64),
    max_output_bytes: usize,
) -> Result<Option<Vec<u8>>> {
    rewrite_shape_bounds_many(xml, &[(position, bounds)], max_output_bytes)
}

/// Rewrite a checked batch of ordinary-shape bounds in one scene scan and one
/// output allocation.  Callers use this for homogeneous property batches so
/// the work remains proportional to the owner XML plus the number of spans,
/// rather than reparsing the owner once per item.
#[cfg(test)]
pub(super) fn rewrite_shape_bounds_many(
    xml: &[u8],
    replacements: &[(Position, (i64, i64, i64, i64))],
    max_output_bytes: usize,
) -> Result<Option<Vec<u8>>> {
    rewrite_shape_bounds_many_with_limits(
        xml,
        replacements,
        max_output_bytes,
        crate::shape::Limits::DEFAULT,
    )
}

/// Rewrite a checked batch under the caller's retained shape policy.
pub(super) fn rewrite_shape_bounds_many_with_limits(
    xml: &[u8],
    replacements: &[(Position, (i64, i64, i64, i64))],
    max_output_bytes: usize,
    shape_limits: crate::shape::Limits,
) -> Result<Option<Vec<u8>>> {
    if replacements.is_empty() {
        return Ok(None);
    }
    for &(_, bounds) in replacements {
        validate_requested_bounds(bounds)?;
    }

    let scene = crate::shape::Scene::read(xml, shape_limits)?;
    refuse_rewritten_scene(&scene)?;
    let scanned = scan_geometry(xml)?;

    let mut selected = HashSet::new();
    selected
        .try_reserve(replacements.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation geometry shape identities",
            source,
        })?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(replacements.len().saturating_mul(4))
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation geometry edit plan",
            source,
        })?;
    let mut selected_values = Vec::new();
    selected_values
        .try_reserve_exact(replacements.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation geometry readback plan",
            source,
        })?;

    for &(position, requested) in replacements {
        let index = position.get();
        if !selected.insert(index) {
            return Err(Error::Invalid(
                "opened-presentation geometry batch selects a shape more than once".into(),
            ));
        }
        let shape = scene.at(index)?;
        if !matches!(
            shape,
            crate::shape::Shape::Auto(_)
                | crate::shape::Shape::Chart(_)
                | crate::shape::Shape::Picture(_)
        ) {
            return Err(Error::Invalid(
                "opened-presentation geometry authoring supports ordinary p:sp shapes, pictures and chart frames only".into(),
            ));
        }
        let span = shape.common().span()?.range(xml.len())?;
        let record = find_record(&scanned, &span)?;
        let current = checked_transform(record, true)?.ok_or_else(|| {
            Error::Invalid("opened-presentation selected shape has no direct geometry".into())
        })?;
        validate_typed_transform(shape)?;
        let current_bounds = (current.x, current.y, current.width, current.height);
        selected_values.push((position, requested));
        if current_bounds == requested {
            continue;
        }
        append_replacements(&mut edits, current, requested)?;
    }

    if edits.is_empty() {
        return Ok(None);
    }
    edits.sort_unstable_by_key(|edit| (edit.span.start, edit.span.end));
    let output = emit_replacements(xml, &edits, max_output_bytes)?;
    let published = crate::shape::Scene::read(&output, shape_limits)?;
    refuse_rewritten_scene(&published)?;
    if !same_shape_identity(&scene, &published) {
        return Err(Error::Invalid(
            "opened-presentation geometry rewrite changed shape identity or count".into(),
        ));
    }
    for (position, requested) in selected_values {
        let shape = published.at(position.get())?;
        validate_typed_transform(shape)?;
        let actual = shape
            .common()
            .bounds()
            .map(|value| (value.x(), value.y(), value.width(), value.height()));
        if actual != Some(requested) {
            return Err(Error::Invalid(
                "opened-presentation geometry rewrite did not round-trip bounds".into(),
            ));
        }
    }
    Ok(Some(output))
}

/// Build a proof surface that masks only validated direct transform values of
/// ordinary shapes.  The result intentionally retains all tags, attributes,
/// namespaces, comments, and unknown extension bytes exactly.
#[cfg(test)]
pub(crate) fn geometry_skeleton(xml: &[u8]) -> Result<Vec<u8>> {
    geometry_skeleton_with_limits(xml, crate::shape::Limits::DEFAULT)
}

/// Build a geometry proof surface under the caller's retained shape policy.
pub(crate) fn geometry_skeleton_with_limits(
    xml: &[u8],
    shape_limits: crate::shape::Limits,
) -> Result<Vec<u8>> {
    let scene = crate::shape::Scene::read(xml, shape_limits)?;
    refuse_rewritten_scene(&scene)?;
    let scanned = scan_geometry(xml)?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(scene.len().saturating_mul(4))
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation geometry skeleton plan",
            source,
        })?;

    for shape in scene.iter() {
        if !matches!(shape, crate::shape::Shape::Auto(_)) {
            continue;
        }
        let span = shape.common().span()?.range(xml.len())?;
        let record = find_record(&scanned, &span)?;
        let Ok(Some(transform)) = checked_transform(record, false) else {
            continue;
        };
        if !typed_transform_is_valid(shape) {
            continue;
        }
        // One-byte canonical values keep the skeleton small while remaining
        // valid Scene geometry.  Zero is accepted for a pre-existing extent;
        // the authoring path itself requires positive requested extents.
        edits.push(Replacement {
            span: transform.x_span.clone(),
            value: b"0".to_vec(),
        });
        edits.push(Replacement {
            span: transform.y_span.clone(),
            value: b"0".to_vec(),
        });
        edits.push(Replacement {
            span: transform.width_span.clone(),
            value: b"1".to_vec(),
        });
        edits.push(Replacement {
            span: transform.height_span.clone(),
            value: b"1".to_vec(),
        });
    }

    if edits.is_empty() {
        let mut output = Vec::new();
        output
            .try_reserve_exact(xml.len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation geometry skeleton XML",
                source,
            })?;
        output.extend_from_slice(xml);
        return Ok(output);
    }
    edits.sort_unstable_by_key(|edit| (edit.span.start, edit.span.end));
    emit_replacements(xml, &edits, usize::MAX)
}

fn validate_requested_bounds((x, y, width, height): (i64, i64, i64, i64)) -> Result<()> {
    if width <= 0 || height <= 0 {
        return Err(Error::Invalid(
            "opened-presentation geometry extents must be positive".into(),
        ));
    }
    validate_bounds_domain((x, y, width, height), true)?;
    Ok(())
}

fn validate_bounds_domain(
    (x, y, width, height): (i64, i64, i64, i64),
    positive_extents: bool,
) -> Result<()> {
    let _ = Coordinate::emu(x).map_err(|_| {
        Error::Invalid("opened-presentation geometry coordinate is outside DrawingML range".into())
    })?;
    let _ = Coordinate::emu(y).map_err(|_| {
        Error::Invalid("opened-presentation geometry coordinate is outside DrawingML range".into())
    })?;
    let width_value = Extent::emu(width).map_err(|_| {
        Error::Invalid("opened-presentation geometry extent is outside DrawingML range".into())
    })?;
    let height_value = Extent::emu(height).map_err(|_| {
        Error::Invalid("opened-presentation geometry extent is outside DrawingML range".into())
    })?;
    if positive_extents && (width_value.as_emu() == 0 || height_value.as_emu() == 0) {
        return Err(Error::Invalid(
            "opened-presentation geometry extents must be positive".into(),
        ));
    }
    Ok(())
}

fn validate_typed_transform(shape: crate::shape::Shape<'_>) -> Result<()> {
    // A graphic frame's transform has no rotation or flips: its bounds are
    // its whole typed geometry.
    if matches!(shape, crate::shape::Shape::Chart(_)) {
        return match shape.common().bounds() {
            Some(_) => Ok(()),
            None => Err(Error::Invalid(
                "opened-presentation selected chart frame has no direct p:xfrm geometry".into(),
            )),
        };
    }
    match shape.common().transform() {
        Ok(Some(_)) => Ok(()),
        Ok(None) => Err(Error::Invalid(
            "opened-presentation selected shape has no direct a:xfrm geometry".into(),
        )),
        Err(_) => Err(Error::Invalid(
            "opened-presentation shape transform is malformed or unsupported".into(),
        )),
    }
}

fn typed_transform_is_valid(shape: crate::shape::Shape<'_>) -> bool {
    matches!(shape.common().transform(), Ok(Some(_)))
}

fn refuse_rewritten_scene(scene: &crate::shape::Scene<'_>) -> Result<()> {
    if scene.is_rewritten() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "rewrite_shape_bounds",
            "markup-compatibility branch selection is not authorable lexically",
        ));
    }
    Ok(())
}

/// Validate the transform record and, for authoring, require that it exists.
/// Scene accepts zero extents for old source documents, so proof normalization
/// accepts them too; negative extents and all incomplete/duplicate transforms
/// remain malformed.
fn checked_transform(record: &ShapeGeometry, required: bool) -> Result<Option<&TransformGeometry>> {
    if record.malformed || record.shape_properties > 1 || record.transforms > 1 {
        return Err(Error::Invalid(
            "opened-presentation shape has a malformed direct transform owner".into(),
        ));
    }
    if record.transforms == 0 || (record.shape_properties == 0 && !record.frame) {
        if required {
            return Err(Error::Invalid(
                "opened-presentation selected shape has no direct a:xfrm geometry".into(),
            ));
        }
        return Ok(None);
    }
    let transform = record.transform.as_ref().ok_or_else(|| {
        Error::Invalid("opened-presentation direct transform owner is missing".into())
    })?;
    if !transform.off_seen || !transform.ext_seen {
        return Err(Error::Invalid(
            "opened-presentation direct transform is missing off or ext".into(),
        ));
    }
    if transform.width < 0 || transform.height < 0 {
        return Err(Error::Invalid(
            "opened-presentation shape transform extents are negative".into(),
        ));
    }
    validate_bounds_domain(
        (transform.x, transform.y, transform.width, transform.height),
        false,
    )?;
    Ok(Some(transform))
}

fn append_replacements(
    edits: &mut Vec<Replacement>,
    transform: &TransformGeometry,
    (x, y, width, height): (i64, i64, i64, i64),
) -> Result<()> {
    if transform.x != x {
        edits.push(Replacement {
            span: transform.x_span.clone(),
            value: x.to_string().into_bytes(),
        });
    }
    if transform.y != y {
        edits.push(Replacement {
            span: transform.y_span.clone(),
            value: y.to_string().into_bytes(),
        });
    }
    if transform.width != width {
        edits.push(Replacement {
            span: transform.width_span.clone(),
            value: width.to_string().into_bytes(),
        });
    }
    if transform.height != height {
        edits.push(Replacement {
            span: transform.height_span.clone(),
            value: height.to_string().into_bytes(),
        });
    }
    Ok(())
}

fn emit_replacements(
    xml: &[u8],
    edits: &[Replacement],
    max_output_bytes: usize,
) -> Result<Vec<u8>> {
    let mut removed = 0usize;
    let mut emitted = 0usize;
    let mut cursor = 0usize;
    for edit in edits {
        if edit.span.start < cursor || edit.span.start > edit.span.end || edit.span.end > xml.len()
        {
            return Err(Error::Invalid(
                "opened-presentation geometry value spans overlap or leave the owner".into(),
            ));
        }
        cursor = edit.span.end;
        removed = removed
            .checked_add(edit.span.len())
            .ok_or_else(|| Error::Invalid("opened-presentation geometry size overflow".into()))?;
        emitted = emitted
            .checked_add(edit.value.len())
            .ok_or_else(|| Error::Invalid("opened-presentation geometry size overflow".into()))?;
    }
    let output_len = xml
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(emitted))
        .ok_or_else(|| Error::Invalid("opened-presentation geometry size overflow".into()))?;
    if output_len > max_output_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: output_len,
            maximum: max_output_bytes,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation geometry XML",
            source,
        })?;
    cursor = 0;
    for edit in edits {
        output.extend_from_slice(&xml[cursor..edit.span.start]);
        output.extend_from_slice(&edit.value);
        cursor = edit.span.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn same_shape_identity(before: &crate::shape::Scene<'_>, after: &crate::shape::Scene<'_>) -> bool {
    before.len() == after.len()
        && before.iter().zip(after.iter()).all(|(left, right)| {
            std::mem::discriminant(&left) == std::mem::discriminant(&right)
                && left.id() == right.id()
                && left.name() == right.name()
                && left.common().source_name() == right.common().source_name()
        })
}

fn scan_geometry(xml: &[u8]) -> Result<Vec<ShapeGeometry>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let origin = ReaderOrigin::of(xml);
    let mut stack: Vec<ElementFrame> = Vec::new();
    stack.try_reserve(32).map_err(|source| Error::Allocation {
        resource: "opened-presentation geometry XML stack",
        source,
    })?;
    let mut shapes = Vec::new();
    loop {
        let start = xml_position(&reader, origin)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|_| malformed_geometry_xml())?;
        let classified = match &event {
            Event::Start(element) | Event::Empty(element) => {
                Some(classify_element(&namespace, element, &stack))
            },
            _ => None,
        };
        let end = xml_position(&reader, origin)?;

        let leaf_owner = stack
            .last()
            .filter(|frame| frame.leaf)
            .and_then(|frame| frame.shape);
        if leaf_owner.is_some() && !matches!(&event, Event::End(_)) {
            mark_malformed(&mut shapes, leaf_owner);
        }

        match event {
            Event::Start(element) => {
                let kind = classified.ok_or_else(|| {
                    Error::Invalid("opened-presentation geometry event lost its kind".into())
                })?;
                let shape = start_element(
                    &mut shapes,
                    &stack,
                    kind,
                    &element,
                    xml,
                    start,
                    end,
                    reader.decoder(),
                )?;
                stack.push(ElementFrame {
                    kind,
                    shape,
                    leaf: matches!(kind, ElementKind::Offset | ElementKind::Extent),
                });
            },
            Event::Empty(element) => {
                let kind = classified.ok_or_else(|| {
                    Error::Invalid("opened-presentation geometry event lost its kind".into())
                })?;
                let shape = start_element(
                    &mut shapes,
                    &stack,
                    kind,
                    &element,
                    xml,
                    start,
                    end,
                    reader.decoder(),
                )?;
                if kind == ElementKind::Shape {
                    if let Some(index) = shape {
                        shapes[index].end = Some(end);
                    }
                }
            },
            Event::End(_) => {
                let frame = stack.pop().ok_or_else(|| {
                    Error::Invalid("opened-presentation geometry XML stack underflow".into())
                })?;
                if frame.kind == ElementKind::Shape {
                    if let Some(index) = frame.shape {
                        shapes[index].end = Some(end);
                    }
                }
            },
            Event::Eof => {
                if !stack.is_empty() {
                    return Err(Error::Invalid(
                        "opened-presentation geometry XML is unterminated".into(),
                    ));
                }
                break;
            },
            _ => {},
        }
    }
    Ok(shapes)
}

fn start_element(
    shapes: &mut Vec<ShapeGeometry>,
    stack: &[ElementFrame],
    kind: ElementKind,
    element: &BytesStart<'_>,
    xml: &[u8],
    start: usize,
    end: usize,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<usize>> {
    let parent = stack.last();
    let inherited_shape = parent.and_then(|frame| frame.shape);
    if kind == ElementKind::Shape {
        let index = shapes.len();
        shapes.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "opened-presentation geometry shape records",
            source,
        })?;
        shapes.push(ShapeGeometry {
            start,
            end: None,
            shape_properties: 0,
            transforms: 0,
            transform: None,
            malformed: false,
            frame: false,
        });
        return Ok(Some(index));
    }

    let Some(shape) = inherited_shape else {
        return Ok(None);
    };
    let direct = |expected: ElementKind| parent.is_some_and(|frame| frame.kind == expected);
    match kind {
        ElementKind::ShapeProperties if direct(ElementKind::Shape) => {
            let record = &mut shapes[shape];
            record.shape_properties = record.shape_properties.saturating_add(1);
            if record.shape_properties > 1 {
                record.malformed = true;
            }
        },
        ElementKind::Transform
            if direct(ElementKind::ShapeProperties) || direct(ElementKind::Shape) =>
        {
            let record = &mut shapes[shape];
            record.frame |= direct(ElementKind::Shape);
            record.transforms = record.transforms.saturating_add(1);
            if record.transforms == 1 {
                record.transform = Some(TransformGeometry::new());
            } else {
                record.malformed = true;
            }
        },
        ElementKind::Offset if direct(ElementKind::Transform) => {
            let record = &mut shapes[shape];
            if record.transforms != 1 {
                record.malformed = true;
            } else if let Some(transform) = record.transform.as_mut() {
                if transform.off_seen {
                    record.malformed = true;
                } else {
                    let pair = parse_coordinate_pair(element, xml, start, end, b"x", b"y", decoder);
                    match pair {
                        Ok(pair) => {
                            transform.off_seen = true;
                            transform.x = pair.first;
                            transform.y = pair.second;
                            transform.x_span = pair.first_span;
                            transform.y_span = pair.second_span;
                        },
                        Err(_) => record.malformed = true,
                    }
                }
            } else {
                record.malformed = true;
            }
        },
        ElementKind::Extent if direct(ElementKind::Transform) => {
            let record = &mut shapes[shape];
            if record.transforms != 1 {
                record.malformed = true;
            } else if let Some(transform) = record.transform.as_mut() {
                if transform.ext_seen {
                    record.malformed = true;
                } else {
                    let pair =
                        parse_coordinate_pair(element, xml, start, end, b"cx", b"cy", decoder);
                    match pair {
                        Ok(pair) => {
                            transform.ext_seen = true;
                            transform.width = pair.first;
                            transform.height = pair.second;
                            transform.width_span = pair.first_span;
                            transform.height_span = pair.second_span;
                        },
                        Err(_) => record.malformed = true,
                    }
                }
            } else {
                record.malformed = true;
            }
        },
        _ => {},
    }
    Ok(Some(shape))
}

impl TransformGeometry {
    fn new() -> Self {
        TransformGeometry {
            x: 0,
            y: 0,
            width: 0,
            height: 0,
            off_seen: false,
            ext_seen: false,
            x_span: 0..0,
            y_span: 0..0,
            width_span: 0..0,
            height_span: 0..0,
        }
    }
}

fn parse_coordinate_pair(
    element: &BytesStart<'_>,
    xml: &[u8],
    start: usize,
    end: usize,
    first_name: &[u8],
    second_name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<CoordinatePair> {
    let raw = xml.get(start..end).ok_or_else(|| {
        Error::Invalid("opened-presentation geometry start tag is outside the owner".into())
    })?;
    let first = parse_coordinate(element, raw, start, first_name, decoder)?;
    let second = parse_coordinate(element, raw, start, second_name, decoder)?;
    Ok(CoordinatePair {
        first: first.0,
        second: second.0,
        first_span: first.1,
        second_span: second.1,
    })
}

#[derive(Debug)]
struct CoordinatePair {
    first: i64,
    second: i64,
    first_span: Range<usize>,
    second_span: Range<usize>,
}

fn parse_coordinate(
    element: &BytesStart<'_>,
    raw: &[u8],
    tag_start: usize,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<(i64, Range<usize>)> {
    let mut decoded = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|_| malformed_geometry_xml())?;
        if attribute.key.prefix().is_none() && attribute.key.local_name().as_ref() == name {
            if decoded.is_some() {
                return Err(Error::Invalid(
                    "opened-presentation geometry coordinate attribute is duplicated".into(),
                ));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|_| malformed_geometry_xml())?
                .into_owned();
            let parsed = value.parse::<i64>().map_err(|_| {
                Error::Invalid("opened-presentation geometry coordinate is not an i64".into())
            })?;
            decoded = Some(parsed);
        }
    }
    let parsed = decoded.ok_or_else(|| {
        Error::Invalid("opened-presentation geometry coordinate is incomplete".into())
    })?;
    let relative = raw_attribute_value_span(raw, name)?;
    let absolute = tag_start
        .checked_add(relative.start)
        .and_then(|start| tag_start.checked_add(relative.end).map(|end| start..end))
        .ok_or_else(|| Error::Invalid("opened-presentation geometry span overflow".into()))?;
    Ok((parsed, absolute))
}

fn raw_attribute_value_span(element: &[u8], selected: &[u8]) -> Result<Range<usize>> {
    if element.first() != Some(&b'<') {
        return Err(Error::Invalid(
            "opened-presentation geometry start tag is malformed".into(),
        ));
    }
    let mut cursor = 1usize;
    while element
        .get(cursor)
        .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(*byte, b'/' | b'>'))
    {
        cursor += 1;
    }
    let mut found = None;
    loop {
        while element.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if cursor >= element.len() || matches!(element[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while element
            .get(cursor)
            .is_some_and(|byte| !byte.is_ascii_whitespace() && !matches!(*byte, b'=' | b'/' | b'>'))
        {
            cursor += 1;
        }
        let name_end = cursor;
        while element.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        if element.get(cursor) != Some(&b'=') {
            return Err(Error::Invalid(
                "opened-presentation geometry attribute has no value".into(),
            ));
        }
        cursor += 1;
        while element.get(cursor).is_some_and(u8::is_ascii_whitespace) {
            cursor += 1;
        }
        let quote = *element.get(cursor).ok_or_else(|| {
            Error::Invalid("opened-presentation geometry attribute value is missing".into())
        })?;
        if !matches!(quote, b'\'' | b'"') {
            return Err(Error::Invalid(
                "opened-presentation geometry attribute value is not quoted".into(),
            ));
        }
        cursor += 1;
        let value_start = cursor;
        while element.get(cursor).is_some_and(|byte| *byte != quote) {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= element.len() {
            return Err(Error::Invalid(
                "opened-presentation geometry attribute value is unterminated".into(),
            ));
        }
        cursor += 1;
        if element.get(name_start..name_end) == Some(selected) {
            if found.is_some() {
                return Err(Error::Invalid(
                    "opened-presentation geometry coordinate attribute is duplicated".into(),
                ));
            }
            found = Some(value_start..value_end);
        }
    }
    found.ok_or_else(|| {
        Error::Invalid("opened-presentation geometry coordinate span is missing".into())
    })
}

fn find_record<'a>(scanned: &'a [ShapeGeometry], span: &Range<usize>) -> Result<&'a ShapeGeometry> {
    let index = scanned
        .binary_search_by_key(&span.start, |record| record.start)
        .map_err(|_| {
            Error::Invalid("opened-presentation shape has no stable native XML span".into())
        })?;
    let record = &scanned[index];
    if record.end == Some(span.end) {
        Ok(record)
    } else {
        Err(Error::Invalid(
            "opened-presentation shape native XML span changed".into(),
        ))
    }
}

fn classify_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    stack: &[ElementFrame],
) -> ElementKind {
    let local = element.local_name();
    if is_pml(namespace) && matches!(local.as_ref(), b"sp" | b"graphicFrame" | b"pic") {
        ElementKind::Shape
    } else if is_pml(namespace)
        && local.as_ref() == b"xfrm"
        && stack
            .last()
            .is_some_and(|frame| frame.kind == ElementKind::Shape && frame.shape.is_some())
    {
        // A graphic frame's own transform (`p:graphicFrame/p:xfrm`).
        ElementKind::Transform
    } else if is_pml(namespace)
        && local.as_ref() == b"spPr"
        && stack
            .last()
            .is_some_and(|frame| frame.kind == ElementKind::Shape && frame.shape.is_some())
    {
        ElementKind::ShapeProperties
    } else if is_dml(namespace)
        && local.as_ref() == b"xfrm"
        && stack.last().is_some_and(|frame| {
            frame.kind == ElementKind::ShapeProperties && frame.shape.is_some()
        })
    {
        ElementKind::Transform
    } else if is_dml(namespace)
        && local.as_ref() == b"off"
        && stack
            .last()
            .is_some_and(|frame| frame.kind == ElementKind::Transform && frame.shape.is_some())
    {
        ElementKind::Offset
    } else if is_dml(namespace)
        && local.as_ref() == b"ext"
        && stack
            .last()
            .is_some_and(|frame| frame.kind == ElementKind::Transform && frame.shape.is_some())
    {
        ElementKind::Extent
    } else {
        ElementKind::Other
    }
}

fn is_pml(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == PML_NAMESPACE || *value == STRICT_PML_NAMESPACE)
}

fn is_dml(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == DML_NAMESPACE || *value == STRICT_DML_NAMESPACE)
}

fn mark_malformed(shapes: &mut [ShapeGeometry], index: Option<usize>) {
    if let Some(index) = index {
        if let Some(shape) = shapes.get_mut(index) {
            shape.malformed = true;
        }
    }
}

fn xml_position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| Error::Invalid("opened-presentation geometry XML offset overflow".into()))
}

fn malformed_geometry_xml() -> Error {
    Error::Xml("malformed opened-presentation geometry XML".into())
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "geometry codec tests use compact fixtures"
)]
mod tests {
    use super::*;

    const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
    const STRICT_DML: &str = "http://purl.oclc.org/ooxml/drawingml/main";

    fn fixture(pml: &str, dml: &str) -> Vec<u8> {
        format!(
            r#"<q:sld xmlns:q="{pml}" xmlns:d="{dml}"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/><q:sp><q:nvSpPr><q:cNvPr id="2" name="Box"/><q:cNvSpPr/><q:nvPr/></q:nvSpPr><q:spPr data-keep=" yes "><d:xfrm rot="900000"><d:off y="20" x="10"/><d:ext cy="40" cx="30"/></d:xfrm><d:foo data="keep"/><d:prstGeom prst="rect"><d:avLst/></d:prstGeom></q:spPr></q:sp></q:spTree></q:cSld></q:sld>"#
        )
        .into_bytes()
    }

    fn replace_bytes(source: Vec<u8>, from: &str, to: &str) -> Vec<u8> {
        String::from_utf8(source)
            .expect("fixture UTF-8")
            .replace(from, to)
            .into_bytes()
    }

    #[test]
    fn rewrites_aliases_and_retains_every_other_byte() {
        for (pml, dml) in [(PML, DML), (STRICT_PML, STRICT_DML)] {
            let source = fixture(pml, dml);
            let output = rewrite_shape_bounds(
                &source,
                Position::new(0),
                (-100, 200, 300, 400),
                source.len() + 100,
            )
            .expect("rewrite")
            .expect("changed");
            assert!(
                String::from_utf8_lossy(&output)
                    .contains(r#"<d:off y="200" x="-100"/><d:ext cy="400" cx="300"/>"#)
            );
            assert!(String::from_utf8_lossy(&output).contains(r#"rot="900000"#));
            assert!(String::from_utf8_lossy(&output).contains(r#"<d:foo data="keep"/>"#));
            assert_eq!(
                crate::shape::Scene::read(&output, crate::shape::Limits::DEFAULT)
                    .expect("scene")
                    .at(0)
                    .expect("shape")
                    .bounds()
                    .map(|b| (b.x(), b.y(), b.width(), b.height())),
                Some((-100, 200, 300, 400))
            );
        }
    }

    #[test]
    fn no_op_is_exact_and_output_limit_is_checked_before_allocation() {
        let source = fixture(PML, DML);
        assert!(
            rewrite_shape_bounds(&source, Position::new(0), (10, 20, 30, 40), 0)
                .expect("no-op")
                .is_none()
        );
        let error = rewrite_shape_bounds(
            &source,
            Position::new(0),
            (10, 20, 3000, 4000),
            source.len(),
        )
        .expect_err("limit");
        assert!(matches!(error, Error::Limit { .. }));
    }

    #[test]
    fn partial_changes_preserve_unchanged_numeric_lexemes() {
        let source = replace_bytes(fixture(PML, DML), "x=\"10\"", "x=\"+001\"");
        let output = rewrite_shape_bounds(&source, Position::new(0), (1, 21, 30, 40), 1_000_000)
            .expect("rewrite")
            .expect("changed");
        let output = String::from_utf8_lossy(&output);
        assert!(
            output.contains(r#"<d:off y="21" x="+001"/>"#),
            "rewritten XML: {output}"
        );
        assert!(output.contains(r#"<d:ext cy="40" cx="30"/>"#));
    }

    #[test]
    fn malformed_or_foreign_coordinate_owners_are_rejected() {
        let duplicate = replace_bytes(fixture(PML, DML), "x=\"10\"", "x=\"10\" x=\"11\"");
        assert!(
            rewrite_shape_bounds(&duplicate, Position::new(0), (1, 2, 3, 4), 1_000_000).is_err()
        );
        let foreign = replace_bytes(
            fixture(PML, DML),
            "<d:off y=\"20\" x=\"10\"/>",
            "<d:off y=\"20\" f:x=\"10\" xmlns:f=\"urn:foreign\"/>",
        );
        assert!(rewrite_shape_bounds(&foreign, Position::new(0), (1, 2, 3, 4), 1_000_000).is_err());
    }

    #[test]
    fn typed_transform_rejection_is_atomic_and_unmasked_in_skeleton() {
        for malformed in [
            replace_bytes(fixture(PML, DML), "rot=\"900000\"", "rot=\"not-an-angle\""),
            replace_bytes(
                fixture(PML, DML),
                "</d:xfrm>",
                "<d:foreign data=\"keep\"/></d:xfrm>",
            ),
        ] {
            assert!(
                rewrite_shape_bounds(&malformed, Position::new(0), (1, 2, 3, 4), 1_000_000)
                    .is_err()
            );
            assert_eq!(
                geometry_skeleton(&malformed).expect("unmasked proof"),
                malformed
            );
        }
    }

    #[test]
    fn requested_bounds_use_native_coordinate_and_extent_domains() {
        let source = fixture(PML, DML);
        assert!(
            rewrite_shape_bounds(&source, Position::new(0), (i64::MAX, 20, 30, 40), 1_000_000,)
                .is_err()
        );
        assert!(
            rewrite_shape_bounds(
                &source,
                Position::new(0),
                (10, 20, litchi_drawingml::coordinate::MAX_EMU + 1, 40),
                1_000_000,
            )
            .is_err()
        );
    }

    #[test]
    fn skeleton_masks_only_validated_native_shape_geometry() {
        let source = fixture(PML, DML);
        let skeleton = geometry_skeleton(&source).expect("skeleton");
        assert!(String::from_utf8_lossy(&skeleton).contains(r#"<d:off y="0" x="0"/>"#));
        assert!(String::from_utf8_lossy(&skeleton).contains(r#"<d:ext cy="1" cx="1"/>"#));
        assert!(String::from_utf8_lossy(&skeleton).contains(r#"rot="900000"#));
        assert!(String::from_utf8_lossy(&skeleton).contains(r#"data-keep=" yes "#));
    }

    #[test]
    fn skeleton_accepts_zero_source_extents_and_skips_invalid_values() {
        let zero = replace_bytes(
            fixture(PML, DML),
            "cy=\"40\" cx=\"30\"",
            "cy=\"0\" cx=\"0\"",
        );
        let skeleton = geometry_skeleton(&zero).expect("zero extent proof");
        assert!(String::from_utf8_lossy(&skeleton).contains(r#"cy="1" cx="1"#));
        let negative = replace_bytes(
            fixture(PML, DML),
            "cy=\"40\" cx=\"30\"",
            "cy=\"-1\" cx=\"30\"",
        );
        assert_eq!(
            geometry_skeleton(&negative).expect("invalid proof"),
            negative
        );
    }

    #[test]
    fn nested_extension_geometry_is_not_treated_as_direct_owner() {
        let source = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:x="urn:foreign"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Box"/></p:nvSpPr><p:spPr/><p:extLst><x:future><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></x:future></p:extLst></p:sp></p:spTree></p:cSld></p:sld>"#
        );
        assert!(
            rewrite_shape_bounds(source.as_bytes(), Position::new(0), (5, 6, 7, 8), 1_000_000)
                .is_err()
        );
        let skeleton = geometry_skeleton(source.as_bytes()).expect("unchanged proof surface");
        assert_eq!(skeleton, source.as_bytes());
    }

    #[test]
    fn mce_rewritten_owners_are_never_edited_or_normalized() {
        let source = format!(
            r#"<p:sld xmlns:p="{PML}" xmlns:a="{DML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><mc:AlternateContent><mc:Choice Requires="x"><p:sp><p:nvSpPr><p:cNvPr id="2" name="Box"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:sp></mc:Choice><mc:Fallback><p:sp/></mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#
        );
        assert!(
            rewrite_shape_bounds(source.as_bytes(), Position::new(0), (5, 6, 7, 8), 1_000_000)
                .is_err()
        );
        assert!(geometry_skeleton(source.as_bytes()).is_err());
    }
}
