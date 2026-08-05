//! Lossless codec for extension-backed native chart drawables.
//!
//! `TSCH.ChartArchive` is protobuf extension field 10000 of
//! `TSCH.ChartDrawableArchive`. `prost` does not expose proto2 extensions, so a
//! direct decode of the generated drawable type loses the chart itself.

use std::collections::{HashMap, HashSet};

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::parse_wire_fields_with_limits;
use prost::Message;

use crate::protobuf::{tsch, tsp};
use crate::wire::{
    WireField, append_length_delimited_field, append_varint_field, parse_wire_fields,
};
use crate::{Error, Result};

const DRAWABLE_SUPER_FIELD: u32 = 1;
const CHART_EXTENSION_FIELD: u32 = 10_000;
const CHART_REFERENCE_LINES_EXTENSION_FIELD: u32 = 10_005;
const CHART_BASE_FIELDS: std::ops::RangeInclusive<u32> = 1..=24;
const MAX_REFERENCE_LINE_GRAPH_AXES: usize = 32;
const MAX_REFERENCE_LINE_GRAPH_ENTRIES: usize = 128;
const MAX_REFERENCE_LINE_GRAPH_FIELDS: usize = 256;

/// A native chart drawable with its extension-backed chart payload retained.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct IWorkChartArchive {
    pub drawable: tsch::ChartDrawableArchive,
    pub chart: Option<tsch::ChartArchive>,
    chart_opaque_fields: Vec<Vec<u8>>,
    opaque_fields: Vec<Vec<u8>>,
}

impl IWorkChartArchive {
    /// Construct a chart drawable from typed protobuf values.
    pub fn new(drawable: tsch::ChartDrawableArchive, chart: tsch::ChartArchive) -> Self {
        Self {
            drawable,
            chart: Some(chart),
            chart_opaque_fields: Vec::new(),
            opaque_fields: Vec::new(),
        }
    }

    pub(crate) fn append_chart_bool_extension(
        &mut self,
        field_number: u32,
        value: bool,
    ) -> Result<()> {
        let mut field = Vec::new();
        append_varint_field(&mut field, field_number, u64::from(value))?;
        self.chart_opaque_fields.push(field);
        Ok(())
    }

    pub(crate) fn append_chart_message_extension(
        &mut self,
        field_number: u32,
        message: &impl Message,
    ) -> Result<()> {
        let mut field = Vec::new();
        append_length_delimited_field(&mut field, field_number, &message.encode_to_vec())?;
        self.chart_opaque_fields.push(field);
        Ok(())
    }

    /// Decode the chart's native reference-line graph, if one is present.
    pub fn reference_lines(&self) -> Result<Option<tsch::ChartReferenceLinesArchive>> {
        reference_lines_from_opaque_fields(&self.chart_opaque_fields)
    }

    pub(crate) fn set_reference_lines(
        &mut self,
        reference_lines: Option<&tsch::ChartReferenceLinesArchive>,
    ) -> Result<()> {
        let existing_index = unique_opaque_field_index(
            &self.chart_opaque_fields,
            CHART_REFERENCE_LINES_EXTENSION_FIELD,
        )?;
        let mut next_chart_opaque_fields = self.chart_opaque_fields.clone();
        match (existing_index, reference_lines) {
            (Some(index), Some(reference_lines)) => {
                let current = next_chart_opaque_fields[index].as_slice();
                let current_fields = parse_wire_fields(current)?;
                let [current_field] = current_fields.as_slice() else {
                    return Err(Error::InvalidFormat(
                        "stored chart extension is not exactly one wire field".to_owned(),
                    ));
                };
                let current_payload = reference_line_extension_payload(current, current_field)?;
                preflight_reference_line_graph(current_payload)?;
                validate_reference_line_graph_wire(current_payload)?;
                let next_payload =
                    preserve_reference_line_graph_wire(current_payload, reference_lines)?;
                preflight_reference_line_graph(&next_payload)?;
                let next = encoded_chart_extension_payload(
                    CHART_REFERENCE_LINES_EXTENSION_FIELD,
                    &next_payload,
                )?;
                let decoded = tsch::ChartReferenceLinesArchive::decode(next_payload.as_slice())?;
                if decoded != *reference_lines {
                    return Err(Error::InvalidFormat(
                        "chart reference-line extension update failed validation".to_owned(),
                    ));
                }
                next_chart_opaque_fields[index] = next;
            },
            (Some(index), None) => {
                let current = next_chart_opaque_fields[index].as_slice();
                let current_fields = parse_wire_fields(current)?;
                let [current_field] = current_fields.as_slice() else {
                    return Err(Error::InvalidFormat(
                        "stored chart extension is not exactly one wire field".to_owned(),
                    ));
                };
                let current_payload = reference_line_extension_payload(current, current_field)?;
                preflight_reference_line_graph(current_payload)?;
                validate_reference_line_graph_wire(current_payload)?;
                tsch::ChartReferenceLinesArchive::decode(current_payload)?;
                next_chart_opaque_fields.remove(index);
            },
            (None, Some(reference_lines)) => {
                let payload = reference_lines.encode_to_vec();
                preflight_reference_line_graph(&payload)?;
                validate_reference_line_graph_wire(&payload)?;
                let decoded = tsch::ChartReferenceLinesArchive::decode(payload.as_slice())?;
                if decoded != *reference_lines {
                    return Err(Error::InvalidFormat(
                        "chart reference-line extension update failed validation".to_owned(),
                    ));
                }
                let encoded = encoded_chart_extension_payload(
                    CHART_REFERENCE_LINES_EXTENSION_FIELD,
                    &payload,
                )?;
                let index = next_chart_opaque_fields
                    .iter()
                    .position(|field| {
                        parse_wire_fields(field)
                            .ok()
                            .and_then(|fields| fields.first().map(|field| field.number()))
                            .is_some_and(|number| number > CHART_REFERENCE_LINES_EXTENSION_FIELD)
                    })
                    .unwrap_or(next_chart_opaque_fields.len());
                next_chart_opaque_fields.insert(index, encoded);
            },
            (None, None) => {},
        }
        if reference_lines_from_opaque_fields(&next_chart_opaque_fields)?
            != reference_lines.cloned()
        {
            return Err(Error::InvalidFormat(
                "chart reference-line extension update failed validation".to_owned(),
            ));
        }
        self.chart_opaque_fields = next_chart_opaque_fields;
        Ok(())
    }

    /// Remap the chart's typed object references while retaining opaque fields.
    ///
    /// iWork records object references in both the protobuf payload and IWA
    /// message metadata. An opaque chart extension could itself contain an
    /// object reference that this codec cannot safely rewrite, so cloning is
    /// rejected when metadata identifies a private reference outside the typed
    /// fields handled here.
    pub(crate) fn remap_references(
        &mut self,
        remap: &HashMap<u64, u64>,
        recorded_references: &[u64],
    ) -> Result<()> {
        let typed_references = self.typed_reference_identifiers()?;
        if let Some(identifier) = recorded_references.iter().copied().find(|identifier| {
            remap.contains_key(identifier) && !typed_references.contains(identifier)
        }) {
            return Err(Error::InvalidFormat(format!(
                "chart payload has an unrecognized private reference {identifier}"
            )));
        }

        if let Some(drawable) = self.drawable.super_.as_mut() {
            remap_optional_reference(&mut drawable.parent, remap);
            remap_optional_reference(&mut drawable.comment, remap);
            for reference in &mut drawable.pencil_annotations {
                remap_reference(reference, remap);
            }
            remap_optional_reference(&mut drawable.title, remap);
            remap_optional_reference(&mut drawable.caption, remap);
        }
        if let Some(chart) = self.chart.as_mut() {
            remap_optional_reference(&mut chart.preset, remap);
            remap_optional_reference(&mut chart.mediator, remap);
            remap_optional_reference(&mut chart.chart_style, remap);
            remap_optional_reference(&mut chart.chart_non_style, remap);
            remap_optional_reference(&mut chart.legend_style, remap);
            remap_optional_reference(&mut chart.legend_non_style, remap);
            remap_references(&mut chart.value_axis_styles, remap);
            remap_references(&mut chart.value_axis_nonstyles, remap);
            remap_references(&mut chart.category_axis_styles, remap);
            remap_references(&mut chart.category_axis_nonstyles, remap);
            remap_references(&mut chart.series_theme_styles, remap);
            remap_sparse_references(chart.series_private_styles.as_mut(), remap);
            remap_sparse_references(chart.series_non_styles.as_mut(), remap);
            remap_references(&mut chart.paragraph_styles, remap);
            remap_optional_reference(&mut chart.owned_preset, remap);
        }
        if let Some(mut reference_lines) = self.reference_lines()? {
            remap_chart_reference_lines(&mut reference_lines, remap);
            self.set_reference_lines(Some(&reference_lines))?;
        }
        Ok(())
    }

    pub(crate) fn typed_reference_identifiers(&self) -> Result<HashSet<u64>> {
        let mut identifiers = HashSet::new();
        if let Some(drawable) = self.drawable.super_.as_ref() {
            collect_optional_reference(&mut identifiers, drawable.parent.as_ref());
            collect_optional_reference(&mut identifiers, drawable.comment.as_ref());
            collect_references(&mut identifiers, &drawable.pencil_annotations);
            collect_optional_reference(&mut identifiers, drawable.title.as_ref());
            collect_optional_reference(&mut identifiers, drawable.caption.as_ref());
        }
        if let Some(chart) = self.chart.as_ref() {
            collect_optional_reference(&mut identifiers, chart.preset.as_ref());
            collect_optional_reference(&mut identifiers, chart.mediator.as_ref());
            collect_optional_reference(&mut identifiers, chart.chart_style.as_ref());
            collect_optional_reference(&mut identifiers, chart.chart_non_style.as_ref());
            collect_optional_reference(&mut identifiers, chart.legend_style.as_ref());
            collect_optional_reference(&mut identifiers, chart.legend_non_style.as_ref());
            collect_references(&mut identifiers, &chart.value_axis_styles);
            collect_references(&mut identifiers, &chart.value_axis_nonstyles);
            collect_references(&mut identifiers, &chart.category_axis_styles);
            collect_references(&mut identifiers, &chart.category_axis_nonstyles);
            collect_references(&mut identifiers, &chart.series_theme_styles);
            collect_sparse_references(&mut identifiers, chart.series_private_styles.as_ref());
            collect_sparse_references(&mut identifiers, chart.series_non_styles.as_ref());
            collect_references(&mut identifiers, &chart.paragraph_styles);
            collect_optional_reference(&mut identifiers, chart.owned_preset.as_ref());
        }
        if let Some(reference_lines) = self.reference_lines()? {
            collect_chart_reference_lines(&mut identifiers, &reference_lines);
        }
        Ok(identifiers)
    }

    /// Decode a chart drawable without discarding extensions or future fields.
    pub fn decode(data: &[u8]) -> Result<Self> {
        let fields = parse_wire_fields(data)?;
        let chart_field = unique_field(&fields, CHART_EXTENSION_FIELD)?;
        let chart_data = chart_field
            .map(|field| length_delimited_payload(data, field))
            .transpose()?;
        let chart = chart_data.map(tsch::ChartArchive::decode).transpose()?;
        let chart_opaque_fields = if let Some(chart_data) = chart_data {
            parse_wire_fields(chart_data)?
                .into_iter()
                .filter(|field| !CHART_BASE_FIELDS.contains(&field.number()))
                .map(|field| chart_data[field.start()..field.end()].to_vec())
                .collect()
        } else {
            Vec::new()
        };
        let opaque_fields = fields
            .iter()
            .filter(|field| {
                field.number() != DRAWABLE_SUPER_FIELD && field.number() != CHART_EXTENSION_FIELD
            })
            .map(|field| data[field.start()..field.end()].to_vec())
            .collect();

        Ok(Self {
            drawable: tsch::ChartDrawableArchive::decode(data)?,
            chart,
            chart_opaque_fields,
            opaque_fields,
        })
    }

    /// Encode the drawable, chart extension, and untouched future fields.
    pub fn encode(&self) -> Result<Vec<u8>> {
        let mut output = self.drawable.encode_to_vec();
        if let Some(chart) = &self.chart {
            let mut chart_data = chart.encode_to_vec();
            for field in &self.chart_opaque_fields {
                chart_data.extend_from_slice(field);
            }
            append_length_delimited_field(&mut output, CHART_EXTENSION_FIELD, &chart_data)?;
        }
        for field in &self.opaque_fields {
            output.extend_from_slice(field);
        }
        Ok(output)
    }
}

fn encoded_chart_extension_payload(field_number: u32, payload: &[u8]) -> Result<Vec<u8>> {
    let mut field = Vec::new();
    append_length_delimited_field(&mut field, field_number, payload)?;
    Ok(field)
}

fn reference_lines_from_opaque_fields(
    chart_opaque_fields: &[Vec<u8>],
) -> Result<Option<tsch::ChartReferenceLinesArchive>> {
    let Some(encoded) =
        unique_opaque_field(chart_opaque_fields, CHART_REFERENCE_LINES_EXTENSION_FIELD)?
    else {
        return Ok(None);
    };
    let fields = parse_wire_fields(encoded)?;
    let field = unique_field(&fields, CHART_REFERENCE_LINES_EXTENSION_FIELD)?.ok_or_else(|| {
        Error::InvalidFormat(
            "chart reference-line extension disappeared during decoding".to_owned(),
        )
    })?;
    let payload = reference_line_extension_payload(encoded, field)?;
    preflight_reference_line_graph(payload)?;
    validate_reference_line_graph_wire(payload)?;
    Ok(Some(tsch::ChartReferenceLinesArchive::decode(payload)?))
}

fn reference_line_extension_payload<'a>(data: &'a [u8], field: &WireField) -> Result<&'a [u8]> {
    if field.wire_type() != 2 {
        return Err(Error::InvalidFormat(
            "chart reference-line extension is not length-delimited".to_owned(),
        ));
    }
    field.validate_canonical_key(data)?;
    field.validate_canonical_length(data)?;
    Ok(field.checked_payload(data)?)
}

#[derive(Clone, Copy)]
enum ReferenceLineWireFieldKind {
    Varint,
    Message(fn(&[u8], &[u8]) -> Result<Vec<u8>>),
    RepeatedMessage(fn(&[u8], &[u8]) -> Result<Vec<u8>>),
}

#[derive(Clone, Copy)]
struct ReferenceLineWireFieldSpec {
    number: u32,
    kind: ReferenceLineWireFieldKind,
    nested: Option<fn(&[u8]) -> Result<()>>,
}

fn reference_line_graph_schema() -> [ReferenceLineWireFieldSpec; 3] {
    [
        ReferenceLineWireFieldSpec {
            number: 1,
            kind: ReferenceLineWireFieldKind::RepeatedMessage(
                preserve_reference_line_non_styles_axis_wire,
            ),
            nested: Some(validate_reference_line_non_styles_axis_wire),
        },
        ReferenceLineWireFieldSpec {
            number: 2,
            kind: ReferenceLineWireFieldKind::RepeatedMessage(
                preserve_reference_line_styles_axis_wire,
            ),
            nested: Some(validate_reference_line_styles_axis_wire),
        },
        ReferenceLineWireFieldSpec {
            number: 3,
            kind: ReferenceLineWireFieldKind::Message(preserve_reference_wire),
            nested: Some(validate_reference_wire),
        },
    ]
}

fn reference_line_non_styles_axis_schema() -> [ReferenceLineWireFieldSpec; 2] {
    [
        ReferenceLineWireFieldSpec {
            number: 1,
            kind: ReferenceLineWireFieldKind::Message(preserve_chart_axis_id_wire),
            nested: Some(validate_chart_axis_id_wire),
        },
        ReferenceLineWireFieldSpec {
            number: 2,
            kind: ReferenceLineWireFieldKind::RepeatedMessage(preserve_reference_line_item_wire),
            nested: Some(validate_reference_line_item_wire),
        },
    ]
}

fn reference_line_styles_axis_schema() -> [ReferenceLineWireFieldSpec; 2] {
    [
        ReferenceLineWireFieldSpec {
            number: 1,
            kind: ReferenceLineWireFieldKind::Message(preserve_chart_axis_id_wire),
            nested: Some(validate_chart_axis_id_wire),
        },
        ReferenceLineWireFieldSpec {
            number: 2,
            kind: ReferenceLineWireFieldKind::Message(preserve_sparse_reference_array_wire),
            nested: Some(validate_sparse_reference_array_wire),
        },
    ]
}

fn reference_line_item_schema() -> [ReferenceLineWireFieldSpec; 2] {
    [
        ReferenceLineWireFieldSpec {
            number: 1,
            kind: ReferenceLineWireFieldKind::Message(preserve_reference_wire),
            nested: Some(validate_reference_wire),
        },
        ReferenceLineWireFieldSpec {
            number: 2,
            kind: ReferenceLineWireFieldKind::Message(preserve_uuid_wire),
            nested: Some(validate_uuid_wire),
        },
    ]
}

fn sparse_reference_array_schema() -> [ReferenceLineWireFieldSpec; 2] {
    [
        ReferenceLineWireFieldSpec {
            number: 1,
            kind: ReferenceLineWireFieldKind::Varint,
            nested: None,
        },
        ReferenceLineWireFieldSpec {
            number: 2,
            kind: ReferenceLineWireFieldKind::RepeatedMessage(preserve_sparse_reference_entry_wire),
            nested: Some(validate_sparse_reference_entry_wire),
        },
    ]
}

fn sparse_reference_entry_schema() -> [ReferenceLineWireFieldSpec; 2] {
    [
        ReferenceLineWireFieldSpec {
            number: 1,
            kind: ReferenceLineWireFieldKind::Varint,
            nested: None,
        },
        ReferenceLineWireFieldSpec {
            number: 2,
            kind: ReferenceLineWireFieldKind::Message(preserve_reference_wire),
            nested: Some(validate_reference_wire),
        },
    ]
}

fn reference_schema() -> [ReferenceLineWireFieldSpec; 3] {
    [
        ReferenceLineWireFieldSpec {
            number: 1,
            kind: ReferenceLineWireFieldKind::Varint,
            nested: None,
        },
        ReferenceLineWireFieldSpec {
            number: 2,
            kind: ReferenceLineWireFieldKind::Varint,
            nested: None,
        },
        ReferenceLineWireFieldSpec {
            number: 3,
            kind: ReferenceLineWireFieldKind::Varint,
            nested: None,
        },
    ]
}

fn uuid_schema() -> [ReferenceLineWireFieldSpec; 2] {
    [
        ReferenceLineWireFieldSpec {
            number: 1,
            kind: ReferenceLineWireFieldKind::Varint,
            nested: None,
        },
        ReferenceLineWireFieldSpec {
            number: 2,
            kind: ReferenceLineWireFieldKind::Varint,
            nested: None,
        },
    ]
}

fn chart_axis_id_schema() -> [ReferenceLineWireFieldSpec; 2] {
    [
        ReferenceLineWireFieldSpec {
            number: 1,
            kind: ReferenceLineWireFieldKind::Varint,
            nested: None,
        },
        ReferenceLineWireFieldSpec {
            number: 2,
            kind: ReferenceLineWireFieldKind::Varint,
            nested: None,
        },
    ]
}

fn preserve_reference_line_graph_wire(
    existing: &[u8],
    replacement: &tsch::ChartReferenceLinesArchive,
) -> Result<Vec<u8>> {
    let replacement_data = replacement.encode_to_vec();
    preflight_reference_line_graph(&replacement_data)?;
    let schema = reference_line_graph_schema();
    merge_reference_line_wire_message(existing, &replacement_data, &schema)
}

fn preserve_reference_line_non_styles_axis_wire(
    existing: &[u8],
    replacement: &[u8],
) -> Result<Vec<u8>> {
    let schema = reference_line_non_styles_axis_schema();
    merge_reference_line_wire_message(existing, replacement, &schema)
}

fn preserve_reference_line_styles_axis_wire(
    existing: &[u8],
    replacement: &[u8],
) -> Result<Vec<u8>> {
    let schema = reference_line_styles_axis_schema();
    merge_reference_line_wire_message(existing, replacement, &schema)
}

fn preserve_reference_line_item_wire(existing: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
    let schema = reference_line_item_schema();
    merge_reference_line_wire_message(existing, replacement, &schema)
}

fn preserve_sparse_reference_array_wire(existing: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
    let schema = sparse_reference_array_schema();
    merge_reference_line_wire_message(existing, replacement, &schema)
}

fn preserve_sparse_reference_entry_wire(existing: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
    let schema = sparse_reference_entry_schema();
    merge_reference_line_wire_message(existing, replacement, &schema)
}

fn preserve_reference_wire(existing: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
    let schema = reference_schema();
    merge_reference_line_wire_message(existing, replacement, &schema)
}

fn preserve_uuid_wire(existing: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
    let schema = uuid_schema();
    merge_reference_line_wire_message(existing, replacement, &schema)
}

fn preserve_chart_axis_id_wire(existing: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
    let schema = chart_axis_id_schema();
    merge_reference_line_wire_message(existing, replacement, &schema)
}

fn validate_reference_line_graph_wire(data: &[u8]) -> Result<()> {
    let schema = reference_line_graph_schema();
    validate_reference_line_wire_message(data, &schema)
}

fn validate_reference_line_non_styles_axis_wire(data: &[u8]) -> Result<()> {
    let schema = reference_line_non_styles_axis_schema();
    validate_reference_line_wire_message(data, &schema)
}

fn validate_reference_line_styles_axis_wire(data: &[u8]) -> Result<()> {
    let schema = reference_line_styles_axis_schema();
    validate_reference_line_wire_message(data, &schema)
}

fn validate_reference_line_item_wire(data: &[u8]) -> Result<()> {
    let schema = reference_line_item_schema();
    validate_reference_line_wire_message(data, &schema)
}

fn validate_sparse_reference_array_wire(data: &[u8]) -> Result<()> {
    let schema = sparse_reference_array_schema();
    validate_reference_line_wire_message(data, &schema)
}

fn validate_sparse_reference_entry_wire(data: &[u8]) -> Result<()> {
    let schema = sparse_reference_entry_schema();
    validate_reference_line_wire_message(data, &schema)
}

fn validate_reference_wire(data: &[u8]) -> Result<()> {
    let schema = reference_schema();
    validate_reference_line_wire_message(data, &schema)
}

fn validate_uuid_wire(data: &[u8]) -> Result<()> {
    let schema = uuid_schema();
    validate_reference_line_wire_message(data, &schema)
}

fn validate_chart_axis_id_wire(data: &[u8]) -> Result<()> {
    let schema = chart_axis_id_schema();
    validate_reference_line_wire_message(data, &schema)
}

fn validate_reference_line_wire_message(
    data: &[u8],
    schema: &[ReferenceLineWireFieldSpec],
) -> Result<()> {
    let limits = WireLimits::default().with_fields(MAX_REFERENCE_LINE_GRAPH_FIELDS)?;
    let fields = parse_wire_fields_with_limits(data, limits)?;
    validate_reference_line_wire_fields(data, &fields, schema)?;
    validate_reference_line_wire_nested_fields(data, &fields, schema)
}

fn validate_reference_line_wire_nested_fields(
    data: &[u8],
    fields: &[WireField],
    schema: &[ReferenceLineWireFieldSpec],
) -> Result<()> {
    for field in fields {
        let Some(spec) = schema.iter().find(|spec| spec.number == field.number()) else {
            continue;
        };
        let Some(validate) = spec.nested else {
            continue;
        };
        validate(field.checked_payload(data)?)?;
    }
    Ok(())
}

fn merge_reference_line_wire_message(
    existing: &[u8],
    replacement: &[u8],
    schema: &[ReferenceLineWireFieldSpec],
) -> Result<Vec<u8>> {
    let limits = WireLimits::default().with_fields(MAX_REFERENCE_LINE_GRAPH_FIELDS)?;
    let existing_fields = parse_wire_fields_with_limits(existing, limits)?;
    let replacement_fields = parse_wire_fields_with_limits(replacement, limits)?;
    validate_reference_line_wire_fields(existing, &existing_fields, schema)?;
    validate_reference_line_wire_fields(replacement, &replacement_fields, schema)?;
    validate_reference_line_wire_nested_fields(existing, &existing_fields, schema)?;
    validate_reference_line_wire_nested_fields(replacement, &replacement_fields, schema)?;

    let mut output = Vec::new();
    for (index, field) in existing_fields.iter().enumerate() {
        let Some(spec) = schema.iter().find(|spec| spec.number == field.number()) else {
            append_reference_line_wire_bytes(
                &mut output,
                &existing[field.start()..field.end()],
                limits,
            )?;
            continue;
        };
        let occurrence = existing_fields[..index]
            .iter()
            .filter(|candidate| candidate.number() == field.number())
            .count();
        let replacement_field = replacement_fields
            .iter()
            .filter(|candidate| candidate.number() == field.number())
            .nth(occurrence);
        let Some(replacement_field) = replacement_field else {
            continue;
        };
        append_reference_line_wire_replacement(
            &mut output,
            existing,
            field,
            replacement,
            replacement_field,
            spec.kind,
            limits,
        )?;
    }

    for (index, field) in replacement_fields.iter().enumerate() {
        let occurrence = replacement_fields[..index]
            .iter()
            .filter(|candidate| candidate.number() == field.number())
            .count();
        let existing_count = existing_fields
            .iter()
            .filter(|candidate| candidate.number() == field.number())
            .count();
        if occurrence >= existing_count {
            append_reference_line_wire_bytes(
                &mut output,
                &replacement[field.start()..field.end()],
                limits,
            )?;
        }
    }
    Ok(output)
}

fn validate_reference_line_wire_fields(
    data: &[u8],
    fields: &[WireField],
    schema: &[ReferenceLineWireFieldSpec],
) -> Result<()> {
    for (index, field) in fields.iter().enumerate() {
        let Some(spec) = schema.iter().find(|spec| spec.number == field.number()) else {
            continue;
        };
        let repeated = matches!(spec.kind, ReferenceLineWireFieldKind::RepeatedMessage(_));
        if !repeated
            && fields[..index]
                .iter()
                .any(|candidate| candidate.number() == field.number())
        {
            return Err(Error::InvalidFormat(format!(
                "reference-line wire field {} occurs more than once",
                field.number()
            )));
        }
        let expected_wire_type = match spec.kind {
            ReferenceLineWireFieldKind::Varint => 0,
            ReferenceLineWireFieldKind::Message(_)
            | ReferenceLineWireFieldKind::RepeatedMessage(_) => 2,
        };
        if field.wire_type() != expected_wire_type {
            return Err(Error::InvalidFormat(format!(
                "reference-line wire field {} has the wrong wire type",
                field.number()
            )));
        }
        field.validate_canonical_key(data)?;
        if expected_wire_type == 2 {
            field.validate_canonical_length(data)?;
        } else {
            let payload = field.checked_payload(data)?;
            let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(payload)
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid reference-line varint: {error}"))
                })?;
            if length != payload.len() || length != litchi_iwa_common::varint::encoded_len(value) {
                return Err(Error::InvalidFormat(format!(
                    "reference-line wire field {} has a noncanonical value",
                    field.number()
                )));
            }
        }
    }
    Ok(())
}

fn append_reference_line_wire_replacement(
    output: &mut Vec<u8>,
    existing: &[u8],
    existing_field: &WireField,
    replacement: &[u8],
    replacement_field: &WireField,
    kind: ReferenceLineWireFieldKind,
    limits: WireLimits,
) -> Result<()> {
    match kind {
        ReferenceLineWireFieldKind::Varint => append_reference_line_wire_bytes(
            output,
            &replacement[replacement_field.start()..replacement_field.end()],
            limits,
        ),
        ReferenceLineWireFieldKind::Message(merge)
        | ReferenceLineWireFieldKind::RepeatedMessage(merge) => {
            let existing_payload = existing_field.checked_payload(existing)?;
            let replacement_payload = replacement_field.checked_payload(replacement)?;
            let merged = merge(existing_payload, replacement_payload)?;
            append_reference_line_length_delimited_field(
                output,
                replacement,
                replacement_field,
                &merged,
                limits,
            )
        },
    }
}

fn append_reference_line_length_delimited_field(
    output: &mut Vec<u8>,
    source: &[u8],
    field: &WireField,
    payload: &[u8],
    limits: WireLimits,
) -> Result<()> {
    append_reference_line_wire_bytes(output, &source[field.start()..field.key_end()], limits)?;
    let mut length = [0_u8; litchi_iwa_common::varint::MAX_BYTES];
    let encoded_length = litchi_iwa_common::varint::encode_varint_to_buffer(
        u64::try_from(payload.len())
            .map_err(|_| Error::InvalidFormat("reference-line payload exceeds u64".to_owned()))?,
        &mut length,
    );
    append_reference_line_wire_bytes(output, encoded_length, limits)?;
    append_reference_line_wire_bytes(output, payload, limits)
}

fn append_reference_line_wire_bytes(
    output: &mut Vec<u8>,
    bytes: &[u8],
    limits: WireLimits,
) -> Result<()> {
    let next_len = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| Error::InvalidFormat("reference-line wire output overflow".to_owned()))?;
    if next_len > limits.max_output_bytes() {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::OutputBytes,
            observed: next_len,
            limit: limits.max_output_bytes(),
        }));
    }
    output.try_reserve(bytes.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "reference-line wire output",
            amount: next_len,
        })
    })?;
    output.extend_from_slice(bytes);
    Ok(())
}

/// Bound repeated reference-line graph nodes before Prost materializes them.
///
/// The generated graph type owns every repeated message, so checking the
/// semantic five-line limit after decoding is too late for hostile archives.
/// This shallow preflight keeps the generated decode behind finite axis,
/// field, and entry budgets while retaining unknown fields for the normal
/// chart extension path.
fn preflight_reference_line_graph(data: &[u8]) -> Result<()> {
    let limits = WireLimits::default().with_fields(MAX_REFERENCE_LINE_GRAPH_FIELDS)?;
    let fields = parse_wire_fields_with_limits(data, limits)?;
    let mut axes = 0usize;
    let mut entries = 0usize;
    for field in fields {
        if !matches!(field.number(), 1 | 2) {
            continue;
        }
        axes = axes
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("reference-line axis count overflow".to_owned()))?;
        if axes > MAX_REFERENCE_LINE_GRAPH_AXES {
            return Err(Error::InvalidFormat(format!(
                "chart reference-line graph has more than {MAX_REFERENCE_LINE_GRAPH_AXES} axes"
            )));
        }
        let axis_payload = length_delimited_payload(data, &field)?;
        let axis_fields = parse_wire_fields_with_limits(axis_payload, limits)?;
        for axis_field in axis_fields {
            if axis_field.number() != 2 {
                continue;
            }
            if field.number() == 1 {
                entries = entries.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("reference-line entry count overflow".to_owned())
                })?;
            } else {
                let sparse_payload = length_delimited_payload(axis_payload, &axis_field)?;
                let sparse_fields = parse_wire_fields_with_limits(sparse_payload, limits)?;
                entries = entries
                    .checked_add(
                        sparse_fields
                            .iter()
                            .filter(|sparse_field| sparse_field.number() == 2)
                            .count(),
                    )
                    .ok_or_else(|| {
                        Error::InvalidFormat("reference-line entry count overflow".to_owned())
                    })?;
            }
            if entries > MAX_REFERENCE_LINE_GRAPH_ENTRIES {
                return Err(Error::InvalidFormat(format!(
                    "chart reference-line graph has more than {MAX_REFERENCE_LINE_GRAPH_ENTRIES} entries"
                )));
            }
        }
    }
    Ok(())
}

fn unique_opaque_field(fields: &[Vec<u8>], field_number: u32) -> Result<Option<&[u8]>> {
    unique_opaque_field_index(fields, field_number)
        .map(|index| index.map(|index| fields[index].as_slice()))
}

fn unique_opaque_field_index(fields: &[Vec<u8>], field_number: u32) -> Result<Option<usize>> {
    let mut result = None;
    for (index, encoded) in fields.iter().enumerate() {
        let parsed = parse_wire_fields(encoded)?;
        let [field] = parsed.as_slice() else {
            return Err(Error::InvalidFormat(
                "stored chart extension is not exactly one wire field".to_owned(),
            ));
        };
        if field.number() != field_number {
            continue;
        }
        if result.replace(index).is_some() {
            return Err(Error::InvalidFormat(format!(
                "chart extension field {field_number} occurs more than once"
            )));
        }
    }
    Ok(result)
}

fn remap_chart_reference_lines(
    reference_lines: &mut tsch::ChartReferenceLinesArchive,
    remap: &HashMap<u64, u64>,
) {
    let mut uuids = reference_lines
        .reference_line_non_styles_map
        .iter()
        .flat_map(|axis| &axis.reference_line_non_style_items)
        .map(|item| (item.uuid.lower, item.uuid.upper))
        .collect::<HashSet<_>>();
    for axis in &mut reference_lines.reference_line_non_styles_map {
        for item in &mut axis.reference_line_non_style_items {
            if remap.contains_key(&item.non_style.identifier) {
                uuids.remove(&(item.uuid.lower, item.uuid.upper));
                item.uuid = fresh_chart_reference_uuid(&mut uuids);
            }
            remap_reference(&mut item.non_style, remap);
        }
    }
    for axis in &mut reference_lines.reference_line_styles_map {
        if let Some(styles) = axis.reference_line_styles.as_mut() {
            for entry in &mut styles.entries {
                remap_reference(&mut entry.reference, remap);
            }
        }
    }
    remap_optional_reference(
        &mut reference_lines.theme_preset_reference_line_style,
        remap,
    );
}

fn fresh_chart_reference_uuid(existing: &mut HashSet<(u64, u64)>) -> tsp::Uuid {
    loop {
        let bytes = litchi_core::id::generate_guid_bytes();
        let mut lower = [0; 8];
        lower.copy_from_slice(&bytes[..8]);
        let mut upper = [0; 8];
        upper.copy_from_slice(&bytes[8..]);
        let uuid = tsp::Uuid {
            lower: u64::from_le_bytes(lower),
            upper: u64::from_le_bytes(upper),
        };
        if existing.insert((uuid.lower, uuid.upper)) {
            return uuid;
        }
    }
}

fn collect_chart_reference_lines(
    identifiers: &mut HashSet<u64>,
    reference_lines: &tsch::ChartReferenceLinesArchive,
) {
    for axis in &reference_lines.reference_line_non_styles_map {
        for item in &axis.reference_line_non_style_items {
            collect_reference(identifiers, &item.non_style);
        }
    }
    for axis in &reference_lines.reference_line_styles_map {
        if let Some(styles) = axis.reference_line_styles.as_ref() {
            for entry in &styles.entries {
                collect_reference(identifiers, &entry.reference);
            }
        }
    }
    collect_optional_reference(
        identifiers,
        reference_lines.theme_preset_reference_line_style.as_ref(),
    );
}

fn remap_reference(reference: &mut tsp::Reference, remap: &HashMap<u64, u64>) {
    if let Some(identifier) = remap.get(&reference.identifier) {
        reference.identifier = *identifier;
    }
}

fn remap_optional_reference(reference: &mut Option<tsp::Reference>, remap: &HashMap<u64, u64>) {
    if let Some(reference) = reference {
        remap_reference(reference, remap);
    }
}

fn remap_references(references: &mut [tsp::Reference], remap: &HashMap<u64, u64>) {
    for reference in references {
        remap_reference(reference, remap);
    }
}

fn remap_sparse_references(
    references: Option<&mut tsp::SparseReferenceArray>,
    remap: &HashMap<u64, u64>,
) {
    if let Some(references) = references {
        for entry in &mut references.entries {
            remap_reference(&mut entry.reference, remap);
        }
    }
}

fn collect_reference(identifiers: &mut HashSet<u64>, reference: &tsp::Reference) {
    if reference.identifier != 0 {
        identifiers.insert(reference.identifier);
    }
}

fn collect_optional_reference(identifiers: &mut HashSet<u64>, reference: Option<&tsp::Reference>) {
    if let Some(reference) = reference {
        collect_reference(identifiers, reference);
    }
}

fn collect_references(identifiers: &mut HashSet<u64>, references: &[tsp::Reference]) {
    for reference in references {
        collect_reference(identifiers, reference);
    }
}

fn collect_sparse_references(
    identifiers: &mut HashSet<u64>,
    references: Option<&tsp::SparseReferenceArray>,
) {
    if let Some(references) = references {
        for entry in &references.entries {
            collect_reference(identifiers, &entry.reference);
        }
    }
}

fn unique_field(fields: &[WireField], field_number: u32) -> Result<Option<&WireField>> {
    let mut matches = fields.iter().filter(|field| field.number() == field_number);
    let first = matches.next();
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart drawable field {field_number} occurs more than once"
        )));
    }
    Ok(first)
}

fn length_delimited_payload<'a>(data: &'a [u8], field: &WireField) -> Result<&'a [u8]> {
    if field.wire_type() != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart drawable field {} is not length-delimited",
            field.number()
        )));
    }
    Ok(field.checked_payload(data)?)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::{tsd, tsp};

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    #[test]
    fn typed_chart_extension_round_trips() {
        let archive = IWorkChartArchive::new(
            tsch::ChartDrawableArchive {
                super_: Some(tsd::DrawableArchive {
                    parent: Some(reference(42)),
                    ..Default::default()
                }),
            },
            tsch::ChartArchive {
                chart_type: Some(tsch::ChartType::ColumnChartType2D as i32),
                contains_default_data: Some(true),
                ..Default::default()
            },
        );

        let encoded = archive.encode().unwrap();
        assert_eq!(IWorkChartArchive::decode(&encoded).unwrap(), archive);
    }

    #[test]
    fn unknown_fields_are_preserved_byte_for_byte() {
        let archive = IWorkChartArchive::default();
        let mut encoded = archive.encode().unwrap();
        append_length_delimited_field(&mut encoded, 77, b"future").unwrap();

        assert_eq!(
            IWorkChartArchive::decode(&encoded)
                .unwrap()
                .encode()
                .unwrap(),
            encoded
        );
    }

    #[test]
    fn unknown_chart_fields_are_preserved_byte_for_byte() {
        let mut archive = IWorkChartArchive::new(
            tsch::ChartDrawableArchive::default(),
            tsch::ChartArchive::default(),
        );
        archive.append_chart_bool_extension(10_026, true).unwrap();
        let encoded = archive.encode().unwrap();

        assert_eq!(
            IWorkChartArchive::decode(&encoded)
                .unwrap()
                .encode()
                .unwrap(),
            encoded
        );
    }

    #[test]
    fn duplicate_chart_extension_is_rejected() {
        let mut encoded = IWorkChartArchive::new(
            tsch::ChartDrawableArchive::default(),
            tsch::ChartArchive::default(),
        )
        .encode()
        .unwrap();
        append_length_delimited_field(
            &mut encoded,
            CHART_EXTENSION_FIELD,
            &tsch::ChartArchive::default().encode_to_vec(),
        )
        .unwrap();

        assert!(IWorkChartArchive::decode(&encoded).is_err());
    }

    #[test]
    fn remaps_all_typed_references_and_rejects_unknown_private_metadata() {
        let mut archive = IWorkChartArchive::new(
            tsch::ChartDrawableArchive {
                super_: Some(tsd::DrawableArchive {
                    parent: Some(reference(1)),
                    comment: Some(reference(2)),
                    pencil_annotations: vec![reference(3)],
                    title: Some(reference(4)),
                    caption: Some(reference(5)),
                    ..Default::default()
                }),
            },
            tsch::ChartArchive {
                preset: Some(reference(6)),
                mediator: Some(reference(7)),
                chart_style: Some(reference(8)),
                chart_non_style: Some(reference(9)),
                legend_style: Some(reference(10)),
                legend_non_style: Some(reference(11)),
                value_axis_styles: vec![reference(12)],
                value_axis_nonstyles: vec![reference(13)],
                category_axis_styles: vec![reference(14)],
                category_axis_nonstyles: vec![reference(15)],
                series_theme_styles: vec![reference(16)],
                series_private_styles: Some(tsp::SparseReferenceArray {
                    count: 1,
                    entries: vec![tsp::sparse_reference_array::Entry {
                        index: 0,
                        reference: reference(17),
                    }],
                }),
                series_non_styles: Some(tsp::SparseReferenceArray {
                    count: 1,
                    entries: vec![tsp::sparse_reference_array::Entry {
                        index: 0,
                        reference: reference(18),
                    }],
                }),
                paragraph_styles: vec![reference(19)],
                owned_preset: Some(reference(20)),
                ..Default::default()
            },
        );
        let remap = (1_u64..=20)
            .map(|identifier| (identifier, identifier + 100))
            .collect::<HashMap<_, _>>();
        let recorded_references = (1_u64..=20).collect::<Vec<_>>();

        archive
            .remap_references(&remap, &recorded_references)
            .unwrap();
        let drawable = archive.drawable.super_.as_ref().unwrap();
        assert_eq!(drawable.parent.as_ref().unwrap().identifier, 101);
        assert_eq!(drawable.comment.as_ref().unwrap().identifier, 102);
        assert_eq!(drawable.pencil_annotations[0].identifier, 103);
        assert_eq!(drawable.title.as_ref().unwrap().identifier, 104);
        assert_eq!(drawable.caption.as_ref().unwrap().identifier, 105);
        let chart = archive.chart.as_ref().unwrap();
        assert_eq!(chart.preset.as_ref().unwrap().identifier, 106);
        assert_eq!(chart.mediator.as_ref().unwrap().identifier, 107);
        assert_eq!(chart.chart_style.as_ref().unwrap().identifier, 108);
        assert_eq!(chart.chart_non_style.as_ref().unwrap().identifier, 109);
        assert_eq!(chart.legend_style.as_ref().unwrap().identifier, 110);
        assert_eq!(chart.legend_non_style.as_ref().unwrap().identifier, 111);
        assert_eq!(chart.value_axis_styles[0].identifier, 112);
        assert_eq!(chart.value_axis_nonstyles[0].identifier, 113);
        assert_eq!(chart.category_axis_styles[0].identifier, 114);
        assert_eq!(chart.category_axis_nonstyles[0].identifier, 115);
        assert_eq!(chart.series_theme_styles[0].identifier, 116);
        assert_eq!(
            chart.series_private_styles.as_ref().unwrap().entries[0]
                .reference
                .identifier,
            117
        );
        assert_eq!(
            chart.series_non_styles.as_ref().unwrap().entries[0]
                .reference
                .identifier,
            118
        );
        assert_eq!(chart.paragraph_styles[0].identifier, 119);
        assert_eq!(chart.owned_preset.as_ref().unwrap().identifier, 120);
        assert_eq!(
            IWorkChartArchive::decode(&archive.encode().unwrap()).unwrap(),
            archive
        );

        let mut unsupported = IWorkChartArchive::default();
        assert!(
            unsupported
                .remap_references(&HashMap::from([(21, 121)]), &[21])
                .is_err()
        );
    }

    #[test]
    fn reference_line_graph_is_bounded_before_materialization() {
        let mut axis = Vec::new();
        for _ in 0..=MAX_REFERENCE_LINE_GRAPH_ENTRIES {
            append_length_delimited_field(&mut axis, 2, &[]).unwrap();
        }
        let mut graph = Vec::new();
        append_length_delimited_field(&mut graph, 1, &axis).unwrap();
        assert!(preflight_reference_line_graph(&graph).is_err());

        let mut axes = Vec::new();
        for _ in 0..=MAX_REFERENCE_LINE_GRAPH_AXES {
            append_length_delimited_field(&mut axes, 1, &[]).unwrap();
        }
        assert!(preflight_reference_line_graph(&axes).is_err());
    }

    #[test]
    fn reference_line_graph_unknown_fields_survive_typed_updates() {
        let old_payload = reference_line_graph_with_unknown_fields(1, 2, 3, 4);
        let replacement = reference_line_graph(11, 12, 13, 14);
        let mut archive = IWorkChartArchive::default();
        archive.chart_opaque_fields.push(
            encoded_chart_extension_payload(CHART_REFERENCE_LINES_EXTENSION_FIELD, &old_payload)
                .unwrap(),
        );

        archive.set_reference_lines(Some(&replacement)).unwrap();

        assert_eq!(archive.reference_lines().unwrap(), Some(replacement));
        let encoded = archive.chart_opaque_fields[0].clone();
        for unknown in [
            [0xd0, 0x05, 9],
            [0xd8, 0x05, 10],
            [0xe0, 0x05, 11],
            [0xe8, 0x05, 12],
            [0xf0, 0x05, 13],
            [0xf8, 0x05, 14],
            [0x80, 0x06, 15],
            [0x88, 0x06, 16],
            [0x90, 0x06, 17],
            [0x98, 0x06, 18],
            [0xa0, 0x06, 19],
            [0xa8, 0x06, 20],
            [0xb0, 0x06, 21],
            [0xb8, 0x06, 22],
        ] {
            assert!(
                contains_bytes(&encoded, &unknown),
                "missing unknown field {unknown:?}"
            );
        }
    }

    #[test]
    fn malformed_reference_line_graph_update_is_atomic() {
        let mut archive = IWorkChartArchive::default();
        archive.chart_opaque_fields.push(
            encoded_chart_extension_payload(
                CHART_REFERENCE_LINES_EXTENSION_FIELD,
                &[0x0a, 0x01, 0x80],
            )
            .unwrap(),
        );
        let before = archive.chart_opaque_fields.clone();

        assert!(
            archive
                .set_reference_lines(Some(&reference_line_graph(11, 12, 13, 14)))
                .is_err()
        );
        assert_eq!(archive.chart_opaque_fields, before);
    }

    #[test]
    fn malformed_nested_reference_line_graph_is_rejected_before_read_or_patch() {
        let malformed_axis_id = [0x88, 0x00, 0x00];
        let mut malformed_axis = Vec::new();
        append_length_delimited_field(&mut malformed_axis, 1, &malformed_axis_id).unwrap();
        let mut malformed_graph = Vec::new();
        append_length_delimited_field(&mut malformed_graph, 1, &malformed_axis).unwrap();

        let mut archive = IWorkChartArchive::default();
        archive.chart_opaque_fields.push(
            encoded_chart_extension_payload(
                CHART_REFERENCE_LINES_EXTENSION_FIELD,
                &malformed_graph,
            )
            .unwrap(),
        );
        let before = archive.chart_opaque_fields.clone();

        assert!(archive.reference_lines().is_err());
        assert!(
            archive
                .set_reference_lines(Some(&tsch::ChartReferenceLinesArchive::default()))
                .is_err()
        );
        assert!(archive.set_reference_lines(None).is_err());
        assert_eq!(archive.chart_opaque_fields, before);
    }

    #[test]
    fn noncanonical_reference_line_extension_is_rejected_atomically() {
        let payload = reference_line_graph(11, 12, 13, 14).encode_to_vec();
        let canonical =
            encoded_chart_extension_payload(CHART_REFERENCE_LINES_EXTENSION_FIELD, &payload)
                .unwrap();
        let canonical_fields = parse_wire_fields(&canonical).unwrap();
        let [field] = canonical_fields.as_slice() else {
            panic!("encoded reference-line extension must have one field");
        };
        let mut noncanonical = Vec::with_capacity(canonical.len() + 1);
        noncanonical.extend_from_slice(&[0xaa, 0xf1, 0x84, 0x00]);
        noncanonical.extend_from_slice(&canonical[field.key_end()..]);

        let mut archive = IWorkChartArchive::default();
        archive.chart_opaque_fields.push(noncanonical);
        let before = archive.chart_opaque_fields.clone();

        assert!(archive.reference_lines().is_err());
        assert!(
            archive
                .set_reference_lines(Some(&reference_line_graph(21, 22, 23, 24)))
                .is_err()
        );
        assert_eq!(archive.chart_opaque_fields, before);
    }

    fn reference_line_graph(
        non_style_identifier: u64,
        non_style_uuid: u64,
        style_identifier: u64,
        style_uuid: u64,
    ) -> tsch::ChartReferenceLinesArchive {
        tsch::ChartReferenceLinesArchive {
            reference_line_non_styles_map: vec![tsch::ChartAxisReferenceLineNonStylesArchive {
                axis_id: primary_axis_id(),
                reference_line_non_style_items: vec![tsch::ChartReferenceLineNonStyleItem {
                    non_style: reference(non_style_identifier),
                    uuid: tsp::Uuid {
                        lower: non_style_uuid,
                        upper: non_style_uuid + 100,
                    },
                }],
            }],
            reference_line_styles_map: vec![tsch::ChartAxisReferenceLineStylesArchive {
                axis_id: primary_axis_id(),
                reference_line_styles: Some(tsp::SparseReferenceArray {
                    count: 1,
                    entries: vec![tsp::sparse_reference_array::Entry {
                        index: 0,
                        reference: reference(style_identifier),
                    }],
                }),
            }],
            theme_preset_reference_line_style: Some(reference(style_uuid)),
        }
    }

    fn reference_line_graph_with_unknown_fields(
        non_style_identifier: u64,
        non_style_uuid: u64,
        style_identifier: u64,
        style_uuid: u64,
    ) -> Vec<u8> {
        let graph = reference_line_graph(
            non_style_identifier,
            non_style_uuid,
            style_identifier,
            style_uuid,
        );
        let canonical = graph.encode_to_vec();
        let fields = parse_wire_fields(&canonical).unwrap();
        let non_style_axis = length_delimited_payload(
            &canonical,
            fields.iter().find(|field| field.number() == 1).unwrap(),
        )
        .unwrap();
        let style_axis = length_delimited_payload(
            &canonical,
            fields.iter().find(|field| field.number() == 2).unwrap(),
        )
        .unwrap();
        let non_style_axis_fields = parse_wire_fields(non_style_axis).unwrap();
        let item = length_delimited_payload(
            non_style_axis,
            non_style_axis_fields
                .iter()
                .find(|field| field.number() == 2)
                .unwrap(),
        )
        .unwrap();
        let item_fields = parse_wire_fields(item).unwrap();
        let mut reference_payload = length_delimited_payload(
            item,
            item_fields
                .iter()
                .find(|field| field.number() == 1)
                .unwrap(),
        )
        .unwrap()
        .to_vec();
        append_varint_field(&mut reference_payload, 90, 9).unwrap();
        let mut uuid_payload = length_delimited_payload(
            item,
            item_fields
                .iter()
                .find(|field| field.number() == 2)
                .unwrap(),
        )
        .unwrap()
        .to_vec();
        append_varint_field(&mut uuid_payload, 91, 10).unwrap();

        let mut item_with_unknown = Vec::new();
        append_length_delimited_field(&mut item_with_unknown, 1, &reference_payload).unwrap();
        append_varint_field(&mut item_with_unknown, 92, 11).unwrap();
        append_length_delimited_field(&mut item_with_unknown, 2, &uuid_payload).unwrap();
        let mut axis_id = primary_axis_id().encode_to_vec();
        append_varint_field(&mut axis_id, 93, 12).unwrap();
        let mut non_style_axis_with_unknown = Vec::new();
        append_varint_field(&mut non_style_axis_with_unknown, 94, 13).unwrap();
        append_length_delimited_field(&mut non_style_axis_with_unknown, 1, &axis_id).unwrap();
        append_varint_field(&mut non_style_axis_with_unknown, 95, 14).unwrap();
        append_length_delimited_field(&mut non_style_axis_with_unknown, 2, &item_with_unknown)
            .unwrap();

        let style_axis_fields = parse_wire_fields(style_axis).unwrap();
        let sparse = length_delimited_payload(
            style_axis,
            style_axis_fields
                .iter()
                .find(|field| field.number() == 2)
                .unwrap(),
        )
        .unwrap();
        let sparse_fields = parse_wire_fields(sparse).unwrap();
        let entry = length_delimited_payload(
            sparse,
            sparse_fields
                .iter()
                .find(|field| field.number() == 2)
                .unwrap(),
        )
        .unwrap();
        let entry_fields = parse_wire_fields(entry).unwrap();
        let sparse_reference = length_delimited_payload(
            entry,
            entry_fields
                .iter()
                .find(|field| field.number() == 2)
                .unwrap(),
        )
        .unwrap();
        let mut sparse_reference_with_unknown = sparse_reference.to_vec();
        append_varint_field(&mut sparse_reference_with_unknown, 96, 15).unwrap();
        let mut entry_with_unknown = Vec::new();
        append_varint_field(&mut entry_with_unknown, 1, 0).unwrap();
        append_length_delimited_field(&mut entry_with_unknown, 2, &sparse_reference_with_unknown)
            .unwrap();
        append_varint_field(&mut entry_with_unknown, 97, 16).unwrap();
        let mut sparse_with_unknown = Vec::new();
        append_varint_field(&mut sparse_with_unknown, 1, 1).unwrap();
        append_varint_field(&mut sparse_with_unknown, 98, 17).unwrap();
        append_length_delimited_field(&mut sparse_with_unknown, 2, &entry_with_unknown).unwrap();
        let mut style_axis_id = primary_axis_id().encode_to_vec();
        append_varint_field(&mut style_axis_id, 99, 18).unwrap();
        let mut style_axis_with_unknown = Vec::new();
        append_length_delimited_field(&mut style_axis_with_unknown, 1, &style_axis_id).unwrap();
        append_varint_field(&mut style_axis_with_unknown, 100, 19).unwrap();
        append_length_delimited_field(&mut style_axis_with_unknown, 2, &sparse_with_unknown)
            .unwrap();
        let mut output = Vec::new();
        append_varint_field(&mut output, 101, 20).unwrap();
        append_length_delimited_field(&mut output, 1, &non_style_axis_with_unknown).unwrap();
        append_varint_field(&mut output, 102, 21).unwrap();
        append_length_delimited_field(&mut output, 2, &style_axis_with_unknown).unwrap();
        append_varint_field(&mut output, 103, 22).unwrap();
        output
    }

    fn primary_axis_id() -> tsch::ChartAxisIdArchive {
        tsch::ChartAxisIdArchive {
            axis_type: Some(tsch::AxisType::Y as i32),
            ordinal: Some(0),
        }
    }

    fn contains_bytes(haystack: &[u8], needle: &[u8]) -> bool {
        haystack
            .windows(needle.len())
            .any(|window| window == needle)
    }
}
