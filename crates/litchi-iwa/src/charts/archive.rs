//! Lossless codec for extension-backed native chart drawables.
//!
//! `TSCH.ChartArchive` is protobuf extension field 10000 of
//! `TSCH.ChartDrawableArchive`. `prost` does not expose proto2 extensions, so a
//! direct decode of the generated drawable type loses the chart itself.

use std::collections::{HashMap, HashSet};

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
        let Some(encoded) = unique_opaque_field(
            &self.chart_opaque_fields,
            CHART_REFERENCE_LINES_EXTENSION_FIELD,
        )?
        else {
            return Ok(None);
        };
        let fields = parse_wire_fields(encoded)?;
        let field =
            unique_field(&fields, CHART_REFERENCE_LINES_EXTENSION_FIELD)?.ok_or_else(|| {
                Error::InvalidFormat(
                    "chart reference-line extension disappeared during decoding".to_owned(),
                )
            })?;
        Ok(Some(tsch::ChartReferenceLinesArchive::decode(
            length_delimited_payload(encoded, field)?,
        )?))
    }

    pub(crate) fn set_reference_lines(
        &mut self,
        reference_lines: Option<&tsch::ChartReferenceLinesArchive>,
    ) -> Result<()> {
        let existing_index = unique_opaque_field_index(
            &self.chart_opaque_fields,
            CHART_REFERENCE_LINES_EXTENSION_FIELD,
        )?;
        match (existing_index, reference_lines) {
            (Some(index), Some(reference_lines)) => {
                self.chart_opaque_fields[index] = encoded_chart_extension(
                    CHART_REFERENCE_LINES_EXTENSION_FIELD,
                    reference_lines,
                )?;
            },
            (Some(index), None) => {
                self.chart_opaque_fields.remove(index);
            },
            (None, Some(reference_lines)) => {
                let encoded = encoded_chart_extension(
                    CHART_REFERENCE_LINES_EXTENSION_FIELD,
                    reference_lines,
                )?;
                let index = self
                    .chart_opaque_fields
                    .iter()
                    .position(|field| {
                        parse_wire_fields(field)
                            .ok()
                            .and_then(|fields| fields.first().map(|field| field.number()))
                            .is_some_and(|number| number > CHART_REFERENCE_LINES_EXTENSION_FIELD)
                    })
                    .unwrap_or(self.chart_opaque_fields.len());
                self.chart_opaque_fields.insert(index, encoded);
            },
            (None, None) => {},
        }
        if self.reference_lines()? != reference_lines.cloned() {
            return Err(Error::InvalidFormat(
                "chart reference-line extension update failed validation".to_owned(),
            ));
        }
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

fn encoded_chart_extension(field_number: u32, message: &impl Message) -> Result<Vec<u8>> {
    let mut field = Vec::new();
    append_length_delimited_field(&mut field, field_number, &message.encode_to_vec())?;
    Ok(field)
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
    Ok(&data[field.payload_start()..field.end()])
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
}
