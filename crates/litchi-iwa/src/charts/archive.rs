//! Lossless codec for extension-backed native chart drawables.
//!
//! `TSCH.ChartArchive` is protobuf extension field 10000 of
//! `TSCH.ChartDrawableArchive`. `prost` does not expose proto2 extensions, so a
//! direct decode of the generated drawable type loses the chart itself.

use prost::Message;

use crate::protobuf::tsch;
use crate::wire::{
    WireField, append_length_delimited_field, append_varint_field, parse_wire_fields,
};
use crate::{Error, Result};

const DRAWABLE_SUPER_FIELD: u32 = 1;
const CHART_EXTENSION_FIELD: u32 = 10_000;
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
                .filter(|field| !CHART_BASE_FIELDS.contains(&field.number))
                .map(|field| chart_data[field.start..field.end].to_vec())
                .collect()
        } else {
            Vec::new()
        };
        let opaque_fields = fields
            .iter()
            .filter(|field| {
                field.number != DRAWABLE_SUPER_FIELD && field.number != CHART_EXTENSION_FIELD
            })
            .map(|field| data[field.start..field.end].to_vec())
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

fn unique_field(fields: &[WireField], field_number: u32) -> Result<Option<&WireField>> {
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let first = matches.next();
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "chart drawable field {field_number} occurs more than once"
        )));
    }
    Ok(first)
}

fn length_delimited_payload<'a>(data: &'a [u8], field: &WireField) -> Result<&'a [u8]> {
    if field.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "chart drawable field {} is not length-delimited",
            field.number
        )));
    }
    Ok(&data[field.payload_start..field.end])
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
}
