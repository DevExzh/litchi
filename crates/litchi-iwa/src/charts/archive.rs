//! Lossless codec for extension-backed native chart drawables.
//!
//! `TSCH.ChartArchive` is protobuf extension field 10000 of
//! `TSCH.ChartDrawableArchive`. `prost` does not expose proto2 extensions, so a
//! direct decode of the generated drawable type loses the chart itself.

use prost::Message;

use crate::protobuf::tsch;
use crate::wire::{WireField, append_length_delimited_field, parse_wire_fields};
use crate::{Error, Result};

const DRAWABLE_SUPER_FIELD: u32 = 1;
const CHART_EXTENSION_FIELD: u32 = 10_000;

/// A native chart drawable with its extension-backed chart payload retained.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct IWorkChartArchive {
    pub drawable: tsch::ChartDrawableArchive,
    pub chart: Option<tsch::ChartArchive>,
    opaque_fields: Vec<Vec<u8>>,
}

impl IWorkChartArchive {
    /// Construct a chart drawable from typed protobuf values.
    pub fn new(drawable: tsch::ChartDrawableArchive, chart: tsch::ChartArchive) -> Self {
        Self {
            drawable,
            chart: Some(chart),
            opaque_fields: Vec::new(),
        }
    }

    /// Decode a chart drawable without discarding extensions or future fields.
    pub fn decode(data: &[u8]) -> Result<Self> {
        let fields = parse_wire_fields(data)?;
        let chart = unique_field(&fields, CHART_EXTENSION_FIELD)?
            .map(|field| {
                tsch::ChartArchive::decode(length_delimited_payload(data, field)?)
                    .map_err(Error::from)
            })
            .transpose()?;
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
            opaque_fields,
        })
    }

    /// Encode the drawable, chart extension, and untouched future fields.
    pub fn encode(&self) -> Result<Vec<u8>> {
        let mut output = self.drawable.encode_to_vec();
        if let Some(chart) = &self.chart {
            append_length_delimited_field(
                &mut output,
                CHART_EXTENSION_FIELD,
                &chart.encode_to_vec(),
            )?;
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
