//! Typed encoding for the extension-backed iWork theme archive.
//!
//! Apple's base `TSS.ThemeArchive` declares its application preset families as
//! proto2 extensions. `prost` intentionally ignores extension declarations, so
//! decoding and re-encoding the generated base message alone silently drops the
//! drawing, text, chart, table, and application preset catalogs. This codec
//! makes those five fields explicit without embedding opaque package data.

use prost::Message;

use crate::protobuf::{tsa, tsch, tsd, tss, tst, tswp};
use crate::wire::{WireField, append_length_delimited_field, parse_wire_fields};
use crate::{Error, Result};

const THEME_SUPER_FIELD: u32 = 1;
const LEGACY_STYLESHEET_FIELD: u32 = 1;
const THEME_IDENTIFIER_FIELD: u32 = 3;
const DOCUMENT_STYLESHEET_FIELD: u32 = 4;
const OLD_PRESET_UUIDS_FIELD: u32 = 5;
const NEW_PRESET_UUIDS_FIELD: u32 = 6;
const COLOR_PRESETS_FIELD: u32 = 10;
const DRAWING_PRESETS_FIELD: u32 = 100;
const TEXT_PRESETS_FIELD: u32 = 110;
const CHART_PRESETS_FIELD: u32 = 120;
const TABLE_PRESETS_FIELD: u32 = 200;
const APPLICATION_PRESETS_FIELD: u32 = 210;

const BASE_THEME_FIELDS: &[u32] = &[
    LEGACY_STYLESHEET_FIELD,
    THEME_IDENTIFIER_FIELD,
    DOCUMENT_STYLESHEET_FIELD,
    OLD_PRESET_UUIDS_FIELD,
    NEW_PRESET_UUIDS_FIELD,
    COLOR_PRESETS_FIELD,
];

/// The five strongly-typed preset families stored as `TSS.ThemeArchive`
/// extensions by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct IWorkThemeExtensions {
    pub drawing: Option<tsd::ThemePresetsArchive>,
    pub text: Option<tswp::ThemePresetsArchive>,
    pub chart: Option<tsch::ChartPresetsArchive>,
    pub table: Option<tst::ThemePresetsArchive>,
    pub application: Option<tsa::ThemePresetsArchive>,
}

/// An iWork application theme wrapper with extension fields retained.
///
/// Pages, Numbers, and Keynote all wrap `TSS.ThemeArchive` in the same single
/// field. Unknown future fields are preserved byte-for-byte on round trip.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct IWorkThemeArchive {
    pub base: tss::ThemeArchive,
    pub extensions: IWorkThemeExtensions,
    opaque_base_fields: Vec<Vec<u8>>,
    opaque_wrapper_fields: Vec<Vec<u8>>,
}

impl IWorkThemeArchive {
    /// Construct a new theme from strongly typed base and extension archives.
    ///
    /// Unlike [`Self::decode`], a newly constructed theme has no opaque fields
    /// inherited from another package, making it suitable for documents built
    /// entirely from scratch.
    pub fn new(base: tss::ThemeArchive, extensions: IWorkThemeExtensions) -> Self {
        Self {
            base,
            extensions,
            opaque_base_fields: Vec::new(),
            opaque_wrapper_fields: Vec::new(),
        }
    }

    /// Decode an application theme without losing proto2 extension fields.
    pub fn decode(data: &[u8]) -> Result<Self> {
        let wrapper_fields = parse_wire_fields(data)?;
        let Some(theme_field) = unique_field(&wrapper_fields, THEME_SUPER_FIELD)? else {
            return Err(Error::InvalidFormat(format!(
                "iWork theme wrapper must contain exactly one field {THEME_SUPER_FIELD}"
            )));
        };
        let theme_data = length_delimited_payload(data, theme_field)?;
        let base_fields = parse_wire_fields(theme_data)?;

        let extensions = IWorkThemeExtensions {
            drawing: decode_extension(theme_data, &base_fields, DRAWING_PRESETS_FIELD)?,
            text: decode_extension(theme_data, &base_fields, TEXT_PRESETS_FIELD)?,
            chart: decode_extension(theme_data, &base_fields, CHART_PRESETS_FIELD)?,
            table: decode_extension(theme_data, &base_fields, TABLE_PRESETS_FIELD)?,
            application: decode_extension(theme_data, &base_fields, APPLICATION_PRESETS_FIELD)?,
        };
        let opaque_base_fields = base_fields
            .iter()
            .filter(|field| {
                !BASE_THEME_FIELDS.contains(&field.number) && !is_known_extension(field.number)
            })
            .map(|field| theme_data[field.start..field.end].to_vec())
            .collect();
        let opaque_wrapper_fields = wrapper_fields
            .iter()
            .filter(|field| field.number != THEME_SUPER_FIELD)
            .map(|field| data[field.start..field.end].to_vec())
            .collect();

        Ok(Self {
            base: tss::ThemeArchive::decode(theme_data)?,
            extensions,
            opaque_base_fields,
            opaque_wrapper_fields,
        })
    }

    /// Encode the complete theme wrapper, including all preset families.
    pub fn encode(&self) -> Result<Vec<u8>> {
        let mut theme = self.base.encode_to_vec();
        append_message(
            &mut theme,
            DRAWING_PRESETS_FIELD,
            self.extensions.drawing.as_ref(),
        )?;
        append_message(
            &mut theme,
            TEXT_PRESETS_FIELD,
            self.extensions.text.as_ref(),
        )?;
        append_message(
            &mut theme,
            CHART_PRESETS_FIELD,
            self.extensions.chart.as_ref(),
        )?;
        append_message(
            &mut theme,
            TABLE_PRESETS_FIELD,
            self.extensions.table.as_ref(),
        )?;
        append_message(
            &mut theme,
            APPLICATION_PRESETS_FIELD,
            self.extensions.application.as_ref(),
        )?;
        for field in &self.opaque_base_fields {
            theme.extend_from_slice(field);
        }

        let mut wrapper = Vec::with_capacity(theme.len() + 8);
        append_length_delimited_field(&mut wrapper, THEME_SUPER_FIELD, &theme)?;
        for field in &self.opaque_wrapper_fields {
            wrapper.extend_from_slice(field);
        }
        Ok(wrapper)
    }
}

fn decode_extension<M: Message + Default>(
    data: &[u8],
    fields: &[WireField],
    field_number: u32,
) -> Result<Option<M>> {
    unique_field(fields, field_number)?
        .map(|field| M::decode(length_delimited_payload(data, field)?).map_err(Error::from))
        .transpose()
}

fn unique_field(fields: &[WireField], field_number: u32) -> Result<Option<&WireField>> {
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let first = matches.next();
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "iWork theme field {field_number} occurs more than once"
        )));
    }
    Ok(first)
}

fn append_message<M: Message>(
    output: &mut Vec<u8>,
    field_number: u32,
    message: Option<&M>,
) -> Result<()> {
    if let Some(message) = message {
        append_length_delimited_field(output, field_number, &message.encode_to_vec())?;
    }
    Ok(())
}

fn length_delimited_payload<'a>(data: &'a [u8], field: &WireField) -> Result<&'a [u8]> {
    if field.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "iWork theme field {} is not length-delimited",
            field.number
        )));
    }
    Ok(&data[field.payload_start..field.end])
}

const fn is_known_extension(field_number: u32) -> bool {
    matches!(
        field_number,
        DRAWING_PRESETS_FIELD
            | TEXT_PRESETS_FIELD
            | CHART_PRESETS_FIELD
            | TABLE_PRESETS_FIELD
            | APPLICATION_PRESETS_FIELD
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tsp;

    #[test]
    fn all_theme_extension_families_round_trip_typed() {
        let theme = IWorkThemeArchive {
            base: tss::ThemeArchive {
                theme_identifier: Some("Litchi Blank".to_owned()),
                document_stylesheet: Some(reference(42)),
                ..Default::default()
            },
            extensions: IWorkThemeExtensions {
                drawing: Some(tsd::ThemePresetsArchive::default()),
                text: Some(tswp::ThemePresetsArchive {
                    list_style_presets: vec![reference(43)],
                    paragraph_style_presets: vec![reference(44)],
                    ..Default::default()
                }),
                chart: Some(tsch::ChartPresetsArchive::default()),
                table: Some(tst::ThemePresetsArchive::default()),
                application: Some(tsa::ThemePresetsArchive::default()),
            },
            ..Default::default()
        };

        let encoded = theme.encode().unwrap();
        assert_eq!(IWorkThemeArchive::decode(&encoded).unwrap(), theme);
    }

    #[test]
    fn unknown_extension_and_wrapper_fields_are_preserved() {
        let theme = IWorkThemeArchive::default();
        let mut encoded = theme.encode().unwrap();
        append_length_delimited_field(&mut encoded, 9, b"wrapper").unwrap();

        let fields = parse_wire_fields(&encoded).unwrap();
        let super_field = fields
            .iter()
            .find(|field| field.number == THEME_SUPER_FIELD)
            .unwrap();
        let mut theme_data = length_delimited_payload(&encoded, super_field)
            .unwrap()
            .to_vec();
        append_length_delimited_field(&mut theme_data, 333, b"future").unwrap();
        let mut encoded_with_unknown = Vec::new();
        append_length_delimited_field(&mut encoded_with_unknown, THEME_SUPER_FIELD, &theme_data)
            .unwrap();
        append_length_delimited_field(&mut encoded_with_unknown, 9, b"wrapper").unwrap();

        let decoded = IWorkThemeArchive::decode(&encoded_with_unknown).unwrap();
        let reencoded = decoded.encode().unwrap();
        assert_eq!(reencoded, encoded_with_unknown);
    }

    #[test]
    fn duplicate_theme_super_is_rejected() {
        let mut encoded = IWorkThemeArchive::default().encode().unwrap();
        append_length_delimited_field(&mut encoded, THEME_SUPER_FIELD, &[]).unwrap();
        assert!(IWorkThemeArchive::decode(&encoded).is_err());
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            deprecated_type: None,
            deprecated_is_external: None,
        }
    }
}
