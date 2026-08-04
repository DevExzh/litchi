//! Strict protobuf-wire decoding for table appearance overrides.

use crate::wire::parse_wire_fields;
use crate::{Error, Result};

const TABLE_STYLE_PROPERTIES_FIELD: u32 = 11;
const BANDED_ROWS_FIELD: u32 = 1;
const AUTO_RESIZE_FIELD: u32 = 22;
const VERTICAL_BODY_GRIDLINES_FIELD: u32 = 33;
const HORIZONTAL_BODY_GRIDLINES_FIELD: u32 = 34;
const LEGACY_VERTICAL_HEADER_ROW_GRIDLINES_FIELD: u32 = 35;
const LEGACY_HORIZONTAL_HEADER_COLUMN_GRIDLINES_FIELD: u32 = 36;
const LEGACY_VERTICAL_FOOTER_ROW_GRIDLINES_FIELD: u32 = 37;
const HORIZONTAL_HEADER_COLUMN_GRIDLINES_FIELD: u32 = 42;
const VERTICAL_HEADER_ROW_GRIDLINES_FIELD: u32 = 43;
const VERTICAL_FOOTER_ROW_GRIDLINES_FIELD: u32 = 44;

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) struct TableAppearanceOverrides {
    pub(super) banded_rows: Option<bool>,
    pub(super) auto_resize: Option<bool>,
    pub(super) horizontal_body_gridlines: Option<bool>,
    pub(super) horizontal_header_column_gridlines: Option<bool>,
    pub(super) vertical_body_gridlines: Option<bool>,
    pub(super) vertical_header_row_gridlines: Option<bool>,
    pub(super) vertical_footer_row_gridlines: Option<bool>,
}

pub(super) fn table_appearance_overrides(data: &[u8]) -> Result<TableAppearanceOverrides> {
    let fields = parse_wire_fields(data)?;
    let mut property_fields = fields
        .iter()
        .filter(|field| field.number() == TABLE_STYLE_PROPERTIES_FIELD);
    let Some(properties) = property_fields.next() else {
        return Ok(TableAppearanceOverrides::default());
    };
    if property_fields.next().is_some() {
        return Err(Error::InvalidFormat(
            "iWork table-style properties occur more than once".to_owned(),
        ));
    }
    if properties.wire_type() != 2 {
        return Err(Error::InvalidFormat(format!(
            "iWork table-style properties use wire type {}, expected 2",
            properties.wire_type()
        )));
    }
    let properties = &data[properties.payload_start()..properties.end()];
    Ok(TableAppearanceOverrides {
        banded_rows: strict_optional_bool(properties, BANDED_ROWS_FIELD, "banded rows")?,
        auto_resize: strict_optional_bool(properties, AUTO_RESIZE_FIELD, "automatic row sizing")?,
        horizontal_body_gridlines: strict_optional_bool(
            properties,
            HORIZONTAL_BODY_GRIDLINES_FIELD,
            "horizontal body gridlines",
        )?,
        horizontal_header_column_gridlines: strict_optional_bool(
            properties,
            HORIZONTAL_HEADER_COLUMN_GRIDLINES_FIELD,
            "horizontal header-column gridlines",
        )?
        .or(strict_optional_bool(
            properties,
            LEGACY_HORIZONTAL_HEADER_COLUMN_GRIDLINES_FIELD,
            "legacy horizontal header-column gridlines",
        )?),
        vertical_body_gridlines: strict_optional_bool(
            properties,
            VERTICAL_BODY_GRIDLINES_FIELD,
            "vertical body gridlines",
        )?,
        vertical_header_row_gridlines: strict_optional_bool(
            properties,
            VERTICAL_HEADER_ROW_GRIDLINES_FIELD,
            "vertical header-row gridlines",
        )?
        .or(strict_optional_bool(
            properties,
            LEGACY_VERTICAL_HEADER_ROW_GRIDLINES_FIELD,
            "legacy vertical header-row gridlines",
        )?),
        vertical_footer_row_gridlines: strict_optional_bool(
            properties,
            VERTICAL_FOOTER_ROW_GRIDLINES_FIELD,
            "vertical footer-row gridlines",
        )?
        .or(strict_optional_bool(
            properties,
            LEGACY_VERTICAL_FOOTER_ROW_GRIDLINES_FIELD,
            "legacy vertical footer-row gridlines",
        )?),
    })
}

fn strict_optional_bool(data: &[u8], field_number: u32, label: &str) -> Result<Option<bool>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number() == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "iWork table {label} field occurs more than once"
        )));
    }
    if field.wire_type() != 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork table {label} uses wire type {}, expected 0",
            field.wire_type()
        )));
    }
    let (value, length) =
        litchi_iwa_common::varint::decode_varint_from_bytes(&data[field.payload_start()..field.end()])
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid iWork table {label}: {error}"))
            })?;
    if field.payload_start() + length != field.end() {
        return Err(Error::InvalidFormat(format!(
            "iWork table {label} contains trailing bytes"
        )));
    }
    match value {
        0 => Ok(Some(false)),
        1 => Ok(Some(true)),
        _ => Err(Error::InvalidFormat(format!(
            "iWork table {label} must be encoded as zero or one, found {value}"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use prost::Message;

    use super::*;
    use crate::protobuf::tst;

    #[test]
    fn appearance_overrides_are_strict_and_typed() {
        let style = tst::TableStyleArchive {
            table_properties: Some(tst::TableStylePropertiesArchive {
                banded_rows: Some(true),
                auto_resize: Some(false),
                h_strokes_visible: Some(false),
                hc_separator_visible: Some(false),
                v_strokes_visible: Some(false),
                hr_separator_visible: Some(false),
                footer_separator_visible: Some(true),
                table_hc_divider_visible: Some(true),
                table_hr_divider_visible: Some(true),
                table_footer_divider_visible: Some(false),
                ..Default::default()
            }),
            ..Default::default()
        }
        .encode_to_vec();
        assert_eq!(
            table_appearance_overrides(&style).unwrap(),
            TableAppearanceOverrides {
                banded_rows: Some(true),
                auto_resize: Some(false),
                horizontal_body_gridlines: Some(false),
                horizontal_header_column_gridlines: Some(true),
                vertical_body_gridlines: Some(false),
                vertical_header_row_gridlines: Some(true),
                vertical_footer_row_gridlines: Some(false),
            }
        );

        let legacy_style = tst::TableStyleArchive {
            table_properties: Some(tst::TableStylePropertiesArchive {
                hc_separator_visible: Some(true),
                hr_separator_visible: Some(false),
                footer_separator_visible: Some(true),
                ..Default::default()
            }),
            ..Default::default()
        }
        .encode_to_vec();
        let legacy = table_appearance_overrides(&legacy_style).unwrap();
        assert_eq!(legacy.horizontal_header_column_gridlines, Some(true));
        assert_eq!(legacy.vertical_header_row_gridlines, Some(false));
        assert_eq!(legacy.vertical_footer_row_gridlines, Some(true));

        let mut duplicate_properties = style;
        duplicate_properties.extend_from_slice(&[0x5a, 0x00]);
        assert!(table_appearance_overrides(&duplicate_properties).is_err());
        assert!(table_appearance_overrides(&[0x5a, 0x02, 0x08, 0x02]).is_err());
        assert!(table_appearance_overrides(&[0x5a, 0x02, 0x0a, 0x00]).is_err());
    }
}
