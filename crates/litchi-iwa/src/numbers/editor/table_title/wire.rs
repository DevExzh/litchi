//! Lossless protobuf wire handling for native iWork table title settings.

use super::*;

const TITLE_VISIBLE_FIELD: u32 = 22;
const TITLE_OUTLINED_FIELD: u32 = 37;

pub(super) fn read_table_title_settings_wire(
    original: &[u8],
    model: &TableModelArchive,
) -> Result<Settings> {
    require_optional_bool(
        original,
        TITLE_VISIBLE_FIELD,
        "title visibility",
        model.table_name_enabled,
    )?;
    require_optional_bool(
        original,
        TITLE_OUTLINED_FIELD,
        "title outline",
        model.table_name_border_enabled,
    )?;
    Ok(Settings::new(
        model.table_name_enabled,
        model.table_name_border_enabled,
    ))
}

pub(super) fn write_table_title_settings_wire(
    original: &[u8],
    model: &TableModelArchive,
    settings: Settings,
) -> Result<Vec<u8>> {
    read_table_title_settings_wire(original, model)?;
    let mut data = patch_varint_field(
        original,
        TITLE_VISIBLE_FIELD,
        model.table_name_enabled.is_some(),
        settings.visible().map(u64::from),
    )?;
    data = patch_varint_field(
        &data,
        TITLE_OUTLINED_FIELD,
        model.table_name_border_enabled.is_some(),
        settings.outlined().map(u64::from),
    )?;
    let verified = TableModelArchive::decode(data.as_slice())?;
    if read_table_title_settings_wire(&data, &verified)? != settings {
        return Err(Error::InvalidFormat(
            "iWork table title wire patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn require_optional_bool(
    original: &[u8],
    field: u32,
    label: &str,
    decoded: Option<bool>,
) -> Result<()> {
    let values = repeated_varint_values(original, field)?;
    let expected = decoded.map(u64::from);
    if values.as_slice() != expected.as_slice() {
        return Err(Error::InvalidFormat(format!(
            "iWork table {label} wire value is missing, duplicated, non-Boolean, or inconsistent"
        )));
    }
    Ok(())
}
