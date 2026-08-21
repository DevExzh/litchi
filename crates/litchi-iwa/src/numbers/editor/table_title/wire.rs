//! Lossless protobuf wire handling for native iWork table title settings.

use super::*;

const TITLE_VISIBLE_FIELD: u32 = 22;
const TITLE_OUTLINED_FIELD: u32 = 37;
const TITLE_CODEC_RECURSION_LIMIT: u32 = 64;

fn decode_title_settings(
    source: &[u8],
) -> Result<litchi_iwa_protos::numbers_table_title_codec::TableTitleSettingsSnapshot> {
    use litchi_iwa_protos::numbers_table_title_codec::DecodeOptions;

    litchi_iwa_protos::numbers_table_title_codec::decode_table_title_settings(
        source,
        DecodeOptions::new(
            source.len().max(1),
            source.len().max(1),
            source.len().saturating_mul(4).max(1),
            TITLE_CODEC_RECURSION_LIMIT,
            2,
        ),
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "iWork table title payload failed strict validation: {error}"
        ))
    })
}

pub(super) fn read_table_title_settings_wire(original: &[u8]) -> Result<Settings> {
    let snapshot = decode_title_settings(original)?;
    Ok(Settings::new(
        snapshot.table_name_enabled(),
        snapshot.table_name_border_enabled(),
    ))
}

pub(super) fn write_table_title_settings_wire(
    original: &[u8],
    settings: Settings,
) -> Result<Vec<u8>> {
    let before = decode_title_settings(original)?;
    let mut data = patch_varint_field(
        original,
        TITLE_VISIBLE_FIELD,
        before.table_name_enabled().is_some(),
        settings.visible().map(u64::from),
    )?;
    data = patch_varint_field(
        &data,
        TITLE_OUTLINED_FIELD,
        before.table_name_border_enabled().is_some(),
        settings.outlined().map(u64::from),
    )?;
    if read_table_title_settings_wire(&data)? != settings {
        return Err(Error::InvalidFormat(
            "iWork table title wire patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strict_title_patch_preserves_unknown_fields_and_presence() {
        let original = [
            0xb0, 0x01, 0x01, // field 22: visible = true
            0xa8, 0x02, 0x01, // field 37: outlined = true
            0x98, 0x06, 0x2a, // unknown field 99
        ];
        let before = read_table_title_settings_wire(&original).unwrap();
        assert_eq!(before, Settings::new(Some(true), Some(true)));

        let settings = Settings::new(None, Some(false));
        let patched = write_table_title_settings_wire(&original, settings).unwrap();
        assert_eq!(read_table_title_settings_wire(&patched).unwrap(), settings);
        assert!(patched.ends_with(&[0x98, 0x06, 0x2a]));
        assert!(!patched.windows(2).any(|window| window == [0xb0, 0x01]));
    }

    #[test]
    fn strict_title_patch_rejects_duplicate_known_fields() {
        let malformed = [
            0xb0, 0x01, 0x01, // field 22: visible = true
            0xb0, 0x01, 0x00, // duplicate field 22
        ];
        assert!(read_table_title_settings_wire(&malformed).is_err());
        assert!(write_table_title_settings_wire(&malformed, Settings::default()).is_err());
    }
}
