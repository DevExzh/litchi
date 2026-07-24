//! Lossless native chart hidden-data participation CRUD.
//!
//! Numbers presents this behavior as a positive `Hidden Data` checkbox, while
//! its generated chart non-style field stores the inverse `skip hidden data`
//! value. Keynote and Pages do not expose this source-table control. This
//! module keeps the native inversion private and exposes positive,
//! developer-facing semantics to the Numbers adapter.

use prost::Message;

use crate::charts::non_style::{
    GENERATED_CHART_NON_STYLE_EXTENSION_FIELD, chart_non_style_slot,
    generated_chart_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefaultskiphiddendata` in
/// `TSCH.Generated.ChartNonStyleArchive`.
const CHART_SKIP_HIDDEN_DATA_FIELD: u32 = 22;

/// Read whether one native chart includes data from hidden rows and columns.
pub(crate) fn chart_includes_hidden_data(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_includes_hidden_data)
}

/// Set whether one native chart includes data from hidden rows and columns.
pub(crate) fn set_chart_includes_hidden_data(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    includes_hidden_data: bool,
) -> Result<()> {
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_includes_hidden_data)? == includes_hidden_data {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_chart_includes_hidden_data(data, includes_hidden_data)
    })?;
    if slot.read(package, read_chart_includes_hidden_data)? != includes_hidden_data {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} hidden-data update failed validation"
        )));
    }
    Ok(())
}

fn read_chart_includes_hidden_data(data: &[u8]) -> Result<bool> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(true);
    };
    tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let skips_hidden_data = strict_optional_skip_hidden_data(extension)?.unwrap_or(false);
    Ok(!skips_hidden_data)
}

fn patch_chart_includes_hidden_data(data: &[u8], includes_hidden_data: bool) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        if includes_hidden_data {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultskiphiddendata: Some(true),
            ..Default::default()
        };
        let extension = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(extension.as_slice()),
        )?;
        validate_patched_hidden_data(&patched, includes_hidden_data)?;
        return Ok(patched);
    };

    let field_present = strict_optional_skip_hidden_data(extension)?.is_some();
    let skip_hidden_data = !includes_hidden_data;
    let replacement = (field_present || skip_hidden_data).then_some(u64::from(skip_hidden_data));
    let extension = patch_varint_field(
        extension,
        CHART_SKIP_HIDDEN_DATA_FIELD,
        field_present,
        replacement,
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_hidden_data(&patched, includes_hidden_data)?;
    Ok(patched)
}

fn strict_optional_skip_hidden_data(data: &[u8]) -> Result<Option<bool>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields
        .iter()
        .filter(|field| field.number == CHART_SKIP_HIDDEN_DATA_FIELD);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart skip-hidden-data field {CHART_SKIP_HIDDEN_DATA_FIELD} occurs more than once"
        )));
    }
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart skip-hidden-data field {CHART_SKIP_HIDDEN_DATA_FIELD} is not a varint"
        )));
    }
    let (value, consumed) = crate::varint::decode_varint_from_bytes(
        &data[field.key_end..field.end],
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "chart skip-hidden-data field {CHART_SKIP_HIDDEN_DATA_FIELD} is invalid: {error}"
        ))
    })?;
    if consumed != 1 || consumed != field.end - field.key_end || value > 1 {
        return Err(Error::InvalidFormat(format!(
            "chart skip-hidden-data field {CHART_SKIP_HIDDEN_DATA_FIELD} is not a canonical boolean"
        )));
    }
    Ok(Some(value == 1))
}

fn validate_patched_hidden_data(data: &[u8], expected: bool) -> Result<()> {
    if read_chart_includes_hidden_data(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart hidden-data wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;
    const NON_CANONICAL_ZERO: &[u8] = &[0x80, 0x00];

    #[test]
    fn hidden_data_defaults_included_and_creates_an_extension_when_excluded() {
        let original = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        assert!(read_chart_includes_hidden_data(&original).unwrap());
        assert_eq!(
            patch_chart_includes_hidden_data(&original, true).unwrap(),
            original
        );

        let excluded = patch_chart_includes_hidden_data(&original, false).unwrap();
        assert!(!read_chart_includes_hidden_data(&excluded).unwrap());
        assert!(
            generated_chart_non_style_extension(&excluded)
                .unwrap()
                .is_some()
        );
    }

    #[test]
    fn hidden_data_patch_is_lossless_and_restores_explicit_native_default() {
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultshowlegend: Some(true),
            tschchartinfodefaultskiphiddendata: Some(false),
            ..Default::default()
        };
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let base = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut original = base.encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        let excluded = patch_chart_includes_hidden_data(&original, false).unwrap();
        assert!(!read_chart_includes_hidden_data(&excluded).unwrap());
        assert_eq!(
            raw_field(&excluded, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_non_style_extension(&excluded)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_chart_non_style_extension(&original)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
            )
        );

        let restored = patch_chart_includes_hidden_data(&excluded, true).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn malformed_native_hidden_data_switches_are_rejected() {
        let base = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();

        let mut duplicate = Vec::new();
        append_varint_field(&mut duplicate, CHART_SKIP_HIDDEN_DATA_FIELD, 0).unwrap();
        append_varint_field(&mut duplicate, CHART_SKIP_HIDDEN_DATA_FIELD, 1).unwrap();
        assert!(read_with_extension(&base, &duplicate).is_err());

        let mut wrong_wire = Vec::new();
        append_length_delimited_field(&mut wrong_wire, CHART_SKIP_HIDDEN_DATA_FIELD, &[]).unwrap();
        assert!(read_with_extension(&base, &wrong_wire).is_err());

        let mut non_boolean = Vec::new();
        append_varint_field(&mut non_boolean, CHART_SKIP_HIDDEN_DATA_FIELD, 2).unwrap();
        assert!(read_with_extension(&base, &non_boolean).is_err());

        let mut non_canonical = Vec::new();
        append_varint_field(&mut non_canonical, CHART_SKIP_HIDDEN_DATA_FIELD, 0).unwrap();
        assert_eq!(non_canonical.pop(), Some(0));
        non_canonical.extend_from_slice(NON_CANONICAL_ZERO);
        assert!(read_with_extension(&base, &non_canonical).is_err());
    }

    fn read_with_extension(base: &[u8], extension: &[u8]) -> Result<bool> {
        let data = patch_length_delimited_field(
            base,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(extension),
        )?;
        read_chart_includes_hidden_data(&data)
    }

    fn raw_field(data: &[u8], number: u32) -> Vec<Vec<u8>> {
        parse_wire_fields(data)
            .unwrap()
            .into_iter()
            .filter(|field| field.number == number)
            .map(|field| data[field.start..field.end].to_vec())
            .collect()
    }
}
