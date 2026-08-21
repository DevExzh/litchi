//! Lossless native chart-legend shadow storage and mutation.
//!
//! Legend shadow is an independently inherited `TSD.ShadowArchive` in the
//! native legend-style extension. The legend inspector supports drop shadows
//! only, so contact and curved shadows are rejected rather than normalized.

use prost::Message;

use crate::charts::legend_style::{GENERATED_LEGEND_STYLE_EXTENSION_FIELD, legend_style_slot};
use crate::protobuf::tsd;
use crate::shapes::{Drop, Shadow, shadow_from_native, shadow_to_native};
use crate::wire::{WireField, parse_wire_fields, patch_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

/// `tschlegendmodeldefaultshadow` in `TSCH.Generated.LegendStyleArchive`.
const LEGEND_SHADOW_FIELD: u32 = 4;

const SHADOW_FIELDS: &[StrictFieldSpec] = &[
    StrictFieldSpec {
        number: 1,
        kind: StrictFieldKind::Message(COLOR_FIELDS),
    },
    StrictFieldSpec {
        number: 2,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 3,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 4,
        kind: StrictFieldKind::Varint,
    },
    StrictFieldSpec {
        number: 5,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 6,
        kind: StrictFieldKind::Bool,
    },
    StrictFieldSpec {
        number: 7,
        kind: StrictFieldKind::Varint,
    },
    StrictFieldSpec {
        number: 8,
        kind: StrictFieldKind::Message(EMPTY_FIELDS),
    },
    StrictFieldSpec {
        number: 9,
        kind: StrictFieldKind::Message(CONTACT_SHADOW_FIELDS),
    },
    StrictFieldSpec {
        number: 10,
        kind: StrictFieldKind::Message(CURVED_SHADOW_FIELDS),
    },
];

const COLOR_FIELDS: &[StrictFieldSpec] = &[
    StrictFieldSpec {
        number: 1,
        kind: StrictFieldKind::Varint,
    },
    StrictFieldSpec {
        number: 3,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 4,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 5,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 6,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 7,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 8,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 9,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 10,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 11,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 12,
        kind: StrictFieldKind::Varint,
    },
];

const EMPTY_FIELDS: &[StrictFieldSpec] = &[];

const CONTACT_SHADOW_FIELDS: &[StrictFieldSpec] = &[
    StrictFieldSpec {
        number: 2,
        kind: StrictFieldKind::Fixed32,
    },
    StrictFieldSpec {
        number: 4,
        kind: StrictFieldKind::Fixed32,
    },
];

const CURVED_SHADOW_FIELDS: &[StrictFieldSpec] = &[StrictFieldSpec {
    number: 1,
    kind: StrictFieldKind::Fixed32,
}];

#[derive(Clone, Copy)]
struct StrictFieldSpec {
    number: u32,
    kind: StrictFieldKind,
}

#[derive(Clone, Copy)]
enum StrictFieldKind {
    Varint,
    Bool,
    Fixed32,
    Message(&'static [StrictFieldSpec]),
}

/// Exact direct shadow state for a native chart legend.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ChartLegendShadow {
    /// No direct override; iWork resolves the legend style's parent chain.
    #[default]
    Inherited,
    /// A direct disabled shadow (the inspector checkbox is off).
    NoShadow,
    /// A direct typed drop shadow.
    Shadow(Drop),
}

/// Read the exact direct legend-shadow state of one native chart.
pub(crate) fn chart_legend_shadow(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartLegendShadow> {
    legend_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_legend_shadow)
}

/// Set or remove the direct legend-shadow override of one native chart.
pub(crate) fn set_chart_legend_shadow(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    shadow: ChartLegendShadow,
) -> Result<()> {
    let mut slot = legend_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_legend_shadow)? == shadow {
        return Ok(());
    }
    slot.ensure_exclusive(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    slot.update(package, |data| patch_legend_shadow(data, shadow))?;
    slot.collapse_if_equivalent(package, chart_archive_name, drawable_object_id)?;
    if chart_legend_shadow(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )? != shadow
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} legend shadow update failed validation"
        )));
    }
    Ok(())
}

fn read_legend_shadow(data: &[u8]) -> Result<ChartLegendShadow> {
    let Some(extension) = strict_optional_message(
        data,
        GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
        "legend style extension",
        EMPTY_FIELDS,
    )?
    else {
        return Ok(ChartLegendShadow::Inherited);
    };
    let Some(shadow_payload) = strict_optional_message(
        extension,
        LEGEND_SHADOW_FIELD,
        "legend shadow",
        SHADOW_FIELDS,
    )?
    else {
        return Ok(ChartLegendShadow::Inherited);
    };
    let native = tsd::ShadowArchive::decode(shadow_payload)?;
    match shadow_from_native(&native)? {
        Shadow::Disabled => Ok(ChartLegendShadow::NoShadow),
        Shadow::Drop(shadow) => Ok(ChartLegendShadow::Shadow(shadow)),
        Shadow::Contact(_) | Shadow::Curved(_) => Err(Error::InvalidFormat(
            "native chart legend uses a non-drop shadow".to_owned(),
        )),
    }
}

fn patch_legend_shadow(data: &[u8], shadow: ChartLegendShadow) -> Result<Vec<u8>> {
    let extension = strict_optional_message(
        data,
        GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
        "legend style extension",
        EMPTY_FIELDS,
    )?;
    let Some(extension) = extension else {
        let ChartLegendShadow::Inherited = shadow else {
            let native = native_shadow_bytes(shadow)?;
            let extension = patch_length_delimited_field(
                &[],
                LEGEND_SHADOW_FIELD,
                false,
                Some(native.as_slice()),
            )?;
            let patched = patch_length_delimited_field(
                data,
                GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
                false,
                Some(extension.as_slice()),
            )?;
            validate_patched_legend_shadow(&patched, shadow)?;
            return Ok(patched);
        };
        return Ok(data.to_vec());
    };

    let existing_shadow = strict_optional_message(
        extension,
        LEGEND_SHADOW_FIELD,
        "legend shadow",
        SHADOW_FIELDS,
    )?;
    let native = match shadow {
        ChartLegendShadow::Inherited => None,
        ChartLegendShadow::NoShadow | ChartLegendShadow::Shadow(_) => {
            let replacement = native_shadow_bytes(shadow)?;
            Some(if let Some(existing) = existing_shadow {
                merge_strict_message(existing, &replacement, SHADOW_FIELDS, "legend shadow")?
            } else {
                replacement
            })
        },
    };
    let extension = patch_length_delimited_field(
        extension,
        LEGEND_SHADOW_FIELD,
        existing_shadow.is_some(),
        native.as_deref(),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_legend_shadow(&patched, shadow)?;
    Ok(patched)
}

fn native_shadow_bytes(shadow: ChartLegendShadow) -> Result<Vec<u8>> {
    let native = match shadow {
        ChartLegendShadow::Inherited => {
            return Err(Error::InvalidFormat(
                "inherited legend shadow has no native payload".to_owned(),
            ));
        },
        ChartLegendShadow::NoShadow => shadow_to_native(Shadow::Disabled),
        ChartLegendShadow::Shadow(shadow) => shadow_to_native(Shadow::Drop(shadow)),
    };
    Ok(native.encode_to_vec())
}

/// Return one schema-selected message while leaving all other fields opaque.
///
/// The generated compatibility type used to decode the entire legend-style
/// extension here.  That eagerly materialized unrelated fill/stroke payloads
/// and silently accepted duplicate or wrong-wire shadow fields.  The private
/// projection below owns only the small shadow schema; unknown fields remain
/// source bytes and are never decoded or re-encoded as generated values.
fn strict_optional_message<'a>(
    data: &'a [u8],
    field_number: u32,
    label: &str,
    schema: &'static [StrictFieldSpec],
) -> Result<Option<&'a [u8]>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields
        .iter()
        .copied()
        .filter(|field| field.number() == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{label} field {field_number} occurs more than once"
        )));
    }
    if field.wire_type() != 2 {
        return Err(Error::InvalidFormat(format!(
            "{label} field {field_number} is not length-delimited"
        )));
    }
    field.validate_canonical_framing(data)?;
    let payload = field.checked_payload(data)?;
    validate_strict_message(payload, schema, label)?;
    Ok(Some(payload))
}

fn validate_strict_message(data: &[u8], schema: &[StrictFieldSpec], label: &str) -> Result<()> {
    let fields = parse_wire_fields(data)?;
    for (index, field) in fields.iter().enumerate() {
        let Some(spec) = schema.iter().find(|spec| spec.number == field.number()) else {
            continue;
        };
        if fields[..index]
            .iter()
            .any(|candidate| candidate.number() == field.number())
        {
            return Err(Error::InvalidFormat(format!(
                "{label} field {} occurs more than once",
                field.number()
            )));
        }
        let expected_wire_type = match spec.kind {
            StrictFieldKind::Varint | StrictFieldKind::Bool => 0,
            StrictFieldKind::Fixed32 => 5,
            StrictFieldKind::Message(_) => 2,
        };
        if field.wire_type() != expected_wire_type {
            return Err(Error::InvalidFormat(format!(
                "{label} field {} has the wrong wire type",
                field.number()
            )));
        }
        field.validate_canonical_framing(data)?;
        match spec.kind {
            StrictFieldKind::Varint => validate_canonical_varint(data, *field, label)?,
            StrictFieldKind::Bool => {
                let value = canonical_varint(data, *field, label)?;
                if value > 1 {
                    return Err(Error::InvalidFormat(format!(
                        "{label} bool field {} must be 0 or 1",
                        field.number()
                    )));
                }
            },
            StrictFieldKind::Fixed32 => {},
            StrictFieldKind::Message(nested) => {
                validate_strict_message(field.checked_payload(data)?, nested, label)?;
            },
        }
    }
    Ok(())
}

fn validate_canonical_varint(data: &[u8], field: WireField, label: &str) -> Result<()> {
    canonical_varint(data, field, label).map(|_| ())
}

fn canonical_varint(data: &[u8], field: WireField, label: &str) -> Result<u64> {
    let payload = field.checked_payload(data)?;
    let (value, consumed) =
        litchi_iwa_common::varint::decode_varint_from_bytes(payload).map_err(|error| {
            Error::InvalidFormat(format!(
                "{label} field {} has an invalid varint: {error}",
                field.number()
            ))
        })?;
    if consumed != payload.len() || consumed != litchi_iwa_common::varint::encoded_len(value) {
        return Err(Error::InvalidFormat(format!(
            "{label} field {} has a noncanonical varint",
            field.number()
        )));
    }
    Ok(value)
}

/// Replace recognized fields in one message while copying every unknown field
/// from the existing source.  Message-valued fields recurse with their own
/// small schema, so unknown bytes inside the shadow color/family archives are
/// retained as well.
fn merge_strict_message(
    existing: &[u8],
    replacement: &[u8],
    schema: &[StrictFieldSpec],
    label: &str,
) -> Result<Vec<u8>> {
    validate_strict_message(existing, schema, label)?;
    validate_strict_message(replacement, schema, label)?;
    let existing_fields = parse_wire_fields(existing)?;
    let replacement_fields = parse_wire_fields(replacement)?;
    let capacity = existing
        .len()
        .checked_add(replacement.len())
        .ok_or_else(|| Error::InvalidFormat("legend shadow wire output overflow".to_owned()))?;
    let mut output = Vec::new();
    output.try_reserve(capacity).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "legend shadow wire output",
            amount: capacity,
        })
    })?;

    for existing_field in &existing_fields {
        let Some(spec) = schema
            .iter()
            .find(|spec| spec.number == existing_field.number())
        else {
            output.extend_from_slice(existing_field.raw(existing)?);
            continue;
        };
        let replacement_field = replacement_fields
            .iter()
            .find(|field| field.number() == existing_field.number());
        let Some(replacement_field) = replacement_field else {
            continue;
        };
        let replacement_bytes = match spec.kind {
            StrictFieldKind::Message(nested) => {
                let merged = merge_strict_message(
                    existing_field.checked_payload(existing)?,
                    replacement_field.checked_payload(replacement)?,
                    nested,
                    label,
                )?;
                encode_replacement_message(replacement, *replacement_field, &merged)?
            },
            StrictFieldKind::Varint | StrictFieldKind::Bool | StrictFieldKind::Fixed32 => {
                replacement_field.raw(replacement)?.to_vec()
            },
        };
        output.extend_from_slice(&replacement_bytes);
    }

    for replacement_field in &replacement_fields {
        let already_present = existing_fields
            .iter()
            .any(|field| field.number() == replacement_field.number());
        if !already_present {
            output.extend_from_slice(replacement_field.raw(replacement)?);
        }
    }
    Ok(output)
}

fn encode_replacement_message(source: &[u8], field: WireField, payload: &[u8]) -> Result<Vec<u8>> {
    let key = field.key(source)?;
    let mut encoded = Vec::new();
    encoded
        .try_reserve(key.len().saturating_add(payload.len()))
        .map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "legend shadow nested wire output",
                amount: key.len().saturating_add(payload.len()),
            })
        })?;
    encoded.extend_from_slice(key);
    let mut length = [0_u8; litchi_iwa_common::varint::MAX_BYTES];
    let length = litchi_iwa_common::varint::encode_varint_to_buffer(
        u64::try_from(payload.len())
            .map_err(|_| Error::InvalidFormat("legend shadow payload exceeds u64".to_owned()))?,
        &mut length,
    );
    encoded.extend_from_slice(length);
    encoded.extend_from_slice(payload);
    Ok(encoded)
}

fn validate_patched_legend_shadow(data: &[u8], expected: ChartLegendShadow) -> Result<()> {
    if read_legend_shadow(data)? != expected {
        return Err(Error::InvalidFormat(
            "Chart legend shadow wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::legend_style::generated_legend_style_extension;
    use crate::protobuf::{tsch, tss};
    use crate::shapes::{
        Appearance, BlurRadius, Offset, Pattern, RgbColorSpace, RgbaColor, ShapeFill, Stroke,
        Width, fill_to_native, stroke_to_native,
    };
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};
    use litchi_iwa_common::shape::shadow::{Angle, Opacity};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_SHADOW_FIELD: u32 = 4_098;
    const UNMAPPED_COLOR_FIELD: u32 = 4_099;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn legend_shadow_is_exact_and_preserves_fill_stroke_and_unknown_fields() {
        let fill =
            ShapeFill::Solid(RgbaColor::new(0.9, 0.95, 1.0, 1.0, RgbColorSpace::Srgb).unwrap());
        let stroke = Stroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            Width::new(2.5).unwrap(),
            Pattern::MediumDash,
        );
        let mut generated = tsch::generated::LegendStyleArchive {
            tschlegendmodeldefaultfill: Some(fill_to_native(&fill)),
            tschlegendmodeldefaultopacity: Some(0.8),
            tschlegendmodeldefaultstroke: Some(stroke_to_native(stroke)),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut generated, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut original = tsch::LegendStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        assert_eq!(
            read_legend_shadow(&original).unwrap(),
            ChartLegendShadow::Inherited
        );
        let shadow = ChartLegendShadow::Shadow(Drop::new(
            Appearance::new(
                RgbaColor::black(),
                BlurRadius::from_points(12).unwrap(),
                Offset::from_points(8.0).unwrap(),
                Opacity::new(0.6).unwrap(),
            ),
            Angle::from_degrees(30.0).unwrap(),
        ));
        let shadowed = patch_legend_shadow(&original, shadow).unwrap();
        assert_eq!(read_legend_shadow(&shadowed).unwrap(), shadow);
        let generated_after = generated_legend_style_extension(&shadowed)
            .unwrap()
            .unwrap();
        let decoded = tsch::generated::LegendStyleArchive::decode(generated_after).unwrap();
        assert_eq!(
            decoded.tschlegendmodeldefaultfill,
            Some(fill_to_native(&fill))
        );
        assert_eq!(
            decoded.tschlegendmodeldefaultstroke,
            Some(stroke_to_native(stroke))
        );
        assert_eq!(decoded.tschlegendmodeldefaultopacity, Some(0.8));
        assert!(
            parse_wire_fields(&shadowed)
                .unwrap()
                .iter()
                .any(|field| field.number() == UNMAPPED_OUTER_FIELD && field.wire_type() == 0)
        );
        assert!(
            parse_wire_fields(generated_after)
                .unwrap()
                .iter()
                .any(|field| field.number() == UNMAPPED_GENERATED_FIELD && field.wire_type() == 0)
        );

        let disabled = patch_legend_shadow(&shadowed, ChartLegendShadow::NoShadow).unwrap();
        assert_eq!(
            read_legend_shadow(&disabled).unwrap(),
            ChartLegendShadow::NoShadow
        );
        let inherited = patch_legend_shadow(&disabled, ChartLegendShadow::Inherited).unwrap();
        assert_eq!(inherited, original);
    }

    #[test]
    fn strict_shadow_projection_rejects_duplicate_or_wrong_wire_fields() {
        let mut duplicate = Vec::new();
        append_varint_field(&mut duplicate, LEGEND_SHADOW_FIELD, 1).unwrap();
        append_varint_field(&mut duplicate, LEGEND_SHADOW_FIELD, 2).unwrap();
        let mut data = Vec::new();
        append_length_delimited_field(
            &mut data,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            &duplicate,
        )
        .unwrap();
        assert!(read_legend_shadow(&data).is_err());

        let mut wrong_wire = Vec::new();
        append_varint_field(&mut wrong_wire, 2, 1).unwrap();
        let mut data = Vec::new();
        append_length_delimited_field(
            &mut data,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            &patch_length_delimited_field(&[], LEGEND_SHADOW_FIELD, false, Some(&wrong_wire))
                .unwrap(),
        )
        .unwrap();
        assert!(read_legend_shadow(&data).is_err());

        // Unselected generated fields remain opaque. A malformed field 1
        // must not block a shadow read or rewrite, and its original bytes
        // must survive the focused patch.
        let mut opaque_extension = Vec::new();
        append_varint_field(&mut opaque_extension, 1, 77).unwrap();
        let mut opaque_data = Vec::new();
        append_length_delimited_field(
            &mut opaque_data,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            &opaque_extension,
        )
        .unwrap();
        assert_eq!(
            read_legend_shadow(&opaque_data).unwrap(),
            ChartLegendShadow::Inherited
        );
        let patched = patch_legend_shadow(&opaque_data, ChartLegendShadow::NoShadow).unwrap();
        assert_eq!(
            read_legend_shadow(&patched).unwrap(),
            ChartLegendShadow::NoShadow
        );
        let extension = strict_optional_message(
            &patched,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            "legend style extension",
            EMPTY_FIELDS,
        )
        .unwrap()
        .unwrap();
        assert!(
            parse_wire_fields(extension)
                .unwrap()
                .iter()
                .any(|field| field.number() == 1 && field.wire_type() == 0)
        );
    }

    #[test]
    fn shadow_patch_preserves_unknown_shadow_and_color_bytes() {
        let original_shadow = Shadow::Drop(Drop::new(
            Appearance::new(
                RgbaColor::black(),
                BlurRadius::from_points(8).unwrap(),
                Offset::from_points(4.0).unwrap(),
                Opacity::new(0.4).unwrap(),
            ),
            Angle::from_degrees(15.0).unwrap(),
        ));
        let native = shadow_to_native(original_shadow);
        let mut color = native.color.as_ref().unwrap().encode_to_vec();
        append_varint_field(&mut color, UNMAPPED_COLOR_FIELD, UNMAPPED_VALUE).unwrap();
        let shadow =
            patch_length_delimited_field(&native.encode_to_vec(), 1, true, Some(&color)).unwrap();
        let mut shadow = shadow;
        append_varint_field(&mut shadow, UNMAPPED_SHADOW_FIELD, UNMAPPED_VALUE).unwrap();

        let mut extension = Vec::new();
        append_length_delimited_field(&mut extension, LEGEND_SHADOW_FIELD, &shadow).unwrap();
        let mut original = Vec::new();
        append_length_delimited_field(
            &mut original,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        let replacement = ChartLegendShadow::Shadow(Drop::new(
            Appearance::new(
                RgbaColor::new(0.2, 0.3, 0.4, 0.8, RgbColorSpace::Srgb).unwrap(),
                BlurRadius::from_points(12).unwrap(),
                Offset::from_points(6.0).unwrap(),
                Opacity::new(0.7).unwrap(),
            ),
            Angle::from_degrees(45.0).unwrap(),
        ));
        assert_eq!(
            read_legend_shadow(&original).unwrap(),
            ChartLegendShadow::Shadow(match original_shadow {
                Shadow::Drop(shadow) => shadow,
                _ => unreachable!(),
            })
        );
        let patched = patch_legend_shadow(&original, replacement).unwrap();
        let extension = strict_optional_message(
            &patched,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            "legend style extension",
            EMPTY_FIELDS,
        )
        .unwrap()
        .unwrap();
        let shadow = strict_optional_message(
            extension,
            LEGEND_SHADOW_FIELD,
            "legend shadow",
            SHADOW_FIELDS,
        )
        .unwrap()
        .unwrap();
        assert!(
            parse_wire_fields(shadow)
                .unwrap()
                .iter()
                .any(|field| field.number() == UNMAPPED_SHADOW_FIELD)
        );
        let color = parse_wire_fields(shadow)
            .unwrap()
            .into_iter()
            .find(|field| field.number() == 1)
            .unwrap()
            .checked_payload(shadow)
            .unwrap();
        assert!(
            parse_wire_fields(color)
                .unwrap()
                .iter()
                .any(|field| field.number() == UNMAPPED_COLOR_FIELD)
        );
        assert_eq!(read_legend_shadow(&patched).unwrap(), replacement);
    }
}
