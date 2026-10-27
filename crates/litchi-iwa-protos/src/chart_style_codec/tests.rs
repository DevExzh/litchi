//! Unit and Prost-differential tests of the chart style codec (test-only
//! oracle: the generated Prost types).

use prost::Message;

use super::field::{chart_non_style, chart_style};
use super::{
    Archive, EXTENSION_FIELD, EmptyExtension, Error, LimitKind, Limits, Vector, Write, decode,
    references, rewrite,
};
use crate::{tsch, tsp, tss};

type TestResult = Result<(), Box<dyn std::error::Error>>;

const UNKNOWN_OUTER_FIELD: u32 = 4_096;
const UNKNOWN_EXTENSION_FIELD: u32 = 4_097;

fn key(number: u32, wire_type: u8) -> Vec<u8> {
    let mut out = Vec::new();
    super::push_varint(&mut out, (u64::from(number) << 3) | u64::from(wire_type));
    out
}

fn length_delimited(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut out = key(number, 2);
    super::push_varint(&mut out, payload.len() as u64);
    out.extend_from_slice(payload);
    out
}

fn varint_field(number: u32, value: u64) -> Vec<u8> {
    let mut out = key(number, 0);
    super::push_varint(&mut out, value);
    out
}

fn base_non_style() -> Vec<u8> {
    tsch::ChartNonStyleArchive {
        super_: Some(tss::StyleArchive::default()),
    }
    .encode_to_vec()
}

fn non_style_with(generated: &tsch::generated::ChartNonStyleArchive) -> Vec<u8> {
    let mut data = base_non_style();
    data.extend(length_delimited(
        EXTENSION_FIELD,
        &generated.encode_to_vec(),
    ));
    data
}

fn extension_of(data: &[u8]) -> Vec<u8> {
    decode(data, Archive::ChartNonStyle, Limits::DEFAULT)
        .unwrap()
        .extension()
        .unwrap_or_default()
        .to_vec()
}

fn float(field: u32, value: Option<f32>) -> Write<'static> {
    Write::Float { field, value }
}

#[test]
fn archives_round_trip_their_message_types() {
    for archive in Archive::ALL {
        assert_eq!(
            Archive::from_message_type(archive.message_type()),
            Some(archive)
        );
    }
    assert_eq!(Archive::from_message_type(5_021), None);
    assert_eq!(Archive::from_message_type(5_030), None);
}

#[test]
fn creating_an_extension_matches_the_generated_encoder() -> TestResult {
    let base = base_non_style();
    let output = rewrite(
        &base,
        Archive::ChartNonStyle,
        &[float(chart_non_style::PIE_START_ANGLE, Some(123.5))],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    assert!(output.changed());
    let expected = non_style_with(&tsch::generated::ChartNonStyleArchive {
        tschchartinfodefaultpiestartangle: Some(123.5),
        ..Default::default()
    });
    assert_eq!(output.as_bytes(), expected.as_slice());

    // A batch of clears on a payload without an extension is a no-op.
    let unchanged = rewrite(
        &base,
        Archive::ChartNonStyle,
        &[float(chart_non_style::PIE_START_ANGLE, None)],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    assert!(!unchanged.changed());
    assert_eq!(unchanged.as_bytes(), base.as_slice());
    Ok(())
}

#[test]
fn created_fields_are_in_field_number_order() -> TestResult {
    let mut base = tsch::ChartStyleArchive {
        super_: Some(tss::StyleArchive::default()),
    }
    .encode_to_vec();
    base.extend(varint_field(UNKNOWN_OUTER_FIELD, 7));
    let output = rewrite(
        &base,
        Archive::ChartStyle,
        &[
            Write::Bool {
                field: chart_style::ROUNDED_CORNER_OUTER_END_ONLY,
                value: Some(true),
            },
            float(chart_style::ROUNDED_CORNER_RADIUS, Some(0.25)),
        ],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    let generated = tsch::generated::ChartStyleArchive {
        tschchartinfodefaultroundedcornerradius: Some(0.25),
        tschchartinfodefaultroundedcornerouterendonly: Some(true),
        ..Default::default()
    };
    let mut expected = base.clone();
    expected.extend(length_delimited(
        EXTENSION_FIELD,
        &generated.encode_to_vec(),
    ));
    assert_eq!(output.as_bytes(), expected.as_slice());
    Ok(())
}

#[test]
fn present_fields_are_replaced_in_place_and_unknown_bytes_survive() -> TestResult {
    let mut extension = tsch::generated::ChartNonStyleArchive {
        tschchartinfodefaultshowlegend: Some(true),
        tschchartinfodefaultpiestartangle: Some(10.0),
        tschchartinfodefaulttitle: Some("Revenue".to_owned()),
        ..Default::default()
    }
    .encode_to_vec();
    extension.extend(varint_field(UNKNOWN_EXTENSION_FIELD, 42));
    let mut data = base_non_style();
    data.extend(length_delimited(EXTENSION_FIELD, &extension));
    data.extend(varint_field(UNKNOWN_OUTER_FIELD, 9));

    let output = rewrite(
        &data,
        Archive::ChartNonStyle,
        &[float(chart_non_style::PIE_START_ANGLE, Some(90.0))],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    // The angle is fixed32: same length, so only its four bytes change.
    assert_eq!(output.as_bytes().len(), data.len());
    let differing: Vec<usize> = output
        .as_bytes()
        .iter()
        .zip(&data)
        .enumerate()
        .filter(|(_, (a, b))| a != b)
        .map(|(index, _)| index)
        .collect();
    assert!(differing.len() <= 4, "{differing:?}");
    let patched = extension_of(output.as_bytes());
    let decoded = tsch::generated::ChartNonStyleArchive::decode(patched.as_slice())?;
    assert_eq!(decoded.tschchartinfodefaultpiestartangle, Some(90.0));
    assert_eq!(
        decoded.tschchartinfodefaulttitle.as_deref(),
        Some("Revenue")
    );
    assert!(
        output
            .as_bytes()
            .ends_with(&varint_field(UNKNOWN_OUTER_FIELD, 9))
    );
    assert!(patched.ends_with(&varint_field(UNKNOWN_EXTENSION_FIELD, 42)));

    // Setting the same value is an exact no-op; clearing removes the field
    // and restoring re-appends it at the end of the extension.
    let same = rewrite(
        output.as_bytes(),
        Archive::ChartNonStyle,
        &[float(chart_non_style::PIE_START_ANGLE, Some(90.0))],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    assert!(!same.changed());
    let cleared = rewrite(
        &data,
        Archive::ChartNonStyle,
        &[float(chart_non_style::PIE_START_ANGLE, None)],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    let decoded =
        tsch::generated::ChartNonStyleArchive::decode(extension_of(cleared.as_bytes()).as_slice())?;
    assert_eq!(decoded.tschchartinfodefaultpiestartangle, None);
    let restored = rewrite(
        cleared.as_bytes(),
        Archive::ChartNonStyle,
        &[float(chart_non_style::PIE_START_ANGLE, Some(10.0))],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    let restored_extension = extension_of(restored.as_bytes());
    let mut angle = key(chart_non_style::PIE_START_ANGLE, 5);
    angle.extend_from_slice(&10.0_f32.to_le_bytes());
    assert!(restored_extension.ends_with(&angle));
    Ok(())
}

#[test]
fn an_emptied_extension_is_kept() -> TestResult {
    let data = non_style_with(&tsch::generated::ChartNonStyleArchive {
        tschchartinfodefaultinnerradius: Some(0.5),
        ..Default::default()
    });
    let output = rewrite(
        &data,
        Archive::ChartNonStyle,
        &[float(chart_non_style::DONUT_INNER_RADIUS, None)],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    let mut expected = base_non_style();
    expected.extend(length_delimited(EXTENSION_FIELD, &[]));
    assert_eq!(output.as_bytes(), expected.as_slice());
    Ok(())
}

#[test]
fn an_emptied_extension_can_be_removed() -> TestResult {
    let mut data = base_non_style();
    let mut extension = Vec::new();
    super::encode_field(
        &mut extension,
        float(
            super::field::series_non_style::PIE_WEDGE_EXPLOSION,
            Some(0.5),
        ),
    );
    data.extend(length_delimited(EXTENSION_FIELD, &extension));
    let output = rewrite(
        &data,
        Archive::SeriesNonStyle,
        &[float(
            super::field::series_non_style::PIE_WEDGE_EXPLOSION,
            None,
        )],
        EmptyExtension::Remove,
        Limits::DEFAULT,
    )?;
    assert_eq!(output.as_bytes(), base_non_style().as_slice());
    Ok(())
}

#[test]
fn vectors_write_whole_or_by_component() -> TestResult {
    let base = base_non_style();
    let rotation = Vector {
        x: 12.0,
        y: -7.5,
        z: 0.0,
        w: 0.0,
    };
    let created = rewrite(
        &base,
        Archive::ChartNonStyle,
        &[Write::Vector {
            field: chart_non_style::ROTATION_3D,
            value: Some(rotation),
        }],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    let expected = non_style_with(&tsch::generated::ChartNonStyleArchive {
        tschchartinfodefault3drotation: Some(tsch::Chart3DVectorArchive {
            x: 12.0,
            y: -7.5,
            z: 0.0,
            w: 0.0,
        }),
        ..Default::default()
    });
    assert_eq!(created.as_bytes(), expected.as_slice());
    let properties = decode(created.as_bytes(), Archive::ChartNonStyle, Limits::DEFAULT)?;
    assert_eq!(
        properties.vector(chart_non_style::ROTATION_3D)?,
        Some(rotation)
    );

    let component = rewrite(
        created.as_bytes(),
        Archive::ChartNonStyle,
        &[Write::VectorComponents {
            field: chart_non_style::ROTATION_3D,
            components: [None, Some(30.0), None, None],
            absent: None,
        }],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    let decoded = tsch::generated::ChartNonStyleArchive::decode(
        extension_of(component.as_bytes()).as_slice(),
    )?;
    let vector = decoded.tschchartinfodefault3drotation.unwrap();
    assert_eq!(
        (vector.x, vector.y, vector.z, vector.w),
        (12.0, 30.0, 0.0, 0.0)
    );
    assert_eq!(component.as_bytes().len(), created.as_bytes().len());

    // An absent vector without a whole replacement is left absent.
    let untouched = rewrite(
        &base,
        Archive::ChartNonStyle,
        &[Write::VectorComponents {
            field: chart_non_style::PIE_SCALE_3D,
            components: [None, None, Some(1.0), None],
            absent: None,
        }],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    assert!(!untouched.changed());
    Ok(())
}

#[test]
fn scalars_read_like_protobuf() -> TestResult {
    let data = non_style_with(&tsch::generated::ChartNonStyleArchive {
        tschchartinfodefault3dintersetdepthgap: Some(1.5),
        tschchartinfodefaultshowlegend: Some(true),
        ..Default::default()
    });
    let properties = decode(&data, Archive::ChartNonStyle, Limits::DEFAULT)?;
    assert!(properties.has_extension());
    assert_eq!(properties.archive(), Archive::ChartNonStyle);
    assert_eq!(properties.float(chart_non_style::SERIES_GAP_3D)?, Some(1.5));
    assert_eq!(properties.float(chart_non_style::PIE_START_ANGLE)?, None);
    assert_eq!(properties.bool(20)?, Some(true));
    assert!(properties.contains(20)?);
    assert_eq!(
        properties.float(20),
        Err(Error::WireType {
            field: 20,
            expected: 5,
            found: 0
        })
    );

    let mut negative = base_non_style();
    let mut extension = Vec::new();
    super::encode_field(
        &mut extension,
        Write::Int32 {
            field: 21,
            value: Some(-3),
        },
    );
    negative.extend(length_delimited(EXTENSION_FIELD, &extension));
    let properties = decode(&negative, Archive::ChartNonStyle, Limits::DEFAULT)?;
    assert_eq!(properties.int32(21)?, Some(-3));
    let base = base_non_style();
    let without = decode(&base, Archive::ChartNonStyle, Limits::DEFAULT)?;
    assert!(!without.has_extension());
    assert_eq!(without.vector(chart_non_style::ROTATION_3D)?, None);
    Ok(())
}

#[test]
fn malformed_and_duplicate_payloads_are_refused() {
    let mut duplicate_extension = base_non_style();
    duplicate_extension.extend(length_delimited(EXTENSION_FIELD, &[]));
    duplicate_extension.extend(length_delimited(EXTENSION_FIELD, &[]));
    assert_eq!(
        decode(
            &duplicate_extension,
            Archive::ChartNonStyle,
            Limits::DEFAULT
        )
        .err(),
        Some(Error::DuplicateField {
            field: EXTENSION_FIELD
        })
    );

    let mut wrong_type = base_non_style();
    wrong_type.extend(varint_field(EXTENSION_FIELD, 1));
    assert!(matches!(
        decode(&wrong_type, Archive::ChartNonStyle, Limits::DEFAULT),
        Err(Error::WireType { .. })
    ));

    let mut padded = base_non_style();
    padded.extend(key(EXTENSION_FIELD, 2));
    padded.extend([0x80, 0x00]);
    assert!(matches!(
        decode(&padded, Archive::ChartNonStyle, Limits::DEFAULT),
        Err(Error::Malformed { .. })
    ));

    let mut truncated = base_non_style();
    truncated.extend(key(EXTENSION_FIELD, 2));
    truncated.push(5);
    assert!(matches!(
        decode(&truncated, Archive::ChartNonStyle, Limits::DEFAULT),
        Err(Error::Malformed { .. })
    ));

    let mut duplicate_property = Vec::new();
    for _ in 0..2 {
        duplicate_property.extend(key(chart_non_style::PIE_START_ANGLE, 5));
        duplicate_property.extend(1.0_f32.to_le_bytes());
    }
    let mut data = base_non_style();
    data.extend(length_delimited(EXTENSION_FIELD, &duplicate_property));
    assert_eq!(
        rewrite(
            &data,
            Archive::ChartNonStyle,
            &[float(chart_non_style::PIE_START_ANGLE, Some(2.0))],
            EmptyExtension::Keep,
            Limits::DEFAULT,
        ),
        Err(Error::DuplicateField {
            field: chart_non_style::PIE_START_ANGLE
        })
    );
    for field in [0, EXTENSION_FIELD] {
        assert_eq!(
            rewrite(
                &data,
                Archive::ChartNonStyle,
                &[float(field, None)],
                EmptyExtension::Keep,
                Limits::DEFAULT
            ),
            Err(Error::InvalidField { field })
        );
    }
}

#[test]
fn limits_are_checked_and_validated() -> TestResult {
    assert!(matches!(
        Limits::DEFAULT.with_max_bytes(0),
        Err(Error::InvalidLimit {
            kind: LimitKind::Bytes
        })
    ));
    assert!(Limits::DEFAULT.with_max_fields(usize::MAX).is_err());
    let data = non_style_with(&tsch::generated::ChartNonStyleArchive {
        tschchartinfodefaultpiestartangle: Some(1.0),
        ..Default::default()
    });
    let tight = Limits::DEFAULT.with_max_bytes(data.len() - 1)?;
    assert!(matches!(
        decode(&data, Archive::ChartNonStyle, tight),
        Err(Error::Limit {
            kind: LimitKind::Bytes,
            ..
        })
    ));
    let one_field = Limits::DEFAULT.with_max_fields(1)?;
    assert!(matches!(
        decode(&data, Archive::ChartNonStyle, one_field),
        Err(Error::Limit {
            kind: LimitKind::Fields,
            ..
        })
    ));
    let exact = Limits::DEFAULT.with_max_bytes(data.len())?;
    assert!(
        rewrite(
            &data,
            Archive::ChartNonStyle,
            &[float(chart_non_style::DONUT_INNER_RADIUS, Some(0.5))],
            EmptyExtension::Keep,
            exact,
        )
        .is_err()
    );
    Ok(())
}

/// Sequential writes on a generated archive agree with the Prost model.
#[test]
fn write_sequences_agree_with_the_generated_model() -> TestResult {
    let slots = [
        chart_non_style::PIE_START_ANGLE,
        chart_non_style::DONUT_INNER_RADIUS,
        chart_non_style::RADAR_START_ANGLE,
        chart_non_style::SERIES_GAP_3D,
    ];
    let mut state = 0x9e37_79b9_u32;
    let mut next = move || {
        state ^= state << 13;
        state ^= state >> 17;
        state ^= state << 5;
        state
    };
    for _ in 0..200 {
        let mut data = base_non_style();
        let mut model = tsch::generated::ChartNonStyleArchive::default();
        for _ in 0..6 {
            let slot = slots[(next() % 4) as usize];
            let value = (next() % 3 != 0).then(|| (next() % 360) as f32);
            let output = rewrite(
                &data,
                Archive::ChartNonStyle,
                &[float(slot, value)],
                EmptyExtension::Keep,
                Limits::DEFAULT,
            )?;
            data = output.into_bytes().into_owned();
            match slot {
                chart_non_style::PIE_START_ANGLE => model.tschchartinfodefaultpiestartangle = value,
                chart_non_style::DONUT_INNER_RADIUS => {
                    model.tschchartinfodefaultinnerradius = value
                },
                chart_non_style::RADAR_START_ANGLE => {
                    model.tschchartinfodefaultradarstartangle = value;
                },
                _ => model.tschchartinfodefault3dintersetdepthgap = value,
            }
            let decoded =
                tsch::generated::ChartNonStyleArchive::decode(extension_of(&data).as_slice())?;
            assert_eq!(decoded, model);
        }
    }
    Ok(())
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_type: None,
        deprecated_is_external: None,
    }
}

#[test]
fn style_references_match_the_generated_chart() -> TestResult {
    let chart = tsch::ChartArchive {
        chart_style: Some(reference(11)),
        chart_non_style: Some(reference(12)),
        legend_style: Some(reference(13)),
        legend_non_style: Some(reference(0)),
        value_axis_styles: vec![reference(21), reference(22)],
        value_axis_nonstyles: vec![reference(23)],
        category_axis_styles: vec![reference(24)],
        category_axis_nonstyles: vec![reference(25)],
        series_theme_styles: vec![reference(31), reference(32), reference(33)],
        series_private_styles: Some(tsp::SparseReferenceArray {
            count: 3,
            entries: vec![tsp::sparse_reference_array::Entry {
                index: 2,
                reference: reference(41),
            }],
        }),
        series_non_styles: Some(tsp::SparseReferenceArray {
            count: 3,
            entries: vec![
                tsp::sparse_reference_array::Entry {
                    index: 0,
                    reference: reference(51),
                },
                tsp::sparse_reference_array::Entry {
                    index: 1,
                    reference: reference(52),
                },
            ],
        }),
        paragraph_styles: vec![reference(61)],
        ..Default::default()
    };
    let mut drawable = tsch::ChartDrawableArchive::default().encode_to_vec();
    drawable.extend(length_delimited(EXTENSION_FIELD, &chart.encode_to_vec()));
    let projected = references(&drawable, Limits::DEFAULT)?;
    assert_eq!(projected.chart_style(), Some(11));
    assert_eq!(projected.chart_non_style(), Some(12));
    assert_eq!(projected.legend_style(), Some(13));
    assert_eq!(projected.legend_non_style(), None);
    assert_eq!(projected.chart_level(Archive::ChartNonStyle), Some(12));
    assert_eq!(projected.chart_level(Archive::SeriesStyle), None);
    assert_eq!(projected.value_axis_styles(), &[21, 22]);
    assert_eq!(projected.value_axis_non_styles(), &[23]);
    assert_eq!(projected.category_axis_styles(), &[24]);
    assert_eq!(projected.category_axis_non_styles(), &[25]);
    assert_eq!(projected.series_theme_styles(), &[31, 32, 33]);
    assert_eq!(projected.paragraph_styles(), &[61]);
    let private = projected.series_private_styles().unwrap();
    assert_eq!((private.count(), private.entries()), (3, &[(2, 41)][..]));
    let non_styles = projected.series_non_styles().unwrap();
    assert_eq!(non_styles.get(1), Some(52));
    assert_eq!(non_styles.get(2), None);

    assert!(matches!(
        references(
            &tsch::ChartDrawableArchive::default().encode_to_vec(),
            Limits::DEFAULT
        ),
        Err(Error::Malformed { .. })
    ));
    Ok(())
}

#[test]
fn conditional_writes_follow_presence() -> TestResult {
    let base = tsch::ChartStyleArchive {
        super_: Some(tss::StyleArchive::default()),
    }
    .encode_to_vec();
    let radius_only = [
        Write::FloatUnlessAbsentDefault {
            field: chart_style::ROUNDED_CORNER_RADIUS,
            value: 0.5,
            default: 0.0,
        },
        Write::BoolUnlessAbsentDefault {
            field: chart_style::ROUNDED_CORNER_OUTER_END_ONLY,
            value: false,
            default: false,
        },
    ];
    let output = rewrite(
        &base,
        Archive::ChartStyle,
        &radius_only,
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    let generated = tsch::generated::ChartStyleArchive::decode(
        decode(output.as_bytes(), Archive::ChartStyle, Limits::DEFAULT)?
            .extension()
            .unwrap(),
    )?;
    assert_eq!(generated.tschchartinfodefaultroundedcornerradius, Some(0.5));
    assert_eq!(
        generated.tschchartinfodefaultroundedcornerouterendonly,
        None
    );

    // A present default-valued field is still written.
    let present = rewrite(
        output.as_bytes(),
        Archive::ChartStyle,
        &[Write::FloatUnlessAbsentDefault {
            field: chart_style::ROUNDED_CORNER_RADIUS,
            value: 0.0,
            default: 0.0,
        }],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    let properties = decode(present.as_bytes(), Archive::ChartStyle, Limits::DEFAULT)?;
    assert_eq!(
        properties.float(chart_style::ROUNDED_CORNER_RADIUS)?,
        Some(0.0)
    );
    Ok(())
}

#[test]
fn depth_writes_scale_the_stored_denominator() -> TestResult {
    let data = non_style_with(&tsch::generated::ChartNonStyleArchive {
        tschchartinfobar3dscale: Some(tsch::Chart3DVectorArchive {
            x: 1.0,
            y: 3.0,
            z: 0.5,
            w: 0.25,
        }),
        ..Default::default()
    });
    let output = rewrite(
        &data,
        Archive::ChartNonStyle,
        &[Write::VectorDepth {
            field: chart_non_style::BAR_SCALE_3D,
            ratio: 2.0,
            denominator: super::field::vector::Y,
            absent: None,
        }],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    let properties = decode(output.as_bytes(), Archive::ChartNonStyle, Limits::DEFAULT)?;
    assert_eq!(
        properties.vector(chart_non_style::BAR_SCALE_3D)?,
        Some(Vector {
            x: 1.0,
            y: 3.0,
            z: 6.0,
            w: 0.25
        })
    );
    let flat = non_style_with(&tsch::generated::ChartNonStyleArchive {
        tschchartinfopie3dscale: Some(tsch::Chart3DVectorArchive {
            x: 0.0,
            y: 1.0,
            z: 1.0,
            w: 1.0,
        }),
        ..Default::default()
    });
    assert_eq!(
        rewrite(
            &flat,
            Archive::ChartNonStyle,
            &[Write::VectorDepth {
                field: chart_non_style::PIE_SCALE_3D,
                ratio: 1.0,
                denominator: super::field::vector::X,
                absent: None,
            }],
            EmptyExtension::Keep,
            Limits::DEFAULT,
        ),
        Err(Error::InvalidScale {
            field: chart_non_style::PIE_SCALE_3D
        })
    );
    let absent = Vector {
        x: 1.5,
        y: 2.2,
        z: 1.0,
        w: 0.5,
    };
    let base = base_non_style();
    let created = rewrite(
        &base,
        Archive::ChartNonStyle,
        &[Write::VectorDepth {
            field: chart_non_style::PIE_SCALE_3D,
            ratio: 1.0,
            denominator: super::field::vector::X,
            absent: Some(absent),
        }],
        EmptyExtension::Keep,
        Limits::DEFAULT,
    )?;
    let properties = decode(created.as_bytes(), Archive::ChartNonStyle, Limits::DEFAULT)?;
    assert_eq!(
        properties.vector(chart_non_style::PIE_SCALE_3D)?,
        Some(absent)
    );
    Ok(())
}

#[test]
fn property_rules_clear_defaults_and_pick_archives() -> TestResult {
    use super::property;
    let base = base_non_style();
    for write in [
        property::pie_start_angle(0.0),
        property::donut_inner_radius(property::DEFAULT_DONUT_INNER_RADIUS),
        property::radar_start_angle(0.0),
        property::series_gap_3d(property::DEFAULT_SERIES_GAP_3D),
        property::rotation_3d(
            property::DEFAULT_ROTATION_3D.0,
            property::DEFAULT_ROTATION_3D.1,
        ),
        property::depth_3d(property::Scale3d::Pie, 1.0, None),
    ] {
        assert_eq!(write.archive(), Archive::ChartNonStyle);
        let output = rewrite(
            &base,
            write.archive(),
            write.writes(),
            write.empty(),
            Limits::DEFAULT,
        )?;
        assert!(!output.changed(), "{write:?}");
    }
    let angle = property::pie_start_angle(90.0);
    let output = rewrite(
        &base,
        angle.archive(),
        angle.writes(),
        angle.empty(),
        Limits::DEFAULT,
    )?;
    let properties = decode(output.as_bytes(), Archive::ChartNonStyle, Limits::DEFAULT)?;
    assert_eq!(
        properties.float(chart_non_style::PIE_START_ANGLE)?,
        Some(90.0)
    );

    let corners = property::rounded_corners(0.0, false);
    assert_eq!(
        (corners.archive(), corners.writes().len()),
        (Archive::ChartStyle, 2)
    );
    let distance = property::pie_label_distance(property::DEFAULT_PIE_LABEL_DISTANCE);
    assert_eq!(distance.archive(), Archive::SeriesNonStyle);
    assert_eq!(distance.empty(), EmptyExtension::Remove);
    assert_eq!(property::pie_wedge_explosion(0.5).writes().len(), 1);
    assert_eq!(
        property::Scale3d::Bar.field(),
        chart_non_style::BAR_SCALE_3D
    );
    Ok(())
}

#[test]
fn legend_frames_rewrite_in_place() -> TestResult {
    let chart = tsch::ChartArchive {
        legend_frame: Some(tsch::RectArchive {
            origin: tsp::Point { x: 1.0, y: 2.0 },
            size: tsp::Size {
                width: 3.0,
                height: 4.0,
            },
        }),
        ..Default::default()
    };
    let mut drawable = tsch::ChartDrawableArchive::default().encode_to_vec();
    drawable.extend(length_delimited(EXTENSION_FIELD, &chart.encode_to_vec()));
    assert_eq!(
        super::legend_frame(&drawable, Limits::DEFAULT)?,
        Some([1.0, 2.0, 3.0, 4.0])
    );
    let moved =
        super::rewrite_legend_frame(&drawable, Some([5.0, 6.0, 0.0, 8.0]), Limits::DEFAULT)?;
    assert_eq!(moved.as_bytes().len(), drawable.len());
    let mut expected_chart = chart.clone();
    expected_chart.legend_frame = Some(tsch::RectArchive {
        origin: tsp::Point { x: 5.0, y: 6.0 },
        size: tsp::Size {
            width: 0.0,
            height: 8.0,
        },
    });
    let mut expected = tsch::ChartDrawableArchive::default().encode_to_vec();
    expected.extend(length_delimited(
        EXTENSION_FIELD,
        &expected_chart.encode_to_vec(),
    ));
    assert_eq!(moved.as_bytes(), expected.as_slice());
    let removed = super::rewrite_legend_frame(&drawable, None, Limits::DEFAULT)?;
    assert_eq!(
        super::legend_frame(removed.as_bytes(), Limits::DEFAULT)?,
        None
    );
    let restored = super::rewrite_legend_frame(
        removed.as_bytes(),
        Some([1.0, 2.0, 3.0, 4.0]),
        Limits::DEFAULT,
    )?;
    assert_eq!(restored.as_bytes(), drawable.as_slice());
    Ok(())
}

fn drawable_with(chart: &tsch::ChartArchive) -> Vec<u8> {
    let mut drawable = tsch::ChartDrawableArchive::default().encode_to_vec();
    drawable.extend(length_delimited(EXTENSION_FIELD, &chart.encode_to_vec()));
    drawable
}

fn sparse(entries: &[(u32, u64)]) -> tsp::SparseReferenceArray {
    tsp::SparseReferenceArray {
        count: entries.len() as u32,
        entries: entries
            .iter()
            .map(|&(index, identifier)| tsp::sparse_reference_array::Entry {
                index,
                reference: tsp::Reference {
                    identifier,
                    ..Default::default()
                },
            })
            .collect(),
    }
}

#[test]
fn series_non_styles_match_the_generated_encoder() -> TestResult {
    let chart = tsch::ChartArchive {
        series_theme_styles: vec![tsp::Reference {
            identifier: 40,
            ..Default::default()
        }],
        paragraph_styles: vec![tsp::Reference {
            identifier: 41,
            ..Default::default()
        }],
        ..Default::default()
    };
    let drawable = drawable_with(&chart);
    // Absent: inserted before the first later field (paragraph styles).
    let added = super::rewrite_series_non_styles(&drawable, &[(0, 7), (2, 9)], Limits::DEFAULT)?;
    let mut expected = chart.clone();
    expected.series_non_styles = Some(sparse(&[(0, 7), (2, 9)]));
    assert_eq!(added.as_bytes(), drawable_with(&expected).as_slice());
    // Present: replaced in place, an empty array keeps its count.
    let emptied = super::rewrite_series_non_styles(added.as_bytes(), &[], Limits::DEFAULT)?;
    expected.series_non_styles = Some(sparse(&[]));
    assert_eq!(emptied.as_bytes(), drawable_with(&expected).as_slice());
    let unchanged = super::rewrite_series_non_styles(emptied.as_bytes(), &[], Limits::DEFAULT)?;
    assert!(!unchanged.changed());
    assert!(matches!(
        super::rewrite_series_non_styles(&drawable, &[(0, 0)], Limits::DEFAULT),
        Err(Error::InvalidField { .. })
    ));
    Ok(())
}

#[test]
fn reference_lines_rewrite_in_place_and_in_field_order() -> TestResult {
    let chart = tsch::ChartArchive {
        series_theme_styles: vec![tsp::Reference {
            identifier: 40,
            ..Default::default()
        }],
        ..Default::default()
    };
    let mut drawable = tsch::ChartDrawableArchive::default().encode_to_vec();
    let mut chart_bytes = chart.encode_to_vec();
    // A later chart extension the insertion must precede.
    chart_bytes.extend(varint_field(10_006, 1));
    drawable.extend(length_delimited(EXTENSION_FIELD, &chart_bytes));
    assert_eq!(super::reference_lines(&drawable, Limits::DEFAULT)?, None);
    let payload = [0x0a, 0x02, 0x08, 0x02];
    let added = super::rewrite_reference_lines(&drawable, Some(&payload), Limits::DEFAULT)?;
    let mut expected_chart = chart.encode_to_vec();
    expected_chart.extend(length_delimited(super::REFERENCE_LINES_FIELD, &payload));
    expected_chart.extend(varint_field(10_006, 1));
    let mut expected = tsch::ChartDrawableArchive::default().encode_to_vec();
    expected.extend(length_delimited(EXTENSION_FIELD, &expected_chart));
    assert_eq!(added.as_bytes(), expected.as_slice());
    assert_eq!(
        super::reference_lines(added.as_bytes(), Limits::DEFAULT)?,
        Some(payload.as_slice())
    );
    let removed = super::rewrite_reference_lines(added.as_bytes(), None, Limits::DEFAULT)?;
    assert_eq!(removed.as_bytes(), drawable.as_slice());
    Ok(())
}
