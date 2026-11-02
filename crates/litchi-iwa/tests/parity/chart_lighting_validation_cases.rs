//! Regression (record 0414): a 3D lighting read refuses every stored
//! package the host's `validate_native_package` refuses (an unnamed light,
//! a non-finite or negative intensity, a non-finite color, an unknown
//! coordinate space, a light with two sources), on all three formats.

use litchi_iwa::charts::Chart3dLightingStyle;
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::wire::transform_length_delimited_fields_at_path;
use litchi_iwa_protos::chart_style_codec::{self as codec, Limits};

use crate::chart_style_cases::{Components, Input, Target, input, publish};
use crate::harness::{Format, TestResult};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const CHART_STYLE_MESSAGE_TYPE: u32 = 5_022;
/// Generated extension, column lighting package, its lights.
const LIGHTS: &[u32] = &[10_000, 4, 2];

/// Fields appended to every light; the last occurrence wins (or, for a
/// submessage, merges), as `prost` decodes.
fn mutations() -> Vec<(&'static str, Vec<u8>)> {
    let float = |key: u8, value: f32| {
        let mut bytes = vec![key];
        bytes.extend_from_slice(&value.to_le_bytes());
        bytes
    };
    let mut ambient_nan = vec![0x12, 5];
    ambient_nan.extend(float(0x0d, f32::NAN));
    vec![
        ("unnamed light", vec![0x0a, 0x00]),
        ("NaN intensity", float(0x2d, f32::NAN)),
        ("negative intensity", float(0x2d, -1.0)),
        ("infinite intensity", float(0x2d, f32::INFINITY)),
        ("NaN ambient color", ambient_nan),
        ("coordinate space 2", vec![0x38, 0x02]),
        // A point and a spot source together: at least two sources.
        ("two sources", vec![0x4a, 0x00, 0x5a, 0x00]),
    ]
}

fn host_read(input: &Input, bytes: &[u8]) -> TestResult<Chart3dLightingStyle> {
    Ok(match input.target {
        Target::Keynote { slide } => {
            keynote::Draft::from_bytes(bytes)?.slide_chart_3d_lighting_style(slide, input.chart)?
        },
        Target::Pages => {
            pages::Draft::from_bytes(bytes)?.body_chart_3d_lighting_style(input.chart)?
        },
        Target::Numbers { sheet } => {
            numbers::Draft::from_bytes(bytes)?.sheet_chart_3d_lighting_style(sheet, input.chart)?
        },
    })
}

fn leaf_read(format: Format, bytes: &[u8]) -> TestResult<String> {
    Ok(match format {
        Format::Keynote => {
            use litchi_keynote::chart::style::{Property, PropertyKind};
            match litchi_keynote::Package::from_bytes(bytes)?.slide_chart_style(
                0,
                1,
                PropertyKind::Lighting3d,
            )? {
                Property::Lighting3d(v) => v.inspector_label().to_owned(),
                _ => return Err("not a lighting value".into()),
            }
        },
        Format::Pages => {
            use litchi_pages::chart::style::{Property, PropertyKind};
            match litchi_pages::Package::from_bytes(bytes)?
                .body_chart_style(0, PropertyKind::Lighting3d)?
            {
                Property::Lighting3d(v) => v.inspector_label().to_owned(),
                _ => return Err("not a lighting value".into()),
            }
        },
        Format::Numbers => {
            use litchi_numbers::sheet::chart::style::{Property, PropertyKind};
            match litchi_numbers::Package::from_bytes(bytes)?.sheet_chart_style(
                0,
                1,
                PropertyKind::Lighting3d,
            )? {
                Property::Lighting3d(v) => v.inspector_label().to_owned(),
                _ => return Err("not a lighting value".into()),
            }
        },
    })
}

/// A host-added 3D column chart with the Glossy preset.
fn glossy(format: Format) -> TestResult<(Input, Vec<u8>)> {
    let input = input(format, Kind::Column3d)?;
    let bytes = match input.target {
        Target::Keynote { slide } => {
            let mut e = keynote::Draft::from_bytes(&input.fixture.bytes)?;
            e.set_slide_chart_3d_lighting_style(slide, input.chart, Chart3dLightingStyle::Glossy)?;
            e.to_bytes()?
        },
        Target::Pages => {
            let mut e = pages::Draft::from_bytes(&input.fixture.bytes)?;
            e.set_body_chart_3d_lighting_style(input.chart, Chart3dLightingStyle::Glossy)?;
            e.to_bytes()?
        },
        Target::Numbers { sheet } => {
            let mut e = numbers::Draft::from_bytes(&input.fixture.bytes)?;
            e.set_sheet_chart_3d_lighting_style(sheet, input.chart, Chart3dLightingStyle::Glossy)?;
            e.to_bytes()?
        },
    };
    Ok((input, bytes))
}

#[test]
fn chart_lighting_invalid_package_refuses() -> TestResult {
    for format in [Format::Numbers, Format::Pages, Format::Keynote] {
        let (input, source) = glossy(format)?;
        assert_eq!(host_read(&input, &source)?, Chart3dLightingStyle::Glossy);
        assert_eq!(leaf_read(format, &source)?, "Glossy");
        for (name, suffix) in mutations() {
            let components = Components::read(&source)?;
            let location = components.locate(input.chart, CHART_MESSAGE_TYPE)?;
            let style =
                codec::references(components.payload(input.chart, location)?, Limits::DEFAULT)
                    .map_err(|error| error.to_string())?
                    .chart_style()
                    .ok_or("no chart style")?;
            let style_location = components.locate(style, CHART_STYLE_MESSAGE_TYPE)?;
            let rewritten = transform_length_delimited_fields_at_path::<_, litchi_iwa_common::Error>(
                components.payload(style, style_location)?,
                LIGHTS,
                |light| {
                    let mut light = light.to_vec();
                    light.extend_from_slice(&suffix);
                    Ok(light)
                },
            )?;
            let bytes = publish(
                &source,
                components,
                vec![(style, CHART_STYLE_MESSAGE_TYPE, rewritten)],
            )?;
            assert!(
                host_read(&input, &bytes).is_err(),
                "{format:?} {name}: the host accepts the package"
            );
            assert!(
                leaf_read(format, &bytes).is_err(),
                "{format:?} {name}: the leaf accepts what the host refuses"
            );
        }
    }
    Ok(())
}
