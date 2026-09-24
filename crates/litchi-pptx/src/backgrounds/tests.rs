#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::format::ImageFormat;

#[test]
fn test_solid_background() {
    let bg = SlideBackground::solid("FF0000");
    assert!(matches!(bg, SlideBackground::Solid { .. }));
}

#[test]
fn test_gradient_background() {
    let stops = vec![
        GradientStop {
            position: 0.0,
            color: "FF0000".to_string(),
        },
        GradientStop {
            position: 1.0,
            color: "0000FF".to_string(),
        },
    ];
    let bg = SlideBackground::linear_gradient(90.0, stops);
    assert!(matches!(bg, SlideBackground::Gradient { .. }));
}

#[test]
fn test_solid_background_xml() {
    let bg = SlideBackground::solid("FF0000");
    let xml = bg.to_xml(None).unwrap();
    assert!(xml.contains("FF0000"));
    assert!(xml.contains("<a:solidFill>"));

    assert_eq!(
        SlideBackground::from_xml(xml.as_bytes()).unwrap(),
        Some(SlideBackground::solid("FF0000"))
    );
}

#[test]
fn gradient_xml_round_trips_fixed_positions_and_angle() {
    let background = SlideBackground::linear_gradient(
        90.0,
        vec![
            GradientStop {
                position: 0.0,
                color: "112233".to_string(),
            },
            GradientStop {
                position: 0.375,
                color: "AABBCC".to_string(),
            },
            GradientStop {
                position: 1.0,
                color: "FFEEDD".to_string(),
            },
        ],
    );
    let xml = background.to_xml(None).unwrap();

    assert_eq!(
        SlideBackground::from_xml(xml.as_bytes()).unwrap(),
        Some(background)
    );
}

#[test]
fn gradient_xml_round_trips_inclusive_position_and_angle_bounds() {
    let background = SlideBackground::linear_gradient(
        360.0,
        vec![
            GradientStop {
                position: 0.0,
                color: "000000".to_string(),
            },
            GradientStop {
                position: 1.0,
                color: "FFFFFF".to_string(),
            },
        ],
    );

    let xml = background.to_xml(None).unwrap();
    assert!(xml.contains("pos=\"0\""));
    assert!(xml.contains("pos=\"100000\""));
    assert!(xml.contains("ang=\"21600000\""));
    assert_eq!(
        SlideBackground::from_xml(xml.as_bytes()).unwrap(),
        Some(background)
    );
}

#[test]
fn pattern_xml_tokens_and_colors_round_trip() {
    let patterns = [
        (PatternType::Pct5, "pct5"),
        (PatternType::Pct10, "pct10"),
        (PatternType::Pct20, "pct20"),
        (PatternType::Pct25, "pct25"),
        (PatternType::Pct30, "pct30"),
        (PatternType::Pct40, "pct40"),
        (PatternType::Pct50, "pct50"),
        (PatternType::Pct60, "pct60"),
        (PatternType::Pct70, "pct70"),
        (PatternType::Pct75, "pct75"),
        (PatternType::Pct80, "pct80"),
        (PatternType::Pct90, "pct90"),
        (PatternType::Horizontal, "horz"),
        (PatternType::Vertical, "vert"),
        (PatternType::LightHorizontal, "ltHorz"),
        (PatternType::LightVertical, "ltVert"),
        (PatternType::DarkHorizontal, "dkHorz"),
        (PatternType::DarkVertical, "dkVert"),
        (PatternType::NarrowHorizontal, "narHorz"),
        (PatternType::NarrowVertical, "narVert"),
        (PatternType::DashedHorizontal, "dashHorz"),
        (PatternType::DashedVertical, "dashVert"),
        (PatternType::DownDiagonal, "dnDiag"),
        (PatternType::UpDiagonal, "upDiag"),
        (PatternType::LightDownDiagonal, "ltDnDiag"),
        (PatternType::LightUpDiagonal, "ltUpDiag"),
        (PatternType::DarkDownDiagonal, "dkDnDiag"),
        (PatternType::DarkUpDiagonal, "dkUpDiag"),
        (PatternType::WideDownDiagonal, "wdDnDiag"),
        (PatternType::WideUpDiagonal, "wdUpDiag"),
        (PatternType::DashedDownDiagonal, "dashDnDiag"),
        (PatternType::DashedUpDiagonal, "dashUpDiag"),
        (PatternType::Cross, "cross"),
        (PatternType::DiagonalCross, "diagCross"),
        (PatternType::SmallCheck, "smCheck"),
        (PatternType::LargeCheck, "lgCheck"),
        (PatternType::SmallGrid, "smGrid"),
        (PatternType::LargeGrid, "lgGrid"),
        (PatternType::DottedGrid, "dotGrid"),
        (PatternType::SmallConfetti, "smConfetti"),
        (PatternType::LargeConfetti, "lgConfetti"),
        (PatternType::HorizontalBrick, "horzBrick"),
        (PatternType::DiagonalBrick, "diagBrick"),
        (PatternType::SolidDiamond, "solidDmnd"),
        (PatternType::OpenDiamond, "openDmnd"),
        (PatternType::DottedDiamond, "dotDmnd"),
        (PatternType::Plaid, "plaid"),
        (PatternType::Sphere, "sphere"),
        (PatternType::Weave, "weave"),
        (PatternType::Divot, "divot"),
        (PatternType::Shingle, "shingle"),
        (PatternType::Wave, "wave"),
        (PatternType::Trellis, "trellis"),
        (PatternType::ZigZag, "zigZag"),
    ];

    for (pattern, token) in patterns {
        let background =
            SlideBackground::pattern(pattern, "112233".to_string(), "AABBCC".to_string());
        let xml = background.to_xml(None).unwrap();
        assert!(xml.contains(&format!("prst=\"{token}\"")));
        assert!(xml.contains("val=\"112233\""));
        assert!(xml.contains("val=\"AABBCC\""));
        assert_eq!(
            SlideBackground::from_xml(xml.as_bytes()).unwrap(),
            Some(background)
        );
    }
}

#[test]
fn picture_xml_keeps_relationship_and_borrows_image_data() {
    let bytes = vec![0x89, b'P', b'N', b'G'];
    let background = SlideBackground::picture(bytes.clone(), ImageFormat::Png, PictureStyle::Tile);

    let (borrowed, format) = background.image_data().expect("picture image data");
    assert_eq!(borrowed, bytes.as_slice());
    assert_eq!(*format, ImageFormat::Png);

    let xml = background.to_xml(Some("rId7")).unwrap();
    assert!(xml.contains("r:embed=\"rId7\""));
    assert!(xml.contains("<a:tile/>"));
    assert_eq!(SlideBackground::from_xml(xml.as_bytes()).unwrap(), None);
}

#[test]
fn malformed_background_xml_is_rejected() {
    assert!(matches!(
        SlideBackground::from_xml(b"<p:bg><p:bgPr></p:bg>"),
        Err(crate::Error::Xml(_))
    ));
}

/// The number of distinct attribute names in an adversarial start tag, and
/// the number of times its last name then repeats.
const CROWDED: usize = 20_000;

/// `CROWDED - 1` distinct filler attributes, then `name="first"`, then
/// `CROWDED` repeats of `name="repeat"`: a checked attribute iterator that
/// skips its duplicate errors scans every earlier name for each repeat.
fn crowded(name: &str, first: &str, repeat: &str) -> String {
    let mut attributes = String::new();
    for index in 0..CROWDED - 1 {
        attributes.push_str(&format!(" n{index}=\"{index}\""));
    }
    attributes.push_str(&format!(" {name}=\"{first}\""));
    let repeated = format!(" {name}=\"{repeat}\"");
    for _ in 0..CROWDED {
        attributes.push_str(&repeated);
    }
    attributes
}

fn background(fill: &str) -> Option<SlideBackground> {
    SlideBackground::from_xml(format!("<p:bg><p:bgPr>{fill}</p:bgPr></p:bg>").as_bytes()).unwrap()
}

fn gradient(stop: &str, descriptor: &str) -> Option<SlideBackground> {
    background(&format!(
        "<a:gradFill><a:gsLst><a:gs{stop}><a:srgbClr val=\"112233\"/></a:gs></a:gsLst>{descriptor}</a:gradFill>"
    ))
}

fn solid(color: &str) -> Option<SlideBackground> {
    Some(SlideBackground::solid(color))
}

fn pattern(prst: &str) -> Option<SlideBackground> {
    background(&format!(
        "<a:pattFill{prst}><a:fgClr><a:srgbClr val=\"111111\"/></a:fgClr><a:bgClr><a:srgbClr val=\"222222\"/></a:bgClr></a:pattFill>"
    ))
}

fn patterned(pattern_type: PatternType) -> Option<SlideBackground> {
    Some(SlideBackground::Pattern {
        pattern_type,
        fg_color: "111111".to_string(),
        bg_color: "222222".to_string(),
    })
}

fn graded(
    gradient_type: GradientType,
    angle: Option<f64>,
    position: f64,
) -> Option<SlideBackground> {
    Some(SlideBackground::Gradient {
        gradient_type,
        angle,
        stops: vec![GradientStop {
            position,
            color: "112233".to_string(),
        }],
    })
}

#[test]
fn duplicate_background_attributes_keep_their_first_occurrence() {
    // One occurrence of each attribute reads as before.
    let single = background("<a:solidFill><a:srgbClr val=\"112233\"/></a:solidFill>");
    assert_eq!(single, solid("112233"));
    assert_eq!(
        gradient(" pos=\"25000\"", "<a:path path=\"rect\"/>"),
        graded(GradientType::Rectangular, None, 0.25)
    );

    // A later occurrence of a name never replaces the first one.
    assert_eq!(
        background("<a:solidFill><a:srgbClr val=\"112233\" val=\"445566\"/></a:solidFill>"),
        solid("112233")
    );
    assert_eq!(
        background("<a:solidFill><a:schemeClr val=\"accent1\" val=\"accent2\"/></a:solidFill>"),
        solid("accent1")
    );
    assert_eq!(
        pattern(" prst=\"dkVert\" prst=\"cross\""),
        patterned(PatternType::DarkVertical)
    );
    assert_eq!(
        gradient(
            " pos=\"25000\" pos=\"75000\"",
            "<a:lin ang=\"5400000\" ang=\"10800000\" scaled=\"0\"/>"
        ),
        graded(GradientType::Linear, Some(90.0), 0.25)
    );
    assert_eq!(
        gradient(" pos=\"0\"", "<a:path path=\"circle\" path=\"rect\"/>"),
        graded(GradientType::Radial, None, 0.0)
    );
    // Even when the first occurrence's value is unusable.
    assert_eq!(
        gradient(
            " pos=\"x\" pos=\"75000\"",
            "<a:lin ang=\"x\" ang=\"5400000\"/>"
        ),
        graded(GradientType::Linear, None, 0.0)
    );
}

#[test]
fn crowded_background_attributes_are_read_first_wins() {
    // The attribute read comes after every repeated name, or never comes.
    let filler = crowded("x", "first", "repeat");
    assert_eq!(
        background(&format!(
            "<a:solidFill><a:srgbClr{filler} val=\"112233\"/></a:solidFill>"
        )),
        solid("112233")
    );
    assert_eq!(
        background(&format!(
            "<a:solidFill><a:schemeClr{filler} val=\"accent2\"/></a:solidFill>"
        )),
        solid("accent2")
    );
    assert_eq!(
        pattern(&format!("{filler} prst=\"dkVert\"")),
        patterned(PatternType::DarkVertical)
    );
    assert_eq!(pattern(&filler), patterned(PatternType::Pct50));

    // The attribute read is the name that repeats.
    assert_eq!(
        gradient(
            &crowded("pos", "25000", "75000"),
            &format!("<a:lin{}/>", crowded("ang", "5400000", "10800000"))
        ),
        graded(GradientType::Linear, Some(90.0), 0.25)
    );
    assert_eq!(
        gradient(
            " pos=\"0\"",
            &format!("<a:path{}/>", crowded("path", "circle", "rect"))
        ),
        graded(GradientType::Radial, None, 0.0)
    );
}
