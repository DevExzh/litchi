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
