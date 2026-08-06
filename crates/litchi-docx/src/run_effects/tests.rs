use super::codec;
use super::*;

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_2010: &str = "http://schemas.microsoft.com/office/word/2010/wordml";

fn all_effects_xml() -> String {
    format!(
        r#"<w:r xmlns:w="{WORD}" xmlns:w14="{WORD_2010}"><w:rPr>
          <w14:glow w14:rad="228600"><w14:srgbClr w14:val="112233"><w14:alpha w14:val="50000"/></w14:srgbClr></w14:glow>
          <w14:shadow w14:blurRad="1000" w14:dist="2" w14:dir="60000"><w14:schemeClr w14:val="accent1"/></w14:shadow>
          <w14:reflection w14:blurRad="3" w14:stA="20000" w14:endA="1000"/>
          <w14:textOutline w14:w="12700" w14:cap="rnd" w14:cmpd="dbl" w14:algn="ctr"><w14:solidFill><w14:srgbClr w14:val="AABBCC"/></w14:solidFill><w14:prstDash w14:val="dash"/><w14:round/></w14:textOutline>
          <w14:textFill><w14:noFill/></w14:textFill>
          <w14:scene3d><w14:camera w14:prst="orthographicFront"/><w14:lightRig w14:rig="threePt" w14:dir="t"><w14:rot w14:lat="0" w14:lon="60000" w14:rev="0"/></w14:lightRig></w14:scene3d>
          <w14:props3d w14:extrusionH="3" w14:prstMaterial="warmMatte"><w14:bevelT w14:w="1" w14:h="2" w14:prst="circle"/><w14:extrusionClr><w14:srgbClr w14:val="010203"/></w14:extrusionClr></w14:props3d>
          <w14:ligatures xmlns:w14="{WORD_2010}" w14:val="standard"/>
        </w:rPr></w:r>"#
    )
}

#[test]
fn typed_effects_round_trip_all_supported_children() {
    let effects = RunEffects::parse(all_effects_xml().as_bytes()).unwrap();
    assert_eq!(effects.len(), 8);
    assert_eq!(effects.glow().unwrap().radius, Some(228600));
    assert_eq!(effects.shadow().unwrap().distance, Some(2));
    assert_eq!(effects.reflection().unwrap().start_alpha, Some(20000));
    assert_eq!(effects.text_outline().unwrap().dash, Some(LineDash::Dash));
    assert_eq!(effects.text_fill().unwrap().fill, Some(Fill::NoFill));
    assert_eq!(
        effects.scene3d().unwrap().camera.preset.as_str(),
        "orthographicFront"
    );
    assert_eq!(effects.props3d().unwrap().extrusion_height, Some(3));
    assert_eq!(effects.unknown().next().unwrap().as_bytes(), br#"<w14:ligatures xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:val="standard"/>"#);

    let mut children = String::new();
    codec::write(&effects, &mut children).unwrap();
    let reparsed = RunEffects::parse(format!("<w:rPr>{children}</w:rPr>").as_bytes()).unwrap();
    assert_eq!(reparsed.len(), effects.len());
    assert_eq!(reparsed.glow(), effects.glow());
    assert_eq!(reparsed.scene3d(), effects.scene3d());
    assert_eq!(
        reparsed.unknown().next().unwrap().as_bytes(),
        effects.unknown().next().unwrap().as_bytes()
    );
}

#[test]
fn malformed_and_duplicate_effects_are_rejected() {
    let duplicate = format!(
        r#"<w:rPr xmlns:w14="{WORD_2010}"><w14:glow><w14:srgbClr w14:val="010203"/></w14:glow><w14:glow><w14:srgbClr w14:val="040506"/></w14:glow></w:rPr>"#
    );
    assert!(RunEffects::parse(duplicate.as_bytes()).is_err());

    let invalid_color = format!(
        r#"<w:rPr xmlns:w14="{WORD_2010}"><w14:glow><w14:srgbClr w14:val="GGGGGG"/></w14:glow></w:rPr>"#
    );
    assert!(RunEffects::parse(invalid_color.as_bytes()).is_err());

    let truncated = format!(r#"<w:rPr xmlns:w14="{WORD_2010}"><w14:future><w14:payload/></w:rPr>"#);
    assert!(RunEffects::parse(truncated.as_bytes()).is_err());

    let opaque =
        OpaqueExtension::new(br#"<w14:future><w14:payload/></w14:future>"#.to_vec()).unwrap();
    let mut effects = RunEffects::new();
    assert!(effects.push_unknown(opaque).is_ok());
    assert!(
        effects
            .push_unknown(OpaqueExtension::new(br#"<w14:broken>"#.to_vec()).unwrap())
            .is_err()
    );
}

#[test]
fn run_path_and_writer_keep_effect_namespace_and_unknown_package_xml() {
    let run = crate::paragraph::Run::new(all_effects_xml().into_bytes());
    let effects = run.run_effects().unwrap();
    match effects.glow().unwrap().color.as_ref().unwrap() {
        Color::Rgb(color) => assert_eq!(color.value, [0x11, 0x22, 0x33]),
        other => panic!("expected RGB glow color, got {other:?}"),
    }

    let mut document = crate::writer::MutableDocument::new();
    let mut authored = RunEffects::new();
    authored
        .set_glow(Some(Glow {
            color: Some(Color::Rgb(RgbColor::new([0xAA, 0xBB, 0xCC]))),
            radius: Some(10),
        }))
        .unwrap();
    document
        .add_paragraph()
        .add_run_with_text("glowing")
        .set_run_effects(authored)
        .unwrap();
    let generated = document.to_xml().unwrap();
    assert!(generated.contains("mc:Ignorable=\"w14\""));
    assert!(generated.contains("w14:glow"));
    assert!(generated.contains("w14:srgbClr"));

    let input = format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{WORD}" xmlns:w14="{WORD_2010}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><w:p><w:r><w:rPr><w14:glow><w14:srgbClr w14:val="010203"/></w14:glow><w14:future xmlns:w14="{WORD_2010}" w14:data="keep"/></w:rPr><w:t>kept</w:t></w:r></w:p></w:body></w:document>"#
    );
    let mut preserved = crate::writer::MutableDocument::from_xml(&input).unwrap();
    preserved.add_paragraph_with_text("appended");
    let output = preserved.to_xml().unwrap();
    assert!(output.contains("w14:future"));
    assert!(output.contains("w14:data=\"keep\""));
    assert!(output.contains("kept"));
    assert!(output.contains("appended"));
}
