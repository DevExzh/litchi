#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use std::borrow::Cow;

use litchi_rtf::{
    PictureShapeProperties, RtfDocument, RtfWriter, Shape, ShapeProperty, ShapeThemeColor,
    ShapeThemeValue, ShapeType,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_real_libreoffice_theme_metadata_and_round_trips_canonically() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/watermark.rtf"
    );
    let producer = RtfDocument::parse_bytes(source).unwrap();
    let property = producer
        .shapes()
        .iter()
        .chain(
            producer
                .sections()
                .iter()
                .flat_map(|section| &section.headers_footers)
                .flat_map(|header_footer| &header_footer.shapes),
        )
        .flat_map(|shape| shape.properties.iter())
        .find(|property| property.name == "fillColor" && property.theme_value.is_some())
        .unwrap();
    assert_eq!(property.value, "4626167");
    assert_eq!(
        property.theme_value,
        Some(ShapeThemeValue {
            color: ShapeThemeColor::Accent6,
            tint: 255,
            shade: 255,
        })
    );

    let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    let mut shape = Shape::new(ShapeType::Rectangle);
    shape.properties.push(ShapeProperty::new_themed(
        Cow::Borrowed(property.name.as_ref()),
        Cow::Borrowed(property.value.as_ref()),
        property.theme_value.unwrap(),
    ));
    document.set_background_shape(shape).unwrap();
    let output = write(&document);
    assert!(String::from_utf8(output.clone()).unwrap().contains(
        "{\\sp{\\sn fillColor}{\\sv 4626167}{\\*\\hsv\\caccentsix\\ctint255\\cshade255}}"
    ));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    let reparsed_property = reparsed
        .background_shape()
        .unwrap()
        .properties
        .iter()
        .find(|property| property.name == "fillColor")
        .unwrap();
    assert_eq!(reparsed_property.theme_value, property.theme_value);
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn typed_picture_property_theme_metadata_round_trips() {
    let mut document =
        RtfDocument::parse(r"{\rtf1{\*\shppict{\pict\pngblip\picw1\pich1 89504e470d0a1a0a}}}")
            .unwrap();
    document
        .set_picture_shape_properties(
            0,
            Some(PictureShapeProperties {
                shape_id: None,
                properties: vec![ShapeProperty::new_themed(
                    Cow::Borrowed("lineColor"),
                    Cow::Borrowed("16711680"),
                    ShapeThemeValue {
                        color: ShapeThemeColor::Accent1,
                        tint: 255,
                        shade: 191,
                    },
                )],
            }),
        )
        .unwrap();

    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reparsed.pictures()[0]
            .shape_properties
            .as_ref()
            .unwrap()
            .properties[0]
            .theme_value,
        document.pictures()[0]
            .shape_properties
            .as_ref()
            .unwrap()
            .properties[0]
            .theme_value
    );
}

#[test]
fn rejects_hostile_hsv_grammar() {
    for source in [
        r"{\rtf1{\*\hsv\caccentone\ctint255\cshade255}}",
        r"{\rtf1{\hsv\caccentone\ctint255\cshade255}}",
        r"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\*\hsv\caccentone\ctint255\cshade255}{\sv 1}}}\pngblip 89504e470d0a1a0a}}",
        r"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}{\hsv\caccentone\ctint255\cshade255}}}\pngblip 89504e470d0a1a0a}}",
        r"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}{\*\hsv1\caccentone\ctint255\cshade255}}}\pngblip 89504e470d0a1a0a}}",
        r"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}{\*\hsv\ctint255\cshade255}}}\pngblip 89504e470d0a1a0a}}",
        r"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}{\*\hsv\caccentone\cshade255}}}\pngblip 89504e470d0a1a0a}}",
        r"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}{\*\hsv\caccentone\ctint255}}}\pngblip 89504e470d0a1a0a}}",
        r"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}{\*\hsv\caccentone\caccenttwo\ctint255\cshade255}}}\pngblip 89504e470d0a1a0a}}",
        r"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}{\*\hsv\caccentone\ctint256\cshade255}}}\pngblip 89504e470d0a1a0a}}",
        r"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}{\*\hsv\caccentone\ctint128\cshade128}}}\pngblip 89504e470d0a1a0a}}",
        r"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}{\*\hsv\caccentone\ctint255\cshade255{\object}}}}\pngblip 89504e470d0a1a0a}}",
        r"{\rtf1{\pict{\*\picprop{\sp{\sn x}{\sv 1}{\*\hsv\caccentone\ctint255\cshade255}{\*\hsv\caccenttwo\ctint255\cshade255}}}\pngblip 89504e470d0a1a0a}}",
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn enforces_typed_theme_invariants() {
    let invalid_theme = ShapeThemeValue {
        color: ShapeThemeColor::Accent3,
        tint: 128,
        shade: 128,
    };
    assert!(invalid_theme.validate().is_err());
    assert!(
        ShapeProperty::new_themed(
            Cow::Borrowed("fillColor"),
            Cow::Borrowed(""),
            ShapeThemeValue {
                color: ShapeThemeColor::Accent3,
                tint: 255,
                shade: 255,
            },
        )
        .validate()
        .is_err()
    );

    let mut binary = ShapeProperty::new_binary(Cow::Borrowed("x"), Cow::Borrowed(&[1]));
    binary.theme_value = Some(ShapeThemeValue {
        color: ShapeThemeColor::Accent3,
        tint: 255,
        shade: 255,
    });
    assert!(binary.validate().is_err());
}
