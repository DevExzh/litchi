#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use crate::Error;

use super::*;

const POI_AUDIO: &[u8] =
    include_bytes!("../../../../test-data/poi/test-data/slideshow/EmbeddedAudio.pptx");
const POI_VIDEO: &[u8] =
    include_bytes!("../../../../test-data/poi/test-data/slideshow/EmbeddedVideo.pptx");

fn extension() -> Extension {
    Extension {
        embed_relationship_id: Some("rIdMedia".into()),
        link_relationship_id: None,
        trim: Some(Trim {
            start: Some(Offset::parse("1.5s").unwrap()),
            end: Some(Offset::ms(250)),
        }),
        fade: Some(Fade {
            fade_in: Some(Offset::ms(1000)),
            fade_out: Some(Offset::secs(2)),
        }),
        bookmarks: vec![Bookmark {
            name: Some("chapter".into()),
            time: Some(Offset::secs(3)),
        }],
        extensions: None,
    }
}
fn value() -> List {
    List {
        pictures: vec![Picture {
            shape_id: 4,
            name: "sample.mp4".into(),
            kind: Kind::Video,
            relationship_id: "rIdVideo".into(),
            resource: Some(Resource::new(
                "/ppt/media/media1.mp4",
                "video/mp4",
                vec![0, 1, 2, 3],
            )),
            poster: Some(Poster {
                relationship_id: "rIdPoster".into(),
                resource: Some(Resource::new(
                    "/ppt/media/image1.png",
                    "image/png",
                    vec![137, 80, 78, 71],
                )),
            }),
            transform: Some(Transform::emu(100, 200, 300, 400).unwrap()),
            office_extension: Some(extension()),
        }],
    }
}
fn slide_xml(conformance: Conformance, inside: &[u8]) -> Vec<u8> {
    [format!("<p:sld xmlns:p=\"{}\" xmlns:a=\"{}\" xmlns:r=\"{}\"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>", conformance.pml(), conformance.dml(), conformance.rel()).as_bytes(), inside, b"</p:spTree></p:cSld></p:sld>"].concat()
}
fn test_package(conformance: Conformance) -> (OpcPackage, PackURI) {
    let mut package = OpcPackage::new();
    let uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        uri.clone(),
        ct::PML_SLIDE.into(),
        slide_xml(conformance, b""),
    )));
    (package, uri)
}

#[test]
fn strict_xml_round_trip_covers_typed_extension_properties() {
    let expected = value();
    let fragment = write_pictures(&expected, Conformance::Strict).unwrap();
    let xml = std::str::from_utf8(&fragment).unwrap();
    assert!(xml.contains(&format!(r#"xmlns:r="{STRICT_REL}""#)));
    assert!(xml.contains(&format!(r#"<p14:media xmlns:r="{REL}" r:embed="rIdMedia""#)));
    let parsed = parse(&slide_xml(Conformance::Strict, &fragment)).unwrap();
    assert_eq!(parsed.pictures[0].shape_id, 4);
    assert_eq!(
        parsed.pictures[0]
            .transform
            .as_ref()
            .unwrap()
            .width()
            .as_emu(),
        300
    );
    let extension = parsed.pictures[0].office_extension.as_ref().unwrap();
    assert_eq!(
        extension.trim.as_ref().unwrap().start,
        Some(Offset::ms(1500))
    );
    assert_eq!(extension.bookmarks[0].name.as_deref(), Some("chapter"));
    assert!(parsed.pictures[0].resource.is_none());
}

#[test]
fn media_transform_round_trips_coordinate_offsets_and_integer_extents() {
    let mut expected = value();
    expected.pictures[0].transform = Some(Transform::new(
        Coordinate::parse("-1.25cm").unwrap(),
        Coordinate::emu(litchi_drawingml::coordinate::MIN_EMU).unwrap(),
        Extent::ZERO,
        Extent::emu(litchi_drawingml::coordinate::MAX_EMU).unwrap(),
    ));

    let fragment = write_pictures(&expected, Conformance::Transitional).unwrap();
    let xml = std::str::from_utf8(&fragment).unwrap();
    assert!(xml.contains(r#"<a:off x="-1.25cm" y="-27273042329600"/>"#));
    assert!(xml.contains(r#"<a:ext cx="0" cy="27273042316900"/>"#));

    let parsed = parse(&slide_xml(Conformance::Transitional, &fragment)).unwrap();
    assert_eq!(parsed.pictures[0].transform, expected.pictures[0].transform);
}

#[test]
fn media_transform_construction_and_parsing_enforce_boundaries() {
    assert!(
        Transform::emu(
            litchi_drawingml::coordinate::MIN_EMU,
            litchi_drawingml::coordinate::MAX_EMU,
            1,
            litchi_drawingml::coordinate::MAX_EMU,
        )
        .is_ok()
    );
    assert!(Transform::emu(litchi_drawingml::coordinate::MIN_EMU - 1, 0, 1, 1).is_err());
    assert!(Transform::emu(0, 0, 0, 1).is_ok());
    assert!(Transform::emu(0, 0, 1, -1).is_err());
    assert!(Transform::emu(0, 0, litchi_drawingml::coordinate::MAX_EMU + 1, 1).is_err());

    let fragment = write_pictures(&value(), Conformance::Transitional).unwrap();
    let invalid = String::from_utf8(fragment)
        .unwrap()
        .replace(r#"cx="300""#, r#"cx="0mm""#);
    assert!(parse(&slide_xml(Conformance::Transitional, invalid.as_bytes(),)).is_err());
}

#[test]
fn trim_and_fade_preserve_absence_and_explicit_zero() {
    let mut expected = value();
    let (authored_trim, authored_fade) = {
        let extension = expected.pictures[0].office_extension.as_mut().unwrap();
        extension.trim = Some(Trim {
            start: Some(Offset::ZERO),
            end: None,
        });
        extension.fade = Some(Fade {
            fade_in: None,
            fade_out: Some(Offset::ZERO),
        });
        (extension.trim.clone(), extension.fade.clone())
    };

    let fragment = write_pictures(&expected, Conformance::Transitional).unwrap();
    let xml = std::str::from_utf8(&fragment).unwrap();
    assert!(xml.contains(r#"<p14:trim st="0"/>"#));
    assert!(xml.contains(r#"<p14:fade out="0"/>"#));

    let parsed = parse(&slide_xml(Conformance::Transitional, &fragment)).unwrap();
    let actual = parsed.pictures[0].office_extension.as_ref().unwrap();
    assert_eq!(actual.trim, authored_trim);
    assert_eq!(actual.fade, authored_fade);
    assert!(actual.trim.as_ref().unwrap().start().is_zero());
    assert!(actual.trim.as_ref().unwrap().end().is_zero());
    assert!(actual.fade.as_ref().unwrap().fade_in().is_zero());
    assert!(actual.fade.as_ref().unwrap().fade_out().is_zero());
}

#[test]
fn opaque_media_extensions_round_trip_canonically() {
    let opaque = ExtensionList::parse(
            format!(
                r#"<p:extLst xmlns:p="{PML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:z="urn:example" mc:Ignorable="z"><p:ext uri="{{EXAMPLE}}"><z:data z:flag="a&amp;b">before<![CDATA[<literal>]]><z:inner/>after &amp; done</z:data></p:ext></p:extLst>"#
            )
            .as_bytes(),
        )
        .unwrap();
    assert!(opaque.as_str().contains("before&lt;literal&gt;<z:inner"));
    assert!(opaque.as_str().contains("after &amp; done"));

    let mut expected = value();
    expected.pictures[0]
        .office_extension
        .as_mut()
        .unwrap()
        .extensions = Some(opaque.clone());
    let fragment = write_pictures(&expected, Conformance::Strict).unwrap();
    let xml = std::str::from_utf8(&fragment).unwrap();
    assert!(xml.contains(&format!(r#"xmlns:p="{PML}""#)));
    assert!(xml.contains(r#"xmlns:z="urn:example""#));

    let parsed = parse(&slide_xml(Conformance::Strict, &fragment)).unwrap();
    assert_eq!(
        parsed.pictures[0]
            .office_extension
            .as_ref()
            .unwrap()
            .extensions
            .as_ref(),
        Some(&opaque)
    );
}

#[test]
fn rejects_duplicate_and_misordered_media_children() {
    fn picture(children: &str) -> Vec<u8> {
        format!(
                r#"<p:pic xmlns:p14="{P14}"><p:nvPicPr><p:cNvPr id="1"/><p:nvPr><a:audioFile r:link="rId1"/><p:extLst><p:ext><p14:media xmlns:r="{REL}" r:embed="rId2">{children}</p14:media></p:ext></p:extLst></p:nvPr></p:nvPicPr></p:pic>"#
            )
            .into_bytes()
    }

    for children in [
        "<p14:trim/><p14:trim/>",
        "<p14:fade/><p14:trim/>",
        "<p:extLst/><p14:bmkLst/>",
    ] {
        assert!(
            parse(&slide_xml(Conformance::Transitional, &picture(children),)).is_err(),
            "accepted invalid media children: {children}"
        );
    }
}

#[test]
fn opaque_media_extension_constructor_is_bounded_and_typed() {
    assert!(ExtensionList::parse(b"<not-an-extension-list/>").is_err());
    assert!(ExtensionList::parse(&vec![b' '; MAX_MEDIA_EXTENSION_XML_BYTES + 1]).is_err());
}

#[test]
fn mce_fallback_selects_supported_media_picture() {
    let fragment = write_pictures(&value(), Conformance::Transitional).unwrap();
    let alternate = [b"<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x=\"urn:unsupported\"><mc:Choice Requires=\"x\"><p:pic/></mc:Choice><mc:Fallback>".as_slice(), fragment.as_slice(), b"</mc:Fallback></mc:AlternateContent>"].concat();
    assert_eq!(
        parse(&slide_xml(Conformance::Transitional, &alternate))
            .unwrap()
            .pictures
            .len(),
        1
    );
}

#[test]
fn loads_poi_audio_and_video_resources_without_decoding() {
    for (bytes, kind, size) in [
        (POI_AUDIO, Kind::Audio, 52_079usize),
        (POI_VIDEO, Kind::Video, 101_799usize),
    ] {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        let uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let media = load(&package, &uri).unwrap();
        assert_eq!(media.pictures.len(), 1);
        let picture = &media.pictures[0];
        assert_eq!(picture.kind, kind);
        assert_eq!(picture.resource.as_ref().unwrap().data.len(), size);
        assert!(
            picture
                .poster
                .as_ref()
                .unwrap()
                .resource
                .as_ref()
                .unwrap()
                .content_type
                .starts_with("image/")
        );
        assert!(
            picture
                .office_extension
                .as_ref()
                .unwrap()
                .embed_relationship_id
                .is_some()
        );
    }
}

#[test]
fn transitional_package_writer_round_trips_complete_graph() {
    let (mut package, uri) = test_package(Conformance::Transitional);
    let expected = value();
    store(&mut package, &uri, &expected, Conformance::Transitional).unwrap();
    assert_eq!(load(&package, &uri).unwrap(), expected);
}

#[test]
fn repeated_media_targets_share_one_immutable_payload_allocation() {
    let (mut package, uri) = test_package(Conformance::Transitional);
    let mut expected = value();
    let mut second = expected.pictures[0].clone();
    second.shape_id = 5;
    second.name = "sample-copy.mp4".into();
    second.relationship_id = "rIdVideo2".into();
    second.poster.as_mut().unwrap().relationship_id = "rIdPoster2".into();
    second
        .office_extension
        .as_mut()
        .unwrap()
        .embed_relationship_id = Some("rIdMedia2".into());
    expected.pictures.push(second);

    store(&mut package, &uri, &expected, Conformance::Transitional).unwrap();
    let loaded = load(&package, &uri).unwrap();
    let first = &loaded.pictures[0].resource.as_ref().unwrap().data;
    let second = &loaded.pictures[1].resource.as_ref().unwrap().data;
    assert!(first.shares_with(second));
}

#[test]
fn rejects_malformed_markup_caps_and_package_graphs() {
    let malformed = format!(
        "<p:sld xmlns:p=\"{PML}\" xmlns:a=\"{DML}\" xmlns:r=\"{REL}\"><p:pic><p:nvPicPr><p:cNvPr id=\"1\"/><p:nvPr><a:audioFile r:link=\"rId1\"/><p:extLst><p:ext><p14:media xmlns:p14=\"{P14}\" r:embed=\"rId2\"><p14:trim st=\"1..2s\"/></p14:media></p:ext></p:extLst></p:nvPr></p:nvPicPr></p:pic></p:sld>"
    );
    assert!(parse(malformed.as_bytes()).is_err());
    assert!(parse(b"<!DOCTYPE x><p:sld/>").is_err());
    assert!(parse(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
    let (mut package, uri) = test_package(Conformance::Transitional);
    let mut expected = value();
    for picture in &mut expected.pictures {
        picture.resource = None;
        if let Some(poster) = picture.poster.as_mut() {
            poster.resource = None;
        }
    }
    let fragment = write_pictures(&expected, Conformance::Transitional).unwrap();
    package
        .get_part_mut(&uri)
        .unwrap()
        .set_blob(slide_xml(Conformance::Transitional, &fragment));
    assert!(load(&package, &uri).is_err());
}

#[test]
fn bookmark_uniqueness_compares_represented_time() {
    let mut value = value();
    let extension = value.pictures[0].office_extension.as_mut().unwrap();
    extension.bookmarks = vec![
        Bookmark {
            name: Some("first".into()),
            time: Some(Offset::parse("1s").unwrap()),
        },
        Bookmark {
            name: Some("second".into()),
            time: Some(Offset::parse("1000ms").unwrap()),
        },
    ];

    assert!(write_pictures(&value, Conformance::Transitional).is_err());
}

#[test]
fn non_image_resource_requires_typed_media_kind() {
    let value = value();
    let resource = value.pictures[0].resource.as_ref().unwrap();

    assert!(matches!(
        resource_uri(resource, false, None),
        Err(Error::Invalid(message)) if message.contains("requires a media kind")
    ));
}

#[test]
fn escaped_media_output_is_preflighted_against_a_typed_budget() {
    let mut value = value();
    value.pictures[0]
        .office_extension
        .as_mut()
        .unwrap()
        .bookmarks[0]
        .name = Some("\"".repeat(1_024));
    validate_value(&value, false).unwrap();

    let maximum = 2_048;
    let mut output = BoundedXml::with_limit(maximum);
    let error =
        write_picture(&mut output, &value.pictures[0], Conformance::Transitional).unwrap_err();

    assert!(matches!(
        error,
        Error::Limit {
            resource: "slide media serialized XML bytes",
            limit,
        } if limit == maximum
    ));
    assert!(output.bytes.len() <= maximum);
}

#[test]
fn media_output_budget_is_aggregate_across_pictures() {
    let picture = value().pictures.pop().unwrap();
    let mut single = BoundedXml::new();
    write_picture(&mut single, &picture, Conformance::Transitional).unwrap();
    let single_len = single.bytes.len();
    let maximum = single_len + single_len / 2;

    let mut output = BoundedXml::with_limit(maximum);
    write_picture(&mut output, &picture, Conformance::Transitional).unwrap();
    let error = write_picture(&mut output, &picture, Conformance::Transitional).unwrap_err();

    assert!(matches!(
        error,
        Error::Limit {
            resource: "slide media serialized XML bytes",
            limit,
        } if limit == maximum
    ));
    assert!(output.bytes.len() >= single_len);
    assert!(output.bytes.len() <= maximum);
}
