#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use std::borrow::Cow;

use litchi_rtf::raw::Tree;
use litchi_rtf::{EmbeddedObject, ObjectKind, ObjectResultKind, write::Writer};

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn round_trips_real_libreoffice_inline_ole_object() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/ole-inline.rtf"
    );
    let document = Tree::parse_bytes(source, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(document.objects().len(), 1);
    let object = &document.objects()[0];
    assert_eq!(object.kind, ObjectKind::Embedded);
    assert!(!object.class_name.is_empty());
    assert!(!object.data.is_empty());
    assert!(document.text().get(..object.position).is_some());

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("{\\object\\objemb"));
    assert!(serialized.contains("{\\*\\objclass "));
    assert!(serialized.contains("{\\*\\objdata "));

    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.objects().len(), 1);
    let reparsed_object = &reparsed.objects()[0];
    assert_eq!(reparsed_object.position, object.position);
    assert_eq!(reparsed_object.kind, object.kind);
    assert_eq!(reparsed_object.class_name, object.class_name);
    assert_eq!(reparsed_object.width, object.width);
    assert_eq!(reparsed_object.height, object.height);
    assert_eq!(reparsed_object.data, object.data);
}

#[test]
fn preserves_typed_object_metadata_and_shared_result_picture_indices() {
    let source = r"{\rtf1\ansi A{\object\objocx\linkself\objlock\objupdate
        {\*\objclass Control.Class}{\*\objname Visible Name}\objsetsize
        \objalign20\objtransy-3\objh40\objw50\objcropt1\objcropb2
        \objcropl3\objcropr4\objscalex125\objscaley80\rsltmerge\rslthtml
        {\*\oleclsid 00112233-4455-6677-8899-AABBCCDDEEFF}
        {\*\objdata 0102a0ff}
        {\result fallback{\pict\pngblip\picw1\pich1 89504e470d0a1a0a}}}B}";
    let document = Tree::parse(source, litchi_rtf::read::OpenOptions::default()).unwrap();
    let object = &document.objects()[0];
    assert_eq!(object.position, 1);
    assert_eq!(object.kind, ObjectKind::OleControl);
    assert!(object.link_self);
    assert_eq!(object.class_id, "00112233-4455-6677-8899-AABBCCDDEEFF");
    assert_eq!(object.alignment, Some(20));
    assert_eq!(object.translation_y, Some(-3));
    assert_eq!(object.crop_top, Some(1));
    assert_eq!(object.crop_bottom, Some(2));
    assert_eq!(object.crop_left, Some(3));
    assert_eq!(object.crop_right, Some(4));
    assert_eq!(object.scale_x, Some(125));
    assert_eq!(object.scale_y, Some(80));
    assert!(object.merge_result);
    assert_eq!(object.result_kind, Some(ObjectResultKind::Html));
    assert_eq!(object.data, [1, 2, 0xa0, 0xff]);
    assert_eq!(object.result_picture_indices.len(), 1);
    let picture_index = object.result_picture_indices[0];
    assert!(std::ptr::eq(
        document.picture_for_object_result(object, 0).unwrap(),
        &raw const document.pictures()[picture_index]
    ));

    let output = write(&document);
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    let reparsed_object = &reparsed.objects()[0];
    assert_eq!(reparsed_object.position, 1);
    assert_eq!(reparsed_object.kind, ObjectKind::OleControl);
    assert_eq!(reparsed_object.class_id, object.class_id);
    assert_eq!(reparsed_object.alignment, Some(20));
    assert_eq!(reparsed_object.result_kind, Some(ObjectResultKind::Html));
    assert_eq!(reparsed_object.data, object.data);
    assert_eq!(reparsed_object.result_text, "fallback");
    assert_eq!(reparsed_object.result_picture_indices.len(), 1);
}

#[test]
fn nested_object_result_groups_preserve_decoding_scope_and_picture_wrappers() {
    use litchi_rtf::PictureCompatibilityKind::{NonShapePicture, ShapePicture};
    let source = br"{\rtf1\ansi\deff0{\fonttbl{\f0\fcharset0 Arial;}{\f1\fcharset128 MS Gothic;}}A{\object\objemb{\*\objclass Package}{\*\objdata 00}{\result{\f1 \'82\'a0}\'e9{\uc0\u945}\u946?{\b {\*\shppict{\pict\pngblip 89504e470d0a1a0a}}{\nonshppict{\pict\wmetafile8 89504e470d0a1a0a}}}}}B}";
    let package = litchi_rtf::Package::from_bytes(source).unwrap();
    let document = package.raw();
    assert_eq!(document.text(), "AB");
    assert_eq!(document.objects().len(), 1);
    let object = &document.objects()[0];
    assert_eq!(object.result_text, "あéαβ");
    assert_eq!(object.result_picture_indices, [0, 1]);
    assert_eq!(
        object.result_picture_wrappers,
        [Some(ShapePicture), Some(NonShapePicture)]
    );
    assert!(document.picture_compatibility_records().is_empty());
    let bytes = write(document);
    let output = String::from_utf8(bytes.clone()).unwrap();
    assert_eq!(output.matches("\\shppict").count(), 1);
    assert_eq!(output.matches("\\nonshppict").count(), 1);
    let reopened = litchi_rtf::Package::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.text(), "AB");
    assert_eq!(reopened.raw().objects()[0].result_text, object.result_text);
    assert_eq!(
        reopened.raw().objects()[0].result_picture_wrappers,
        object.result_picture_wrappers
    );
    assert_eq!(reopened.raw().pictures(), document.pictures());
}

#[test]
fn unknown_object_result_destinations_remain_inert_and_block_lossy_writing() {
    let package = litchi_rtf::Package::from_bytes(br"{\rtf1 A{\object\objemb{\*\objdata 00}{\result Visible{\*\unknownObjectResult never activate}}}B}").unwrap();
    assert_eq!(package.text(), "AB");
    assert_eq!(package.raw().objects()[0].result_text, "Visible");
    let opaque = package.raw().opaque_nodes();
    assert_eq!(opaque.len(), 1);
    assert!(matches!(
        opaque[0].anchor(),
        litchi_rtf::opaque::Anchor::Structural {
            context: litchi_rtf::opaque::Context::Drawing,
            ..
        }
    ));
    assert_eq!(
        opaque[0].source(),
        br"{\*\unknownObjectResult never activate}"
    );
    let mut output = Vec::new();
    let error = Writer::new(&mut output)
        .write_document(package.raw())
        .unwrap_err();
    assert_eq!(error.kind(), std::io::ErrorKind::InvalidInput);
    assert!(output.is_empty());
}

#[test]
fn malformed_or_deep_object_result_wrappers_are_rejected() {
    for result in [
        r"{\shppict{\pict\pngblip 00}}".to_owned(),
        r"{\*\nonshppict{\pict\pngblip 00}}".to_owned(),
        r"{\*\shppict1{\pict\pngblip 00}}".to_owned(),
        r"{\*\shppict{\pict\pngblip 00}{\pict\pngblip 00}}".to_owned(),
        r"{\*\shppict{\pict\pngblip}}".to_owned(),
        format!("{}text{}", r"{\b ".repeat(40), "}".repeat(40)),
    ] {
        let source = format!(r"{{\rtf1{{\object\objemb{{\*\objdata 00}}{{\result {result}}}}}}}");
        assert!(
            matches!(
                litchi_rtf::Package::from_bytes(source.as_bytes()),
                Err(litchi_rtf::Error::MalformedDocument(_))
            ),
            "accepted {result}"
        );
    }
}

#[test]
fn typed_mutation_validates_positions_and_clear_keeps_shared_pictures() {
    let mut document = Tree::parse(
        r"{\rtf1 Body{\object\objemb{\*\objdata 00}{\result{\pict\pngblip\picw1\pich1 89504e470d0a1a0a}}}}", litchi_rtf::read::OpenOptions::default()
    )
    .unwrap();
    assert_eq!(document.pictures().len(), 1);
    document.clear_objects();
    assert_eq!(document.pictures().len(), 1);

    let mut object = EmbeddedObject::new();
    object.position = 2;
    object.kind = ObjectKind::Publisher;
    object.class_name = Cow::Borrowed("Publisher.Class");
    object.data = vec![0x12, 0x34];
    object.result_picture_indices.push(0);
    document.push_object(object.clone()).unwrap();
    assert!(
        document
            .picture_for_object_result(&document.objects()[0], 0)
            .is_some()
    );

    let mut invalid = object;
    invalid.position = 99;
    assert!(document.push_object(invalid).is_err());

    let output = write(&document);
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.objects()[0].position, 2);
    assert_eq!(reparsed.objects()[0].kind, ObjectKind::Publisher);
    assert_eq!(reparsed.objects()[0].result_picture_indices.len(), 1);

    document.clear_objects();
    assert_eq!(document.pictures().len(), 1);
    assert!(
        !String::from_utf8(write(&document))
            .unwrap()
            .contains("\\object")
    );
}

#[test]
fn rejects_hostile_object_destination_grammar() {
    for source in [
        r"{\rtf1{\*\object\objemb{\*\objdata 00}}}",
        r"{\rtf1{\object1\objemb{\*\objdata 00}}}",
        r"{\rtf1{\object\objemb1{\*\objdata 00}}}",
        r"{\rtf1{\object\objemb\objalign{\*\objdata 00}}}",
        r"{\rtf1{\object\objemb\objalign1\objalign2{\*\objdata 00}}}",
        r"{\rtf1{\object\objemb{\objclass X}{\*\objdata 00}}}",
        r"{\rtf1{\object\objemb{\*\oleclsid1 X}{\*\objdata 00}}}",
        r"{\rtf1{\object\objemb{\*\oleclsid X}{\*\oleclsid Y}{\*\objdata 00}}}",
        r"{\rtf1{\object\objemb{\*\objdata 00}{\*\oleclsid X}}}",
        r"{\rtf1{\object\objemb{\*\objdata 00}{\*\result x}}}",
        r"{\rtf1{\object\objemb{\*\objclass X{\object}}{\*\objdata 00}}}",
        r"{\rtf1\intbl{\object\objemb{\*\objdata 00}}}",
    ] {
        assert!(
            Tree::parse(source, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn round_trips_object_alias_and_section_metadata() {
    let source = r"{\rtf1\ansi A{\object\objemb\objh40\objw50
        {\*\objclass Package}{\*\objname Visible Name}
        {\*\objalias Alias Name}{\*\objsect Section1}
        {\*\oleclsid 00112233-4455-6677-8899-AABBCCDDEEFF}
        {\*\objdata 0102a0ff}
        {\result fallback}}B}";
    let document = Tree::parse(source, litchi_rtf::read::OpenOptions::default()).unwrap();
    let object = &document.objects()[0];
    assert_eq!(object.alias.as_deref(), Some("Alias Name"));
    assert_eq!(object.section.as_deref(), Some("Section1"));

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("{\\*\\objalias Alias Name}"));
    assert!(serialized.contains("{\\*\\objsect Section1}"));

    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    let reparsed_object = &reparsed.objects()[0];
    assert_eq!(reparsed_object.alias, object.alias);
    assert_eq!(reparsed_object.section, object.section);
    assert_eq!(reparsed_object.class_id, object.class_id);
    assert_eq!(reparsed_object.data, object.data);
}

#[test]
fn rejects_misplaced_object_alias_and_section_destinations() {
    for source in [
        // Unstarred destinations.
        r"{\rtf1{\object\objemb{\objalias X}{\*\objdata 00}}}",
        r"{\rtf1{\object\objemb{\objsect X}{\*\objdata 00}}}",
        // Duplicate destinations.
        r"{\rtf1{\object\objemb{\*\objalias X}{\*\objalias Y}{\*\objdata 00}}}",
        r"{\rtf1{\object\objemb{\*\objsect X}{\*\objsect Y}{\*\objdata 00}}}",
        // Metadata after objdata.
        r"{\rtf1{\object\objemb{\*\objdata 00}{\*\objalias X}}}",
        r"{\rtf1{\object\objemb{\*\objdata 00}{\*\objsect X}}}",
        // Metadata after the CLSID.
        r"{\rtf1{\object\objemb{\*\oleclsid X}{\*\objalias A}{\*\objdata 00}}}",
    ] {
        assert!(
            Tree::parse(source, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn round_trips_object_time_metadata() {
    let source = r"{\rtf1\ansi A{\object\objlink\objh40\objw50
        {\*\objclass Package}{\*\objname Visible Name}
        {\*\objsect Section1}{\*\objtime \yr1994\mo6\dy16\hr12\min30}
        {\*\oleclsid 00112233-4455-6677-8899-AABBCCDDEEFF}
        {\*\objdata 0102a0ff}
        {\result fallback}}B}";
    let document = Tree::parse(source, litchi_rtf::read::OpenOptions::default()).unwrap();
    let object = &document.objects()[0];
    let time = object.time.unwrap();
    assert_eq!(time.year, Some(1994));
    assert_eq!(time.month, Some(6));
    assert_eq!(time.day, Some(16));
    assert_eq!(time.hour, Some(12));
    assert_eq!(time.minute, Some(30));
    assert_eq!(time.second, None);

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("{\\*\\objtime \\yr1994\\mo6\\dy16\\hr12\\min30}"));

    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.objects()[0].time, object.time);
    assert_eq!(reparsed.objects()[0].section, object.section);
}

#[test]
fn rejects_misplaced_object_time_destinations() {
    for source in [
        // Unstarred destination.
        r"{\rtf1{\object\objemb{\objtime \yr1994}{\*\objdata 00}}}",
        // Duplicate destination.
        r"{\rtf1{\object\objemb{\*\objtime \yr1994}{\*\objtime \yr1995}{\*\objdata 00}}}",
        // Destination after objdata.
        r"{\rtf1{\object\objemb{\*\objdata 00}{\*\objtime \yr1994}}}",
        // Text payload in the destination.
        r"{\rtf1{\object\objemb{\*\objtime noon}{\*\objdata 00}}}",
        // Active non-time control in the destination.
        r"{\rtf1{\object\objemb{\*\objtime \b\yr1994}{\*\objdata 00}}}",
    ] {
        assert!(
            Tree::parse(source, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted malformed {source}"
        );
    }
}
