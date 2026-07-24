use litchi_rtf::{ImageType, Picture, PictureIdentity, RtfDocument, RtfWriter};
use std::borrow::Cow;

const UID: &[u8] = &[
    0x00, 0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd, 0xee, 0xff,
];

#[test]
fn parses_and_canonically_round_trips_picture_identity() {
    let source = concat!(
        r#"{\rtf1{\pict\pngblip\picw1\pich1\bliptag-7\blipupi96"#,
        r#"{\*\blipuid 00112233445566778899aabbccddeeff}"#,
        r#" 89504e470d0a1a0a}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let picture = &document.pictures()[0];
    assert_eq!(picture.image_type, ImageType::Png);
    assert_eq!(picture.identity.as_ref().unwrap().tag, Some(-7));
    assert_eq!(picture.identity.as_ref().unwrap().units_per_inch, Some(96));
    assert_eq!(picture.identity.as_ref().unwrap().uid.as_deref(), Some(UID));

    let mut first = br"{\rtf1".to_vec();
    RtfWriter::new(&mut first).write_picture(picture).unwrap();
    first.push(b'}');
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(reparsed.pictures(), document.pictures());

    let mut second = br"{\rtf1".to_vec();
    RtfWriter::new(&mut second)
        .write_picture(&reparsed.pictures()[0])
        .unwrap();
    second.push(b'}');
    assert_eq!(first, second);

    let mut borrowed = Picture::new(ImageType::Png, Cow::Borrowed(&[0x89, b'P', b'N', b'G']));
    borrowed.identity = Some(PictureIdentity {
        tag: None,
        units_per_inch: None,
        uid: Some(Cow::Borrowed(UID)),
    });
    assert!(matches!(
        borrowed.identity.unwrap().uid,
        Some(Cow::Borrowed(_))
    ));
}

#[test]
fn accepts_real_empty_uid_but_rejects_malformed_identity_content() {
    let empty =
        RtfDocument::parse(r#"{\rtf1{\pict\pngblip{\*\blipuid } 89504e470d0a1a0a}}"#).unwrap();
    assert_eq!(
        empty.pictures()[0]
            .identity
            .as_ref()
            .unwrap()
            .uid
            .as_deref(),
        Some(&[][..])
    );

    let malformed = [
        r#"{\rtf1\bliptag1 X}"#,
        r#"{\rtf1\blipupi96 X}"#,
        r#"{\rtf1{\*\blipuid 00112233445566778899aabbccddeeff}}"#,
        r#"{\rtf1{\pict\pngblip{\blipuid 00112233445566778899aabbccddeeff} 00}}"#,
        r#"{\rtf1{\pict\pngblip\blipuid 00112233445566778899aabbccddeeff 00}}"#,
        r#"{\rtf1{\pict\pngblip\bliptag1\bliptag2 00}}"#,
        r#"{\rtf1{\pict\pngblip\blipupi96\blipupi97 00}}"#,
        r#"{\rtf1{\pict\pngblip{\*\blipuid 00112233445566778899aabbccddeeff}{\*\blipuid 00112233445566778899aabbccddeeff} 00}}"#,
        r#"{\rtf1{\pict\pngblip\blipupi-1 00}}"#,
        r#"{\rtf1{\pict\pngblip\blipupi0 00}}"#,
        r#"{\rtf1{\pict\pngblip\blipupi65536 00}}"#,
        r#"{\rtf1{\pict\pngblip\blipupi96\bliptag1 00}}"#,
        r#"{\rtf1{\pict\pngblip{\*\blipuid 00112233445566778899aabbccddeeff}\blipupi96 00}}"#,
        r#"{\rtf1{\pict\pngblip{\*\blipuid 0} 00}}"#,
        r#"{\rtf1{\pict\pngblip{\*\blipuid zz} 00}}"#,
        r#"{\rtf1{\pict\pngblip{\*\blipuid 00112233445566778899aabbccddee} 00}}"#,
        r#"{\rtf1{\pict\pngblip{\*\blipuid 00112233445566778899aabbccddeeff00} 00}}"#,
        r#"{\rtf1{\pict\pngblip{\*\blipuid{\field danger}} 00}}"#,
        r#"{\rtf1{\pict\pngblip{\*\blipuid{\object danger}} 00}}"#,
        r#"{\rtf1{\pict\pngblip{\*\blipuid\bin2 xx} 00}}"#,
        r#"{\rtf1{\pict\pngblip 00{\*\blipuid 00112233445566778899aabbccddeeff}}}"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed RTF: {source}"
        );
    }

    let mut picture = Picture::new(ImageType::Png, Cow::Borrowed(&[0x89, b'P', b'N', b'G']));
    picture.identity = Some(PictureIdentity {
        tag: None,
        units_per_inch: Some(96),
        uid: Some(Cow::Borrowed(&[1])),
    });
    assert!(RtfWriter::new(Vec::new()).write_picture(&picture).is_err());
}

fn isolated_picture(fixture: &[u8], required: &[u8]) -> Vec<u8> {
    let marker = br"{\pict";
    let mut cursor = 0usize;
    loop {
        let start = cursor
            + fixture[cursor..]
                .windows(marker.len())
                .position(|window| window == marker)
                .unwrap();
        let mut depth = 0usize;
        let mut end = None;
        for (offset, byte) in fixture[start..].iter().enumerate() {
            match byte {
                b'{' => depth += 1,
                b'}' => {
                    depth -= 1;
                    if depth == 0 {
                        end = Some(start + offset + 1);
                        break;
                    }
                },
                _ => {},
            }
        }
        let end = end.unwrap();
        if fixture[start..end]
            .windows(required.len())
            .any(|window| window == required)
        {
            let mut source = br"{\rtf1".to_vec();
            source.extend_from_slice(&fixture[start..end]);
            source.push(b'}');
            return source;
        }
        cursor = end;
    }
}

#[test]
fn parses_multiple_bundled_libreoffice_identity_fixtures() {
    let uid_only = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf127806.rtf"
    );
    let source = isolated_picture(uid_only, br"\blipuid");
    let document = RtfDocument::parse_bytes(&source).unwrap();
    assert_eq!(
        document.pictures()[0]
            .identity
            .as_ref()
            .unwrap()
            .uid
            .as_deref(),
        Some(
            &[
                0xc0, 0xd4, 0x86, 0xc2, 0x6e, 0xfd, 0x24, 0x45, 0x9c, 0x14, 0x59, 0x2d, 0x82, 0x49,
                0xa3, 0x2b,
            ][..]
        )
    );

    let with_upi = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf153613_inlineAfterPgBreak.rtf"
    );
    let source = isolated_picture(with_upi, br"\blipupi2");
    let document = RtfDocument::parse_bytes(&source).unwrap();
    let identity = document.pictures()[0].identity.as_ref().unwrap();
    assert_eq!(identity.tag, Some(1_890_591_635));
    assert_eq!(identity.units_per_inch, Some(2));
    assert_eq!(
        identity.uid.as_deref(),
        Some(
            &[
                0x70, 0xb0, 0x23, 0x93, 0xad, 0x94, 0xa8, 0x7d, 0x9e, 0xc5, 0x96, 0xdf, 0xc3, 0xb1,
                0x42, 0x83,
            ][..]
        )
    );
}
