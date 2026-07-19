use litchi_rtf::{BodyStoryEvent, RtfDocument, RtfWriter};

fn round_trip(source: &str) -> RtfDocument<'static> {
    let document = RtfDocument::parse(source).unwrap();
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes)
        .write_document(&document)
        .unwrap();
    RtfDocument::parse(&String::from_utf8(bytes).unwrap()).unwrap()
}

#[test]
fn preserves_equal_offset_order_across_all_atomic_body_destinations() {
    let source = concat!(
        r#"{\rtf1{\*\bkmkstart B}"#,
        r#"{\listtext L}{\xe X}{\footnote N}"#,
        r#"{\*\shppict{\pict\pngblip 89504e470d0a1a0a}}"#,
        r#"{\object\objemb{\*\objdata 00}}"#,
        r#"{\*\do\dobxpage\dobypara\dodhgt1\dpline\dpptx1\dppty2\dpptx3\dppty4\dpx5\dpy6\dpxsize7\dpysize8}"#,
        r#"{\*\atnid A}{\*\atnauthor Ada}\chatn{\*\annotation point}"#,
        r#"{\*\bkmkend B}Z}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let expected = [
        BodyStoryEvent::BookmarkStart(0),
        BodyStoryEvent::GeneratedListMarker(0),
        BodyStoryEvent::NavigationEntry(0),
        BodyStoryEvent::Note(0),
        BodyStoryEvent::PictureCompatibility(0),
        BodyStoryEvent::Object(0),
        BodyStoryEvent::LegacyDrawing(0),
        BodyStoryEvent::AnnotationStart(0),
        BodyStoryEvent::AnnotationEnd(0),
        BodyStoryEvent::BookmarkEnd(0),
    ];
    assert_eq!(document.body_story_events(), expected);
    assert_eq!(round_trip(source).body_story_events(), expected);
}

#[test]
fn preserves_revision_and_form_field_boundaries_without_canonical_ties() {
    let revisions = r#"{\rtf1{\*\revtbl{A;}}{\revised\revauth0\revdttm1 X}{\deleted\revauthdel0\revdttmdel2 Y}Z}"#;
    let document = RtfDocument::parse(revisions).unwrap();
    let expected = [
        BodyStoryEvent::RevisionStart(0),
        BodyStoryEvent::RevisionEnd(0),
        BodyStoryEvent::RevisionDeletion(1),
    ];
    assert_eq!(document.body_story_events(), expected);
    assert_eq!(round_trip(revisions).body_story_events(), expected);

    let adjacent = r#"{\rtf1{\*\revtbl{A;}}{\revised\revauth0 X}{\revised\revauth0 Y}}"#;
    let document = round_trip(adjacent);
    assert_eq!(document.revisions().len(), 2);
    assert_eq!(
        document.body_story_events(),
        [
            BodyStoryEvent::RevisionStart(0),
            BodyStoryEvent::RevisionEnd(0),
            BodyStoryEvent::RevisionStart(1),
            BodyStoryEvent::RevisionEnd(1),
        ]
    );

    let form = r#"{\rtf1{\field{\*\fldinst FORMCHECKBOX{\*\formfield{\fftype1\fftypetxt0\ffhps20\ffdefres0\ffres0}}}{\fldrslt }}}"#;
    let document = round_trip(form);
    assert_eq!(
        document.body_story_events(),
        [
            BodyStoryEvent::FormFieldStart(0),
            BodyStoryEvent::FormFieldEnd(0)
        ]
    );
}
