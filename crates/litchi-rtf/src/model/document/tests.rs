use super::super::error::RtfError;
use super::*;
use crate::{Alignment, ListFollow, ListJustification, ListLevelType, RevisionType, StyleType};
use std::borrow::Cow;

#[test]
fn test_simple_document() {
    let rtf = r#"{\rtf1\ansi Hello World!\par}"#;
    let doc = RtfDocument::parse(rtf).unwrap();
    let text = doc.text();
    assert!(text.contains("Hello World"));
}

#[test]
fn root_drawing_mutations_preflight_corrupt_indices_and_events() {
    let mut document = RtfDocument::parse(r#"{\rtf1 body}"#).unwrap();
    let mut first = crate::Shape::new(crate::ShapeType::Rectangle);
    first.position = document.text().len();
    let mut second = crate::Shape::new(crate::ShapeType::Ellipse);
    second.position = document.text().len();
    document.add_shape(first).unwrap();
    document.add_shape(second).unwrap();

    *document.drawing_order.last_mut().unwrap() = crate::StoryDrawing::Shape(usize::MAX);
    let shapes_before = document.shapes.clone();
    let order_before = document.drawing_order.clone();
    let events_before = document.body_story_events.clone();
    let mut third = crate::Shape::new(crate::ShapeType::TextBox);
    third.position = document.text().len();
    assert!(document.add_shape(third).is_err());
    assert!(document.move_drawing(0, 1).is_err());
    assert_eq!(document.shapes, shapes_before);
    assert_eq!(document.drawing_order, order_before);
    assert_eq!(document.body_story_events, events_before);

    let mut document = RtfDocument::parse(r#"{\rtf1 body}"#).unwrap();
    let mut first = crate::Shape::new(crate::ShapeType::Rectangle);
    first.position = document.text().len();
    let mut second = crate::Shape::new(crate::ShapeType::Ellipse);
    second.position = document.text().len();
    document.add_shape(first).unwrap();
    document.add_shape(second).unwrap();
    let drawing_event = document
        .body_story_events
        .iter()
        .rposition(|event| matches!(event, crate::BodyStoryEvent::Drawing(_)))
        .unwrap();
    document.body_story_events.remove(drawing_event);
    let order_before = document.drawing_order.clone();
    let events_before = document.body_story_events.clone();
    assert!(document.move_drawing(0, 1).is_err());
    assert_eq!(document.drawing_order, order_before);
    assert_eq!(document.body_story_events, events_before);
}

#[test]
fn background_mutations_are_total_for_a_corrupt_private_index() {
    let mut document = RtfDocument::parse(r#"{\rtf1 body}"#).unwrap();
    let mut shape = crate::Shape::new(crate::ShapeType::Rectangle);
    shape.position = document.text().len();
    document.add_shape(shape).unwrap();
    document.background_shape_index = Some(usize::MAX);
    let shapes_before = document.shapes.clone();

    assert!(
        document
            .set_background_shape(crate::Shape::new(crate::ShapeType::Ellipse))
            .is_err()
    );
    assert_eq!(document.shapes, shapes_before);
    assert!(document.clear_background_shape().is_none());
    assert_eq!(document.shapes, shapes_before);

    document.background_shape_index = Some(usize::MAX);
    document.clear_shapes();
    assert!(document.shapes.is_empty());
    assert!(document.drawing_order.is_empty());
    assert!(document.background_shape_index.is_none());
}

#[test]
fn test_formatted_text() {
    let rtf = r#"{\rtf1\ansi{\b Bold}{\i Italic}\par}"#;
    let doc = RtfDocument::parse(rtf).unwrap();
    let runs = doc.runs();
    assert!(!runs.is_empty());
}

#[test]
fn parses_and_owns_shapes_and_shape_groups() {
    let rtf = r#"{\rtf1\ansi
            {\shp\shpleft10\shptop20\shpright310\shpbottom140
                \shprotation15\shpz4\shpwr4\shplockanchor{\*\shpinst
                    {\sp{\sn shapeType}{\sv 202}}
                    {\sp{\sn wzName}{\sv Owned Text Box}}
                    {\sp{\sn fBehindDocument}{\sv 1}}
                    {\sp{\sn fLockPosition}{\sv 1}}
                    {\sp{\sn fillType}{\sv 7}}
                    {\sp{\sn fillColor}{\sv 66051}}
                    {\sp{\sn fillBackColor}{\sv 263430}}
                    {\sp{\sn fillOpacity}{\sv 32768}}
                    {\sp{\sn fLine}{\sv 0}}
                    {\sp{\sn lineColor}{\sv 460809}}
                    {\sp{\sn lineWidth}{\sv 12700}}
                    {\sp{\sn futureOfficeArtProperty}{\sv retained}}
                    {\shptxt Hello \u20320?}}}
            {\shpgrp{\*\shpinst\shpleft1\shptop2\shpright801\shpbottom602
                {\sp{\sn wzName}{\sv Owned Group}}
                {\shp{\*\shpinst\shpleft5\shptop6\shpwidth70\shpheight80\shpfblwtxt1{\sp{\sn shapeType}{\sv 1}}}}
                {\shp{\*\shpinst\shpleft15\shptop16\shpwidth90\shpheight100{\sp{\sn shapeType}{\sv 3}}}}
                {\shpgrp{\*\shpinst\shpleft100\shptop110\shpright400\shpbottom510
                    {\sp{\sn wzName}{\sv Owned Nested Group}}
                    {\shp{\*\shpinst\shpleft1\shptop2\shpwidth3\shpheight4{\sp{\sn shapeType}{\sv 20}}}}}}}}
        }"#;
    let doc = RtfDocument::parse(rtf).unwrap();

    let shape = &doc.shapes()[0];
    assert_eq!(shape.shape_type, crate::ShapeType::TextBox);
    assert_eq!(shape.geometry.x, 10);
    assert_eq!(shape.geometry.y, 20);
    assert_eq!(shape.geometry.width, 300);
    assert_eq!(shape.geometry.height, 120);
    assert_eq!(shape.geometry.rotation, 15);
    assert_eq!(shape.geometry.z_order, 4);
    assert_eq!(shape.text, "Hello 你");
    assert_eq!(shape.name, "Owned Text Box");
    assert!(shape.behind_doc);
    assert!(shape.locked);
    assert_eq!(shape.wrap_mode, crate::WrapMode::Tight);
    assert_eq!(shape.fill.fill_type, crate::FillType::Gradient);
    assert_eq!(shape.fill.color.raw(), 66_051);
    assert_eq!(shape.fill.color.red(), 1);
    assert_eq!(shape.fill.color.green(), 2);
    assert_eq!(shape.fill.color.blue(), 3);
    assert_eq!(shape.fill.color2.unwrap().raw(), 263_430);
    assert_eq!(shape.fill.opacity.raw(), 32_768);
    assert_eq!(shape.fill.opacity.as_fraction(), 0.5);
    assert!(!shape.line.visible);
    assert_eq!(shape.line.color.raw(), 460_809);
    assert_eq!(shape.line.width_emu, 12_700);
    assert!(shape.properties.iter().any(|property| {
        property.name == "futureOfficeArtProperty" && property.value == "retained"
    }));
    assert!(
        shape
            .properties
            .iter()
            .all(|property| matches!(property.name, Cow::Owned(_)))
    );
    assert!(matches!(shape.text, Cow::Owned(_)));

    let group = &doc.shape_groups()[0];
    assert_eq!(group.geometry, crate::ShapeGeometry::new(1, 2, 800, 600));
    assert_eq!(group.shapes.len(), 2);
    assert_eq!(group.shapes[0].shape_type, crate::ShapeType::Rectangle);
    assert!(group.shapes[0].behind_doc);
    assert_eq!(group.shapes[0].wrap_mode, crate::WrapMode::Behind);
    assert_eq!(group.shapes[1].shape_type, crate::ShapeType::Ellipse);
    assert_eq!(group.name, "Owned Group");
    assert!(matches!(group.name, Cow::Owned(_)));
    assert!(
        group
            .properties
            .iter()
            .all(|property| matches!(property.value, Cow::Owned(_)))
    );
    assert_eq!(group.groups().len(), 1);
    let nested = &group.groups()[0];
    assert_eq!(nested.name, "Owned Nested Group");
    assert_eq!(
        nested.geometry,
        crate::ShapeGeometry::new(100, 110, 300, 400)
    );
    assert_eq!(nested.shapes().len(), 1);
    assert_eq!(nested.shapes()[0].shape_type, crate::ShapeType::Line);
    assert!(matches!(nested.name, Cow::Owned(_)));
}

#[test]
fn rejects_excessively_nested_shape_groups() {
    let mut rtf = String::from("{\\rtf1");
    for _ in 0..=64 {
        rtf.push_str("{\\shpgrp{\\*\\shpinst");
    }
    for _ in 0..=64 {
        rtf.push_str("}}");
    }
    rtf.push('}');

    let error = match RtfDocument::parse(&rtf) {
        Ok(_) => panic!("excessive shape-group nesting should fail"),
        Err(error) => error,
    };
    assert!(
        matches!(error, RtfError::MalformedDocument(message) if message.contains("shape group nesting"))
    );
}

#[test]
fn parses_real_background_shape_fixture_with_trailing_newline() {
    let rtf = include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/rtf/background.rtf"
    ));
    let doc = RtfDocument::parse(rtf).unwrap();

    assert_eq!(doc.shapes().len(), 2);
    assert_eq!(doc.shapes()[0].geometry.x, 2633);
    assert_eq!(doc.shapes()[0].geometry.width, 2220);
    assert_eq!(doc.shapes()[0].fill.color.raw(), 5_880_731);
    assert_eq!(doc.shapes()[0].fill.fill_type, crate::FillType::Solid);
    assert_eq!(doc.shapes()[0].properties.len(), 4);
    assert_eq!(doc.shapes()[1].geometry.x, 488);
    assert_eq!(doc.shapes()[1].geometry.width, 1515);
    assert_eq!(doc.shapes()[1].fill.color.raw(), 5_066_944);
    assert_eq!(
        doc.text(),
        "First should be foreground, the second should be background.\n"
    );
}

#[test]
fn preserves_real_watermark_office_art_properties() {
    let rtf = include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/rtf/watermark.rtf"
    ));
    let doc = RtfDocument::parse(rtf).unwrap();

    assert!(doc.shapes().is_empty());
    let header_shapes: Vec<_> = doc
        .sections()
        .iter()
        .flat_map(|section| &section.headers_footers)
        .flat_map(|header_footer| &header_footer.shapes)
        .collect();
    assert_eq!(header_shapes.len(), 3);
    let shape = header_shapes[0];
    assert_eq!(shape.shape_type, crate::ShapeType::Custom(136));
    assert_eq!(shape.geometry.rotation, 315);
    assert_eq!(shape.fill.color.raw(), 6_108_695);
    assert_eq!(shape.fill.opacity.raw(), 32_768);
    assert!(!shape.line.visible);
    assert_eq!(shape.name, "PowerPlusWaterMarkObject142907");
    assert!(shape.behind_doc);
    assert_eq!(shape.property("gtextUNICODE"), Some("ASAP"));
}

#[test]
fn parses_shape_from_ignorable_page_background_destination() {
    let rtf = include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/rtf/page-background.rtf"
    ));
    let doc = RtfDocument::parse(rtf).unwrap();

    assert_eq!(doc.shapes().len(), 1);
    let shape = &doc.shapes()[0];
    assert_eq!(shape.shape_type, crate::ShapeType::Rectangle);
    assert_eq!(shape.fill.color.raw(), 5_296_274);
    assert!(shape.is_background);
    assert!(shape.behind_doc);
    assert_eq!(shape.property("bWMode"), Some("9"));
    assert!(!doc.text().contains("shapeType"));
    assert!(!doc.text().contains("fillColor"));
}

#[test]
fn extracts_embedded_object_metadata_and_native_bytes_without_activation() {
    let rtf = r#"{\rtf1\ansi Before {\object\objemb\objw1440\objh720\objlock\objupdate\objsetsize
            {\*\objclass Package}{\*\objname Owned \u20320? Object}
            {\*\objdata
                01050000 02000000 08000000 5061636b61676500
                00000000 00000000 08000000 d0cf11e0a1b11ae1}
            {\result fallback {\pict\pngblip\picw10\pich20 89504e470d0a1a0a}}} After}"#;
    let doc = RtfDocument::parse(rtf).unwrap();

    assert_eq!(doc.text(), "Before  After");
    assert_eq!(doc.objects().len(), 1);
    let object = &doc.objects()[0];
    assert_eq!(object.kind, crate::ObjectKind::Embedded);
    assert_eq!(object.class_name, "Package");
    assert_eq!(object.name, "Owned 你 Object");
    assert_eq!(object.width, 1440);
    assert_eq!(object.height, 720);
    assert!(object.locked);
    assert!(object.update_requested);
    assert!(object.set_size);
    assert!(matches!(object.class_name, Cow::Owned(_)));
    assert_eq!(object.result_text, "fallback");
    assert_eq!(object.result_picture_indices, [0]);
    assert!(matches!(object.result_text, Cow::Owned(_)));
    assert_eq!(doc.pictures().len(), 1);
    assert_eq!(doc.pictures()[0].image_type, crate::ImageType::Png);
    assert_eq!(doc.pictures()[0].width, Some(10));
    assert_eq!(doc.pictures()[0].height, Some(20));

    let header = object.ole_header().unwrap();
    assert_eq!(header.ole_version, 0x501);
    assert_eq!(header.format_id, 2);
    assert_eq!(header.class_name, b"Package");
    assert!(header.is_compound_file());
}

#[test]
fn rejects_invalid_embedded_object_hex() {
    let error = match RtfDocument::parse(r#"{\rtf1{\object{\*\objdata 0xz}}}"#) {
        Ok(_) => panic!("invalid object hex should fail"),
        Err(error) => error,
    };
    assert!(
        matches!(error, RtfError::MalformedDocument(message) if message.contains("non-hexadecimal"))
    );
}

#[test]
fn preserves_exact_binary_picture_payload() {
    let payload = [
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, b'{', b'\\', b'}',
    ];
    let mut rtf = br"{\rtf1{\pict\pngblip\bin11 ".to_vec();
    rtf.extend_from_slice(&payload);
    rtf.extend_from_slice(b"}}");

    let doc = RtfDocument::parse_bytes(&rtf).unwrap();
    assert_eq!(doc.pictures().len(), 1);
    assert_eq!(doc.pictures()[0].image_type, crate::ImageType::Png);
    assert_eq!(doc.pictures()[0].data(), payload);
}

#[test]
fn rejects_unclosed_document_and_destination_groups() {
    for rtf in [
        r#"{\rtf1 body"#,
        r#"{\rtf1{\*\unknown destination"#,
        r#"{\rtf1{\shp{\*\shpinst\shpleft1}}"#,
        r#"{\rtf1{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 1}"#,
    ] {
        let result = RtfDocument::parse(rtf);
        match result {
            Err(RtfError::UnexpectedEof) => {},
            Err(error) => panic!("unexpected error for {rtf:?}: {error}"),
            Ok(_) => panic!("unexpected success for {rtf:?}"),
        }
    }
}

#[test]
fn parses_complete_document_info_without_leaking_into_body() {
    let rtf = r#"{\rtf1\ansi\ansicpg1252
            {\info
                {\title Annual \u20320? Report}
                {\subject Results}{\author Ada}{\manager Grace}
                {\company Caf\'e9 Corp \u8364?}{\operator Linus}{\category Finance}
                {\keywords alpha; beta}{\comment Reviewed}
                {\creatim\yr2025\mo7\dy14\hr9\min8\sec7}
                {\revtim\yr2026\mo1\dy2\hr3\min4\sec5}
                {\printim\yr2026\mo2\dy3}{\buptim\yr2024\mo12\dy31}
                \version4\vern9\edmins120\nofpages8\nofwords900
                \nofchars4200\nofcharsws5000\id77
            }
            Body text\par}"#;
    let doc = RtfDocument::parse(rtf).unwrap();
    let info = doc.info();
    assert_eq!(info.title.as_deref(), Some("Annual 你 Report"));
    assert_eq!(info.subject.as_deref(), Some("Results"));
    assert_eq!(info.author.as_deref(), Some("Ada"));
    assert_eq!(info.manager.as_deref(), Some("Grace"));
    assert_eq!(info.company.as_deref(), Some("Café Corp €"));
    assert_eq!(info.operator.as_deref(), Some("Linus"));
    assert_eq!(info.category.as_deref(), Some("Finance"));
    assert_eq!(info.keywords.as_deref(), Some("alpha; beta"));
    assert_eq!(info.comment.as_deref(), Some("Reviewed"));
    assert_eq!(info.creation_time.as_deref(), Some("2025-07-14T09:08:07"));
    assert_eq!(info.revision_time.as_deref(), Some("2026-01-02T03:04:05"));
    assert_eq!(info.print_time.as_deref(), Some("2026-02-03T00:00:00"));
    assert_eq!(info.backup_time.as_deref(), Some("2024-12-31T00:00:00"));
    assert_eq!(info.version, Some(4));
    assert_eq!(info.revision, Some(9));
    assert_eq!(info.editing_time, Some(120));
    assert_eq!(info.pages, Some(8));
    assert_eq!(info.words, Some(900));
    assert_eq!(info.characters, Some(4200));
    assert_eq!(info.characters_with_spaces, Some(5000));
    assert_eq!(info.id, Some(77));
    assert_eq!(doc.text().trim(), "Body text");
}

#[test]
fn ignores_unknown_nested_info_destinations() {
    let rtf = r#"{\rtf1{\info{\*\unknown nested {data}}{\title Kept}}Text}"#;
    let doc = RtfDocument::parse(rtf).unwrap();
    assert_eq!(doc.info().title.as_deref(), Some("Kept"));
    assert_eq!(doc.text(), "Text");
}

#[test]
fn parses_complete_stylesheet_without_leaking_names_into_body() {
    let rtf = r#"{\rtf1\ansi
            {\stylesheet
                {\s0\fs22\ql\snext0\sqformat\spriority0 Normal;}
                {\s1\b\qc\sb120\li240\keepn\sbasedon0\snext0\slink2
                    \sautoupd\shidden\slocked\ssemihidden\sunhideused\sqformat
                    \spriority9\styrsid42\spersonal\scompose\sreply Heading \u20320?;}
                {\*\cs2\i\additive\sbasedon0\slink1 Emphasis;}
                {\*\ds3 Section Style;}
                {\*\ts4{\*\unknown ignored} Table Style;}
            }
            Body}"#;
    let doc = RtfDocument::parse(rtf).unwrap();
    assert_eq!(doc.text().trim(), "Body");
    assert_eq!(doc.stylesheet().styles().len(), 5);

    let heading = doc.stylesheet().get_typed(StyleType::Paragraph, 1).unwrap();
    assert_eq!(heading.name, "Heading 你");
    assert_eq!(heading.based_on, Some(0));
    assert_eq!(heading.next_style, Some(0));
    assert_eq!(heading.linked_style, Some(2));
    assert!(heading.formatting.bold);
    let paragraph = heading.paragraph.unwrap();
    assert_eq!(paragraph.alignment, Alignment::Center);
    assert_eq!(paragraph.spacing.before, 120);
    assert_eq!(paragraph.indentation.left, 240);
    assert!(paragraph.keep_next);
    assert!(heading.auto_update);
    assert!(heading.hidden);
    assert!(heading.locked);
    assert!(heading.semi_hidden);
    assert!(heading.unhide_when_used);
    assert!(heading.quick_format);
    assert_eq!(heading.priority, Some(9));
    assert_eq!(heading.revision_id, Some(42));
    assert!(heading.personal);
    assert!(heading.compose);
    assert!(heading.reply);

    let emphasis = doc.stylesheet().get_typed(StyleType::Character, 2).unwrap();
    assert_eq!(emphasis.name, "Emphasis");
    assert!(emphasis.formatting.italic);
    assert!(emphasis.additive);
    assert!(emphasis.paragraph.is_none());
    assert!(doc.stylesheet().get_typed(StyleType::Section, 3).is_some());
    assert!(doc.stylesheet().get_typed(StyleType::Table, 4).is_some());
}

#[test]
fn parses_list_and_override_tables_without_leaking_labels() {
    let rtf = r#"{\rtf1\ansi
            {\*\listtable
                {\list\listtemplateid42\listhybrid
                    {\listlevel\levelnfc0\leveljc2\levelfollow1\levelstartat3
                        \levelspace120\levelindent360
                        {\leveltext\'02\'00.;}{\levelnumbers\'01;}\f2}
                    {\listlevel\levelnfc23\leveljc0\levelfollow2\levelstartat1
                        {\leveltext\'01\u8226?;}{\levelnumbers;}}
                    {\listname Outline;}\listid77}
            }
            {\listoverridetable
                {\listoverride\listid77\listoverridecount1
                    {\lfolevel\listoverridestartat\levelstartat9}\ls4}}
            Body}"#;
    let doc = RtfDocument::parse(rtf).unwrap();
    assert_eq!(doc.text().trim(), "Body");
    assert_eq!(doc.list_table().lists().len(), 1);

    let list = doc.list_table().get(77).unwrap();
    assert_eq!(list.template_id, 42);
    assert!(!list.simple);
    assert!(list.hybrid);
    assert_eq!(list.name, "Outline");
    assert_eq!(list.levels.len(), 2);
    let decimal = &list.levels[0];
    assert_eq!(decimal.level_type, ListLevelType::Decimal);
    assert_eq!(decimal.number_text, "\0.");
    assert_eq!(decimal.start_at, 3);
    assert_eq!(decimal.justification, ListJustification::Right);
    assert_eq!(decimal.follow, ListFollow::Space);
    assert_eq!(decimal.font_ref, 2);
    assert_eq!(decimal.indent, 360);
    assert_eq!(decimal.space, 120);
    let bullet = &list.levels[1];
    assert_eq!(bullet.level_type, ListLevelType::Bullet);
    assert_eq!(bullet.number_text, "•");
    assert_eq!(bullet.follow, ListFollow::Nothing);

    let list_override = doc.list_override_table().get(4).unwrap();
    assert_eq!(list_override.list_id, 77);
    assert_eq!(list_override.level_count_override, Some(1));
    assert_eq!(list_override.start_at_override, Some(9));
}

#[test]
fn preserves_paragraph_list_instance_and_level() {
    let doc = RtfDocument::parse(r#"{\rtf1\pard\ls4\ilvl2 Listed text}"#).unwrap();
    assert_eq!(doc.text(), "Listed text");
    let paragraph = doc.blocks().last().unwrap().paragraph;
    assert_eq!(paragraph.list_override, Some(4));
    assert_eq!(paragraph.list_level, Some(2));
}

#[test]
fn parses_tracked_insertions_and_deletions_with_author_ranges() {
    let rtf = r#"{\rtf1\ansi
            {\*\revtbl {Unknown;}{Max \u20320?;}}
            Before {\deleted\revauthdel1\revdttmdel1199059860 old \u20320? text}
            and {\revised\revauth1\revdttm-1501115711 new text} after}"#;
    let doc = RtfDocument::parse(rtf).unwrap();
    let body = doc.text();
    assert!(!body.contains("old 你 text"));
    assert!(body.contains("and new text after"));
    assert_eq!(doc.revisions().len(), 2);

    let deletion = &doc.revisions()[0];
    assert_eq!(deletion.revision_type, RevisionType::Deletion);
    assert_eq!(deletion.id, 1);
    assert_eq!(deletion.author, "Max 你");
    assert_eq!(deletion.date.as_deref(), Some("1199059860"));
    assert_eq!(deletion.content, "old 你 text");
    assert_eq!(deletion.position, deletion.range_end);
    assert!(!body.contains(deletion.content.as_ref()));

    let insertion = &doc.revisions()[1];
    assert_eq!(insertion.revision_type, RevisionType::Insertion);
    assert_eq!(insertion.author, "Max 你");
    assert_eq!(insertion.date.as_deref(), Some("-1501115711"));
    assert_eq!(insertion.content, "new text");
    assert_eq!(
        body.get(insertion.position..insertion.range_end),
        Some(insertion.content.as_ref())
    );
}

#[test]
fn revision_toggle_boundaries_flush_preceding_text() {
    let doc = RtfDocument::parse(
        r#"{\rtf1{\*\revtbl Unknown;}plain \revised\revauth0\revdttm1 changed\revised0 plain}"#,
    )
    .unwrap();
    assert_eq!(doc.text(), "plain changedplain");
    assert_eq!(doc.revisions().len(), 1);
    assert_eq!(doc.revisions()[0].content, "changed");
    assert_eq!(doc.revisions()[0].author, "Unknown");
}

#[test]
fn parses_nested_bookmarks_with_range_metadata() {
    let rtf = r#"{\rtf1\ansi Before {\*\bkmkstart\bkmkcolf1\bkmkcoll3\bkmkpub Outer}alpha {\*\bkmkstart Inner}\u20320?{\*\bkmkend Inner} omega{\*\bkmkend Outer} After}"#;
    let doc = RtfDocument::parse(rtf).unwrap();
    assert_eq!(doc.text(), "Before alpha 你 omega After");

    let outer = doc.bookmarks().get("Outer").unwrap();
    assert_eq!(outer.position, "Before ".len());
    assert_eq!(outer.content, "alpha 你 omega");
    assert_eq!(outer.first_column, Some(1));
    assert_eq!(outer.last_column, Some(3));
    assert!(outer.is_public);

    let inner = doc.bookmarks().get("Inner").unwrap();
    assert_eq!(inner.position, "Before alpha ".len());
    assert_eq!(inner.content, "你");
}

#[test]
fn parses_annotation_range_author_and_body_without_text_leakage() {
    let rtf = r#"{\rtf1\ansi aaa {\*\atrfstart 7}bbb{\*\atrfend 7}{\*\atnid MM}{\*\atnauthor Max Mustermann}\chatn{\*\annotation{\*\atnref 7}{\*\atndate 667322855}{\*\atnparent root}{\*\atnicn 2}{\*\atntime 42}Comment \u20320?} ccc}"#;
    let doc = RtfDocument::parse(rtf).unwrap();
    assert_eq!(doc.text(), "aaa bbb ccc");
    assert_eq!(doc.annotations().len(), 1);

    let annotation = &doc.annotations()[0];
    assert_eq!(annotation.id, 7);
    assert_eq!(annotation.author, "Max Mustermann");
    assert_eq!(annotation.initials, "MM");
    assert_eq!(annotation.date.as_deref(), Some("667322855"));
    assert_eq!(annotation.text, "Comment 你");
    assert_eq!(annotation.position, "aaa ".len());
    assert_eq!(annotation.range_end, "aaa bbb".len());
    assert_eq!(annotation.parent_id.as_deref(), Some("root"));
    assert_eq!(annotation.icon.as_deref(), Some("2"));
    assert_eq!(annotation.time.as_deref(), Some("42"));
}

#[test]
fn preserves_parsed_headers_and_footers_in_owned_document() {
    let rtf = r#"{\rtf1\ansi\sectd\sbkeven\pgwsxn10000\pghsxn14000\marglsxn900\margrsxn800\margtsxn700\margbsxn600\guttersxn120\headery300\footery400\lndscpsxn\cols2\colsx360\pgnstarts5\pgnucrm\vertalc\linemod1\lineppage{\header Main \u20320? header\par Second line}{\footer Page footer}Body}"#;
    let doc = RtfDocument::parse(rtf).unwrap();
    assert_eq!(doc.text(), "Body");
    assert_eq!(doc.sections().len(), 1);
    let section = &doc.sections()[0];
    assert_eq!(
        section.properties.break_type,
        crate::SectionBreakType::EvenPage
    );
    assert_eq!(section.properties.page_width, 10000);
    assert_eq!(section.properties.page_height, 14000);
    assert_eq!(section.properties.margin_left, 900);
    assert_eq!(section.properties.margin_right, 800);
    assert_eq!(section.properties.margin_top, 700);
    assert_eq!(section.properties.margin_bottom, 600);
    assert_eq!(section.properties.margin_gutter, 120);
    assert_eq!(section.properties.header_distance, 300);
    assert_eq!(section.properties.footer_distance, 400);
    assert_eq!(
        section.properties.orientation,
        crate::PageOrientation::Landscape
    );
    assert_eq!(section.properties.columns.count, 2);
    assert_eq!(section.properties.columns.default_spacing, 360);
    assert_eq!(section.properties.page_number_start, 5);
    assert_eq!(
        section.properties.page_number_format,
        crate::PageNumberFormat::UpperRoman
    );
    assert_eq!(
        section.properties.vertical_alignment,
        crate::VerticalAlignment::Center
    );
    assert!(section.properties.line_numbering.is_enabled());
    assert_eq!(
        section.properties.line_numbering.restart,
        Some(crate::SectionLineNumberRestart::Page)
    );
    assert_eq!(
        section
            .get_header(super::super::section::HeaderFooterType::Header)
            .unwrap()
            .text(),
        "Main 你 header\nSecond line"
    );
    assert_eq!(
        section
            .get_header(super::super::section::HeaderFooterType::Footer)
            .unwrap()
            .text(),
        "Page footer"
    );
}

#[test]
fn decodes_hex_escapes_after_declared_codepage() {
    let cyrillic =
        RtfDocument::parse(r#"{\rtf1\ansi\ansicpg1251 \'cf\'f0\'e8\'e2\'e5\'f2}"#).unwrap();
    assert_eq!(cyrillic.text(), "Привет");

    let japanese = RtfDocument::parse(r#"{\rtf1\ansi\ansicpg932 \'82\'a0\'82\'a2}"#).unwrap();
    assert_eq!(japanese.text(), "あい");
}

#[test]
fn font_switches_select_exact_explicit_and_charset_codepages() {
    let explicit = RtfDocument::parse(
            r#"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0\fnil\cpg1252 Latin;}{\f1\fnil\cpg932 Japanese;}}\f0\'e9|\f1\'82\'a0|\plain\'e9}"#,
        )
        .unwrap();
    assert_eq!(explicit.text(), "é|あ|é");

    let inferred = RtfDocument::parse(
            r#"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0\fnil\fcharset0 Latin;}{\f1\fnil\fcharset128 Japanese;}}\f0\'e9|\f1\'82\'a0|\f0\'e9}"#,
        )
        .unwrap();
    assert_eq!(inferred.text(), "é|あ|é");

    let inherited = RtfDocument::parse(
            r#"{\rtf1\ansi\ansicpg1251{\fonttbl{\f0\fnil\fcharset1 Default;}}\f0\'cf\'f0\'e8\'e2\'e5\'f2}"#,
        )
        .unwrap();
    assert_eq!(inherited.text(), "Привет");

    let override_page = RtfDocument::parse(
        r#"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0\fnil\fcharset78\cpg932 Japanese;}}\f0\'82\'a0}"#,
    )
    .unwrap();
    assert_eq!(override_page.text(), "あ");

    assert!(
        RtfDocument::parse(
            r#"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0\fnil\fcharset78 Japanese;}}\f0\'82\'a0}"#,
        )
        .is_err()
    );
}

#[test]
fn rejects_unsupported_or_wide_ansi_codepages() {
    for page in [-1, 1200, 65000, 99_999] {
        let source = format!(r"{{\rtf1\ansi\ansicpg{page} text}}");
        assert!(RtfDocument::parse(&source).is_err(), "page {page}");
    }
}

#[test]
fn decodes_macintosh_and_exact_dos_character_sets() {
    let scoped = RtfDocument::parse(r#"{\rtf1\ansi \'80{\mac \'80}\'80}"#).unwrap();
    assert_eq!(scoped.text(), "€Ä€");

    let cp437 = RtfDocument::parse(r#"{\rtf1\pc \'9b}"#).unwrap();
    assert_eq!(cp437.text(), "¢");
    let cp850 = RtfDocument::parse(r#"{\rtf1\pca \'9b}"#).unwrap();
    assert_eq!(cp850.text(), "ø");
    let explicit_cp437 = RtfDocument::parse(r#"{\rtf1\ansi\ansicpg437 \'9b}"#).unwrap();
    assert_eq!(explicit_cp437.text(), "¢");
}

#[test]
fn decodes_unescaped_legacy_bytes_and_semantic_control_symbols() {
    let mut bytes = br#"{\rtf1\ansi\ansicpg1252 "#.to_vec();
    bytes.push(0xE9);
    bytes.push(b'}');
    assert_eq!(RtfDocument::parse_bytes(&bytes).unwrap().text(), "é");

    let symbols = RtfDocument::parse(r#"{\rtf1\pc A\~B\-C\_D}"#).unwrap();
    assert_eq!(symbols.text(), "A\u{00A0}B\u{00AD}C\u{2011}D");
}

#[test]
fn decodes_declared_codepage_in_semantic_destinations() {
    let doc = RtfDocument::parse(
            r#"{\rtf1\ansi\ansicpg1251{\info{\author \'cf\'f0\'e8\'e2\'e5\'f2\~X}}{\*\revtbl {\'cf\'f0\'e8\'e2\'e5\'f2;}}{\header \'cf\'f0\'e8\'e2\'e5\'f2\_X}Body{\footnote \'cf\'f0\'e8\'e2\'e5\'f2\~X}{\revised\revauth0 X}}"#,
        )
        .unwrap();

    assert_eq!(doc.info().author.as_deref(), Some("Привет\u{00A0}X"));
    assert_eq!(
        doc.sections()[0].headers_footers[0].text(),
        "Привет\u{2011}X"
    );
    assert_eq!(doc.notes()[0].content, "Привет\u{00A0}X");
    assert_eq!(doc.revisions()[0].author, "Привет");
}
