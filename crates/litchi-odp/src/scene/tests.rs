use super::change::plan;
use super::*;

const HEAD: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?>"#,
    r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" "#,
    r#"xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" "#,
    r#"xmlns:xlink="http://www.w3.org/1999/xlink" "#,
    r#"xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">"#,
    r#"<office:body><office:presentation>"#,
);
const TAIL: &str = "</office:presentation></office:body></office:document-content>";

fn document(pages: &str) -> String {
    format!("{HEAD}{pages}{TAIL}")
}

fn sample() -> String {
    document(concat!(
        r#"<draw:page draw:name="page1" draw:master-page-name="Default" presentation:presentation-page-layout-name="AL1T0">"#,
        r#"<draw:frame presentation:class="title" draw:name="Title" svg:x="2cm" svg:y="1cm" svg:width="20cm" svg:height="3cm">"#,
        r#"<draw:text-box><text:p text:style-name="P1">Hello <text:span text:style-name="T1">big</text:span> world</text:p></draw:text-box></draw:frame>"#,
        r#"<draw:frame presentation:class="outline" draw:name="Body"><draw:text-box>"#,
        r#"<text:list text:style-name="L1"><text:list-item><text:p>First</text:p></text:list-item>"#,
        r#"<text:list-item><text:p>Second<text:s text:c="2"/>gap<text:tab/>tab</text:p></text:list-item></text:list>"#,
        r#"</draw:text-box></draw:frame>"#,
        r#"<draw:custom-shape draw:name="Star" draw:id="id9"><text:p>Shape text</text:p><draw:enhanced-geometry draw:type="star5"/></draw:custom-shape>"#,
        r#"<draw:frame draw:name="Pic"><draw:image xlink:href="Pictures/a.png"><text:p/></draw:image><svg:title>Alt</svg:title></draw:frame>"#,
        r#"<draw:frame draw:name="Grid"><table:table table:name="T"><table:table-column table:number-columns-repeated="2"/>"#,
        r#"<table:table-row><table:table-cell><text:p>A1</text:p></table:table-cell><table:table-cell><text:p>B1</text:p></table:table-cell></table:table-row>"#,
        r#"</table:table></draw:frame>"#,
        r#"<draw:g><draw:rect draw:name="R"/><draw:ellipse draw:name="E"><text:p>Oval</text:p></draw:ellipse></draw:g>"#,
        r#"<presentation:notes><draw:page-thumbnail draw:page-number="1"/><draw:frame presentation:class="notes"><draw:text-box><text:p>Note one</text:p><text:p>Note two</text:p></draw:text-box></draw:frame></presentation:notes>"#,
        r#"</draw:page>"#,
        r#"<draw:page draw:name="page2" draw:master-page-name="Default"/>"#,
    ))
}

#[test]
fn indexes_pages_objects_and_stories() {
    let xml = sample();
    let scene = Scene::parse(&xml).unwrap();
    assert_eq!(scene.pages().len(), 2);
    let page = &scene.pages()[0];
    assert_eq!(page.name(), Some("page1"));
    assert_eq!(page.master_page(), Some("Default"));
    assert_eq!(page.layout(), Some("AL1T0"));
    assert_eq!(page.title().as_deref(), Some("Hello big world"));
    let objects = page.objects();
    assert_eq!(objects.len(), 6);
    let title = &objects[0];
    assert_eq!(title.class(), Some("title"));
    assert_eq!(title.content(), FrameContent::TextBox);
    assert_eq!(title.x(), Some("2cm"));
    let story = title.story().unwrap();
    assert!(story.replaceable);
    let paragraph = &story.paragraphs()[0];
    let runs: Vec<_> = paragraph
        .runs()
        .iter()
        .map(|run| (run.kind(), run.text()))
        .collect();
    assert_eq!(
        runs,
        [
            (RunKind::Text, "Hello"),
            (RunKind::Span, "big"),
            (RunKind::Text, "world"),
        ]
    );
    assert_eq!(&xml[paragraph.runs()[1].content.clone()], "big");

    let body = objects[1].story().unwrap();
    assert_eq!(body.paragraphs().len(), 2);
    assert_eq!(body.paragraphs()[0].level(), 1);
    assert_eq!(body.paragraphs()[1].text(), "Second  gap\ttab");
    assert!(xml[body.paragraphs()[0].removal.clone()].starts_with("<text:list-item>"));
    assert_eq!(body.text(), "First\nSecond  gap\ttab");

    let star = &objects[2];
    assert_eq!(star.kind(), ObjectKind::CustomShape);
    assert_eq!(star.id(), Some("id9"));
    assert_eq!(star.text(), "Shape text");
    assert_eq!(
        &xml[star.story().unwrap().range()],
        "<text:p>Shape text</text:p>"
    );

    let picture = &objects[3];
    assert_eq!(picture.content(), FrameContent::Image);
    assert_eq!(picture.href(), Some("Pictures/a.png"));
    assert_eq!(picture.title(), Some("Alt"));
    assert!(picture.story().is_none());

    let table = objects[4].table().unwrap();
    assert_eq!(table.columns(), 2);
    assert_eq!(table.rows()[0].cells()[1].text(), "B1");

    let group = &objects[5];
    assert_eq!(group.kind(), ObjectKind::Group);
    assert_eq!(group.children().len(), 2);
    assert_eq!(group.children()[1].text(), "Oval");

    let notes = page.notes().unwrap();
    assert_eq!(notes.objects().len(), 2);
    assert_eq!(notes.text().as_deref(), Some("Note one\nNote two"));

    assert!(scene.pages()[1].objects().is_empty());
    assert_eq!(scene.pages()[1].insertion, None);
}

fn apply(xml: &str, changes: &[Change]) -> String {
    let scene = Scene::parse(xml).unwrap();
    plan(xml, &scene, changes).unwrap().assemble(xml)
}

fn title(object: usize) -> StoryRef {
    StoryRef::object(ObjectRef::new(0, object))
}

#[test]
fn text_changes_splice_only_their_ranges() {
    let xml = sample();
    let changed = apply(
        &xml,
        &[
            Change::SetRunText {
                story: title(0),
                paragraph: 0,
                run: 1,
                text: "small".into(),
            },
            Change::SetParagraphText {
                story: title(1),
                paragraph: 1,
                text: "  lead and  double ".into(),
            },
            Change::SetStoryText {
                story: title(2),
                text: "One\nTwo".into(),
            },
        ],
    );
    let scene = Scene::parse(&changed).unwrap();
    let objects = scene.pages()[0].objects();
    assert_eq!(objects[0].text(), "Hello small world");
    assert_eq!(
        objects[1].story().unwrap().paragraphs()[1].text(),
        "  lead and  double "
    );
    assert_eq!(objects[2].text(), "One\nTwo");
    // Everything before the first change is untouched.
    let first = xml.find("big").unwrap();
    assert_eq!(&changed[..first], &xml[..first]);
}

#[test]
fn story_text_keeps_list_wrappers() {
    let xml = sample();
    let changed = apply(
        &xml,
        &[Change::SetStoryText {
            story: title(1),
            text: "a\nb\nc".into(),
        }],
    );
    let scene = Scene::parse(&changed).unwrap();
    let story = scene.pages()[0].objects()[1].story().unwrap();
    assert_eq!(story.paragraphs().len(), 3);
    assert!(
        story
            .paragraphs()
            .iter()
            .all(|paragraph| paragraph.level() == 1)
    );
    assert!(changed.contains(
        r#"<text:list text:style-name="L1"><text:list-item><text:p>a</text:p></text:list-item>"#
    ));
}

#[test]
fn paragraphs_insert_and_remove_with_their_list_items() {
    let xml = sample();
    let changed = apply(
        &xml,
        &[Change::InsertParagraph {
            story: title(1),
            index: 1,
            text: "Middle".into(),
        }],
    );
    let scene = Scene::parse(&changed).unwrap();
    let texts: Vec<_> = scene.pages()[0].objects()[1]
        .story()
        .unwrap()
        .paragraphs()
        .iter()
        .map(|paragraph| (paragraph.text().to_owned(), paragraph.level()))
        .collect();
    assert_eq!(
        texts,
        [
            ("First".to_owned(), 1),
            ("Middle".to_owned(), 1),
            ("Second  gap\ttab".to_owned(), 1)
        ]
    );
    let removed = apply(
        &changed,
        &[Change::RemoveParagraph {
            story: title(1),
            paragraph: 0,
        }],
    );
    let scene = Scene::parse(&removed).unwrap();
    assert_eq!(
        scene.pages()[0].objects()[1].text(),
        "Middle\nSecond  gap\ttab"
    );
    assert!(!removed.contains("<text:list-item><text:p>First"));
}

#[test]
fn attributes_merge_into_one_start_tag() {
    let xml = sample();
    let object = ObjectRef::new(0, 0);
    let changed = apply(
        &xml,
        &[
            Change::SetAttribute {
                object: object.clone(),
                attribute: Attribute::X,
                value: Some("3cm".into()),
            },
            Change::SetAttribute {
                object: object.clone(),
                attribute: Attribute::Name,
                value: Some("Renamed".into()),
            },
            Change::SetAttribute {
                object: ObjectRef::new(0, 3),
                attribute: Attribute::Width,
                value: Some("5cm".into()),
            },
        ],
    );
    let scene = Scene::parse(&changed).unwrap();
    let objects = scene.pages()[0].objects();
    assert_eq!(objects[0].x(), Some("3cm"));
    assert_eq!(objects[0].name(), Some("Renamed"));
    assert_eq!(objects[3].width(), Some("5cm"));
}

#[test]
fn objects_and_pages_move_remove_and_insert() {
    let xml = sample();
    let changed = apply(
        &xml,
        &[
            Change::InsertTextBox {
                page: 0,
                text_box: NewTextBox {
                    name: Some("New".into()),
                    geometry: ["1cm".into(), "1cm".into(), "4cm".into(), "2cm".into()],
                    text: "Fresh".into(),
                },
            },
            Change::RemoveObject {
                object: ObjectRef::new(0, 3),
            },
        ],
    );
    let scene = Scene::parse(&changed).unwrap();
    let objects = scene.pages()[0].objects();
    assert_eq!(objects.len(), 6);
    assert_eq!(objects[5].name(), Some("New"));
    assert_eq!(objects[5].text(), "Fresh");
    assert!(objects.iter().all(|object| object.name() != Some("Pic")));

    let moved = apply(
        &xml,
        &[Change::MoveObject {
            object: ObjectRef::new(0, 0),
            to: 2,
        }],
    );
    let scene = Scene::parse(&moved).unwrap();
    let names: Vec<_> = scene.pages()[0]
        .objects()
        .iter()
        .map(|object| object.name().unwrap_or(""))
        .collect();
    assert_eq!(names, ["Body", "Star", "Title", "Pic", "Grid", ""]);

    let pages = apply(
        &xml,
        &[Change::InsertPage {
            index: 0,
            page: NewPage::default(),
        }],
    );
    let scene = Scene::parse(&pages).unwrap();
    assert_eq!(scene.pages().len(), 3);
    assert_eq!(scene.pages()[0].name(), Some("page3"));
    assert_eq!(scene.pages()[0].master_page(), Some("Default"));

    let swapped = apply(
        &xml,
        &[Change::SwapPages {
            first: 0,
            second: 1,
        }],
    );
    let scene = Scene::parse(&swapped).unwrap();
    assert_eq!(scene.pages()[0].name(), Some("page2"));
    assert_eq!(scene.pages()[1].name(), Some("page1"));

    let removed = apply(&xml, &[Change::RemovePage { page: 1 }]);
    assert_eq!(Scene::parse(&removed).unwrap().pages().len(), 1);
}

#[test]
fn references_and_overlaps_are_refused() {
    let xml = document(concat!(
        r#"<draw:page draw:name="A" draw:master-page-name="M"><draw:frame draw:name="Link"><draw:text-box>"#,
        r##"<text:p><text:a xlink:href="#B">go</text:a></text:p></draw:text-box></draw:frame></draw:page>"##,
        r#"<draw:page draw:name="B" draw:master-page-name="M"/>"#,
    ));
    let scene = Scene::parse(&xml).unwrap();
    let error = plan(&xml, &scene, &[Change::RemovePage { page: 1 }]).unwrap_err();
    assert!(
        matches!(error, litchi_core::Error::Unsupported(_)),
        "{error:?}"
    );
    let error = plan(
        &xml,
        &scene,
        &[
            Change::SetStoryText {
                story: StoryRef::object(ObjectRef::new(0, 0)),
                text: "x".into(),
            },
            Change::SetParagraphText {
                story: StoryRef::object(ObjectRef::new(0, 0)),
                paragraph: 0,
                text: "y".into(),
            },
        ],
    )
    .unwrap_err();
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(_)),
        "{error:?}"
    );
    // An exact no-op plans nothing.
    let noop = plan(
        &xml,
        &scene,
        &[Change::SetStoryText {
            story: StoryRef::object(ObjectRef::new(0, 0)),
            text: "go".into(),
        }],
    )
    .unwrap();
    assert!(noop.is_empty());
}

#[test]
fn byte_order_mark_offsets_ranges() {
    let xml = format!("\u{feff}{}", sample());
    let scene = Scene::parse(&xml).unwrap();
    let paragraph = &scene.pages()[0].objects()[0].story().unwrap().paragraphs()[0];
    assert!(xml[paragraph.range()].starts_with("<text:p"));
}

/// Record 0295: a field run's displayed text is replaced in place, its
/// element and attributes kept; a text run, and a run that does not exist
/// (bookmarks are not runs), are refused.
#[test]
fn field_text_keeps_the_field_element() {
    let xml = document(concat!(
        r#"<draw:page draw:name="p">"#,
        r#"<draw:frame draw:name="F"><draw:text-box><text:p>No. <text:page-number text:select-page="current">7</text:page-number><text:page-count>2</text:page-count><text:bookmark text:name="b"/></text:p></draw:text-box></draw:frame>"#,
        r#"</draw:page>"#,
    ));
    let changed = apply(
        &xml,
        &[
            Change::SetFieldText {
                story: title(0),
                paragraph: 0,
                run: 1,
                text: "1".into(),
            },
            Change::SetFieldText {
                story: title(0),
                paragraph: 0,
                run: 2,
                text: "3".into(),
            },
        ],
    );
    assert!(
        changed.contains(r#"<text:page-number text:select-page="current">1</text:page-number><text:page-count>3</text:page-count>"#),
        "{changed}"
    );
    let scene = Scene::parse(&xml).unwrap();
    for run in [0, 3] {
        assert!(
            plan(
                &xml,
                &scene,
                &[Change::SetFieldText {
                    story: title(0),
                    paragraph: 0,
                    run,
                    text: "x".into(),
                }],
            )
            .is_err(),
            "run {run}"
        );
    }
}
