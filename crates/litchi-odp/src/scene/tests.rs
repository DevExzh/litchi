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
        r##"<text:p><text:a xlink:href="#B">go</text:a></text:p></draw:text-box></draw:frame>"##,
        r#"<draw:frame draw:name="Plain"><draw:text-box><text:p>plain</text:p></draw:text-box></draw:frame></draw:page>"#,
        r#"<draw:page draw:name="B" draw:master-page-name="M"/>"#,
    ));
    let scene = Scene::parse(&xml).unwrap();
    // A text set over the link would drop it (record 0304).
    let error = plan(
        &xml,
        &scene,
        &[Change::SetStoryText {
            story: StoryRef::object(ObjectRef::new(0, 0)),
            text: "x".into(),
        }],
    )
    .unwrap_err();
    assert!(
        matches!(&error, litchi_core::Error::Refused(_)),
        "{error:?}"
    );
    let error = plan(&xml, &scene, &[Change::RemovePage { page: 1 }]).unwrap_err();
    assert!(
        matches!(
            error,
            litchi_core::Error::Unsupported(_) | litchi_core::Error::Refused(_)
        ),
        "{error:?}"
    );
    let error = plan(
        &xml,
        &scene,
        &[
            Change::SetStoryText {
                story: StoryRef::object(ObjectRef::new(0, 1)),
                text: "x".into(),
            },
            Change::SetParagraphText {
                story: StoryRef::object(ObjectRef::new(0, 1)),
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

/// The sample with style and fo bindings and two automatic styles: the
/// title paragraph's `P1` (no text properties) and the span's `T1` (bold).
fn styled() -> String {
    sample()
        .replacen(
            r#"xmlns:table="#,
            concat!(
                r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
                r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" "#,
                r#"xmlns:table="#
            ),
            1,
        )
        .replacen(
            "<office:body>",
            concat!(
                r#"<office:automatic-styles>"#,
                r#"<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>"#,
                r##"<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:color="#112233"/></style:style>"##,
                r#"</office:automatic-styles><office:body>"#
            ),
            1,
        )
}

fn edit(build: impl FnOnce(&mut TextPropertyUpdate)) -> TextPropertyUpdate {
    let mut properties = TextPropertyUpdate::new();
    build(&mut properties);
    properties
}

#[test]
fn automatic_styles_are_indexed() {
    let xml = styled();
    let scene = Scene::parse(&xml).unwrap();
    let styles = scene.automatic_styles(&xml).unwrap();
    assert_eq!(styles.as_slice().len(), 2);
    let span = &scene.pages()[0].objects()[0].story().unwrap().paragraphs()[0].runs()[1];
    let properties = styles.run_text_properties(span).unwrap();
    assert_eq!(properties.bold(), Some(true));
    assert_eq!(properties.color(), Some("#112233"));
    assert_eq!(properties.italic(), None);
    let paragraph = &scene.pages()[0].objects()[0].story().unwrap().paragraphs()[0];
    assert_eq!(
        styles.paragraph_text_properties(paragraph),
        Some(&TextProperties::default())
    );
    assert!(styles.get("text", "P1").is_none());
}

#[test]
fn span_formatting_derives_a_new_style() {
    let xml = styled();
    let changed = apply(
        &xml,
        &[Change::SetTextProperties {
            story: title(0),
            paragraph: 0,
            run: Some(1),
            properties: edit(|properties| {
                properties.italic = Some(true);
                properties.size = Some("24pt".into());
            }),
        }],
    );
    let scene = Scene::parse(&changed).unwrap();
    let styles = scene.automatic_styles(&changed).unwrap();
    let span = &scene.pages()[0].objects()[0].story().unwrap().paragraphs()[0].runs()[1];
    assert_ne!(span.style_name(), Some("T1"));
    let properties = styles.run_text_properties(span).unwrap();
    // Kept from T1, and added.
    assert_eq!(properties.bold(), Some(true));
    assert_eq!(properties.color(), Some("#112233"));
    assert_eq!(properties.italic(), Some(true));
    assert_eq!(properties.size(), Some("24pt"));
    // T1 itself is unchanged.
    assert_eq!(
        styles.get("text", "T1").unwrap().text_properties().italic(),
        None
    );
    assert_eq!(scene.pages()[0].objects()[0].text(), "Hello big world");
}

#[test]
fn plain_text_runs_are_wrapped_and_paragraphs_restyled() {
    let xml = styled();
    let changed = apply(
        &xml,
        &[
            Change::SetTextProperties {
                story: title(0),
                paragraph: 0,
                run: Some(0),
                properties: edit(|properties| properties.underline = Some(true)),
            },
            Change::SetRunText {
                story: title(0),
                paragraph: 0,
                run: 0,
                text: "Hi".into(),
            },
            Change::SetTextProperties {
                story: title(0),
                paragraph: 0,
                run: None,
                properties: edit(|properties| properties.font = Some("Liberation Sans".into())),
            },
            Change::SetTextProperties {
                story: title(1),
                paragraph: 0,
                run: None,
                properties: edit(|properties| properties.bold = Some(true)),
            },
        ],
    );
    let scene = Scene::parse(&changed).unwrap();
    let styles = scene.automatic_styles(&changed).unwrap();
    let title_story = scene.pages()[0].objects()[0].story().unwrap();
    let paragraph = &title_story.paragraphs()[0];
    let first = &paragraph.runs()[0];
    assert_eq!(first.text(), "Hi");
    assert_eq!(first.kind(), RunKind::Span);
    assert_eq!(
        styles.run_text_properties(first).unwrap().underline(),
        Some(true)
    );
    let restyled = styles.paragraph_text_properties(paragraph).unwrap();
    assert_eq!(restyled.font(), Some("Liberation Sans"));
    // The derived paragraph style keeps P1's paragraph properties.
    let derived = styles
        .get("paragraph", paragraph.style_name().unwrap())
        .unwrap();
    assert!(changed[derived.range()].contains(r#"fo:text-align="center""#));
    // A paragraph with no style gets one with no parent.
    let body = &scene.pages()[0].objects()[1].story().unwrap().paragraphs()[0];
    let style = styles.get("paragraph", body.style_name().unwrap()).unwrap();
    assert_eq!(style.parent(), None);
    assert_eq!(style.text_properties().bold(), Some(true));
}

#[test]
fn formatting_no_ops_and_refusals() {
    let xml = styled();
    let scene = Scene::parse(&xml).unwrap();
    // T1 is already bold: an exact no-op.
    let plan_bold = plan(
        &xml,
        &scene,
        &[Change::SetTextProperties {
            story: title(0),
            paragraph: 0,
            run: Some(1),
            properties: edit(|properties| properties.bold = Some(true)),
        }],
    )
    .unwrap();
    assert!(plan_bold.is_empty());
    let invalid = plan(
        &xml,
        &scene,
        &[Change::SetTextProperties {
            story: title(0),
            paragraph: 0,
            run: None,
            properties: edit(|properties| properties.color = Some("red".into())),
        }],
    );
    assert!(matches!(invalid, Err(litchi_core::Error::InvalidFormat(_))));
    // Content without automatic styles gains the element before the body.
    let bare = sample().replacen(
        r#"xmlns:table="#,
        concat!(
            r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
            r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" "#,
            r#"xmlns:table="#
        ),
        1,
    );
    let changed = apply(
        &bare,
        &[Change::SetTextProperties {
            story: title(0),
            paragraph: 0,
            run: Some(1),
            properties: edit(|properties| properties.bold = Some(true)),
        }],
    );
    assert!(changed.contains("<office:automatic-styles><style:style style:name=\"T1\""));
    // Without the style namespace bound on the root, the edit is refused.
    let unbound = Scene::parse(&sample()).unwrap();
    let refused = plan(
        &sample(),
        &unbound,
        &[Change::SetTextProperties {
            story: title(0),
            paragraph: 0,
            run: Some(1),
            properties: edit(|properties| properties.bold = Some(true)),
        }],
    );
    assert!(matches!(
        refused,
        Err(litchi_core::Error::Unsupported(_) | litchi_core::Error::Refused(_))
    ));
}

/// A styled page whose one text box holds a link and a page-number field.
fn links_and_fields() -> String {
    styled().replacen(
        r#"<draw:custom-shape draw:name="Star""#,
        concat!(
            r#"<draw:frame draw:name="Links"><draw:text-box><text:p>"#,
            r#"<text:a xlink:href="https://example.com" text:style-name="Internet_20_link">site</text:a>"#,
            r#" page <text:page-number>4</text:page-number>"#,
            r#"</text:p></draw:text-box></draw:frame>"#,
            r#"<draw:custom-shape draw:name="Star""#,
        ),
        1,
    )
}

fn links(object: usize) -> StoryRef {
    StoryRef::object(ObjectRef::new(0, object))
}

#[test]
fn link_and_field_runs_take_formatting() {
    let xml = links_and_fields();
    let scene = Scene::parse(&xml).unwrap();
    let story = scene.story(&links(2)).unwrap();
    let kinds: Vec<_> = story.paragraphs()[0].runs().iter().map(Run::kind).collect();
    assert_eq!(kinds, [RunKind::Link, RunKind::Text, RunKind::Field]);
    let bold = edit(|properties| properties.bold = Some(true));
    let changed = apply(
        &xml,
        &[
            Change::SetTextProperties {
                story: links(2),
                paragraph: 0,
                run: Some(0),
                properties: bold.clone(),
            },
            Change::SetTextProperties {
                story: links(2),
                paragraph: 0,
                run: Some(2),
                properties: bold.clone(),
            },
        ],
    );
    // The link keeps its link style; a span inside it formats the text,
    // and a span around the field formats the field.
    assert!(
        changed.contains(r#"text:style-name="Internet_20_link"><text:span text:style-name="T3">site</text:span></text:a>"#),
        "{changed}"
    );
    assert!(
        changed.contains(
            r#"<text:span text:style-name="T4"><text:page-number>4</text:page-number></text:span>"#
        ),
        "{changed}"
    );
    let scene = Scene::parse(&changed).unwrap();
    let styles = scene.automatic_styles(&changed).unwrap();
    let common = CommonStyles::default();
    let runs = scene.story(&links(2)).unwrap().paragraphs()[0]
        .runs()
        .to_vec();
    assert_eq!(runs[0].kind(), RunKind::Link);
    assert_eq!(runs[0].href(), Some("https://example.com"));
    assert_eq!(runs[0].link_text_style_name(), Some("T3"));
    assert_eq!(runs[2].kind(), RunKind::Field);
    assert_eq!(runs[2].style_name(), Some("T4"));
    for run in [&runs[0], &runs[2]] {
        assert_eq!(styles.resolve_run(&common, run).bold(), Some(true));
    }
    // The formatted field still refreshes, inside its span.
    let refreshed = apply(
        &changed,
        &[Change::SetFieldText {
            story: links(2),
            paragraph: 0,
            run: 2,
            text: "5".into(),
        }],
    );
    assert!(refreshed.contains(
        r#"<text:span text:style-name="T4"><text:page-number>5</text:page-number></text:span>"#
    ));
    // A second format derives from the link's text span.
    let italic = apply(
        &changed,
        &[Change::SetTextProperties {
            story: links(2),
            paragraph: 0,
            run: Some(0),
            properties: edit(|properties| properties.italic = Some(true)),
        }],
    );
    let scene = Scene::parse(&italic).unwrap();
    let styles = scene.automatic_styles(&italic).unwrap();
    let link = &scene.story(&links(2)).unwrap().paragraphs()[0].runs()[0];
    let resolved = styles.resolve_run(&common, link);
    assert_eq!(
        (resolved.bold(), resolved.italic()),
        (Some(true), Some(true))
    );
    assert_eq!(
        italic.matches("<text:span").count(),
        changed.matches("<text:span").count()
    );
}

#[test]
fn formatting_resolves_through_common_styles() {
    let styles_xml = concat!(
        r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
        r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:styles>"#,
        r#"<style:default-style style:family="paragraph"><style:text-properties fo:font-size="18pt"/></style:default-style>"#,
        r#"<style:style style:name="Base" style:family="paragraph"><style:text-properties fo:font-style="italic" style:font-name="Liberation Sans"/></style:style>"#,
        r#"<style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Base"><style:text-properties fo:font-weight="bold"/></style:style>"#,
        r#"<style:style style:name="Loop" style:family="paragraph" style:parent-style-name="Loop"/>"#,
        r#"</office:styles></office:document-styles>"#,
    );
    let common = CommonStyles::parse(styles_xml).unwrap();
    // A byte-order mark is skipped; the styles are the same.
    let marked = CommonStyles::parse(&format!("\u{feff}{styles_xml}")).unwrap();
    let none = AutomaticStyles::default();
    assert_eq!(
        none.resolve(&marked, "paragraph", Some("Heading")),
        none.resolve(&common, "paragraph", Some("Heading"))
    );
    let xml = styled().replacen(
        r#"<style:style style:name="P1" style:family="paragraph">"#,
        r#"<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Heading">"#,
        1,
    );
    let scene = Scene::parse(&xml).unwrap();
    let styles = scene.automatic_styles(&xml).unwrap();
    let paragraph = &scene.story(&title(0)).unwrap().paragraphs()[0];
    let resolved = styles.resolve_paragraph(&common, paragraph);
    // Heading's bold, Base's italic and font, the default size.
    assert_eq!(resolved.bold(), Some(true));
    assert_eq!(resolved.italic(), Some(true));
    assert_eq!(resolved.font(), Some("Liberation Sans"));
    assert_eq!(resolved.size(), Some("18pt"));
    // The local properties alone are unchanged.
    assert_eq!(
        styles.paragraph_text_properties(paragraph),
        Some(&TextProperties::default())
    );
    // A cycle ends; a missing style resolves to the defaults.
    assert_eq!(
        styles.resolve(&common, "paragraph", Some("Loop")).size(),
        Some("18pt")
    );
    assert_eq!(
        styles.resolve(&common, "paragraph", Some("Missing")).bold(),
        None
    );
    // A document without office:styles has none.
    assert_eq!(CommonStyles::parse(&xml).unwrap(), CommonStyles::default());
}

/// A page with the record-0328 elements: a link run, line breaks, a click
/// listener, a `draw:a` link, a plain shape, declared columns, a plugin and
/// an inline formula.
fn elements() -> String {
    let head = HEAD.replace(
        "xmlns:presentation=",
        concat!(
            r#"xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" "#,
            r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
            r#"xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" "#,
            r#"xmlns:presentation="#
        ),
    );
    let head = head.replace(
        "<office:body>",
        concat!(
            r#"<office:automatic-styles><style:style style:name="co1" style:family="table-column">"#,
            r#"<style:table-column-properties style:column-width="3cm"/></style:style></office:automatic-styles>"#,
            r#"<office:body>"#
        ),
    );
    format!(
        "{head}{}{TAIL}",
        concat!(
            r#"<draw:page draw:name="One">"#,
            r#"<draw:frame draw:name="Text"><draw:text-box>"#,
            r#"<text:p>See <text:a xlink:type="simple" xlink:href="https://a.example/" office:title="A">a</text:a> and b</text:p>"#,
            r#"<text:p>One<text:line-break/>Two<text:span>x<text:line-break/>y</text:span></text:p>"#,
            r#"</draw:text-box></draw:frame>"#,
            r#"<draw:custom-shape draw:name="Button"><office:event-listeners><presentation:event-listener script:event-name="dom:click" presentation:action="show" xlink:href="https://b.example/"/></office:event-listeners><text:p>Go</text:p></draw:custom-shape>"#,
            r#"<draw:a xlink:type="simple" xlink:href="https://c.example/"><draw:frame draw:name="Linked"><draw:text-box><text:p>Box</text:p></draw:text-box></draw:frame></draw:a>"#,
            r#"<draw:custom-shape draw:name="Plain"><svg:title>Alt</svg:title><text:p>Plain</text:p></draw:custom-shape>"#,
            r#"<draw:frame draw:name="Grid"><table:table><table:table-column table:style-name="co1" table:number-columns-repeated="2"/><table:table-row><table:table-cell><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell></table:table-row></table:table></draw:frame>"#,
            r#"<draw:frame draw:name="Movie"><draw:plugin xlink:href="Media/m.webm" draw:mime-type="application/vnd.sun.star.media"><draw:param draw:name="Loop" draw:value="true"/></draw:plugin></draw:frame>"#,
            r#"<draw:frame draw:name="F"><draw:object><math xmlns="http://www.w3.org/1998/Math/MathML"><semantics><mi>z</mi><annotation encoding="StarMath 5.0">z</annotation></semantics></math></draw:object></draw:frame>"#,
            r#"</draw:page><draw:page draw:name="Two"/>"#,
        )
    )
}

#[test]
fn parts_read_links_breaks_columns_media_and_formulas() {
    let xml = elements();
    let scene = Scene::parse(&xml).unwrap();
    let objects = scene.pages()[0].objects();
    let story = objects[0].story().unwrap();
    let link = &story.paragraphs()[0].runs()[1];
    assert_eq!(link.kind(), RunKind::Link);
    assert_eq!(parts::link_title(&xml, link).as_deref(), Some("A"));
    // Breaks inside the span count; the paragraph has two.
    assert_eq!(parts::line_breaks(&xml, &story.paragraphs()[1]).len(), 2);
    let click = parts::shape_click(&xml, &objects[1]).unwrap();
    assert_eq!(click.action(), Some("show"));
    assert_eq!(click.href(), Some("https://b.example/"));
    let wrapped = parts::shape_click(&xml, &objects[2]).unwrap();
    assert!(wrapped.is_wrapper());
    assert_eq!(wrapped.href(), Some("https://c.example/"));
    assert!(parts::shape_click(&xml, &objects[3]).is_none());
    let table = objects[4].table().unwrap();
    let columns = parts::table_columns(&xml, table);
    assert_eq!(columns.len(), 1);
    assert_eq!(columns[0].repeat(), 2);
    let styles = scene.automatic_styles(&xml).unwrap();
    assert_eq!(columns[0].width(&xml, &styles).as_deref(), Some("3cm"));
    let plugin = parts::plugin(&xml, &objects[5]).unwrap();
    assert_eq!(plugin.media_kind(), Some(parts::MediaKind::Video));
    assert_eq!(plugin.parameter("loop"), Some("true"));
    let math = parts::inline_math(&xml, &objects[6]).unwrap();
    assert_eq!(parts::formula_text(math).as_deref(), Some("z"));
}

#[test]
fn link_click_and_line_break_changes_splice_their_markup() {
    let xml = elements();
    let story = StoryRef::object(ObjectRef::new(0, 0));
    // Retarget a link and drop its title.
    let retargeted = apply(
        &xml,
        &[Change::SetRunLink {
            story: story.clone(),
            paragraph: 0,
            run: 1,
            href: Some("https://d.example/".to_owned()),
            title: Some(None),
        }],
    );
    assert!(
        retargeted
            .contains(r#"<text:a xlink:type="simple" xlink:href="https://d.example/">a</text:a>"#)
    );
    // Unwrap it.
    let unwrapped = apply(
        &xml,
        &[Change::SetRunLink {
            story: story.clone(),
            paragraph: 0,
            run: 1,
            href: None,
            title: None,
        }],
    );
    assert!(unwrapped.contains("<text:p>See a and b</text:p>"));
    // Link a plain run.
    let linked = apply(
        &xml,
        &[Change::LinkRun {
            story: story.clone(),
            paragraph: 0,
            run: 0,
            href: "https://e.example/".to_owned(),
            title: None,
        }],
    );
    assert!(linked.contains(
        r#"<text:p><text:a xlink:type="simple" xlink:href="https://e.example/">See </text:a>"#
    ));
    // Line breaks.
    let added = apply(
        &xml,
        &[Change::InsertLineBreak {
            story: story.clone(),
            paragraph: 0,
        }],
    );
    assert!(added.contains("and b<text:line-break/></text:p>"));
    let removed = apply(
        &xml,
        &[Change::RemoveLineBreak {
            story: story.clone(),
            paragraph: 1,
            index: 1,
        }],
    );
    assert!(removed.contains("<text:span>xy</text:span>"));
    // Click actions: retarget a listener, add one after the title, remove a
    // draw:a link.
    let retargeted = apply(
        &xml,
        &[Change::SetShapeClick {
            object: ObjectRef::new(0, 1),
            click: Some(ShapeAction::Page("Two".to_owned())),
        }],
    );
    assert!(retargeted.contains(r##"presentation:action="show" xlink:href="#Two""##));
    let added = apply(
        &xml,
        &[Change::SetShapeClick {
            object: ObjectRef::new(0, 3),
            click: Some(ShapeAction::NextPage),
        }],
    );
    assert!(added.contains(concat!(
        r#"<svg:title>Alt</svg:title><office:event-listeners><presentation:event-listener "#,
        r#"script:event-name="dom:click" presentation:action="next-page"/></office:event-listeners><text:p>Plain"#
    )));
    let removed = apply(
        &xml,
        &[Change::SetShapeClick {
            object: ObjectRef::new(0, 1),
            click: None,
        }],
    );
    assert!(removed.contains(r#"<draw:custom-shape draw:name="Button"><text:p>Go</text:p>"#));
    let unwrapped = apply(
        &xml,
        &[Change::SetShapeClick {
            object: ObjectRef::new(0, 2),
            click: None,
        }],
    );
    assert!(!unwrapped.contains("<draw:a"));
    assert!(unwrapped.contains(r#"<draw:frame draw:name="Linked">"#));
    // The same target is an exact no-op.
    let scene = Scene::parse(&xml).unwrap();
    let noop = plan(
        &xml,
        &scene,
        &[Change::SetShapeClick {
            object: ObjectRef::new(0, 1),
            click: Some(ShapeAction::Url("https://b.example/".to_owned())),
        }],
    )
    .unwrap();
    assert!(noop.is_empty());
}
