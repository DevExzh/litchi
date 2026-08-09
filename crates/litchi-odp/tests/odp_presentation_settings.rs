#![allow(
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::{
    Builder, Presentation, edit, page,
    show::{self, CustomShow, Settings},
};

fn settings(start: &str, pages: &[&str]) -> Settings {
    Settings {
        start_page: Some(start.to_string()),
        show: Some("Review".to_string()),
        custom_shows: vec![
            CustomShow::new(
                "Review",
                pages.iter().map(|value| (*value).to_string()).collect(),
            )
            .unwrap(),
        ],
        ..Settings::default()
    }
}

fn metadata(names: &[&str]) -> page::Collection {
    page::Collection::new(
        names
            .iter()
            .enumerate()
            .map(|(slide_index, name)| {
                let mut page = page::Page::new(slide_index);
                page.name = Some((*name).to_string());
                page
            })
            .collect(),
    )
    .unwrap()
}

#[test]
fn builder_rejects_dangling_and_ambiguous_references_but_allows_repeats() {
    let mut builder = Builder::new();
    builder.add_slide("one").unwrap();
    builder
        .set_settings(Some(settings("missing", &["page1"])))
        .unwrap();
    assert!(builder.build().is_err());

    let mut builder = Builder::new();
    builder.add_slide("one").unwrap();
    builder
        .set_settings(Some(settings("page1", &["missing"])))
        .unwrap();
    assert!(builder.build().is_err());

    let mut builder = Builder::new();
    builder.add_slide("one").unwrap();
    builder
        .set_settings(Some(settings("page1", &["page1", "page1"])))
        .unwrap();
    assert!(builder.build().is_ok());

    let mut builder = Builder::new();
    builder.add_slide("one").unwrap();
    builder.set_pages(Some(metadata(&["duplicate"]))).unwrap();
    builder.add_slide("two").unwrap();
    builder
        .set_pages(Some(metadata(&["duplicate", "duplicate"])))
        .unwrap();
    builder
        .set_settings(Some(settings("duplicate", &["duplicate"])))
        .unwrap();
    assert!(builder.build().is_err());
}

#[test]
fn source_checked_insert_remove_preserves_page_identity_and_references() {
    let mut builder = Builder::new();
    for title in ["one", "two", "three"] {
        builder.add_slide(title).unwrap();
    }
    builder
        .set_settings(Some(settings("page2", &["page3", "page2", "page3"])))
        .unwrap();
    let source = edit::Snapshot::from_bytes(builder.build().unwrap()).unwrap();
    let source_bytes = source.bytes().to_vec();
    let mut transaction = source.transaction().unwrap();
    transaction
        .add_before(1, "inserted", "new")
        .unwrap()
        .unwrap();
    let inserted = transaction.commit().unwrap();
    assert_eq!(source.bytes(), source_bytes);

    let presentation = inserted.snapshot().to_presentation().unwrap();
    let pages = presentation.pages().unwrap();
    let names = pages
        .pages()
        .iter()
        .map(|page| page.name.as_deref().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(names, ["page1", "page4", "page2", "page3"]);
    assert!(presentation.slides().unwrap()[2].all_text().contains("two"));

    let mut refused = inserted.snapshot().transaction().unwrap();
    assert!(refused.remove(2).is_err());
    let noop = refused.commit().unwrap();
    assert!(!noop.changed());
    assert_eq!(noop.snapshot().bytes(), inserted.snapshot().bytes());

    let mut remove = inserted.snapshot().transaction().unwrap();
    assert_eq!(
        remove.remove(1).unwrap().unwrap().title.as_deref(),
        Some("inserted")
    );
    let removed = remove.commit().unwrap();
    let presentation = removed.snapshot().to_presentation().unwrap();
    let pages = presentation.pages().unwrap();
    let names = pages
        .pages()
        .iter()
        .map(|page| page.name.as_deref().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(names, ["page1", "page2", "page3"]);
    assert_eq!(
        presentation.settings().unwrap(),
        Some(settings("page2", &["page3", "page2", "page3"]))
    );
    let restored = inserted
        .patch()
        .inverse()
        .apply(inserted.snapshot())
        .unwrap();
    assert_eq!(restored.bytes(), source.bytes());
}

#[test]
fn explicit_metadata_reindexes_without_losing_identity() {
    let mut builder = Builder::new();
    builder
        .add_slide("alpha")
        .unwrap()
        .add_slide("beta")
        .unwrap();
    let mut pages = metadata(&["Alpha", "Beta"]);
    let mut records = pages.pages().to_vec();
    records[1].xml_id = Some("beta-id".to_string());
    records[1].draw_id = Some("beta-id".to_string());
    pages = page::Collection::new(records).unwrap();
    builder.set_pages(Some(pages)).unwrap();
    builder
        .set_settings(Some(settings("Beta", &["Alpha", "Beta"])))
        .unwrap();
    let source = edit::Snapshot::from_bytes(builder.build().unwrap()).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction.add_before(0, "new", "new").unwrap().unwrap();
    let inserted = transaction.commit().unwrap();
    let presentation = inserted.snapshot().to_presentation().unwrap();
    let pages = presentation.pages().unwrap();
    assert_eq!(pages.pages()[1].name.as_deref(), Some("Alpha"));
    assert_eq!(pages.pages()[2].name.as_deref(), Some("Beta"));
    assert_eq!(pages.pages()[2].xml_id.as_deref(), Some("beta-id"));

    let mut transaction = inserted.snapshot().transaction().unwrap();
    transaction.remove(0).unwrap().unwrap();
    let removed = transaction.commit().unwrap();
    let presentation = removed.snapshot().to_presentation().unwrap();
    let pages = presentation.pages().unwrap();
    assert_eq!(pages.pages()[0].name.as_deref(), Some("Alpha"));
    assert_eq!(pages.pages()[1].name.as_deref(), Some("Beta"));
    assert_eq!(pages.pages()[1].xml_id.as_deref(), Some("beta-id"));
}

#[test]
fn libreoffice_and_odfpy_settings_fixtures_parse_save_and_reopen() {
    for fixture in [
        include_str!("fixtures/libreoffice-presentation-settings.xml"),
        include_str!("fixtures/odfpy-presentation-settings.xml"),
    ] {
        let parsed = show::parse(fixture).unwrap().unwrap();
        assert_eq!(parsed.custom_shows[0].pages.len(), 3);
        assert_eq!(
            parsed.custom_shows[0].pages[0],
            parsed.custom_shows[0].pages[2]
        );
        let page_names = if parsed.start_page.as_deref() == Some("Details") {
            ["Intro", "Details", "Summary"]
        } else {
            ["One", "Two", "Three"]
        };
        let mut builder = Builder::new();
        for name in page_names {
            builder.add_slide(name).unwrap();
        }
        builder.set_pages(Some(metadata(&page_names))).unwrap();
        builder.set_settings(Some(parsed.clone())).unwrap();
        let reopened = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(reopened.settings().unwrap(), Some(parsed));
    }
}
