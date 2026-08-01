use litchi_odf::{
    CustomPresentationShow, MutablePresentation, Presentation, PresentationBuilder,
    PresentationPageMetadata, PresentationPageMetadataCollection, PresentationSettings,
};

fn settings(start: &str, pages: &[&str]) -> PresentationSettings {
    PresentationSettings {
        start_page: Some(start.to_string()),
        show: Some("Review".to_string()),
        custom_shows: vec![
            CustomPresentationShow::new(
                "Review",
                pages.iter().map(|value| (*value).to_string()).collect(),
            )
            .unwrap(),
        ],
        ..PresentationSettings::default()
    }
}

fn metadata(names: &[&str]) -> PresentationPageMetadataCollection {
    PresentationPageMetadataCollection::new(
        names
            .iter()
            .enumerate()
            .map(|(slide_index, name)| {
                let mut page = PresentationPageMetadata::new(slide_index);
                page.name = Some((*name).to_string());
                page
            })
            .collect(),
    )
    .unwrap()
}

#[test]
fn builder_rejects_dangling_and_ambiguous_references_but_allows_repeats() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide("one").unwrap();
    builder
        .set_settings(Some(settings("missing", &["page1"])))
        .unwrap();
    assert!(builder.build().is_err());

    let mut builder = PresentationBuilder::new();
    builder.add_slide("one").unwrap();
    builder
        .set_settings(Some(settings("page1", &["missing"])))
        .unwrap();
    assert!(builder.build().is_err());

    let mut builder = PresentationBuilder::new();
    builder.add_slide("one").unwrap();
    builder
        .set_settings(Some(settings("page1", &["page1", "page1"])))
        .unwrap();
    assert!(builder.build().is_ok());

    let mut builder = PresentationBuilder::new();
    builder.add_slide("one").unwrap();
    builder
        .set_page_metadata(Some(metadata(&["duplicate"])))
        .unwrap();
    builder.add_slide("two").unwrap();
    builder
        .set_page_metadata(Some(metadata(&["duplicate", "duplicate"])))
        .unwrap();
    builder
        .set_settings(Some(settings("duplicate", &["duplicate"])))
        .unwrap();
    assert!(builder.build().is_err());
}

#[test]
fn direct_settings_and_metadata_edits_fail_only_at_final_serialization() {
    let mut builder = PresentationBuilder::new();
    builder.add_slide("one").unwrap().add_slide("two").unwrap();
    builder
        .set_settings(Some(settings("page2", &["page2"])))
        .unwrap();
    let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();

    let mut mutable = MutablePresentation::from_presentation(presentation).unwrap();
    mutable.settings_mut().unwrap().start_page = Some("missing".to_string());
    assert!(mutable.to_bytes().is_err());
    mutable.settings_mut().unwrap().start_page = Some("page2".to_string());
    mutable
        .set_page_metadata(Some(metadata(&["page1", "renamed"])))
        .unwrap();
    assert!(mutable.to_bytes().is_err());
    assert_eq!(mutable.slides().len(), 2);
}

#[test]
fn fallback_insert_remove_preserves_identity_and_rejects_referenced_removal() {
    let mut mutable = MutablePresentation::new();
    mutable.add_slide("one", "1").unwrap();
    mutable.add_slide("two", "2").unwrap();
    mutable.add_slide("three", "3").unwrap();
    mutable
        .set_settings(Some(settings("page2", &["page3", "page2", "page3"])))
        .unwrap();

    mutable.insert_slide(1, "inserted", "new").unwrap();
    let names = mutable
        .page_metadata()
        .unwrap()
        .pages()
        .iter()
        .map(|page| page.name.as_deref().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(names, ["page1", "page4", "page2", "page3"]);
    assert_eq!(mutable.slides()[2].title.as_deref(), Some("two"));

    let before_titles = mutable
        .slides()
        .iter()
        .map(|slide| slide.title.clone())
        .collect::<Vec<_>>();
    assert!(mutable.remove_slide(2).is_err());
    assert_eq!(
        mutable
            .slides()
            .iter()
            .map(|slide| slide.title.clone())
            .collect::<Vec<_>>(),
        before_titles
    );

    assert_eq!(
        mutable.remove_slide(1).unwrap().title.as_deref(),
        Some("inserted")
    );
    let names = mutable
        .page_metadata()
        .unwrap()
        .pages()
        .iter()
        .map(|page| page.name.as_deref().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(names, ["page1", "page2", "page3"]);
    let reopened = Presentation::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.settings().unwrap(), mutable.settings().cloned());
}

#[test]
fn explicit_metadata_reindexes_without_losing_page_identity() {
    let mut builder = PresentationBuilder::new();
    builder
        .add_slide("alpha")
        .unwrap()
        .add_slide("beta")
        .unwrap();
    let mut pages = metadata(&["Alpha", "Beta"]);
    let mut records = pages.pages().to_vec();
    records[1].xml_id = Some("beta-id".to_string());
    records[1].draw_id = Some("beta-id".to_string());
    pages = PresentationPageMetadataCollection::new(records).unwrap();
    builder.set_page_metadata(Some(pages)).unwrap();
    builder
        .set_settings(Some(settings("Beta", &["Alpha", "Beta"])))
        .unwrap();
    let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
    let mut mutable = MutablePresentation::from_presentation(presentation).unwrap();
    mutable.insert_slide(0, "new", "new").unwrap();

    let pages = mutable.page_metadata().unwrap().pages();
    assert_eq!(pages[1].name.as_deref(), Some("Alpha"));
    assert_eq!(pages[2].name.as_deref(), Some("Beta"));
    assert_eq!(pages[2].xml_id.as_deref(), Some("beta-id"));
    assert_eq!(
        mutable.remove_slide(0).unwrap().title.as_deref(),
        Some("new")
    );
    let pages = mutable.page_metadata().unwrap().pages();
    assert_eq!(pages[0].name.as_deref(), Some("Alpha"));
    assert_eq!(pages[1].name.as_deref(), Some("Beta"));
    assert_eq!(pages[1].xml_id.as_deref(), Some("beta-id"));
}

#[test]
fn libreoffice_and_odfpy_lexical_fixtures_parse_save_and_reopen() {
    for fixture in [
        include_str!("fixtures/libreoffice-presentation-settings.xml"),
        include_str!("fixtures/odfpy-presentation-settings.xml"),
    ] {
        let parsed = litchi_odf::parse_presentation_settings(fixture)
            .unwrap()
            .unwrap();
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
        let mut builder = PresentationBuilder::new();
        for name in page_names {
            builder.add_slide(name).unwrap();
        }
        builder
            .set_page_metadata(Some(metadata(&page_names)))
            .unwrap();
        builder.set_settings(Some(parsed.clone())).unwrap();
        let reopened = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(reopened.settings().unwrap(), Some(parsed));
    }
}
