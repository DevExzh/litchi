use litchi_odp::{Builder, Presentation};

#[test]
fn builder_and_presentation_facade_round_trip() {
    let mut builder = Builder::new();
    builder
        .add_slide_with_title("Welcome", "Hello from ODP")
        .unwrap();
    let bytes = builder.build().unwrap();
    let presentation = Presentation::from_bytes(bytes.clone()).unwrap();
    assert_eq!(presentation.slide_count().unwrap(), 1);
    assert_eq!(presentation.text().unwrap(), "Welcome\nHello from ODP");
    assert_eq!(presentation.to_bytes().unwrap(), bytes);
}

#[test]
fn presentation_facade_reports_empty_packages_without_slides() {
    let bytes = Builder::new().build().unwrap();
    let presentation = Presentation::from_bytes(bytes).unwrap();
    assert_eq!(presentation.slide_count().unwrap(), 0);
    assert_eq!(presentation.text().unwrap(), "");
}

#[test]
fn semantic_page_and_layout_facades_round_trip() {
    let measure =
        litchi_odp::layout::Measure::new(1.0, litchi_odp::layout::Unit::Centimeter).unwrap();
    let mut layouts = litchi_odp::layout::Collection::default();
    let mut layout = litchi_odp::layout::Layout::new("title_layout").unwrap();
    layout
        .placeholders
        .push(litchi_odp::layout::Placeholder::new(
            litchi_odp::layout::Role::Title,
            measure,
            measure,
            measure,
            measure,
        ));
    layouts.layouts.push(layout);
    layouts.validate().unwrap();

    let page = litchi_odp::page::Page {
        slide_index: 0,
        name: Some("page1".to_string()),
        style_name: None,
        master_page_name: None,
        page_layout_name: Some("title_layout".to_string()),
        draw_id: None,
        xml_id: None,
        href: None,
        navigation_order: Vec::new(),
    };
    let pages = litchi_odp::page::Collection::new(vec![page]).unwrap();

    let mut builder = Builder::new();
    builder.add_slide_with_title("Title", "Body").unwrap();
    builder.set_layouts(layouts).unwrap();
    builder.set_pages(Some(pages)).unwrap();
    let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();

    assert_eq!(
        presentation.layouts().unwrap().layouts[0].name,
        "title_layout"
    );
    assert_eq!(
        presentation.pages().unwrap().pages()[0].name.as_deref(),
        Some("page1")
    );
    let _: Option<litchi_odp::slide::Slide> = presentation.slide(0).unwrap();
}
