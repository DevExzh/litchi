//! Unified reads of OpenDocument presentations (record 0278): path
//! resolution, node views, queries and view modes over the scene index.

#![cfg(feature = "odp")]

use litchi_core::FileFormat;
use litchi_ops::view::Mode;
use litchi_ops::{Code, File};

fn fixture(relative: &str) -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

fn open(relative: &str) -> File {
    let file = litchi_ops::open(fixture(relative)).expect("open ODP fixture");
    assert_eq!(file.format(), Some(FileFormat::Odp));
    file
}

fn text(file: &File, path: &str) -> Option<String> {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .expect("text")
        .map(|text| text.into_owned())
}

#[test]
fn paths_resolve_slides_shapes_paragraphs_and_runs() {
    let file = open("odf/corpus/impress-basic.odp");
    let title = file.get("/slide[1]/shape[1]").unwrap();
    assert_eq!(title.kind(), "title");
    assert_eq!(title.path().to_string(), "/slide[1]/shape[1]");
    assert_eq!(text(&file, "/slide[1]/shape[1]").as_deref(), Some("OL"));
    assert_eq!(
        text(&file, "/slide[1]/shape[2]/paragraph[2]").as_deref(),
        Some("Tsvey")
    );
    assert_eq!(
        text(&file, "/slide[1]/shape[2]/p[2]/r[1]").as_deref(),
        Some("Tsvey")
    );
    // A placeholder resolves by its presentation class.
    let body = file.get("/slide[1]/placeholder[outline]").unwrap();
    assert_eq!(body.path().to_string(), "/slide[1]/shape[2]");
    assert_eq!(
        file.get("/slide[1]/placeholder[title]")
            .unwrap()
            .path()
            .to_string(),
        "/slide[1]/shape[1]"
    );
    let slide = file.get("/slide[1]").unwrap();
    let format: Vec<String> = slide
        .view()
        .format()
        .unwrap()
        .map(|prop| format!("{}={prop}", prop.key()))
        .collect();
    assert!(format.contains(&"layout=AL1T1".to_owned()), "{format:?}");
    assert!(format.contains(&"master=Default".to_owned()), "{format:?}");
    // Notes are the notes page's text frame.
    let notes = file.get("/slide[1]/notes").unwrap();
    assert_eq!(notes.kind(), "notes");
    // The root projects meta.xml.
    let root: Vec<String> = file
        .root()
        .unwrap()
        .view()
        .format()
        .unwrap()
        .map(|prop| format!("{}={prop}", prop.key()))
        .collect();
    assert!(root.contains(&"slideCount=1".to_owned()), "{root:?}");
    assert_eq!(
        file.metadata().unwrap().unwrap().title.as_deref(),
        Some("PowerPoint Presentation")
    );
}

#[test]
fn missing_gaps_and_absent_concepts_are_typed() {
    let file = open("odf/corpus/impress-basic.odp");
    assert_eq!(
        file.get("/slide[9]").unwrap_err().code(),
        Code::PathNotFound
    );
    assert!(file.find("/slide[9]").unwrap().is_none());
    // Record 0328: video frames are typed media, so a slide without one
    // has no `video[1]`; animations remain a gap.
    assert_eq!(
        file.get("/slide[1]/video[1]").unwrap_err().code(),
        Code::PathNotFound
    );
    assert_eq!(
        file.get("/slide[1]/animation[1]").unwrap_err().code(),
        Code::UnsupportedCapability
    );
    assert_eq!(
        file.get("/slide[1]/smartart[1]").unwrap_err().code(),
        Code::UnsupportedPath
    );
    assert_eq!(
        file.query("smartart").unwrap_err().code(),
        Code::UnsupportedPath
    );
}

#[test]
fn tables_pictures_masters_and_layouts_read() {
    let file = open("odf/odp/cellspan.odp");
    let table = file.get("/slide[1]/table[1]").unwrap();
    assert_eq!(table.kind(), "table");
    let cells = file.query("tc").unwrap().count();
    assert_eq!(cells, 12);
    assert!(
        file.get("/slide[1]/table[1]/tr[1]/tc[1]/paragraph[1]")
            .is_ok()
    );

    let file = open("odf/odp/text-in-image.odp");
    let picture = file.get("/slide[1]/picture[1]").unwrap();
    assert_eq!(picture.kind(), "picture");
    let master = file.get("/slideMaster[1]").unwrap();
    let props: Vec<String> = master
        .view()
        .format()
        .unwrap()
        .map(|prop| format!("{}={prop}", prop.key()))
        .collect();
    assert!(props.contains(&"name=Default".to_owned()), "{props:?}");
    assert!(file.get("/slideLayout[1]").is_ok());
}

#[test]
fn queries_are_deterministic_and_stop_early() {
    let file = open("odf/odp/tdf169979.odp");
    let digest = |selector: &str| {
        file.query(selector)
            .unwrap()
            .map(|node| {
                let node = node.unwrap();
                format!(
                    "{} {} {:?}",
                    node.kind(),
                    node.path(),
                    node.view().text().unwrap()
                )
            })
            .collect::<Vec<_>>()
    };
    for selector in ["shape", "title", "paragraph", "notes paragraph", "ole"] {
        assert_eq!(digest(selector), digest(selector), "{selector}");
    }
    assert_eq!(digest("title").len(), 3);
    assert_eq!(digest("ole").len(), 3);
    let first = file.query("paragraph").unwrap().first().unwrap().unwrap();
    assert_eq!(first.path().to_string(), "/slide[1]/shape[1]/paragraph[1]");
    assert_eq!(file.query("paragraph").unwrap().limit(2).count(), 2);
}

#[test]
fn view_modes_render() {
    let file = open("odf/corpus/impress-basic.odp");
    let text = file.view(Mode::Text).unwrap();
    assert!(text.starts_with("=== /slide[1] ===\nOL\nEins"), "{text}");
    let outline = file.view(Mode::Outline).unwrap();
    assert!(outline.contains("Slide 1: \"OL\""), "{outline}");
    assert!(file.view(Mode::Stats).unwrap().contains("Slides: 1"));
    assert!(
        file.view(Mode::Annotated)
            .unwrap()
            .contains("[Title] \"OL\"")
    );
    assert_eq!(
        file.view(Mode::Html).unwrap_err().code(),
        Code::UnsupportedCapability
    );
    let json = serde_json::to_value(file.view_json(Mode::Outline).unwrap()).unwrap();
    assert_eq!(json["totalSlides"], 1);
}

#[test]
fn flat_presentations_open_for_reads() {
    let file = open("libreoffice-core/sd/qa/unit/data/draw-image-link.fodp");
    assert!(file.get("/slide[1]").is_ok());
    assert!(file.view(Mode::Outline).unwrap().starts_with("1 slides"));
    // Record 0341: XML's predefined entities no longer make the leaf
    // refuse a flat presentation.
    let theme = open("libreoffice-core/xmloff/qa/unit/data/theme.fodp");
    assert!(theme.get("/slide[1]").is_ok());
    // A flat document the leaf refuses (an empty office:body) stays a
    // typed open error.
    let error = litchi_ops::open(fixture("odf/drawing/multicolor-gradient.fodp")).unwrap_err();
    assert_eq!(error.code(), Code::CorruptFile);
}
