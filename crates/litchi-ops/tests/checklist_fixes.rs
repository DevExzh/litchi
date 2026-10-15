//! Regressions for the claimed-supported cells record 0289's CRUD checklist
//! left failing (F1, F2, F4, F6, F7) and the metadata `language` read-back
//! record 0288 left open; record 0290 fixed each. Every test fails on the
//! base build.

#[cfg(any(feature = "pptx", feature = "odp", feature = "rtf"))]
fn fixture(relative: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(relative),
    )
    .unwrap()
}

#[cfg(any(feature = "pptx", feature = "odp"))]
fn paths(file: &litchi_ops::File, selector: &str) -> Vec<String> {
    file.query(selector)
        .unwrap()
        .map(|node| node.unwrap().positional_path().to_string())
        .collect()
}

/// F1: a bare `placeholder` selects the shapes with a placeholder binding
/// (PPTX) or presentation class (ODP), title placeholders included, at their
/// `/shape[N]` addresses; `shape` still selects them too.
#[cfg(any(feature = "pptx", feature = "odp"))]
fn assert_placeholders_are_shapes(file: &litchi_ops::File) {
    let placeholders = paths(file, "placeholder");
    assert!(!placeholders.is_empty());
    let shapes = paths(file, "shape");
    assert!(placeholders.iter().all(|path| shapes.contains(path)));
    let mut typed = paths(file, "shape[type=placeholder]");
    typed.extend(paths(file, "shape[type=title]"));
    typed.sort();
    let mut sorted = placeholders.clone();
    sorted.sort();
    assert_eq!(sorted, typed);
    let on_slides: Vec<String> = placeholders
        .iter()
        .filter(|path| path.starts_with("/slide["))
        .cloned()
        .collect();
    assert_eq!(paths(file, "slide > placeholder"), on_slides);
}

#[cfg(feature = "pptx")]
#[test]
fn pptx_bare_placeholder_query_matches_placeholder_shapes() {
    let file = litchi_ops::open_bytes(fixture("ooxml/pptx/radar-chart.pptx")).unwrap();
    assert_placeholders_are_shapes(&file);
    let created = litchi_ops::create(litchi_core::FileFormat::Pptx).unwrap();
    if created.query("shape").unwrap().next().is_some() {
        assert_placeholders_are_shapes(&created);
    }
}

#[cfg(feature = "odp")]
#[test]
fn odp_bare_placeholder_query_matches_presentation_class_shapes() {
    let file = litchi_ops::open_bytes(fixture("odf/corpus/impress-basic.odp")).unwrap();
    assert_placeholders_are_shapes(&file);
}

/// F2: a chart whose producer wrote negative `c:axId` values has a node
/// view, and its axes read.
#[cfg(feature = "pptx")]
#[test]
fn pptx_chart_with_negative_axis_ids_has_a_node_view() {
    let file = litchi_ops::open_bytes(fixture("ooxml/pptx/radar-chart.pptx")).unwrap();
    let chart = file.get("/slide[1]/chart[1]").unwrap();
    chart.view().format().unwrap();
    assert!(file.query("axis").unwrap().count() >= 2);
}

/// F4: `hidden` on a slide without slide-show information, from `create()`.
#[cfg(feature = "ppt")]
#[test]
fn ppt_hidden_on_a_created_slide() {
    use litchi_ops::{NewNode, props};

    let file = litchi_ops::create(litchi_core::FileFormat::Ppt)
        .unwrap()
        .add("/", NewNode::new("slide"))
        .unwrap()
        .file;
    let hidden = |file: &litchi_ops::File| {
        file.get("/slide[1]")
            .unwrap()
            .view()
            .format()
            .unwrap()
            .into_iter()
            .find(|prop| prop.key() == "hidden")
            .map(|prop| prop.to_string())
    };
    let shown = file.set("/slide[1]", props! { "hidden" => false }).unwrap();
    assert_eq!(shown.to_bytes().unwrap(), file.to_bytes().unwrap());
    let set = file.set("/slide[1]", props! { "hidden" => true }).unwrap();
    let reopened = litchi_ops::open_bytes(set.to_bytes().unwrap()).unwrap();
    assert_eq!(hidden(&reopened).as_deref(), Some("true"));
    let back = reopened
        .set("/slide[1]", props! { "hidden" => false })
        .unwrap();
    assert_eq!(hidden(&back).as_deref(), Some("false"));
}

/// F7: LibreOffice `n192129.rtf` stores GIF data under `\pngblip`; removing
/// the picture moves the group's bytes unchanged.
#[cfg(feature = "rtf")]
#[test]
fn rtf_picture_with_mismatched_payload_is_removed() {
    let input = fixture("libreoffice-core/sw/qa/extras/rtfexport/data/n192129.rtf");
    let file = litchi_ops::open_bytes(input.clone()).unwrap();
    let removed = file.remove("/body/p[1]/picture[1]").unwrap();
    assert_eq!(removed.query("picture").unwrap().count(), 0);
    let output = String::from_utf8(removed.to_bytes().unwrap()).unwrap();
    let source = String::from_utf8(input).unwrap();
    let start = source.find("{\\pict").unwrap();
    let end = start + source[start..].find('}').unwrap() + 1;
    assert_eq!(output, format!("{}{}", &source[..start], &source[end..]));
}

/// F6 and the language read-back: every ODF format made by `create()`
/// takes a `/metadata` set (OTH and ODM packages have no `meta.xml` until
/// then) and `File::metadata()` reports the written language.
#[cfg(all(
    feature = "ods",
    feature = "odp",
    feature = "odt",
    feature = "oth",
    feature = "odm",
    feature = "odg",
    feature = "odc",
    feature = "odi",
    feature = "odb",
    feature = "odf-formula"
))]
#[test]
fn odf_metadata_language_round_trips_from_create() {
    use litchi_core::FileFormat;
    use litchi_ops::props;

    for format in [
        FileFormat::Odt,
        FileFormat::Ods,
        FileFormat::Odp,
        FileFormat::Oth,
        FileFormat::Odm,
        FileFormat::Odg,
        FileFormat::Odc,
        FileFormat::Odi,
        FileFormat::Odb,
        FileFormat::Odf,
    ] {
        let file = litchi_ops::create(format).unwrap();
        let set = file
            .set(
                "/metadata",
                props! { "title" => "Litchi", "language" => "de-DE" },
            )
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        let reopened = litchi_ops::open_bytes(set.to_bytes().unwrap()).unwrap();
        let metadata = reopened.metadata().unwrap().unwrap();
        assert_eq!(metadata.title.as_deref(), Some("Litchi"), "{format:?}");
        assert_eq!(metadata.language.as_deref(), Some("de-DE"), "{format:?}");
    }
}
