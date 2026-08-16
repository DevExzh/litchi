use std::io::{self, Write};

use litchi_opc::OpcError;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::phys_pkg::PhysPkgReader;
use litchi_pptx::{
    Error, Package, StreamingPresentationLimits, StreamingPresentationOptions,
    StreamingPresentationWriter, TextBoxSpec,
};

fn deck(slide_count: usize) -> Vec<u8> {
    let mut writer = StreamingPresentationWriter::new(Vec::new(), slide_count).unwrap();
    for index in 0..slide_count {
        let title = (index == 0).then_some("Title & one");
        let mut slide = writer.start_slide(title).unwrap();
        slide
            .write_text_box(TextBoxSpec::new(
                if index == 0 {
                    "Hello <world>"
                } else {
                    "Second"
                },
                914_400,
                914_400,
                2_743_200,
                914_400,
            ))
            .unwrap();
        writer = slide.finish().unwrap();
    }
    writer.finish().unwrap()
}

fn empty_deck() -> Vec<u8> {
    let writer = StreamingPresentationWriter::new(Vec::new(), 1).unwrap();
    let slide = writer.start_slide(None).unwrap();
    let writer = slide.finish().unwrap();
    writer.finish().unwrap()
}

#[test]
fn streaming_deck_is_deterministic_and_reopens() {
    let first = deck(2);
    let second = deck(2);
    assert_eq!(first, second);

    let package = Package::from_bytes(&first).unwrap();
    let presentation = package.presentation().unwrap();
    assert_eq!(presentation.slide_count().unwrap(), 2);
    assert_eq!(presentation.slide_size().unwrap(), (9_144_000, 6_858_000));
    let first_slide = presentation.slide(0).unwrap().unwrap();
    assert_eq!(first_slide.shape_count().unwrap(), 2);
    assert!(first_slide.text().unwrap().contains("Hello <world>"));
    let scene = first_slide.shapes().unwrap();
    let shape = scene.shape(0).unwrap();
    assert_eq!(shape.bounds().unwrap().x(), 914_400);
    assert_eq!(shape.bounds().unwrap().width(), 7_315_200);

    let reader = PhysPkgReader::new(&first).unwrap();
    assert_eq!(reader.len(), 41);
    let names = reader.member_names().unwrap();
    assert!(names.iter().any(|name| name == "ppt/slides/slide1.xml"));
    assert!(
        names
            .iter()
            .any(|name| name == "ppt/slides/_rels/slide2.xml.rels")
    );
    for name in names
        .iter()
        .filter(|name| name.ends_with(".xml") || name.ends_with(".rels"))
    {
        let xml = reader.read_member(name).unwrap();
        assert!(!xml.windows(2).any(|window| window == b">\n"));
        assert!(!xml.windows(2).any(|window| window == b">\r"));
    }
}

#[test]
fn title_multiple_boxes_have_deterministic_order_and_ids() {
    let mut writer = StreamingPresentationWriter::new(Vec::new(), 1).unwrap();
    let mut slide = writer.start_slide(Some("Title")).unwrap();
    slide
        .write_text_box(TextBoxSpec::new("First", 0, 0, 1, 1))
        .unwrap();
    slide
        .write_text_box(TextBoxSpec::new("Second", 1, 1, 1, 1))
        .unwrap();
    assert_eq!(slide.text_box_count(), 2);
    writer = slide.finish().unwrap();
    let bytes = writer.finish().unwrap();

    let reader = PhysPkgReader::new(&bytes).unwrap();
    let xml = String::from_utf8(reader.read_member("ppt/slides/slide1.xml").unwrap()).unwrap();
    let title = xml.find("id=\"2\"").unwrap();
    let first = xml.find("id=\"3\"").unwrap();
    let second = xml.find("id=\"4\"").unwrap();
    assert!(title < first && first < second);
    assert!(xml.find(">Title</a:t>").unwrap() < xml.find(">First</a:t>").unwrap());
    assert!(xml.find(">First</a:t>").unwrap() < xml.find(">Second</a:t>").unwrap());

    let package = Package::from_bytes(&bytes).unwrap();
    let semantic_slide = package.presentation().unwrap().slide(0).unwrap().unwrap();
    assert_eq!(semantic_slide.shape_count().unwrap(), 3);
    let shapes = semantic_slide.shapes().unwrap();
    assert_eq!(shapes.shape(0).unwrap().text(), Some("Title"));
    assert_eq!(shapes.shape(1).unwrap().text(), Some("First"));
    assert_eq!(shapes.shape(2).unwrap().text(), Some("Second"));
}

#[test]
fn title_only_slide_is_semantic_text_without_caller_boxes() {
    let writer = StreamingPresentationWriter::new(Vec::new(), 1).unwrap();
    let slide = writer.start_slide(Some("Title only")).unwrap();
    assert_eq!(slide.text_box_count(), 0);
    let writer = slide.finish().unwrap();
    let bytes = writer.finish().unwrap();

    let package = Package::from_bytes(&bytes).unwrap();
    let semantic_slide = package.presentation().unwrap().slide(0).unwrap().unwrap();
    assert_eq!(semantic_slide.shape_count().unwrap(), 1);
    assert_eq!(semantic_slide.text().unwrap(), "Title only");
}

#[test]
fn empty_widescreen_slide_reopens_with_fixed_topology() {
    let mut writer = StreamingPresentationWriter::with_options(
        Vec::new(),
        1,
        StreamingPresentationOptions::widescreen(),
        StreamingPresentationLimits::default(),
    )
    .unwrap();
    let slide = writer.start_slide(None).unwrap();
    assert_eq!(slide.text_box_count(), 0);
    writer = slide.finish().unwrap();
    let bytes = writer.finish().unwrap();

    let package = Package::from_bytes(&bytes).unwrap();
    let presentation = package.presentation().unwrap();
    assert_eq!(presentation.slide_count().unwrap(), 1);
    assert_eq!(presentation.slide_size().unwrap(), (9_144_000, 5_143_500));
    assert_eq!(
        presentation
            .slide(0)
            .unwrap()
            .unwrap()
            .shape_count()
            .unwrap(),
        0
    );
    assert_eq!(PhysPkgReader::new(&bytes).unwrap().len(), 39);
}

#[test]
fn static_content_types_and_relationships_close_the_one_slide_graph() {
    let bytes = empty_deck();
    let reader = PhysPkgReader::new(&bytes).unwrap();
    let names = reader.member_names().unwrap();
    assert_eq!(names.len(), 39);
    for name in [
        "[Content_Types].xml",
        "_rels/.rels",
        "docProps/core.xml",
        "docProps/app.xml",
        "ppt/presProps.xml",
        "ppt/viewProps.xml",
        "ppt/tableStyles.xml",
        "ppt/slideMasters/slideMaster1.xml",
        "ppt/slideMasters/_rels/slideMaster1.xml.rels",
        "ppt/theme/theme1.xml",
        "ppt/theme/theme2.xml",
        "ppt/notesMasters/notesMaster1.xml",
        "ppt/notesMasters/_rels/notesMaster1.xml.rels",
        "ppt/presentation.xml",
        "ppt/_rels/presentation.xml.rels",
        "ppt/slides/slide1.xml",
        "ppt/slides/_rels/slide1.xml.rels",
    ] {
        assert!(
            names.iter().any(|candidate| candidate == name),
            "missing {name}"
        );
    }
    for index in 1..=11 {
        assert!(
            names
                .iter()
                .any(|name| name == &format!("ppt/slideLayouts/slideLayout{index}.xml"))
        );
        assert!(names.iter().any(|name| {
            name == &format!("ppt/slideLayouts/_rels/slideLayout{index}.xml.rels")
        }));
    }

    let content_types =
        String::from_utf8(reader.read_member("[Content_Types].xml").unwrap()).unwrap();
    assert_eq!(content_types.matches("<Override ").count(), 22);
    for (path, content_type) in [
        ("/docProps/app.xml", ct::OFC_EXTENDED_PROPERTIES),
        ("/docProps/core.xml", ct::OPC_CORE_PROPERTIES),
        ("/ppt/presProps.xml", ct::PML_PRES_PROPS),
        ("/ppt/presentation.xml", ct::PML_PRESENTATION_MAIN),
        ("/ppt/slideMasters/slideMaster1.xml", ct::PML_SLIDE_MASTER),
        ("/ppt/slides/slide1.xml", ct::PML_SLIDE),
        ("/ppt/tableStyles.xml", ct::PML_TABLE_STYLES),
        ("/ppt/theme/theme1.xml", ct::OFC_THEME),
        ("/ppt/theme/theme2.xml", ct::OFC_THEME),
        ("/ppt/viewProps.xml", ct::PML_VIEW_PROPS),
        ("/ppt/notesMasters/notesMaster1.xml", ct::PML_NOTES_MASTER),
    ] {
        let needle = format!("PartName=\"{path}\" ContentType=\"{content_type}\"");
        assert!(content_types.contains(&needle), "missing {needle}");
    }
    for index in 1..=11 {
        let needle = format!(
            "PartName=\"/ppt/slideLayouts/slideLayout{index}.xml\" ContentType=\"{}\"",
            ct::PML_SLIDE_LAYOUT
        );
        assert!(content_types.contains(&needle), "missing {needle}");
    }

    let root_rels = String::from_utf8(reader.read_member("_rels/.rels").unwrap()).unwrap();
    for (id, relation_type, target) in [
        ("rId1", rt::OFFICE_DOCUMENT, "ppt/presentation.xml"),
        ("rId2", rt::CORE_PROPERTIES, "docProps/core.xml"),
        ("rId3", rt::EXTENDED_PROPERTIES, "docProps/app.xml"),
    ] {
        let needle = format!("Id=\"{id}\" Type=\"{relation_type}\" Target=\"{target}\"");
        assert!(root_rels.contains(&needle), "missing {needle}");
    }
    let presentation_rels = String::from_utf8(
        reader
            .read_member("ppt/_rels/presentation.xml.rels")
            .unwrap(),
    )
    .unwrap();
    for (id, relation_type, target) in [
        ("rId1", rt::SLIDE_MASTER, "slideMasters/slideMaster1.xml"),
        ("rId2", rt::VIEW_PROPS, "viewProps.xml"),
        ("rId3", rt::PRES_PROPS, "presProps.xml"),
        ("rId4", rt::SLIDE, "slides/slide1.xml"),
        (
            "rIdNotesMaster",
            rt::NOTES_MASTER,
            "notesMasters/notesMaster1.xml",
        ),
        ("rIdTableStyles", rt::TABLE_STYLES, "tableStyles.xml"),
    ] {
        let needle = format!("Id=\"{id}\" Type=\"{relation_type}\" Target=\"{target}\"");
        assert!(presentation_rels.contains(&needle), "missing {needle}");
    }
    let master_rels = String::from_utf8(
        reader
            .read_member("ppt/slideMasters/_rels/slideMaster1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    for index in 1..=11 {
        let needle = format!(
            "Id=\"rId{index}\" Type=\"{}\" Target=\"../slideLayouts/slideLayout{index}.xml\"",
            rt::SLIDE_LAYOUT
        );
        assert!(master_rels.contains(&needle), "missing {needle}");
    }
    assert!(master_rels.contains(&format!(
        "Id=\"rId12\" Type=\"{}\" Target=\"../theme/theme1.xml\"",
        rt::THEME
    )));
    let layout_rels = String::from_utf8(
        reader
            .read_member("ppt/slideLayouts/_rels/slideLayout1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    assert!(layout_rels.contains(&format!(
        "Id=\"rId1\" Type=\"{}\" Target=\"../slideMasters/slideMaster1.xml\"",
        rt::SLIDE_MASTER
    )));
    let notes_rels = String::from_utf8(
        reader
            .read_member("ppt/notesMasters/_rels/notesMaster1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    assert!(notes_rels.contains(&format!(
        "Id=\"rId1\" Type=\"{}\" Target=\"../theme/theme2.xml\"",
        rt::THEME
    )));
    let slide_rels = String::from_utf8(
        reader
            .read_member("ppt/slides/_rels/slide1.xml.rels")
            .unwrap(),
    )
    .unwrap();
    assert!(slide_rels.contains(&format!(
        "Id=\"rId1\" Type=\"{}\" Target=\"../slideLayouts/slideLayout1.xml\"",
        rt::SLIDE_LAYOUT
    )));
}

#[test]
fn invalid_global_configuration_does_not_touch_sink() {
    let mut sink = Vec::new();
    let error = match StreamingPresentationWriter::new(&mut sink, 0) {
        Ok(_) => panic!("zero slides must be refused"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Invalid(_)));
    assert!(sink.is_empty());

    let error = match StreamingPresentationWriter::with_options(
        &mut sink,
        1,
        StreamingPresentationOptions::new(1, 1).unwrap_or_default(),
        StreamingPresentationLimits {
            max_output_bytes: 0,
            ..StreamingPresentationLimits::default()
        },
    ) {
        Ok(_) => panic!("zero output limit must be refused"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Invalid(_)));
    assert!(sink.is_empty());

    let overflow_limits = StreamingPresentationLimits {
        max_slides: usize::MAX,
        ..StreamingPresentationLimits::default()
    };
    let error = match StreamingPresentationWriter::with_options(
        &mut sink,
        usize::MAX,
        StreamingPresentationOptions::default(),
        overflow_limits,
    ) {
        Ok(_) => panic!("unbounded slide count must be refused before slide-ID arithmetic"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Limit { .. }));
    assert!(sink.is_empty());

    let mut sink = Vec::new();
    let error = match StreamingPresentationWriter::with_options(
        &mut sink,
        32_749,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits::default(),
    ) {
        Ok(_) => panic!("the physical ZIP entry limit must be refused"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Limit { .. }));
    assert!(sink.is_empty());

    let mut sink = Vec::new();
    let error = match StreamingPresentationWriter::with_options(
        &mut sink,
        1,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits {
            max_output_bytes: 1,
            ..StreamingPresentationLimits::default()
        },
    ) {
        Ok(_) => panic!("an output budget below structural ZIP metadata must be refused"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Limit { .. }));
    assert!(sink.is_empty());

    let mut sink = Vec::new();
    let error = match StreamingPresentationWriter::with_options(
        &mut sink,
        1,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits {
            max_output_bytes: 512 * 1024 * 1024 + 1,
            ..StreamingPresentationLimits::default()
        },
    ) {
        Ok(_) => panic!("the public output budget must not exceed the physical ZIP budget"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Limit { .. }));
    assert!(sink.is_empty());
}

#[test]
fn order_and_limit_refusals_are_checked() {
    let mut writer = StreamingPresentationWriter::with_options(
        Vec::new(),
        1,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits {
            max_text_boxes_per_slide: 1,
            max_text_bytes_per_box: 5,
            max_total_text_bytes: 5,
            ..StreamingPresentationLimits::default()
        },
    )
    .unwrap();
    let mut slide = writer.start_slide(None).unwrap();
    assert!(matches!(
        slide.write_text_box(TextBoxSpec::new("abcdef", 0, 0, 1, 1)),
        Err(Error::Limit { .. })
    ));
    slide
        .write_text_box(TextBoxSpec::new("hello", 0, 0, 1, 1))
        .unwrap();
    assert!(matches!(
        slide.write_text_box(TextBoxSpec::new("x", 0, 0, 1, 1)),
        Err(Error::Limit { .. })
    ));
    writer = slide.finish().unwrap();
    writer.finish().unwrap();

    let mut missing = StreamingPresentationWriter::new(Vec::new(), 2).unwrap();
    let slide = missing.start_slide(None).unwrap();
    missing = slide.finish().unwrap();
    assert!(matches!(missing.finish(), Err(Error::Invalid(_))));

    let mut extra = StreamingPresentationWriter::new(Vec::new(), 1).unwrap();
    let slide = extra.start_slide(None).unwrap();
    extra = slide.finish().unwrap();
    assert!(matches!(extra.start_slide(None), Err(Error::Invalid(_))));
}

#[test]
fn slide_xml_limit_reserves_the_final_suffix_before_box_bytes() {
    let mut reference = StreamingPresentationWriter::new(Vec::new(), 1).unwrap();
    let mut slide = reference.start_slide(None).unwrap();
    slide
        .write_text_box(TextBoxSpec::new("exact", 0, 0, 1, 1))
        .unwrap();
    reference = slide.finish().unwrap();
    let full_bytes = reference.finish().unwrap();
    let full_xml_len = PhysPkgReader::new(&full_bytes)
        .unwrap()
        .read_member("ppt/slides/slide1.xml")
        .unwrap()
        .len();

    let mut limits = StreamingPresentationLimits {
        max_slide_xml_bytes: full_xml_len - 1,
        ..StreamingPresentationLimits::default()
    };
    let writer = StreamingPresentationWriter::with_options(
        Vec::new(),
        1,
        StreamingPresentationOptions::default(),
        limits,
    )
    .unwrap();
    let mut slide = writer.start_slide(None).unwrap();
    let before = slide.output_bytes();
    assert!(matches!(
        slide.write_text_box(TextBoxSpec::new("exact", 0, 0, 1, 1)),
        Err(Error::Limit { .. })
    ));
    assert_eq!(slide.output_bytes(), before);

    limits.max_slide_xml_bytes = full_xml_len;
    let mut writer = StreamingPresentationWriter::with_options(
        Vec::new(),
        1,
        StreamingPresentationOptions::default(),
        limits,
    )
    .unwrap();
    let mut slide = writer.start_slide(None).unwrap();
    slide
        .write_text_box(TextBoxSpec::new("exact", 0, 0, 1, 1))
        .unwrap();
    writer = slide.finish().unwrap();
    writer.finish().unwrap();
}

#[test]
fn hostile_text_and_geometry_are_refused_before_box_bytes() {
    let mut sink = Vec::new();
    let writer = StreamingPresentationWriter::new(&mut sink, 1).unwrap();
    let before_title = writer.output_bytes();
    assert!(matches!(
        writer.start_slide(Some("bad\u{1}")),
        Err(Error::Invalid(_))
    ));
    assert_eq!(sink.len() as u64, before_title);

    let writer = StreamingPresentationWriter::new(Vec::new(), 1).unwrap();
    let mut slide = writer.start_slide(None).unwrap();
    let before = slide.output_bytes();
    assert!(matches!(
        slide.write_text_box(TextBoxSpec::new("bad\u{1}", 0, 0, 1, 1)),
        Err(Error::Invalid(_))
    ));
    assert_eq!(slide.output_bytes(), before);
    assert!(matches!(
        slide.write_text_box(TextBoxSpec::new("bad", -1, 0, 1, 1)),
        Err(Error::Invalid(_))
    ));
}

#[test]
fn unicode_text_budget_is_utf8_exact_and_whitespace_is_preserved() {
    let text = "  é & <  ";
    let mut limits = StreamingPresentationLimits {
        max_text_bytes_per_box: text.len(),
        max_total_text_bytes: text.len(),
        ..StreamingPresentationLimits::default()
    };
    let mut writer = StreamingPresentationWriter::with_options(
        Vec::new(),
        1,
        StreamingPresentationOptions::default(),
        limits,
    )
    .unwrap();
    let mut slide = writer.start_slide(None).unwrap();
    slide
        .write_text_box(TextBoxSpec::new(text, 0, 0, 1, 1))
        .unwrap();
    writer = slide.finish().unwrap();
    let bytes = writer.finish().unwrap();
    let reader = PhysPkgReader::new(&bytes).unwrap();
    let xml = String::from_utf8(reader.read_member("ppt/slides/slide1.xml").unwrap()).unwrap();
    assert!(xml.contains("xml:space=\"preserve\">  é &amp; &lt;  </a:t>"));

    let package = Package::from_bytes(&bytes).unwrap();
    let semantic_slide = package.presentation().unwrap().slide(0).unwrap().unwrap();
    let shapes = semantic_slide.shapes().unwrap();
    assert_eq!(shapes.shape(0).unwrap().text(), Some(text));

    limits.max_text_bytes_per_box = 1;
    let writer = StreamingPresentationWriter::with_options(
        Vec::new(),
        1,
        StreamingPresentationOptions::default(),
        limits,
    )
    .unwrap();
    let mut slide = writer.start_slide(None).unwrap();
    let before = slide.output_bytes();
    assert!(matches!(
        slide.write_text_box(TextBoxSpec::new("é", 0, 0, 1, 1)),
        Err(Error::Limit { .. })
    ));
    assert_eq!(slide.output_bytes(), before);

    limits.max_text_bytes_per_box = text.len();
    limits.max_total_text_bytes = 1;
    let writer = StreamingPresentationWriter::with_options(
        Vec::new(),
        1,
        StreamingPresentationOptions::default(),
        limits,
    )
    .unwrap();
    let mut slide = writer.start_slide(None).unwrap();
    assert!(matches!(
        slide.write_text_box(TextBoxSpec::new("é", 0, 0, 1, 1)),
        Err(Error::Limit { .. })
    ));
}

struct FailingSink {
    bytes: Vec<u8>,
    remaining: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(io::ErrorKind::WriteZero, "test sink"));
        }
        let count = bytes.len().min(self.remaining);
        self.bytes.extend_from_slice(&bytes[..count]);
        self.remaining -= count;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn sink_failure_reports_incomplete_output_progress() {
    let result = StreamingPresentationWriter::new(
        FailingSink {
            bytes: Vec::new(),
            remaining: 128,
        },
        1,
    );
    let error = match result {
        Ok(_) => panic!("failing sink must fail"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Opc(OpcError::IncompleteOutput { written, .. }) if written > 0
    ));
}

#[test]
fn post_start_sink_failure_reports_progress_and_poisoning() {
    let alphabet = b"0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
    let mut state = 0x1234_5678_u32;
    let mut text = String::with_capacity(128 * 1024);
    for _ in 0..128 * 1024 {
        state = state.wrapping_mul(1_664_525).wrapping_add(1_013_904_223);
        text.push(char::from(alphabet[(state as usize) % alphabet.len()]));
    }
    let mut reference_slide = StreamingPresentationWriter::new(Vec::new(), 1)
        .unwrap()
        .start_slide(None)
        .unwrap();
    let before_box = reference_slide.output_bytes();
    reference_slide
        .write_text_box(TextBoxSpec::new(&text, 0, 0, 1, 1))
        .unwrap();
    let after_box = reference_slide.output_bytes();
    assert!(after_box > before_box + 8);
    drop(reference_slide);

    let remaining = usize::try_from(before_box + (after_box - before_box) / 2).unwrap();
    let writer = StreamingPresentationWriter::new(
        FailingSink {
            bytes: Vec::new(),
            remaining,
        },
        1,
    )
    .unwrap();
    let mut slide = writer.start_slide(None).unwrap();
    let result = slide.write_text_box(TextBoxSpec::new(&text, 0, 0, 1, 1));
    let error = match result {
        Ok(()) => panic!("the sink must fail while the box is being streamed"),
        Err(error) => error,
    };
    let written = match error {
        Error::Opc(OpcError::IncompleteOutput { written, .. }) => written,
        other => panic!("expected incomplete output, got {other:?}"),
    };
    assert!(written > before_box);
    assert_eq!(slide.output_bytes(), written);
    assert!(matches!(
        slide.write_text_box(TextBoxSpec::new("after failure", 0, 0, 1, 1)),
        Err(Error::Invalid(message)) if message.contains("poisoned")
    ));
    assert_eq!(slide.output_bytes(), written);
}

#[test]
fn runtime_output_limit_is_exact_and_one_under_is_incomplete() {
    // The structural preflight is only a lower bound. This empirically exact
    // budget includes streamed payloads and ZIP finalization; one byte less
    // must therefore fail after output has begun with typed progress.
    let reference = empty_deck();
    let exact_limits = StreamingPresentationLimits {
        max_output_bytes: reference.len() as u64,
        ..StreamingPresentationLimits::default()
    };
    let mut exact = StreamingPresentationWriter::with_options(
        Vec::new(),
        1,
        StreamingPresentationOptions::default(),
        exact_limits,
    )
    .unwrap();
    let slide = exact.start_slide(None).unwrap();
    exact = slide.finish().unwrap();
    let exact_bytes = exact.finish().unwrap();
    assert_eq!(exact_bytes.len(), reference.len());

    let mut under_limits = exact_limits;
    under_limits.max_output_bytes -= 1;
    let result = StreamingPresentationWriter::with_options(
        Vec::new(),
        1,
        StreamingPresentationOptions::default(),
        under_limits,
    )
    .and_then(|writer| writer.start_slide(None))
    .and_then(|slide| slide.finish())
    .and_then(StreamingPresentationWriter::finish);
    let error = match result {
        Ok(_) => panic!("one byte below the output ceiling must fail"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Opc(OpcError::IncompleteOutput { written, .. }) if written > 0
    ));
}
