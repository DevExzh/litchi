use litchi_rtf::{GeneratedListMarkerKind, RtfDocument, RtfWriter};

const SYNTHETIC: &str = concat!(
    r#"{\rtf1\ansi\ansicpg1250\uc1 A"#,
    r#"{\listtext\pard\plain\u8226?\'8a\tab}B"#,
    r#"{\pntext\pard\plain 1.\tab}C}"#,
);

#[test]
fn parses_inert_markers_and_round_trips_without_visible_duplication() {
    let document = RtfDocument::parse(SYNTHETIC).unwrap();
    assert_eq!(document.text(), "ABC");
    assert_eq!(document.generated_list_markers().len(), 2);
    let modern = &document.generated_list_markers()[0];
    assert_eq!(modern.kind, GeneratedListMarkerKind::Modern);
    assert_eq!(modern.text, "•Š\t");
    assert_eq!(modern.position, 1);
    let legacy = &document.generated_list_markers()[1];
    assert_eq!(legacy.kind, GeneratedListMarkerKind::Legacy);
    assert_eq!(legacy.text, "1.\t");
    assert_eq!(legacy.position, 2);

    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(
        reparsed.generated_list_markers(),
        document.generated_list_markers()
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn rejects_malformed_placement_duplicates_active_content_and_bounds() {
    let malformed = [
        r#"{\rtf1\listtext x}"#,
        r#"{\rtf1{\*\listtext x}}"#,
        r#"{\rtf1{\header{\listtext x}}Body}"#,
        r#"{\rtf1{\listtext}}"#,
        r#"{\rtf1{\listtext x}{\listtext y}Body}"#,
        r#"{\rtf1{\listtext\bin2 xx}Body}"#,
        r#"{\rtf1{\listtext{\*\unknown danger}}Body}"#,
        r#"{\rtf1{\listtext{\field danger}}Body}"#,
        r#"{\rtf1{\listtext{\object danger}}Body}"#,
        r#"{\rtf1{\listtext{\pict 00}}Body}"#,
        r#"{\rtf1{\listtext{\shp danger}}Body}"#,
        r#"{\rtf1{\listtext{\formfield danger}}Body}"#,
        r#"{\rtf1{\listtext x\par}Body}"#,
        r#"{\rtf1{\pntext x\line}Body}"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed RTF: {source}"
        );
    }

    let oversized = format!(r#"{{\rtf1{{\listtext {}}}}}"#, "x".repeat(4097));
    assert!(RtfDocument::parse(&oversized).is_err());
}

fn isolated_marker(fixture: &[u8], marker: &[u8], header: &[u8]) -> Vec<u8> {
    let start = fixture
        .windows(marker.len())
        .position(|window| window == marker)
        .unwrap();
    let mut depth = 0usize;
    let mut end = None;
    for (offset, byte) in fixture[start..].iter().enumerate() {
        match byte {
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    end = Some(start + offset + 1);
                    break;
                }
            },
            _ => {},
        }
    }
    let mut source = br"{\rtf1".to_vec();
    source.extend_from_slice(header);
    source.extend_from_slice(b"A");
    source.extend_from_slice(&fixture[start..end.unwrap()]);
    source.extend_from_slice(b"B}");
    source
}

#[test]
fn parses_modern_and_legacy_bundled_libreoffice_markers() {
    let modern_fixture = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf167569.rtf"
    );
    let source = isolated_marker(modern_fixture, br"{\listtext", br"\ansi\ansicpg1252\uc1");
    let document = RtfDocument::parse_bytes(&source).unwrap();
    assert_eq!(document.text(), "AB");
    let marker = &document.generated_list_markers()[0];
    assert_eq!(marker.kind, GeneratedListMarkerKind::Modern);
    assert_eq!(marker.text, "·\t");
    assert_eq!(marker.position, 1);

    let legacy_fixture = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/core/data/rtf/fail/forcepoint-4.rtf"
    );
    let source = isolated_marker(legacy_fixture, br"{\pntext", br"\ansi\ansicpg1252\uc1");
    let document = RtfDocument::parse_bytes(&source).unwrap();
    assert_eq!(document.text(), "AB");
    let marker = &document.generated_list_markers()[0];
    assert_eq!(marker.kind, GeneratedListMarkerKind::Legacy);
    assert!(marker.text.contains("1."));
    assert!(marker.text.ends_with('\t'));
    assert_eq!(marker.position, 1);
}
