use litchi_rtf::{
    MAX_PARAGRAPH_TAB_STOPS, RtfDocument, RtfWriter, TabAlignment, TabLeader,
};

fn block<'a>(document: &'a RtfDocument<'a>, needle: &str) -> &'a litchi_rtf::StyleBlock<'a> {
    document
        .blocks()
        .iter()
        .find(|block| block.text.contains(needle))
        .unwrap_or_else(|| panic!("missing body block containing {needle:?}"))
}

#[test]
fn parses_libreoffice_paragraph_and_style_tab_fixture() {
    let source = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
            "../../3rdparty/libreoffice-core/sw/qa/extras/rtfimport/data/tdf96308-tabpos.rtf",
        ),
    )
    .unwrap();
    let document = RtfDocument::parse_bytes(&source).unwrap();

    let style_tabs = document
        .stylesheet()
        .get(30)
        .unwrap()
        .paragraph
        .unwrap()
        .tab_stops;
    assert_eq!(style_tabs.len(), 1);
    assert_eq!(style_tabs.as_slice()[0].position, 2552);

    assert!(document.tables().iter().any(|table| {
        table
            .rows()
            .iter()
            .flat_map(|row| row.cells())
            .any(|cell| cell.text().contains("A1"))
    }));

    let leader_source = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
            "../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/tab-stop-fill-chars.rtf",
        ),
    )
    .unwrap();
    let leader_document = RtfDocument::parse_bytes(&leader_source).unwrap();
    assert!(leader_document.blocks().iter().any(|block| {
        block
            .paragraph
            .tab_stops
            .iter()
            .any(|tab| tab.position == 2520 && tab.leader == TabLeader::MiddleDot)
    }));
    assert!(leader_document.blocks().iter().any(|block| {
        block
            .paragraph
            .tab_stops
            .iter()
            .any(|tab| tab.position == 2520 && tab.leader == TabLeader::Equal)
    }));
}

#[test]
fn parses_inherits_resets_and_deterministically_writes_all_tab_forms() {
    let source = concat!(
        r#"{\rtf1\ansi\pard"#,
        r#"\tqr\tldot\tx720\tqc\tlmdot\tx1440\tqdec\tlhyph\tx2160"#,
        r#"\tlul\tb2880\tlth\tx3600\tleq\tx4320\tql\tx5040 Outer"#,
        r#"{\tqr\tx5760 Inner}Tail{\pard Reset}"#,
        r#"{\*\unknown\tqr1\tx\tlmdot0 ignored}Visible}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let outer = block(&document, "Outer");
    let tabs = outer.paragraph.tab_stops.as_slice();
    assert_eq!(tabs.len(), 7);
    assert_eq!(tabs[0].alignment, TabAlignment::Right);
    assert_eq!(tabs[0].leader, TabLeader::Dot);
    assert_eq!(tabs[1].alignment, TabAlignment::Center);
    assert_eq!(tabs[1].leader, TabLeader::MiddleDot);
    assert_eq!(tabs[2].alignment, TabAlignment::Decimal);
    assert_eq!(tabs[2].leader, TabLeader::Hyphen);
    assert_eq!(tabs[3].alignment, TabAlignment::Bar);
    assert_eq!(tabs[3].leader, TabLeader::Underscore);
    assert_eq!(tabs[4].leader, TabLeader::ThickLine);
    assert_eq!(tabs[5].leader, TabLeader::Equal);
    assert_eq!(tabs[6].alignment, TabAlignment::Left);

    assert_eq!(block(&document, "Inner").paragraph.tab_stops.len(), 8);
    assert_eq!(block(&document, "Tail").paragraph.tab_stops, outer.paragraph.tab_stops);
    assert!(block(&document, "Reset").paragraph.tab_stops.is_empty());
    assert_eq!(block(&document, "Visible").paragraph.tab_stops, outer.paragraph.tab_stops);

    let mut output = Vec::new();
    RtfWriter::new(&mut output).write_document(&document).unwrap();
    let written = String::from_utf8(output).unwrap();
    assert!(written.contains(concat!(
        r#"\tqr\tldot\tx720\tqc\tlmdot\tx1440\tqdec\tlhyph\tx2160"#,
        r#"\tlul\tb2880\tlth\tx3600\tleq\tx4320\tx5040"#,
    )));
    assert!(!written.contains(r#"\tql"#));
    assert!(!written.contains(r#"\tb\tx"#));

    let reparsed = RtfDocument::parse(&written).unwrap();
    assert_eq!(
        block(&reparsed, "Outer").paragraph.tab_stops,
        outer.paragraph.tab_stops
    );
}

#[test]
fn parses_and_writes_tabs_in_paragraph_styles() {
    let document = RtfDocument::parse(
        r#"{\rtf1{\stylesheet{\s30\tqr\tldot\tx2552 Body Text 3;}}Body}"#,
    )
    .unwrap();
    let tabs = document.stylesheet().get(30).unwrap().paragraph.unwrap().tab_stops;
    assert_eq!(tabs.len(), 1);
    assert_eq!(tabs.as_slice()[0].alignment, TabAlignment::Right);

    let mut output = Vec::new();
    let mut writer = RtfWriter::new(&mut output);
    writer.write_document_header().unwrap();
    writer.write_stylesheet(document.stylesheet()).unwrap();
    writer.write_str("}").unwrap();
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reparsed.stylesheet().get(30).unwrap().paragraph.unwrap().tab_stops,
        tabs
    );
}

#[test]
fn rejects_malformed_and_over_limit_tab_definitions() {
    for source in [
        r#"{\rtf1\tx X}"#,
        r#"{\rtf1\tb X}"#,
        r#"{\rtf1\tqr1\tx20 X}"#,
        r#"{\rtf1\tldot0\tx20 X}"#,
        r#"{\rtf1\tqr\tqc\tx20 X}"#,
        r#"{\rtf1\tldot\tleq\tx20 X}"#,
        r#"{\rtf1\tldot\tqr\tx20 X}"#,
        r#"{\rtf1\tqr\tb20 X}"#,
        r#"{\rtf1\tx2147483648 X}"#,
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    let mut over_limit = String::from(r#"{\rtf1"#);
    for position in 0..=MAX_PARAGRAPH_TAB_STOPS {
        over_limit.push_str(&format!(r#"\tx{}"#, position * 20));
    }
    over_limit.push_str(" X}");
    assert!(RtfDocument::parse(&over_limit).is_err());

    assert!(RtfDocument::parse(
        r#"{\rtf1{\*\unknown\tqr1\tx\tlmdot0 ignored}Visible}"#,
    )
    .is_ok());
}
