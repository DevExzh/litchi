use litchi_rtf::{FontPage, FontPitch, RtfDocument, RtfWriter};

const SYNTHETIC: &str = r#"{\rtf1\ansi\ansicpg1250\uc1
{\fonttbl
{\f0\fswiss\fcharset238\fprq2\cpg1250{\*\panose 020b0604020202020204}{\*\fname Non\u20320?}{\*\falt \'8a Alt}Primary\u20320?;}
{\f2\fmodern\fcharset0\fprq1{\*\panose 02070309020205020404}Courier New;}}
\f0 Body}"#;

#[test]
fn parses_decodes_and_round_trips_extended_font_metadata() {
    let doc = RtfDocument::parse(SYNTHETIC).unwrap();
    let table = doc.font_table();
    let primary = table.get(0).unwrap();
    assert_eq!(primary.name, "Primary你");
    assert_eq!(primary.alternate_name.as_deref(), Some("Š Alt"));
    assert_eq!(primary.non_tagged_name.as_deref(), Some("Non你"));
    assert_eq!(primary.panose, Some([2, 11, 6, 4, 2, 2, 2, 2, 2, 4]));
    assert_eq!(primary.pitch, FontPitch::Variable);
    assert_eq!(primary.code_page.map(FontPage::id), Some(1250));
    assert!(!table.is_defined(1));
    assert!(table.get(1).is_none());
    assert_eq!(table.get(2).unwrap().pitch, FontPitch::Fixed);

    let mut first = Vec::new();
    RtfWriter::new(&mut first).write_document(&doc).unwrap();
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(doc.font_table(), reparsed.font_table());
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn font_pages_decode_names_independently_from_the_header() {
    let source = r#"{\rtf1\ansi\ansicpg1252
{\fonttbl
{\f0\fnil\fcharset128{\*\fname \'82\'a0;}\'82\'a2{\*\falt \'82\'a4};}
{\f1\fnil\fcharset0\cpg932{\*\fname \'82\'a0;}\'82\'a2{\*\falt \'82\'a4};}}}"#;
    let doc = RtfDocument::parse(source).unwrap();

    for font_ref in [0, 1] {
        let font = doc.font_table().get(font_ref).unwrap();
        assert_eq!(font.name, "い");
        assert_eq!(font.non_tagged_name.as_deref(), Some("あ"));
        assert_eq!(font.alternate_name.as_deref(), Some("う"));
    }
}

#[test]
fn rejects_malformed_extended_font_metadata() {
    let malformed = [
        r#"{\rtf1{\fonttbl{\f0\fnil Arial;}}{\fonttbl{\f1\fnil Times;}}}"#,
        r#"{\rtf1 Body{\fonttbl{\f0\fnil Arial;}}}"#,
        r#"{\rtf1{\fonttbl{\f0\fnil Arial;}{\f0\fnil Times;}}}"#,
        r#"{\rtf1{\fonttbl{\f-1\fnil Arial;}}}"#,
        r#"{\rtf1{\fonttbl{\f0\fnil\fprq3 Arial;}}}"#,
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\panose 0202}Arial;}}}"#,
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\panose 0202060305040502030Z}Arial;}}}"#,
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\falt A}{\*\falt B}Arial;}}}"#,
        r#"{\rtf1{\fonttbl{\f0\fnil{\*\falt {\field X}}Arial;}}}"#,
        r#"{\rtf1{\fonttbl{\f0\fnil;}}}"#,
        r#"{\rtf1{\fonttbl{\f0\fnil\cpg70000 Arial;}}}"#,
        r#"{\rtf1{\fonttbl{\f0\fnil\cpg1200 Arial;}}}"#,
        r#"{\rtf1{\fonttbl{\f0\fnil\cpg65000 Arial;}}}"#,
        r#"{\rtf1{\fonttbl{\f0\fnil\fcharset3 Arial;}}}"#,
        r#"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0\fnil\fcharset78 \'82\'a0;}}}"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed RTF: {source}"
        );
    }
}

#[test]
fn parses_real_libreoffice_font_metadata_fixture() {
    let fixture = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/core/data/rtf/fail/forcepoint-5.rtf"
    );
    let marker = br"{\fonttbl";
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
    let mut isolated = br"{\rtf1\ansi\ansicpg1252".to_vec();
    isolated.extend_from_slice(&fixture[start..end.unwrap()]);
    isolated.push(b'}');
    let doc = RtfDocument::parse_bytes(&isolated).unwrap();
    let first = doc.font_table().get(0).unwrap();
    assert_eq!(first.name, "Times New Roman");
    assert_eq!(first.alternate_name.as_deref(), Some("Bookman Old Style"));
    assert_eq!(first.panose, Some([2, 2, 6, 3, 5, 4, 5, 2, 3, 4]));
    assert_eq!(first.pitch, FontPitch::Variable);
}
