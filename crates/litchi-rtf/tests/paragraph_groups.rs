use litchi_rtf::{BorderStyle, RtfDocument, RtfWriter};

const SYNTHETIC: &str = r#"{\rtf1\ansi\ansicpg1250
{\*\pgptbl
{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}
{\pgp\ipgp1\itap1\li-300\ri120\sb40\sa80\brdrt\brdrs\brdrw15\brsp20\brdrcf3}
{\pgp\ipgp2\itap2\li360\ri0\sb0\sa120}}
Body}"#;

#[test]
fn parses_resolves_and_round_trips_paragraph_group_table() {
    let doc = RtfDocument::parse(SYNTHETIC).unwrap();
    let table = doc.paragraph_group_table().unwrap();
    assert_eq!(table.entries().len(), 3);
    assert_eq!(table.get(2).unwrap().parent_id, 1);
    assert_eq!(table.parent_of(3).unwrap().id, 2);
    let second = table.get(2).unwrap();
    assert_eq!(second.table_nesting_level, 1);
    assert_eq!(second.left_indent, -300);
    assert_eq!(second.right_indent, 120);
    assert_eq!(second.borders.top.style, BorderStyle::Single);
    assert_eq!(second.borders.top.width, 15);
    assert_eq!(second.borders.top.space, 20);
    assert_eq!(second.borders.top.color_ref, 3);
    assert_eq!(doc.text().trim(), "Body");

    let mut first = Vec::new();
    RtfWriter::new(&mut first).write_document(&doc).unwrap();
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(
        doc.paragraph_group_table(),
        reparsed.paragraph_group_table()
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, second);
}

#[test]
fn rejects_malformed_paragraph_group_tables() {
    let malformed = [
        r#"{\rtf1{\pgptbl{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}}}"#,
        r#"{\rtf1{\*\pgptbl}}"#,
        r#"{\rtf1{\*\pgptbl{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}}{\*\pgptbl{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}}}"#,
        r#"{\rtf1 Body{\*\pgptbl{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0}}}"#,
        r#"{\rtf1{\*\pgptbl{\pgp\ipgp0\itap0\li0\ri0\sb0}}}"#,
        r#"{\rtf1{\*\pgptbl{\pgp\ipgp2\itap0\li0\ri0\sb0\sa0}}}"#,
        r#"{\rtf1{\*\pgptbl{\pgp\ipgp2\itap0\li0\ri0\sb0\sa0}{\pgp\ipgp1\itap0\li0\ri0\sb0\sa0}}}"#,
        r#"{\rtf1{\*\pgptbl{\pgp\ipgp0\itap-1\li0\ri0\sb0\sa0}}}"#,
        r#"{\rtf1{\*\pgptbl{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0{\field X}}}}"#,
        r#"{\rtf1{\*\pgptbl{\pgp\ipgp0\itap0\li0\ri0\sb0\sa0\u20320?}}}"#,
    ];
    for source in malformed {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed RTF: {source}"
        );
    }
}

#[test]
fn parses_real_libreoffice_paragraph_group_fixture() {
    let fixture = include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/tdf167569-2.rtf"
    );
    let marker = br"{\*\pgptbl";
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
    let mut isolated = br"{\rtf1\ansi".to_vec();
    isolated.extend_from_slice(&fixture[start..end.unwrap()]);
    isolated.push(b'}');
    let doc = RtfDocument::parse_bytes(&isolated).unwrap();
    let table = doc.paragraph_group_table().unwrap();
    assert_eq!(table.entries().len(), 17);
    assert_eq!(table.get(2).unwrap().parent_id, 13);
    assert_eq!(table.get(13).unwrap().parent_id, 17);
    assert_eq!(table.parent_of(17).unwrap().id, 11);
    assert!(
        table
            .entries()
            .iter()
            .any(|entry| entry.borders.has_any_border())
    );
}
