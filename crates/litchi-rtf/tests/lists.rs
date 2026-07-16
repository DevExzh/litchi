use litchi_rtf::{ListLevelType, RtfDocument, RtfWriter};

const SYNTHETIC: &str = r#"{\rtf1\ansi\ansicpg1250
{\*\listtable
{\*\listpicture}
{\list\listtemplateid42\listhybrid
{\listlevel\levelnfc99\levelnfcn99\leveljc2\leveljcn2\levelfollow1\levelstartat3\levelspace120\levelindent360\levelpicture0\li360\fi-180\tx720\f2
{\leveltext\'02\'00.;}{\levelnumbers\'01;}}
{\listname N\'8a;}{\*\liststylename Style;}\spriority5\listid77}}
{\*\listoverridetable
{\listoverride\listid77\listoverridecount2
{\lfolevel\listoverridestartat\levelstartat9}
{\lfolevel\listoverrideformat}
\ls4}}
\pard\ls4\ilvl0 Body}"#;

#[test]
fn retains_and_resolves_extended_list_metadata() {
    let doc = RtfDocument::parse(SYNTHETIC).unwrap();
    let table = doc.list_table();
    assert_eq!(table.picture_bullet_count, 1);
    let list = table.get(77).unwrap();
    assert_eq!(list.name, "NŠ");
    assert_eq!(list.style_name, "Style");
    assert_eq!(list.style_priority, Some(5));
    let level = &list.levels[0];
    assert_eq!(level.level_type, ListLevelType::Other(99));
    assert_eq!(level.number_text.as_bytes(), b"\0.");
    assert_eq!(level.number_positions.as_bytes(), b"\x01");
    assert_eq!(level.left_indent, Some(360));
    assert_eq!(level.first_line_indent, Some(-180));
    assert_eq!(level.tabs, [720]);
    assert_eq!(level.picture_index, Some(0));

    let list_override = doc.list_override_table().get(4).unwrap();
    assert_eq!(list_override.levels.len(), 2);
    assert_eq!(list_override.levels[0].start_at, Some(9));
    assert!(!list_override.levels[0].format_override);
    assert_eq!(list_override.levels[1].start_at, None);
    assert!(list_override.levels[1].format_override);

    let body = doc.blocks().iter().find(|block| block.text.contains("Body")).unwrap();
    let (resolved_override, resolved_level, start_at) =
        doc.resolve_paragraph_list(&body.paragraph).unwrap();
    assert_eq!(resolved_override.index, 4);
    assert_eq!(resolved_level.level, 0);
    assert_eq!(start_at, Some(9));
}

#[test]
fn writer_round_trips_list_models_deterministically() {
    let doc = RtfDocument::parse(SYNTHETIC).unwrap();
    let mut first = Vec::new();
    RtfWriter::new(&mut first).write_document(&doc).unwrap();
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(doc.list_table(), reparsed.list_table());
    assert_eq!(doc.list_override_table(), reparsed.list_override_table());

    let mut second = Vec::new();
    RtfWriter::new(&mut second).write_document(&reparsed).unwrap();
    assert_eq!(first, second);
}

#[test]
fn rejects_malformed_list_roots_ids_and_overrides() {
    let malformed = [
        r#"{\rtf1{\*\listtable}{\*\listtable}}"#,
        r#"{\rtf1{\*\listoverridetable}{\*\listtable}}"#,
        r#"{\rtf1 Body{\*\listtable}}"#,
        r#"{\rtf1{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0{\leveltext\'01.;}{\levelnumbers;} }\listid1}{\list\listtemplateid2{\listlevel\levelnfc0{\leveltext\'01.;}{\levelnumbers;} }\listid1}}}"#,
        r#"{\rtf1{\*\listtable{\list\listtemplateid1\listsimple\listhybrid{\listlevel\levelnfc0{\leveltext\'01.;}{\levelnumbers;} }\listid1}}}"#,
        r#"{\rtf1{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0{\leveltext\'01.;}{\levelnumbers;} }\listid1}}{\*\listoverridetable{\listoverride\listid2\listoverridecount0\ls1}}}"#,
        r#"{\rtf1{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0{\leveltext\'01.;}{\levelnumbers;} }\listid1}}{\*\listoverridetable{\listoverride\listid1\listoverridecount1\ls1}}}"#,
    ];
    for source in malformed {
        assert!(RtfDocument::parse(source).is_err(), "accepted malformed RTF: {source}");
    }
}

#[test]
fn parses_rare_libreoffice_list_fixtures() {
    let picture = RtfDocument::parse_bytes(include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/i120928.rtf"
    ))
    .unwrap();
    assert!(picture.list_table().picture_bullet_count > 0);
    assert!(picture
        .list_table()
        .lists()
        .iter()
        .flat_map(|list| &list.levels)
        .any(|level| level.picture_index.is_some()));

    let style = RtfDocument::parse_bytes(include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/tdf125719_case_2.rtf"
    ))
    .unwrap();
    assert!(style
        .list_table()
        .lists()
        .iter()
        .any(|list| !list.style_name.is_empty()));

    let starts = RtfDocument::parse_bytes(include_bytes!(
        "../../../3rdparty/libreoffice-core/sw/qa/extras/rtfexport/data/num-override-start.rtf"
    ))
    .unwrap();
    assert!(starts
        .list_override_table()
        .overrides()
        .iter()
        .any(|entry| entry.levels.len() >= 2
            && entry.levels[0].start_at != entry.levels[1].start_at));
}
