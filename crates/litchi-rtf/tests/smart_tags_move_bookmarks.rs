use litchi_rtf::{Document, RtfResult, SmartTagAttribute};
use std::borrow::Cow;

#[test]
fn parses_smart_tag_and_move_bookmark_metadata() -> RtfResult<()> {
    let source = r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}{\xmlattr\xmlattrns0{\xmlattrname Month}{\xmlattrvalue 4}}}4/11{\*\xmlclose}{\*\mvfmf tag 3412D4C3B2A1}m{\*\mvfml tag}{\*\mvtof tag 7856D4C3B2A1}n{\*\mvtol tag}}"#;
    let document = Document::parse(source)?;
    assert_eq!(document.text(), "4/11mn");
    assert_eq!(document.smart_tags().len(), 1);
    assert_eq!(document.smart_tags()[0].name, "date");
    assert_eq!(document.smart_tags()[0].namespace, Some(2));
    assert_eq!(document.smart_tags()[0].attributes[0].namespace, Some(0));
    assert_eq!(document.smart_tags()[0].content, "4/11");
    assert_eq!(document.move_bookmarks().len(), 2);
    assert_eq!(document.move_bookmarks()[0].author, 0x1234);
    assert_eq!(document.move_bookmarks()[0].date, 0xA1B2C3D4);
    assert_eq!(document.move_bookmarks()[0].content, "m");
    assert_eq!(
        document.move_bookmarks()[0].tag,
        document.move_bookmarks()[1].tag
    );
    assert_eq!(
        document.move_bookmarks()[1].kind,
        litchi_rtf::MoveBookmarkKind::To
    );
    assert_eq!(document.move_bookmarks()[1].content, "n");
    Ok(())
}

#[test]
fn smart_tag_accepts_zero_xml_namespace_reference() -> RtfResult<()> {
    let source = r#"{\rtf1\ansi{\*\xmlopen\xmlns0{\factoidname date}}x{\*\xmlclose}}"#;
    let document = Document::parse(source)?;
    assert_eq!(document.smart_tags().len(), 1);
    assert_eq!(document.smart_tags()[0].namespace, Some(0));
    assert_eq!(document.smart_tags()[0].content, "x");
    assert_eq!(document.to_bytes().unwrap(), source.as_bytes());
    Ok(())
}

#[test]
fn smart_tag_accepts_normative_direct_attribute_pcdata_and_preserves_source()
-> Result<(), Box<dyn std::error::Error>> {
    let source = r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}{\xmlattr\xmlattrns0\xmlattrname Month\xmlattrvalue 4}}4/11{\*\xmlclose}}"#;
    let document = Document::parse(source)?;
    assert_eq!(document.smart_tags()[0].attributes[0].namespace, Some(0));
    assert_eq!(document.smart_tags()[0].attributes[0].name, "Month");
    assert_eq!(document.smart_tags()[0].attributes[0].value, "4");
    assert_eq!(document.smart_tags()[0].content, "4/11");
    assert_eq!(document.to_bytes().unwrap(), source.as_bytes());

    let mut edited_tag = document.smart_tags()[0].clone();
    edited_tag.name = Cow::Borrowed("month");
    let mut edit = document.edit();
    edit.set_smart_tag(0, edited_tag.clone())?;
    let commit = edit.commit()?;
    let bytes = commit.snapshot().to_bytes()?;
    let reopened = Document::from_bytes(&bytes)?;
    assert_eq!(reopened.smart_tags()[0].name, edited_tag.name);
    assert_eq!(reopened.smart_tags()[0].attributes[0].value, "4");
    Ok(())
}

#[test]
fn unmatched_move_start_is_ignored_in_the_semantic_model() -> Result<(), Box<dyn std::error::Error>>
{
    let source = r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}}x{\*\xmlclose}{\*\mvfmf orphan 0100EFCDAB00}}"#;
    let document = Document::parse(source)?;
    assert_eq!(document.smart_tags().len(), 1);
    assert!(document.move_bookmarks().is_empty());
    assert_eq!(document.text(), "x");
    assert_eq!(document.to_bytes()?, source.as_bytes());
    Ok(())
}

#[test]
fn malformed_move_marker_fields_are_rejected() {
    assert!(Document::parse(r#"{\rtf1\ansi{\*\mvfmf bad-tag 0100EFCDAB00}}"#).is_err());
    assert!(Document::parse(r#"{\rtf1\ansi{\*\mvfml tag trailing}}"#).is_err());
    for source in [
        r#"{\rtf1\ansi{\*\mvfmf1 tag 0100EFCDAB00}x{\*\mvfml tag}}"#,
        r#"{\rtf1\ansi{\*\mvfml1 tag}}"#,
        r#"{\rtf1\ansi{\*\mvtof1 tag 0100EFCDAB00}x{\*\mvtol tag}}"#,
        r#"{\rtf1\ansi{\*\mvtol1 tag}}"#,
    ] {
        assert!(Document::parse(source).is_err());
    }
}

#[test]
fn smart_tags_and_moves_are_main_body_only() {
    let smart_tag_in_header =
        r#"{\rtf1\ansi{\header{\*\xmlopen\xmlns1{\factoidname date}}x{\*\xmlclose}}body}"#;
    let move_in_footer = r#"{\rtf1\ansi{\footer{\*\mvfmf tag 3412D4C3B2A1}x{\*\mvfml tag}}body}"#;
    let smart = Document::parse(smart_tag_in_header).unwrap();
    assert!(smart.smart_tags().is_empty());
    assert_eq!(smart.to_bytes().unwrap(), smart_tag_in_header.as_bytes());
    let moved = Document::parse(move_in_footer).unwrap();
    assert!(moved.move_bookmarks().is_empty());
    assert_eq!(moved.to_bytes().unwrap(), move_in_footer.as_bytes());
}

#[test]
fn smart_tag_enforces_page_153_namespace_and_attribute_grammar() {
    let missing_xml_namespace = r#"{\rtf1\ansi{\*\xmlopen{\factoidname date}}x{\*\xmlclose}}"#;
    assert!(Document::parse(missing_xml_namespace).is_err());

    let xmlattr_alias = r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}{\xmlattr1{\xmlattrname Month}{\xmlattrvalue 4}}}x{\*\xmlclose}}"#;
    assert!(Document::parse(xmlattr_alias).is_err());

    let missing_attribute_namespace = r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}{\xmlattr{\xmlattrname Month}{\xmlattrvalue 4}}}x{\*\xmlclose}}"#;
    assert!(Document::parse(missing_attribute_namespace).is_err());

    let value_before_name = r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}{\xmlattr\xmlattrns0{\xmlattrvalue 4}{\xmlattrname Month}}}x{\*\xmlclose}}"#;
    assert!(Document::parse(value_before_name).is_err());

    let direct_value_before_name = r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}{\xmlattr\xmlattrns0\xmlattrvalue 4\xmlattrname Month}}x{\*\xmlclose}}"#;
    assert!(Document::parse(direct_value_before_name).is_err());

    let negative_namespace = r#"{\rtf1\ansi{\*\xmlopen\xmlns-1{\factoidname date}}x{\*\xmlclose}}"#;
    assert!(Document::parse(negative_namespace).is_err());

    let negative_attribute_namespace = r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}{\xmlattr\xmlattrns-1{\xmlattrname Month}{\xmlattrvalue 4}}}x{\*\xmlclose}}"#;
    assert!(Document::parse(negative_attribute_namespace).is_err());

    let missing_namespace_parameter =
        r#"{\rtf1\ansi{\*\xmlopen\xmlns{\factoidname date}}x{\*\xmlclose}}"#;
    assert!(Document::parse(missing_namespace_parameter).is_err());

    let missing_attribute_namespace_parameter = r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}{\xmlattr\xmlattrns{\xmlattrname Month}{\xmlattrvalue 4}}}x{\*\xmlclose}}"#;
    assert!(Document::parse(missing_attribute_namespace_parameter).is_err());

    for source in [
        r#"{\rtf1\ansi{\*\xmlopen1\xmlns2{\factoidname date}}x{\*\xmlclose}}"#,
        r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname1 date}}x{\*\xmlclose}}"#,
        r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}}x{\*\xmlclose1}}"#,
        r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}{\xmlattr\xmlattrns0{\xmlattrname1 Month}{\xmlattrvalue 4}}}x{\*\xmlclose}}"#,
        r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}{\xmlattr\xmlattrns0{\xmlattrname Month}{\xmlattrvalue1 4}}}x{\*\xmlclose}}"#,
    ] {
        assert!(Document::parse(source).is_err());
    }
}

fn smart_tag_with_attribute_count(count: usize) -> String {
    let mut source = String::from(r#"{\rtf1\ansi{\*\xmlopen\xmlns1{\factoidname date}"#);
    for index in 0..count {
        source.push_str(&format!(
            r#"{{\xmlattr\xmlattrns0{{\xmlattrname a{index}}}{{\xmlattrvalue v}}}}"#
        ));
    }
    source.push_str(r#"}x{\*\xmlclose}}"#);
    source
}

#[test]
fn smart_tag_attribute_count_is_checked_before_parsing_the_next_attribute() {
    assert!(Document::parse(&smart_tag_with_attribute_count(1_024)).is_ok());
    assert!(Document::parse(&smart_tag_with_attribute_count(1_025)).is_err());
}

#[test]
fn smart_tag_and_move_aggregate_limits_are_checked_at_the_boundary() -> RtfResult<()> {
    const LIMIT: usize = 16 * 1_048_576;

    let mut smart_source = String::from(r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}}"#);
    smart_source.push_str(&"x".repeat(LIMIT - "date".len()));
    smart_source.push_str(r#"{\*\xmlclose}}"#);
    assert!(Document::parse(&smart_source).is_ok());
    let close = smart_source.len() - r#"{\*\xmlclose}}"#.len();
    smart_source.insert(close, 'x');
    assert!(Document::parse(&smart_source).is_err());

    let mut move_source = String::from(r#"{\rtf1\ansi{\*\mvfmf tag 3412D4C3B2A1}"#);
    move_source.push_str(&"x".repeat(LIMIT - "tag".len()));
    move_source.push_str(r#"{\*\mvfml tag}}"#);
    assert!(Document::parse(&move_source).is_ok());
    let close = move_source.len() - r#"{\*\mvfml tag}}"#.len();
    move_source.insert(close, 'x');
    assert!(Document::parse(&move_source).is_err());
    Ok(())
}

#[test]
fn canonical_writer_emits_metadata_and_reopens() -> RtfResult<()> {
    let source = r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}}x{\*\xmlclose}{\*\mvfmf tag 3412D4C3B2A1}m{\*\mvfml tag}}"#;
    let document = Document::parse(source)?;
    let mut output = Vec::new();
    let mut writer = litchi_rtf::write::Writer::with_options(
        &mut output,
        litchi_rtf::write::Options {
            indent: true,
            ..Default::default()
        },
    );
    writer.write(&document).unwrap();
    let written = String::from_utf8(output).unwrap();
    assert!(written.contains(r#"{\*\xmlopen"#));
    assert!(written.contains(r#"{\*\mvfmf tag 3412D4C3B2A1}"#));
    let reopened = Document::parse(&written)?;
    assert_eq!(reopened.smart_tags()[0].content, "x");
    assert_eq!(reopened.move_bookmarks()[0].content, "m");
    Ok(())
}

#[test]
fn source_bound_metadata_edits_are_atomic_reversible_and_library_generated()
-> Result<(), Box<dyn std::error::Error>> {
    let source = Document::parse(
        r#"{\rtf1\ansi{\*\xmlopen\xmlns2{\factoidname date}}x{\*\xmlclose}{\*\mvfmf tag 3412D4C3B2A1}m{\*\mvfml tag}}"#,
    )?;
    let mut smart_tag = source.smart_tags()[0].clone();
    smart_tag.name = Cow::Borrowed("time");
    smart_tag.namespace = Some(3);
    smart_tag.attributes = vec![SmartTagAttribute::new(
        Some(1),
        Cow::Borrowed("format"),
        Cow::Borrowed("short"),
    )?];
    let mut move_bookmark = source.move_bookmarks()[0].clone();
    move_bookmark.author = 7;
    move_bookmark.date = 0x0102_0304;

    let mut edit = source.edit();
    edit.set_smart_tag(0, smart_tag.clone())?
        .set_move_bookmark(0, move_bookmark.clone())?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.snapshot().text(), source.text());
    assert_eq!(commit.snapshot().smart_tags()[0], smart_tag);
    assert_eq!(commit.snapshot().move_bookmarks()[0], move_bookmark);
    assert!(
        commit
            .patch()
            .apply(&source)?
            .same_snapshot(commit.snapshot())
    );
    assert!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())?
            .same_snapshot(&source)
    );
    let foreign = Document::parse(r#"{\rtf1\ansi other}"#)?;
    assert!(matches!(
        commit.patch().apply(&foreign),
        Err(litchi_rtf::edit::Error::PatchConflict)
    ));

    let bytes = commit.snapshot().to_bytes()?;
    std::fs::write("/var/tmp/litchi-spec-gap-smart-move.rtf", &bytes).map_err(|error| {
        litchi_rtf::Error::MalformedDocument(format!("artifact write failed: {error}"))
    })?;
    let reopened = Document::from_bytes(&bytes)?;
    assert_eq!(reopened.smart_tags()[0], smart_tag);
    assert_eq!(reopened.move_bookmarks()[0], move_bookmark);

    let mut noop = source.edit();
    noop.set_smart_tag(0, source.smart_tags()[0].clone())?
        .set_move_bookmark(0, source.move_bookmarks()[0].clone())?;
    let noop = noop.commit()?;
    assert!(!noop.diagnostics().changed());
    assert!(noop.snapshot().same_snapshot(&source));
    Ok(())
}

#[test]
fn changed_publication_refuses_unmatched_move_source() -> Result<(), Box<dyn std::error::Error>> {
    let source_bytes = br#"{\rtf1\ansi{\*\mvfmf orphan 0100EFCDAB00}{\*\xmlopen\xmlns2{\factoidname date}}x{\*\xmlclose}}"#;
    let source = Document::parse(std::str::from_utf8(source_bytes)?)?;
    let mut noop = source.edit();
    noop.set_smart_tag(0, source.smart_tags()[0].clone())?;
    let noop = noop.commit()?;
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.snapshot().to_bytes()?, source_bytes);

    let mut edit = source.edit();
    edit.set_smart_tag(0, {
        let mut value = source.smart_tags()[0].clone();
        value.name = Cow::Borrowed("time");
        value
    })?;
    assert!(matches!(
        edit.commit(),
        Err(litchi_rtf::edit::Error::UnsupportedSource(
            "changed publication refuses unmatched move-bookmark destinations"
        ))
    ));
    assert_eq!(source.to_bytes()?, source_bytes);
    Ok(())
}

#[test]
fn changed_publication_refuses_unmatched_move_end() -> Result<(), Box<dyn std::error::Error>> {
    let source_bytes =
        br#"{\rtf1\ansi{\*\mvfml orphan}{\*\xmlopen\xmlns2{\factoidname date}}x{\*\xmlclose}}"#;
    let source = Document::parse(std::str::from_utf8(source_bytes)?)?;
    let mut edit = source.edit();
    let mut value = source.smart_tags()[0].clone();
    value.name = Cow::Borrowed("time");
    edit.set_smart_tag(0, value)?;
    assert!(matches!(
        edit.commit(),
        Err(litchi_rtf::edit::Error::UnsupportedSource(
            "changed publication refuses unmatched move-bookmark destinations"
        ))
    ));
    assert_eq!(source.to_bytes()?, source_bytes);
    Ok(())
}

#[test]
fn raw_canonical_writer_refuses_unmatched_move_source() -> Result<(), Box<dyn std::error::Error>> {
    let source =
        litchi_rtf::raw::Document::parse(r#"{\rtf1\ansi{\*\mvfmf orphan 0100EFCDAB00}x}"#)?;
    let mut output = Vec::new();
    let error = litchi_rtf::write::Writer::new(&mut output).write_document(&source);
    assert!(matches!(error, Err(error) if error.kind() == std::io::ErrorKind::InvalidInput));
    assert!(output.is_empty());
    Ok(())
}

#[test]
fn duplicate_move_tag_within_one_location_is_rejected() {
    let duplicate_open = r#"{\rtf1\ansi{\*\mvfmf tag 3412D4C3B2A1}{\*\mvfmf tag 3412D4C3B2A1}x{\*\mvfml tag}{\*\mvfml tag}}"#;
    assert!(Document::parse(duplicate_open).is_err());

    let duplicate_completed = r#"{\rtf1\ansi{\*\mvfmf tag 3412D4C3B2A1}x{\*\mvfml tag}{\*\mvfmf tag 3412D4C3B2A1}y{\*\mvfml tag}}"#;
    assert!(Document::parse(duplicate_completed).is_err());
}

#[test]
fn smart_tag_count_is_checked_before_event_growth() {
    let mut source = String::with_capacity(3_500_000);
    source.push_str(r#"{\rtf1\ansi"#);
    for _ in 0..=65_536 {
        source.push_str(r#"{\*\xmlopen\xmlns1{\factoidname x}}{\*\xmlclose}"#);
    }
    source.push('}');
    assert!(Document::parse(&source).is_err());
}

#[test]
fn move_bookmark_indexing_handles_the_safety_limit_without_quadratic_scans() {
    const COUNT: usize = 65_536;
    let mut source = String::with_capacity(COUNT * 64 + 16);
    source.push_str(r#"{\rtf1\ansi"#);
    for index in 0..COUNT {
        let tag = format!("m{index:05X}");
        source.push_str(&format!(
            r#"{{\*\mvfmf {tag} 0100EFCDAB00}}{{\*\mvfml {tag}}}"#
        ));
    }
    source.push('}');

    let document = Document::parse(&source).expect("bounded move-bookmark parse");
    assert_eq!(document.move_bookmarks().len(), COUNT);
}
