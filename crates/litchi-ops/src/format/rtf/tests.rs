use super::*;
use crate::detect::{DetectedFormat, detect_format_smart};
use litchi_core::{Position, detection::FileFormat, metadata::Timestamp};
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

/// The facade's byte route: smart detection, then the RTF arm.
fn open_detected(bytes: Vec<u8>) -> Rtf {
    match detect_format_smart(bytes, &crate::detect::SourceLimits::default()) {
        Some(DetectedFormat::Rtf(bytes)) => {
            Rtf::from_bytes(bytes, &OpenOptions::default()).expect("open detected RTF")
        },
        other => panic!("expected RTF detection, got {other:?}"),
    }
}

#[test]
fn handle_is_one_pointer_and_shares_its_state() {
    assert_eq!(size_of::<Rtf>(), size_of::<usize>());
    let rtf = open_detected(std::fs::read(fixture("rtf/testUnicode.rtf")).expect("read"));
    let clone = rtf.clone();
    assert!(Arc::ptr_eq(&rtf.0, &clone.0));
    assert_eq!(format!("{rtf:?}"), "Rtf { .. }");
}

#[test]
fn detected_bytes_match_the_leaf_reader() {
    for name in ["rtf/testUnicode.rtf", "rtf/chtoutline.rtf"] {
        let bytes = std::fs::read(fixture(name)).expect("read RTF fixture");
        let leaf = Package::from_owned_bytes(bytes.clone()).expect("leaf RTF");
        let rtf = open_detected(bytes);
        assert_eq!(rtf.text().expect("text"), leaf.text(), "{name}");
        assert_eq!(
            rtf.paragraph_count().expect("paragraph count"),
            leaf.paragraph_count(),
            "{name}"
        );
        assert_eq!(
            &rtf.tolerance_report().expect("report"),
            leaf.tolerance_report(),
            "{name}"
        );
    }
}

#[test]
fn compressed_transports_open_through_the_same_route() {
    let plain = br"{\rtf1\ansi Plain owned route\par Second paragraph}".to_vec();
    let expected = open_detected(plain.clone()).text().expect("plain text");
    assert!(expected.contains("Plain owned route"));
    for use_lzfu in [true, false] {
        let compressed = litchi_rtf::transport::compress(
            &plain,
            use_lzfu,
            litchi_rtf::transport::CompressionLimits::default(),
        )
        .expect("compress RTF");
        assert_eq!(open_detected(compressed).text().expect("text"), expected);
    }
}

#[test]
fn metadata_matches_the_facade_projection() {
    let bytes = std::fs::read(fixture("rtf/chtoutline.rtf")).expect("read");
    let leaf = Tree::from_owned_bytes(bytes.clone(), litchi_rtf::read::OpenOptions::default())
        .expect("leaf RTF");
    let info = leaf.info();
    let metadata = open_detected(bytes)
        .metadata()
        .expect("metadata")
        .expect("RTF metadata is Some");
    // `{\author Hunglm}{\operator Hunglm}{\creatim\yr2014\mo11\dy15\hr23\min9}
    // {\revtim\yr2014\mo11\dy15\hr23\min9}`
    assert_eq!(metadata.author.as_deref(), Some("Hunglm"));
    assert_eq!(metadata.last_modified_by.as_deref(), Some("Hunglm"));
    let stamp =
        chrono::NaiveDate::from_ymd_opt(2014, 11, 15).and_then(|date| date.and_hms_opt(23, 9, 0));
    assert_eq!(metadata.created.and_then(Timestamp::unzoned), stamp);
    assert_eq!(metadata.modified.and_then(Timestamp::unzoned), stamp);
    // The remaining fields are the facade's field-for-field projection of
    // the leaf's `\info` model.
    assert_eq!(metadata.title.as_deref(), info.title.as_deref());
    assert_eq!(
        metadata.revision,
        info.version.map(|value| value.to_string())
    );
    assert_eq!(
        metadata.version,
        info.revision.map(|value| value.to_string())
    );
    assert_eq!(metadata.page_count, info.pages);
    assert_eq!(metadata.word_count, info.words);
    assert_eq!(metadata.character_count, info.characters);
    assert_eq!(
        metadata.character_count_with_spaces,
        info.characters_with_spaces
    );
    assert_eq!(metadata.editing_time_minutes, info.editing_time);
    assert_eq!(
        metadata.application,
        leaf.generator()
            .map(|generator| generator.value.to_string())
    );
    // RTF timestamps carry no zone; fields RTF has no source for stay empty.
    assert_eq!(metadata.created.and_then(Timestamp::utc), None);
    assert_eq!(metadata.modified.and_then(Timestamp::utc), None);
    assert_eq!(metadata.language, None);
    assert_eq!(metadata.codepage, None);
}

#[test]
fn malformed_input_is_a_typed_rtf_error() {
    let error = Rtf::from_bytes(b"not rtf at all".to_vec(), &OpenOptions::default())
        .expect_err("non-RTF bytes must not open");
    assert!(matches!(error.leaf(), crate::error::Source::Rtf(_)));
    assert_eq!(error.format(), Some(FileFormat::Rtf));
}

#[test]
fn paragraph_move_matches_the_native_final_position_contract() {
    let input = br"{\rtf1\ansi First\par Second\par Third\par Fourth}".to_vec();
    let source = open_detected(input.clone());

    let native = Package::from_bytes(&input).expect("native RTF");
    let mut native_edit = native.edit();
    native_edit
        .move_paragraph(Position::new(0), Position::new(2))
        .expect("native move");
    let expected = native_edit
        .commit()
        .expect("native commit")
        .into_package()
        .to_bytes()
        .expect("native bytes");

    let actual = source
        .move_to(
            crate::path::PathArg::parse("/body/p[1]").expect("source path"),
            crate::Placement::Position(Position::new(2)),
        )
        .expect("unified move");
    let mut actual_bytes = Vec::new();
    actual
        .after
        .write_to(&mut actual_bytes)
        .expect("unified bytes");
    assert_eq!(actual_bytes, expected);
    let mut source_bytes = Vec::new();
    source.write_to(&mut source_bytes).expect("source bytes");
    assert_eq!(source_bytes, input);
}

#[test]
fn paragraph_move_resolves_before_and_after_after_detach() {
    let source = open_detected(br"{\rtf1\ansi First\par Second\par Third}".to_vec());
    let before = source
        .move_to(
            crate::path::PathArg::parse("/body/p[1]").expect("source path"),
            crate::Placement::before("/body/p[3]").expect("before path"),
        )
        .expect("before move");
    assert_eq!(before.after.text().expect("text"), "Second\nFirst\nThird");

    let after = source
        .move_to(
            crate::path::PathArg::parse("/body/p[3]").expect("source path"),
            crate::Placement::after("/body/p[1]").expect("after path"),
        )
        .expect("after move");
    assert_eq!(after.after.text().expect("text"), "First\nThird\nSecond");
}
