use super::super::error::RtfError;
use super::codec::parser_classification_error;
use crate::model::document::RtfDocument;

#[test]
fn parser_classification_failures_use_the_typed_error_channel() {
    assert!(matches!(
        parser_classification_error(),
        RtfError::ParserError(message)
            if message == "RTF parser control classification invariant failed"
    ));
}

#[test]
fn accepts_only_nul_padding_and_unicode_whitespace_after_the_root() {
    let mut source = br"{\rtf1\ansi body}".to_vec();
    source.extend_from_slice(b"\0 \t\r\n\xa0\n\0");
    let error = RtfDocument::parse_bytes(&source).err();
    assert!(
        error.is_none(),
        "rejected permitted trailing padding: {error:?}"
    );

    let mut source = String::from(r"{\rtf1\ansi body}");
    source.push('\0');
    let error = RtfDocument::parse(&source).err();
    assert!(
        error.is_none(),
        "rejected permitted trailing NUL padding: {error:?}"
    );
}

#[test]
fn ignores_visible_trailing_data_and_rejects_non_rtf_roots() {
    // Word and LibreOffice stop at the root group's closing brace; visible
    // trailing data is ignored and recorded at its first visible byte.
    for (suffix, visible) in [
        (b" visible".as_slice(), 1),
        (b"\0visible", 1),
        (b"\xa0\0visible", 2),
        (br"\par".as_slice(), 0),
        (br"{trailing}".as_slice(), 0),
        (br"}".as_slice(), 0),
    ] {
        let mut source = br"{\rtf1\ansi body}".to_vec();
        let root_len = source.len();
        source.extend_from_slice(suffix);
        let document = RtfDocument::parse_bytes(&source)
            .unwrap_or_else(|error| panic!("rejected trailing data {suffix:?}: {error}"));
        assert_eq!(document.text(), "body");
        assert_eq!(
            document.tolerance_report().defects(),
            &[crate::leniency::ToleratedDefect {
                defect: crate::leniency::SyntaxDefect::TrailingContentAfterRoot.into(),
                offset: (root_len + visible) as u64,
            }],
            "trailing data {suffix:?}"
        );
    }

    const MISSING_VERSION_ERROR: &str = "RTF \\rtf control requires a version parameter";
    for source in [r"{\rtf}", r"{\rtf }", r"{\rtf\ansi}", r"{\rtf text}"] {
        assert!(
            matches!(
                RtfDocument::parse(source).err(),
                Some(RtfError::MalformedDocument(message)) if message == MISSING_VERSION_ERROR
            ),
            "accepted bare \\rtf root {source:?}"
        );
    }

    const ROOT_ERROR: &str = "RTF document must begin with the supported \\rtf1 root header";
    for source in [r"{\rtf0 body}", r"{\rtf2 body}", r"{\ansi body}"] {
        assert!(
            matches!(
                RtfDocument::parse(source).err(),
                Some(RtfError::MalformedDocument(message)) if message == ROOT_ERROR
            ),
            "accepted invalid RTF root {source:?}"
        );
    }
}
