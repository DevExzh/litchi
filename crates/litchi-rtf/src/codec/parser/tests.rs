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
fn rejects_visible_trailing_data_and_non_rtf_roots() {
    const TRAILING_ERROR: &str = "RTF document contains trailing non-whitespace tokens";

    for suffix in [
        b" visible".as_slice(),
        b"\0visible",
        b"\xa0\0visible",
        br"\par".as_slice(),
        br"{trailing}".as_slice(),
    ] {
        let mut source = br"{\rtf1\ansi body}".to_vec();
        source.extend_from_slice(suffix);
        assert!(
            matches!(
                RtfDocument::parse_bytes(&source).err(),
                Some(RtfError::MalformedDocument(message)) if message == TRAILING_ERROR
            ),
            "accepted visible trailing data {suffix:?}"
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
