#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_rtf::{RtfDocument, RtfError};

const STORY_FONT_CODEPAGES: &str = r"{\rtf1\ansi\ansicpg1252\uc1
{\fonttbl{\f0\fnil\fcharset0 ANSI;}{\f1\fnil\fcharset128 JIS;}{\f2\fnil\fcharset0\cpg932 Explicit;}}
{\header\f1\'82\'a0}
{\footer\f2\'82\'a0}
Body
{\footnote\f1\'82\'a0}
{\endnote\f2\'82\'a0}
}";

#[test]
fn stories_decode_hex_bytes_with_the_selected_font_code_page() {
    let document = RtfDocument::parse(STORY_FONT_CODEPAGES).unwrap();

    assert_eq!(document.sections()[0].headers_footers[0].text(), "あ");
    assert_eq!(document.sections()[0].headers_footers[1].text(), "あ");
    assert_eq!(document.notes()[0].content, "あ");
    assert_eq!(document.notes()[1].content, "あ");
}

const STORY_UNICODE_FALLBACKS: &str = r"{\rtf1\ansi\ansicpg1252
{\fonttbl{\f0\fnil\fcharset128 JIS;}}
{\header\f0\uc1\u12354?\'82\'a0}
{\footer\f0\uc2\u12354\'82\'a0}
Body
{\footnote\f0\uc1\u12354?\'82\'a0}
{\endnote\f0\uc2\u12354\'82\'a0}
}";

#[test]
fn story_unicode_fallbacks_skip_the_declared_count_and_decode_remainder_by_font() {
    let document = RtfDocument::parse(STORY_UNICODE_FALLBACKS).unwrap();

    assert_eq!(document.sections()[0].headers_footers[0].text(), "ああ");
    assert_eq!(document.sections()[0].headers_footers[1].text(), "あ");
    assert_eq!(document.notes()[0].content, "ああ");
    assert_eq!(document.notes()[1].content, "あ");
}

const NESTED_STORY_FONT_RESTORATION: &str = r"{\rtf1\ansi\ansicpg1252
{\fonttbl{\f0\fnil\fcharset0 ANSI;}{\f1\fnil\fcharset128 JIS;}}
{\header\f0 A{\f1\'82\'a0}\'e9}
{\footer\f0 A{\f1\'82\'a0}\'e9}
Body
{\footnote\f0 A{\f1\'82\'a0}\'e9}
{\endnote\f0 A{\f1\'82\'a0}\'e9}
}";

#[test]
fn stories_restore_the_outer_font_after_nested_groups() {
    let document = RtfDocument::parse(NESTED_STORY_FONT_RESTORATION).unwrap();

    assert_eq!(document.sections()[0].headers_footers[0].text(), "Aあé");
    assert_eq!(document.sections()[0].headers_footers[1].text(), "Aあé");
    assert_eq!(document.notes()[0].content, "Aあé");
    assert_eq!(document.notes()[1].content, "Aあé");
}

const STORY_FONT_TABLE: &str = r"{\fonttbl{\f0\fnil\fcharset0 ANSI;}{\f1\fnil\fcharset128 JIS;}}";

fn malformed_story(destination: &str, payload: &str) -> String {
    let mut source = String::from(r"{\rtf1\ansi\ansicpg1252");
    source.push_str(STORY_FONT_TABLE);
    source.push_str("{\\");
    source.push_str(destination);
    source.push_str("\\f1");
    source.push_str(payload);
    source.push_str("}Body}");
    source
}

#[test]
fn malformed_cp932_story_bytes_fail_with_typed_errors() {
    for destination in ["header", "footer", "footnote", "endnote"] {
        for payload in [r"\'82", r"\'82\'7f"] {
            let source = malformed_story(destination, payload);
            assert!(
                matches!(
                    RtfDocument::parse(&source),
                    Err(RtfError::MalformedDocument(_))
                ),
                "accepted malformed {destination} payload {payload}"
            );
        }
    }
}
