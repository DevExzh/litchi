#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use super::super::error::RtfError;
use super::super::limits::ParseLimits;
use super::codec::{CharacterSet, ControlWord, Lexer, Token};
use bumpalo::Bump;
use std::borrow::Cow;

#[test]
fn test_simple_tokenization() {
    let arena = Bump::new();
    let input = r"{\rtf1\ansi Hello}";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();

    assert_eq!(tokens.len(), 5);
    assert!(matches!(tokens[0], Token::OpenBrace));
    assert!(matches!(tokens[1], Token::Control(ControlWord::Rtf(1))));
    assert!(matches!(tokens[2], Token::Control(ControlWord::Ansi)));
}

#[test]
fn test_lexer_new() {
    let arena = Bump::new();
    let lexer = Lexer::new(r"{\rtf1}", &arena);
    assert_eq!(lexer.pos, 0);
}

#[test]
fn test_character_set_variants() {
    assert_eq!(CharacterSet::default(), CharacterSet::Ansi);
    assert_ne!(CharacterSet::Ansi, CharacterSet::Mac);
}

#[test]
fn test_token_variants() {
    let token = Token::OpenBrace;
    assert!(matches!(token, Token::OpenBrace));
    let token = Token::CloseBrace;
    assert!(matches!(token, Token::CloseBrace));
}

#[test]
fn test_control_word_variants() {
    let word = ControlWord::Rtf(1);
    assert!(matches!(word, ControlWord::Rtf(1)));
    let word = ControlWord::Bold(true);
    assert!(matches!(word, ControlWord::Bold(true)));
    let word = ControlWord::FontNumber(0);
    assert!(matches!(word, ControlWord::FontNumber(0)));
}

#[test]
fn test_tokenize_empty_braces() {
    let arena = Bump::new();
    let input = "{}";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 2);
    assert!(matches!(tokens[0], Token::OpenBrace));
    assert!(matches!(tokens[1], Token::CloseBrace));
}

#[test]
fn test_tokenize_control_word_with_param() {
    let arena = Bump::new();
    let input = r"{\rtf1}";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 3);
    assert!(matches!(tokens[1], Token::Control(ControlWord::Rtf(1))));
}

#[test]
fn test_tokenize_control_word_without_param() {
    let arena = Bump::new();
    let input = r"{\b}";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert!(matches!(tokens[1], Token::Control(ControlWord::Bold(true))));
}

#[test]
fn test_tokenize_text() {
    let arena = Bump::new();
    let input = r"{Hello}";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 3);
    assert!(matches!(tokens[1], Token::Text(_)));
}

#[test]
fn plain_text_scan_preserves_utf8_delimiters_and_physical_line_breaks() {
    let arena = Bump::new();
    let input = "alpha 你好\r\n\n beta\\par gamma{delta}";
    let mut lexer = Lexer::new(input, &arena);
    let (tokens, spans) = lexer.tokenize_with_spans().unwrap();

    assert_eq!(
        tokens,
        vec![
            Token::Text(Cow::Borrowed("alpha 你好")),
            Token::Text(Cow::Borrowed(" beta")),
            Token::Control(ControlWord::Par),
            Token::Text(Cow::Borrowed("gamma")),
            Token::OpenBrace,
            Token::Text(Cow::Borrowed("delta")),
            Token::CloseBrace,
        ]
    );
    assert_eq!(
        spans,
        vec![0..13, 13..20, 20..25, 25..30, 30..31, 31..36, 36..37]
    );
}

#[test]
fn test_tokenize_multiple_control_words() {
    let arena = Bump::new();
    let input = r"{\rtf1\ansi\deff0}";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 5);
    assert!(matches!(tokens[1], Token::Control(ControlWord::Rtf(1))));
    assert!(matches!(tokens[2], Token::Control(ControlWord::Ansi)));
}

#[test]
fn test_tokenize_escaped_braces() {
    let arena = Bump::new();
    let input = r"\{\}";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 2);
    // Should be text tokens containing { and }
    assert!(matches!(tokens[0], Token::Text(_)));
    assert!(matches!(tokens[1], Token::Text(_)));
}

#[test]
fn test_tokenize_backslash_escape() {
    let arena = Bump::new();
    let input = r"\\";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(tokens[0], Token::Text(_)));
}

#[test]
fn test_tokenize_hex_escape() {
    let arena = Bump::new();
    let input = r"\'41"; // 'A' in hex
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(tokens[0], Token::Text(_)));
}

#[test]
fn test_hex_escape_preserves_following_literal_space() {
    let arena = Bump::new();
    let mut lexer = Lexer::new(r"\'80 value", &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 2);
    assert_eq!(tokens[0], Token::Text(Cow::Borrowed("\u{0080}")));
    assert_eq!(tokens[1], Token::Text(Cow::Borrowed(" value")));
}

#[test]
fn test_tokenize_asterisk_destination() {
    let arena = Bump::new();
    let input = r"\*";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(
        tokens[0],
        Token::Control(ControlWord::IgnorableDestination)
    ));
}

#[test]
fn test_tokenize_par_control_word() {
    let arena = Bump::new();
    let input = r"\par";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(tokens[0], Token::Control(ControlWord::Par)));
}

#[test]
fn test_tokenize_non_breaking_space() {
    let arena = Bump::new();
    let input = r"\~";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert_eq!(tokens[0], Token::Control(ControlWord::NonBreakingSpace));
}

#[test]
fn test_tokenize_optional_hyphen() {
    let arena = Bump::new();
    let input = r"\-";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert_eq!(tokens[0], Token::Control(ControlWord::OptionalHyphen));
}

#[test]
fn test_tokenize_non_breaking_hyphen() {
    let arena = Bump::new();
    let input = r"\_";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert_eq!(tokens[0], Token::Control(ControlWord::NonBreakingHyphen));
}

#[test]
fn test_tokenize_special_character_control_words() {
    let cases: &[(&str, ControlWord<'_>)] = &[
        (r"\emdash", ControlWord::EmDash),
        (r"\endash", ControlWord::EnDash),
        (r"\emspace", ControlWord::EmSpace),
        (r"\enspace", ControlWord::EnSpace),
        (r"\qmspace", ControlWord::QuarterEmSpace),
        (r"\bullet", ControlWord::Bullet),
        (r"\ltrmark", ControlWord::LeftToRightMark),
        (r"\rtlmark", ControlWord::RightToLeftMark),
        (r"\zwj", ControlWord::ZeroWidthJoiner),
        (r"\zwnj", ControlWord::ZeroWidthNonJoiner),
        (r"\chdate", ControlWord::CurrentDate),
        (r"\chdpl", ControlWord::CurrentDateLong),
        (r"\chdpa", ControlWord::CurrentDateAbbreviated),
        (r"\chtime", ControlWord::CurrentTime),
    ];
    for (input, expected) in cases {
        let arena = Bump::new();
        let mut lexer = Lexer::new(input, &arena);
        let tokens = lexer.tokenize().unwrap();
        assert_eq!(tokens.len(), 1, "tokenized {input}");
        assert_eq!(tokens[0], Token::Control(*expected), "lexed {input}");
    }
}

#[test]
fn test_tokenize_column_break() {
    let arena = Bump::new();
    let input = r"\column";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert_eq!(tokens[0], Token::Control(ControlWord::Column(None)));
}

#[test]
fn test_tokenize_tab() {
    let arena = Bump::new();
    let input = r"\tab";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(tokens[0], Token::Control(ControlWord::Tab)));
}

#[test]
fn test_tokenize_par() {
    let arena = Bump::new();
    let input = r"\par";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(tokens[0], Token::Control(ControlWord::Par)));
}

#[test]
fn test_tokenize_line() {
    let arena = Bump::new();
    let input = r"\line";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(tokens[0], Token::Control(ControlWord::Line)));
}

#[test]
fn test_tokenize_page() {
    let arena = Bump::new();
    let input = r"\page";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(tokens[0], Token::Control(ControlWord::Page(None))));
}

#[test]
fn test_tokenize_font_size() {
    let arena = Bump::new();
    let input = r"\fs24";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(
        tokens[0],
        Token::Control(ControlWord::FontSize(24))
    ));
}

#[test]
fn test_tokenize_font_number() {
    let arena = Bump::new();
    let input = r"\f0";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(
        tokens[0],
        Token::Control(ControlWord::FontNumber(0))
    ));
}

#[test]
fn test_tokenize_bold_toggle() {
    let arena = Bump::new();
    let input = r"\b\b0";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 2);
    assert!(matches!(tokens[0], Token::Control(ControlWord::Bold(true))));
    assert!(matches!(
        tokens[1],
        Token::Control(ControlWord::Bold(false))
    ));
}

#[test]
fn test_tokenize_italic_toggle() {
    let arena = Bump::new();
    let input = r"\i\i0";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 2);
    assert!(matches!(
        tokens[0],
        Token::Control(ControlWord::Italic(true))
    ));
    assert!(matches!(
        tokens[1],
        Token::Control(ControlWord::Italic(false))
    ));
}

#[test]
fn test_tokenize_underline_variants() {
    let arena = Bump::new();
    let input = r"\ul\ulnone\uldb";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 3);
    assert!(matches!(
        tokens[0],
        Token::Control(ControlWord::Underline(true))
    ));
    assert!(matches!(
        tokens[1],
        Token::Control(ControlWord::UnderlineNone)
    ));
    assert!(matches!(
        tokens[2],
        Token::Control(ControlWord::UnderlineDouble)
    ));
}

#[test]
fn test_tokenize_alignment() {
    let arena = Bump::new();
    let input = r"\ql\qr\qc\qj";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 4);
    assert!(matches!(tokens[0], Token::Control(ControlWord::LeftAlign)));
    assert!(matches!(tokens[1], Token::Control(ControlWord::RightAlign)));
    assert!(matches!(tokens[2], Token::Control(ControlWord::Center)));
    assert!(matches!(tokens[3], Token::Control(ControlWord::Justify)));
}

#[test]
fn test_tokenize_color_table() {
    let arena = Bump::new();
    let input = r"{\colortbl;\red255\green0\blue0;}";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert!(matches!(tokens[1], Token::Control(ControlWord::ColorTable)));
    assert!(matches!(tokens[3], Token::Control(ControlWord::Red(255))));
    assert!(matches!(tokens[4], Token::Control(ControlWord::Green(0))));
    assert!(matches!(tokens[5], Token::Control(ControlWord::Blue(0))));
}

#[test]
fn test_tokenize_font_table() {
    let arena = Bump::new();
    let input = r"{\fonttbl{\f0\fnil Arial;}}";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert!(matches!(tokens[1], Token::Control(ControlWord::FontTable)));
}

#[test]
fn test_tokenize_unknown_control_word() {
    let arena = Bump::new();
    let input = r"\xyz123";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(
        tokens[0],
        Token::Control(ControlWord::Unknown(_, _))
    ));
}

#[test]
fn test_tokenize_section_break() {
    let arena = Bump::new();
    let input = r"\sect";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 1);
    assert!(matches!(tokens[0], Token::Control(ControlWord::Section)));
}

#[test]
fn test_tokenize_page_dimensions() {
    let arena = Bump::new();
    let input = r"\paperw12240\paperh15840";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert!(matches!(
        tokens[0],
        Token::Control(ControlWord::PageWidth(12240))
    ));
    assert!(matches!(
        tokens[1],
        Token::Control(ControlWord::PageHeight(15840))
    ));
}

#[test]
fn test_tokenize_margins() {
    let arena = Bump::new();
    let input = r"\margl1440\margr1440\margt1440\margb1440";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert!(matches!(
        tokens[0],
        Token::Control(ControlWord::MarginLeft(1440))
    ));
    assert!(matches!(
        tokens[1],
        Token::Control(ControlWord::MarginRight(1440))
    ));
    assert!(matches!(
        tokens[2],
        Token::Control(ControlWord::MarginTop(1440))
    ));
    assert!(matches!(
        tokens[3],
        Token::Control(ControlWord::MarginBottom(1440))
    ));
}

#[test]
fn test_tokenize_lists() {
    let arena = Bump::new();
    let input = r"\listtable\listid1\listlevel1";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 3);
    assert!(matches!(tokens[0], Token::Control(ControlWord::ListTable)));
    assert!(matches!(tokens[1], Token::Control(ControlWord::ListId(1))));
    assert!(matches!(tokens[2], Token::Control(ControlWord::ListLevel)));
}

#[test]
fn test_tokenize_table() {
    let arena = Bump::new();
    let input = r"\trowd\cellx4320";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert_eq!(tokens.len(), 2);
    assert!(matches!(
        tokens[0],
        Token::Control(ControlWord::TableRowDefaults)
    ));
    assert!(matches!(
        tokens[1],
        Token::Control(ControlWord::CellX(4320))
    ));
}

#[test]
fn test_tokenize_field() {
    let arena = Bump::new();
    let input = r"\field\fldinst HYPERLINK";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert!(matches!(tokens[0], Token::Control(ControlWord::Field)));
    assert!(matches!(
        tokens[1],
        Token::Control(ControlWord::FieldInstruction)
    ));
    // HYPERLINK is parsed as text since it's after a space
    assert!(matches!(tokens[2], Token::Text(_)));
}

#[test]
fn test_tokenize_picture() {
    let arena = Bump::new();
    let input = r"{\pict\pngblip\picw100\pich100}";
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert!(matches!(tokens[1], Token::Control(ControlWord::Picture)));
    assert!(matches!(tokens[2], Token::Control(ControlWord::Pngblip)));
    assert!(matches!(
        tokens[3],
        Token::Control(ControlWord::PictureWidth(100))
    ));
    assert!(matches!(
        tokens[4],
        Token::Control(ControlWord::PictureHeight(100))
    ));
}

#[test]
fn test_tokenize_binary() {
    let arena = Bump::new();
    let input = r"\bin4 ABCD"; // 4 bytes of binary data
    let mut lexer = Lexer::new(input, &arena);
    let tokens = lexer.tokenize().unwrap();
    assert!(matches!(&tokens[0], Token::Binary(data) if data.as_ref() == b"ABCD"));
}

#[test]
fn test_rejects_truncated_binary_data() {
    let arena = Bump::new();
    let mut lexer = Lexer::new(r"\bin4 AB", &arena);
    assert!(matches!(lexer.tokenize(), Err(RtfError::UnexpectedEof)));
}

#[test]
fn binary_declaration_is_validated_before_allocation() {
    let arena = Bump::new();
    let limits = ParseLimits::default()
        .with_max_binary_bytes(usize::MAX)
        .with_max_total_binary_bytes(usize::MAX);
    let mut lexer = Lexer::new_with_limits(r"\bin2147483647 x", &arena, limits);
    assert!(matches!(lexer.tokenize(), Err(RtfError::UnexpectedEof)));
}

#[test]
fn binary_limits_report_per_payload_and_aggregate_usage() {
    let arena = Bump::new();
    let limits = ParseLimits::default().with_max_binary_bytes(3);
    let mut lexer = Lexer::new_with_limits(r"\bin4 ABCD", &arena, limits);
    assert!(matches!(
        lexer.tokenize(),
        Err(RtfError::LimitExceeded {
            resource: "binary payload bytes",
            observed: 4,
            limit: 3,
        })
    ));

    let arena = Bump::new();
    let limits = ParseLimits::default()
        .with_max_binary_bytes(2)
        .with_max_total_binary_bytes(3);
    let mut lexer = Lexer::new_with_limits(r"\bin2 ab\bin2 cd", &arena, limits);
    assert!(matches!(
        lexer.tokenize(),
        Err(RtfError::LimitExceeded {
            resource: "aggregate binary payload bytes",
            observed: 4,
            limit: 3,
        })
    ));
}

#[test]
fn token_limit_is_checked_before_emitting_an_extra_token() {
    let arena = Bump::new();
    let limits = ParseLimits::default().with_max_tokens(1);
    let mut lexer = Lexer::new_with_limits("{}", &arena, limits);
    assert!(matches!(
        lexer.tokenize(),
        Err(RtfError::LimitExceeded {
            resource: "lexer tokens",
            observed: 2,
            limit: 1,
        })
    ));
}

#[test]
fn plain_text_borrows_the_source_and_binary_preserves_latin1_bytes() {
    let arena = Bump::new();
    let mut lexer = Lexer::new("plain text", &arena);
    let tokens = lexer.tokenize().unwrap();
    assert!(matches!(
        tokens.as_slice(),
        [Token::Text(Cow::Borrowed("plain text"))]
    ));

    let arena = Bump::new();
    let mut lexer = Lexer::new("\\bin2 \u{80}\u{ff}", &arena);
    let tokens = lexer.tokenize().unwrap();
    assert!(matches!(&tokens[0], Token::Binary(data) if data.as_ref() == [0x80, 0xff]));
}

#[test]
fn invalid_private_utf8_cursor_returns_a_typed_error() {
    let arena = Bump::new();
    let mut lexer = Lexer::new("你", &arena);
    lexer.pos = 1;
    assert!(matches!(
        lexer.next_token(),
        Err(RtfError::MalformedDocument(message))
            if message.contains("UTF-8 boundary")
    ));
}

#[test]
fn test_tokenize_document_with_trailing_line_breaks() {
    let arena = Bump::new();
    let mut lexer = Lexer::new("{\\rtf1 body}\r\n", &arena);
    let tokens = lexer.tokenize().unwrap();

    assert!(matches!(tokens.last(), Some(Token::Text(text)) if text.is_empty()));
    assert!(matches!(
        tokens.get(tokens.len().saturating_sub(2)),
        Some(Token::CloseBrace)
    ));
}
