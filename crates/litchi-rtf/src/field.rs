//! Safe, structured RTF field-code support.

use std::borrow::Cow;

const MAX_INSTRUCTION_LEN: usize = 65_536;
const MAX_TOKENS: usize = 256;

/// Field type in RTF documents.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldType {
    Hyperlink,
    Reference,
    PageReference,
    NoteReference,
    Page,
    Date,
    Toc,
    Bookmark,
    Equation,
    Index,
    Unknown,
}

/// One token from a field instruction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldCodeToken<'a> {
    pub value: Cow<'a, str>,
    pub quoted: bool,
}

/// A preserved field-code switch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldSwitch<'a> {
    pub name: Cow<'a, str>,
    pub value: Option<Cow<'a, str>>,
}

/// A parsed HYPERLINK field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HyperlinkCode<'a> {
    pub external_target: Option<Cow<'a, str>>,
    pub bookmark: Option<Cow<'a, str>>,
    pub screen_tip: Option<Cow<'a, str>>,
    pub target_frame: Option<Cow<'a, str>>,
    pub coordinates: Option<Cow<'a, str>>,
    pub new_window: bool,
    pub unknown_switches: Vec<FieldSwitch<'a>>,
}

/// A parsed REF, PAGEREF, or NOTEREF field.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferenceCode<'a> {
    pub bookmark: Cow<'a, str>,
    pub hyperlink: bool,
    pub position: bool,
    pub footnote_mark: bool,
    pub unknown_switches: Vec<FieldSwitch<'a>>,
}

/// Why a recognized field code is non-actionable.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FieldCodeError {
    InstructionTooLong,
    TooManyTokens,
    UnterminatedQuote,
    MissingKeyword,
    MissingOperand(&'static str),
    DuplicateOperand(&'static str),
    UnexpectedOperand(String),
}

/// Typed field semantics. Malformed input is represented, never activated.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ParsedFieldCode<'a> {
    Hyperlink(HyperlinkCode<'a>),
    Reference(ReferenceCode<'a>),
    PageReference(ReferenceCode<'a>),
    NoteReference(ReferenceCode<'a>),
    Other {
        keyword: Cow<'a, str>,
        arguments: Vec<FieldCodeToken<'a>>,
    },
    Malformed(FieldCodeError),
}

/// Parsed RTF field.
#[derive(Debug, Clone)]
pub struct Field<'a> {
    pub field_type: FieldType,
    pub instruction: Cow<'a, str>,
    pub result: Cow<'a, str>,
}

impl<'a> Field<'a> {
    #[inline]
    pub fn new(field_type: FieldType, instruction: Cow<'a, str>, result: Cow<'a, str>) -> Self {
        Self {
            field_type,
            instruction,
            result,
        }
    }

    /// Parse the instruction keyword with an exact, case-insensitive boundary.
    pub fn parse_instruction(instruction: &'a str) -> Self {
        let parsed = parse_field_code(instruction);
        let field_type = match parsed {
            ParsedFieldCode::Hyperlink(_) => FieldType::Hyperlink,
            ParsedFieldCode::Reference(_) => FieldType::Reference,
            ParsedFieldCode::PageReference(_) => FieldType::PageReference,
            ParsedFieldCode::NoteReference(_) => FieldType::NoteReference,
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("PAGE") => {
                FieldType::Page
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("DATE")
                    || keyword.eq_ignore_ascii_case("TIME") =>
            {
                FieldType::Date
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("TOC") => {
                FieldType::Toc
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("BOOKMARK") =>
            {
                FieldType::Bookmark
            },
            ParsedFieldCode::Other { ref keyword, .. } if keyword.eq_ignore_ascii_case("EQ") => {
                FieldType::Equation
            },
            ParsedFieldCode::Other { ref keyword, .. }
                if keyword.eq_ignore_ascii_case("INDEX") || keyword.eq_ignore_ascii_case("XE") =>
            {
                FieldType::Index
            },
            _ => FieldType::Unknown,
        };
        Self {
            field_type,
            instruction: Cow::Borrowed(instruction),
            result: Cow::Borrowed(""),
        }
    }

    /// Parse this field's instruction into bounded, typed semantics.
    pub fn parsed_code(&self) -> ParsedFieldCode<'_> {
        parse_field_code(self.instruction.as_ref())
    }

    /// Compatibility URL helper. Internal-only links return `#bookmark`.
    pub fn extract_url(&self) -> Option<String> {
        let ParsedFieldCode::Hyperlink(code) = self.parsed_code() else {
            return None;
        };
        code.external_target
            .map(Cow::into_owned)
            .or_else(|| code.bookmark.map(|bookmark| format!("#{bookmark}")))
    }

    /// Compatibility bookmark helper for reference and hyperlink fields.
    pub fn extract_bookmark(&self) -> Option<String> {
        match self.parsed_code() {
            ParsedFieldCode::Hyperlink(code) => code.bookmark.map(Cow::into_owned),
            ParsedFieldCode::Reference(code)
            | ParsedFieldCode::PageReference(code)
            | ParsedFieldCode::NoteReference(code) => Some(code.bookmark.into_owned()),
            _ => None,
        }
    }

    #[inline]
    pub fn display_text(&self) -> &str {
        if !self.result.is_empty() {
            &self.result
        } else {
            &self.instruction
        }
    }
}

/// Parse a field instruction without evaluating it.
pub fn parse_field_code(instruction: &str) -> ParsedFieldCode<'_> {
    match parse_field_code_inner(instruction) {
        Ok(parsed) => parsed,
        Err(error) => ParsedFieldCode::Malformed(error),
    }
}

fn parse_field_code_inner(instruction: &str) -> Result<ParsedFieldCode<'_>, FieldCodeError> {
    let mut tokens = tokenize(instruction)?;
    if tokens.is_empty() {
        return Err(FieldCodeError::MissingKeyword);
    }
    let keyword = tokens.remove(0);
    if keyword.value.eq_ignore_ascii_case("HYPERLINK") {
        return parse_hyperlink(tokens).map(ParsedFieldCode::Hyperlink);
    }
    for (name, constructor) in [
        ("REF", 0u8),
        ("PAGEREF", 1u8),
        ("NOTEREF", 2u8),
    ] {
        if keyword.value.eq_ignore_ascii_case(name) {
            let code = parse_reference(tokens)?;
            return Ok(match constructor {
                0 => ParsedFieldCode::Reference(code),
                1 => ParsedFieldCode::PageReference(code),
                _ => ParsedFieldCode::NoteReference(code),
            });
        }
    }
    Ok(ParsedFieldCode::Other {
        keyword: keyword.value,
        arguments: tokens,
    })
}

fn parse_hyperlink(tokens: Vec<FieldCodeToken<'_>>) -> Result<HyperlinkCode<'_>, FieldCodeError> {
    let mut code = HyperlinkCode {
        external_target: None,
        bookmark: None,
        screen_tip: None,
        target_frame: None,
        coordinates: None,
        new_window: false,
        unknown_switches: Vec::new(),
    };
    let mut index = 0;
    while index < tokens.len() {
        let token = &tokens[index];
        if let Some(name) = switch_name(token) {
            let normalized = name.to_ascii_lowercase();
            match normalized.as_str() {
                "n" => {
                    if code.new_window {
                        return Err(FieldCodeError::DuplicateOperand("\\n"));
                    }
                    code.new_window = true;
                    index += 1;
                },
                "l" | "o" | "t" | "m" => {
                    let value = switch_value(&tokens, index, name)?;
                    let slot = match normalized.as_str() {
                        "l" => &mut code.bookmark,
                        "o" => &mut code.screen_tip,
                        "t" => &mut code.target_frame,
                        _ => &mut code.coordinates,
                    };
                    if slot.replace(value).is_some() {
                        return Err(FieldCodeError::DuplicateOperand(match normalized.as_str() {
                            "l" => "\\l",
                            "o" => "\\o",
                            "t" => "\\t",
                            _ => "\\m",
                        }));
                    }
                    index += 2;
                },
                _ => {
                    let value = tokens.get(index + 1).filter(|next| switch_name(next).is_none());
                    code.unknown_switches.push(FieldSwitch {
                        name: Cow::Owned(name.to_string()),
                        value: value.map(|token| token.value.clone()),
                    });
                    index += 1 + usize::from(value.is_some());
                },
            }
        } else {
            if code.external_target.replace(token.value.clone()).is_some() {
                return Err(FieldCodeError::UnexpectedOperand(token.value.to_string()));
            }
            index += 1;
        }
    }
    if code.external_target.is_none() && code.bookmark.is_none() {
        return Err(FieldCodeError::MissingOperand("hyperlink target or \\l bookmark"));
    }
    Ok(code)
}

fn parse_reference(tokens: Vec<FieldCodeToken<'_>>) -> Result<ReferenceCode<'_>, FieldCodeError> {
    let Some(first) = tokens.first() else {
        return Err(FieldCodeError::MissingOperand("bookmark"));
    };
    if switch_name(first).is_some() {
        return Err(FieldCodeError::MissingOperand("bookmark"));
    }
    let mut code = ReferenceCode {
        bookmark: first.value.clone(),
        hyperlink: false,
        position: false,
        footnote_mark: false,
        unknown_switches: Vec::new(),
    };
    let mut index = 1;
    while index < tokens.len() {
        let token = &tokens[index];
        let Some(name) = switch_name(token) else {
            return Err(FieldCodeError::UnexpectedOperand(token.value.to_string()));
        };
        match name.to_ascii_lowercase().as_str() {
            "h" if !code.hyperlink => code.hyperlink = true,
            "p" if !code.position => code.position = true,
            "f" if !code.footnote_mark => code.footnote_mark = true,
            "h" => return Err(FieldCodeError::DuplicateOperand("\\h")),
            "p" => return Err(FieldCodeError::DuplicateOperand("\\p")),
            "f" => return Err(FieldCodeError::DuplicateOperand("\\f")),
            _ => {
                let value = tokens.get(index + 1).filter(|next| switch_name(next).is_none());
                code.unknown_switches.push(FieldSwitch {
                    name: Cow::Owned(name.to_string()),
                    value: value.map(|token| token.value.clone()),
                });
                if value.is_some() {
                    index += 1;
                }
            },
        }
        index += 1;
    }
    Ok(code)
}

fn switch_name<'a>(token: &'a FieldCodeToken<'_>) -> Option<&'a str> {
    token
        .value
        .strip_prefix('\\')
        .filter(|name| !name.is_empty())
}

fn switch_value<'a>(
    tokens: &[FieldCodeToken<'a>],
    index: usize,
    name: &str,
) -> Result<Cow<'a, str>, FieldCodeError> {
    let value = tokens
        .get(index + 1)
        .filter(|value| switch_name(value).is_none())
        .ok_or(FieldCodeError::MissingOperand("switch value"))?;
    if name.is_empty() {
        return Err(FieldCodeError::MissingOperand("switch name"));
    }
    Ok(value.value.clone())
}

fn tokenize(instruction: &str) -> Result<Vec<FieldCodeToken<'_>>, FieldCodeError> {
    if instruction.len() > MAX_INSTRUCTION_LEN {
        return Err(FieldCodeError::InstructionTooLong);
    }
    let bytes = instruction.as_bytes();
    let mut tokens = Vec::new();
    let mut index = 0;
    while index < bytes.len() {
        while index < bytes.len() && bytes[index].is_ascii_whitespace() {
            index += 1;
        }
        if index == bytes.len() {
            break;
        }
        if tokens.len() >= MAX_TOKENS {
            return Err(FieldCodeError::TooManyTokens);
        }
        if bytes[index] == b'"' {
            index += 1;
            let mut value = String::new();
            let mut closed = false;
            while index < bytes.len() {
                match bytes[index] {
                    b'"' => {
                        index += 1;
                        closed = true;
                        break;
                    },
                    b'\\' if index + 1 < bytes.len()
                        && matches!(bytes[index + 1], b'\\' | b'"') =>
                    {
                        value.push(bytes[index + 1] as char);
                        index += 2;
                    },
                    _ => {
                        let character = instruction[index..]
                            .chars()
                            .next()
                            .expect("index is inside instruction");
                        value.push(character);
                        index += character.len_utf8();
                    },
                }
            }
            if !closed {
                return Err(FieldCodeError::UnterminatedQuote);
            }
            tokens.push(FieldCodeToken {
                value: Cow::Owned(value),
                quoted: true,
            });
        } else {
            let start = index;
            while index < bytes.len() && !bytes[index].is_ascii_whitespace() {
                index += 1;
            }
            tokens.push(FieldCodeToken {
                value: Cow::Borrowed(&instruction[start..index]),
                quoted: false,
            });
        }
    }
    Ok(tokens)
}

pub(crate) fn quoted_field_operand(value: &str) -> String {
    let mut output = String::with_capacity(value.len() + 2);
    output.push('"');
    for character in value.chars() {
        match character {
            '\\' => output.push_str("\\\\"),
            '"' => output.push_str("\\\""),
            _ => output.push(character),
        }
    }
    output.push('"');
    output
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn exact_case_insensitive_keywords_and_distinct_references() {
        assert!(matches!(parse_field_code("hyperlink \"https://e\""), ParsedFieldCode::Hyperlink(_)));
        for invalid in ["HYPERLINKER x", "REFRESH x", "PAGEREFERENCE x"] {
            assert!(matches!(parse_field_code(invalid), ParsedFieldCode::Other { .. }));
            assert_eq!(Field::parse_instruction(invalid).field_type, FieldType::Unknown);
        }
        assert!(matches!(parse_field_code("REF mark \\h"), ParsedFieldCode::Reference(_)));
        assert!(matches!(parse_field_code("PAGEREF mark \\p"), ParsedFieldCode::PageReference(_)));
        assert!(matches!(parse_field_code("NOTEREF mark \\f"), ParsedFieldCode::NoteReference(_)));
    }

    #[test]
    fn parses_internal_external_and_switch_semantics() {
        let ParsedFieldCode::Hyperlink(code) = parse_field_code(
            r#"HYPERLINK "https://example/a b" \l "_Toc1" \o "Tip" \t "_blank" \n"#,
        ) else {
            panic!("expected hyperlink");
        };
        assert_eq!(code.external_target.as_deref(), Some("https://example/a b"));
        assert_eq!(code.bookmark.as_deref(), Some("_Toc1"));
        assert_eq!(code.screen_tip.as_deref(), Some("Tip"));
        assert_eq!(code.target_frame.as_deref(), Some("_blank"));
        assert!(code.new_window);
        let field = Field::parse_instruction(r#"HYPERLINK \l "_Toc1""#);
        assert_eq!(field.extract_url().as_deref(), Some("#_Toc1"));
        assert_eq!(field.extract_bookmark().as_deref(), Some("_Toc1"));
    }

    #[test]
    fn writer_operand_cannot_inject_switches_and_round_trips_specials() {
        let target = "c:\\docs\\a \" \\l \"attacker{one}";
        let instruction = format!("HYPERLINK {}", quoted_field_operand(target));
        let ParsedFieldCode::Hyperlink(code) = parse_field_code(&instruction) else {
            panic!("expected hyperlink");
        };
        assert_eq!(code.external_target.as_deref(), Some(target));
        assert!(code.bookmark.is_none());

        let mut rtf = br#"{\rtf1\ansi "#.to_vec();
        crate::RtfWriter::new(&mut rtf)
            .write_hyperlink(target, "safe link")
            .unwrap();
        rtf.push(b'}');
        let document = crate::RtfDocument::from_bytes(&rtf).unwrap();
        let ParsedFieldCode::Hyperlink(code) = document.fields()[0].parsed_code() else {
            panic!("expected serialized hyperlink");
        };
        assert_eq!(code.external_target.as_deref(), Some(target));
        assert!(code.bookmark.is_none());
    }

    #[test]
    fn malformed_recognized_fields_are_non_actionable() {
        for instruction in [
            "HYPERLINK",
            r#"HYPERLINK "unterminated"#,
            r#"HYPERLINK \l"#,
            r#"HYPERLINK x \l a \l b"#,
            "REF",
            r#"REF a \h \h"#,
        ] {
            assert!(matches!(parse_field_code(instruction), ParsedFieldCode::Malformed(_)));
        }
    }

    #[test]
    fn parses_libreoffice_internal_hyperlink_fixtures() {
        for (relative, expected) in [
            ("rtfexport/data/fdo86750.rtf", "anchor"),
            ("rtfexport/data/tdf134614_toc_indent.rtf", "_Toc1"),
        ] {
            let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../3rdparty/libreoffice-core/sw/qa/extras")
                .join(relative);
            let document = crate::RtfDocument::from_bytes(&std::fs::read(path).unwrap()).unwrap();
            assert!(document.fields().iter().any(|field| {
                field.extract_bookmark().as_deref() == Some(expected)
                    && field.extract_url().as_deref() == Some(format!("#{expected}").as_str())
            }), "fixture {relative} fields: {:?}", document.fields());
        }

        let base = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../3rdparty/libreoffice-core/sw/qa/extras");
        let formatted = crate::RtfDocument::from_bytes(
            &std::fs::read(base.join("rtfimport/data/fdo82071.rtf")).unwrap(),
        )
        .unwrap();
        assert!(formatted.fields().iter().any(|field| matches!(
            field.parsed_code(),
            ParsedFieldCode::PageReference(ReferenceCode { ref bookmark, hyperlink: true, .. })
                if bookmark == "_Toc363816075"
        )));

        let backslashes = crate::RtfDocument::from_bytes(
            &std::fs::read(base.join("rtfexport/data/hyperlink-with-backslashes.rtf")).unwrap(),
        )
        .unwrap();
        assert_eq!(
            backslashes.fields()[0].extract_url().as_deref(),
            Some(r"c:\temp\doc1.doc")
        );

        let target = crate::RtfDocument::from_bytes(
            &std::fs::read(base.join("rtfexport/data/hyperlink-target.rtf")).unwrap(),
        )
        .unwrap();
        let ParsedFieldCode::Hyperlink(code) = target.fields()[0].parsed_code() else {
            panic!("expected target-frame hyperlink");
        };
        assert_eq!(code.target_frame.as_deref(), Some("_blank"));
    }
}
