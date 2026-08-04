//! Bounded Word field-instruction and document XML codecs.

#[allow(
    clippy::wildcard_imports,
    reason = "the codec shares the field model vocabulary with its parent"
)]
use super::*;

use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};

impl Field {
    /// Extract all fields from document XML bytes.
    ///
    /// # Arguments
    ///
    /// * `doc_xml` - The document XML bytes
    ///
    /// # Returns
    ///
    /// A vector of fields
    pub fn extract_from_document(doc_xml: &[u8]) -> Result<Vec<Field>> {
        let mut reader = Reader::from_reader(doc_xml);
        reader.config_mut().trim_text(false);

        let mut fields = Vec::new();
        let mut next_order = 0usize;
        let mut in_instr_text = false;
        let mut in_field_result = false;
        let mut in_result_text = false;
        let mut in_simple_result_text = false;
        let mut current_instruction = String::new();
        let mut current_result = String::new();
        let mut current_dirty = false;
        let mut current_locked = false;
        let mut current_order = 0usize;
        let mut field_depth: i32 = 0;
        let mut simple_fields = Vec::new();

        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) if e.local_name().as_ref() == b"t" => {
                    in_result_text = false;
                    in_simple_result_text = false;
                },
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"fldSimple" => {
                    simple_fields.push(PendingSimpleField::parse(
                        &e,
                        reader.decoder(),
                        next_order,
                    )?);
                    next_order += 1;
                    in_simple_result_text = false;
                },
                Ok(Event::Empty(e)) if e.local_name().as_ref() == b"fldSimple" => {
                    let field = PendingSimpleField::parse(&e, reader.decoder(), next_order)?;
                    next_order += 1;
                    fields.push((field.order, field.finish()));
                },
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    match e.local_name().as_ref() {
                        b"fldChar" => {
                            // Field character marks field boundaries
                            let mut fld_char_type = None;
                            let mut dirty = None;
                            let mut locked = None;

                            for attr in e.attributes() {
                                let attr = attr
                                    .map_err(|error| Error::Xml(error.to_string()))?;
                                let value = attr
                                    .decoded_and_normalized_value(
                                        XmlVersion::Explicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(|error| Error::Xml(error.to_string()))?;
                                if attr.key.local_name().as_ref() == b"fldCharType" {
                                    fld_char_type = Some(value.to_string());
                                }
                                if attr.key.local_name().as_ref() == b"dirty" {
                                    dirty = Some(is_on(&value));
                                }
                                if attr.key.local_name().as_ref() == b"fldLock" {
                                    locked = Some(is_on(&value));
                                }
                            }

                            if let Some(ref char_type) = fld_char_type {
                                match char_type.as_str() {
                                    "begin" => {
                                        // Start of field
                                        field_depth += 1;
                                        if field_depth == 1 {
                                            current_order = next_order;
                                            next_order += 1;
                                            current_instruction.clear();
                                            current_result.clear();
                                            current_dirty = dirty.unwrap_or(false);
                                            current_locked = locked.unwrap_or(false);
                                            in_instr_text = false;
                                            in_field_result = false;
                                            in_result_text = false;
                                        }
                                    },
                                    "separate"
                                        // Separator between instruction and result
                                        if field_depth == 1 => {
                                            current_dirty |= dirty.unwrap_or(false);
                                            current_locked |= locked.unwrap_or(false);
                                            in_instr_text = false;
                                            in_field_result = true;
                                            in_result_text = false;
                                        },
                                    "end" => {
                                        // End of field
                                        if field_depth == 1 {
                                            in_field_result = false;
                                            in_instr_text = false;
                                            in_result_text = false;

                                            if !current_instruction.is_empty() {
                                                let result = if current_result.is_empty() {
                                                    None
                                                } else {
                                                    Some(current_result.clone())
                                                };
                                                fields.push((current_order, Field::with_flags(
                                                    current_instruction.trim().to_string(),
                                                    result,
                                                    current_dirty,
                                                    current_locked,
                                                )));
                                            }
                                        }
                                        field_depth = field_depth.saturating_sub(1);
                                    },
                                    _ => {},
                                }
                            }
                        },
                        b"instrText"
                            // Field instruction text
                            if field_depth > 0 => {
                                in_instr_text = true;
                            },
                        b"t" => {
                            if in_field_result {
                                in_result_text = true;
                            }
                            if !simple_fields.is_empty() {
                                in_simple_result_text = true;
                            }
                        },
                        b"tab" if in_field_result && field_depth == 1 => {
                            current_result.push('\t');
                        },
                        b"br" | b"cr" if in_field_result && field_depth == 1 => {
                            current_result.push('\n');
                        },
                        b"noBreakHyphen" if in_field_result && field_depth == 1 => {
                            current_result.push('\u{2011}');
                        },
                        b"softHyphen" if in_field_result && field_depth == 1 => {
                            current_result.push('\u{00ad}');
                        },
                        _ => {},
                    }

                    if !simple_fields.is_empty() {
                        let character = match e.local_name().as_ref() {
                            b"tab" => Some('\t'),
                            b"br" | b"cr" => Some('\n'),
                            b"noBreakHyphen" => Some('\u{2011}'),
                            b"softHyphen" => Some('\u{00ad}'),
                            _ => None,
                        };
                        if let Some(character) = character {
                            for field in &mut simple_fields {
                                field.result.push(character);
                            }
                        }
                    }
                },
                Ok(Event::Text(e)) => {
                    let has_complex_target = (in_instr_text && field_depth == 1)
                        || (in_field_result && in_result_text && field_depth == 1);
                    if has_complex_target || in_simple_result_text {
                        let decoded = e
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        let unescaped = quick_xml::escape::unescape(&decoded)
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        if in_instr_text && field_depth == 1 {
                            current_instruction.push_str(&unescaped);
                        } else if in_field_result && in_result_text && field_depth == 1 {
                            current_result.push_str(&unescaped);
                        }
                        if in_simple_result_text {
                            for field in &mut simple_fields {
                                field.result.push_str(&unescaped);
                            }
                        }
                    }
                },
                Ok(Event::CData(e)) => {
                    let has_complex_target = (in_instr_text && field_depth == 1)
                        || (in_field_result && in_result_text && field_depth == 1);
                    if has_complex_target || in_simple_result_text {
                        let decoded = e
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        if in_instr_text && field_depth == 1 {
                            current_instruction.push_str(&decoded);
                        } else if in_field_result && in_result_text && field_depth == 1 {
                            current_result.push_str(&decoded);
                        }
                        if in_simple_result_text {
                            for field in &mut simple_fields {
                                field.result.push_str(&decoded);
                            }
                        }
                    }
                },
                Ok(Event::GeneralRef(reference)) => {
                    let has_complex_target = (in_instr_text && field_depth == 1)
                        || (in_field_result && in_result_text && field_depth == 1);
                    if has_complex_target || in_simple_result_text {
                        let decoded = decode_xml_reference(&reference)
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        if in_instr_text && field_depth == 1 {
                            current_instruction.push_str(&decoded);
                        } else if in_field_result && in_result_text && field_depth == 1 {
                            current_result.push_str(&decoded);
                        }
                        if in_simple_result_text {
                            for field in &mut simple_fields {
                                field.result.push_str(&decoded);
                            }
                        }
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"instrText" => {
                    in_instr_text = false;
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"t" => {
                    in_result_text = false;
                    in_simple_result_text = false;
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"fldSimple" => {
                    in_simple_result_text = false;
                    let field = simple_fields.pop().ok_or_else(|| {
                        Error::Invalid(
                            "DOCX simple field ended without a matching start".to_string(),
                        )
                    })?;
                    fields.push((field.order, field.finish()));
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(Error::Xml(e.to_string())),
                _ => {},
            }
        }

        fields.sort_unstable_by_key(|(order, _)| *order);
        Ok(fields.into_iter().map(|(_, field)| field).collect())
    }
}
pub(super) fn has_field_switch(switches: &[Switch], name: char) -> bool {
    switches
        .iter()
        .any(|switch| switch.name.eq_ignore_ascii_case(&name))
}

pub(super) fn optional_field_switch_argument<'a>(
    switches: &'a [Switch],
    name: char,
    field_type: &str,
) -> Result<Option<&'a str>> {
    let mut matching = switches
        .iter()
        .filter(|switch| switch.name.eq_ignore_ascii_case(&name));
    let Some(switch) = matching.next() else {
        return Ok(None);
    };
    if matching.next().is_some() {
        return Err(Error::Invalid(format!(
            "{field_type} field has duplicate \\{name} switches"
        )));
    }
    switch
        .argument
        .as_deref()
        .map(Some)
        .ok_or_else(|| Error::Invalid(format!("{field_type} \\{name} switch requires an argument")))
}

pub(super) fn parse_authority_category(value: &str, minimum: u8, field_type: &str) -> Result<u8> {
    let value = value.parse::<u8>().map_err(|_| {
        Error::Invalid(format!("{field_type} authority category is not an integer"))
    })?;
    if !(minimum..=16).contains(&value) {
        return Err(Error::Invalid(format!(
            "{field_type} authority category must be in {minimum}..=16"
        )));
    }
    Ok(value)
}

pub(super) fn parse_index_columns(value: &str) -> Result<u8> {
    let columns = value
        .parse::<u8>()
        .map_err(|_| Error::Invalid("INDEX column count is not an integer".to_string()))?;
    if !(1..=4).contains(&columns) {
        return Err(Error::Invalid(
            "INDEX column count must be in 1..=4".to_string(),
        ));
    }
    Ok(columns)
}

pub(super) fn parse_index_sort_order(value: &str) -> Result<IndexOrder> {
    match value {
        "S" | "s" => Ok(IndexOrder::Stroke),
        "P" | "p" => Ok(IndexOrder::Pronunciation),
        _ => Err(Error::Invalid(format!(
            "INDEX \\o sort order must be S or P, got {value:?}"
        ))),
    }
}

pub(super) fn field_instruction_remainder<'a>(
    instruction: &'a str,
    field_type: &str,
) -> Option<&'a str> {
    let instruction = instruction.trim_start();
    let field_type_end = field_type.len();
    let candidate = instruction.get(..field_type_end)?;
    let remainder = instruction.get(field_type_end..)?;
    if !candidate.eq_ignore_ascii_case(field_type) {
        return None;
    }
    match remainder.chars().next() {
        None | Some('\\') | Some('"') => Some(remainder),
        Some(character) if character.is_whitespace() => Some(remainder),
        Some(_) => None,
    }
}

pub(super) fn parse_field_switches(
    instruction: &str,
    field_type: &str,
) -> Result<Option<Vec<Switch>>> {
    let Some(remainder) = field_instruction_remainder(instruction, field_type) else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    Ok(Some(parse_field_switches_from_characters(
        &mut characters,
        field_type,
    )?))
}

pub(super) fn parse_field_operand_and_switches(
    instruction: &str,
    field_type: &str,
) -> Result<Option<(Option<String>, Vec<Switch>)>> {
    let Some(remainder) = field_instruction_remainder(instruction, field_type) else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    skip_field_whitespace(&mut characters);
    let operand = match characters.peek().copied() {
        None | Some('\\') => None,
        Some('"') => {
            characters.next();
            Some(parse_field_quoted_argument(&mut characters, field_type)?)
        },
        Some(_) => Some(parse_field_unquoted_argument(&mut characters)),
    };
    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    Ok(Some((operand, switches)))
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_mail_merge_data_field_parts(
    instruction: &str,
) -> Result<Option<(String, Option<String>, Vec<Switch>)>> {
    let Some(remainder) = field_instruction_remainder(instruction, "DATA") else {
        return Ok(None);
    };
    if instruction.len() > MAX_MAIL_MERGE_DATA_FIELD_INSTRUCTION_BYTES {
        return Err(Error::Invalid(format!(
            "DATA field instruction exceeds {MAX_MAIL_MERGE_DATA_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let mut characters = remainder.chars().peekable();
    let data_source = parse_next_field_argument(&mut characters, "DATA")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            Error::Invalid("DATA field is missing its data-source identifier".to_string())
        })?;
    let header_source = match parse_next_field_argument(&mut characters, "DATA")? {
        Some(header_source) if !header_source.is_empty() => Some(header_source),
        Some(_) => {
            return Err(Error::Invalid(
                "DATA field header-source identifier is empty".to_string(),
            ));
        },
        None => None,
    };
    let switches = parse_field_switches_from_characters(&mut characters, "DATA")?;

    Ok(Some((data_source, header_source, switches)))
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_info_field_parts(
    instruction: &str,
) -> Result<Option<(String, Option<String>, Vec<Switch>)>> {
    let Some(remainder) = field_instruction_remainder(instruction, "INFO") else {
        return Ok(None);
    };
    if instruction.len() > MAX_INFO_FIELD_INSTRUCTION_BYTES {
        return Err(Error::Invalid(format!(
            "INFO field instruction exceeds {MAX_INFO_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let mut characters = remainder.chars().peekable();
    let information_type = parse_next_field_argument(&mut characters, "INFO")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid("INFO field is missing its property selector".to_string()))?;
    let new_value = parse_next_field_argument(&mut characters, "INFO")?;
    let switches = parse_field_switches_from_characters(&mut characters, "INFO")?;

    Ok(Some((information_type, new_value, switches)))
}

pub(super) fn parse_macro_button_operands(instruction: &str) -> Result<Option<(String, String)>> {
    let Some(remainder) = field_instruction_remainder(instruction, "MACROBUTTON") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let macro_name = parse_next_field_argument(&mut characters, "MACROBUTTON")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            Error::Invalid("MACROBUTTON field is missing its macro or command name".to_string())
        })?;
    let display_text = parse_next_field_argument(&mut characters, "MACROBUTTON")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            Error::Invalid("MACROBUTTON field is missing its button text".to_string())
        })?;
    skip_field_whitespace(&mut characters);
    if characters.next().is_some() {
        return Err(Error::Invalid(
            "MACROBUTTON field must contain exactly two arguments and no switches".to_string(),
        ));
    }
    Ok(Some((macro_name, display_text)))
}

pub(super) fn parse_go_to_button_operands(instruction: &str) -> Result<Option<(String, String)>> {
    let Some(remainder) = field_instruction_remainder(instruction, "GOTOBUTTON") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let target = parse_next_field_argument(&mut characters, "GOTOBUTTON")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid("GOTOBUTTON field is missing its destination".to_string()))?;
    let button_text = parse_next_field_argument(&mut characters, "GOTOBUTTON")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid("GOTOBUTTON field is missing its button text".to_string()))?;
    skip_field_whitespace(&mut characters);
    if characters.next().is_some() {
        return Err(Error::Invalid(
            "GOTOBUTTON field must contain exactly two arguments and no switches".to_string(),
        ));
    }
    Ok(Some((target, button_text)))
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_user_identity_field_parts(
    instruction: &str,
) -> Result<Option<(UserIdentityKind, Option<String>, Option<UserIdentityFormat>)>> {
    let (kind, field_type, remainder) =
        if let Some(remainder) = field_instruction_remainder(instruction, "USERADDRESS") {
            (UserIdentityKind::Address, "USERADDRESS", remainder)
        } else if let Some(remainder) = field_instruction_remainder(instruction, "USERINITIALS") {
            (UserIdentityKind::Initials, "USERINITIALS", remainder)
        } else if let Some(remainder) = field_instruction_remainder(instruction, "USERNAME") {
            (UserIdentityKind::Name, "USERNAME", remainder)
        } else {
            return Ok(None);
        };

    let mut characters = remainder.chars().peekable();
    let override_value = parse_next_field_argument(&mut characters, field_type)?;
    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    let mut formatting = None;
    for switch in switches {
        if switch.name != '*' {
            return Err(Error::Invalid(format!(
                "{field_type} field has an unsupported \\{} switch",
                switch.name
            )));
        }
        if formatting.is_some() {
            return Err(Error::Invalid(format!(
                "{field_type} field repeats its \\* switch"
            )));
        }
        let argument = switch.argument.ok_or_else(|| {
            Error::Invalid(format!(
                "{field_type} \\* switch requires a general-formatting argument"
            ))
        })?;
        formatting = Some(if argument.eq_ignore_ascii_case("Caps") {
            UserIdentityFormat::Caps
        } else if argument.eq_ignore_ascii_case("FirstCap") {
            UserIdentityFormat::FirstCap
        } else if argument.eq_ignore_ascii_case("Lower") {
            UserIdentityFormat::Lower
        } else if argument.eq_ignore_ascii_case("Upper") {
            UserIdentityFormat::Upper
        } else {
            return Err(Error::Invalid(format!(
                "{field_type} \\* switch must be Caps, FirstCap, Lower, or Upper"
            )));
        });
    }

    Ok(Some((kind, override_value, formatting)))
}

pub(super) fn parse_advance_field_adjustments(
    instruction: &str,
) -> Result<Option<Vec<AdvanceAdjustment>>> {
    let Some(switches) = parse_field_switches(instruction, "ADVANCE")? else {
        return Ok(None);
    };

    let mut adjustments = Vec::with_capacity(switches.len());
    for switch in switches {
        let operation = match switch.name {
            'd' => AdvanceOperation::Down,
            'l' => AdvanceOperation::Left,
            'r' => AdvanceOperation::Right,
            'u' => AdvanceOperation::Up,
            'x' => AdvanceOperation::HorizontalPosition,
            'y' => AdvanceOperation::VerticalPosition,
            name => {
                return Err(Error::Invalid(format!(
                    "ADVANCE field has an unsupported \\{name} switch"
                )));
            },
        };
        let points = switch.argument.ok_or_else(|| {
            Error::Invalid(format!(
                "ADVANCE \\{} switch requires an integral number of points",
                switch.name
            ))
        })?;
        let points = points.parse::<i64>().map_err(|_| {
            Error::Invalid(format!(
                "ADVANCE \\{} switch must specify an integral number of points",
                switch.name
            ))
        })?;
        adjustments.push(AdvanceAdjustment { operation, points });
    }

    Ok(Some(adjustments))
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_link_operands_and_switches(
    instruction: &str,
) -> Result<Option<(String, String, Option<String>, Vec<Switch>)>> {
    parse_external_link_operands_and_switches(instruction, "LINK")
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_dde_operands_and_switches(
    instruction: &str,
) -> Result<Option<(DdeKind, String, String, Option<String>, Vec<Switch>)>> {
    if let Some((application, source, item, switches)) =
        parse_external_link_operands_and_switches(instruction, "DDEAUTO")?
    {
        return Ok(Some((
            DdeKind::DdeAuto,
            application,
            source,
            item,
            switches,
        )));
    }

    Ok(
        parse_external_link_operands_and_switches(instruction, "DDE")?.map(
            |(application, source, item, switches)| {
                (DdeKind::Dde, application, source, item, switches)
            },
        ),
    )
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_external_include_operands_and_switches(
    instruction: &str,
) -> Result<Option<(IncludeKind, String, Option<String>, Vec<Switch>)>> {
    let (kind, field_type) = if field_instruction_remainder(instruction, "INCLUDETEXT").is_some() {
        (IncludeKind::Text, "INCLUDETEXT")
    } else if field_instruction_remainder(instruction, "INCLUDE").is_some() {
        (IncludeKind::Text, "INCLUDE")
    } else if field_instruction_remainder(instruction, "INCLUDEPICTURE").is_some() {
        (IncludeKind::Picture, "INCLUDEPICTURE")
    } else if field_instruction_remainder(instruction, "IMPORT").is_some() {
        (IncludeKind::Picture, "IMPORT")
    } else {
        return Ok(None);
    };
    let remainder =
        field_instruction_remainder(instruction, field_type).expect("recognized include field");
    let mut characters = remainder.chars().peekable();
    let source = parse_next_field_argument(&mut characters, field_type)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid(format!("{field_type} field is missing its source")))?;
    let bookmark = match kind {
        IncludeKind::Text => parse_next_field_argument(&mut characters, field_type)?,
        IncludeKind::Picture => None,
    };
    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    Ok(Some((kind, source, bookmark, switches)))
}

pub(super) fn required_external_include_option_argument(
    switch: &Switch,
    kind: IncludeKind,
) -> Result<String> {
    let field_type = match kind {
        IncludeKind::Text => "INCLUDETEXT",
        IncludeKind::Picture => "INCLUDEPICTURE",
    };
    switch.argument.clone().ok_or_else(|| {
        Error::Invalid(format!(
            "{field_type} {} switch requires an argument",
            switch.name
        ))
    })
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_external_link_operands_and_switches(
    instruction: &str,
    field_type: &str,
) -> Result<Option<(String, String, Option<String>, Vec<Switch>)>> {
    let Some(remainder) = field_instruction_remainder(instruction, field_type) else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let application_type = parse_next_field_argument(&mut characters, field_type)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| {
            Error::Invalid(format!(
                "{field_type} field is missing its application type"
            ))
        })?;
    let source = parse_next_field_argument(&mut characters, field_type)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid(format!("{field_type} field is missing its source")))?;
    let item = parse_next_field_argument(&mut characters, field_type)?;
    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    Ok(Some((application_type, source, item, switches)))
}

pub(super) fn parse_next_field_argument(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
    field_type: &str,
) -> Result<Option<String>> {
    skip_field_whitespace(characters);
    match characters.peek().copied() {
        None | Some('\\') => Ok(None),
        Some('"') => {
            characters.next();
            Ok(Some(parse_field_quoted_argument(characters, field_type)?))
        },
        Some(_) => Ok(Some(parse_field_unquoted_argument(characters))),
    }
}

pub(super) fn parse_set_field_parts(instruction: &str) -> Result<Option<(String, String)>> {
    if instruction.len() > MAX_SET_FIELD_INSTRUCTION_BYTES {
        return Err(Error::Invalid(format!(
            "SET field instruction exceeds {MAX_SET_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let Some(remainder) = field_instruction_remainder(instruction, "SET") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let target_name = parse_next_field_argument(&mut characters, "SET")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid("SET field is missing its target name".to_string()))?;
    skip_field_whitespace(&mut characters);
    let expression = characters.collect::<String>();
    if expression.trim().is_empty() {
        return Err(Error::Invalid(
            "SET field is missing its expression".to_string(),
        ));
    }

    Ok(Some((target_name, expression)))
}

pub(super) fn parse_sequence_field_parts(
    instruction: &str,
) -> Result<Option<(String, Option<String>, String)>> {
    if instruction.len() > MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES {
        return Err(Error::Invalid(format!(
            "SEQ field instruction exceeds {MAX_SEQUENCE_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let Some(remainder) = field_instruction_remainder(instruction, "SEQ") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let identifier = parse_next_field_argument(&mut characters, "SEQ")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid("SEQ field is missing its identifier".to_string()))?;
    skip_field_whitespace(&mut characters);
    let bookmark = match characters.peek().copied() {
        None | Some('\\') => None,
        Some(_) => Some(
            parse_next_field_argument(&mut characters, "SEQ")?
                .filter(|value| !value.is_empty())
                .ok_or_else(|| Error::Invalid("SEQ field bookmark is empty".to_string()))?,
        ),
    };
    skip_field_whitespace(&mut characters);
    let tail = characters.collect::<String>().trim().to_string();

    Ok(Some((identifier, bookmark, tail)))
}

pub(super) fn parse_formula_field_formula(instruction: &str) -> Result<Option<String>> {
    if instruction.len() > MAX_FORMULA_FIELD_INSTRUCTION_BYTES {
        return Err(Error::Invalid(format!(
            "formula field instruction exceeds {MAX_FORMULA_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let Some(formula) = instruction.trim().strip_prefix('=') else {
        return Ok(None);
    };
    let formula = formula.trim();
    if formula.is_empty() {
        return Err(Error::Invalid(
            "formula field is missing its formula".to_string(),
        ));
    }

    Ok(Some(formula.to_string()))
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_style_reference_field_parts(
    instruction: &str,
) -> Result<Option<(String, Vec<StyleOption>, Vec<Switch>)>> {
    if instruction.len() > MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES {
        return Err(Error::Invalid(format!(
            "STYLEREF field instruction exceeds {MAX_STYLE_REFERENCE_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let Some(remainder) = field_instruction_remainder(instruction, "STYLEREF") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let style_name = parse_next_field_argument(&mut characters, "STYLEREF")?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid("STYLEREF field is missing its style name".to_string()))?;
    let switches = parse_field_switches_from_characters(&mut characters, "STYLEREF")?;
    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    for switch in switches {
        let option = match switch.name {
            'l' => Some(StyleOption::FollowingText),
            'n' => Some(StyleOption::ParagraphNumber),
            'p' => Some(StyleOption::RelativePosition),
            'r' => Some(StyleOption::ParagraphNumberRelativeContext),
            't' => Some(StyleOption::SuppressNonNumberText),
            'w' => Some(StyleOption::ParagraphNumberFullContext),
            _ => None,
        };
        if let Some(option) = option {
            if switch.argument.is_some() {
                return Err(Error::Invalid(format!(
                    "STYLEREF \\\\{} switch does not take an argument",
                    switch.name
                )));
            }
            options.push(option);
        } else {
            unknown_switches.push(switch);
        }
    }

    Ok(Some((style_name, options, unknown_switches)))
}

pub(super) fn parse_auto_text_field_parts(
    instruction: &str,
) -> Result<Option<(AutoTextKind, String, Vec<Switch>)>> {
    let (kind, field_type, remainder) =
        if let Some(remainder) = field_instruction_remainder(instruction, "GLOSSARY") {
            (AutoTextKind::Glossary, "GLOSSARY", remainder)
        } else if let Some(remainder) = field_instruction_remainder(instruction, "AUTOTEXT") {
            (AutoTextKind::AutoText, "AUTOTEXT", remainder)
        } else {
            return Ok(None);
        };
    if instruction.len() > MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES {
        return Err(Error::Invalid(format!(
            "{field_type} field instruction exceeds {MAX_AUTO_TEXT_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let mut characters = remainder.chars().peekable();
    let entry_name = parse_next_field_argument(&mut characters, field_type)?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::Invalid(format!("{field_type} field is missing its entry name")))?;
    let unknown_switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    Ok(Some((kind, entry_name, unknown_switches)))
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_auto_text_list_field_parts(
    instruction: &str,
) -> Result<Option<(Option<String>, Vec<AutoTextListOption>, Vec<Switch>)>> {
    let Some(remainder) = field_instruction_remainder(instruction, "AUTOTEXTLIST") else {
        return Ok(None);
    };
    if instruction.len() > MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES {
        return Err(Error::Invalid(format!(
            "AUTOTEXTLIST field instruction exceeds {MAX_AUTO_TEXT_LIST_FIELD_INSTRUCTION_BYTES} bytes"
        )));
    }

    let mut characters = remainder.chars().peekable();
    skip_field_whitespace(&mut characters);
    let display_text = match characters.peek().copied() {
        None | Some('\\') => None,
        Some(_) => parse_next_field_argument(&mut characters, "AUTOTEXTLIST")?,
    };
    let switches = parse_field_switches_from_characters(&mut characters, "AUTOTEXTLIST")?;
    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    for switch in switches {
        match switch.name {
            's' => {
                let style = switch.argument.ok_or_else(|| {
                    Error::Invalid("AUTOTEXTLIST \\s switch requires an argument".to_string())
                })?;
                options.push(AutoTextListOption::Style(style));
            },
            't' => {
                let tip = switch.argument.ok_or_else(|| {
                    Error::Invalid("AUTOTEXTLIST \\t switch requires an argument".to_string())
                })?;
                options.push(AutoTextListOption::Tip(tip));
            },
            _ => unknown_switches.push(switch),
        }
    }
    Ok(Some((display_text, options, unknown_switches)))
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_prompt_field_parts(
    instruction: &str,
) -> Result<
    Option<(
        PromptKind,
        Option<String>,
        Option<String>,
        Option<String>,
        bool,
    )>,
> {
    let (kind, field_type, remainder) =
        if let Some(remainder) = field_instruction_remainder(instruction, "ASK") {
            (PromptKind::Ask, "ASK", remainder)
        } else if let Some(remainder) = field_instruction_remainder(instruction, "FILLIN") {
            (PromptKind::FillIn, "FILLIN", remainder)
        } else {
            return Ok(None);
        };

    let mut characters = remainder.chars().peekable();
    let (bookmark, prompt) = match kind {
        PromptKind::Ask => {
            let bookmark = parse_next_field_argument(&mut characters, field_type)?
                .filter(|value| !value.is_empty())
                .ok_or_else(|| {
                    Error::Invalid("ASK field is missing its bookmark name".to_string())
                })?;
            let prompt =
                parse_next_field_argument(&mut characters, field_type)?.ok_or_else(|| {
                    Error::Invalid("ASK field is missing its prompt text".to_string())
                })?;
            (Some(bookmark), Some(prompt))
        },
        PromptKind::FillIn => (
            None,
            parse_next_field_argument(&mut characters, field_type)?,
        ),
    };

    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    let mut default_response = None;
    let mut prompts_once_per_mail_merge = false;
    for switch in switches {
        match switch.name {
            'd' => {
                if default_response.is_some() {
                    return Err(Error::Invalid(format!(
                        "{field_type} field repeats its \\d switch"
                    )));
                }
                default_response = Some(switch.argument.ok_or_else(|| {
                    Error::Invalid(format!(
                        "{field_type} field requires an argument for its \\d switch"
                    ))
                })?);
            },
            'o' => {
                if prompts_once_per_mail_merge {
                    return Err(Error::Invalid(format!(
                        "{field_type} field repeats its \\o switch"
                    )));
                }
                if switch.argument.is_some() {
                    return Err(Error::Invalid(format!(
                        "{field_type} field does not allow an argument for its \\o switch"
                    )));
                }
                prompts_once_per_mail_merge = true;
            },
            _ => {
                return Err(Error::Invalid(format!(
                    "{field_type} field has an unsupported \\{} switch",
                    switch.name
                )));
            },
        }
    }

    Ok(Some((
        kind,
        bookmark,
        prompt,
        default_response,
        prompts_once_per_mail_merge,
    )))
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_mail_merge_recipient_field_parts(
    instruction: &str,
) -> Result<
    Option<(
        RecipientKind,
        Option<CountryInclusion>,
        bool,
        Vec<String>,
        Option<String>,
        Option<String>,
        Option<String>,
        Vec<Switch>,
    )>,
> {
    let (kind, field_type, remainder) =
        if let Some(remainder) = field_instruction_remainder(instruction, "ADDRESSBLOCK") {
            (RecipientKind::AddressBlock, "ADDRESSBLOCK", remainder)
        } else if let Some(remainder) = field_instruction_remainder(instruction, "GREETINGLINE") {
            (RecipientKind::GreetingLine, "GREETINGLINE", remainder)
        } else {
            return Ok(None);
        };

    let mut characters = remainder.chars().peekable();
    let switches = parse_field_switches_from_characters(&mut characters, field_type)?;
    let mut country_inclusion = None;
    let mut formats_using_recipient_country = false;
    let mut excluded_countries = Vec::new();
    let mut format_template = None;
    let mut language = None;
    let mut greeting_fallback_text = None;
    let mut unknown_switches = Vec::new();

    for switch in switches {
        match (kind, switch.name) {
            (RecipientKind::AddressBlock, 'c') => {
                if country_inclusion.is_some() {
                    return Err(Error::Invalid(
                        "ADDRESSBLOCK field repeats its \\c switch".to_string(),
                    ));
                }
                let argument = switch.argument.ok_or_else(|| {
                    Error::Invalid("ADDRESSBLOCK \\c switch requires an argument".to_string())
                })?;
                country_inclusion = Some(match argument.as_str() {
                    "0" => CountryInclusion::Omit,
                    "1" => CountryInclusion::Always,
                    "2" => CountryInclusion::UnlessExcluded,
                    _ => {
                        return Err(Error::Invalid(format!(
                            "ADDRESSBLOCK \\c switch must be 0, 1, or 2, got {argument:?}"
                        )));
                    },
                });
            },
            (RecipientKind::AddressBlock, 'd') => {
                if formats_using_recipient_country {
                    return Err(Error::Invalid(
                        "ADDRESSBLOCK field repeats its \\d switch".to_string(),
                    ));
                }
                if switch.argument.is_some() {
                    return Err(Error::Invalid(
                        "ADDRESSBLOCK \\d switch does not accept an argument".to_string(),
                    ));
                }
                formats_using_recipient_country = true;
            },
            (RecipientKind::AddressBlock, 'e') => {
                let argument = switch.argument.ok_or_else(|| {
                    Error::Invalid("ADDRESSBLOCK \\e switch requires an argument".to_string())
                })?;
                excluded_countries.push(argument);
            },
            (_, 'f') => {
                if format_template.is_some() {
                    return Err(Error::Invalid(format!(
                        "{field_type} field repeats its \\f switch"
                    )));
                }
                format_template = Some(switch.argument.ok_or_else(|| {
                    Error::Invalid(format!("{field_type} \\f switch requires an argument"))
                })?);
            },
            (_, 'l') => {
                if language.is_some() {
                    return Err(Error::Invalid(format!(
                        "{field_type} field repeats its \\l switch"
                    )));
                }
                language = Some(switch.argument.ok_or_else(|| {
                    Error::Invalid(format!("{field_type} \\l switch requires an argument"))
                })?);
            },
            (RecipientKind::GreetingLine, 'c' | 'e') => {
                if greeting_fallback_text.is_some() {
                    return Err(Error::Invalid(
                        "GREETINGLINE field repeats its fallback-text switch".to_string(),
                    ));
                }
                greeting_fallback_text = Some(switch.argument.ok_or_else(|| {
                    Error::Invalid(
                        "GREETINGLINE fallback-text switch requires an argument".to_string(),
                    )
                })?);
            },
            _ => unknown_switches.push(switch),
        }
    }

    Ok(Some((
        kind,
        country_inclusion,
        formats_using_recipient_country,
        excluded_countries,
        format_template,
        language,
        greeting_fallback_text,
        unknown_switches,
    )))
}

/// Parse a `CITATION` instruction while accepting Word's documented leading
/// `\\l` locale switch. Other switches still follow the primary source tag or
/// a preceding `\\m` source tag.
pub(super) fn parse_citation_operand_and_switches(
    instruction: &str,
) -> Result<Option<(Option<String>, Vec<Switch>)>> {
    let Some(remainder) = field_instruction_remainder(instruction, "CITATION") else {
        return Ok(None);
    };
    let mut characters = remainder.chars().peekable();
    let mut switches = Vec::new();
    skip_field_whitespace(&mut characters);
    while characters
        .peek()
        .is_some_and(|character| *character == '\\')
    {
        let switch = parse_field_switch_from_characters(&mut characters, "CITATION")?;
        if switch.name != 'l' {
            return Err(Error::Invalid(
                "CITATION field requires its primary source tag before this switch".to_string(),
            ));
        }
        if switches.len() >= MAX_FIELD_SWITCHES {
            return Err(Error::Invalid(format!(
                "CITATION field exceeds {MAX_FIELD_SWITCHES} switches"
            )));
        }
        switches.push(switch);
        skip_field_whitespace(&mut characters);
    }
    let operand = match characters.peek().copied() {
        None | Some('\\') => None,
        Some('"') => {
            characters.next();
            Some(parse_field_quoted_argument(&mut characters, "CITATION")?)
        },
        Some(_) => Some(parse_field_unquoted_argument(&mut characters)),
    };
    let remaining = parse_field_switches_from_characters(&mut characters, "CITATION")?;
    if switches.len() + remaining.len() > MAX_FIELD_SWITCHES {
        return Err(Error::Invalid(format!(
            "CITATION field exceeds {MAX_FIELD_SWITCHES} switches"
        )));
    }
    switches.extend(remaining);
    Ok(Some((operand, switches)))
}

pub(super) fn parse_field_switches_from_characters(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
    field_type: &str,
) -> Result<Vec<Switch>> {
    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(characters);
        let Some(character) = characters.next() else {
            break;
        };
        if character != '\\' {
            return Err(Error::Invalid(format!(
                "{field_type} field contains text outside a field switch"
            )));
        }
        if switches.len() >= MAX_FIELD_SWITCHES {
            return Err(Error::Invalid(format!(
                "{field_type} field exceeds {MAX_FIELD_SWITCHES} switches"
            )));
        }
        switches.push(parse_field_switch_after_intro(characters, field_type)?);
    }
    Ok(switches)
}

pub(super) fn parse_field_switch_from_characters(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
    field_type: &str,
) -> Result<Switch> {
    let introducer = characters.next().ok_or_else(|| {
        Error::Invalid(format!("{field_type} field ends with a switch introducer"))
    })?;
    if introducer != '\\' {
        return Err(Error::Invalid(format!(
            "{field_type} field has an invalid switch introducer"
        )));
    }
    parse_field_switch_after_intro(characters, field_type)
}

pub(super) fn parse_field_switch_after_intro(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
    field_type: &str,
) -> Result<Switch> {
    let name = characters.next().ok_or_else(|| {
        Error::Invalid(format!("{field_type} field ends with a switch introducer"))
    })?;
    if name == '\\' || name.is_whitespace() {
        return Err(Error::Invalid(format!(
            "{field_type} field has an invalid switch name"
        )));
    }
    skip_field_whitespace(characters);
    let argument = match characters.peek().copied() {
        None | Some('\\') => None,
        Some('"') => {
            characters.next();
            Some(parse_field_quoted_argument(characters, field_type)?)
        },
        Some(_) => Some(parse_field_unquoted_argument(characters)),
    };
    Ok(Switch {
        name: name.to_ascii_lowercase(),
        argument,
    })
}

pub(super) fn skip_field_whitespace(characters: &mut std::iter::Peekable<std::str::Chars<'_>>) {
    while characters
        .peek()
        .is_some_and(|character| character.is_whitespace())
    {
        characters.next();
    }
}

pub(super) fn parse_field_quoted_argument(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
    field_type: &str,
) -> Result<String> {
    let mut argument = String::new();
    let mut escaped = false;
    while let Some(character) = characters.next() {
        if escaped {
            if character != '\\' && character != '"' {
                argument.push('\\');
            }
            argument.push(character);
            escaped = false;
            continue;
        }
        match character {
            '\\' => escaped = true,
            '"' => {
                if characters
                    .peek()
                    .is_some_and(|next| !next.is_whitespace() && *next != '\\')
                {
                    return Err(Error::Invalid(format!(
                        "{field_type} quoted switch argument has trailing text"
                    )));
                }
                return Ok(argument);
            },
            _ => argument.push(character),
        }
    }
    Err(Error::Invalid(format!(
        "{field_type} field has an unterminated quoted switch argument"
    )))
}

pub(super) fn parse_field_unquoted_argument(
    characters: &mut std::iter::Peekable<std::str::Chars<'_>>,
) -> String {
    let mut argument = String::new();
    while characters
        .peek()
        .is_some_and(|character| !character.is_whitespace() && *character != '\\')
    {
        argument.push(characters.next().expect("checked field argument character"));
    }
    argument
}

pub(super) fn parse_toc_level_range(value: &str) -> Result<TocLevelRange> {
    let mut levels = value.split('-').map(str::trim);
    let start = levels
        .next()
        .ok_or_else(|| Error::Invalid("TOC level range is empty".to_string()))?
        .parse::<u8>()
        .map_err(|_| Error::Invalid("invalid TOC start level".to_string()))?;
    let end = levels
        .next()
        .ok_or_else(|| Error::Invalid("TOC level range is incomplete".to_string()))?
        .parse::<u8>()
        .map_err(|_| Error::Invalid("invalid TOC end level".to_string()))?;
    if levels.next().is_some() {
        return Err(Error::Invalid(
            "TOC level range contains too many separators".to_string(),
        ));
    }
    TocLevelRange::new(start, end)
}

struct PendingSimpleField {
    order: usize,
    instruction: String,
    result: String,
    dirty: bool,
    locked: bool,
}

impl PendingSimpleField {
    fn parse(element: &BytesStart<'_>, decoder: Decoder, order: usize) -> Result<Self> {
        let mut instruction = None;
        let mut dirty = false;
        let mut locked = false;
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?;
            match attribute.key.local_name().as_ref() {
                b"instr" => instruction = Some(value.into_owned()),
                b"dirty" => dirty = is_on(&value),
                b"fldLock" => locked = is_on(&value),
                _ => {},
            }
        }
        let instruction = instruction
            .ok_or_else(|| Error::Invalid("DOCX simple field is missing w:instr".to_string()))?;
        Ok(Self {
            order,
            instruction,
            result: String::new(),
            dirty,
            locked,
        })
    }

    fn finish(self) -> Field {
        let result = (!self.result.is_empty()).then_some(self.result);
        Field::with_flags(
            self.instruction.trim().to_string(),
            result,
            self.dirty,
            self.locked,
        )
    }
}

pub(super) fn is_on(value: &str) -> bool {
    matches!(value, "true" | "1" | "on")
}
