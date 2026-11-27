//! Bounded external field-instruction parsers.

use super::prelude::*;

pub(in crate::parts::fields) fn parse_dde_field_parts(instruction: &str) -> Option<DdeParts> {
    if instruction.len() > MAX_DDE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind = if keyword.eq_ignore_ascii_case("DDE") {
        DdeFieldKind::Dde
    } else if keyword.eq_ignore_ascii_case("DDEAUTO") {
        DdeFieldKind::DdeAuto
    } else {
        return None;
    };

    let application = next_field_argument(instruction, &mut position).ok()??;
    if application.is_empty() {
        return None;
    }
    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let item = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut automatic_updates = kind == DdeFieldKind::DdeAuto;
    let mut saw_automatic_update = false;
    let mut representation = None;
    let mut omit_graphic_data = false;
    let mut unknown_switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || unknown_switches.len() >= MAX_DDE_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        match name {
            'a' if kind == DdeFieldKind::Dde => {
                if saw_automatic_update || argument.is_some() {
                    return None;
                }
                automatic_updates = true;
                saw_automatic_update = true;
            },
            'a' => return None,
            'd' => {
                if representation.is_some() || omit_graphic_data || argument.is_some() {
                    return None;
                }
                omit_graphic_data = true;
            },
            'b' | 'h' | 'p' | 'r' | 't' | 'u' => {
                if representation.is_some() || omit_graphic_data || argument.is_some() {
                    return None;
                }
                representation = Some(match name {
                    'b' => DdeRepresentation::Bitmap,
                    'h' => DdeRepresentation::Html,
                    'p' => DdeRepresentation::Picture,
                    'r' => DdeRepresentation::RichText,
                    't' => DdeRepresentation::Text,
                    'u' => DdeRepresentation::UnicodeText,
                    _ => unreachable!("DDE representation switch was matched above"),
                });
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(DdeParts {
        kind,
        application,
        source,
        item,
        automatic_updates,
        representation,
        omit_graphic_data,
        unknown_switches,
    })
}

pub(in crate::parts::fields) fn parse_link_field_parts(instruction: &str) -> Option<LinkParts> {
    if instruction.len() > MAX_LINK_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    if !keyword.eq_ignore_ascii_case("LINK") {
        return None;
    }

    let application_type = next_field_argument(instruction, &mut position).ok()??;
    if application_type.is_empty() {
        return None;
    }
    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let item = match peek_field_character(instruction, position) {
        None | Some('\\') => None,
        Some(_) => Some(next_field_argument(instruction, &mut position).ok()??),
    };

    let mut switches = Vec::new();
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switches.len() >= MAX_LINK_FIELD_SWITCHES {
            return None;
        }

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };
        switches.push(MergeFieldSwitch { name, argument });
    }

    let mut automatic_updates = false;
    let mut result_options = Vec::new();
    let mut formatting_modes = Vec::new();
    for switch in &switches {
        match switch.name {
            'a' => {
                if switch.argument.is_some() {
                    return None;
                }
                automatic_updates = true;
            },
            'f' => {
                let value = switch.argument.as_deref()?.parse::<i64>().ok()?;
                formatting_modes.push(match value {
                    0 => LinkFormatting::Source,
                    2 => LinkFormatting::Destination,
                    4 => LinkFormatting::SpreadsheetSource,
                    5 => LinkFormatting::SpreadsheetDestination,
                    other => LinkFormatting::Unsupported(other),
                });
            },
            'b' | 'd' | 'h' | 'p' | 'r' | 't' | 'u' => {
                if switch.argument.is_some() {
                    return None;
                }
                result_options.push(match switch.name {
                    'b' => LinkResultOption::Bitmap,
                    'd' => LinkResultOption::OmitGraphicData,
                    'h' => LinkResultOption::Html,
                    'p' => LinkResultOption::Picture,
                    'r' => LinkResultOption::RichText,
                    't' => LinkResultOption::Text,
                    'u' => LinkResultOption::UnicodeText,
                    _ => unreachable!("LINK result switch was matched above"),
                });
            },
            _ => {},
        }
    }

    Some(LinkParts {
        application_type,
        source,
        item,
        automatic_updates,
        result_options,
        formatting_modes,
        switches,
    })
}

pub(in crate::parts::fields) fn parse_external_include_field_parts(
    instruction: &str,
) -> Option<ExternalIncludeParts> {
    if instruction.len() > MAX_EXTERNAL_INCLUDE_FIELD_INSTRUCTION_BYTES {
        return None;
    }

    let mut position = 0;
    let keyword = next_field_argument(instruction, &mut position).ok()??;
    let kind =
        if keyword.eq_ignore_ascii_case("INCLUDETEXT") || keyword.eq_ignore_ascii_case("INCLUDE") {
            IncludeFieldKind::Text
        } else if keyword.eq_ignore_ascii_case("INCLUDEPICTURE")
            || keyword.eq_ignore_ascii_case("IMPORT")
        {
            IncludeFieldKind::Picture
        } else {
            return None;
        };

    let source = next_field_argument(instruction, &mut position).ok()??;
    if source.is_empty() {
        return None;
    }

    skip_field_whitespace(instruction, &mut position);
    let bookmark = match (kind, peek_field_character(instruction, position)) {
        (IncludeFieldKind::Text, None | Some('\\')) => None,
        (IncludeFieldKind::Text, Some(_)) => {
            Some(next_field_argument(instruction, &mut position).ok()??)
        },
        (IncludeFieldKind::Picture, _) => None,
    };

    let mut suppress_nested_field_updates = false;
    let mut omit_picture_data = false;
    let mut options = Vec::new();
    let mut unknown_switches = Vec::new();
    let mut switch_count = 0;
    loop {
        skip_field_whitespace(instruction, &mut position);
        let Some(introducer) = next_field_character(instruction, &mut position) else {
            break;
        };
        if introducer != '\\' || switch_count >= MAX_EXTERNAL_INCLUDE_FIELD_SWITCHES {
            return None;
        }
        switch_count += 1;

        let name = next_field_character(instruction, &mut position)?;
        if name == '\\' || name.is_whitespace() {
            return None;
        }
        let name = name.to_ascii_lowercase();

        skip_field_whitespace(instruction, &mut position);
        let argument = match peek_field_character(instruction, position) {
            None | Some('\\') => None,
            Some(_) => next_field_argument(instruction, &mut position).ok()?,
        };

        match (kind, name) {
            (_, 'c') => options.push(ExternalIncludeOption::Converter(argument?)),
            (IncludeFieldKind::Text, 'e') => {
                options.push(ExternalIncludeOption::Encoding(argument?));
            },
            (IncludeFieldKind::Text, 'm') => {
                options.push(ExternalIncludeOption::MimeType(argument?));
            },
            (IncludeFieldKind::Text, 'n') => {
                options.push(ExternalIncludeOption::NamespaceMapping(argument?));
            },
            (IncludeFieldKind::Text, 't') => {
                options.push(ExternalIncludeOption::Xslt(argument?));
            },
            (IncludeFieldKind::Text, 'x') => {
                options.push(ExternalIncludeOption::XPath(argument?));
            },
            (IncludeFieldKind::Text, '!') => {
                if suppress_nested_field_updates || argument.is_some() {
                    return None;
                }
                suppress_nested_field_updates = true;
            },
            (IncludeFieldKind::Picture, 'd') => {
                if omit_picture_data || argument.is_some() {
                    return None;
                }
                omit_picture_data = true;
            },
            _ => unknown_switches.push(MergeFieldSwitch { name, argument }),
        }
    }

    Some(ExternalIncludeParts {
        kind,
        source,
        bookmark,
        suppress_nested_field_updates,
        omit_picture_data,
        options,
        unknown_switches,
    })
}
