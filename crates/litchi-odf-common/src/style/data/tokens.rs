use super::{
    Attribute, Calendar, Clock, Currency, EmbeddedText, Fraction, Kind, LOEXT, Locale, MAX_PARTS,
    Map, Month, NUMBER, Node, NumberToken, Result, STYLE, Scientific, Seconds, ShortLong, Token,
    Version, WeekOfYear, ensure_empty_node, ensure_no_children, ensure_whitespace, invalid,
    parse_locale, reject_remaining, required, required_i64, take, take_bool, take_f64, take_i64,
    take_versioned_bool, take_versioned_i64, take_versioned_u64, validate_locale,
    validate_optional_string, validate_text,
};

pub(crate) fn parse_part_node(mut node: Node, version: Version) -> Result<(Token, bool)> {
    let standard = node.namespace.as_deref() == Some(NUMBER);
    let lo_fill = node.namespace.as_deref() == Some(LOEXT) && node.local == "fill-character";
    if !standard && !lo_fill {
        return invalid(format!(
            "unexpected data-style child {}:{}",
            node.namespace.as_deref().unwrap_or(""),
            node.local
        ));
    }
    if lo_fill {
        reject_remaining(&node.attributes, "loext:fill-character")?;
        ensure_no_children(&node, "loext:fill-character")?;
        validate_text(&node.text, "loext:fill-character")?;
        return Ok((Token::FillCharacter(node.text), true));
    }
    let mut alias = false;
    let token = match node.local.as_str() {
        "text" => {
            reject_remaining(&node.attributes, "number:text")?;
            ensure_no_children(&node, "number:text")?;
            Token::Text(node.text)
        },
        "fill-character" => {
            if version == Version::V1_2 {
                return invalid("number:fill-character requires ODF 1.3");
            }
            reject_remaining(&node.attributes, "number:fill-character")?;
            ensure_no_children(&node, "number:fill-character")?;
            Token::FillCharacter(node.text)
        },
        "number" => {
            ensure_whitespace(&node.text, "number:number")?;
            let mut embedded_text = Vec::new();
            for mut child in node.children.drain(..) {
                if child.namespace.as_deref() != Some(NUMBER) || child.local != "embedded-text" {
                    return invalid("number:number may contain only number:embedded-text");
                }
                ensure_no_children(&child, "number:embedded-text")?;
                let position = required_i64(&mut child.attributes, NUMBER, "position")?;
                reject_remaining(&child.attributes, "number:embedded-text")?;
                embedded_text.push(EmbeddedText {
                    position,
                    text: child.text,
                });
            }
            let min_decimal_places = take_versioned_i64(
                &mut node.attributes,
                "min-decimal-places",
                version,
                &mut alias,
            )?;
            let value = NumberToken {
                decimal_replacement: take(&mut node.attributes, NUMBER, "decimal-replacement"),
                display_factor: take_f64(&mut node.attributes, NUMBER, "display-factor")?,
                decimal_places: take_i64(&mut node.attributes, NUMBER, "decimal-places")?,
                min_decimal_places,
                min_integer_digits: take_i64(&mut node.attributes, NUMBER, "min-integer-digits")?,
                grouping: take_bool(&mut node.attributes, NUMBER, "grouping")?,
                embedded_text,
            };
            reject_remaining(&node.attributes, "number:number")?;
            Token::Number(value)
        },
        "scientific-number" => {
            ensure_empty_node(&node, "number:scientific-number")?;
            let exponent_interval = take_versioned_u64(
                &mut node.attributes,
                "exponent-interval",
                version,
                &mut alias,
            )?;
            let forced_exponent_sign = take_versioned_bool(
                &mut node.attributes,
                "forced-exponent-sign",
                version,
                &mut alias,
            )?;
            let min_decimal_places = take_versioned_i64(
                &mut node.attributes,
                "min-decimal-places",
                version,
                &mut alias,
            )?;
            let value = Scientific {
                min_exponent_digits: take_i64(&mut node.attributes, NUMBER, "min-exponent-digits")?,
                exponent_interval,
                forced_exponent_sign,
                decimal_places: take_i64(&mut node.attributes, NUMBER, "decimal-places")?,
                min_decimal_places,
                min_integer_digits: take_i64(&mut node.attributes, NUMBER, "min-integer-digits")?,
                grouping: take_bool(&mut node.attributes, NUMBER, "grouping")?,
            };
            reject_remaining(&node.attributes, "number:scientific-number")?;
            Token::ScientificNumber(value)
        },
        "fraction" => {
            ensure_empty_node(&node, "number:fraction")?;
            let max_denominator_value = take_versioned_u64(
                &mut node.attributes,
                "max-denominator-value",
                version,
                &mut alias,
            )?;
            let value = Fraction {
                min_numerator_digits: take_i64(
                    &mut node.attributes,
                    NUMBER,
                    "min-numerator-digits",
                )?,
                min_denominator_digits: take_i64(
                    &mut node.attributes,
                    NUMBER,
                    "min-denominator-digits",
                )?,
                denominator_value: take_i64(&mut node.attributes, NUMBER, "denominator-value")?,
                max_denominator_value,
                min_integer_digits: take_i64(&mut node.attributes, NUMBER, "min-integer-digits")?,
                grouping: take_bool(&mut node.attributes, NUMBER, "grouping")?,
            };
            reject_remaining(&node.attributes, "number:fraction")?;
            Token::Fraction(value)
        },
        "currency-symbol" => {
            ensure_no_children(&node, "number:currency-symbol")?;
            let locale = parse_locale(&mut node.attributes);
            reject_remaining(&node.attributes, "number:currency-symbol")?;
            Token::CurrencySymbol(Currency {
                locale,
                text: node.text,
            })
        },
        "day" => Token::Day(parse_calendar(&mut node)?),
        "month" => {
            ensure_empty_node(&node, "number:month")?;
            let value = Month {
                style: take_style(&mut node.attributes)?,
                textual: take_bool(&mut node.attributes, NUMBER, "textual")?,
                possessive_form: take_bool(&mut node.attributes, NUMBER, "possessive-form")?,
                calendar: take(&mut node.attributes, NUMBER, "calendar"),
            };
            reject_remaining(&node.attributes, "number:month")?;
            Token::Month(value)
        },
        "year" => Token::Year(parse_calendar(&mut node)?),
        "era" => Token::Era(parse_calendar(&mut node)?),
        "day-of-week" => Token::DayOfWeek(parse_calendar(&mut node)?),
        "week-of-year" => {
            ensure_empty_node(&node, "number:week-of-year")?;
            let value = WeekOfYear {
                calendar: take(&mut node.attributes, NUMBER, "calendar"),
            };
            reject_remaining(&node.attributes, "number:week-of-year")?;
            Token::WeekOfYear(value)
        },
        "quarter" => Token::Quarter(parse_calendar(&mut node)?),
        "hours" => Token::Hours(parse_clock(&mut node)?),
        "minutes" => Token::Minutes(parse_clock(&mut node)?),
        "seconds" => {
            ensure_empty_node(&node, "number:seconds")?;
            let value = Seconds {
                style: take_style(&mut node.attributes)?,
                decimal_places: take_i64(&mut node.attributes, NUMBER, "decimal-places")?,
            };
            reject_remaining(&node.attributes, "number:seconds")?;
            Token::Seconds(value)
        },
        "am-pm" => {
            ensure_empty_node(&node, "number:am-pm")?;
            reject_remaining(&node.attributes, "number:am-pm")?;
            Token::AmPm
        },
        "boolean" => {
            ensure_empty_node(&node, "number:boolean")?;
            reject_remaining(&node.attributes, "number:boolean")?;
            Token::Boolean
        },
        "text-content" => {
            ensure_empty_node(&node, "number:text-content")?;
            reject_remaining(&node.attributes, "number:text-content")?;
            Token::TextContent
        },
        _ => return invalid(format!("unexpected number:{} token", node.local)),
    };
    Ok((token, alias))
}

pub(crate) fn parse_map(mut node: Node) -> Result<Map> {
    ensure_empty_node(&node, "style:map")?;
    let value = Map {
        condition: required(&mut node.attributes, STYLE, "condition")?,
        apply_style_name: required(&mut node.attributes, STYLE, "apply-style-name")?,
        base_cell_address: take(&mut node.attributes, STYLE, "base-cell-address"),
    };
    reject_remaining(&node.attributes, "style:map")?;
    Ok(value)
}

pub(crate) fn parse_calendar(node: &mut Node) -> Result<Calendar> {
    ensure_empty_node(node, "calendar token")?;
    let value = Calendar {
        style: take_style(&mut node.attributes)?,
        calendar: take(&mut node.attributes, NUMBER, "calendar"),
    };
    reject_remaining(&node.attributes, "calendar token")?;
    Ok(value)
}

pub(crate) fn parse_clock(node: &mut Node) -> Result<Clock> {
    ensure_empty_node(node, "clock token")?;
    let value = Clock {
        style: take_style(&mut node.attributes)?,
    };
    reject_remaining(&node.attributes, "clock token")?;
    Ok(value)
}

pub(crate) fn take_style(attributes: &mut Vec<Attribute>) -> Result<Option<ShortLong>> {
    take(attributes, NUMBER, "style")
        .map(|value| ShortLong::parse(&value))
        .transpose()
}

pub(crate) fn validate_sequence(
    kind: Kind,
    parts: &[Token],
    version: Version,
    allow_lo: bool,
) -> Result<()> {
    let allow_fill = version == Version::V1_3 || allow_lo;
    let mut index = 0usize;
    match kind {
        Kind::Boolean => {
            consume_plain_text(parts, &mut index);
            if matches!(parts.get(index), Some(Token::Boolean)) {
                index += 1;
                consume_plain_text(parts, &mut index);
            }
        },
        Kind::Number => {
            consume_separator(parts, &mut index, allow_fill);
            if matches!(
                parts.get(index),
                Some(Token::Number(_) | Token::ScientificNumber(_) | Token::Fraction(_))
            ) {
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
            }
        },
        Kind::Percentage => {
            consume_separator(parts, &mut index, allow_fill);
            if matches!(parts.get(index), Some(Token::Number(_))) {
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
            }
        },
        Kind::Currency => {
            consume_separator(parts, &mut index, allow_fill);
            if matches!(parts.get(index), Some(Token::Number(_))) {
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
                if matches!(parts.get(index), Some(Token::CurrencySymbol(_))) {
                    index += 1;
                    consume_separator(parts, &mut index, allow_fill);
                }
            } else if matches!(parts.get(index), Some(Token::CurrencySymbol(_))) {
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
                if matches!(parts.get(index), Some(Token::Number(_))) {
                    index += 1;
                    consume_separator(parts, &mut index, allow_fill);
                }
            }
        },
        Kind::Date => {
            consume_separator(parts, &mut index, allow_fill);
            let mut count = 0usize;
            while parts.get(index).is_some_and(is_date_token) {
                count += 1;
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
            }
            if count == 0 {
                return invalid("number:date-style requires at least one date token");
            }
        },
        Kind::Time => {
            consume_separator(parts, &mut index, allow_fill);
            let mut count = 0usize;
            while parts.get(index).is_some_and(is_time_token) {
                count += 1;
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
            }
            if count == 0 {
                return invalid("number:time-style requires at least one time token");
            }
        },
        Kind::Text => {
            consume_separator(parts, &mut index, allow_fill);
            while matches!(parts.get(index), Some(Token::TextContent)) {
                index += 1;
                consume_separator(parts, &mut index, allow_fill);
            }
        },
    }
    if index != parts.len() {
        return invalid(format!(
            "invalid ordered token sequence for number:{}",
            kind.local()
        ));
    }
    Ok(())
}

pub(crate) fn consume_plain_text(parts: &[Token], index: &mut usize) {
    if matches!(parts.get(*index), Some(Token::Text(_))) {
        *index += 1;
    }
}

pub(crate) fn consume_separator(parts: &[Token], index: &mut usize, allow_fill: bool) {
    consume_plain_text(parts, index);
    if allow_fill && matches!(parts.get(*index), Some(Token::FillCharacter(_))) {
        *index += 1;
        consume_plain_text(parts, index);
    }
}

pub(crate) fn is_date_token(part: &Token) -> bool {
    matches!(
        part,
        Token::Day(_)
            | Token::Month(_)
            | Token::Year(_)
            | Token::Era(_)
            | Token::DayOfWeek(_)
            | Token::WeekOfYear(_)
            | Token::Quarter(_)
            | Token::Hours(_)
            | Token::Minutes(_)
            | Token::Seconds(_)
            | Token::AmPm
    )
}

pub(crate) fn is_time_token(part: &Token) -> bool {
    matches!(
        part,
        Token::Hours(_) | Token::Minutes(_) | Token::Seconds(_) | Token::AmPm
    )
}

pub(crate) fn validate_part(part: &Token, version: Version, allow_lo: bool) -> Result<()> {
    match part {
        Token::Text(value) => validate_text(value, "number:text")?,
        Token::FillCharacter(value) => {
            validate_text(value, "number:fill-character")?;
            require_1_3(true, version, allow_lo)?;
        },
        Token::Number(value) => {
            validate_optional_string(
                value.decimal_replacement.as_deref(),
                "number:decimal-replacement",
            )?;
            if value.embedded_text.len() > MAX_PARTS {
                return invalid("too many number:embedded-text elements");
            }
            for embedded in &value.embedded_text {
                validate_text(&embedded.text, "number:embedded-text")?;
            }
            require_1_3(value.min_decimal_places.is_some(), version, allow_lo)?;
        },
        Token::ScientificNumber(value) => {
            require_1_3(
                value.min_decimal_places.is_some()
                    || value.exponent_interval.is_some()
                    || value.forced_exponent_sign.is_some(),
                version,
                allow_lo,
            )?;
            if value.exponent_interval == Some(0) {
                return invalid("number:exponent-interval must be positive");
            }
        },
        Token::Fraction(value) => {
            require_1_3(value.max_denominator_value.is_some(), version, allow_lo)?;
            if value.max_denominator_value == Some(0) {
                return invalid("number:max-denominator-value must be positive");
            }
        },
        Token::CurrencySymbol(value) => {
            validate_locale(&value.locale)?;
            validate_text(&value.text, "number:currency-symbol")?;
        },
        Token::Day(value)
        | Token::Year(value)
        | Token::Era(value)
        | Token::DayOfWeek(value)
        | Token::Quarter(value) => {
            validate_optional_string(value.calendar.as_deref(), "number:calendar")?;
        },
        Token::Month(value) => {
            validate_optional_string(value.calendar.as_deref(), "number:calendar")?;
        },
        Token::WeekOfYear(value) => {
            validate_optional_string(value.calendar.as_deref(), "number:calendar")?;
        },
        Token::Hours(_)
        | Token::Minutes(_)
        | Token::Seconds(_)
        | Token::AmPm
        | Token::Boolean
        | Token::TextContent => {},
    }
    Ok(())
}

pub(crate) fn require_1_3(present: bool, version: Version, allow_lo: bool) -> Result<()> {
    if present && version == Version::V1_2 && !allow_lo {
        return invalid("ODF 1.3 number-format feature used in ODF 1.2");
    }
    Ok(())
}

pub(crate) fn write_part(out: &mut String, part: &Token, version: Version) -> Result<()> {
    match part {
        Token::Text(value) => element_text(out, "number:text", value),
        Token::FillCharacter(value) => {
            if version != Version::V1_3 {
                return invalid("number:fill-character requires ODF 1.3 output");
            }
            element_text(out, "number:fill-character", value);
        },
        Token::Number(value) => {
            out.push_str("<number:number");
            attr(
                out,
                "number:decimal-replacement",
                value.decimal_replacement.as_deref(),
            );
            f64_attr(out, "number:display-factor", value.display_factor);
            i64_attr(out, "number:decimal-places", value.decimal_places);
            i64_attr(out, "number:min-decimal-places", value.min_decimal_places);
            i64_attr(out, "number:min-integer-digits", value.min_integer_digits);
            bool_attr(out, "number:grouping", value.grouping);
            if value.embedded_text.is_empty() {
                out.push_str("/>");
            } else {
                out.push('>');
                for embedded in &value.embedded_text {
                    out.push_str("<number:embedded-text");
                    i64_attr(out, "number:position", Some(embedded.position));
                    out.push('>');
                    out.push_str(&esc(&embedded.text));
                    out.push_str("</number:embedded-text>");
                }
                out.push_str("</number:number>");
            }
        },
        Token::ScientificNumber(value) => {
            out.push_str("<number:scientific-number");
            i64_attr(out, "number:min-exponent-digits", value.min_exponent_digits);
            u64_attr(out, "number:exponent-interval", value.exponent_interval);
            bool_attr(
                out,
                "number:forced-exponent-sign",
                value.forced_exponent_sign,
            );
            i64_attr(out, "number:decimal-places", value.decimal_places);
            i64_attr(out, "number:min-decimal-places", value.min_decimal_places);
            i64_attr(out, "number:min-integer-digits", value.min_integer_digits);
            bool_attr(out, "number:grouping", value.grouping);
            out.push_str("/>");
        },
        Token::Fraction(value) => {
            out.push_str("<number:fraction");
            i64_attr(
                out,
                "number:min-numerator-digits",
                value.min_numerator_digits,
            );
            i64_attr(
                out,
                "number:min-denominator-digits",
                value.min_denominator_digits,
            );
            i64_attr(out, "number:denominator-value", value.denominator_value);
            u64_attr(
                out,
                "number:max-denominator-value",
                value.max_denominator_value,
            );
            i64_attr(out, "number:min-integer-digits", value.min_integer_digits);
            bool_attr(out, "number:grouping", value.grouping);
            out.push_str("/>");
        },
        Token::CurrencySymbol(value) => {
            out.push_str("<number:currency-symbol");
            locale_attrs(out, &value.locale);
            out.push('>');
            out.push_str(&esc(&value.text));
            out.push_str("</number:currency-symbol>");
        },
        Token::Day(value) => write_calendar(out, "day", value),
        Token::Year(value) => write_calendar(out, "year", value),
        Token::Era(value) => write_calendar(out, "era", value),
        Token::DayOfWeek(value) => write_calendar(out, "day-of-week", value),
        Token::Quarter(value) => write_calendar(out, "quarter", value),
        Token::Month(value) => {
            out.push_str("<number:month");
            short_long_attr(out, value.style);
            bool_attr(out, "number:textual", value.textual);
            bool_attr(out, "number:possessive-form", value.possessive_form);
            attr(out, "number:calendar", value.calendar.as_deref());
            out.push_str("/>");
        },
        Token::WeekOfYear(value) => {
            out.push_str("<number:week-of-year");
            attr(out, "number:calendar", value.calendar.as_deref());
            out.push_str("/>");
        },
        Token::Hours(value) => write_clock(out, "hours", value),
        Token::Minutes(value) => write_clock(out, "minutes", value),
        Token::Seconds(value) => {
            out.push_str("<number:seconds");
            short_long_attr(out, value.style);
            i64_attr(out, "number:decimal-places", value.decimal_places);
            out.push_str("/>");
        },
        Token::AmPm => out.push_str("<number:am-pm/>"),
        Token::Boolean => out.push_str("<number:boolean/>"),
        Token::TextContent => out.push_str("<number:text-content/>"),
    }
    Ok(())
}

pub(crate) fn write_calendar(out: &mut String, local: &str, value: &Calendar) {
    out.push_str("<number:");
    out.push_str(local);
    short_long_attr(out, value.style);
    attr(out, "number:calendar", value.calendar.as_deref());
    out.push_str("/>");
}

pub(crate) fn write_clock(out: &mut String, local: &str, value: &Clock) {
    out.push_str("<number:");
    out.push_str(local);
    short_long_attr(out, value.style);
    out.push_str("/>");
}

pub(crate) fn element_text(out: &mut String, qname: &str, value: &str) {
    out.push('<');
    out.push_str(qname);
    out.push('>');
    out.push_str(&esc(value));
    out.push_str("</");
    out.push_str(qname);
    out.push('>');
}

pub(crate) fn locale_attrs(out: &mut String, locale: &Locale) {
    attr(out, "number:language", locale.language.as_deref());
    attr(out, "number:country", locale.country.as_deref());
    attr(out, "number:script", locale.script.as_deref());
    attr(
        out,
        "number:rfc-language-tag",
        locale.rfc_language_tag.as_deref(),
    );
}

pub(crate) fn short_long_attr(out: &mut String, style_option: Option<ShortLong>) {
    if let Some(style_value) = style_option {
        attr(out, "number:style", Some(style_value.as_str()));
    }
}

pub(crate) fn attr(out: &mut String, name: &str, attribute_value_option: Option<&str>) {
    if let Some(attribute_value) = attribute_value_option {
        out.push(' ');
        out.push_str(name);
        out.push_str("=\"");
        out.push_str(&esc(attribute_value));
        out.push('"');
    }
}

pub(crate) fn bool_attr(out: &mut String, name: &str, boolean_option: Option<bool>) {
    if let Some(boolean_value) = boolean_option {
        attr(
            out,
            name,
            Some(if boolean_value { "true" } else { "false" }),
        );
    }
}

pub(crate) fn i64_attr(out: &mut String, name: &str, integer_option: Option<i64>) {
    if let Some(integer_value) = integer_option {
        attr(out, name, Some(&integer_value.to_string()));
    }
}

pub(crate) fn u64_attr(out: &mut String, name: &str, integer_option: Option<u64>) {
    if let Some(integer_value) = integer_option {
        attr(out, name, Some(&integer_value.to_string()));
    }
}

pub(crate) fn f64_attr(out: &mut String, name: &str, number_option: Option<f64>) {
    if let Some(number_value) = number_option {
        let lexical = if number_value.is_nan() {
            "NaN".to_string()
        } else if number_value == f64::INFINITY {
            "INF".to_string()
        } else if number_value == f64::NEG_INFINITY {
            "-INF".to_string()
        } else {
            number_value.to_string()
        };
        attr(out, name, Some(&lexical));
    }
}

pub(crate) fn esc(value: &str) -> String {
    litchi_core::xml::escape_xml(value)
}
