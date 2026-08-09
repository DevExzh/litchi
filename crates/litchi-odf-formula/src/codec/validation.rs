//! `MathML` 2.0 presentation-schema validation.
//!
//! ODF 1.4 Part 3, 2.2.1(B.4), requires a `math`-rooted `content.xml`
//! to validate against the `MathML` 2.0 schema.  The ODF Relax NG schema in
//! `3rdparty/specs/OpenDocument-v1.4-os.zip` deliberately uses an `anyName`
//! placeholder for `MathML`, so this module implements the bounded subset of
//! the referenced `MathML` 2 presentation schema that the model exposes.
//! ODF Part 4 `OpenFormula` governs spreadsheet formula attributes, not this
//! standalone `MathML`-rooted document family.

use litchi_core::{Error, Result};

use crate::model::{Content, Element, MATHML_NAMESPACE};

/// Validate a complete `MathML` tree without recursion.
///
/// # Errors
///
/// Returns an error when the tree lacks a `math` root or an element has an
/// invalid content model, arity, or recognized schema-constrained attribute
/// value.
pub fn validate(root: &Element) -> Result<()> {
    if root.namespace_uri() != Some(MATHML_NAMESPACE) || root.local_name() != "math" {
        return Err(invalid("formula content must have a MathML math root"));
    }
    validate_subtree(root)
}

pub(crate) fn validate_subtree(root: &Element) -> Result<()> {
    let mut pending = vec![root];
    while let Some(element) = pending.pop() {
        validate_element(element)?;
        for child in element.children() {
            pending.push(child);
        }
    }
    Ok(())
}

pub(crate) fn validate_element(element: &Element) -> Result<()> {
    if element.namespace_uri() != Some(MATHML_NAMESPACE) {
        return Ok(());
    }
    validate_attributes(element)?;
    match element.local_name() {
        "math" | "mrow" | "msqrt" | "mstyle" | "merror" | "mpadded" | "mphantom" | "mfenced"
        | "menclose" | "mtd" => validate_expression_sequence(element),
        "semantics" => validate_semantics(element),
        "annotation" => validate_text_only(element),
        "annotation-xml" => Ok(()),
        "mi" | "mn" | "mo" | "mtext" | "ms" => validate_token(element),
        "mspace" | "mglyph" | "maligngroup" | "malignmark" | "none" | "mprescripts" => {
            validate_empty(element)
        },
        "mfrac" | "mroot" | "msub" | "msup" | "munder" | "mover" => {
            validate_exact_expressions(element, 2)
        },
        "msubsup" | "munderover" => validate_exact_expressions(element, 3),
        "mmultiscripts" => validate_multiscripts(element),
        "mtable" => validate_named_children(element, &["mtr", "mlabeledtr"]),
        "mtr" | "mlabeledtr" => validate_named_children(element, &["mtd"]),
        "maction" => validate_nonempty_expressions(element),
        local_name => Err(invalid(format!(
            "MathML 2.0 schema does not define element '{local_name}'"
        ))),
    }
}

fn validate_attributes(element: &Element) -> Result<()> {
    for attribute in element.attributes() {
        if attribute.namespace_uri().is_some() {
            continue;
        }
        let value = attribute.value();
        let valid = match attribute.local_name() {
            "align" => is_table_align(value),
            "accent" | "accentunder" | "alignmentscope" | "bevelled" | "displaystyle"
            | "equalcolumns" | "equalrows" | "fence" | "largeop" | "movablelimits"
            | "separator" | "stretchy" | "symmetric" => is_boolean(value),
            "columnalign" => value_list(value, &["left", "center", "right"]),
            "columnlines" | "rowlines" => value_list(value, &["none", "solid", "dashed"]),
            "denomalign" | "numalign" => value_in(value, &["left", "center", "right"]),
            "dir" => value_in(value, &["ltr", "rtl"]),
            "display" => value_in(value, &["block", "inline"]),
            "edge" => value_in(value, &["left", "right"]),
            "fontstyle" => value_in(value, &["normal", "italic"]),
            "fontweight" => value_in(value, &["normal", "bold"]),
            "form" => value_in(value, &["prefix", "infix", "postfix"]),
            "frame" => value_in(value, &["none", "solid", "dashed"]),
            "groupalign" => value_list(value, &["left", "center", "right", "decimalpoint"]),
            "linebreak" => value_in(
                value,
                &["auto", "newline", "nobreak", "goodbreak", "badbreak"],
            ),
            "linebreakstyle" => value_in(value, &["before", "after", "duplicate"]),
            "overflow" => value_in(
                value,
                &["linebreak", "scroll", "elide", "truncate", "scale"],
            ),
            "linethickness" => value_in(value, &["thin", "medium", "thick"]) || is_length(value),
            "columnspacing" | "rowspacing" => length_list(value),
            "columnwidth" => {
                let mut items = value.split_ascii_whitespace();
                let valid = items
                    .clone()
                    .all(|item| matches!(item, "auto" | "fit") || is_length(item));
                valid && items.next().is_some()
            },
            "framespacing" => {
                let values: Vec<_> = value.split_ascii_whitespace().collect();
                values.len() == 2 && values.into_iter().all(is_length)
            },
            "mathvariant" => is_math_variant(value),
            "notation" => value_list(
                value,
                &[
                    "longdiv",
                    "actuarial",
                    "radical",
                    "box",
                    "roundedbox",
                    "circle",
                    "left",
                    "right",
                    "top",
                    "bottom",
                    "updiagonalstrike",
                    "downdiagonalstrike",
                    "verticalstrike",
                    "horizontalstrike",
                ],
            ),
            "rowalign" => value_list(value, &["top", "bottom", "center", "baseline", "axis"]),
            "index" | "rowspan" | "columnspan" | "selection" => is_positive_integer(value),
            "scriptlevel" => is_script_level(value),
            "side" => value_in(value, &["left", "right", "leftoverlap", "rightoverlap"]),
            "actiontype" => value_in(value, &["toggle", "statusline", "tooltip", "highlight"]),
            // These MathML 2 attributes use the schema's length family.
            "mathsize" => is_length(value) || matches!(value, "normal" | "small" | "big"),
            "width" => is_length(value) || value == "auto",
            "depth" | "height" | "lspace" | "minlabelspacing" | "rspace" | "subscriptshift"
            | "superscriptshift" => is_length(value),
            // Free text, URI, color, or forward-compatible attributes retain
            // their lexical value. Their schemas do not have a finite domain.
            _ => true,
        };
        if !valid {
            return Err(invalid(format!(
                "invalid MathML {} value '{}' on {}",
                attribute.local_name(),
                value,
                element.local_name()
            )));
        }
    }
    Ok(())
}

fn validate_empty(element: &Element) -> Result<()> {
    if element.content().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!(
            "MathML {} must be empty",
            element.local_name()
        )))
    }
}

fn validate_exact_expressions(element: &Element, expected: usize) -> Result<()> {
    validate_expression_sequence(element)?;
    let actual = element.children().count();
    if actual == expected {
        Ok(())
    } else {
        Err(invalid(format!(
            "MathML {} requires exactly {expected} element children, found {actual}",
            element.local_name()
        )))
    }
}

fn validate_expression_sequence(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    for child in element.children() {
        if !is_expression(child) {
            return Err(invalid(format!(
                "MathML {} cannot contain {}",
                element.local_name(),
                child.local_name()
            )));
        }
    }
    Ok(())
}

fn validate_multiscripts(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    let Some(base) = children.first() else {
        return Err(invalid("MathML mmultiscripts requires a base expression"));
    };
    if !is_expression(base) {
        return Err(invalid("MathML mmultiscripts base is not an expression"));
    }
    let mut marker = None;
    for (index, child) in children.iter().enumerate().skip(1) {
        if child.namespace_uri() == Some(MATHML_NAMESPACE) && child.local_name() == "mprescripts" {
            if marker.replace(index).is_some() {
                return Err(invalid(
                    "MathML mmultiscripts has multiple mprescripts markers",
                ));
            }
        } else if !is_script(child) {
            return Err(invalid(format!(
                "MathML mmultiscripts has invalid script {}",
                child.local_name()
            )));
        }
    }
    let marker_index = marker.unwrap_or(children.len());
    let post_count = marker_index.saturating_sub(1);
    let pre_count = children
        .len()
        .saturating_sub(marker_index.saturating_add(1));
    if post_count.is_multiple_of(2) && pre_count.is_multiple_of(2) {
        Ok(())
    } else {
        Err(invalid(
            "MathML mmultiscripts scripts must occur in subscript/superscript pairs",
        ))
    }
}

fn validate_named_children(element: &Element, names: &[&str]) -> Result<()> {
    validate_no_character_data(element)?;
    for child in element.children() {
        if child.namespace_uri() != Some(MATHML_NAMESPACE) || !names.contains(&child.local_name()) {
            return Err(invalid(format!(
                "MathML {} cannot contain {}",
                element.local_name(),
                child.local_name()
            )));
        }
    }
    Ok(())
}

fn validate_no_character_data(element: &Element) -> Result<()> {
    if element.content().iter().all(|content| match content {
        Content::Text(text) => text.chars().all(char::is_whitespace),
        Content::Element(_) => true,
    }) {
        Ok(())
    } else {
        Err(invalid(format!(
            "MathML {} has character data outside a token",
            element.local_name()
        )))
    }
}

fn validate_nonempty_expressions(element: &Element) -> Result<()> {
    validate_expression_sequence(element)?;
    if element.children().next().is_some() {
        Ok(())
    } else {
        Err(invalid("MathML maction requires at least one expression"))
    }
}

fn validate_semantics(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let mut children = element.children();
    let first = children
        .next()
        .ok_or_else(|| invalid("MathML semantics requires a primary expression"))?;
    if !is_expression(first) {
        return Err(invalid(
            "MathML semantics primary child is not an expression",
        ));
    }
    for child in children {
        if child.namespace_uri() != Some(MATHML_NAMESPACE)
            || !matches!(child.local_name(), "annotation" | "annotation-xml")
        {
            return Err(invalid(
                "MathML semantics permits only annotations after its primary expression",
            ));
        }
    }
    Ok(())
}

fn validate_text_only(element: &Element) -> Result<()> {
    if element
        .content()
        .iter()
        .all(|content| matches!(content, Content::Text(_)))
    {
        Ok(())
    } else {
        Err(invalid("MathML annotation may contain character data only"))
    }
}

fn validate_token(element: &Element) -> Result<()> {
    for content in element.content() {
        let valid = match content {
            Content::Text(_) => true,
            Content::Element(child) => {
                child.namespace_uri() == Some(MATHML_NAMESPACE)
                    && matches!(child.local_name(), "mglyph" | "malignmark")
            },
        };
        if !valid {
            return Err(invalid(format!(
                "MathML token {} contains a non-token child",
                element.local_name()
            )));
        }
    }
    Ok(())
}

fn is_boolean(value: &str) -> bool {
    matches!(value, "true" | "false")
}

fn is_expression(element: &Element) -> bool {
    if element.namespace_uri() != Some(MATHML_NAMESPACE) {
        return false;
    }
    matches!(
        element.local_name(),
        "semantics"
            | "mi"
            | "mn"
            | "mo"
            | "mtext"
            | "mspace"
            | "ms"
            | "mrow"
            | "mfrac"
            | "msqrt"
            | "mroot"
            | "mstyle"
            | "merror"
            | "mpadded"
            | "mphantom"
            | "mfenced"
            | "menclose"
            | "msub"
            | "msup"
            | "msubsup"
            | "munder"
            | "mover"
            | "munderover"
            | "mmultiscripts"
            | "mtable"
            | "maction"
    )
}

fn is_length(value: &str) -> bool {
    const NAMED: &[&str] = &[
        "veryverythinmathspace",
        "verythinmathspace",
        "thinmathspace",
        "mediummathspace",
        "thickmathspace",
        "verythickmathspace",
        "veryverythickmathspace",
        "negativeveryverythinmathspace",
        "negativeverythinmathspace",
        "negativethinmathspace",
        "negativemediummathspace",
        "negativethickmathspace",
        "negativeverythickmathspace",
        "negativeveryverythickmathspace",
        "infinity",
    ];
    if NAMED.contains(&value) {
        return true;
    }
    let split = value
        .char_indices()
        .find(|(_, character)| character.is_ascii_alphabetic() || *character == '%')
        .map_or(value.len(), |(index, _)| index);
    let (number, unit) = value.split_at(split);
    let number_valid =
        number.parse::<f64>().is_ok() && number.bytes().any(|byte| byte.is_ascii_digit());
    number_valid
        && matches!(
            unit,
            "" | "em" | "ex" | "px" | "in" | "cm" | "mm" | "pt" | "pc" | "%"
        )
}

fn length_list(value: &str) -> bool {
    let mut items = value.split_ascii_whitespace();
    let valid = items.clone().all(is_length);
    valid && items.next().is_some()
}

fn is_math_variant(value: &str) -> bool {
    value_in(
        value,
        &[
            "normal",
            "bold",
            "italic",
            "bold-italic",
            "double-struck",
            "bold-fraktur",
            "script",
            "bold-script",
            "fraktur",
            "sans-serif",
            "bold-sans-serif",
            "sans-serif-italic",
            "sans-serif-bold-italic",
            "monospace",
        ],
    )
}

fn is_positive_integer(value: &str) -> bool {
    value.parse::<usize>().is_ok_and(|number| number > 0)
}

fn is_script(element: &Element) -> bool {
    is_expression(element)
        || (element.namespace_uri() == Some(MATHML_NAMESPACE) && element.local_name() == "none")
}

fn is_script_level(value: &str) -> bool {
    value.parse::<i32>().is_ok()
        || value.strip_prefix(['+', '-']).is_some_and(|digits| {
            !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit())
        })
}

fn is_table_align(value: &str) -> bool {
    let mut items = value.split_ascii_whitespace();
    let Some(alignment) = items.next() else {
        return false;
    };
    if !value_in(alignment, &["top", "bottom", "center", "baseline", "axis"]) {
        return false;
    }
    let row_valid = items.next().is_none_or(|row| row.parse::<i32>().is_ok());
    row_valid && items.next().is_none()
}

fn value_in(value: &str, values: &[&str]) -> bool {
    values.contains(&value)
}

fn value_list(value: &str, values: &[&str]) -> bool {
    let mut items = value.split_ascii_whitespace();
    let valid = items.clone().all(|item| values.contains(&item));
    valid && items.next().is_some()
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
