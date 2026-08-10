//! `MathML` 2.0 schema-derived validation.
//!
//! ODF 1.4 Part 3, 2.2.1(B.4), requires a `math`-rooted `content.xml`
//! to validate against the `MathML` 2.0 schema.  The ODF Relax NG schema in
//! `3rdparty/specs/OpenDocument-v1.4-os.zip` deliberately uses an `anyName`
//! placeholder for `MathML`, so this module implements the bounded presentation
//! and content subset of the referenced `MathML` 2 schema that the model
//! exposes.
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
        if element.namespace_uri() == Some(MATHML_NAMESPACE)
            && element.local_name() == "annotation-xml"
        {
            continue;
        }
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
        "math" => validate_math_sequence(element),
        "mrow" | "msqrt" | "mstyle" | "merror" | "mpadded" | "mphantom" | "mfenced"
        | "menclose" | "mtd" => validate_presentation_sequence(element),
        "semantics" => validate_semantics(element),
        "annotation" => validate_text_only(element),
        "annotation-xml" => Ok(()),
        "mi" | "mn" | "mo" | "mtext" | "ms" => validate_token(element),
        "mspace" | "mglyph" | "maligngroup" | "malignmark" | "none" | "mprescripts" | "sep" => {
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
        "apply" => validate_application(element),
        "reln" => validate_relation(element),
        "ci" | "cn" | "csymbol" => validate_content_token(element),
        "interval" => validate_interval(element),
        "piece" => validate_exact_content_expressions(element, 2),
        "condition" => validate_condition(element),
        "degree"
        | "domainofapplication"
        | "logbase"
        | "lowlimit"
        | "momentabout"
        | "otherwise"
        | "uplimit" => validate_exact_content_expressions(element, 1),
        "fn" => validate_exact_math_expressions(element, 1),
        "piecewise" => validate_piecewise(element),
        "lambda" => validate_lambda(element),
        "bvar" => validate_bound_variable(element),
        "declare" => validate_declare(element),
        "list" | "set" => validate_list_or_set(element),
        "matrixrow" => validate_matrix_row(element),
        "vector" => validate_vector(element),
        "matrix" => validate_matrix(element),
        local_name if is_content_symbol(local_name) => validate_empty(element),
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
            "rowspan" | "columnspan" | "selection" => is_positive_integer(value),
            "base" => {
                if element.local_name() == "cn" {
                    value
                        .parse::<u8>()
                        .is_ok_and(|base| (2..=36).contains(&base))
                } else {
                    is_positive_integer(value)
                }
            },
            "index" | "nargs" => is_nonnegative_integer(value),
            "occurrence" => value_in(value, &["prefix", "infix", "function-model"]),
            "closure" => value_in(value, &["open", "closed", "open-closed", "closed-open"]),
            "order" => value_in(value, &["numeric", "lexicographic"]),
            "scriptlevel" => is_script_level(value),
            "side" => value_in(value, &["left", "right", "leftoverlap", "rightoverlap"]),
            "actiontype" => value_in(value, &["toggle", "statusline", "tooltip", "highlight"]),
            "type" => validate_type_attribute(element.local_name(), value),
            "scope" => value_in(value, &["local", "global"]),
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
    validate_required_attributes(element)
}

fn validate_required_attributes(element: &Element) -> Result<()> {
    if element.local_name() == "mglyph" {
        for required in ["alt", "fontfamily", "index"] {
            if element.attribute(None, required).is_none_or(str::is_empty) {
                return Err(invalid(format!(
                    "MathML mglyph requires a nonempty {required} attribute"
                )));
            }
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
    validate_presentation_sequence(element)?;
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

fn validate_exact_math_expressions(element: &Element, expected: usize) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    if children.len() == expected && children.iter().all(|child| is_math_expression(child)) {
        Ok(())
    } else {
        Err(invalid(format!(
            "MathML {} requires exactly {expected} MathML expression children",
            element.local_name()
        )))
    }
}

fn validate_presentation_sequence(element: &Element) -> Result<()> {
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
    validate_presentation_sequence(element)?;
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
    if !is_math_expression(first) {
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

fn validate_bound_variable(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    if children.is_empty() || children.len() > 2 || !is_identifier_token(children[0]) {
        return Err(invalid(
            "MathML bvar requires one content expression and an optional degree",
        ));
    }
    if children.len() == 2
        && (children[1].namespace_uri() != Some(MATHML_NAMESPACE)
            || children[1].local_name() != "degree")
    {
        return Err(invalid("MathML bvar second child must be degree"));
    }
    Ok(())
}

fn validate_content_expression_sequence(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    for child in element.children() {
        if !is_content_expression(child) {
            return Err(invalid(format!(
                "MathML {} cannot contain content child {}",
                element.local_name(),
                child.local_name()
            )));
        }
    }
    Ok(())
}

fn validate_content_token(element: &Element) -> Result<()> {
    let mut separators = 0_usize;
    for content in element.content() {
        let valid = match content {
            Content::Text(_) => true,
            Content::Element(child) => {
                let is_separator = element.local_name() == "cn"
                    && child.namespace_uri() == Some(MATHML_NAMESPACE)
                    && child.local_name() == "sep";
                separators += usize::from(is_separator);
                is_separator || is_expression(child)
            },
        };
        if !valid {
            return Err(invalid(format!(
                "MathML content token {} has an invalid child",
                element.local_name()
            )));
        }
    }
    if element.local_name() == "cn" {
        let kind = element.attribute(None, "type").unwrap_or("real");
        let expected = usize::from(matches!(
            kind,
            "e-notation" | "rational" | "complex-cartesian" | "complex-polar"
        ));
        if separators != expected {
            return Err(invalid(format!(
                "MathML cn type {kind} requires exactly {expected} sep children"
            )));
        }
    }
    Ok(())
}

fn validate_declare(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    if matches!(children.len(), 1 | 2) && children.iter().all(|child| is_content_expression(child))
    {
        Ok(())
    } else {
        Err(invalid(
            "MathML declare requires one content object and an optional content value",
        ))
    }
}

fn validate_condition(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    if children.len() == 1
        && children[0].namespace_uri() == Some(MATHML_NAMESPACE)
        && matches!(
            children[0].local_name(),
            "apply" | "reln" | "set" | "true" | "false"
        )
    {
        Ok(())
    } else {
        Err(invalid(
            "MathML condition requires one predicate expression",
        ))
    }
}

fn validate_interval(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    if children.len() == 2 && children.iter().all(|child| is_content_expression(child)) {
        return Ok(());
    }
    let bound_end = consume_bound_variables(&children, 0);
    if bound_end > 0
        && children.len() == bound_end.saturating_add(1)
        && is_named(children.get(bound_end).copied(), "condition")
    {
        Ok(())
    } else {
        Err(invalid(
            "MathML interval requires two endpoints or bound variables followed by one condition",
        ))
    }
}

fn validate_exact_content_expressions(element: &Element, expected: usize) -> Result<()> {
    validate_content_expression_sequence(element)?;
    let actual = element.children().count();
    if actual == expected {
        Ok(())
    } else {
        Err(invalid(format!(
            "MathML {} requires exactly {expected} content children, found {actual}",
            element.local_name()
        )))
    }
}

fn validate_lambda(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    let mut index = consume_bound_variables(&children, 0);
    index = consume_domain_qualifier(&children, index);
    if children.len() != index.saturating_add(1)
        || !children
            .get(index)
            .is_some_and(|child| is_content_expression(child))
    {
        return Err(invalid(
            "MathML lambda permits bound variables, one optional domain qualifier, and one body expression",
        ));
    }
    Ok(())
}

fn validate_list_or_set(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    if children.iter().all(|child| is_content_expression(child)) {
        return Ok(());
    }
    let mut index = consume_bound_variables(&children, 0);
    let domain_end = consume_domain_qualifier(&children, index);
    if domain_end == index {
        return Err(invalid(format!(
            "MathML {} generated form requires a domain qualifier",
            element.local_name()
        )));
    }
    index = domain_end;
    validate_single_trailing_argument(element.local_name(), &children, index)
}

fn validate_vector(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    if children.is_empty() {
        return Err(invalid("MathML vector requires at least one entry"));
    }
    if children.iter().all(|child| is_content_expression(child)) {
        return Ok(());
    }
    let mut index = consume_bound_variables(&children, 0);
    index = consume_domain_qualifier(&children, index);
    validate_single_trailing_argument("vector", &children, index)
}

fn validate_math_sequence(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let mut argument_seen = false;
    for child in element.children() {
        if child.namespace_uri() == Some(MATHML_NAMESPACE) && child.local_name() == "declare" {
            if argument_seen {
                return Err(invalid(
                    "MathML declare elements must precede math arguments",
                ));
            }
            continue;
        }
        argument_seen = true;
        if !is_math_expression(child) {
            return Err(invalid(format!(
                "MathML math cannot contain {}",
                child.local_name()
            )));
        }
    }
    Ok(())
}

fn validate_application(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    let Some(operator) = children.first() else {
        return Err(invalid("MathML apply requires an operator"));
    };
    if operator.namespace_uri() != Some(MATHML_NAMESPACE) {
        return Err(invalid(
            "MathML apply operator must be in the MathML namespace",
        ));
    }
    let name = operator.local_name();
    let arguments = &children[1..];
    if matches!(name, "curl" | "divergence" | "grad" | "laplacian") {
        return validate_bound_application(name, arguments, false);
    }
    if is_relation_operator(name) {
        return validate_relation_arguments(name, arguments);
    }
    if is_unary_operator(name) {
        return validate_exact_arguments(name, arguments, 1);
    }
    if name == "minus" {
        return validate_argument_range(name, arguments, 1, 2);
    }
    if is_binary_operator(name) {
        return validate_exact_arguments(name, arguments, 2);
    }
    if is_nary_operator(name) {
        return validate_enumerated_or_bound(name, arguments);
    }
    match name {
        "diff" | "partialdiff" => validate_bound_application(name, arguments, false),
        "int" => {
            if arguments.len() == 1 && is_content_expression(arguments[0]) {
                Ok(())
            } else {
                validate_bound_application(name, arguments, true)
            }
        },
        "sum" | "product" => validate_enumerated_or_bound(name, arguments),
        "log" => validate_optional_qualifier(name, arguments, "logbase"),
        "root" => validate_optional_qualifier(name, arguments, "degree"),
        "moment" => validate_moment(arguments),
        "limit" => validate_limit(arguments),
        "forall" | "exists" => validate_bound_application(name, arguments, true),
        local_name if is_special_application_operator(local_name) => {
            validate_enumerated_or_bound(local_name, arguments)
        },
        _ => Err(invalid(format!("MathML apply has invalid operator {name}"))),
    }
}

fn validate_matrix(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    if children.iter().all(|child| {
        child.namespace_uri() == Some(MATHML_NAMESPACE) && child.local_name() == "matrixrow"
    }) {
        let Some(first) = children.first() else {
            return Err(invalid(
                "MathML enumerated matrix requires at least one row",
            ));
        };
        let width = first.children().count();
        if width == 0 || children.iter().any(|row| row.children().count() != width) {
            return Err(invalid(
                "MathML enumerated matrix rows must be nonempty and rectangular",
            ));
        }
        return Ok(());
    }
    let mut index = consume_bound_variables(&children, 0);
    index = consume_domain_qualifier(&children, index);
    validate_single_trailing_argument("matrix", &children, index)
}

fn validate_matrix_row(element: &Element) -> Result<()> {
    validate_content_expression_sequence(element)?;
    if element.children().next().is_some() {
        Ok(())
    } else {
        Err(invalid("MathML matrixrow requires at least one cell"))
    }
}

fn validate_relation(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    let Some(relation) = children.first() else {
        return Err(invalid("MathML reln requires a relation operator"));
    };
    if relation.namespace_uri() != Some(MATHML_NAMESPACE)
        || !is_relation_operator(relation.local_name())
    {
        return Err(invalid("MathML reln has an invalid relation operator"));
    }
    validate_relation_arguments(relation.local_name(), &children[1..])
}

fn validate_relation_arguments(name: &str, arguments: &[&Element]) -> Result<()> {
    let mut index = consume_bound_variables(arguments, 0);
    if is_named(arguments.get(index).copied(), "condition") {
        index += 1;
    }
    let values = &arguments[index..];
    if is_binary_relation(name) {
        validate_exact_arguments(name, values, 2)
    } else {
        validate_argument_sequence(name, values)
    }
}

fn validate_bound_application(
    name: &str,
    arguments: &[&Element],
    allow_domain: bool,
) -> Result<()> {
    let mut index = consume_bound_variables(arguments, 0);
    if allow_domain {
        index = consume_domain_qualifier(arguments, index);
    }
    validate_single_trailing_argument(name, arguments, index)
}

fn validate_limit(arguments: &[&Element]) -> Result<()> {
    let mut index = consume_bound_variables(arguments, 0);
    if is_named(arguments.get(index).copied(), "lowlimit") {
        index += 1;
    }
    if is_named(arguments.get(index).copied(), "condition") {
        index += 1;
    }
    validate_single_trailing_argument("limit", arguments, index)
}

fn validate_moment(arguments: &[&Element]) -> Result<()> {
    let mut index = usize::from(is_named(arguments.first().copied(), "degree"));
    if is_named(arguments.get(index).copied(), "momentabout") {
        index += 1;
    }
    validate_argument_sequence("moment", &arguments[index..])
}

fn validate_optional_qualifier(name: &str, arguments: &[&Element], qualifier: &str) -> Result<()> {
    let index = usize::from(is_named(arguments.first().copied(), qualifier));
    validate_single_trailing_argument(name, arguments, index)
}

fn validate_single_trailing_argument(
    name: &str,
    arguments: &[&Element],
    index: usize,
) -> Result<()> {
    if arguments.len() == index.saturating_add(1)
        && arguments
            .get(index)
            .is_some_and(|child| is_content_expression(child))
    {
        Ok(())
    } else {
        Err(invalid(format!(
            "MathML {name} requires one trailing content argument"
        )))
    }
}

fn validate_exact_arguments(name: &str, arguments: &[&Element], expected: usize) -> Result<()> {
    validate_argument_range(name, arguments, expected, expected)
}

fn validate_enumerated_or_bound(name: &str, arguments: &[&Element]) -> Result<()> {
    if arguments.iter().all(|child| is_content_expression(child)) {
        Ok(())
    } else {
        validate_bound_application(name, arguments, true)
    }
}

fn validate_argument_range(
    name: &str,
    arguments: &[&Element],
    minimum: usize,
    maximum: usize,
) -> Result<()> {
    if (minimum..=maximum).contains(&arguments.len())
        && arguments.iter().all(|child| is_content_expression(child))
    {
        Ok(())
    } else {
        Err(invalid(format!(
            "MathML {name} requires {minimum}..={maximum} content arguments"
        )))
    }
}

fn validate_argument_sequence(name: &str, arguments: &[&Element]) -> Result<()> {
    if arguments.iter().all(|child| is_content_expression(child)) {
        Ok(())
    } else {
        Err(invalid(format!(
            "MathML {name} has a non-expression argument"
        )))
    }
}

fn consume_bound_variables(arguments: &[&Element], mut index: usize) -> usize {
    while is_named(arguments.get(index).copied(), "bvar") {
        index += 1;
    }
    index
}

fn consume_domain_qualifier(arguments: &[&Element], index: usize) -> usize {
    if is_named(arguments.get(index).copied(), "lowlimit") {
        return index + 1 + usize::from(is_named(arguments.get(index + 1).copied(), "uplimit"));
    }
    if arguments.get(index).is_some_and(|element| {
        element.namespace_uri() == Some(MATHML_NAMESPACE)
            && matches!(
                element.local_name(),
                "condition" | "domainofapplication" | "interval" | "uplimit"
            )
    }) {
        index + 1
    } else {
        index
    }
}

fn is_named(candidate: Option<&Element>, name: &str) -> bool {
    candidate.is_some_and(|element| {
        element.namespace_uri() == Some(MATHML_NAMESPACE) && element.local_name() == name
    })
}

fn validate_piecewise(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let mut otherwise_seen = false;
    for child in element.children() {
        if child.namespace_uri() != Some(MATHML_NAMESPACE) {
            return Err(invalid("MathML piecewise has a foreign child"));
        }
        match child.local_name() {
            "piece" if !otherwise_seen => {},
            "otherwise" if !otherwise_seen => otherwise_seen = true,
            _ => return Err(invalid("MathML piecewise child order is invalid")),
        }
    }
    Ok(())
}

fn is_boolean(value: &str) -> bool {
    matches!(value, "true" | "false")
}

fn is_nonnegative_integer(value: &str) -> bool {
    !value.is_empty() && value.bytes().all(|byte| byte.is_ascii_digit())
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

fn is_content_expression(element: &Element) -> bool {
    if element.namespace_uri() != Some(MATHML_NAMESPACE) {
        return false;
    }
    matches!(
        element.local_name(),
        "apply"
            | "reln"
            | "ci"
            | "cn"
            | "csymbol"
            | "fn"
            | "interval"
            | "lambda"
            | "list"
            | "matrix"
            | "piecewise"
            | "semantics"
            | "set"
            | "vector"
    ) || is_content_symbol(element.local_name())
}

fn is_identifier_token(element: &Element) -> bool {
    if element.namespace_uri() != Some(MATHML_NAMESPACE) {
        return false;
    }
    if element.local_name() == "ci" {
        return true;
    }
    if element.local_name() != "semantics" {
        return false;
    }
    element.children().next().is_some_and(|child| {
        child.namespace_uri() == Some(MATHML_NAMESPACE) && child.local_name() == "ci"
    })
}

fn is_binary_operator(local_name: &str) -> bool {
    matches!(
        local_name,
        "approx"
            | "divide"
            | "equivalent"
            | "factorof"
            | "implies"
            | "outerproduct"
            | "power"
            | "quotient"
            | "rem"
            | "scalarproduct"
            | "setdiff"
            | "vectorproduct"
    )
}

fn is_binary_relation(local_name: &str) -> bool {
    matches!(
        local_name,
        "in" | "neq" | "notin" | "notprsubset" | "notsubset" | "tendsto"
    )
}

fn is_nary_operator(local_name: &str) -> bool {
    matches!(
        local_name,
        "and"
            | "cartesianproduct"
            | "compose"
            | "gcd"
            | "intersect"
            | "lcm"
            | "max"
            | "mean"
            | "median"
            | "min"
            | "mode"
            | "or"
            | "plus"
            | "sdev"
            | "selector"
            | "times"
            | "union"
            | "variance"
            | "xor"
    )
}

fn is_relation_operator(local_name: &str) -> bool {
    is_binary_relation(local_name)
        || matches!(
            local_name,
            "eq" | "geq" | "gt" | "leq" | "lt" | "prsubset" | "subset"
        )
}

fn is_special_application_operator(local_name: &str) -> bool {
    matches!(
        local_name,
        "apply" | "ci" | "csymbol" | "fn" | "lambda" | "reln" | "semantics"
    )
}

fn is_unary_operator(local_name: &str) -> bool {
    matches!(
        local_name,
        "abs"
            | "arccos"
            | "arccosh"
            | "arccot"
            | "arccoth"
            | "arccsc"
            | "arccsch"
            | "arcsec"
            | "arcsech"
            | "arcsin"
            | "arcsinh"
            | "arctan"
            | "arctanh"
            | "arg"
            | "card"
            | "ceiling"
            | "codomain"
            | "conjugate"
            | "cos"
            | "cosh"
            | "cot"
            | "coth"
            | "csc"
            | "csch"
            | "curl"
            | "determinant"
            | "divergence"
            | "domain"
            | "exp"
            | "factorial"
            | "floor"
            | "grad"
            | "ident"
            | "image"
            | "imaginary"
            | "inverse"
            | "laplacian"
            | "ln"
            | "not"
            | "real"
            | "sec"
            | "sech"
            | "sin"
            | "sinh"
            | "tan"
            | "tanh"
            | "transpose"
    )
}

fn is_content_symbol(local_name: &str) -> bool {
    matches!(
        local_name,
        "abs"
            | "and"
            | "approx"
            | "arccos"
            | "arccosh"
            | "arccot"
            | "arccoth"
            | "arccsc"
            | "arccsch"
            | "arcsec"
            | "arcsech"
            | "arcsin"
            | "arcsinh"
            | "arctan"
            | "arctanh"
            | "arg"
            | "card"
            | "cartesianproduct"
            | "ceiling"
            | "complexes"
            | "compose"
            | "conjugate"
            | "codomain"
            | "cos"
            | "cosh"
            | "cot"
            | "coth"
            | "csc"
            | "csch"
            | "curl"
            | "determinant"
            | "diff"
            | "divergence"
            | "divide"
            | "domain"
            | "emptyset"
            | "eq"
            | "equivalent"
            | "eulergamma"
            | "exists"
            | "exp"
            | "exponentiale"
            | "factorial"
            | "factorof"
            | "false"
            | "floor"
            | "forall"
            | "gcd"
            | "geq"
            | "grad"
            | "gt"
            | "ident"
            | "image"
            | "imaginary"
            | "imaginaryi"
            | "implies"
            | "in"
            | "infinity"
            | "integers"
            | "intersect"
            | "int"
            | "inverse"
            | "laplacian"
            | "lcm"
            | "leq"
            | "limit"
            | "ln"
            | "log"
            | "lt"
            | "max"
            | "mean"
            | "median"
            | "min"
            | "minus"
            | "mode"
            | "moment"
            | "naturalnumbers"
            | "neq"
            | "not"
            | "notanumber"
            | "notin"
            | "notprsubset"
            | "notsubset"
            | "or"
            | "outerproduct"
            | "partialdiff"
            | "pi"
            | "plus"
            | "power"
            | "primes"
            | "product"
            | "prsubset"
            | "quotient"
            | "rationals"
            | "real"
            | "reals"
            | "rem"
            | "root"
            | "scalarproduct"
            | "sdev"
            | "sec"
            | "sech"
            | "selector"
            | "setdiff"
            | "sin"
            | "sinh"
            | "subset"
            | "sum"
            | "tan"
            | "tanh"
            | "tendsto"
            | "times"
            | "transpose"
            | "true"
            | "union"
            | "variance"
            | "vectorproduct"
            | "xor"
    )
}

fn is_math_expression(element: &Element) -> bool {
    is_expression(element) || is_content_expression(element)
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

fn validate_type_attribute(local_name: &str, value: &str) -> bool {
    match local_name {
        "cn" => value_in(
            value,
            &[
                "integer",
                "real",
                "e-notation",
                "rational",
                "complex-cartesian",
                "complex-polar",
                "constant",
            ],
        ),
        "tendsto" => value_in(value, &["above", "below", "two-sided"]),
        _ => !value.is_empty(),
    }
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
