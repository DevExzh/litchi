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
        "apply" | "reln" => validate_nonempty_application(element),
        "ci" | "cn" | "csymbol" => validate_content_token(element),
        "interval" | "piece" => validate_exact_content_expressions(element, 2),
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
        "list" | "matrixrow" | "set" | "vector" => validate_content_expression_sequence(element),
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
            "rowspan" | "columnspan" | "selection" | "base" => is_positive_integer(value),
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
            "rational" | "complex-cartesian" | "complex-polar"
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
    let valid_head = children.first().is_some_and(|child| {
        child.namespace_uri() == Some(MATHML_NAMESPACE) && child.local_name() == "ci"
    });
    let valid_tail = children.get(1).is_none_or(|child| {
        child.namespace_uri() == Some(MATHML_NAMESPACE)
            && (child.local_name() == "fn" || is_constructor(child.local_name()))
    });
    if matches!(children.len(), 1 | 2) && valid_head && valid_tail {
        Ok(())
    } else {
        Err(invalid(
            "MathML declare requires ci and an optional fn or constructor",
        ))
    }
}

fn validate_condition(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    if children.len() == 1
        && children[0].namespace_uri() == Some(MATHML_NAMESPACE)
        && matches!(children[0].local_name(), "apply" | "reln" | "set")
    {
        Ok(())
    } else {
        Err(invalid(
            "MathML condition requires one apply, reln, or set child",
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
    let mut index = 0_usize;
    while children.get(index).is_some_and(|child| {
        child.namespace_uri() == Some(MATHML_NAMESPACE) && child.local_name() == "bvar"
    }) {
        index += 1;
    }
    if children
        .get(index)
        .is_some_and(|child| is_domain_qualifier(child))
    {
        index += 1;
    }
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

fn validate_nonempty_application(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let mut children = element.children();
    let Some(first) = children.next() else {
        return Err(invalid(format!(
            "MathML {} requires at least one content expression",
            element.local_name()
        )));
    };
    if !is_content_expression(first) {
        return Err(invalid(format!(
            "MathML {} has an invalid operator expression",
            element.local_name()
        )));
    }
    if children.all(|child| is_content_expression(child) || is_application_qualifier(child)) {
        Ok(())
    } else {
        Err(invalid(format!(
            "MathML {} has an invalid argument or qualifier",
            element.local_name()
        )))
    }
}

fn validate_matrix(element: &Element) -> Result<()> {
    validate_no_character_data(element)?;
    let children: Vec<_> = element.children().collect();
    if children.iter().all(|child| {
        child.namespace_uri() == Some(MATHML_NAMESPACE) && child.local_name() == "matrixrow"
    }) {
        return Ok(());
    }
    Err(invalid(
        "MathML matrix currently supports only enumerated matrixrow children",
    ))
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

fn is_application_qualifier(element: &Element) -> bool {
    element.namespace_uri() == Some(MATHML_NAMESPACE)
        && matches!(
            element.local_name(),
            "bvar"
                | "condition"
                | "degree"
                | "domainofapplication"
                | "interval"
                | "logbase"
                | "lowlimit"
                | "momentabout"
                | "uplimit"
        )
}

fn is_constructor(local_name: &str) -> bool {
    matches!(
        local_name,
        "interval"
            | "list"
            | "matrix"
            | "matrixrow"
            | "otherwise"
            | "piece"
            | "piecewise"
            | "set"
            | "vector"
    )
}

fn is_domain_qualifier(element: &Element) -> bool {
    element.namespace_uri() == Some(MATHML_NAMESPACE)
        && matches!(
            element.local_name(),
            "condition" | "domainofapplication" | "interval" | "lowlimit" | "uplimit"
        )
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
