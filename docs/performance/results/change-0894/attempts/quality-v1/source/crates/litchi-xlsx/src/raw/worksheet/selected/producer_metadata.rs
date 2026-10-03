//! Validators for the small worksheet-metadata subset admitted by the
//! selected-cell stream.
//!
//! The selected route does not materialize layout, view, or print metadata.
//! It may therefore admit only fields whose lexical form is independently
//! checked and whose presence cannot change the selected cell values.  An
//! unknown or repairable field is reported to the caller as an ineligible
//! worksheet, which leaves the authoritative worksheet parser in charge.

use litchi_ooxml_common::mce::SemanticElement;
use litchi_sheet::{COLUMNS, ROWS, Rect};

const RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_SQREF_TOKENS: usize = 64;
const MAX_RELATIONSHIP_ID_BYTES: usize = 4096;

/// Validate a direct worksheet `sheetViews` container.
pub(super) fn sheet_views(element: &SemanticElement<'_>) -> bool {
    attributes(element, &[])
}

/// Validate the bounded `sheetView` subset used by the selected route.
pub(super) fn sheet_view(element: &SemanticElement<'_>) -> bool {
    attributes(element, &["workbookViewId"])
        && required(element, "workbookViewId").is_some_and(|value| parse_u32(value).is_some())
}

/// Validate one inert selection projection.
pub(super) fn selection(element: &SemanticElement<'_>) -> bool {
    if !attributes(element, &["activeCell", "sqref"]) {
        return false;
    }
    let active = element_attribute(element, "activeCell");
    let sqref = element_attribute(element, "sqref");
    if active.is_none() && sqref.is_none() {
        return false;
    }
    if let Some(active) = active {
        if parse_a1(active).is_err() {
            return false;
        }
    }
    if let Some(sqref) = sqref {
        let mut count = 0usize;
        for token in sqref.split_ascii_whitespace() {
            count = count.saturating_add(1);
            if count > MAX_SQREF_TOKENS || Rect::from_a1(token).is_err() {
                return false;
            }
        }
        if count == 0 {
            return false;
        }
    }
    true
}

/// Validate the modeled, finite `sheetFormatPr` values.  The x14ac descent
/// attribute is checked by the raw extension observer and is admitted here
/// only at its recognized namespace and local name.
pub(super) fn sheet_format(element: &SemanticElement<'_>) -> bool {
    const ALLOWED: &[&str] = &[
        "baseColWidth",
        "defaultColWidth",
        "defaultRowHeight",
        "customHeight",
        "zeroHeight",
        "thickTop",
        "thickBottom",
        "outlineLevelRow",
        "outlineLevelCol",
    ];
    if !attributes_with_x14ac(element, ALLOWED) {
        return false;
    }
    let Some(default_row_height) = required(element, "defaultRowHeight") else {
        return false;
    };
    if !parse_f64(default_row_height, |value| {
        crate::layout::Height::new(value).is_ok()
    }) {
        return false;
    }
    if let Some(value) = element_attribute(element, "defaultColWidth")
        && !parse_f64(value, |value| crate::layout::Width::new(value).is_ok())
    {
        return false;
    }
    if let Some(value) = element_attribute(element, "baseColWidth")
        && !parse_u32(value).is_some_and(|value| value <= 255)
    {
        return false;
    }
    for name in ["outlineLevelRow", "outlineLevelCol"] {
        if let Some(value) = element_attribute(element, name)
            && !parse_u32(value).is_some_and(|value| value <= 7)
        {
            return false;
        }
    }
    for name in ["customHeight", "zeroHeight", "thickTop", "thickBottom"] {
        if let Some(value) = element_attribute(element, name)
            && !parse_bool(value)
        {
            return false;
        }
    }
    true
}

/// Validate the direct `cols` container.  Its records are checked separately
/// because non-overlap and source order are scanner state, not tag-local facts.
pub(super) fn cols(element: &SemanticElement<'_>) -> bool {
    attributes(element, &[])
}

/// Validate a style-free column record and return its inclusive maximum.
pub(super) fn col(element: &SemanticElement<'_>, previous_max: u32) -> Option<u32> {
    const ALLOWED: &[&str] = &[
        "min",
        "max",
        "width",
        "hidden",
        "bestFit",
        "customWidth",
        "phonetic",
        "outlineLevel",
        "collapsed",
    ];
    if !attributes(element, ALLOWED) {
        return None;
    }
    // A style-bearing column changes the effective cell presentation and is
    // deliberately outside this value-only selected route.
    if element_attribute(element, "style").is_some() {
        return None;
    }
    let min = required(element, "min").and_then(parse_u32)?;
    let max = required(element, "max").and_then(parse_u32)?;
    if min == 0 || max > COLUMNS || min > max || min <= previous_max {
        return None;
    }
    if let Some(value) = element_attribute(element, "width")
        && !parse_f64(value, |value| crate::column::Width::new(value).is_ok())
    {
        return None;
    }
    if let Some(value) = element_attribute(element, "outlineLevel")
        && !parse_u32(value).is_some_and(|value| value <= 7)
    {
        return None;
    }
    for name in ["hidden", "bestFit", "customWidth", "phonetic", "collapsed"] {
        if let Some(value) = element_attribute(element, name)
            && !parse_bool(value)
        {
            return None;
        }
    }
    Some(max)
}

/// Validate a complete, direct page-margin record.
pub(super) fn page_margins(element: &SemanticElement<'_>) -> bool {
    const NAMES: &[&str] = &["left", "right", "top", "bottom", "header", "footer"];
    if !attributes(element, NAMES) {
        return false;
    }
    NAMES.iter().all(|name| {
        required(element, name).is_some_and(|value| {
            parse_f64(value, |value| {
                value.is_finite() && (0.0..49.0).contains(&value)
            })
        })
    })
}

/// Validate a bounded, value-safe page-setup subset.  Other valid print
/// settings remain authoritative-parser territory because the selected route
/// does not materialize their semantics.
pub(super) fn page_setup(element: &SemanticElement<'_>) -> bool {
    const ALLOWED: &[&str] = &["paperSize", "orientation"];
    if !attributes_with_relationship_id(element, ALLOWED) {
        return false;
    }
    if let Some(value) = element_attribute(element, "paperSize")
        && !parse_u32(value).is_some_and(|value| crate::page_setup::Paper::try_from(value).is_ok())
    {
        return false;
    }
    if let Some(value) = element_attribute(element, "orientation")
        && !matches!(value, "default" | "portrait" | "landscape")
    {
        return false;
    }
    if let Some(value) = relationship_id(element) {
        if value.is_empty()
            || value.len() > MAX_RELATIONSHIP_ID_BYTES
            || !litchi_ooxml_common::xml::is_ncname(value)
        {
            return false;
        }
    }
    true
}

/// Validate a row's optional, advisory `spans` range.
pub(super) fn row_spans(value: &str) -> bool {
    let Some((first, last)) = value.split_once(':') else {
        return false;
    };
    let Some(first) = parse_u32(first) else {
        return false;
    };
    let Some(last) = parse_u32(last) else {
        return false;
    };
    (1..=COLUMNS).contains(&first) && (first..=COLUMNS).contains(&last)
}

fn attributes(element: &SemanticElement<'_>, allowed: &[&str]) -> bool {
    attributes_inner(element, allowed, false, false)
}

fn attributes_with_x14ac(element: &SemanticElement<'_>, allowed: &[&str]) -> bool {
    attributes_inner(element, allowed, true, false)
}

fn attributes_with_relationship_id(element: &SemanticElement<'_>, allowed: &[&str]) -> bool {
    attributes_inner(element, allowed, false, true)
}

fn attributes_inner(
    element: &SemanticElement<'_>,
    allowed: &[&str],
    allow_x14ac: bool,
    allow_relationship_id: bool,
) -> bool {
    let mut seen = 0u128;
    let mut relationship_seen = false;
    for attribute in &element.attributes {
        let namespace = attribute.expanded_name.namespace.as_bytes();
        let local = attribute.expanded_name.local_name.as_str();
        if namespace.is_empty() {
            let Some(index) = allowed.iter().position(|name| *name == local) else {
                return false;
            };
            let Some(bit) = 1u128.checked_shl(index as u32) else {
                return false;
            };
            if seen & bit != 0 {
                return false;
            }
            seen |= bit;
        } else if allow_x14ac
            && namespace == b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
            && local == "dyDescent"
        {
            // The raw x14ac observer owns the lexical/value check.
        } else if allow_relationship_id
            && local == "id"
            && (namespace == RELATIONSHIP_NAMESPACE || namespace == STRICT_RELATIONSHIP_NAMESPACE)
        {
            if relationship_seen {
                return false;
            }
            relationship_seen = true;
            // The relationship identifier is validated below.
        } else {
            return false;
        }
    }
    true
}

fn relationship_id<'a>(element: &'a SemanticElement<'_>) -> Option<&'a str> {
    element.attributes.iter().find_map(|attribute| {
        let namespace = attribute.expanded_name.namespace.as_bytes();
        (attribute.expanded_name.local_name == "id"
            && (namespace == RELATIONSHIP_NAMESPACE || namespace == STRICT_RELATIONSHIP_NAMESPACE))
            .then_some(attribute.decoded_value.as_ref())
    })
}

fn element_attribute<'a>(element: &'a SemanticElement<'_>, name: &str) -> Option<&'a str> {
    element
        .attributes
        .iter()
        .find(|attribute| {
            attribute.expanded_name.namespace.is_empty()
                && attribute.expanded_name.local_name == name
        })
        .map(|attribute| attribute.decoded_value.as_ref())
}

fn required<'a>(element: &'a SemanticElement<'_>, name: &str) -> Option<&'a str> {
    element_attribute(element, name)
}

fn parse_u32(value: &str) -> Option<u32> {
    (!value.is_empty()).then(|| value.parse().ok()).flatten()
}

fn parse_f64(value: &str, valid: impl FnOnce(f64) -> bool) -> bool {
    value.parse::<f64>().ok().is_some_and(valid)
}

fn parse_bool(value: &str) -> bool {
    matches!(value, "0" | "1" | "true" | "false")
}

fn parse_a1(value: &str) -> Result<(u32, u32), ()> {
    let split = value
        .as_bytes()
        .iter()
        .position(u8::is_ascii_digit)
        .ok_or(())?;
    if split == 0 || split == value.len() {
        return Err(());
    }
    let column = value[..split]
        .bytes()
        .try_fold(0u32, |column, byte| {
            byte.is_ascii_alphabetic()
                .then(|| {
                    column
                        .checked_mul(26)?
                        .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
                })
                .flatten()
        })
        .ok_or(())?;
    if !(1..=COLUMNS).contains(&column) {
        return Err(());
    }
    let row = value[split..].parse::<u32>().map_err(|_| ())?;
    if (1..=ROWS).contains(&row) {
        Ok((row, column))
    } else {
        Err(())
    }
}
