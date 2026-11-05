//! Detached classification edits with atomic commit semantics.

use super::model::Outcome;
use super::validation::validate_outcome;
use crate::Result;

use super::codec::{self, Source};

/// Stage a typed outcome into the selected raw shape while preserving every
/// unrelated extension and every untouched byte range.
pub(super) fn set(
    xml: &[u8],
    source: &Source,
    outcome: Outcome,
    max_bytes: usize,
) -> Result<Vec<u8>> {
    validate_outcome(outcome)?;
    let generated = codec::write_classification(outcome);
    if let Some(known) = &source.layout.known {
        if let Some(element) = &known.classification {
            return codec::replace(xml, element.span.clone(), &generated, max_bytes);
        }
        if known.element.empty {
            return expand_empty(xml, &known.element, &generated, max_bytes);
        }
        return codec::replace(
            xml,
            known.element.close_start..known.element.close_start,
            &generated,
            max_bytes,
        );
    }

    if let Some(ext_lst) = &source.layout.ext_lst {
        let extension = codec::write_extension(source.layout.conformance, outcome)?;
        if ext_lst.element.empty {
            return expand_empty(xml, &ext_lst.element, &extension, max_bytes);
        }
        return codec::replace(
            xml,
            ext_lst.element.close_start..ext_lst.element.close_start,
            &extension,
            max_bytes,
        );
    }

    let list = codec::write_extension_list(source.layout.conformance, outcome)?;
    if source.layout.nv_pr.empty {
        expand_empty(xml, &source.layout.nv_pr, &list, max_bytes)
    } else {
        codec::replace(
            xml,
            source.layout.nv_pr.close_start..source.layout.nv_pr.close_start,
            &list,
            max_bytes,
        )
    }
}

/// Remove the typed classification element. Unknown extension entries remain
/// in their original source order and byte spelling.
pub(super) fn remove(xml: &[u8], source: &Source, max_bytes: usize) -> Result<Option<Vec<u8>>> {
    let Some(known) = &source.layout.known else {
        return Ok(None);
    };
    let Some(classification) = &known.classification else {
        return Ok(None);
    };

    let replacement = if !known.other_content
        && source
            .layout
            .ext_lst
            .as_ref()
            .is_some_and(|list| list.child_elements == 1 && !list.other_content)
    {
        source
            .layout
            .ext_lst
            .as_ref()
            .map(|list| list.element.span.clone())
            .ok_or_else(|| {
                crate::Error::Invalid("classification extension list disappeared".into())
            })?
    } else if !known.other_content {
        known.element.span.clone()
    } else {
        classification.span.clone()
    };
    Ok(Some(codec::replace(xml, replacement, &[], max_bytes)?))
}

fn expand_empty(
    xml: &[u8],
    element: &codec::Element,
    child: &[u8],
    max_bytes: usize,
) -> Result<Vec<u8>> {
    let raw = xml.get(element.span.clone()).ok_or_else(|| {
        crate::Error::Invalid("classification empty element range is invalid".into())
    })?;
    let slash = raw
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| crate::Error::Invalid("classification empty element has no slash".into()))?;
    let qname = qualified_name(raw)?;
    let mut replacement = Vec::new();
    let len = raw
        .len()
        .checked_sub(1)
        .and_then(|value| value.checked_add(child.len()))
        .and_then(|value| value.checked_add(qname.len()))
        .and_then(|value| value.checked_add(3))
        .ok_or_else(|| {
            crate::Error::Invalid("classification empty element length overflow".into())
        })?;
    replacement
        .try_reserve_exact(len)
        .map_err(|source| crate::Error::Allocation {
            resource: "classification empty-element expansion",
            source,
        })?;
    replacement.extend_from_slice(&raw[..slash]);
    replacement.extend_from_slice(&raw[slash + 1..]);
    replacement.extend_from_slice(child);
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(qname);
    replacement.push(b'>');
    codec::replace(xml, element.span.clone(), &replacement, max_bytes)
}

fn qualified_name(element: &[u8]) -> Result<&[u8]> {
    if element.first() != Some(&b'<') {
        return Err(crate::Error::Invalid(
            "classification element does not start with '<'".into(),
        ));
    }
    let end = element[1..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(*byte, b'/' | b'>'))
        .map(|offset| offset + 1)
        .ok_or_else(|| {
            crate::Error::Invalid("classification element name is unterminated".into())
        })?;
    element
        .get(1..end)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| crate::Error::Invalid("classification element name is empty".into()))
}
