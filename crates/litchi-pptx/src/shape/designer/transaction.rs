//! Detached and range-based Designer design-element edits.

use super::codec::{self, Source};
use super::validation::validate_value;
use crate::Result;

/// Stage a value into one selected raw shape while preserving unrelated XML.
pub(super) fn set(xml: &[u8], source: &Source, value: bool, max_bytes: usize) -> Result<Vec<u8>> {
    validate_value(value)?;
    let generated = codec::write_design_element(value);
    if let Some(known) = &source.layout.known {
        if let Some(element) = &known.design {
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
        let extension = codec::write_extension(source.layout.conformance, value)?;
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

    let list = codec::write_extension_list(source.layout.conformance, value)?;
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

/// Remove the typed design element while retaining unrelated extension bytes.
pub(super) fn remove(xml: &[u8], source: &Source, max_bytes: usize) -> Result<Option<Vec<u8>>> {
    let Some(known) = &source.layout.known else {
        return Ok(None);
    };
    let Some(design) = &known.design else {
        return Ok(None);
    };

    if !known.other_content
        && source
            .layout
            .ext_lst
            .as_ref()
            .is_some_and(|list| list.child_elements == 1 && !list.other_content)
    {
        return Ok(Some(collapse_empty(xml, &source.layout.nv_pr, max_bytes)?));
    }
    let replacement = if known.other_content {
        design.span.clone()
    } else {
        known.element.span.clone()
    };
    Ok(Some(codec::replace(xml, replacement, &[], max_bytes)?))
}

fn collapse_empty(xml: &[u8], element: &codec::Element, max_bytes: usize) -> Result<Vec<u8>> {
    let raw = xml
        .get(element.span.clone())
        .ok_or_else(|| crate::Error::Invalid("designer element range is invalid".into()))?;
    let end = raw
        .iter()
        .position(|byte| *byte == b'>')
        .ok_or_else(|| crate::Error::Invalid("designer element start tag is invalid".into()))?;
    let mut replacement = Vec::with_capacity(end + 2);
    replacement.extend_from_slice(&raw[..end]);
    replacement.extend_from_slice(b"/>");
    codec::replace(xml, element.span.clone(), &replacement, max_bytes)
}

fn expand_empty(
    xml: &[u8],
    element: &codec::Element,
    child: &[u8],
    max_bytes: usize,
) -> Result<Vec<u8>> {
    let raw = xml
        .get(element.span.clone())
        .ok_or_else(|| crate::Error::Invalid("designer empty element range is invalid".into()))?;
    let slash = raw
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| crate::Error::Invalid("designer empty element has no slash".into()))?;
    let qname = qualified_name(raw)?;
    let mut replacement = Vec::new();
    let len = raw
        .len()
        .checked_sub(1)
        .and_then(|value| value.checked_add(child.len()))
        .and_then(|value| value.checked_add(qname.len()))
        .and_then(|value| value.checked_add(3))
        .ok_or_else(|| crate::Error::Invalid("designer empty-element length overflow".into()))?;
    replacement
        .try_reserve_exact(len)
        .map_err(|source| crate::Error::Allocation {
            resource: "designer empty-element expansion",
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
            "designer element does not start with '<'".into(),
        ));
    }
    let end = element[1..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(*byte, b'/' | b'>'))
        .map(|offset| offset + 1)
        .ok_or_else(|| crate::Error::Invalid("designer element name is unterminated".into()))?;
    element
        .get(1..end)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| crate::Error::Invalid("designer element name is empty".into()))
}
