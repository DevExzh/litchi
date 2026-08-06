//! Transaction-ready XML edits for a shape-owned `p:tags` anchor.
//!
//! This owner contains only the local PresentationML mutations around the
//! anchor. Relationship graph validation and package staging remain in the
//! parent codec, while shape discovery remains in the semantic scanner.

use super::model::{Anchor, Container, Element, Layout};
use crate::Result;
use crate::tag::{allocation, invalid, replace_xml};

pub(super) fn add_anchor(xml: &[u8], layout: &Layout, relationship_id: &str) -> Result<Vec<u8>> {
    if layout.anchor.is_some() {
        return Err(invalid("selected shape already has a p:tags anchor"));
    }
    let anchor = format!(
        "<p:tags xmlns:p=\"{}\" xmlns:r=\"{}\" r:id=\"{}\"/>",
        layout.conformance.namespace(),
        layout.conformance.relationship_namespace(),
        relationship_id
    );

    if let Some(container) = &layout.container {
        if !container.element.empty {
            return replace_xml(
                xml,
                container.element.close_start..container.element.close_start,
                anchor.as_bytes(),
            );
        }
        return expand_empty(xml, &container.element, anchor.as_bytes());
    }

    let container = format!(
        "<p:custDataLst xmlns:p=\"{}\">{anchor}</p:custDataLst>",
        layout.conformance.namespace()
    );
    if layout.nv_pr.empty {
        expand_empty(xml, &layout.nv_pr, container.as_bytes())
    } else {
        replace_xml(
            xml,
            layout.insertion..layout.insertion,
            container.as_bytes(),
        )
    }
}

pub(super) fn replace_anchor_id(
    xml: &[u8],
    layout: &Layout,
    relationship_id: &str,
) -> Result<Vec<u8>> {
    let anchor = layout
        .anchor
        .as_ref()
        .ok_or_else(|| invalid("selected shape has no p:tags anchor"))?;
    replace_xml(xml, anchor.id_value.clone(), relationship_id.as_bytes())
}

pub(super) fn remove_anchor(xml: &[u8], layout: &Layout) -> Result<Vec<u8>> {
    let anchor = layout
        .anchor
        .as_ref()
        .ok_or_else(|| invalid("selected shape has no p:tags anchor"))?;
    let container = layout
        .container
        .as_ref()
        .ok_or_else(|| invalid("shape p:tags has no p:custDataLst parent"))?;
    if container.child_elements == 1
        && !container.preserve_when_empty
        && container_contains_only_anchor(xml, container, anchor)?
    {
        replace_xml(xml, container.element.span.clone(), &[])
    } else {
        replace_xml(xml, anchor.span.clone(), &[])
    }
}

fn expand_empty(xml: &[u8], element: &Element, child: &[u8]) -> Result<Vec<u8>> {
    let raw = xml
        .get(element.span.clone())
        .ok_or_else(|| invalid("empty shape element span is outside owner XML"))?;
    let slash = raw
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| invalid("empty shape element has no closing slash"))?;
    let qualified_name = element_qname(raw)?;
    let len = raw
        .len()
        .checked_sub(1)
        .and_then(|value| value.checked_add(child.len()))
        .and_then(|value| value.checked_add(qualified_name.len()))
        .and_then(|value| value.checked_add(3))
        .ok_or_else(|| invalid("expanded shape element length overflow"))?;
    let mut replacement = Vec::new();
    replacement
        .try_reserve_exact(len)
        .map_err(|source| allocation("expanded shape-tag element", source))?;
    replacement.extend_from_slice(&raw[..slash]);
    replacement.extend_from_slice(&raw[slash + 1..]);
    replacement.extend_from_slice(child);
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(qualified_name);
    replacement.push(b'>');
    replace_xml(xml, element.span.clone(), &replacement)
}

fn element_qname(element: &[u8]) -> Result<&[u8]> {
    if element.first() != Some(&b'<') {
        return Err(invalid("shape element does not start with '<'"));
    }
    let start = 1usize;
    let end = element[start..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(*byte, b'/' | b'>'))
        .map(|offset| start + offset)
        .ok_or_else(|| invalid("shape element name is unterminated"))?;
    element
        .get(start..end)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid("shape element name is empty"))
}

fn container_contains_only_anchor(
    xml: &[u8],
    container: &Container,
    anchor: &Anchor,
) -> Result<bool> {
    if container.element.empty {
        return Ok(false);
    }
    let before = xml
        .get(container.element.open_end..anchor.span.start)
        .ok_or_else(|| invalid("shape customer-data prefix range is invalid"))?;
    let after = xml
        .get(anchor.span.end..container.element.close_start)
        .ok_or_else(|| invalid("shape customer-data suffix range is invalid"))?;
    Ok(before.iter().all(u8::is_ascii_whitespace) && after.iter().all(u8::is_ascii_whitespace))
}
