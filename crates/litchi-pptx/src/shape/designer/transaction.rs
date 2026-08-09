//! Detached and range-based Designer design-element edits.

use super::codec::{self, Source};
use super::model::Snapshot;
use super::validation::{validate_snapshot, validate_value};
use crate::Result;

/// A detached atomic edit of one design-element snapshot.
#[derive(Debug, Clone)]
pub struct Editor {
    original: Snapshot,
    working: Snapshot,
}

impl Editor {
    pub(super) fn new(snapshot: Snapshot) -> Self {
        Self {
            original: snapshot.clone(),
            working: snapshot,
        }
    }

    /// Set the typed boolean while retaining every opaque extension.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set(&mut self, value: bool) -> Result<()> {
        validate_value(value)?;
        self.working.value = Some(value);
        Ok(())
    }

    /// Clear only the typed value; opaque extension entries remain retained.
    pub fn clear(&mut self) {
        self.working.value = None;
    }

    /// Borrow the projected snapshot without mutating the original.
    #[inline]
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.working
    }

    /// Whether this edit changes the detached snapshot.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.original != self.working
    }

    /// Validate and consume the edit into a new snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<Snapshot> {
        validate_snapshot(&self.working)?;
        Ok(self.working)
    }
}

/// Stage a value into one selected raw shape while preserving unrelated XML.
pub(super) fn set(xml: &[u8], source: &Source, value: bool) -> Result<Vec<u8>> {
    validate_value(value)?;
    let generated = codec::write_design_element(value);
    if let Some(known) = &source.layout.known {
        if let Some(element) = &known.design {
            return codec::replace(xml, element.span.clone(), &generated);
        }
        if known.element.empty {
            return expand_empty(xml, &known.element, &generated);
        }
        return codec::replace(
            xml,
            known.element.close_start..known.element.close_start,
            &generated,
        );
    }

    if let Some(ext_lst) = &source.layout.ext_lst {
        let extension = codec::write_extension(source.layout.conformance, value)?;
        if ext_lst.element.empty {
            return expand_empty(xml, &ext_lst.element, &extension);
        }
        return codec::replace(
            xml,
            ext_lst.element.close_start..ext_lst.element.close_start,
            &extension,
        );
    }

    let list = codec::write_extension_list(source.layout.conformance, value)?;
    if source.layout.nv_pr.empty {
        expand_empty(xml, &source.layout.nv_pr, &list)
    } else {
        codec::replace(
            xml,
            source.layout.nv_pr.close_start..source.layout.nv_pr.close_start,
            &list,
        )
    }
}

/// Remove the typed design element while retaining unrelated extension bytes.
pub(super) fn remove(xml: &[u8], source: &Source) -> Result<Option<Vec<u8>>> {
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
        return Ok(Some(collapse_empty(xml, &source.layout.nv_pr)?));
    }
    let replacement = if known.other_content {
        design.span.clone()
    } else {
        known.element.span.clone()
    };
    Ok(Some(codec::replace(xml, replacement, &[])?))
}

fn collapse_empty(xml: &[u8], element: &codec::Element) -> Result<Vec<u8>> {
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
    codec::replace(xml, element.span.clone(), &replacement)
}

fn expand_empty(xml: &[u8], element: &codec::Element, child: &[u8]) -> Result<Vec<u8>> {
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
    codec::replace(xml, element.span.clone(), &replacement)
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
