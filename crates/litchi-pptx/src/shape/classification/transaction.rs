//! Detached classification edits with atomic commit semantics.

use super::model::{Outcome, Snapshot};
use super::validation::{validate_outcome, validate_snapshot};
use crate::Result;

use super::codec::{self, Source};

/// A detached classification edit.
///
/// The source snapshot is never mutated. Callers can inspect the projected
/// snapshot during the edit and consume the editor with [`Editor::commit`].
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

    /// Replace the typed outcome while retaining every opaque extension.
    pub fn set(&mut self, outcome: Outcome) -> Result<()> {
        validate_outcome(outcome)?;
        self.working.outcome = Some(outcome);
        Ok(())
    }

    /// Clear only the typed outcome, retaining opaque extensions for a later
    /// package-level decision.
    pub fn clear(&mut self) {
        self.working.outcome = None;
    }

    /// Borrow the projected snapshot.
    #[inline]
    pub fn snapshot(&self) -> &Snapshot {
        &self.working
    }

    /// Whether the edit changes the semantic snapshot.
    #[inline]
    pub fn is_changed(&self) -> bool {
        self.original != self.working
    }

    /// Validate and consume the edit into a new snapshot.
    pub fn commit(self) -> Result<Snapshot> {
        validate_snapshot(&self.working)?;
        Ok(self.working)
    }
}

/// Stage a typed outcome into the selected raw shape while preserving every
/// unrelated extension and every untouched byte range.
pub(super) fn set(xml: &[u8], source: &Source, outcome: Outcome) -> Result<Vec<u8>> {
    validate_outcome(outcome)?;
    let generated = codec::write_classification(outcome);
    if let Some(known) = &source.layout.known {
        if let Some(element) = &known.classification {
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
        let extension = codec::write_extension(source.layout.conformance, outcome)?;
        if ext_lst.element.empty {
            return expand_empty(xml, &ext_lst.element, &extension);
        }
        return codec::replace(
            xml,
            ext_lst.element.close_start..ext_lst.element.close_start,
            &extension,
        );
    }

    let list = codec::write_extension_list(source.layout.conformance, outcome)?;
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

/// Remove the typed classification element. Unknown extension entries remain
/// in their original source order and byte spelling.
pub(super) fn remove(xml: &[u8], source: &Source) -> Result<Option<Vec<u8>>> {
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
    Ok(Some(codec::replace(xml, replacement, &[])?))
}

fn expand_empty(xml: &[u8], element: &codec::Element, child: &[u8]) -> Result<Vec<u8>> {
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
    codec::replace(xml, element.span.clone(), &replacement)
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
