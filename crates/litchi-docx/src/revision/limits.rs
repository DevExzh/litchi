//! Finite resource policy for tracked-revision readback.

use crate::{Error, Result};

/// Resource ceilings for one paragraph, table, row, or cell revision read.
///
/// Nested annotations charge each retained copy of their affected text.
/// Limits apply to semantic ownership and parser work, not exact allocator RSS.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum source XML bytes.
    pub max_source_bytes: usize,
    /// Maximum XML events, including text and closing elements.
    pub max_events: usize,
    /// Maximum XML element depth; the hard safety ceiling is 128.
    pub max_depth: usize,
    /// Maximum recognized revision records.
    pub max_revisions: usize,
    /// Maximum aggregate decoded author, identifier, and timestamp bytes.
    pub max_metadata_bytes: usize,
    /// Maximum aggregate retained revision text bytes, including nested copies.
    pub max_text_bytes: usize,
    /// Maximum encoded bytes of one metadata attribute or entity reference.
    pub max_value_bytes: usize,
    /// Maximum attributes on one revision element.
    pub max_attributes: usize,
    /// Maximum namespace bindings inherited from an enclosing source owner.
    pub max_inherited_namespaces: usize,
    /// Maximum aggregate inherited prefix and namespace URI bytes.
    pub max_inherited_namespace_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_source_bytes: 64 * 1024 * 1024,
            max_events: 1_000_000,
            max_depth: 128,
            max_revisions: 100_000,
            max_metadata_bytes: 8 * 1024 * 1024,
            max_text_bytes: 64 * 1024 * 1024,
            max_value_bytes: 64 * 1024,
            max_attributes: 256,
            max_inherited_namespaces: 4096,
            max_inherited_namespace_bytes: 4 * 1024 * 1024,
        }
    }
}

impl Limits {
    /// Validate finite limits against hard safety ceilings.
    ///
    /// Zero quotas permit only sources that consume none of that resource.
    ///
    /// # Errors
    /// Returns a counted revision limit error for a ceiling above its hard cap.
    pub fn validate(self) -> Result<Self> {
        for (resource, value, maximum) in [
            ("source bytes", self.max_source_bytes, 256 * 1024 * 1024),
            ("events", self.max_events, 16_000_000),
            ("depth", self.max_depth, 128),
            ("records", self.max_revisions, 1_000_000),
            ("metadata bytes", self.max_metadata_bytes, 64 * 1024 * 1024),
            ("text bytes", self.max_text_bytes, 256 * 1024 * 1024),
            ("value bytes", self.max_value_bytes, 1024 * 1024),
            ("attributes", self.max_attributes, 4096),
            (
                "inherited namespaces",
                self.max_inherited_namespaces,
                65_536,
            ),
            (
                "inherited namespace bytes",
                self.max_inherited_namespace_bytes,
                16 * 1024 * 1024,
            ),
        ] {
            check(resource, value, maximum)?;
        }
        Ok(self)
    }
}

pub(super) fn check(resource: &'static str, actual: usize, maximum: usize) -> Result<()> {
    if actual > maximum {
        Err(Error::RevisionLimit {
            resource,
            actual,
            maximum,
        })
    } else {
        Ok(())
    }
}

pub(super) fn charge(
    resource: &'static str,
    used: &mut usize,
    additional: usize,
    maximum: usize,
) -> Result<()> {
    let actual = used.checked_add(additional).ok_or(Error::RevisionLimit {
        resource,
        actual: usize::MAX,
        maximum,
    })?;
    check(resource, actual, maximum)?;
    *used = actual;
    Ok(())
}
