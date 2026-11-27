//! Source-checked linked-section edits: the `link` family.
//!
//! [`Package::edit_link`](crate::Package::edit_link) stages an [`Edit`] of
//! one `text:section-source` target; [`Edit::commit`] publishes a new
//! [`Package`] with its [`Patch`] (which records the
//! [`Change`]) and [`Diagnostics`];
//! [`Package::apply_link`](crate::Package::apply_link) replays a patch onto
//! its exact source.

use litchi_core::{Error as CoreError, Position, Result};
use std::{borrow::Cow, ops::Range, sync::Arc};

use crate::Package;

type FamilyResult<T> = std::result::Result<T, Error>;

/// Whether a commit changed the package bytes.
pub use litchi_core::edit::Diagnostics;
/// The family error: typed limits, refusals and patch conflicts plus the
/// common-core payload.
pub use litchi_odf_common::Error;
/// The dimension named by [`Error::Limit`] and [`Error::InvalidLimit`].
pub use litchi_odf_common::LimitKind;

/// A committed link edit: the published package, its patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

const MAX_HREF_BYTES: usize = 16 * 1024;
const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;

/// A semantic selector for an existing linked section.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Selector<'a> {
    /// Selects the unique containing `text:section` by its exact name.
    Section(Cow<'a, str>),
    /// Selects a link by its checked zero-based semantic position.
    Position(Position),
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Section(Cow::Borrowed(value))
    }
}

impl From<String> for Selector<'_> {
    fn from(value: String) -> Self {
        Self::Section(Cow::Owned(value))
    }
}

impl From<Position> for Selector<'_> {
    fn from(value: Position) -> Self {
        Self::Position(value)
    }
}

/// An isolated edit of one existing `text:section-source` link.
pub struct Edit<'source> {
    source: &'source Package,
    reference: Position,
    span: Range<usize>,
    before: String,
    after: String,
}

impl<'source> Edit<'source> {
    pub(crate) fn new(source: &'source Package, selector: Selector<'_>) -> Result<Self> {
        let reference = resolve(source, selector)?;
        let selected = source
            .subdocuments()
            .get(reference.get())
            .ok_or_else(|| invalid("ODM subdocument selector is out of bounds"))?;
        let span = source
            .href_span(reference.get())
            .cloned()
            .ok_or_else(|| invalid("ODM subdocument source span is missing"))?;
        Ok(Self {
            source,
            reference,
            span,
            before: selected.href().to_owned(),
            after: selected.href().to_owned(),
        })
    }

    /// Returns the checked zero-based reference position resolved by this edit.
    #[must_use]
    pub const fn reference(&self) -> Position {
        self.reference
    }

    /// Returns the inert target staged for publication.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.after
    }

    /// Replaces the target of the selected linked section.
    ///
    /// The target stays inert: publication does not open, resolve, fetch, or
    /// execute it.
    ///
    /// # Errors
    ///
    /// Returns an error when the target exceeds the bounded input size or is
    /// not representable as XML 1.0 character data.
    pub fn set_href(&mut self, value: impl Into<String>) -> FamilyResult<()> {
        let href = value.into();
        validate_href(&href)?;
        self.after = href;
        Ok(())
    }

    /// Validates, reparses, and atomically publishes the staged link change.
    ///
    /// # Errors
    ///
    /// Returns an error when the compact source cannot be patched losslessly,
    /// the package is signed, or semantic readback differs from the request.
    pub fn commit(self) -> std::result::Result<Commit, Error> {
        Ok(self.commit_core()?)
    }

    fn commit_core(self) -> Result<Commit> {
        if self.before == self.after {
            let snapshot = self.source.clone();
            return Ok(commit(
                self.source,
                snapshot,
                Change::new(self.reference, self.before, self.after),
            ));
        }
        let content = replace_attribute_value(self.source.content_xml(), &self.span, &self.after)?;
        let snapshot = self.source.with_content_xml(&content)?;
        let actual = snapshot
            .subdocuments()
            .get(self.reference.get())
            .ok_or_else(|| invalid("ODM edited subdocument disappeared during readback"))?;
        if actual.href() != self.after {
            return Err(invalid(
                "ODM linked-section transaction readback differs from the staged target",
            ));
        }
        Ok(commit(
            self.source,
            snapshot,
            Change::new(self.reference, self.before, self.after),
        ))
    }
}

/// The semantic effect of one linked-section patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Change {
    reference: Position,
    before: String,
    after: String,
}

impl Change {
    fn new(reference: Position, before: String, after: String) -> Self {
        Self {
            reference,
            before,
            after,
        }
    }

    /// Returns the checked zero-based reference position.
    #[must_use]
    pub const fn reference(&self) -> Position {
        self.reference
    }

    /// Returns the source target.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Returns the published target.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

fn commit(source: &Package, snapshot: Package, change: Change) -> Commit {
    let patch = Patch {
        before: source.shared_bytes(),
        after: snapshot.shared_bytes(),
        change,
    };
    let diagnostics = Diagnostics::new(!patch.is_noop());
    Commit::new(snapshot, patch, diagnostics)
}

/// An exact-source-checked reversible linked-section patch.
#[derive(Clone)]
pub struct Patch {
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
    change: Change,
}

impl Patch {
    /// Returns the semantic link change.
    #[must_use]
    pub const fn change(&self) -> &Change {
        &self.change
    }

    /// Returns whether the patch applies to this exact source artifact.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Package) -> bool {
        source.as_bytes() == self.before.as_slice()
    }

    /// Applies this patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied source differs byte-for-byte.
    pub(crate) fn apply(&self, source: &Package) -> std::result::Result<Commit, Error> {
        if !self.is_applicable_to(source) {
            return Err(Error::patch_conflict(
                "link.apply",
                "the package is not the patch source",
            ));
        }
        let package = Package::from_shared_bytes(Arc::clone(&self.after))?;
        Ok(Commit::new(
            package,
            self.clone(),
            Diagnostics::new(!self.is_noop()),
        ))
    }

    /// Returns the patch that restores the exact source package bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            change: Change::new(
                self.change.reference,
                self.change.after.clone(),
                self.change.before.clone(),
            ),
        }
    }

    /// Returns whether this patch leaves the source bytes unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_slice() == self.after.as_slice()
    }
}

pub(crate) fn replace_attribute_value(
    source: &str,
    span: &Range<usize>,
    value: &str,
) -> Result<String> {
    if span.start > span.end || span.end > source.len() {
        return Err(invalid("ODM subdocument source span is invalid"));
    }
    let replacement = quick_xml::escape::escape(value);
    let capacity = source
        .len()
        .checked_sub(span.end - span.start)
        .and_then(|size| size.checked_add(replacement.len()))
        .ok_or_else(|| invalid("ODM edited content size overflow"))?;
    if capacity > MAX_CONTENT_BYTES {
        return Err(invalid("ODM edited content exceeds the output limit"));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation| CoreError::Allocation {
            resource: "ODM edited content",
            source: allocation,
        })?;
    output.push_str(&source[..span.start]);
    output.push_str(&replacement);
    output.push_str(&source[span.end..]);
    Ok(output)
}

fn resolve(source: &Package, selector: Selector<'_>) -> Result<Position> {
    match selector {
        Selector::Position(position) => source
            .subdocuments()
            .get(position.get())
            .map(|_| position)
            .ok_or_else(|| invalid("ODM subdocument selector is out of bounds")),
        Selector::Section(name) => source
            .subdocuments()
            .iter()
            .position(|reference| reference.section() == name.as_ref())
            .map(Position::new)
            .ok_or_else(|| invalid("ODM linked section name was not found")),
    }
}

pub(crate) fn validate_href(href: &str) -> Result<()> {
    if href.len() > MAX_HREF_BYTES {
        return Err(invalid("ODM subdocument target exceeds the 16 KiB limit"));
    }
    if href.chars().any(|value| {
        !matches!(
            value,
            '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(invalid(
            "ODM subdocument target contains a character forbidden by XML 1.0",
        ));
    }
    Ok(())
}

fn invalid(message: &str) -> CoreError {
    CoreError::InvalidFormat(message.to_owned())
}
