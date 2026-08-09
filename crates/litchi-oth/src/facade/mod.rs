//! Concise family entry points.

use litchi_core::{Error, Metadata, Position, Result};
use std::path::Path;

pub use crate::authoring::Builder;

const MAX_PARAGRAPH_BYTES: usize = 16 * 1024 * 1024;

/// A read-only semantic text-web body projection.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextBody {
    paragraphs: Vec<crate::paragraph::Paragraph>,
}

impl TextBody {
    /// Returns projected paragraph character data in document order.
    #[must_use]
    pub fn paragraphs(&self) -> &[crate::paragraph::Paragraph] {
        &self.paragraphs
    }
}

/// Immutable document snapshot.
#[derive(Clone)]
pub struct Template {
    package: crate::package::Snapshot,
}

impl Template {
    /// Opens a web-template package from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        crate::package::Snapshot::open(path).map(|package| Self { package })
    }

    /// Opens a web-template package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not a valid package.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        crate::package::Snapshot::from_bytes(bytes).map(|package| Self { package })
    }

    /// Returns the `content.xml` document.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    /// Returns the `styles.xml` document, if present.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    /// Returns the document metadata, if present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
    }

    /// Returns the raw package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Lists the file entries stored in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be enumerated.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    /// Projects inert paragraph character data from the validated text body.
    ///
    /// Fields, links, scripts, forms, resources, and embedded objects are not
    /// evaluated, followed, activated, or otherwise executed.
    pub fn text_body(&self) -> Result<TextBody> {
        Ok(TextBody {
            paragraphs: self.package.paragraphs().to_vec(),
        })
    }

    /// Starts a source-bound text-body transaction.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            source: self,
            change: None,
        }
    }

    /// Consumes the snapshot and returns the raw package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }
}

/// A source-bound, one-operation web-template text edit.
pub struct Edit<'a> {
    source: &'a Template,
    change: Option<ParagraphChange>,
}

impl Edit<'_> {
    /// Replaces one paragraph's sole direct character-data XML span.
    ///
    /// Nested, split, CDATA, and empty paragraphs remain readable but are not
    /// rewritten, because recreating their markup could lose unknown content.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is invalid, the paragraph has no
    /// lossless replacement span, or another paragraph is already staged.
    pub fn set_paragraph_text(
        &mut self,
        paragraph: Position,
        text: impl Into<String>,
    ) -> Result<()> {
        let after = text.into();
        if after.len() > MAX_PARAGRAPH_BYTES {
            return Err(Error::InvalidFormat(
                "OTH replacement paragraph text exceeds the limit".to_string(),
            ));
        }
        if self
            .change
            .as_ref()
            .is_some_and(|change| change.paragraph != paragraph)
        {
            return Err(Error::InvalidFormat(
                "an OTH transaction supports one paragraph-text edit".to_string(),
            ));
        }
        let before = self
            .source
            .package
            .paragraphs()
            .get(paragraph.get())
            .ok_or_else(|| {
                Error::InvalidFormat("OTH paragraph selector is out of bounds".to_string())
            })?
            .text()
            .to_owned();
        if before == after {
            self.change = None;
            return Ok(());
        }
        if self
            .source
            .package
            .replacement_site(paragraph.get())
            .is_none()
        {
            return Err(Error::InvalidFormat(
                "OTH paragraph is not one losslessly replaceable XML text span".to_string(),
            ));
        }
        self.change = Some(ParagraphChange {
            paragraph,
            before,
            after,
        });
        Ok(())
    }

    /// Atomically validates, publishes, and records this text edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the source cannot be losslessly rewritten or the
    /// fully reopened candidate fails semantic readback.
    pub fn commit(self) -> Result<Commit> {
        let Some(change) = self.change else {
            return Ok(Commit::unchanged(self.source.clone()));
        };
        let span = self
            .source
            .package
            .replacement_site(change.paragraph.get())
            .ok_or_else(|| {
                Error::InvalidFormat("OTH paragraph edit site disappeared".to_string())
            })?;
        let content = replace_text(self.source.content_xml(), span, &change.after)?;
        let snapshot = Template {
            package: self.source.package.rebuild_with_content(&content)?,
        };
        let actual = snapshot
            .package
            .paragraphs()
            .get(change.paragraph.get())
            .ok_or_else(|| Error::InvalidFormat("OTH edited paragraph disappeared".to_string()))?;
        if actual.text() != change.after {
            return Err(Error::InvalidFormat(
                "OTH package edit failed semantic readback".to_string(),
            ));
        }
        Ok(Commit {
            snapshot: snapshot.clone(),
            patch: Patch {
                source: self.source.clone(),
                target: snapshot,
                change: Some(change),
            },
            changed: true,
        })
    }
}

/// One reversible semantic paragraph-text operation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ParagraphChange {
    paragraph: Position,
    before: String,
    after: String,
}

impl ParagraphChange {
    /// The zero-based source-order paragraph position.
    #[must_use]
    pub const fn paragraph(&self) -> Position {
        self.paragraph
    }

    /// The text expected before applying the patch.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// The replacement text.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// A committed immutable template and its exact-source patch.
pub struct Commit {
    snapshot: Template,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn unchanged(snapshot: Template) -> Self {
        Self {
            patch: Patch {
                source: snapshot.clone(),
                target: snapshot.clone(),
                change: None,
            },
            snapshot,
            changed: false,
        }
    }

    /// Whether the committed package differs from its source snapshot.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Returns the committed immutable template snapshot.
    #[must_use]
    pub fn template(&self) -> &Template {
        &self.snapshot
    }

    /// Returns the source-checked reversible patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit and returns the published template.
    #[must_use]
    pub fn into_template(self) -> Template {
        self.snapshot
    }
}

/// A source-checked reversible OTH paragraph-text patch.
#[derive(Clone)]
pub struct Patch {
    source: Template,
    target: Template,
    change: Option<ParagraphChange>,
}

impl Patch {
    /// Returns whether this patch authorizes the supplied exact source bytes.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Template) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Applies this patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied source differs byte-for-byte.
    pub fn apply(&self, source: &Template) -> Result<Template> {
        if !self.is_applicable_to(source) {
            return Err(Error::InvalidFormat(
                "OTH patch source does not match its expected snapshot".to_string(),
            ));
        }
        Ok(self.target.clone())
    }

    /// Returns the semantic change, if this is not an exact no-op patch.
    #[must_use]
    pub fn change(&self) -> Option<&ParagraphChange> {
        self.change.as_ref()
    }

    /// Returns the patch that restores the exact source snapshot.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            change: self.change.as_ref().map(|change| ParagraphChange {
                paragraph: change.paragraph,
                before: change.after.clone(),
                after: change.before.clone(),
            }),
        }
    }
}

fn replace_text(source: &str, span: &std::ops::Range<usize>, replacement: &str) -> Result<String> {
    if span.start > span.end || span.end > source.len() {
        return Err(Error::InvalidFormat(
            "OTH paragraph source span is invalid".to_string(),
        ));
    }
    let replacement = quick_xml::escape::escape(replacement);
    let capacity = source
        .len()
        .checked_sub(span.end - span.start)
        .and_then(|size| size.checked_add(replacement.len()))
        .ok_or_else(|| Error::InvalidFormat("OTH edited content size overflow".to_string()))?;
    if capacity > MAX_PARAGRAPH_BYTES.saturating_mul(16) {
        return Err(Error::InvalidFormat(
            "OTH edited content exceeds the output limit".to_string(),
        ));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "OTH edited content",
            source,
        })?;
    output.push_str(&source[..span.start]);
    output.push_str(&replacement);
    output.push_str(&source[span.end..]);
    Ok(output)
}
