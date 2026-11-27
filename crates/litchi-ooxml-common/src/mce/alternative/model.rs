//! Typed snapshots for one `mc:AlternateContent` element.

use std::sync::Arc;

use super::super::Capabilities;

/// Bounds for one retained `AlternateContent` snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum source bytes retained by one snapshot.
    pub bytes: usize,
    /// Maximum nested XML element depth, including `AlternateContent`.
    pub depth: usize,
    /// Maximum XML events scanned while validating the snapshot.
    pub nodes: usize,
    /// Maximum direct `Choice`/`Fallback` branches.
    pub branches: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            bytes: 16 * 1024 * 1024,
            depth: 256,
            nodes: 65_536,
            branches: 1_024,
        }
    }
}

/// A lossless, bounded snapshot of one `AlternateContent` element.
///
/// The source XML is retained exactly once. Branch metadata stores only byte
/// spans and expanded namespace requirements, so inactive branches remain
/// available without allocating a second XML buffer for each branch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Alternatives {
    pub(crate) source: Arc<[u8]>,
    pub(crate) branches: Box<[Stored]>,
}

impl Alternatives {
    /// Borrow the exact source element, including all inactive branches.
    #[must_use]
    pub fn as_xml(&self) -> &[u8] {
        &self.source
    }

    /// Return the number of direct branches in source order.
    #[must_use]
    pub fn len(&self) -> usize {
        self.branches.len()
    }

    /// Return whether the snapshot contains no direct branches.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.branches.is_empty()
    }

    /// Visit direct branches in their source order.
    pub fn branches(&self) -> impl Iterator<Item = Branch<'_>> + '_ {
        self.branches
            .iter()
            .filter_map(|branch| branch.view(&self.source))
    }

    /// Return a checked direct-branch view.
    #[must_use]
    pub fn branch(&self, index: usize) -> Option<Branch<'_>> {
        self.branches
            .get(index)
            .and_then(|branch| branch.view(&self.source))
    }

    /// Select the first supported `Choice`, or the `Fallback` when no choice
    /// is supported. The returned view retains the exact branch bytes.
    #[must_use]
    pub fn select(&self, capabilities: &Capabilities) -> Option<Branch<'_>> {
        self.branches.iter().find_map(|branch| {
            if branch.kind == Kind::Fallback
                || branch
                    .requirements
                    .iter()
                    .all(|namespace| capabilities.understands(namespace))
            {
                branch.view(&self.source)
            } else {
                None
            }
        })
    }
}

/// One direct branch of an `AlternateContent` snapshot.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Branch<'a> {
    /// A `Choice` selected only when every required namespace is understood.
    Choice(Choice<'a>),
    /// The final branch used when no preceding choice is supported.
    Fallback(Fallback<'a>),
}

impl Branch<'_> {
    /// Borrow the exact branch element, including its MCE wrapper.
    #[must_use]
    pub fn as_xml(&self) -> &[u8] {
        match self {
            Self::Choice(choice) => choice.as_xml(),
            Self::Fallback(fallback) => fallback.as_xml(),
        }
    }

    /// Borrow the exact branch content between its opening and closing tags.
    #[must_use]
    pub fn content(&self) -> &[u8] {
        match self {
            Self::Choice(choice) => choice.content(),
            Self::Fallback(fallback) => fallback.content(),
        }
    }
}

/// A typed `Choice` branch view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Choice<'a> {
    xml: &'a [u8],
    content: &'a [u8],
    requirements: &'a [Box<str>],
}

impl Choice<'_> {
    /// Borrow the exact `mc:Choice` element.
    #[must_use]
    pub fn as_xml(&self) -> &[u8] {
        self.xml
    }

    /// Borrow the exact bytes contained by the choice wrapper.
    #[must_use]
    pub fn content(&self) -> &[u8] {
        self.content
    }

    /// Visit the expanded namespace URIs listed by `Requires`.
    pub fn requirements(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.requirements.iter().map(Box::as_ref)
    }
}

/// A typed `Fallback` branch view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Fallback<'a> {
    xml: &'a [u8],
    content: &'a [u8],
}

impl Fallback<'_> {
    /// Borrow the exact `mc:Fallback` element.
    #[must_use]
    pub fn as_xml(&self) -> &[u8] {
        self.xml
    }

    /// Borrow the exact bytes contained by the fallback wrapper.
    #[must_use]
    pub fn content(&self) -> &[u8] {
        self.content
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Kind {
    Choice,
    Fallback,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Span {
    pub(crate) start: usize,
    pub(crate) content_start: usize,
    pub(crate) content_end: usize,
    pub(crate) end: usize,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Stored {
    pub(crate) kind: Kind,
    pub(crate) span: Span,
    pub(crate) requirements: Box<[Box<str>]>,
}

impl Stored {
    pub(crate) fn view<'a>(&'a self, source: &'a [u8]) -> Option<Branch<'a>> {
        let xml = source.get(self.span.start..self.span.end)?;
        let content = source.get(self.span.content_start..self.span.content_end)?;
        Some(match self.kind {
            Kind::Choice => Branch::Choice(Choice {
                xml,
                content,
                requirements: &self.requirements,
            }),
            Kind::Fallback => Branch::Fallback(Fallback { xml, content }),
        })
    }
}
