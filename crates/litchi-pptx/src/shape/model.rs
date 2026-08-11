//! Borrowed semantic shape views.

use std::ops::Range;

use crate::{Error, Result};

/// A checked byte span in the markup-compatibility-processed owner XML.
///
/// Spans are compact so a large scene can retain its index without copying
/// individual shape subtrees. Use [`Scene::xml`](super::Scene::xml) to obtain
/// the owner against which this span is defined.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Span {
    pub(super) start: u32,
    pub(super) len: u32,
}

impl Span {
    /// Byte offset from the beginning of the processed owner XML.
    #[inline]
    #[must_use]
    pub const fn start(self) -> u32 {
        self.start
    }

    /// Length of this element, including its start and end tags.
    #[inline]
    #[must_use]
    pub const fn len(self) -> u32 {
        self.len
    }

    /// Whether the element occupies no bytes.
    #[inline]
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.len == 0
    }

    /// Exclusive byte end, checked for arithmetic overflow.
    #[inline]
    #[must_use]
    pub const fn end(self) -> Option<u32> {
        self.start.checked_add(self.len)
    }

    pub(crate) fn range(self, owner_len: usize) -> Result<Range<usize>> {
        let start = usize::try_from(self.start)
            .map_err(|_err| Error::Invalid("shape offset does not fit usize".into()))?;
        let len = usize::try_from(self.len)
            .map_err(|_err| Error::Invalid("shape length does not fit usize".into()))?;
        let end = start
            .checked_add(len)
            .ok_or_else(|| Error::Invalid("shape byte span overflows usize".into()))?;
        if end > owner_len {
            return Err(Error::Invalid(
                "shape byte span is outside its processed owner XML".into(),
            ));
        }
        Ok(start..end)
    }
}

/// Local shape bounds in English Metric Units (EMUs).
///
/// A value is exposed only when one complete offset/extent pair is present;
/// inherited layout transforms are deliberately not guessed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Bounds {
    x: i64,
    y: i64,
    width: i64,
    height: i64,
}

impl Bounds {
    #[inline]
    pub(super) const fn new(x: i64, y: i64, width: i64, height: i64) -> Self {
        Self {
            x,
            y,
            width,
            height,
        }
    }

    /// Left edge in EMUs.
    #[inline]
    #[must_use]
    pub const fn x(self) -> i64 {
        self.x
    }

    /// Top edge in EMUs.
    #[inline]
    #[must_use]
    pub const fn y(self) -> i64 {
        self.y
    }

    /// Width in EMUs.
    #[inline]
    #[must_use]
    pub const fn width(self) -> i64 {
        self.width
    }

    /// Height in EMUs.
    #[inline]
    #[must_use]
    pub const fn height(self) -> i64 {
        self.height
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct TextSpan {
    pub(super) start: u32,
    pub(super) len: u32,
}

impl TextSpan {
    pub(super) fn get(self, arena: &str) -> Option<&str> {
        let start = self.start as usize;
        let len = self.len as usize;
        let end = start.checked_add(len)?;
        arena.get(start..end)
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct PlaceholderRecord {
    pub(super) kind: Option<TextSpan>,
    pub(super) index: u32,
}

/// Placeholder metadata declared by one shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Placeholder<'a> {
    kind: Option<&'a str>,
    index: u32,
}

impl<'a> Placeholder<'a> {
    /// Producer token such as `title`, `body`, or `ctrTitle`.
    ///
    /// `None` preserves the schema-defaulted spelling instead of inventing a
    /// token that was not present in the owner XML.
    #[inline]
    #[must_use]
    pub const fn kind(self) -> Option<&'a str> {
        self.kind
    }

    /// Placeholder index. OOXML defaults an omitted index to zero.
    #[inline]
    #[must_use]
    pub const fn index(self) -> u32 {
        self.index
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Kind {
    Auto,
    Picture,
    Table,
    Chart,
    Diagram,
    Ole,
    Frame,
    Group,
    Connector,
    Content,
    Unknown,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct Record {
    pub(super) span: Span,
    pub(super) subtree_end: u32,
    pub(super) parent: Option<u32>,
    pub(super) kind: Kind,
    pub(super) name: Option<TextSpan>,
    pub(super) id: Option<u32>,
    pub(super) bounds: Option<Bounds>,
    pub(super) placeholder: Option<PlaceholderRecord>,
    pub(super) text: Option<TextSpan>,
    pub(super) source_name: Option<TextSpan>,
}

/// Borrowed common properties shared by every semantic shape variant.
#[derive(Debug, Clone, Copy)]
pub struct Common<'a> {
    pub(super) xml: &'a [u8],
    pub(super) records: &'a [Record],
    pub(super) strings: &'a str,
    pub(super) record: &'a Record,
    pub(super) index: usize,
}

impl<'a> Common<'a> {
    /// Zero-based pre-order position in the owning scene.
    #[inline]
    #[must_use]
    pub const fn index(self) -> usize {
        self.index
    }

    /// Checked span in the processed owner XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn span(self) -> Result<Span> {
        Ok(self.record.span)
    }

    /// Exact borrowed element XML without a per-shape copy.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn xml(self) -> Result<&'a [u8]> {
        let range = self.record.span.range(self.xml.len())?;
        self.xml
            .get(range)
            .ok_or_else(|| Error::Invalid("shape XML span became invalid".into()))
    }

    /// Decoded producer name from `cNvPr`.
    #[must_use]
    pub fn name(self) -> Option<&'a str> {
        self.record.name?.get(self.strings)
    }

    /// Numeric non-visual ID used by low-level animation references.
    ///
    /// Semantic selectors should normally use [`name`](Self::name) or a
    /// checked pre-order position instead.
    #[must_use]
    pub fn id(self) -> Option<u32> {
        self.record.id
    }

    /// Complete local transform bounds when explicitly present.
    #[must_use]
    pub fn bounds(self) -> Option<Bounds> {
        self.record.bounds
    }

    /// Placeholder metadata when declared by the shape.
    #[must_use]
    pub fn placeholder(self) -> Option<Placeholder<'a>> {
        let value = self.record.placeholder?;
        Some(Placeholder {
            kind: value.kind.and_then(|span| span.get(self.strings)),
            index: value.index,
        })
    }

    /// Decoded `DrawingML` text captured in full-owner namespace context.
    ///
    /// `None` distinguishes shapes with no `DrawingML` text from an explicitly
    /// empty text body.
    #[must_use]
    pub fn text(self) -> Option<&'a str> {
        self.record.text?.get(self.strings)
    }

    /// Qualified source element name, useful for an unknown extension shape.
    #[must_use]
    pub fn source_name(self) -> Option<&'a str> {
        self.record.source_name?.get(self.strings)
    }
}

macro_rules! leaf {
    ($(#[$meta:meta])* $name:ident) => {
        $(#[$meta])*
        #[derive(Debug, Clone, Copy)]
        pub struct $name<'a>(pub(super) Common<'a>);

        impl<'a> $name<'a> {
            /// Borrow the properties common to all shape kinds.
            #[inline]
            pub const fn common(self) -> Common<'a> {
                self.0
            }
        }
    };
}

leaf!(/// An ordinary `PresentationML` auto shape (`p:sp`).
    Auto);
leaf!(/// A picture (`p:pic`).
    Picture);
leaf!(/// A graphic frame whose payload is a `DrawingML` table.
    Table);
leaf!(/// A graphic frame whose payload is a `DrawingML` chart.
    Chart);
leaf!(/// A graphic frame whose payload is a `DrawingML` diagram.
    Diagram);
leaf!(/// A graphic frame whose payload is an embedded or linked OLE object.
    Ole);
leaf!(/// A graphic frame with an unclassified or extension payload.
    Frame);
leaf!(/// A connector (`p:cxnSp`).
    Connector);
leaf!(/// A Microsoft content-part shape.
    Content);
leaf!(/// A retained shape-like extension not understood by this release.
    Unknown);

/// A grouped shape with direct hierarchical child access.
#[derive(Debug, Clone, Copy)]
pub struct Group<'a>(pub(super) Common<'a>);

impl<'a> Group<'a> {
    /// Borrow the properties common to all shape kinds.
    #[inline]
    #[must_use]
    pub const fn common(self) -> Common<'a> {
        self.0
    }

    /// Iterate direct children in source order.
    #[must_use]
    pub fn shapes(self) -> Shapes<'a> {
        let first = self.0.index.saturating_add(1);
        let end = self.0.record.subtree_end as usize;
        Shapes {
            xml: self.0.xml,
            records: self.0.records,
            strings: self.0.strings,
            cursor: first,
            end,
            parent: Some(self.0.index as u32),
            preorder: false,
        }
    }
}

/// One semantic `PresentationML` shape.
///
/// The enum carries a typed view rather than requiring users to compare a
/// separate numeric or discriminator value.
#[derive(Debug, Clone, Copy)]
#[non_exhaustive]
pub enum Shape<'a> {
    Auto(Auto<'a>),
    Picture(Picture<'a>),
    Table(Table<'a>),
    Chart(Chart<'a>),
    Diagram(Diagram<'a>),
    Ole(Ole<'a>),
    Frame(Frame<'a>),
    Group(Group<'a>),
    Connector(Connector<'a>),
    Content(Content<'a>),
    Unknown(Unknown<'a>),
}

impl<'a> Shape<'a> {
    pub(super) const fn from_common(common: Common<'a>) -> Self {
        match common.record.kind {
            Kind::Auto => Self::Auto(Auto(common)),
            Kind::Picture => Self::Picture(Picture(common)),
            Kind::Table => Self::Table(Table(common)),
            Kind::Chart => Self::Chart(Chart(common)),
            Kind::Diagram => Self::Diagram(Diagram(common)),
            Kind::Ole => Self::Ole(Ole(common)),
            Kind::Frame => Self::Frame(Frame(common)),
            Kind::Group => Self::Group(Group(common)),
            Kind::Connector => Self::Connector(Connector(common)),
            Kind::Content => Self::Content(Content(common)),
            Kind::Unknown => Self::Unknown(Unknown(common)),
        }
    }

    /// Borrow the properties common to every variant.
    #[must_use]
    pub const fn common(self) -> Common<'a> {
        match self {
            Self::Auto(value) => value.0,
            Self::Picture(value) => value.0,
            Self::Table(value) => value.0,
            Self::Chart(value) => value.0,
            Self::Diagram(value) => value.0,
            Self::Ole(value) => value.0,
            Self::Frame(value) => value.0,
            Self::Group(value) => value.0,
            Self::Connector(value) => value.0,
            Self::Content(value) => value.0,
            Self::Unknown(value) => value.0,
        }
    }

    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn span(self) -> Result<Span> {
        self.common().span()
    }

    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn xml(self) -> Result<&'a [u8]> {
        self.common().xml()
    }

    #[inline]
    #[must_use]
    pub fn name(self) -> Option<&'a str> {
        self.common().name()
    }

    #[inline]
    #[must_use]
    pub fn id(self) -> Option<u32> {
        self.common().id()
    }

    #[inline]
    #[must_use]
    pub fn bounds(self) -> Option<Bounds> {
        self.common().bounds()
    }

    #[inline]
    #[must_use]
    pub fn placeholder(self) -> Option<Placeholder<'a>> {
        self.common().placeholder()
    }

    #[inline]
    #[must_use]
    pub fn text(self) -> Option<&'a str> {
        self.common().text()
    }
}

/// A borrowed shape iterator.
#[derive(Debug, Clone)]
pub struct Shapes<'a> {
    pub(super) xml: &'a [u8],
    pub(super) records: &'a [Record],
    pub(super) strings: &'a str,
    pub(super) cursor: usize,
    pub(super) end: usize,
    pub(super) parent: Option<u32>,
    pub(super) preorder: bool,
}

impl<'a> Iterator for Shapes<'a> {
    type Item = Shape<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        while self.cursor < self.end {
            let index = self.cursor;
            let record = self.records.get(index)?;
            self.cursor = if self.preorder {
                index.checked_add(1)?
            } else {
                (record.subtree_end as usize).max(index.checked_add(1)?)
            };
            if self.preorder || record.parent == self.parent {
                let common = Common {
                    xml: self.xml,
                    records: self.records,
                    strings: self.strings,
                    record,
                    index,
                };
                return Some(Shape::from_common(common));
            }
        }
        None
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.end.saturating_sub(self.cursor);
        if self.preorder {
            (remaining, Some(remaining))
        } else {
            (0, Some(remaining))
        }
    }
}

impl std::iter::FusedIterator for Shapes<'_> {}
