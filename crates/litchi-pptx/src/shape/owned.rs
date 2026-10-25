//! Independently retained shape scenes.
//!
//! A [`Scene`] owns the processed XML and its index behind one shared allocation.
//! Its [`Handle`]s outlive the reader and do not retain a presentation or package.
//! Lookup uses exact, case-sensitive producer names or checked positions; missing
//! selections are `None` and duplicate names are typed errors.
//!
//! ```
//! use litchi_pptx::shape::owned::Scene;
//! let xml = br#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
//!   <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/></p:nvSpPr></p:sp>
//! </p:spTree>"#;
//! let scene = Scene::read(xml.to_vec(), litchi_pptx::shape::Limits::DEFAULT)?;
//! let title = scene.shape("Title")?.expect("title");
//! drop(scene);
//! assert_eq!(title.view().name(), Some("Title"));
//! assert_eq!(title.position().get(), 0);
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use super::model::{Common, Record};
use super::{Limits, Shape};
use crate::leniency::ToleranceReport;
use litchi_core::Position;
use std::{borrow::Cow, fmt, iter::FusedIterator, sync::Arc};

/// Semantic name or zero-based pre-order position within a scene.
#[derive(Clone, PartialEq, Eq)]
pub enum Selector<'a> {
    /// Exact, case-sensitive producer name.
    Name(Cow<'a, str>),
    /// Checked positional vocabulary shared with other format owners.
    Position(Position),
}
impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(Cow::Borrowed(value))
    }
}
impl From<String> for Selector<'_> {
    fn from(value: String) -> Self {
        Self::Name(Cow::Owned(value))
    }
}
impl From<Position> for Selector<'_> {
    fn from(value: Position) -> Self {
        Self::Position(value)
    }
}
impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Position(Position::new(value))
    }
}
impl fmt::Debug for Selector<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Name(_) => f.write_str("Name(..)"),
            Self::Position(position) => f.debug_tuple("Position").field(position).finish(),
        }
    }
}

/// Content-free scene selection error.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// More than one shape has the requested exact name.
    #[error("shape name is ambiguous ({matches} matches)")]
    Ambiguous { matches: usize },
}

/// A cheap shared snapshot of one parsed shape tree.
///
/// Cloning retains the same immutable index. The original allocation is retained
/// when markup compatibility leaves it unchanged; otherwise only the processed
/// XML is retained. This value has no live source owner or cache.
#[derive(Clone)]
pub struct Scene {
    data: Arc<Data>,
}
struct Data {
    _reservation: Option<Arc<litchi_core::Reservation>>,
    _scene_reservation: Option<Arc<litchi_core::Reservation>>,
    xml: Xml,
    records: Vec<Record>,
    strings: String,
    limits: Limits,
    tolerance: ToleranceReport,
}
enum Xml {
    Source(litchi_opc::PartData),
    Input(Arc<Vec<u8>>),
    Processed(Vec<u8>),
}
impl Xml {
    fn bytes(&self) -> &[u8] {
        match self {
            Self::Source(data) => data.as_bytes(),
            Self::Input(bytes) => bytes,
            Self::Processed(bytes) => bytes,
        }
    }
}
impl Scene {
    /// Index owned XML with caller-selected finite limits.
    ///
    /// # Errors
    /// Returns the same parse and limit failures as [`super::Scene::read`].
    pub fn read(xml: impl Into<Arc<Vec<u8>>>, limits: Limits) -> crate::Result<Self> {
        let input = xml.into();
        let parsed = super::Scene::read(input.as_slice(), limits)?;
        Self::from_parsed(parsed, Arc::clone(&input))
    }
    pub(crate) fn from_parsed(
        parsed: super::Scene<'_>,
        input: Arc<Vec<u8>>,
    ) -> crate::Result<Self> {
        Self::retain(parsed, Xml::Input(input), None)
    }
    pub(crate) fn from_source(
        parsed: super::Scene<'_>,
        input: litchi_opc::PartData,
    ) -> crate::Result<Self> {
        Self::retain(parsed, Xml::Source(input), None)
    }
    pub(crate) fn from_reserved(
        parsed: super::Scene<'_>,
        input: Arc<Vec<u8>>,
        reservation: Arc<litchi_core::Reservation>,
    ) -> crate::Result<Self> {
        Self::retain(parsed, Xml::Input(input), Some(reservation))
    }
    fn retain(
        parsed: super::Scene<'_>,
        input: Xml,
        reservation: Option<Arc<litchi_core::Reservation>>,
    ) -> crate::Result<Self> {
        let (xml, records, strings, limits, tolerance) = parsed.into_parts();
        let processed = match xml {
            Cow::Owned(xml) => Some(xml),
            Cow::Borrowed(bytes) if std::ptr::eq(bytes, input.bytes()) => None,
            Cow::Borrowed(bytes) => {
                let mut copied = Vec::new();
                copied.try_reserve_exact(bytes.len()).map_err(|source| {
                    crate::Error::Allocation {
                        resource: "owned shape scene XML",
                        source,
                    }
                })?;
                copied.extend_from_slice(bytes);
                Some(copied)
            },
        };
        let xml = processed.map_or(input, Xml::Processed);
        Ok(Self {
            data: Arc::new(Data {
                _reservation: reservation,
                _scene_reservation: None,
                xml,
                records,
                strings,
                limits,
                tolerance,
            }),
        })
    }
    /// Attach a read-admission charge before a newly built scene is shared.
    pub(crate) fn with_scene_reservation(
        mut self,
        reservation: Arc<litchi_core::Reservation>,
    ) -> crate::Result<Self> {
        let data = Arc::get_mut(&mut self.data).ok_or_else(|| {
            crate::Error::Invalid("scene admission requires an unpublished owner".into())
        })?;
        data._scene_reservation = Some(reservation);
        Ok(self)
    }

    pub(crate) fn retained_bytes(&self) -> usize {
        let xml = match &self.data.xml {
            Xml::Input(bytes) => bytes.capacity(),
            Xml::Source(data) => data.as_bytes().len(),
            Xml::Processed(bytes) => bytes.capacity(),
        };
        size_of::<Data>()
            .saturating_add(xml)
            .saturating_add(
                self.data
                    .records
                    .capacity()
                    .saturating_mul(size_of::<Record>()),
            )
            .saturating_add(self.data.strings.capacity())
            .saturating_add(self.data.tolerance.heap_bytes())
    }
    pub(crate) fn is_unshared(&self) -> bool {
        Arc::strong_count(&self.data) == 1
    }
    pub(crate) fn ptr_eq(left: &Self, right: &Self) -> bool {
        Arc::ptr_eq(&left.data, &right.data)
    }
    /// Number of shapes in depth-first pre-order.
    #[must_use]
    pub fn len(&self) -> usize {
        self.data.records.len()
    }
    /// Whether the tree contains no shapes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.data.records.is_empty()
    }
    /// Limits used when constructing the immutable scene.
    #[must_use]
    pub fn limits(&self) -> Limits {
        self.data.limits
    }
    /// Whether markup compatibility changed the owner bytes.
    #[must_use]
    pub fn is_rewritten(&self) -> bool {
        matches!(self.data.xml, Xml::Processed(_))
    }
    /// Explicit tolerance report from the existing leaf reader.
    #[must_use]
    pub fn tolerance_report(&self) -> &ToleranceReport {
        &self.data.tolerance
    }
    /// Select a shape by exact producer name or pre-order position.
    /// Missing names and positions, including `usize::MAX`, return `None`.
    ///
    /// # Errors
    /// Returns [`Error::Ambiguous`] when multiple shapes have the requested name.
    pub fn shape<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Option<Handle>, Error> {
        let position = match selector.into() {
            Selector::Position(position) => {
                let index = position.get();
                (index < self.len()).then_some(index)
            },
            Selector::Name(name) => {
                let mut found = None;
                let mut matches = 0;
                for (position, record) in self.data.records.iter().enumerate() {
                    if record.name.and_then(|span| span.get(&self.data.strings))
                        == Some(name.as_ref())
                    {
                        found = Some(position);
                        matches += 1;
                    }
                }
                if matches > 1 {
                    return Err(Error::Ambiguous { matches });
                }
                found
            },
        };
        Ok(position.map(|index| self.handle(index)))
    }
    fn handle(&self, index: usize) -> Handle {
        Handle {
            scene: self.clone(),
            index,
        }
    }
    /// Iterate all shapes in depth-first pre-order; the iterator owns its scene.
    #[must_use]
    pub fn iter(&self) -> Handles {
        Handles {
            scene: self.clone(),
            cursor: 0,
            end: self.len(),
            parent: None,
            preorder: true,
        }
    }
    /// Borrow every shape's typed leaf view in depth-first pre-order, as
    /// [`Handle::view`] would for each item of [`Self::iter`], without
    /// retaining the scene per shape (record 0286).
    ///
    /// ```
    /// use litchi_pptx::shape::owned::Scene;
    ///
    /// let xml = br#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></p:spTree>"#;
    /// let scene = Scene::read(xml.to_vec(), litchi_pptx::shape::Limits::DEFAULT)?;
    /// assert_eq!(scene.views().count(), scene.iter().count());
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    pub fn views(&self) -> impl ExactSizeIterator<Item = Shape<'_>> + '_ {
        let data = &*self.data;
        (0..data.records.len()).map(move |index| {
            Shape::from_common(Common {
                xml: data.xml.bytes(),
                records: &data.records,
                strings: &data.strings,
                limits: data.limits,
                record: &data.records[index],
                index,
            })
        })
    }
    /// Iterate only direct children of the owner shape tree.
    #[must_use]
    pub fn roots(&self) -> Handles {
        Handles {
            preorder: false,
            ..self.iter()
        }
    }
}
impl fmt::Debug for Scene {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Scene")
            .field("shapes", &self.len())
            .field("rewritten", &self.is_rewritten())
            .finish_non_exhaustive()
    }
}

/// A checked position retaining its immutable scene, without an owner borrow.
#[derive(Clone)]
pub struct Handle {
    scene: Scene,
    index: usize,
}
impl Handle {
    /// Zero-based depth-first pre-order position in the owning scene.
    #[must_use]
    pub fn position(&self) -> Position {
        Position::new(self.index)
    }
    /// Borrow the existing typed leaf shape view.
    #[must_use]
    pub fn view(&self) -> Shape<'_> {
        Shape::from_common(Common {
            xml: self.scene.data.xml.bytes(),
            records: &self.scene.data.records,
            strings: &self.scene.data.strings,
            limits: self.scene.data.limits,
            record: &self.scene.data.records[self.index],
            index: self.index,
        })
    }

    /// Borrow the table projection from this graphic-frame shape.
    ///
    /// The projection is indexed against the retained owner XML, so DrawingML
    /// prefixes declared by an ancestor remain resolvable even when the
    /// graphic-frame span itself does not repeat those declarations.
    ///
    /// # Errors
    ///
    /// Returns the table reader's errors when the selected shape does not hold
    /// a readable table or its owner span is invalid.
    pub fn table(&self) -> crate::Result<crate::table::Table<'_>> {
        let owner = self.scene.data.xml.bytes();
        let range = self.scene.data.records[self.index]
            .span
            .range(owner.len())?;
        let xml = owner
            .get(range.clone())
            .ok_or_else(|| crate::Error::Invalid("shape XML span became invalid".into()))?;
        crate::table::Table::from_graphic_frame_with_context(xml, owner, range.start)
    }

    /// The paragraphs and text runs of the cells of the table this
    /// graphic frame holds, as [`crate::table::Table::cell_text_index`]
    /// gives them, from one pass over the frame when it can (record 0348).
    ///
    /// # Errors
    ///
    /// As [`Self::table`] and [`crate::table::Table::cell_text_index`].
    pub fn cell_text_index(&self) -> crate::Result<crate::table::CellTextIndex> {
        let owner = self.scene.data.xml.bytes();
        let range = self.scene.data.records[self.index]
            .span
            .range(owner.len())?;
        let xml = owner
            .get(range.clone())
            .ok_or_else(|| crate::Error::Invalid("shape XML span became invalid".into()))?;
        if let Some(index) =
            crate::table::shape::scan_frame_cell_text_index(xml, owner, range.start)?
        {
            return Ok(index);
        }
        self.table()?.cell_text_index()
    }

    /// Clone the detached scene retained by this handle.
    ///
    /// The returned scene shares the same immutable allocation and carries no
    /// presentation or package owner. This is useful to adapters that need to
    /// enumerate sibling roots for a derived property such as z-order.
    #[must_use]
    pub fn scene(&self) -> Scene {
        self.scene.clone()
    }
    /// Parent group, or `None` for a root shape.
    #[must_use]
    pub fn parent(&self) -> Option<Self> {
        self.scene.data.records[self.index]
            .parent
            .map(|index| self.scene.handle(index as usize))
    }
    /// Direct child shapes, in source order; empty for non-group shapes.
    #[must_use]
    pub fn children(&self) -> Handles {
        let record = &self.scene.data.records[self.index];
        Handles {
            scene: self.scene.clone(),
            cursor: self.index + 1,
            end: (record.subtree_end as usize).min(self.scene.len()),
            parent: u32::try_from(self.index).ok(),
            preorder: false,
        }
    }
}
impl fmt::Debug for Handle {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Handle")
            .field("position", &self.position())
            .finish_non_exhaustive()
    }
}

/// Owning, fused iteration over checked scene positions.
#[derive(Clone)]
pub struct Handles {
    scene: Scene,
    cursor: usize,
    end: usize,
    parent: Option<u32>,
    preorder: bool,
}
impl Iterator for Handles {
    type Item = Handle;
    fn next(&mut self) -> Option<Handle> {
        while self.cursor < self.end {
            let index = self.cursor;
            let record = &self.scene.data.records[index];
            self.cursor = if self.preorder {
                index + 1
            } else {
                (record.subtree_end as usize).max(index + 1)
            };
            if self.preorder || record.parent == self.parent {
                return Some(self.scene.handle(index));
            }
        }
        None
    }
    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.end.saturating_sub(self.cursor);
        (if self.preorder { remaining } else { 0 }, Some(remaining))
    }
}
impl FusedIterator for Handles {}
impl fmt::Debug for Handles {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Handles")
            .field("remaining_upper_bound", &self.size_hint().1)
            .finish_non_exhaustive()
    }
}

/// Explicit access to the physical scene representation.
pub mod raw {
    /// The processed XML against which borrowed shape spans are defined.
    #[must_use]
    pub fn xml(scene: &super::Scene) -> &[u8] {
        scene.data.xml.bytes()
    }
}

#[cfg(test)]
mod tests;
