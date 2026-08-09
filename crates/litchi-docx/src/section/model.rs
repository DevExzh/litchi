#![expect(
    clippy::assigning_clones,
    reason = "clone assignment preserves validation-before-replacement behavior"
)]
#![expect(
    clippy::cast_possible_truncation,
    reason = "OOXML numeric values are bounded before conversion"
)]
#![expect(
    clippy::cast_precision_loss,
    reason = "OOXML unit conversion intentionally uses floating-point output"
)]
#![expect(
    clippy::expect_used,
    reason = "the invariant is established immediately before extraction"
)]
#![expect(
    clippy::iter_without_into_iter,
    reason = "the established collection API exposes explicit iterators"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Owned semantic section values and lazy section access.

use super::codec::{self, Snapshot};
use crate::error::{Error, Result};
use crate::header_footer::Kind;
use litchi_core::unit::{EMUS_PER_CM, EMUS_PER_INCH, EMUS_PER_PT, EMUS_PER_TWIP};
use std::fmt;

/// Maximum number of bytes accepted for one `w:sectPr` fragment.
pub(super) const MAX_XML_BYTES: usize = 4 * 1024 * 1024;

/// Maximum number of XML elements accepted in one section fragment.
pub(super) const MAX_XML_NODES: usize = 4096;

/// Maximum XML nesting accepted in one section fragment.
pub(super) const MAX_XML_DEPTH: usize = 64;

/// Maximum `WordprocessingML` measurement in twips accepted by Word's section
/// page geometry and margin domains (`[MS-OI29500]` §17.6.11 and §17.6.13).
pub(super) const MAX_TWIPS: i64 = 31_680;

/// Page layout orientation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Orientation {
    /// Portrait orientation.
    Portrait = 0,
    /// Landscape orientation.
    Landscape = 1,
}

impl Orientation {
    /// Convert the orientation to its `WordprocessingML` lexical value.
    #[inline]
    #[must_use]
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::Portrait => "portrait",
            Self::Landscape => "landscape",
        }
    }

    /// Parse a `WordprocessingML` orientation value.
    #[inline]
    #[must_use]
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "portrait" => Some(Self::Portrait),
            "landscape" => Some(Self::Landscape),
            _ => None,
        }
    }
}

impl Default for Orientation {
    #[inline]
    fn default() -> Self {
        Self::Portrait
    }
}

impl fmt::Display for Orientation {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Portrait => formatter.write_str("Portrait"),
            Self::Landscape => formatter.write_str("Landscape"),
        }
    }
}

/// Section-break placement.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Start {
    /// Continue on the current page.
    Continuous = 0,
    /// Begin in the next column.
    NewColumn = 1,
    /// Begin on the next page.
    NewPage = 2,
    /// Begin on the next even page.
    EvenPage = 3,
    /// Begin on the next odd page.
    OddPage = 4,
}

impl Start {
    /// Convert the section-break placement to its XML lexical value.
    #[inline]
    #[must_use]
    pub const fn to_xml(self) -> &'static str {
        match self {
            Self::Continuous => "continuous",
            Self::NewColumn => "nextColumn",
            Self::NewPage => "nextPage",
            Self::EvenPage => "evenPage",
            Self::OddPage => "oddPage",
        }
    }

    /// Parse a `WordprocessingML` section-break placement.
    #[inline]
    #[must_use]
    pub fn from_xml(value: &str) -> Option<Self> {
        match value {
            "continuous" => Some(Self::Continuous),
            "nextColumn" => Some(Self::NewColumn),
            "nextPage" => Some(Self::NewPage),
            "evenPage" => Some(Self::EvenPage),
            "oddPage" => Some(Self::OddPage),
            _ => None,
        }
    }
}

impl Default for Start {
    #[inline]
    fn default() -> Self {
        Self::NewPage
    }
}

impl fmt::Display for Start {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Continuous => formatter.write_str("Continuous"),
            Self::NewColumn => formatter.write_str("New Column"),
            Self::NewPage => formatter.write_str("New Page"),
            Self::EvenPage => formatter.write_str("Even Page"),
            Self::OddPage => formatter.write_str("Odd Page"),
        }
    }
}

/// A checked English Metric Unit value.
///
/// `WordprocessingML` section measurements are encoded as twips. `Emu` keeps
/// the ergonomic format-neutral unit used by the public DOCX facade and
/// performs checked conversion at the XML boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Emu(pub i64);

impl Emu {
    /// Construct from inches, retaining the historical infallible facade.
    ///
    /// Fallible callers should use [`Self::try_from_inches`]. Floating-point
    /// casts are saturating in Rust; XML serialization still requires an
    /// exact checked twip value.
    #[inline]
    #[must_use]
    pub const fn from_inches(inches: f64) -> Self {
        Self((inches * EMUS_PER_INCH as f64) as i64)
    }

    /// Construct from centimeters, retaining the historical infallible
    /// facade.
    #[inline]
    #[must_use]
    pub const fn from_cm(cm: f64) -> Self {
        Self((cm * EMUS_PER_CM as f64) as i64)
    }

    /// Construct from points, retaining the historical infallible facade.
    #[inline]
    #[must_use]
    pub const fn from_pt(pt: f64) -> Self {
        Self((pt * EMUS_PER_PT as f64) as i64)
    }

    /// Construct from twips with saturating arithmetic.
    #[inline]
    #[must_use]
    pub const fn from_twips(twips: i64) -> Self {
        Self(twips.saturating_mul(EMUS_PER_TWIP))
    }

    /// Construct from twips without allowing EMU overflow.
    #[inline]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn try_from_twips(twips: i64) -> Result<Self> {
        twips
            .checked_mul(EMUS_PER_TWIP)
            .map(Self)
            .ok_or_else(|| Error::InvalidFormat("twip value overflows EMU".into()))
    }

    /// Construct from inches after checking finiteness and the `i64` EMU
    /// domain.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn try_from_inches(inches: f64) -> Result<Self> {
        Self::try_from_float(inches, EMUS_PER_INCH, "inches")
    }

    /// Construct from centimeters after checking finiteness and the `i64`
    /// EMU domain.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn try_from_cm(cm: f64) -> Result<Self> {
        Self::try_from_float(cm, EMUS_PER_CM, "centimeters")
    }

    /// Construct from points after checking finiteness and the `i64` EMU
    /// domain.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn try_from_pt(pt: f64) -> Result<Self> {
        Self::try_from_float(pt, EMUS_PER_PT, "points")
    }

    fn try_from_float(value: f64, scale: i64, unit: &str) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::InvalidFormat(format!(
                "{unit} measurement is not finite"
            )));
        }
        let scaled = value * scale as f64;
        if scaled < i64::MIN as f64 || scaled > i64::MAX as f64 {
            return Err(Error::InvalidFormat(format!(
                "{unit} measurement overflows EMU"
            )));
        }
        Ok(Self(scaled as i64))
    }

    /// Convert to inches.
    #[inline]
    #[must_use]
    pub fn to_inches(self) -> f64 {
        self.0 as f64 / EMUS_PER_INCH as f64
    }

    /// Convert to centimeters.
    #[inline]
    #[must_use]
    pub fn to_cm(self) -> f64 {
        self.0 as f64 / EMUS_PER_CM as f64
    }

    /// Convert to points.
    #[inline]
    #[must_use]
    pub fn to_pt(self) -> f64 {
        self.0 as f64 / EMUS_PER_PT as f64
    }

    /// Convert to the nearest twip without overflow-prone floating-point
    /// arithmetic.
    #[inline]
    #[must_use]
    pub fn to_twips(self) -> i64 {
        let quotient = self.0 / EMUS_PER_TWIP;
        let remainder = self.0 % EMUS_PER_TWIP;
        if remainder.unsigned_abs() * 2 >= EMUS_PER_TWIP as u64 {
            quotient + remainder.signum()
        } else {
            quotient
        }
    }

    /// Return the exact twip representation required by `WordprocessingML`.
    #[inline]
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn try_to_twips(self) -> Result<i64> {
        if self.0 % EMUS_PER_TWIP != 0 {
            return Err(Error::InvalidFormat(
                "section measurement is not an exact twip".into(),
            ));
        }
        Ok(self.0 / EMUS_PER_TWIP)
    }

    /// Return the raw EMU count.
    #[inline]
    #[must_use]
    pub const fn get(self) -> i64 {
        self.0
    }
}

impl From<i64> for Emu {
    #[inline]
    fn from(value: i64) -> Self {
        Self(value)
    }
}

/// Page margins for a section. Values are optional because a `sectPr` may
/// inherit or default an omitted attribute.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Margins {
    /// Distance from the top page edge.
    pub top: Option<Emu>,
    /// Distance from the right page edge.
    pub right: Option<Emu>,
    /// Distance from the bottom page edge.
    pub bottom: Option<Emu>,
    /// Distance from the left page edge.
    pub left: Option<Emu>,
    /// Distance from the top page edge to the header.
    pub header: Option<Emu>,
    /// Distance from the bottom page edge to the footer.
    pub footer: Option<Emu>,
    /// Binding gutter.
    pub gutter: Option<Emu>,
}

/// Page geometry for a section.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageSize {
    /// Page width.
    pub width: Option<Emu>,
    /// Page height.
    pub height: Option<Emu>,
    /// Page orientation.
    pub orientation: Orientation,
}

impl Default for PageSize {
    fn default() -> Self {
        Self {
            width: None,
            height: None,
            orientation: Orientation::Portrait,
        }
    }
}

/// One explicitly sized section column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Column {
    /// Column width.
    pub width: Emu,
    /// Space after this column, when explicitly encoded.
    pub space: Option<Emu>,
}

/// Section newspaper-style column layout.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Columns {
    /// Whether Word distributes the available width equally.
    pub equal_width: bool,
    /// Declared column count.
    pub count: u16,
    /// Default space between columns.
    pub space: Option<Emu>,
    /// Whether a separator is drawn between columns.
    pub separator: bool,
    /// Explicit widths and spaces for unequal columns.
    pub columns: Vec<Column>,
}

impl Default for Columns {
    fn default() -> Self {
        Self {
            equal_width: true,
            count: 1,
            space: None,
            separator: false,
            columns: Vec::new(),
        }
    }
}

/// A relationship reference carried by a section's header or footer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    /// Header/footer role.
    pub kind: Kind,
    /// Relationship ID in the owning document part.
    pub relationship_id: String,
}

#[derive(Debug, Clone, Default)]
pub(crate) struct State {
    pub(crate) page_size: Option<PageSize>,
    pub(crate) margins: Option<Margins>,
    pub(crate) start: Option<Start>,
    pub(crate) columns: Option<Columns>,
    pub(crate) headers: Vec<Reference>,
    pub(crate) footers: Vec<Reference>,
}

/// A document section backed by an owned `w:sectPr` XML fragment.
///
/// The original fragment is retained until a mutation is requested. Semantic
/// values are decoded lazily, so simply inspecting a document does not copy or
/// normalize its unknown children. An unchanged section round-trips byte for
/// byte; mutations retain unknown direct children and their authored order.
#[derive(Debug, Clone)]
pub struct Section {
    xml_bytes: Vec<u8>,
    snapshot: Option<Snapshot>,
}

impl Default for Section {
    fn default() -> Self {
        Self::from_xml_bytes(b"<w:sectPr/>".to_vec()).expect("built-in section is valid")
    }
}

impl Section {
    /// Construct a section from a bounded `w:sectPr` fragment.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml_bytes(xml_bytes: Vec<u8>) -> Result<Self> {
        codec::validate(&xml_bytes)?;
        super::footnote_columns::Snapshot::from_xml(xml_bytes.clone())?;
        Ok(Self {
            xml_bytes,
            snapshot: None,
        })
    }

    /// Return the authored XML bytes currently held by this section.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml_bytes
    }

    /// Return an owned XML snapshot, preserving unknown content.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn to_xml_bytes(&mut self) -> Result<Vec<u8>> {
        if let Some(snapshot) = &self.snapshot
            && snapshot.is_dirty()
        {
            let bytes = codec::encode(snapshot)?;
            self.xml_bytes = bytes.clone();
            self.snapshot = Some(codec::decode(&bytes)?);
            return Ok(bytes);
        }
        Ok(self.xml_bytes.clone())
    }

    /// Return the current section XML as UTF-8 text.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn to_xml(&mut self) -> Result<String> {
        String::from_utf8(self.to_xml_bytes()?)
            .map_err(|error| Error::InvalidFormat(format!("section XML is not UTF-8: {error}")))
    }

    /// Validate the fragment and, on success, cache its semantic snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn validate(&mut self) -> Result<()> {
        self.snapshot_mut().map(|_| ())
    }

    /// Return the page width, if the local section specifies one.
    pub fn page_width(&mut self) -> Option<Emu> {
        self.page_size_checked().ok().and_then(|size| size.width)
    }

    /// Return the page height, if the local section specifies one.
    pub fn page_height(&mut self) -> Option<Emu> {
        self.page_size_checked().ok().and_then(|size| size.height)
    }

    /// Return the local page geometry, defaulting omitted geometry to an empty
    /// portrait value.
    pub fn page_size(&mut self) -> PageSize {
        self.page_size_checked().unwrap_or_default()
    }

    /// Return the local page geometry with parsing errors exposed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn page_size_checked(&mut self) -> Result<PageSize> {
        Ok(self.snapshot_mut()?.state.page_size.unwrap_or_default())
    }

    /// Return the local page orientation.
    pub fn orientation(&mut self) -> Orientation {
        self.page_size().orientation
    }

    /// Return the local margins, with omitted attributes represented by
    /// `None`.
    pub fn margins(&mut self) -> Margins {
        self.margins_checked().unwrap_or_default()
    }

    /// Return local margins with parsing errors exposed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn margins_checked(&mut self) -> Result<Margins> {
        Ok(self.snapshot_mut()?.state.margins.unwrap_or_default())
    }

    /// Return the top margin.
    pub fn top_margin(&mut self) -> Option<Emu> {
        self.margins_checked().ok().and_then(|margins| margins.top)
    }

    /// Return the right margin.
    pub fn right_margin(&mut self) -> Option<Emu> {
        self.margins_checked()
            .ok()
            .and_then(|margins| margins.right)
    }

    /// Return the bottom margin.
    pub fn bottom_margin(&mut self) -> Option<Emu> {
        self.margins_checked()
            .ok()
            .and_then(|margins| margins.bottom)
    }

    /// Return the left margin.
    pub fn left_margin(&mut self) -> Option<Emu> {
        self.margins_checked().ok().and_then(|margins| margins.left)
    }

    /// Return the header distance.
    pub fn header_distance(&mut self) -> Option<Emu> {
        self.margins_checked()
            .ok()
            .and_then(|margins| margins.header)
    }

    /// Return the footer distance.
    pub fn footer_distance(&mut self) -> Option<Emu> {
        self.margins_checked()
            .ok()
            .and_then(|margins| margins.footer)
    }

    /// Return the gutter distance.
    pub fn gutter(&mut self) -> Option<Emu> {
        self.margins_checked()
            .ok()
            .and_then(|margins| margins.gutter)
    }

    /// Return the local section-break placement.
    pub fn start_type(&mut self) -> Start {
        self.snapshot_mut()
            .ok()
            .and_then(|snapshot| snapshot.state.start)
            .unwrap_or_default()
    }

    /// Return the local section-break placement with parsing errors exposed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn start(&mut self) -> Result<Option<Start>> {
        Ok(self.snapshot_mut()?.state.start)
    }

    /// Return local columns, if a `w:cols` child exists.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn columns(&mut self) -> Result<Option<Columns>> {
        Ok(self.snapshot_mut()?.state.columns.clone())
    }

    /// Return the direct Word 2012 footnote-area layout, if present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn footnote_columns(&mut self) -> Result<Option<super::footnote_columns::Layout>> {
        Ok(self.footnote_columns_snapshot()?.layout())
    }

    /// Return a detached, lossless snapshot for the section's Word 2012
    /// footnote-column extension.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn footnote_columns_snapshot(&mut self) -> Result<super::footnote_columns::Snapshot> {
        let xml = self.to_xml_bytes()?;
        super::footnote_columns::Snapshot::from_xml(xml)
    }

    /// Return local header references.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn headers(&mut self) -> Result<Vec<Reference>> {
        Ok(self.snapshot_mut()?.state.headers.clone())
    }

    /// Return local footer references.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn footers(&mut self) -> Result<Vec<Reference>> {
        Ok(self.snapshot_mut()?.state.footers.clone())
    }

    /// Replace local page geometry and commit one lossless XML edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_page_size(&mut self, page_size: PageSize) -> Result<()> {
        codec::validate_page_size(&page_size)?;
        self.update(|state| {
            state.page_size = Some(page_size);
        })
    }

    /// Remove the local `w:pgSz` child.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn clear_page_size(&mut self) -> Result<()> {
        self.update(|state| state.page_size = None)
    }

    /// Replace local margins and commit one lossless XML edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_margins(&mut self, margins: Margins) -> Result<()> {
        codec::validate_margins(&margins)?;
        self.update(|state| state.margins = Some(margins))
    }

    /// Remove the local `w:pgMar` child.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn clear_margins(&mut self) -> Result<()> {
        self.update(|state| state.margins = None)
    }

    /// Set or clear the local section-break placement.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_start(&mut self, start: Option<Start>) -> Result<()> {
        self.update(|state| state.start = start)
    }

    /// Set or clear the local column layout.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_columns(&mut self, columns: Option<Columns>) -> Result<()> {
        if let Some(columns) = &columns {
            codec::validate_columns(columns)?;
        }
        self.update(|state| state.columns = columns)
    }

    /// Set or remove the direct Word 2012 footnote-area layout atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_footnote_columns(
        &mut self,
        value: Option<super::footnote_columns::Layout>,
    ) -> Result<()> {
        let snapshot = self.footnote_columns_snapshot()?;
        let mut edit = snapshot.edit();
        edit.set_layout(value)?;
        let commit = edit.commit()?;
        self.xml_bytes = commit.snapshot().xml_bytes().to_vec();
        self.snapshot = None;
        Ok(())
    }

    /// Remove the direct Word 2012 footnote-area layout marker.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn clear_footnote_columns(&mut self) -> Result<()> {
        self.set_footnote_columns(None)
    }

    /// Apply a preconditioned footnote-layout patch atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply_footnote_columns_patch(
        &mut self,
        patch: &super::footnote_columns::Patch,
    ) -> Result<()> {
        let snapshot = self.footnote_columns_snapshot()?;
        let next = patch.apply(&snapshot)?;
        self.xml_bytes = next.xml_bytes().to_vec();
        self.snapshot = None;
        Ok(())
    }

    fn snapshot_mut(&mut self) -> Result<&mut Snapshot> {
        if self.snapshot.is_none() {
            self.snapshot = Some(codec::decode(&self.xml_bytes)?);
        }
        Ok(self
            .snapshot
            .as_mut()
            .expect("section snapshot initialized"))
    }

    fn update(&mut self, edit: impl FnOnce(&mut State)) -> Result<()> {
        let mut candidate = self.snapshot_mut()?.clone();
        let previous = candidate.state.clone();
        edit(&mut candidate.state);
        candidate.mark_dirty_since(&previous);
        let bytes = codec::encode(&candidate)?;
        self.xml_bytes = bytes.clone();
        self.snapshot = Some(codec::decode(&bytes)?);
        Ok(())
    }

    fn local_state(&mut self) -> Result<State> {
        Ok(self.snapshot_mut()?.state.clone())
    }
}

/// An owned ordered collection of document sections.
#[derive(Debug, Clone, Default)]
pub struct Sections {
    sections: Vec<Section>,
}

impl Sections {
    /// Construct a section collection.
    #[inline]
    #[must_use]
    pub fn new(sections: Vec<Section>) -> Self {
        Self { sections }
    }

    /// Number of sections.
    #[inline]
    #[must_use]
    pub fn len(&self) -> usize {
        self.sections.len()
    }

    /// Whether the collection is empty.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.sections.is_empty()
    }

    /// Append a section.
    #[inline]
    pub fn push(&mut self, section: Section) {
        self.sections.push(section);
    }

    /// Insert a section at an existing position.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn insert(&mut self, index: usize, section: Section) -> Result<()> {
        if index > self.sections.len() {
            return Err(Error::OutOfBounds {
                object: "section",
                index,
                len: self.sections.len(),
            });
        }
        self.sections.insert(index, section);
        Ok(())
    }

    /// Remove and return a section.
    #[inline]
    pub fn remove(&mut self, index: usize) -> Option<Section> {
        (index < self.sections.len()).then(|| self.sections.remove(index))
    }

    /// Get a mutable section by index.
    #[inline]
    pub fn get_mut(&mut self, index: usize) -> Option<&mut Section> {
        self.sections.get_mut(index)
    }

    /// Get a section by index.
    #[inline]
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&Section> {
        self.sections.get(index)
    }

    /// Iterate mutably in document order.
    #[inline]
    pub fn iter_mut(&mut self) -> std::slice::IterMut<'_, Section> {
        self.sections.iter_mut()
    }

    /// Iterate in document order.
    #[inline]
    pub fn iter(&self) -> std::slice::Iter<'_, Section> {
        self.sections.iter()
    }

    /// Resolve page geometry through the section inheritance chain.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn effective_page_size(&mut self, index: usize) -> Result<PageSize> {
        let states = self.states_through(index)?;
        Ok(states
            .into_iter()
            .fold(PageSize::default(), |mut current, state| {
                if let Some(page_size) = state.page_size {
                    current = page_size;
                }
                current
            }))
    }

    /// Resolve margins through the section inheritance chain, retaining each
    /// attribute until a later section overrides it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn effective_margins(&mut self, index: usize) -> Result<Margins> {
        let states = self.states_through(index)?;
        Ok(states
            .into_iter()
            .fold(Margins::default(), |mut current, state| {
                if let Some(margins) = state.margins {
                    if margins.top.is_some() {
                        current.top = margins.top;
                    }
                    if margins.right.is_some() {
                        current.right = margins.right;
                    }
                    if margins.bottom.is_some() {
                        current.bottom = margins.bottom;
                    }
                    if margins.left.is_some() {
                        current.left = margins.left;
                    }
                    if margins.header.is_some() {
                        current.header = margins.header;
                    }
                    if margins.footer.is_some() {
                        current.footer = margins.footer;
                    }
                    if margins.gutter.is_some() {
                        current.gutter = margins.gutter;
                    }
                }
                current
            }))
    }

    /// Resolve the header references inherited by a section.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn effective_headers(&mut self, index: usize) -> Result<Vec<Reference>> {
        self.effective_references(index, true)
    }

    /// Resolve the footer references inherited by a section.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn effective_footers(&mut self, index: usize) -> Result<Vec<Reference>> {
        self.effective_references(index, false)
    }

    fn states_through(&mut self, index: usize) -> Result<Vec<State>> {
        if index >= self.sections.len() {
            return Err(Error::OutOfBounds {
                object: "section",
                index,
                len: self.sections.len(),
            });
        }
        self.sections[..=index]
            .iter_mut()
            .map(Section::local_state)
            .collect()
    }

    fn effective_references(&mut self, index: usize, headers: bool) -> Result<Vec<Reference>> {
        let states = self.states_through(index)?;
        let mut effective = Vec::new();
        for state in states {
            let references = if headers {
                state.headers
            } else {
                state.footers
            };
            for reference in references {
                if let Some(existing) = effective
                    .iter_mut()
                    .find(|existing: &&mut Reference| existing.kind == reference.kind)
                {
                    *existing = reference;
                } else {
                    effective.push(reference);
                }
            }
        }
        Ok(effective)
    }
}

impl std::ops::Index<usize> for Sections {
    type Output = Section;

    #[inline]
    fn index(&self, index: usize) -> &Self::Output {
        &self.sections[index]
    }
}

impl std::ops::IndexMut<usize> for Sections {
    #[inline]
    fn index_mut(&mut self, index: usize) -> &mut Self::Output {
        &mut self.sections[index]
    }
}
