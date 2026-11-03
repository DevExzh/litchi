//! Paragraph-list engine (record 0374).
//!
//! A storage names a `TSWP.ListStyleArchive` (2023) per run of paragraphs
//! in its list-style table (field 7); each paragraph's nesting level lives
//! in the paragraph-data table (field 6) and a numbering restart in the
//! paragraph-start table (field 14). A list style is one of the canonical
//! nine-level None/Bullet/Numbered presets of its stylesheet, or a variation
//! of one whose per-level arrays (bullet strings, label geometries, label
//! and text indents, label color, number formats, number tiering) override
//! the preset's.
//!
//! The engine reads all of this ([`placements`], [`level`], [`numbering`],
//! [`property`]) and plans the retired host's (commit `27b4840684`) per-paragraph property
//! changes ([`plan`], [`plan_collapse`]): rewrite an isolated exclusive
//! variation in place, derive a new variation for one paragraph, collapse
//! a variation that no longer overrides anything, or strip one property
//! from it. Level and numbering changes rewrite the storage only
//! ([`with_level`], [`with_numbering`]). Preset assignment reuses or
//! creates the stylesheet's preset objects ([`find_preset`],
//! [`preset_payload`]) and rewrites the boundaries ([`replace_boundaries`]).

pub mod publish;
mod storage;
mod style;

use std::collections::{BTreeMap, HashMap, HashSet};
use std::hash::BuildHasher;

use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListBullet, ParagraphListBulletBaselineOffset,
    ParagraphListBulletGeometry, ParagraphListBulletScale, ParagraphListIndentation,
    ParagraphListLabelColor, ParagraphListLabelIndent, ParagraphListLevel,
    ParagraphListLevelPlacement, ParagraphListNumberFormat, ParagraphListNumberScale,
    ParagraphListNumberTiering, ParagraphListNumbering, ParagraphListPlacement, ParagraphListStart,
    ParagraphListTextGap,
};
use litchi_iwa_text::position::TextPosition;
use thiserror::Error as ThisError;

use crate::native::{Budget, NativeError};
use style::{Arrays, Direct, Geometry, ListStyle};

/// `TSWP.ListStyleArchive`.
pub const MESSAGE_TYPE: u32 = 2_023;
/// The message version a new list style is stored with.
pub const MESSAGE_VERSION: [u32; 3] = [1, 0, 5];

/// Why a well-formed list change is refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The position is not a paragraph start.
    NotAParagraphStart,
    /// The paragraph's list kind has no such property (a bullet on a
    /// numbered paragraph, a label on an ordinary paragraph, …).
    NotApplicable,
    /// Placements must start at zero and strictly increase.
    InvalidPlacements,
}

/// Why a list operation failed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A payload is malformed, a style is missing, or a native value cannot
    /// be represented.
    #[error("invalid paragraph-list source: {reason}")]
    InvalidSource {
        /// What is wrong.
        reason: &'static str,
    },
    /// A well-formed change this engine refuses.
    #[error("paragraph-list change refused: {0:?}")]
    Refused(Refusal),
    /// A configured wire resource was exhausted.
    #[error("paragraph-list {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The attempted aggregate consumption.
        observed: usize,
        /// The caller's configured bound.
        maximum: usize,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} paragraph-list items or bytes")]
    Allocation {
        /// The attempted item or byte capacity.
        amount: usize,
    },
}

impl From<NativeError> for Error {
    fn from(error: NativeError) -> Self {
        match error {
            NativeError::Invalid(reason) => Self::InvalidSource { reason },
            NativeError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind,
                observed,
                maximum,
            },
            NativeError::Allocation { amount } => Self::Allocation { amount },
        }
    }
}

/// The result of a list operation.
pub type Result<T> = std::result::Result<T, Error>;

fn invalid<T>(reason: &'static str) -> Result<T> {
    Err(Error::InvalidSource { reason })
}

/// List-style payloads by object identifier.
pub trait Styles {
    /// The `TSWP.ListStyleArchive` payload of `identifier`.
    fn list_style(&self, identifier: u64) -> Option<&[u8]>;
}

impl<S: BuildHasher> Styles for HashMap<u64, Vec<u8>, S> {
    fn list_style(&self, identifier: u64) -> Option<&[u8]> {
        self.get(&identifier).map(Vec::as_slice)
    }
}

impl Styles for BTreeMap<u64, Vec<u8>> {
    fn list_style(&self, identifier: u64) -> Option<&[u8]> {
        self.get(&identifier).map(Vec::as_slice)
    }
}

fn locate(
    styles: &(impl Styles + ?Sized),
    identifier: u64,
    budget: &mut Budget,
) -> Result<ListStyle> {
    let Some(payload) = styles.list_style(identifier) else {
        return invalid("iWork list style is missing");
    };
    Ok(ListStyle::decode(payload, budget)?)
}

fn parent_of(style: &ListStyle) -> Result<u64> {
    style.parent().ok_or(Error::InvalidSource {
        reason: "iWork list-style variation has no parent",
    })
}

/// The list preset a style resolves to through its parent chain.
fn resolve(
    styles: &(impl Styles + ?Sized),
    first: u64,
    budget: &mut Budget,
) -> Result<ParagraphList> {
    let mut current = first;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current) {
            return invalid("iWork list-style inheritance contains a cycle");
        }
        let style = locate(styles, current, budget)?;
        if let Some(preset) = style::preset_of(&style) {
            return Ok(preset);
        }
        current = parent_of(&style)?;
    }
}

/// The first style on the chain whose `pick` is present.
fn first_with<T>(
    styles: &(impl Styles + ?Sized),
    first: u64,
    budget: &mut Budget,
    mut pick: impl FnMut(&ListStyle) -> Option<T>,
) -> Result<T> {
    let mut current = first;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current) {
            return invalid("iWork list-style inheritance contains a cycle");
        }
        let style = locate(styles, current, budget)?;
        if let Some(value) = pick(&style) {
            return Ok(value);
        }
        current = parent_of(&style)?;
    }
}

fn strings(
    styles: &(impl Styles + ?Sized),
    style: u64,
    budget: &mut Budget,
) -> Result<Vec<String>> {
    if resolve(styles, style, budget)? != ParagraphList::Bullet {
        return Err(Error::Refused(Refusal::NotApplicable));
    }
    let strings = first_with(styles, style, budget, |s| {
        (!s.strings.is_empty()).then(|| s.strings.clone())
    })?;
    style::validate_strings(&strings)?;
    Ok(strings)
}

fn geometries(
    styles: &(impl Styles + ?Sized),
    style: u64,
    budget: &mut Budget,
) -> Result<Vec<Geometry>> {
    if resolve(styles, style, budget)? == ParagraphList::None {
        return Err(Error::Refused(Refusal::NotApplicable));
    }
    let geometries = first_with(styles, style, budget, |s| {
        (!s.geometries.is_empty()).then(|| s.geometries.clone())
    })?;
    style::validate_geometries(&geometries)?;
    Ok(geometries)
}

fn floats(
    styles: &(impl Styles + ?Sized),
    style: u64,
    text: bool,
    budget: &mut Budget,
) -> Result<Vec<f32>> {
    if resolve(styles, style, budget)? == ParagraphList::None {
        return Err(Error::Refused(Refusal::NotApplicable));
    }
    let values = first_with(styles, style, budget, |s| {
        let values = if text { &s.text_indents } else { &s.indents };
        (!values.is_empty()).then(|| values.clone())
    })?;
    style::validate_floats(&values)?;
    Ok(values)
}

fn number_types(
    styles: &(impl Styles + ?Sized),
    style: u64,
    budget: &mut Budget,
) -> Result<Vec<i32>> {
    if resolve(styles, style, budget)? != ParagraphList::Numbered {
        return Err(Error::Refused(Refusal::NotApplicable));
    }
    let values = first_with(styles, style, budget, |s| {
        (!s.number_types.is_empty()).then(|| s.number_types.clone())
    })?;
    style::validate_number_types(&values)?;
    Ok(values)
}

fn tiered(styles: &(impl Styles + ?Sized), style: u64, budget: &mut Budget) -> Result<Vec<bool>> {
    if resolve(styles, style, budget)? != ParagraphList::Numbered {
        return Err(Error::Refused(Refusal::NotApplicable));
    }
    let values = first_with(styles, style, budget, |s| {
        (!s.tiered_numbers.is_empty()).then(|| s.tiered_numbers.clone())
    })?;
    style::validate_tiered(&values)?;
    Ok(values)
}

fn label_color(
    styles: &(impl Styles + ?Sized),
    first: u64,
    budget: &mut Budget,
) -> Result<ParagraphListLabelColor> {
    if resolve(styles, first, budget)? == ParagraphList::None {
        return Err(Error::Refused(Refusal::NotApplicable));
    }
    let mut current = first;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current) {
            return invalid("iWork list-style inheritance contains a cycle");
        }
        let style = locate(styles, current, budget)?;
        if let Some(color) = style::label_color_from_native(&style)? {
            return Ok(color);
        }
        let Some(parent) = style.parent() else {
            return Ok(ParagraphListLabelColor::Automatic);
        };
        current = parent;
    }
}

/// The stylesheet a list style belongs to (its own, else its nearest
/// ancestor's).
///
/// # Errors
///
/// [`Error::InvalidSource`] for a missing style, a cycle, or a chain
/// without stylesheet.
pub fn stylesheet(styles: &(impl Styles + ?Sized), style: u64, limits: WireLimits) -> Result<u64> {
    let mut budget = Budget::new(limits);
    first_with(styles, style, &mut budget, |s| {
        s.base
            .stylesheet
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0)
    })
    .or_else(|error| match error {
        Error::InvalidSource {
            reason: "iWork list-style variation has no parent",
        } => invalid("iWork list style has neither a stylesheet nor a parent"),
        other => Err(other),
    })
}

/// The list preset a style resolves to.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a missing style, a cycle, or a chain that
/// reaches no canonical preset.
pub fn preset(
    styles: &(impl Styles + ?Sized),
    style: u64,
    limits: WireLimits,
) -> Result<ParagraphList> {
    resolve(styles, style, &mut Budget::new(limits))
}

/// The storage's uniform list style: its only list-style entry, at zero.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the storage has no table, several
/// boundaries or no style.
pub fn uniform_style(storage: &[u8], limits: WireLimits) -> Result<u64> {
    let mut budget = Budget::new(limits);
    let source = storage::decode(storage, &mut budget)?;
    let Some(table) = source.list_table else {
        return invalid("text storage must contain one list-style table");
    };
    let entries = storage::references(table, &mut budget)?;
    let [entry] = entries.as_slice() else {
        return invalid("text storage must have one uniform list-style boundary");
    };
    if entry.index != 0 {
        return invalid("text storage list style must begin at UTF-16 index zero");
    }
    entry
        .object
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource {
            reason: "text storage has no uniform list style",
        })
}

/// The storage's list-style boundaries `(UTF-16 index, style)`.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a missing or malformed table.
pub fn boundaries(storage: &[u8], limits: WireLimits) -> Result<Vec<(u32, u64)>> {
    let mut budget = Budget::new(limits);
    let source = storage::decode(storage, &mut budget)?;
    Ok(storage::boundaries(&source, &mut budget)?)
}

/// The storage's paragraph starts (UTF-16 indexes).
///
/// # Errors
///
/// [`Error::InvalidSource`] for malformed text.
pub fn paragraph_starts(storage: &[u8], limits: WireLimits) -> Result<Vec<u32>> {
    Ok(storage::decode(storage, &mut Budget::new(limits))?.starts)
}

/// The list preset of every boundary.
///
/// # Errors
///
/// As [`boundaries`] and [`preset`].
pub fn placements(
    storage: &[u8],
    styles: &(impl Styles + ?Sized),
    limits: WireLimits,
) -> Result<Vec<ParagraphListPlacement>> {
    let mut budget = Budget::new(limits);
    let source = storage::decode(storage, &mut budget)?;
    storage::boundaries(&source, &mut budget)?
        .into_iter()
        .map(|(index, style)| {
            Ok(ParagraphListPlacement::new(
                TextPosition::from_utf16_code_units(index),
                resolve(styles, style, &mut budget)?,
            ))
        })
        .collect()
}

fn effective_style(boundaries: &[(u32, u64)], starts: &[u32], paragraph: u32) -> Result<u64> {
    if starts.binary_search(&paragraph).is_err() {
        return Err(Error::Refused(Refusal::NotAParagraphStart));
    }
    boundaries
        .iter()
        .take_while(|entry| entry.0 <= paragraph)
        .last()
        .map(|entry| entry.1)
        .ok_or(Error::InvalidSource {
            reason: "text storage has no list-style boundary at zero",
        })
}

fn isolated(boundaries: &[(u32, u64)], starts: &[u32], paragraph: u32) -> bool {
    if boundaries
        .binary_search_by_key(&paragraph, |entry| entry.0)
        .is_err()
    {
        return false;
    }
    let Some(next) = starts.iter().copied().find(|start| *start > paragraph) else {
        return true;
    };
    boundaries
        .binary_search_by_key(&next, |entry| entry.0)
        .is_ok()
}

/// The boundaries with the paragraph at `paragraph` styled by `style` and
/// every other paragraph unchanged, coalesced (the host's
/// `paragraph_boundaries_with_style`).
///
/// # Errors
///
/// [`Error::Refused`] when `paragraph` is not a paragraph start; as
/// [`boundaries`] otherwise.
pub fn boundaries_with_style(
    storage: &[u8],
    paragraph: TextPosition,
    style: u64,
    limits: WireLimits,
) -> Result<Vec<(u32, u64)>> {
    let mut budget = Budget::new(limits);
    let source = storage::decode(storage, &mut budget)?;
    let current = storage::boundaries(&source, &mut budget)?;
    let target = paragraph.utf16_index();
    if source.starts.binary_search(&target).is_err() {
        return Err(Error::Refused(Refusal::NotAParagraphStart));
    }
    let mut result: Vec<(u32, u64)> = Vec::new();
    for &start in &source.starts {
        let identifier = if start == target {
            style
        } else {
            current
                .iter()
                .take_while(|entry| entry.0 <= start)
                .last()
                .map(|entry| entry.1)
                .ok_or(Error::InvalidSource {
                    reason: "text storage has no list-style boundary at zero",
                })?
        };
        if result
            .last()
            .is_none_or(|previous| previous.1 != identifier)
        {
            result.push((start, identifier));
        }
    }
    Ok(result)
}

/// Rewrite the list-style table to `boundaries`, keeping the bytes of
/// entries at unchanged indexes. The storage message's references are the
/// caller's: replace the old boundary styles, in place, by the new ones.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed table.
pub fn replace_boundaries(
    storage: &[u8],
    boundaries: &[(u32, u64)],
    limits: WireLimits,
) -> Result<Vec<u8>> {
    Ok(storage::replace(
        storage,
        boundaries,
        &mut Budget::new(limits),
    )?)
}

fn level_of(
    source: &storage::Source<'_>,
    entries: &[storage::Data],
    paragraph: u32,
) -> Result<ParagraphListLevel> {
    if source.starts.binary_search(&paragraph).is_err() {
        return Err(Error::Refused(Refusal::NotAParagraphStart));
    }
    storage::validate_levels(source, entries)?;
    let entry = entries
        .iter()
        .take_while(|entry| entry.index <= paragraph)
        .last()
        .ok_or(Error::InvalidSource {
            reason: "text storage has no paragraph-data boundary at zero",
        })?;
    ParagraphListLevel::new(u8::try_from(entry.first).unwrap_or(u8::MAX)).map_err(|_| {
        Error::InvalidSource {
            reason: "native paragraph list level is out of range",
        }
    })
}

fn level_source<'a>(
    storage: &'a [u8],
    budget: &mut Budget,
) -> Result<(storage::Source<'a>, Vec<storage::Data>)> {
    let source = storage::decode(storage, budget)?;
    let Some(table) = source.data_table else {
        return invalid("text storage must contain one paragraph-data table");
    };
    let entries = storage::data_entries(table, budget)?;
    Ok((source, entries))
}

/// Every paragraph-level boundary.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a missing or malformed table.
pub fn levels(storage: &[u8], limits: WireLimits) -> Result<Vec<ParagraphListLevelPlacement>> {
    let mut budget = Budget::new(limits);
    let (source, entries) = level_source(storage, &mut budget)?;
    storage::validate_levels(&source, &entries)?;
    entries
        .iter()
        .map(|entry| {
            Ok(ParagraphListLevelPlacement::new(
                TextPosition::from_utf16_code_units(entry.index),
                ParagraphListLevel::new(u8::try_from(entry.first).unwrap_or(u8::MAX)).map_err(
                    |_| Error::InvalidSource {
                        reason: "native paragraph list level is out of range",
                    },
                )?,
            ))
        })
        .collect()
}

/// The level of the paragraph at `paragraph`.
///
/// # Errors
///
/// [`Error::Refused`] when `paragraph` is not a paragraph start; as
/// [`levels`] otherwise.
pub fn level(
    storage: &[u8],
    paragraph: TextPosition,
    limits: WireLimits,
) -> Result<ParagraphListLevel> {
    let mut budget = Budget::new(limits);
    let (source, entries) = level_source(storage, &mut budget)?;
    level_of(&source, &entries, paragraph.utf16_index())
}

/// Set the level of one paragraph; `None` when it already has it.
///
/// # Errors
///
/// As [`level`].
pub fn with_level(
    storage: &[u8],
    paragraph: TextPosition,
    level: ParagraphListLevel,
    limits: WireLimits,
) -> Result<Option<Vec<u8>>> {
    if self::level(storage, paragraph, limits)? == level {
        return Ok(None);
    }
    Ok(Some(storage::with_level(
        storage,
        paragraph.utf16_index(),
        u32::from(level.get()),
        &mut Budget::new(limits),
    )?))
}

fn numbering_from_native(value: u32) -> Result<ParagraphListNumbering> {
    if value == 0 {
        return Ok(ParagraphListNumbering::Continue);
    }
    ParagraphListStart::new(value)
        .map(ParagraphListNumbering::StartAt)
        .map_err(|_| Error::InvalidSource {
            reason: "native list numbering start is out of range",
        })
}

/// The numbering of the paragraph at `paragraph`.
///
/// # Errors
///
/// [`Error::Refused`] when `paragraph` is not a paragraph start;
/// [`Error::InvalidSource`] for a malformed table.
pub fn numbering(
    storage: &[u8],
    paragraph: TextPosition,
    limits: WireLimits,
) -> Result<ParagraphListNumbering> {
    let mut budget = Budget::new(limits);
    let source = storage::decode(storage, &mut budget)?;
    if source
        .starts
        .binary_search(&paragraph.utf16_index())
        .is_err()
    {
        return Err(Error::Refused(Refusal::NotAParagraphStart));
    }
    let Some(table) = source.start_table else {
        return Ok(ParagraphListNumbering::Continue);
    };
    let entries = storage::data_entries(table, &mut budget)?;
    storage::validate_numbering(&source, &entries)?;
    for entry in &entries {
        numbering_from_native(entry.first)?;
    }
    entries
        .iter()
        .find(|entry| entry.index == paragraph.utf16_index())
        .map_or(Ok(ParagraphListNumbering::Continue), |entry| {
            numbering_from_native(entry.first)
        })
}

/// Set one paragraph's numbering; `None` when it already has it. A
/// restart requires a numbered paragraph.
///
/// # Errors
///
/// [`Error::Refused`] for a restart on a paragraph that is not numbered;
/// as [`numbering`] and [`placements`] otherwise.
pub fn with_numbering(
    storage: &[u8],
    styles: &(impl Styles + ?Sized),
    paragraph: TextPosition,
    numbering: ParagraphListNumbering,
    limits: WireLimits,
) -> Result<Option<Vec<u8>>> {
    if self::numbering(storage, paragraph, limits)? == numbering {
        return Ok(None);
    }
    if matches!(numbering, ParagraphListNumbering::StartAt(_)) {
        let list = placements(storage, styles, limits)?
            .iter()
            .take_while(|placement| placement.paragraph <= paragraph)
            .last()
            .map(|placement| placement.list)
            .ok_or(Error::InvalidSource {
                reason: "text storage has no list-style boundary at zero",
            })?;
        if list != ParagraphList::Numbered {
            return Err(Error::Refused(Refusal::NotApplicable));
        }
    }
    let start = match numbering {
        ParagraphListNumbering::Continue => 0,
        ParagraphListNumbering::StartAt(start) => start.get(),
    };
    Ok(Some(storage::with_numbering(
        storage,
        paragraph.utf16_index(),
        start,
        &mut Budget::new(limits),
    )?))
}

/// One per-paragraph list property with its value.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Property {
    /// The text-bullet marker of the paragraph's level.
    Bullet(ParagraphListBullet),
    /// The bullet's size and baseline.
    BulletGeometry(ParagraphListBulletGeometry),
    /// Label indent and text gap.
    Indentation(ParagraphListIndentation),
    /// The label color.
    LabelColor(ParagraphListLabelColor),
    /// The number format of the paragraph's level.
    NumberFormat(ParagraphListNumberFormat),
    /// The number label's size.
    NumberScale(ParagraphListNumberScale),
    /// Flat or tiered numbering.
    NumberTiering(ParagraphListNumberTiering),
}

/// The property a [`Property`] value sets.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PropertyKind {
    /// [`Property::Bullet`].
    Bullet,
    /// [`Property::BulletGeometry`].
    BulletGeometry,
    /// [`Property::Indentation`].
    Indentation,
    /// [`Property::LabelColor`].
    LabelColor,
    /// [`Property::NumberFormat`].
    NumberFormat,
    /// [`Property::NumberScale`].
    NumberScale,
    /// [`Property::NumberTiering`].
    NumberTiering,
}

impl PropertyKind {
    /// Every kind.
    pub const ALL: [Self; 7] = [
        Self::Bullet,
        Self::BulletGeometry,
        Self::Indentation,
        Self::LabelColor,
        Self::NumberFormat,
        Self::NumberScale,
        Self::NumberTiering,
    ];
}

impl Property {
    /// The property this value sets.
    #[must_use]
    pub const fn kind(&self) -> PropertyKind {
        match self {
            Self::Bullet(_) => PropertyKind::Bullet,
            Self::BulletGeometry(_) => PropertyKind::BulletGeometry,
            Self::Indentation(_) => PropertyKind::Indentation,
            Self::LabelColor(_) => PropertyKind::LabelColor,
            Self::NumberFormat(_) => PropertyKind::NumberFormat,
            Self::NumberScale(_) => PropertyKind::NumberScale,
            Self::NumberTiering(_) => PropertyKind::NumberTiering,
        }
    }
}

struct Context<'a> {
    storage: storage::Source<'a>,
    boundaries: Vec<(u32, u64)>,
    level: usize,
    style: u64,
}

fn context<'a>(
    storage: &'a [u8],
    paragraph: TextPosition,
    budget: &mut Budget,
) -> Result<Context<'a>> {
    let source = storage::decode(storage, budget)?;
    let boundaries = storage::boundaries(&source, budget)?;
    let target = paragraph.utf16_index();
    let style = effective_style(&boundaries, &source.starts, target)?;
    let Some(table) = source.data_table else {
        return invalid("text storage must contain one paragraph-data table");
    };
    let entries = storage::data_entries(table, budget)?;
    let level = usize::from(level_of(&source, &entries, target)?.get());
    Ok(Context {
        storage: source,
        boundaries,
        level,
        style,
    })
}

fn read(
    styles: &(impl Styles + ?Sized),
    style: u64,
    level: usize,
    kind: PropertyKind,
    budget: &mut Budget,
) -> Result<Property> {
    let invalid_value = |_| Error::InvalidSource {
        reason: "native list value is out of range",
    };
    Ok(match kind {
        PropertyKind::Bullet => Property::Bullet(
            ParagraphListBullet::from_boxed(strings(styles, style, budget)?[level].clone().into())
                .map_err(invalid_value)?,
        ),
        PropertyKind::BulletGeometry => {
            if resolve(styles, style, budget)? != ParagraphList::Bullet {
                return Err(Error::Refused(Refusal::NotApplicable));
            }
            let geometry = geometries(styles, style, budget)?[level];
            Property::BulletGeometry(ParagraphListBulletGeometry::new(
                ParagraphListBulletScale::from_ratio(
                    geometry
                        .scale
                        .unwrap_or(ParagraphListBulletScale::ONE.ratio()),
                )
                .map_err(invalid_value)?,
                ParagraphListBulletBaselineOffset::from_points(
                    geometry
                        .baseline_offset
                        .unwrap_or(ParagraphListBulletBaselineOffset::ZERO.points()),
                )
                .map_err(invalid_value)?,
            ))
        },
        PropertyKind::Indentation => {
            let indents = floats(styles, style, false, budget)?;
            let text_indents = floats(styles, style, true, budget)?;
            Property::Indentation(ParagraphListIndentation::new(
                ParagraphListLabelIndent::from_points(indents[level]).map_err(invalid_value)?,
                ParagraphListTextGap::from_em(text_indents[level]).map_err(invalid_value)?,
            ))
        },
        PropertyKind::LabelColor => Property::LabelColor(label_color(styles, style, budget)?),
        PropertyKind::NumberFormat => Property::NumberFormat(style::number_format_from_native(
            number_types(styles, style, budget)?[level],
        )?),
        PropertyKind::NumberScale => {
            if resolve(styles, style, budget)? != ParagraphList::Numbered {
                return Err(Error::Refused(Refusal::NotApplicable));
            }
            Property::NumberScale(
                ParagraphListNumberScale::from_ratio(
                    geometries(styles, style, budget)?[level]
                        .scale
                        .unwrap_or(ParagraphListNumberScale::ONE.ratio()),
                )
                .map_err(invalid_value)?,
            )
        },
        PropertyKind::NumberTiering => {
            Property::NumberTiering(if tiered(styles, style, budget)?[level] {
                ParagraphListNumberTiering::Tiered
            } else {
                ParagraphListNumberTiering::Flat
            })
        },
    })
}

/// The effective value of `kind` for the paragraph at `paragraph`.
///
/// # Errors
///
/// [`Error::Refused`] when `paragraph` is not a paragraph start or the
/// paragraph's list kind has no such property; [`Error::InvalidSource`]
/// otherwise.
pub fn property(
    storage: &[u8],
    styles: &(impl Styles + ?Sized),
    paragraph: TextPosition,
    kind: PropertyKind,
    limits: WireLimits,
) -> Result<Property> {
    let mut budget = Budget::new(limits);
    let context = context(storage, paragraph, &mut budget)?;
    read(styles, context.style, context.level, kind, &mut budget)
}

/// The value a reset restores for `kind` at the paragraph.
///
/// # Errors
///
/// As [`property`].
pub fn standard(
    storage: &[u8],
    styles: &(impl Styles + ?Sized),
    paragraph: TextPosition,
    kind: PropertyKind,
    limits: WireLimits,
) -> Result<Property> {
    let mut budget = Budget::new(limits);
    let context = context(storage, paragraph, &mut budget)?;
    standard_of(styles, context.style, context.level, kind, &mut budget)
}

fn standard_of(
    styles: &(impl Styles + ?Sized),
    style: u64,
    level: usize,
    kind: PropertyKind,
    budget: &mut Budget,
) -> Result<Property> {
    let level = ParagraphListLevel::new(u8::try_from(level).unwrap_or(u8::MAX)).map_err(|_| {
        Error::InvalidSource {
            reason: "native paragraph list level is out of range",
        }
    })?;
    Ok(match kind {
        PropertyKind::Bullet => Property::Bullet(ParagraphListBullet::default()),
        PropertyKind::BulletGeometry => {
            Property::BulletGeometry(ParagraphListBulletGeometry::standard(level))
        },
        PropertyKind::Indentation => Property::Indentation(
            ParagraphListIndentation::standard(resolve(styles, style, budget)?, level)
                .map_err(|_| Error::Refused(Refusal::NotApplicable))?,
        ),
        PropertyKind::LabelColor => Property::LabelColor(ParagraphListLabelColor::Automatic),
        PropertyKind::NumberFormat => Property::NumberFormat(ParagraphListNumberFormat::DECIMAL),
        PropertyKind::NumberScale => Property::NumberScale(ParagraphListNumberScale::ONE),
        PropertyKind::NumberTiering => Property::NumberTiering(ParagraphListNumberTiering::Flat),
    })
}

/// What a list change does to one paragraph's list style.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Plan {
    /// Nothing changes.
    Unchanged,
    /// Replace the payload of the paragraph's own variation in place.
    Rewrite {
        /// The variation.
        style: u64,
        /// Its new payload.
        payload: Vec<u8>,
    },
    /// Store a new variation of `parent` (references `[parent]`), register
    /// it in `stylesheet`, and style the paragraph with it
    /// ([`boundaries_with_style`], [`replace_boundaries`]).
    Derive {
        /// The paragraph's current style, the variation's parent.
        parent: u64,
        /// The stylesheet.
        stylesheet: u64,
        /// The variation payload.
        payload: Vec<u8>,
    },
    /// Style the paragraph with `parent` again and remove the variation
    /// `style`, which no longer overrides anything.
    Collapse {
        /// The removed variation.
        style: u64,
        /// Its parent.
        parent: u64,
        /// The stylesheet.
        stylesheet: u64,
    },
}

/// A requested property change.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Change {
    /// Set the property.
    Set(Property),
    /// Set the property to its standard value; after publishing a changed
    /// plan the caller runs [`plan_collapse`] for the same kind.
    Reset(PropertyKind),
}

fn arrays_for(
    styles: &(impl Styles + ?Sized),
    style: u64,
    level: usize,
    property: &Property,
    budget: &mut Budget,
) -> Result<Arrays> {
    Ok(match property {
        Property::Bullet(bullet) => {
            let mut values = strings(styles, style, budget)?;
            values[level] = bullet.as_str().to_owned();
            Arrays::Strings(values)
        },
        Property::BulletGeometry(geometry) => {
            let mut values = geometries(styles, style, budget)?;
            values[level].scale = Some(geometry.scale.ratio());
            values[level].baseline_offset = Some(geometry.baseline_offset.points());
            values[level].scale_with_text = Some(true);
            Arrays::Geometries(values)
        },
        Property::Indentation(indentation) => {
            let mut indents = floats(styles, style, false, budget)?;
            let mut text_indents = floats(styles, style, true, budget)?;
            indents[level] = indentation.label_from_margin.points();
            text_indents[level] = indentation.text_from_label.em();
            Arrays::Indentation(indents, text_indents)
        },
        Property::LabelColor(color) => Arrays::LabelColor(*color),
        Property::NumberFormat(format) => {
            let mut values = number_types(styles, style, budget)?;
            values[level] = style::number_format_to_native(*format);
            Arrays::NumberTypes(values)
        },
        Property::NumberScale(scale) => {
            let mut values = geometries(styles, style, budget)?;
            values[level].scale = Some(scale.ratio());
            Arrays::Geometries(values)
        },
        Property::NumberTiering(tiering) => {
            let mut values = tiered(styles, style, budget)?;
            values[level] = matches!(tiering, ParagraphListNumberTiering::Tiered);
            Arrays::Tiered(values)
        },
    })
}

/// Plan a property change of the paragraph at `paragraph`. `exclusive`
/// says whether a style is referenced by one storage boundary only and is
/// nobody's parent.
///
/// # Errors
///
/// As [`property`].
pub fn plan(
    storage: &[u8],
    styles: &(impl Styles + ?Sized),
    paragraph: TextPosition,
    change: &Change,
    exclusive: impl Fn(u64) -> bool,
    limits: WireLimits,
) -> Result<Plan> {
    let mut budget = Budget::new(limits);
    let context = context(storage, paragraph, &mut budget)?;
    let target = match change {
        Change::Set(property) => property.clone(),
        Change::Reset(kind) => {
            standard_of(styles, context.style, context.level, *kind, &mut budget)?
        },
    };
    let current = read(
        styles,
        context.style,
        context.level,
        target.kind(),
        &mut budget,
    )?;
    if current == target {
        return Ok(Plan::Unchanged);
    }
    let style = context.style;
    let decoded = locate(styles, style, &mut budget)?;
    let stylesheet = stylesheet(styles, style, limits)?;
    let arrays = arrays_for(styles, style, context.level, &target, &mut budget)?;
    let in_place = decoded.base.parent.is_some()
        && isolated(
            &context.boundaries,
            &context.storage.starts,
            paragraph.utf16_index(),
        )
        && exclusive(style);
    if in_place {
        let payload = styles.list_style(style).ok_or(Error::InvalidSource {
            reason: "iWork list style is missing",
        })?;
        return Ok(Plan::Rewrite {
            style,
            payload: style::with_arrays(payload, &decoded, &arrays, limits)?,
        });
    }
    Ok(Plan::Derive {
        parent: style,
        stylesheet,
        payload: style::variation(style, stylesheet, arrays.overrides(), limits)?,
    })
}

/// Plan removing the paragraph variation's own `kind` override once it
/// equals its parent's standard value (the host's reset clean-up, run
/// after the reset's [`plan`] is published).
///
/// # Errors
///
/// As [`property`].
pub fn plan_collapse(
    storage: &[u8],
    styles: &(impl Styles + ?Sized),
    paragraph: TextPosition,
    kind: PropertyKind,
    exclusive: impl Fn(u64) -> bool,
    limits: WireLimits,
) -> Result<Plan> {
    let mut budget = Budget::new(limits);
    let context = context(storage, paragraph, &mut budget)?;
    let style = context.style;
    if !isolated(
        &context.boundaries,
        &context.storage.starts,
        paragraph.utf16_index(),
    ) || !exclusive(style)
    {
        return Ok(Plan::Unchanged);
    }
    let decoded = locate(styles, style, &mut budget)?;
    let parent = if kind == PropertyKind::Bullet {
        match decoded.parent() {
            Some(parent) => parent,
            None => return Ok(Plan::Unchanged),
        }
    } else {
        parent_of(&decoded)?
    };
    let stylesheet = stylesheet(styles, style, limits)?;
    let collapse = Plan::Collapse {
        style,
        parent,
        stylesheet,
    };
    if kind == PropertyKind::Bullet {
        if decoded.override_count != Some(1) || decoded.strings.is_empty() {
            return Ok(Plan::Unchanged);
        }
        if strings(styles, style, &mut budget)? != strings(styles, parent, &mut budget)? {
            return Ok(Plan::Unchanged);
        }
        return Ok(collapse);
    }
    let parent_value = read(styles, parent, context.level, kind, &mut budget)?;
    let standard = standard_of(styles, parent, context.level, kind, &mut budget)?;
    let parent_is_standard = match (&parent_value, &standard) {
        (Property::BulletGeometry(value), Property::BulletGeometry(standard)) => {
            let raw = geometries(styles, parent, &mut budget)?[context.level];
            raw.scale.unwrap_or(ParagraphListBulletScale::ONE.ratio()) == standard.scale.ratio()
                && raw
                    .baseline_offset
                    .unwrap_or(ParagraphListBulletBaselineOffset::ZERO.points())
                    == standard.baseline_offset.points()
                && value.scale.ratio().is_finite()
        },
        (Property::Indentation(_), Property::Indentation(standard)) => {
            let indents = floats(styles, parent, false, &mut budget)?;
            let text_indents = floats(styles, parent, true, &mut budget)?;
            indents[context.level] == standard.label_from_margin.points()
                && text_indents[context.level] == standard.text_from_label.em()
        },
        (value, standard) => value == standard,
    };
    if !parent_is_standard {
        return Ok(Plan::Unchanged);
    }
    let Some(override_count) = decoded.override_count else {
        return Ok(Plan::Unchanged);
    };
    let (direct, present, direct_count) = match kind {
        PropertyKind::BulletGeometry | PropertyKind::NumberScale => {
            (Direct::Geometries, !decoded.geometries.is_empty(), 1)
        },
        PropertyKind::Indentation => {
            let count = u32::from(!decoded.indents.is_empty())
                + u32::from(!decoded.text_indents.is_empty());
            (Direct::Indentation, count > 0, count)
        },
        PropertyKind::LabelColor => (
            Direct::LabelColor,
            decoded.font_color_null.is_some() || decoded.font_color.is_some(),
            1,
        ),
        PropertyKind::NumberFormat => (Direct::NumberTypes, !decoded.number_types.is_empty(), 1),
        PropertyKind::NumberTiering => (Direct::Tiered, !decoded.tiered_numbers.is_empty(), 1),
        PropertyKind::Bullet => return Ok(Plan::Unchanged),
    };
    if !present {
        return Ok(Plan::Unchanged);
    }
    if override_count == direct_count {
        return Ok(collapse);
    }
    let payload = styles.list_style(style).ok_or(Error::InvalidSource {
        reason: "iWork list style is missing",
    })?;
    Ok(match style::without(payload, &decoded, direct)? {
        Some(payload) => Plan::Rewrite { style, payload },
        None => Plan::Unchanged,
    })
}

/// The stylesheet's preset object for `preset`: the smallest identifier
/// among `candidates` (the list styles stored with the stylesheet) whose
/// stylesheet is `stylesheet` and that matches the canonical preset.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed candidate.
pub fn find_preset<'a>(
    candidates: impl IntoIterator<Item = (u64, &'a [u8])>,
    stylesheet: u64,
    preset: ParagraphList,
    limits: WireLimits,
) -> Result<Option<u64>> {
    let mut budget = Budget::new(limits);
    let mut found = None::<u64>;
    for (identifier, payload) in candidates {
        let style = ListStyle::decode(payload, &mut budget)?;
        if style.base.stylesheet.map(|reference| reference.identifier) == Some(stylesheet)
            && style::matches_preset(&style, preset)
        {
            found = Some(found.map_or(identifier, |current| current.min(identifier)));
        }
    }
    Ok(found)
}

/// The canonical named preset object `preset` of `stylesheet` (no object
/// references; register it in the stylesheet's styles).
///
/// # Errors
///
/// Resource errors as configured.
pub fn preset_payload(
    stylesheet: u64,
    preset: ParagraphList,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    Ok(style::preset_payload(stylesheet, preset, limits)?)
}

/// Normalize list placements against the storage's paragraph starts (the
/// host's `normalize_placements`): they must begin at zero, strictly
/// increase and sit on paragraph starts; equal neighbours coalesce.
///
/// # Errors
///
/// [`Error::Refused`] for invalid placements.
pub fn normalize(
    storage: &[u8],
    placements: &[ParagraphListPlacement],
    limits: WireLimits,
) -> Result<Vec<ParagraphListPlacement>> {
    let starts = paragraph_starts(storage, limits)?;
    if placements
        .first()
        .map(|placement| placement.paragraph.utf16_index())
        != Some(0)
    {
        return Err(Error::Refused(Refusal::InvalidPlacements));
    }
    let mut normalized: Vec<ParagraphListPlacement> = Vec::new();
    let mut previous = None;
    for &placement in placements {
        let index = placement.paragraph.utf16_index();
        if previous.is_some_and(|previous| previous >= index) {
            return Err(Error::Refused(Refusal::InvalidPlacements));
        }
        if starts.binary_search(&index).is_err() {
            return Err(Error::Refused(Refusal::NotAParagraphStart));
        }
        previous = Some(index);
        if normalized
            .last()
            .is_some_and(|last| last.list == placement.list)
        {
            continue;
        }
        normalized.push(placement);
    }
    Ok(normalized)
}
