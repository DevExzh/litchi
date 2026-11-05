//! Storage-scoped plain-text Drop Cap engine (record 0374).
//!
//! A storage's `table_drop_cap_style` (field 28) names, per paragraph
//! start, a `TSWP.DropCapStyleArchive` (10024). A paragraph shows a drop
//! cap when it names a variation of its stylesheet's Drop Cap base style
//! (found by the format package in its theme's Drop Cap presets). The
//! engine reads the placements ([`placements`]) and a variation's model
//! ([`model`]), plans a set or removal as the retired host (commit `4d8a3e161f`) did
//! ([`plan`]), encodes canonical variations ([`variation`]) and repoints
//! one paragraph's table entry ([`point`]).

use litchi_iwa_common::wire::{
    patch_length_delimited_field, patch_varint_field, rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_text::paragraph::drop_cap::{
    CharacterCount, CharacterScale, CornerRadius, DropCap, LineCount, Outdent, Padding,
    RaisedLines, Wrap,
};
use litchi_iwa_text::position::TextPosition;
use thiserror::Error as ThisError;

use crate::native::{
    self, Budget, NativeError, NativeResult, Writer, has_exact_fields, required_payload,
};

/// `TSWP.DropCapStyleArchive`.
pub const MESSAGE_TYPE: u32 = 10_024;
/// The message version a new variation is stored with.
pub const MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
/// The storage field of the Drop Cap table (field-info paths start here).
pub const TABLE_FIELD: u32 = 28;

const STYLESHEET_FIELD: u32 = 2;
const TEXT_FIELD: u32 = 3;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_OVERRIDE_COUNT_FIELD: u32 = 10;
const STYLE_CHARACTER_PROPERTIES_FIELD: u32 = 11;
const STYLE_DROP_CAP_PROPERTIES_FIELD: u32 = 12;
const SUPER_PARENT_FIELD: u32 = 3;
const SUPER_VARIATION_FIELD: u32 = 4;
const SUPER_STYLESHEET_FIELD: u32 = 5;
const PROPERTIES_MODEL_FIELD: u32 = 1;
const MODEL_FIELDS: [u32; 10] = [1, 2, 3, 6, 7, 10, 11, 12, 13, 14];
const TEXT_DROP_CAP_TYPE: i32 = 0;

/// Why a well-formed Drop Cap change is refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The position is not the start of a nonempty paragraph.
    NotAParagraphStart,
}

/// Why a Drop Cap operation failed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A payload is malformed, or a variation is not a canonical plain-text
    /// Drop Cap.
    #[error("invalid Drop Cap source: {reason}")]
    InvalidSource {
        /// What is wrong.
        reason: &'static str,
    },
    /// A well-formed change this engine refuses.
    #[error("Drop Cap change refused: {0:?}")]
    Refused(Refusal),
    /// A configured wire resource was exhausted.
    #[error("Drop Cap {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// The exhausted resource.
        kind: LimitKind,
        /// The attempted aggregate consumption.
        observed: usize,
        /// The caller's configured bound.
        maximum: usize,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} Drop Cap items or bytes")]
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

/// The result of a Drop Cap operation.
pub type Result<T> = std::result::Result<T, Error>;

fn invalid<T>(reason: &'static str) -> Result<T> {
    Err(Error::InvalidSource { reason })
}

/// A storage's Drop Cap table: its stylesheet and per-paragraph styles.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Placements {
    stylesheet: u64,
    entries: Vec<(TextPosition, Option<u64>)>,
}

impl Placements {
    /// The storage's stylesheet.
    #[must_use]
    pub const fn stylesheet(&self) -> u64 {
        self.stylesheet
    }

    /// Each table entry: a paragraph start and the style it names.
    #[must_use]
    pub fn entries(&self) -> &[(TextPosition, Option<u64>)] {
        &self.entries
    }

    /// The style named at `paragraph`.
    #[must_use]
    pub fn style_at(&self, paragraph: TextPosition) -> Option<u64> {
        self.entries
            .iter()
            .find(|(position, _)| *position == paragraph)
            .and_then(|(_, style)| *style)
    }
}

struct Source<'a> {
    text: String,
    stylesheet: Option<u64>,
    table: Option<&'a [u8]>,
}

fn decode<'a>(storage: &'a [u8], budget: &mut Budget) -> NativeResult<Source<'a>> {
    let mut text = String::new();
    let mut stylesheet = None;
    let mut tables = Vec::new();
    native::fields(storage, 0, budget, |field, budget| {
        match field.number {
            TEXT_FIELD => {
                let fragment = std::str::from_utf8(field.bytes()?)
                    .map_err(|_| NativeError::Invalid("storage text is not UTF-8"))?;
                text.try_reserve(fragment.len())
                    .map_err(|_| NativeError::Allocation {
                        amount: fragment.len(),
                    })?;
                text.push_str(fragment);
            },
            STYLESHEET_FIELD => {
                let reference = stylesheet.get_or_insert_with(native::Reference::default);
                reference.merge(field.bytes()?, 1, budget)?;
            },
            TABLE_FIELD => {
                native::reserve(&mut tables, 1)?;
                tables.push(field.bytes()?);
            },
            _ => {},
        }
        Ok(())
    })?;
    let table = match tables.as_slice() {
        [] => None,
        [table] => Some(*table),
        _ => return Err(NativeError::Invalid("storage repeats its Drop Cap table")),
    };
    Ok(Source {
        text,
        stylesheet: stylesheet.map(|reference| reference.identifier),
        table,
    })
}

fn is_boundary(text: &str, index: u32) -> bool {
    let Ok(target) = usize::try_from(index) else {
        return false;
    };
    let mut units = 0usize;
    for character in text.chars() {
        if units == target {
            return true;
        }
        units += character.len_utf16();
        if units > target {
            return false;
        }
    }
    units == target
}

fn is_paragraph_start(text: &str, index: u32) -> bool {
    let Ok(target) = usize::try_from(index) else {
        return false;
    };
    let mut units = 0usize;
    let mut previous = None;
    for character in text.chars() {
        if units == target {
            return (target == 0 || previous == Some('\n')) && character != '\n';
        }
        units += character.len_utf16();
        previous = Some(character);
    }
    false
}

struct Entry {
    index: u32,
    style: Option<u64>,
    raw: Vec<u8>,
}

fn entries(table: &[u8], budget: &mut Budget) -> NativeResult<Vec<Entry>> {
    let mut entries = Vec::new();
    native::fields(table, 1, budget, |field, budget| {
        if field.number != TABLE_ENTRIES_FIELD {
            return Ok(());
        }
        let raw = field.bytes()?;
        let mut index = 0u32;
        let mut style: Option<native::Reference> = None;
        native::fields(raw, 2, budget, |inner, budget| {
            match inner.number {
                ENTRY_INDEX_FIELD => index = inner.uint32()?,
                ENTRY_OBJECT_FIELD => style.get_or_insert_with(native::Reference::default).merge(
                    inner.bytes()?,
                    3,
                    budget,
                )?,
                _ => {},
            }
            Ok(())
        })?;
        native::reserve(&mut entries, 1)?;
        entries.push(Entry {
            index,
            style: style
                .map(|reference| reference.identifier)
                .filter(|identifier| *identifier != 0),
            raw: raw.to_vec(),
        });
        Ok(())
    })?;
    Ok(entries)
}

/// The storage's stylesheet and Drop Cap table.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a storage without stylesheet or with a
/// malformed table; resource errors as configured.
pub fn placements(storage: &[u8], limits: WireLimits) -> Result<Placements> {
    let mut budget = Budget::new(limits);
    let source = decode(storage, &mut budget)?;
    let Some(stylesheet) = source.stylesheet.filter(|identifier| *identifier != 0) else {
        return invalid("text storage has no stylesheet");
    };
    let mut list = Vec::new();
    if let Some(table) = source.table {
        let mut previous = None;
        for entry in entries(table, &mut budget)? {
            if !is_boundary(&source.text, entry.index) {
                return invalid("Drop Cap boundary is not a scalar boundary");
            }
            if previous.is_some_and(|index| index >= entry.index) {
                return invalid("Drop Cap boundaries are not strictly increasing");
            }
            previous = Some(entry.index);
            list.push((
                TextPosition::from_utf16_code_units(entry.index),
                entry.style,
            ));
        }
    }
    Ok(Placements {
        stylesheet,
        entries: list,
    })
}

fn required<T>(value: Option<T>) -> Result<T> {
    value.ok_or(Error::InvalidSource {
        reason: "native Drop Cap model is incomplete",
    })
}

#[derive(Default)]
struct Model {
    kind: Option<i32>,
    lines: Option<u32>,
    raised: Option<u32>,
    deprecated: bool,
    outdent: Option<f64>,
    padding: Option<f64>,
    wrap: Option<i32>,
    shape_enabled: Option<bool>,
    corner_radius: Option<f64>,
    character_scale: Option<f64>,
    characters: Option<u32>,
}

/// A variation's parent, stylesheet and plain-text model.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Variation {
    parent: u64,
    stylesheet: u64,
    drop_cap: DropCap,
}

impl Variation {
    /// The parent style.
    #[must_use]
    pub const fn parent(self) -> u64 {
        self.parent
    }

    /// The stylesheet.
    #[must_use]
    pub const fn stylesheet(self) -> u64 {
        self.stylesheet
    }

    /// The Drop Cap model.
    #[must_use]
    pub const fn drop_cap(self) -> DropCap {
        self.drop_cap
    }
}

/// The model of a canonical plain-text Drop Cap variation.
///
/// # Errors
///
/// [`Error::InvalidSource`] unless the payload is exactly the canonical
/// variation the host writes; resource errors as configured.
#[allow(clippy::too_many_lines, reason = "one check per native field")]
pub fn model(style: &[u8], limits: WireLimits) -> Result<Variation> {
    let mut budget = Budget::new(limits);
    let mut base = native::StyleBase::default();
    let mut override_count = None;
    let mut character = false;
    let mut model = None::<Model>;
    native::fields(style, 0, &mut budget, |field, budget| {
        match field.number {
            STYLE_SUPER_FIELD => base.merge(field.bytes()?, 1, budget)?,
            STYLE_OVERRIDE_COUNT_FIELD => override_count = Some(field.uint32()?),
            STYLE_CHARACTER_PROPERTIES_FIELD => {
                field.bytes()?;
                character = true;
            },
            STYLE_DROP_CAP_PROPERTIES_FIELD => {
                native::fields(field.bytes()?, 1, budget, |inner, budget| {
                    if inner.number == PROPERTIES_MODEL_FIELD {
                        let model = model.get_or_insert_with(Model::default);
                        native::fields(inner.bytes()?, 2, budget, |value, _| {
                            match value.number {
                                1 => model.kind = Some(value.int32()?),
                                2 => model.lines = Some(value.uint32()?),
                                3 => model.raised = Some(value.uint32()?),
                                4 | 5 | 8 | 9 => {
                                    value.float()?;
                                    model.deprecated = true;
                                },
                                6 => model.wrap = Some(value.int32()?),
                                7 => model.shape_enabled = Some(value.bool()?),
                                10 => model.characters = Some(value.uint32()?),
                                11 => model.outdent = Some(value.double()?),
                                12 => model.padding = Some(value.double()?),
                                13 => model.corner_radius = Some(value.double()?),
                                14 => model.character_scale = Some(value.double()?),
                                _ => {},
                            }
                            Ok(())
                        })?;
                    }
                    Ok(())
                })?;
            },
            _ => {},
        }
        Ok(())
    })?;
    let Some(model) = model else {
        return invalid("Drop Cap style has no model");
    };
    if model.kind != Some(TEXT_DROP_CAP_TYPE)
        || model.shape_enabled != Some(false)
        || model.deprecated
    {
        return invalid("unsupported native Drop Cap model");
    }
    let invalid_value = |_| Error::InvalidSource {
        reason: "native Drop Cap value is out of range",
    };
    let drop_cap = DropCap {
        lines: LineCount::new(u8::try_from(required(model.lines)?).map_err(|_| {
            Error::InvalidSource {
                reason: "native Drop Cap line count exceeds u8",
            }
        })?)
        .map_err(invalid_value)?,
        characters: CharacterCount::new(required(model.characters)?).map_err(invalid_value)?,
        raised_lines: RaisedLines::new(u8::try_from(required(model.raised)?).map_err(|_| {
            Error::InvalidSource {
                reason: "native Drop Cap raised-line count exceeds u8",
            }
        })?)
        .map_err(invalid_value)?,
        wrap: match required(model.wrap)? {
            0 => Wrap::Rectangular,
            1 => Wrap::Contour,
            2 => Wrap::None,
            _ => return invalid("unsupported native Drop Cap wrap type"),
        },
        padding: Padding::from_points(required(model.padding)?).map_err(invalid_value)?,
        outdent: Outdent::from_ratio(required(model.outdent)?).map_err(invalid_value)?,
        corner_radius: CornerRadius::from_ratio(required(model.corner_radius)?)
            .map_err(invalid_value)?,
        character_scale: CharacterScale::from_ratio(required(model.character_scale)?)
            .map_err(invalid_value)?,
    };
    let (Some(parent), Some(stylesheet)) = (base.parent, base.stylesheet) else {
        return invalid("Drop Cap style is not a canonical plain-text variation");
    };
    let super_raw = required_payload(style, STYLE_SUPER_FIELD, &mut budget)?;
    let character_raw = required_payload(style, STYLE_CHARACTER_PROPERTIES_FIELD, &mut budget)?;
    let properties_raw = required_payload(style, STYLE_DROP_CAP_PROPERTIES_FIELD, &mut budget)?;
    let model_raw = required_payload(properties_raw, PROPERTIES_MODEL_FIELD, &mut budget)?;
    let canonical = override_count == Some(1)
        && character
        && base.name.is_none()
        && base.style_identifier.is_none()
        && base.is_variation == Some(true)
        && has_exact_fields(
            style,
            &[
                STYLE_SUPER_FIELD,
                STYLE_OVERRIDE_COUNT_FIELD,
                STYLE_CHARACTER_PROPERTIES_FIELD,
                STYLE_DROP_CAP_PROPERTIES_FIELD,
            ],
            &mut budget,
        )?
        && has_exact_fields(
            super_raw,
            &[
                SUPER_PARENT_FIELD,
                SUPER_VARIATION_FIELD,
                SUPER_STYLESHEET_FIELD,
            ],
            &mut budget,
        )?
        && has_exact_fields(character_raw, &[], &mut budget)?
        && has_exact_fields(properties_raw, &[PROPERTIES_MODEL_FIELD], &mut budget)?
        && has_exact_fields(model_raw, &MODEL_FIELDS, &mut budget)?;
    if !canonical {
        return invalid("Drop Cap style is not a canonical plain-text variation");
    }
    Ok(Variation {
        parent: parent.identifier,
        stylesheet: stylesheet.identifier,
        drop_cap,
    })
}

/// The canonical payload of a plain-text Drop Cap variation of `parent`;
/// its message's object references are `[parent]`.
///
/// # Errors
///
/// Resource errors as configured.
pub fn variation(
    parent: u64,
    stylesheet: u64,
    drop_cap: DropCap,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let mut writer = Writer::new(limits);
    writer.message(STYLE_SUPER_FIELD, |writer| {
        writer.message(SUPER_PARENT_FIELD, |writer| {
            native::Reference { identifier: parent }.encode(writer)
        })?;
        writer.bool(SUPER_VARIATION_FIELD, true)?;
        writer.message(SUPER_STYLESHEET_FIELD, |writer| {
            native::Reference {
                identifier: stylesheet,
            }
            .encode(writer)
        })
    })?;
    writer.uint32(STYLE_OVERRIDE_COUNT_FIELD, 1)?;
    writer.bytes(STYLE_CHARACTER_PROPERTIES_FIELD, &[])?;
    writer.message(STYLE_DROP_CAP_PROPERTIES_FIELD, |writer| {
        writer.message(PROPERTIES_MODEL_FIELD, |writer| {
            writer.int32(1, TEXT_DROP_CAP_TYPE)?;
            writer.uint32(2, u32::from(drop_cap.lines.get()))?;
            writer.uint32(3, u32::from(drop_cap.raised_lines.get()))?;
            writer.int32(
                6,
                match drop_cap.wrap {
                    Wrap::Rectangular => 0,
                    Wrap::Contour => 1,
                    Wrap::None => 2,
                },
            )?;
            writer.bool(7, false)?;
            writer.uint32(10, drop_cap.characters.get())?;
            writer.double(11, drop_cap.outdent.ratio())?;
            writer.double(12, drop_cap.padding.points())?;
            writer.double(13, drop_cap.corner_radius.ratio())?;
            writer.double(14, drop_cap.character_scale.ratio())
        })
    })?;
    Ok(writer.into_bytes())
}

/// What a Drop Cap change does to one paragraph.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Plan {
    /// The paragraph already shows the requested Drop Cap (or none).
    Unchanged,
    /// Replace the payload of the paragraph's own variation in place.
    Rewrite {
        /// The variation.
        style: u64,
        /// Its new payload.
        payload: Vec<u8>,
    },
    /// Store a new variation of the base style and point the paragraph at
    /// it ([`point`] from `previous`).
    Derive {
        /// The style the paragraph named before (base or a shared
        /// variation), if any.
        previous: Option<u64>,
        /// The new variation's payload (parent: the base style).
        payload: Vec<u8>,
    },
    /// Clear the paragraph's entry ([`point`] to `None`); an exclusive
    /// variation is then removed from its stylesheet and the package.
    Detach {
        /// The variation.
        style: u64,
        /// Its parent.
        parent: u64,
        /// Whether only this paragraph uses it.
        exclusive: bool,
    },
}

/// A requested Drop Cap change.
#[derive(Debug, Clone, Copy, PartialEq)]
#[non_exhaustive]
pub enum Change {
    /// Show `DropCap` at the paragraph.
    Set(DropCap),
    /// Show none.
    Remove,
}

/// Plan `change` at `paragraph` of `storage`.
///
/// `base` is the stylesheet's Drop Cap base style; `style` resolves a
/// variation payload; `exclusive` says whether a variation is used by this
/// storage entry only and is nobody's parent.
///
/// # Errors
///
/// [`Error::Refused`] when `paragraph` is not a nonempty paragraph start;
/// [`Error::InvalidSource`] for a malformed table or a variation that is
/// not canonical or not of the storage's stylesheet.
pub fn plan<'a>(
    storage: &[u8],
    paragraph: TextPosition,
    base: u64,
    style: impl Fn(u64) -> Option<&'a [u8]>,
    exclusive: impl Fn(u64) -> bool,
    change: Change,
    limits: WireLimits,
) -> Result<Plan> {
    let mut budget = Budget::new(limits);
    let source = decode(storage, &mut budget)?;
    let placements = placements(storage, limits)?;
    if !is_paragraph_start(&source.text, paragraph.utf16_index()) {
        return Err(Error::Refused(Refusal::NotAParagraphStart));
    }
    let named = placements.style_at(paragraph);
    let current = named.filter(|identifier| *identifier != base);
    let current = current
        .map(|identifier| {
            let payload = style(identifier).ok_or(Error::InvalidSource {
                reason: "Drop Cap style is missing",
            })?;
            let existing = model(payload, limits)?;
            if existing.stylesheet != placements.stylesheet {
                return invalid("Drop Cap style is not owned by the storage's stylesheet");
            }
            Ok((identifier, existing))
        })
        .transpose()?;
    match change {
        Change::Set(drop_cap) => {
            if let Some((identifier, existing)) = current {
                if existing.drop_cap == drop_cap {
                    return Ok(Plan::Unchanged);
                }
                if exclusive(identifier) {
                    return Ok(Plan::Rewrite {
                        style: identifier,
                        payload: variation(
                            existing.parent,
                            placements.stylesheet,
                            drop_cap,
                            limits,
                        )?,
                    });
                }
            }
            Ok(Plan::Derive {
                previous: named,
                payload: variation(base, placements.stylesheet, drop_cap, limits)?,
            })
        },
        Change::Remove => Ok(match current {
            None => Plan::Unchanged,
            Some((identifier, existing)) => Plan::Detach {
                style: identifier,
                parent: existing.parent,
                exclusive: exclusive(identifier),
            },
        }),
    }
}

/// A repointed storage and the change of its message's references (also
/// applied to the field infos whose path starts at [`TABLE_FIELD`]).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Rewrite {
    payload: Vec<u8>,
    added: Option<u64>,
    removed: Option<u64>,
}

impl Rewrite {
    /// The storage payload.
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    /// Take the storage payload.
    #[must_use]
    pub fn into_payload(self) -> Vec<u8> {
        self.payload
    }

    /// The style the references gain.
    #[must_use]
    pub const fn added(&self) -> Option<u64> {
        self.added
    }

    /// The style the references lose.
    #[must_use]
    pub const fn removed(&self) -> Option<u64> {
        self.removed
    }
}

fn reference_bytes(identifier: u64, limits: WireLimits) -> NativeResult<Vec<u8>> {
    let mut writer = Writer::new(limits);
    native::Reference { identifier }.encode(&mut writer)?;
    Ok(writer.into_bytes())
}

/// Point the table entry of `paragraph` from `previous` to `next`, adding
/// the entry when the paragraph has none.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the entry does not name `previous`, the
/// table repeats the paragraph or the storage is malformed.
pub fn point(
    storage: &[u8],
    paragraph: TextPosition,
    previous: Option<u64>,
    next: Option<u64>,
    limits: WireLimits,
) -> Result<Rewrite> {
    let mut budget = Budget::new(limits);
    let source = decode(storage, &mut budget)?;
    let table = source.table.unwrap_or_default();
    let mut list = entries(table, &mut budget)?;
    let count = |list: &[Entry], identifier: Option<u64>| {
        identifier.map_or(0, |identifier| {
            list.iter()
                .filter(|entry| entry.style == Some(identifier))
                .count()
        })
    };
    let old_before = count(&list, previous);
    let new_before = count(&list, next);
    let matches: Vec<usize> = list
        .iter()
        .enumerate()
        .filter(|(_, entry)| entry.index == paragraph.utf16_index())
        .map(|(position, _)| position)
        .collect();
    if matches.len() > 1 {
        return invalid("storage has duplicate Drop Cap boundaries");
    }
    if matches.first().and_then(|position| list[*position].style) != previous
        || (matches.is_empty() && previous.is_some())
    {
        return invalid("Drop Cap reference changed unexpectedly");
    }
    let wire = native::common;
    if let Some(position) = matches.first().copied() {
        let entry = &mut list[position];
        entry.raw = match (previous, next) {
            (Some(_), Some(new)) => {
                let reference = required_payload(&entry.raw, ENTRY_OBJECT_FIELD, &mut budget)?;
                let patched =
                    patch_varint_field(reference, REFERENCE_IDENTIFIER_FIELD, true, Some(new))
                        .map_err(wire)?;
                patch_length_delimited_field(&entry.raw, ENTRY_OBJECT_FIELD, true, Some(&patched))
                    .map_err(wire)?
            },
            (Some(_), None) => {
                patch_length_delimited_field(&entry.raw, ENTRY_OBJECT_FIELD, true, None)
                    .map_err(wire)?
            },
            (None, Some(new)) => patch_length_delimited_field(
                &entry.raw,
                ENTRY_OBJECT_FIELD,
                false,
                Some(&reference_bytes(new, limits)?),
            )
            .map_err(wire)?,
            (None, None) => entry.raw.clone(),
        };
        entry.style = next;
    } else if let Some(new) = next {
        let mut writer = Writer::new(limits);
        writer.uint32(ENTRY_INDEX_FIELD, paragraph.utf16_index())?;
        writer.message(ENTRY_OBJECT_FIELD, |writer| {
            native::Reference { identifier: new }.encode(writer)
        })?;
        native::reserve(&mut list, 1)?;
        list.push(Entry {
            index: paragraph.utf16_index(),
            style: next,
            raw: writer.into_bytes(),
        });
    }
    list.sort_by_key(|entry| entry.index);
    if list.windows(2).any(|pair| pair[0].index == pair[1].index) {
        return invalid("storage has duplicate Drop Cap boundaries");
    }
    let old_after = count(&list, previous);
    let new_after = count(&list, next);
    let encoded: Vec<Vec<u8>> = list.into_iter().map(|entry| entry.raw).collect();
    let rewritten = rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &encoded)
        .map_err(wire)?;
    budget.admit_bytes(storage.len())?;
    let payload = patch_length_delimited_field(
        storage,
        TABLE_FIELD,
        source.table.is_some(),
        Some(&rewritten),
    )
    .map_err(wire)?;
    native::admit(
        LimitKind::OutputBytes,
        payload.len(),
        limits.max_output_bytes(),
    )?;
    Ok(Rewrite {
        payload,
        removed: previous.filter(|_| old_before > 0 && old_after == 0),
        added: next.filter(|_| new_before == 0 && new_after > 0),
    })
}
