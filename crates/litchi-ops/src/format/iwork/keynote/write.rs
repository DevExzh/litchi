//! Keynote writes over the leaf's focused transactions (record 0298).
//!
//! Every verb stages its changes in a [`Pending`] set first and publishes
//! them in as few leaf commits as the families allow: one slide-drawable
//! edit for every frame, drawable text, drawable removal and new drawable;
//! one slide-text or speaker-notes edit per text (a single minimal span
//! replacement however many of its paragraphs changed); one slide-creation
//! edit for consecutive new slides. Batches coalesce the same way
//! ([`super::coalesce`]).

use std::collections::BTreeMap;
use std::sync::Arc;

use litchi_core::selector::Path;
use litchi_core::{FileFormat, Position};
use litchi_keynote::slide::creation::LayoutSelector;
use litchi_keynote::slide::drawable::{
    Created, Geometry, NewTable, Preset, Style, TableCellValue, TextSpan as DrawableSpan,
};
use litchi_keynote::slide::placeholder::Kind as Role;
use litchi_keynote::{
    DrawableSelector, Package, SlideSelector, TableSelector, TextPosition, TextSpan,
};

use super::super::write::{
    Chain, Commit, LeafPatch, bound, destination, element_refusal, finish, key_refusal,
    leaf_refusal, node_at, refusal, resolve, unwritable,
};
use super::super::{Iwork, LeafError, Node, Owner};
use super::{BODY, DRAWABLE, Keynote, NONE, NOTES, Patch, TITLE};
use crate::error::{Code, Error};
use crate::new_node::NewNode;
use crate::path::{IntoPath, PathArg};
use crate::props::{PropValue, Props};
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::{LengthSpec, Scalar, ValueType};
use crate::{Placement, Result};

/// The leaf failure of a Keynote transaction.
pub(super) fn failure(error: LeafError) -> Error {
    leaf_refusal(error, FileFormat::Keynote)
}

/// Record one committed Keynote transaction on the chain.
pub(super) fn push(chain: &mut Chain<Package>, package: Package, patch: Patch) {
    let changed = !patch.is_noop();
    chain.push(package, LeafPatch::Keynote(patch), changed);
}

/// One text of a slide: the title or body placeholder, the speaker notes,
/// or a drawable's own storage.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct Text {
    pub(super) slide: usize,
    pub(super) container: usize,
    /// The drawable's position for [`DRAWABLE`]; [`NONE`] otherwise.
    pub(super) drawable: usize,
}

impl Text {
    /// The text a placeholder, notes, shape or paragraph item belongs to.
    pub(super) fn of(node: &Node<'_>) -> Option<Self> {
        let [slide, container, _, drawable] = node.item().leaf().map(|value| value as usize);
        let container = match node.item().type_name() {
            "placeholder" | "notes" | "paragraph" => container,
            "shape" => DRAWABLE,
            _ => return None,
        };
        Some(Self {
            slide,
            container,
            drawable: if container == DRAWABLE {
                drawable
            } else {
                NONE
            },
        })
    }

    /// The item holding this text in `file`'s tree.
    fn item(self, file: &Iwork) -> Option<&super::super::tree::Item> {
        self.at(file).map(|at| file.item(at))
    }

    /// The place of the item holding this text in `file`'s tree.
    fn at(self, file: &Iwork) -> Option<super::super::At> {
        let slide = file.kids(super::super::At::ROOT).get(self.slide)?;
        // The stand-ins: reading a table would decode it (record 0312).
        file.kids(slide).iter().find(|&child| {
            let item = file.stand_in(child);
            let leaf = item.leaf().map(|value| value as usize);
            match (item.type_name(), self.container) {
                ("placeholder", TITLE | BODY) | ("notes", NOTES) => leaf[1] == self.container,
                ("shape", DRAWABLE) => leaf[3] == self.drawable,
                _ => false,
            }
        })
    }
}

/// A new drawable staged on a slide.
#[derive(Debug)]
pub(super) enum Add {
    Text {
        slide: usize,
        preset: Option<Preset>,
        geometry: Geometry,
        text: String,
    },
    Image {
        slide: usize,
        bytes: Arc<[u8]>,
        name: String,
        geometry: Geometry,
    },
    /// A table cloned from the deck's first slide table (record 0329).
    Table { slide: usize, table: NewTable },
}

impl Add {
    pub(super) const fn slide(&self) -> usize {
        match self {
            Self::Text { slide, .. } | Self::Image { slide, .. } | Self::Table { slide, .. } => {
                *slide
            },
        }
    }

    /// The element kind the new drawable reads as.
    pub(super) const fn element(&self) -> &'static str {
        match self {
            Self::Text { .. } => "shape",
            Self::Image { .. } => "picture",
            Self::Table { .. } => "table",
        }
    }
}

/// Staged changes, each with the batch item index that staged it (for
/// error attribution; 0 outside batches).
#[derive(Debug, Default)]
pub(super) struct Pending {
    /// Texts to publish: text → (new text, item).
    texts: BTreeMap<Text, (Staged, usize)>,
    /// Frames to publish: (slide, drawable) → (frame, item).
    frames: BTreeMap<(usize, usize), (Geometry, usize)>,
    /// Drawables to remove: (slide, drawable) → item.
    removals: BTreeMap<(usize, usize), usize>,
    /// New drawables, in order.
    adds: Vec<(Add, usize)>,
    /// Slide visibility: slide → (hidden, item).
    hidden: BTreeMap<usize, (bool, usize)>,
    /// Character styles: (slide, drawable, first char, end char) → (style,
    /// item) (record 0309).
    styles: BTreeMap<(usize, usize, usize, usize), (Style, usize)>,
    /// Table cells: (slide, drawable, row, column) → (value, item).
    cells: BTreeMap<(usize, usize, u32, u32), (TableCellValue, usize)>,
    /// Root table-cell comments: (slide, table, row, column) → (text, item).
    comments: BTreeMap<(usize, usize, u32, u32), (Option<String>, usize)>,
    /// Placeholder visibility: (slide, TITLE or BODY) → (hidden, item).
    visibility: BTreeMap<(usize, usize), (bool, usize)>,
    /// What the sets of each staged text touched (record 0314).
    groups: BTreeMap<Text, TextGroup>,
}

/// A staged text.
#[derive(Debug)]
enum Staged {
    /// A whole new text.
    Text(String),
    /// The base of the text's [`TextGroup`] with the paragraphs its sets
    /// replaced (record 0340): a paragraph set costs its own paragraph, and
    /// the text is joined once, when the run publishes.
    Paragraphs,
}

/// The sets one text received in a run: its text before them, the region
/// they touched, and the largest region the leaf admitted for one combined
/// splice (record 0314).
#[derive(Debug)]
struct TextGroup {
    base: String,
    /// The byte offset of each paragraph of `base`, then `base.len() + 1`.
    starts: Vec<usize>,
    /// The UTF-16 offset of each paragraph of `base`, then its UTF-16
    /// length + 1.
    units: Vec<usize>,
    /// The paragraphs set while the text is [`Staged::Paragraphs`]; `None`
    /// keeps the base's paragraph.
    lines: Vec<Option<String>>,
    /// How many of `lines` differ from the base's paragraph.
    differs: usize,
    touched: Touched,
    admitted: Option<Touched>,
}

impl TextGroup {
    fn new(base: String) -> Self {
        let mut starts = vec![0];
        let mut units = vec![0];
        let mut unit = 0_usize;
        for (at, byte) in base.char_indices() {
            if byte == '\n' {
                starts.push(at + 1);
                units.push(unit + 1);
            }
            unit += byte.len_utf16();
        }
        starts.push(base.len() + 1);
        units.push(unit + 1);
        let count = starts.len() - 1;
        Self {
            base,
            starts,
            units,
            lines: std::iter::repeat_with(|| None).take(count).collect(),
            differs: 0,
            touched: Touched::Lines(0, 0),
            admitted: None,
        }
    }

    /// The number of paragraphs of the base.
    fn count(&self) -> usize {
        self.lines.len()
    }

    /// Paragraph `line` of the base.
    fn base_line(&self, line: usize) -> Option<&str> {
        let (start, end) = (*self.starts.get(line)?, *self.starts.get(line + 1)?);
        self.base.get(start..end - 1)
    }

    /// Paragraph `line` of the staged text.
    fn line(&self, line: usize) -> Option<&str> {
        match self.lines.get(line)? {
            Some(set) => Some(set),
            None => self.base_line(line),
        }
    }

    /// Replace paragraph `line` of the staged text.
    fn set_line(&mut self, line: usize, value: String) {
        let Some(base) = self.base_line(line) else {
            return;
        };
        let now = value != base;
        let before = self.lines[line].as_deref().is_some_and(|held| held != base);
        self.differs = self.differs + usize::from(now) - usize::from(before);
        self.lines[line] = now.then_some(value);
    }

    /// The staged text, joined.
    fn text(&self) -> String {
        if self.differs == 0 {
            return self.base.clone();
        }
        let mut text = String::with_capacity(self.base.len());
        for line in 0..self.count() {
            if line > 0 {
                text.push('\n');
            }
            text.push_str(self.line(line).unwrap_or_default());
        }
        text
    }

    /// The UTF-16 region of the base that `touched` covers, both ends
    /// inclusive.
    fn hull(&self, touched: Touched) -> Option<std::ops::Range<usize>> {
        let (first, last) = match touched {
            Touched::Whole => (0, self.count().checked_sub(1)?),
            Touched::Lines(first, last) => (first, last),
        };
        let start = *self.units.get(first)?;
        let end = self.units.get(last + 1)? - 1;
        (start <= end).then_some(start..end)
    }
}

/// The region of a text its sets touched.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Touched {
    /// Paragraphs `first..=last` of the base text.
    Lines(usize, usize),
    /// The whole text.
    Whole,
}

impl Touched {
    /// Whether `self` lies within `other`.
    const fn within(self, other: Self) -> bool {
        match (self, other) {
            (Self::Lines(a, b), Self::Lines(c, d)) => c <= a && b <= d,
            (_, Self::Whole) => true,
            (Self::Whole, Self::Lines(..)) => false,
        }
    }

    const fn union(self, other: Self) -> Self {
        match (self, other) {
            (Self::Lines(a, b), Self::Lines(c, d)) => {
                Self::Lines(if a < c { a } else { c }, if b > d { b } else { d })
            },
            _ => Self::Whole,
        }
    }
}

/// Whether the leaf admits one combined splice over `hull` of `text`.
fn leaf_admits(package: &Package, text: Text, hull: std::ops::Range<usize>) -> bool {
    let slide = SlideSelector::index(text.slide);
    match text.container {
        TITLE | BODY => {
            let role = if text.container == TITLE {
                Role::Title
            } else {
                Role::Body
            };
            package
                .edit_slide_text(slide, role)
                .and_then(|edit| edit.admits_combined_splice(hull))
                .unwrap_or(false)
        },
        NOTES => package
            .edit_slide_notes(slide)
            .and_then(|edit| edit.admits_combined_splice(hull))
            .unwrap_or(false),
        DRAWABLE => package
            .edit_slide_drawables()
            .text_admits_combined_splice(slide, DrawableSelector::index(text.drawable), hull)
            .unwrap_or(false),
        _ => false,
    }
}

/// The widest region of `text` holding `union` that the leaf admits for one
/// combined splice: the whole text, else `union` widened paragraph by
/// paragraph as far as the leaf admits; `None` when it refuses `union`.
///
/// The leaf's answer holds for every region within an admitted one, so one
/// search per text group replaces a leaf call per set (record 0340).
fn widest_admitted(
    package: &Package,
    text: Text,
    group: &TextGroup,
    union: Touched,
) -> Option<Touched> {
    let admits = |touched| {
        group
            .hull(touched)
            .is_some_and(|hull| leaf_admits(package, text, hull))
    };
    if admits(Touched::Whole) {
        return Some(Touched::Whole);
    }
    let Touched::Lines(first, last) = union else {
        return None;
    };
    if !admits(union) {
        return None;
    }
    // The first paragraph the widened region can start at.
    let (mut low, mut high) = (0, first);
    while low < high {
        let middle = low + (high - low) / 2;
        if admits(Touched::Lines(middle, last)) {
            high = middle;
        } else {
            low = middle + 1;
        }
    }
    let start = low;
    // The last paragraph it can end at.
    let (mut low, mut high) = (last, group.count().saturating_sub(1).max(last));
    while low < high {
        let middle = high - (high - low) / 2;
        if admits(Touched::Lines(first, middle)) {
            low = middle;
        } else {
            high = middle - 1;
        }
    }
    let widest = Touched::Lines(start, low);
    Some(if widest == union || admits(widest) {
        widest
    } else {
        union
    })
}

/// Whether a set writing `written` over `touched` brings `current`, a whole
/// text staged after `base`, back to `base`: one at a time the sets publish
/// twice, joined they would publish nothing.
fn cancels_text(base: &str, current: &str, touched: Touched, written: &str) -> bool {
    match touched {
        Touched::Whole => written == base && current != base,
        Touched::Lines(line, _) => {
            base.split('\n').nth(line) == Some(written)
                && current.split('\n').nth(line) != Some(written)
                && current
                    .split('\n')
                    .zip(base.split('\n'))
                    .enumerate()
                    .all(|(at, (held, base))| at == line || held == base)
                && current.split('\n').count() == base.split('\n').count()
        },
    }
}

/// What publishing a pending set made.
#[derive(Debug, Default)]
pub(super) struct Published {
    /// The drawables the slide-drawable commit created, in staging order.
    pub(super) created: Vec<Created>,
}

/// A refusal of `key` on a run: only character properties are written.
fn run_refusal(key: &str) -> Error {
    refusal(
        Code::UnsupportedCapability,
        "a Keynote run writes its character properties (bold, italic, underline, size, color, font.latin)",
    )
    .with_property(key.to_owned())
    .with_suggestion("set the paragraph's or the shape's text to change the characters")
}

/// Fold one run property into `style`.
fn run_style(style: &mut Style, key: &str, value: &PropValue<'_>) -> Result<()> {
    let cleared = || {
        refusal(
            Code::InvalidValue,
            "a run's character property cannot be cleared",
        )
        .with_property(key.to_owned())
    };
    match key {
        "bold" => style.bold = Some(super::super::write::boolean(value, "bold")?),
        "italic" => style.italic = Some(super::super::write::boolean(value, "italic")?),
        "underline" => {
            style.underline = Some(super::super::write::boolean(value, "underline")?);
        },
        "size" => {
            let points = value
                .with_coerced(super::FONT_SIZE, |scalar| {
                    Ok(scalar.and_then(|scalar| match scalar {
                        Scalar::Length(length) => Some(length.points()),
                        Scalar::Number(value) => Some(value),
                        #[allow(clippy::cast_precision_loss, reason = "point sizes are small")]
                        Scalar::Int(value) => Some(value as f64),
                        _ => None,
                    }))
                })
                .map_err(|error| error.with_property("size"))?
                .ok_or_else(cleared)?;
            #[allow(
                clippy::cast_possible_truncation,
                reason = "Keynote stores font sizes as 32-bit floats"
            )]
            let points = points as f32;
            if !points.is_finite() || points <= 0.0 {
                return Err(refusal(Code::InvalidValue, "a font size must be positive")
                    .with_property("size".to_owned()));
            }
            style.font_size = Some(points);
        },
        "color" => {
            let color = value
                .with_coerced(super::COLOR, |scalar| {
                    Ok(scalar.and_then(|scalar| match scalar {
                        Scalar::Color(color) if color.transforms().is_empty() => {
                            match color.base() {
                                crate::value::Base::Rgb(rgba) => {
                                    Some([rgba.red(), rgba.green(), rgba.blue()])
                                },
                                _ => None,
                            }
                        },
                        _ => None,
                    }))
                })
                .map_err(|error| error.with_property("color"))?
                .ok_or_else(cleared)?;
            style.color = Some(color);
        },
        _ => {
            let name = super::super::write::text(value, "font.latin")?.ok_or_else(cleared)?;
            style.font_name = Some(name);
        },
    }
    Ok(())
}

/// The value of a length key in points; `null` is refused.
fn points(value: &PropValue<'_>, key: &'static str, extent: bool) -> Result<f32> {
    let ty = ValueType::Length(if extent {
        LengthSpec::EXTENT
    } else {
        LengthSpec::POSITION
    });
    let points = value
        .with_coerced(ty, |scalar| {
            Ok(scalar.and_then(|scalar| match scalar {
                Scalar::Length(length) => Some(length.points()),
                _ => None,
            }))
        })
        .map_err(|error| error.with_property(key))?
        .ok_or_else(|| {
            refusal(Code::InvalidValue, "a frame property cannot be cleared")
                .with_property(key.to_owned())
        })?;
    #[allow(
        clippy::cast_possible_truncation,
        reason = "Keynote stores frames as 32-bit floats"
    )]
    let points = points as f32;
    if !points.is_finite() || (extent && points <= 0.0) {
        return Err(refusal(
            Code::InvalidValue,
            "a frame needs finite coordinates and a positive size",
        )
        .with_property(key.to_owned()));
    }
    Ok(points)
}

/// The value of `rotation` in degrees; `null` is refused.
fn degrees(value: &PropValue<'_>) -> Result<f32> {
    let degrees = value
        .with_coerced(ValueType::Number, |scalar| {
            Ok(scalar.and_then(|scalar| match scalar {
                Scalar::Number(value) => Some(value),
                #[allow(clippy::cast_precision_loss, reason = "degrees are small")]
                Scalar::Int(value) => Some(value as f64),
                _ => None,
            }))
        })
        .map_err(|error| error.with_property("rotation"))?
        .ok_or_else(|| {
            refusal(Code::InvalidValue, "a frame property cannot be cleared")
                .with_property("rotation".to_owned())
        })?;
    #[allow(
        clippy::cast_possible_truncation,
        reason = "Keynote stores rotations as 32-bit floats"
    )]
    let degrees = degrees as f32;
    if !degrees.is_finite() {
        return Err(refusal(Code::InvalidValue, "the rotation must be finite")
            .with_property("rotation".to_owned()));
    }
    Ok(degrees)
}

/// The text of a paragraph value: a newline would split the paragraph.
fn paragraph_text(value: &PropValue<'_>) -> Result<String> {
    let value = super::super::write::text(value, "text")?.unwrap_or_default();
    if value.contains('\n') {
        return Err(refusal(
            Code::InvalidValue,
            "a paragraph's text cannot contain a newline, which separates Keynote paragraphs",
        )
        .with_property("text".to_owned())
        .with_suggestion("add each paragraph separately, or set the whole text"));
    }
    Ok(value)
}

/// The frame of an item: its `x`, `y`, `width`, `height` and `rotation`.
fn frame_of(node: &Node<'_>) -> Option<Geometry> {
    let item = node.item();
    let length = |key| match item.value(key) {
        Some(crate::node::Value::Length(length)) =>
        {
            #[allow(
                clippy::cast_possible_truncation,
                reason = "Keynote stores frames as 32-bit floats"
            )]
            Some(length.points() as f32)
        },
        _ => None,
    };
    let mut geometry = Geometry::new(
        length("x")?,
        length("y")?,
        length("width")?,
        length("height")?,
    );
    if let Some(crate::node::Value::Number(rotation)) = item.value("rotation") {
        #[allow(
            clippy::cast_possible_truncation,
            reason = "Keynote stores rotations as 32-bit floats"
        )]
        {
            geometry = geometry.with_rotation(*rotation as f32);
        }
    }
    Some(geometry)
}

impl Pending {
    pub(super) fn is_empty(&self) -> bool {
        self.texts.is_empty()
            && self.frames.is_empty()
            && self.removals.is_empty()
            && self.adds.is_empty()
            && self.hidden.is_empty()
            && self.styles.is_empty()
            && self.cells.is_empty()
            && self.visibility.is_empty()
    }

    /// The new drawables staged on `slide` so far.
    pub(super) fn adds_on(&self, slide: usize) -> impl Iterator<Item = &Add> {
        self.adds
            .iter()
            .map(|(add, _)| add)
            .filter(move |add| add.slide() == slide)
    }

    /// The staged text of `text`, joined; `None` when it is not staged.
    fn staged(&self, text: Text) -> Option<String> {
        match self.texts.get(&text)? {
            (Staged::Text(value), _) => Some(value.clone()),
            (Staged::Paragraphs, _) => self.groups.get(&text).map(TextGroup::text),
        }
    }

    /// The current text of `text`: staged, or the tree's.
    fn current(&self, file: &Iwork, text: Text) -> Result<String> {
        if let Some(value) = self.staged(text) {
            return Ok(value);
        }
        text.item(file)
            .map(|item| item.text().unwrap_or_default().to_owned())
            .ok_or_else(|| refusal(Code::PathNotFound, "the text's element is missing"))
    }

    /// Stage a whole text.
    pub(super) fn stage_text(&mut self, text: Text, value: String, index: usize) {
        self.texts.insert(text, (Staged::Text(value), index));
    }

    /// Record that a set touches `touched` of `text`, opening the text's
    /// group at its current text.
    fn touch(&mut self, file: &Iwork, text: Text, touched: Touched) -> Result<()> {
        if !self.groups.contains_key(&text) {
            let base = self.current(file, text)?;
            let mut group = TextGroup::new(base);
            group.touched = touched;
            self.groups.insert(text, group);
            return Ok(());
        }
        if let Some(group) = self.groups.get_mut(&text) {
            group.touched = group.touched.union(touched);
        }
        Ok(())
    }

    /// Stage paragraph `line` of `text` as `value`.
    fn stage_paragraph(
        &mut self,
        file: &Iwork,
        text: Text,
        line: usize,
        value: String,
        index: usize,
    ) -> Result<()> {
        let missing = || refusal(Code::PathNotFound, "the paragraph is no longer in its text");
        self.touch(file, text, Touched::Lines(line, line))?;
        let group = self
            .groups
            .get_mut(&text)
            .ok_or_else(|| refusal(Code::InternalError, "the text's group is missing"))?;
        match self.texts.get_mut(&text) {
            // A whole text staged before: replace its paragraph.
            Some((Staged::Text(current), at)) => {
                let mut lines = current.split('\n').collect::<Vec<_>>();
                let slot = lines.get_mut(line).ok_or_else(missing)?;
                *slot = &value;
                *current = lines.join("\n");
                *at = index;
            },
            Some((Staged::Paragraphs, at)) => {
                if line >= group.count() {
                    return Err(missing());
                }
                group.set_line(line, value);
                *at = index;
            },
            None => {
                if line >= group.count() {
                    return Err(missing());
                }
                group.set_line(line, value);
                self.texts.insert(text, (Staged::Paragraphs, index));
            },
        }
        Ok(())
    }

    /// Whether setting `key` to `value` on `node` would bring a staged
    /// scalar (a frame, a slide's or a placeholder's visibility) back to its
    /// value in the run's base, or sets a staged table cell again: one at a
    /// time those sets publish twice, joined they would publish nothing (or
    /// the last only).
    fn cancels(&self, node: &Node<'_>, key: &str, value: &PropValue<'_>) -> bool {
        let item = node.item();
        let [slide, slot, _, drawable] = item.leaf().map(|value| value as usize);
        let base_flag = || match item.value("hidden") {
            Some(crate::node::Value::Bool(flag)) => Some(*flag),
            _ => None,
        };
        match (item.type_name(), key) {
            ("slide", "hidden") => self.hidden.get(&slide).is_some_and(|(pending, _)| {
                super::super::write::boolean(value, "hidden")
                    .is_ok_and(|flag| Some(flag) == base_flag() && Some(*pending) != base_flag())
            }),
            ("placeholder", "hidden") => {
                self.visibility
                    .get(&(slide, slot))
                    .is_some_and(|(pending, _)| {
                        super::super::write::boolean(value, "hidden").is_ok_and(|flag| {
                            Some(flag) == base_flag() && Some(*pending) != base_flag()
                        })
                    })
            },
            ("placeholder" | "shape" | "picture", "x" | "y" | "width" | "height" | "rotation") => {
                let (Some((pending, _)), Some(base)) =
                    (self.frames.get(&(slide, drawable)), frame_of(node))
                else {
                    return false;
                };
                let mut geometry = *pending;
                let parsed = match key {
                    "x" => points(value, "x", false).map(|v| geometry.x = v),
                    "y" => points(value, "y", false).map(|v| geometry.y = v),
                    "width" => points(value, "width", true).map(|v| geometry.width = v),
                    "height" => points(value, "height", true).map(|v| geometry.height = v),
                    _ => degrees(value).map(|v| geometry.rotation = Some(v)),
                };
                parsed.is_ok() && geometry == base && *pending != base
            },
            ("tc", _) => {
                let [_, _, row, _] = item.leaf();
                self.cells
                    .keys()
                    .any(|cell| cell.0 == slide && cell.1 == drawable && cell.2 == row)
            },
            _ => false,
        }
    }

    /// Whether a set of `props` on `node` publishes, joined to the staged
    /// changes, what it publishes after them on its own (record 0314):
    /// `false` asks the caller to publish the staged changes first.
    ///
    /// A text that already holds a set takes another while one splice from
    /// its base text to the result is what the sets publish one at a time:
    /// the leaf admits the touched region, the text was not emptied, and the
    /// sets do not cancel out. A text and the character properties of its
    /// slide's runs never share a run.
    pub(super) fn joins(&mut self, file: &Iwork, node: &Node<'_>, props: &Props<'_>) -> bool {
        let item = node.item();
        let line = item.leaf()[2] as usize;
        for (key, value) in props.iter() {
            let touched = match (item.type_name(), key.as_str()) {
                ("placeholder" | "notes" | "shape", "text") => Touched::Whole,
                ("paragraph", "text") => Touched::Lines(line, line),
                ("run", _) => {
                    let slide = item.leaf()[0] as usize;
                    if self.texts.keys().any(|text| text.slide == slide) {
                        return false;
                    }
                    continue;
                },
                _ => {
                    if self.cancels(node, key.as_str(), value) {
                        return false;
                    }
                    continue;
                },
            };
            let Some(text) = Text::of(node) else {
                continue;
            };
            if self.styles.keys().any(|style| style.0 == text.slide) {
                return false;
            }
            let (Some(group), Some((staged, _))) = (self.groups.get(&text), self.texts.get(&text))
            else {
                continue;
            };
            let written = match touched {
                Touched::Whole => {
                    super::super::write::text(value, "text").map(Option::unwrap_or_default)
                },
                Touched::Lines(..) => paragraph_text(value),
            };
            // A refused value is reported by staging.
            let Ok(written) = written else {
                continue;
            };
            let (emptied, cancels) = match staged {
                Staged::Text(current) => (
                    current.is_empty(),
                    cancels_text(&group.base, current, touched, &written),
                ),
                Staged::Paragraphs => {
                    let emptied = group.count() == 1 && group.line(0) == Some("");
                    // The sets cancel out when this one brings the text back
                    // to its base: it writes the base's own text (or
                    // paragraph), and the rest of the text is the base's
                    // already.
                    let cancels = match touched {
                        Touched::Whole => written == group.base && group.differs > 0,
                        Touched::Lines(line, _) => {
                            let base = group.base_line(line);
                            let held = group.line(line);
                            base == Some(written.as_str())
                                && held != Some(written.as_str())
                                && group.differs == usize::from(held != base)
                        },
                    };
                    (emptied, cancels)
                },
            };
            if emptied && !group.base.is_empty() {
                return false;
            }
            if cancels {
                return false;
            }
            let union = group.touched.union(touched);
            if group
                .admitted
                .is_some_and(|admitted| union.within(admitted))
            {
                continue;
            }
            let Some(package) = file.keynote() else {
                return false;
            };
            let Some(admitted) = widest_admitted(package, text, group, union) else {
                return false;
            };
            if let Some(group) = self.groups.get_mut(&text) {
                group.admitted = Some(admitted);
            }
        }
        true
    }

    /// Stage one character property of a run of a text box, shape, title
    /// or body placeholder, or the speaker notes: the run's characters in
    /// its text (records 0309, 0312, 0329).
    fn stage_run_style(
        &mut self,
        file: &Iwork,
        node: &Node<'_>,
        key: &str,
        value: &PropValue<'_>,
        index: usize,
    ) -> Result<()> {
        let [slide, container, line, drawable] = node.item().leaf().map(|value| value as usize);
        let unsupported = || {
            refusal(
                Code::UnsupportedCapability,
                "character properties are written on runs of text boxes, shapes, title and body placeholders and speaker notes",
            )
            .with_property(key.to_owned())
        };
        let notes = container == NOTES;
        if !notes && (container != DRAWABLE || drawable == NONE) {
            return Err(unsupported());
        }
        // The text holding the run: a shape's, a placeholder's whose frame
        // is the run's drawable, or the slide's notes (record 0329).
        let slide_at = file
            .kids(super::super::At::ROOT)
            .get(slide)
            .ok_or_else(|| refusal(Code::PathNotFound, "the run's slide is missing"))?;
        let shape = file
            .kids(slide_at)
            .iter()
            .find(|&child| {
                let item = file.stand_in(child);
                if notes {
                    return item.type_name() == "notes";
                }
                matches!(item.type_name(), "shape" | "placeholder")
                    && item.leaf()[3] as usize == drawable
            })
            .ok_or_else(unsupported)?;
        let text = match file.stand_in(shape).leaf()[1] as usize {
            NOTES if notes => Text {
                slide,
                container: NOTES,
                drawable: NONE,
            },
            slot @ (TITLE | BODY) if file.stand_in(shape).type_name() == "placeholder" => Text {
                slide,
                container: slot,
                drawable: NONE,
            },
            _ => Text {
                slide,
                container: DRAWABLE,
                drawable,
            },
        };
        if self.texts.contains_key(&text) {
            return Err(refusal(
                Code::InvalidArgument,
                "set a text and its runs' character properties in separate writes",
            )
            .with_property(key.to_owned()));
        }
        let of_kind = |parent, kind: &str| -> Vec<super::super::At> {
            file.kids(parent)
                .iter()
                .filter(|&child| file.item(child).type_name() == kind)
                .collect()
        };
        let paragraphs = of_kind(shape, "paragraph");
        let mut start = 0;
        for &paragraph in paragraphs.iter().take(line) {
            start += file
                .item(paragraph)
                .text()
                .unwrap_or_default()
                .chars()
                .count()
                + 1;
        }
        let paragraph = *paragraphs
            .get(line)
            .ok_or_else(|| refusal(Code::PathNotFound, "the run's paragraph is missing"))?;
        let runs: Vec<_> = of_kind(paragraph, "run")
            .into_iter()
            .map(|run| (run, file.item(run)))
            .collect();
        let ordinal = runs
            .iter()
            .position(|(run, _)| *run == node.at())
            .ok_or_else(|| refusal(Code::InternalError, "the run is not in its paragraph"))?;
        let runs: Vec<_> = runs.into_iter().map(|(_, run)| run).collect();
        for run in runs.iter().take(ordinal) {
            start += run.text().unwrap_or_default().chars().count();
        }
        let length = node.item().text().unwrap_or_default().chars().count();
        if length == 0 {
            return Err(refusal(
                Code::InvalidArgument,
                "the run holds no characters",
            ));
        }
        let entry = self
            .styles
            .entry((slide, drawable, start, start + length))
            .or_insert_with(|| (Style::new(), index));
        run_style(&mut entry.0, key, value)?;
        Ok(())
    }

    /// Stage one slide-table cell value (record 0309).
    fn stage_cell(
        &mut self,
        file: &Iwork,
        node: &Node<'_>,
        key: &str,
        value: &PropValue<'_>,
        index: usize,
    ) -> Result<()> {
        let [slide, _, row, drawable] = node.item().leaf().map(|value| value as usize);
        let slide_at = file
            .kids(super::super::At::ROOT)
            .get(slide)
            .ok_or_else(|| refusal(Code::PathNotFound, "the cell's slide is missing"))?;
        let table = file
            .kids(slide_at)
            .iter()
            .find(|&child| {
                let item = file.stand_in(child);
                item.type_name() == "table" && item.leaf()[3] as usize == drawable
            })
            .ok_or_else(|| refusal(Code::PathNotFound, "the cell's table is missing"))?;
        let table_ordinal = file
            .kids(slide_at)
            .iter()
            .filter(|&child| file.item(child).type_name() == "table")
            .position(|child| child == table)
            .ok_or_else(|| refusal(Code::InternalError, "the cell's table ordinal is missing"))?;
        let row_at = file
            .kids(table)
            .iter()
            .filter(|&child| file.item(child).type_name() == "tr")
            .nth(row)
            .ok_or_else(|| refusal(Code::PathNotFound, "the cell's row is missing"))?;
        let column = file
            .kids(row_at)
            .iter()
            .filter(|&child| file.item(child).type_name() == "tc")
            .position(|child| child == node.at())
            .ok_or_else(|| refusal(Code::InternalError, "the cell is not in its row"))?;
        if key.eq_ignore_ascii_case("comment") {
            self.comments.insert(
                (
                    slide,
                    table_ordinal,
                    u32::try_from(row).unwrap_or(u32::MAX),
                    u32::try_from(column).unwrap_or(u32::MAX),
                ),
                (super::super::write::text(value, "comment")?, index),
            );
            return Ok(());
        }
        let cell = table_cell_value(key, value)?;
        self.cells.insert(
            (
                slide,
                drawable,
                u32::try_from(row).unwrap_or(u32::MAX),
                u32::try_from(column).unwrap_or(u32::MAX),
            ),
            (cell, index),
        );
        Ok(())
    }

    /// Stage one `key = value` assignment on `node`.
    pub(super) fn stage_set(
        &mut self,
        file: &Iwork,
        node: &Node<'_>,
        key: &str,
        value: &PropValue<'_>,
        index: usize,
    ) -> Result<()> {
        let item = node.item();
        let [slide, _, line, drawable] = item.leaf().map(|value| value as usize);
        match (item.type_name(), key) {
            ("slide", "hidden") => {
                let hidden = super::super::write::boolean(value, "hidden")?;
                self.hidden.insert(slide, (hidden, index));
            },
            ("placeholder" | "notes" | "shape", "text") => {
                let text = Text::of(node)
                    .ok_or_else(|| refusal(Code::InternalError, "the text is missing"))?;
                let value = super::super::write::text(value, "text")?.unwrap_or_default();
                self.touch(file, text, Touched::Whole)?;
                self.stage_text(text, value, index);
            },
            ("paragraph", "text") => {
                let text = Text::of(node)
                    .ok_or_else(|| refusal(Code::InternalError, "the text is missing"))?;
                let value = paragraph_text(value)?;
                self.stage_paragraph(file, text, line, value, index)?;
            },
            ("tc", "comment") => self.stage_cell(file, node, key, value, index)?,
            ("placeholder" | "shape" | "picture", "x" | "y" | "width" | "height" | "rotation") => {
                if drawable == NONE {
                    return Err(refusal(
                        Code::UnsupportedCapability,
                        "this placeholder has no frame on its slide",
                    )
                    .with_property(key.to_owned()));
                }
                let current = match self.frames.get(&(slide, drawable)) {
                    Some((geometry, _)) => *geometry,
                    None => frame_of(node).ok_or_else(|| {
                        refusal(Code::UnsupportedCapability, "the drawable records no frame")
                            .with_property(key.to_owned())
                    })?,
                };
                let mut geometry = current;
                match key {
                    "x" => geometry.x = points(value, "x", false)?,
                    "y" => geometry.y = points(value, "y", false)?,
                    "width" => geometry.width = points(value, "width", true)?,
                    "height" => geometry.height = points(value, "height", true)?,
                    _ => geometry.rotation = Some(degrees(value)?),
                }
                self.frames.insert((slide, drawable), (geometry, index));
            },
            ("run", "bold" | "italic" | "underline" | "size" | "color" | "font.latin") => {
                self.stage_run_style(file, node, key, value, index)?;
            },
            ("run", _) => return Err(run_refusal(key)),
            ("tc", "text" | "value") => self.stage_cell(file, node, key, value, index)?,
            ("placeholder", "hidden") => {
                let [_, slot, _, _] = node.item().leaf().map(|value| value as usize);
                if slot != TITLE && slot != BODY {
                    return Err(key_refusal(node, key));
                }
                let hidden = super::super::write::boolean(value, "hidden")?;
                self.visibility.insert((slide, slot), (hidden, index));
            },
            (
                "slide" | "placeholder" | "notes" | "paragraph" | "shape" | "picture" | "connector"
                | "video" | "audio" | "table" | "tc" | "chart" | "group",
                _,
            ) => return Err(key_refusal(node, key)),
            _ => return Err(element_refusal("set", node)),
        }
        Ok(())
    }

    /// Stage the removal of one drawable.
    pub(super) fn stage_removal(&mut self, slide: usize, drawable: usize, index: usize) {
        self.removals.insert((slide, drawable), index);
    }

    /// Stage one new drawable.
    pub(super) fn stage_add(&mut self, add: Add, index: usize) {
        self.adds.push((add, index));
    }

    /// Publish every staged change onto `chain`: one slide-drawable commit,
    /// one commit per changed placeholder or notes text, one per changed
    /// slide visibility. A failure names the batch item it belongs to.
    pub(super) fn flush(
        &mut self,
        chain: &mut Chain<Package>,
    ) -> std::result::Result<Published, (Error, Option<usize>)> {
        let mut published = Published::default();
        // Join the texts staged as paragraphs, once.
        let groups = &self.groups;
        let texts = std::mem::take(&mut self.texts)
            .into_iter()
            .map(|(text, (staged, index))| {
                let value = match staged {
                    Staged::Text(value) => value,
                    Staged::Paragraphs => {
                        groups.get(&text).map(TextGroup::text).unwrap_or_default()
                    },
                };
                (text, (value, index))
            })
            .collect::<BTreeMap<_, _>>();
        let drawable_texts = texts
            .iter()
            .filter(|(text, _)| text.container == DRAWABLE)
            .map(|(text, (value, index))| (*text, value.as_str(), *index))
            .collect::<Vec<_>>();
        if !drawable_texts.is_empty()
            || !self.frames.is_empty()
            || !self.removals.is_empty()
            || !self.adds.is_empty()
            || !self.styles.is_empty()
            || !self.cells.is_empty()
        {
            let first = drawable_texts
                .iter()
                .map(|(_, _, index)| *index)
                .chain(self.frames.values().map(|(_, index)| *index))
                .chain(self.removals.values().copied())
                .chain(self.adds.iter().map(|(_, index)| *index))
                .chain(self.styles.values().map(|(_, index)| *index))
                .chain(self.cells.values().map(|(_, index)| *index))
                .min();
            let at = |index: usize| {
                move |error| (failure(LeafError::KeynoteDrawables(error)), Some(index))
            };
            let commit = {
                let mut edit = chain.package.edit_slide_drawables();
                for ((slide, drawable), (geometry, index)) in &self.frames {
                    edit.set_geometry(
                        SlideSelector::index(*slide),
                        DrawableSelector::index(*drawable),
                        *geometry,
                    )
                    .map_err(at(*index))?;
                }
                for (text, value, index) in &drawable_texts {
                    edit.set_text(
                        SlideSelector::index(text.slide),
                        DrawableSelector::index(text.drawable),
                        value,
                    )
                    .map_err(at(*index))?;
                }
                for ((slide, drawable, start, end), (style, index)) in &self.styles {
                    if *drawable == NONE {
                        edit.set_notes_text_style(
                            SlideSelector::index(*slide),
                            DrawableSpan::new(*start, *end),
                            style.clone(),
                        )
                    } else {
                        edit.set_text_style(
                            SlideSelector::index(*slide),
                            DrawableSelector::index(*drawable),
                            DrawableSpan::new(*start, *end),
                            style.clone(),
                        )
                    }
                    .map_err(at(*index))?;
                }
                for ((slide, drawable, row, column), (value, index)) in &self.cells {
                    edit.set_table_cell(
                        SlideSelector::index(*slide),
                        DrawableSelector::index(*drawable),
                        *row,
                        *column,
                        value.clone(),
                    )
                    .map_err(at(*index))?;
                }
                for ((slide, drawable), index) in &self.removals {
                    edit.remove(
                        SlideSelector::index(*slide),
                        DrawableSelector::index(*drawable),
                    )
                    .map_err(at(*index))?;
                }
                for (add, index) in &self.adds {
                    match add {
                        Add::Text {
                            slide,
                            preset: None,
                            geometry,
                            text,
                        } => edit.add_text_box(SlideSelector::index(*slide), *geometry, text),
                        Add::Text {
                            slide,
                            preset: Some(preset),
                            geometry,
                            text,
                        } => edit.add_shape(SlideSelector::index(*slide), *preset, *geometry, text),
                        Add::Image {
                            slide,
                            bytes,
                            name,
                            geometry,
                        } => edit.add_image(SlideSelector::index(*slide), bytes, name, *geometry),
                        Add::Table { slide, table } => {
                            edit.add_table(SlideSelector::index(*slide), table.clone())
                        },
                    }
                    .map_err(at(*index))?;
                }
                edit.commit()
                    .map_err(|error| (failure(LeafError::KeynoteDrawables(error)), first))?
            };
            published.created = commit.created().to_vec();
            let (package, patch) = commit.into_parts();
            push(chain, package, Patch::Drawables(patch));
        }
        for ((slide, table, row, column), (comment, index)) in std::mem::take(&mut self.comments) {
            let error = |error| (failure(LeafError::KeynoteTableComments(error)), Some(index));
            let cell = litchi_keynote::slide::table::cells::At::new(row, column);
            let edit = chain
                .package
                .edit_slide_table_cell_comment(
                    SlideSelector::index(slide),
                    TableSelector::index(table),
                    cell,
                )
                .map_err(error)?;
            let commit = match comment {
                Some(comment) => edit.set(&comment).map_err(error)?.commit(),
                None => edit.clear().map_err(error)?.commit(),
            }
            .map_err(error)?;
            let patch = commit.patch().clone();
            push(chain, commit.into_package(), Patch::TableComments(patch));
        }
        for (text, (value, index)) in texts {
            let at = |error| (error, Some(index));
            match text.container {
                TITLE | BODY => {
                    let role = if text.container == TITLE {
                        Role::Title
                    } else {
                        Role::Body
                    };
                    let error = |error| failure(LeafError::KeynoteText(error));
                    let commit = {
                        let mut edit = chain
                            .package
                            .edit_slide_text(SlideSelector::index(text.slide), role)
                            .map_err(error)
                            .map_err(at)?;
                        if let Some((span, replacement)) =
                            minimal_span(edit.text().map_err(error).map_err(at)?, &value)
                                .map_err(at)?
                        {
                            edit.replace(span, replacement).map_err(error).map_err(at)?;
                        }
                        edit.commit().map_err(error).map_err(at)?
                    };
                    let patch = commit.patch().clone();
                    push(chain, commit.into_package(), Patch::Text(patch));
                },
                NOTES => {
                    let error = |error| failure(LeafError::KeynoteNotes(error));
                    let commit = {
                        let mut edit = chain
                            .package
                            .edit_slide_notes(SlideSelector::index(text.slide))
                            .map_err(error)
                            .map_err(at)?;
                        if let Some((span, replacement)) =
                            minimal_span(edit.text(), &value).map_err(at)?
                        {
                            edit.replace(span, replacement).map_err(error).map_err(at)?;
                        }
                        edit.commit().map_err(error).map_err(at)?
                    };
                    let patch = commit.patch().clone();
                    push(chain, commit.into_package(), Patch::Notes(patch));
                },
                _ => {},
            }
        }
        for ((slide, slot), (hidden, index)) in std::mem::take(&mut self.visibility) {
            let error = |error| (failure(LeafError::KeynotePlaceholder(error)), Some(index));
            let kind = if slot == TITLE {
                litchi_keynote::slide::placeholder::Kind::Title
            } else {
                litchi_keynote::slide::placeholder::Kind::Body
            };
            let edit = chain
                .package
                .edit_slide_placeholder_visibility(SlideSelector::index(slide), kind)
                .map_err(error)?;
            let edit = if hidden { edit.hide() } else { edit.show() };
            let commit = edit.commit().map_err(error)?;
            let patch = commit.patch().clone();
            push(chain, commit.into_package(), Patch::Placeholder(patch));
        }
        for (slide, (hidden, index)) in std::mem::take(&mut self.hidden) {
            let error = |error| (failure(LeafError::KeynoteEdit(error)), Some(index));
            let commit = {
                let mut edit = chain.package.edit();
                edit.set_slide_skipped(SlideSelector::index(slide), hidden)
                    .map_err(error)?;
                edit.commit().map_err(error)?
            };
            let patch = commit.patch().clone();
            push(chain, commit.into_package(), Patch::Skip(patch));
        }
        self.frames.clear();
        self.removals.clear();
        self.styles.clear();
        self.cells.clear();
        self.comments.clear();
        self.adds.clear();
        self.groups.clear();
        Ok(published)
    }
}

/// The smallest UTF-16 span of `before` whose replacement gives `after`,
/// with that replacement; `None` when the texts are equal. The common prefix
/// and suffix keep their paragraph and character attributes.
fn minimal_span<'a>(before: &str, after: &'a str) -> Result<Option<(TextSpan, &'a str)>> {
    if before == after {
        return Ok(None);
    }
    let prefix = before
        .char_indices()
        .zip(after.chars())
        .find(|((_, left), right)| left != right)
        .map_or_else(|| before.len().min(after.len()), |((index, _), _)| index);
    let (before_rest, after_rest) = (&before[prefix..], &after[prefix..]);
    let suffix = before_rest
        .chars()
        .rev()
        .zip(after_rest.chars().rev())
        .take_while(|(left, right)| left == right)
        .map(|(left, _)| left.len_utf8())
        .sum::<usize>();
    let removed = &before_rest[..before_rest.len() - suffix];
    let inserted = &after_rest[..after_rest.len() - suffix];
    let start = before[..prefix].encode_utf16().count();
    let end = start + removed.encode_utf16().count();
    let position = |units| {
        TextPosition::from_utf16_index(units)
            .map_err(|_| refusal(Code::LimitExceeded, "the text is too long to address"))
    };
    let span = TextSpan::new(position(start)?, position(end)?)
        .map_err(|_| refusal(Code::InternalError, "the text span is reversed"))?;
    Ok(Some((span, inserted)))
}

/// Publish `pending` from `package` and adopt the result.
fn publish(
    file: &Iwork,
    package: &Package,
    mut pending: Pending,
    operations: usize,
    limits: &SemanticLimits,
) -> Result<(Commit, Published)> {
    let mut chain = Chain::new(package.clone());
    let published = pending.flush(&mut chain).map_err(|(error, _)| error)?;
    let commit = finish(
        file,
        chain.changed,
        Owner::Keynote(chain.package),
        chain.patches,
        operations,
        limits,
    )?;
    Ok((commit, published))
}

/// The leaf value of a slide-table cell's `text` or `value`.
fn table_cell_value(key: &str, value: &PropValue<'_>) -> Result<TableCellValue> {
    use super::super::write::CellScalar;
    Ok(match super::super::write::cell_scalar(key, value)? {
        CellScalar::Text(text) => TableCellValue::Text(text),
        CellScalar::Number(number) => TableCellValue::Number(number),
        CellScalar::Boolean(flag) => TableCellValue::Boolean(flag),
        CellScalar::Empty => TableCellValue::Empty,
    })
}

/// A one-value set of a slide-table cell named by the direct path
/// `/slide[N]/table[M]/tr[R]/tc[C]` (exactly those step names, each with one
/// positional index), committed without reading the table (record 0312):
/// the slide's part locates the table's drawable through its stand-in, and
/// the leaf writer checks the cell against the table. `None` sends any
/// other target, and any failure, to the general route, which reports it
/// (the leaf refuses a cell outside the table, `InvalidSpan`, so the
/// general route reports the missing path).
fn direct_cell(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Option<Commit> {
    let mut entries = props.iter();
    let (Some((key, value)), None) = (entries.next(), entries.next()) else {
        return None;
    };
    let key = key.as_str();
    if key != "text" && key != "value" {
        return None;
    }
    let path = target.path()?.path();
    let [slide, table, row, column] = path.steps() else {
        return None;
    };
    if !path.is_absolute()
        || slide.name() != "slide"
        || table.name() != "table"
        || row.name() != "tr"
        || column.name() != "tc"
    {
        return None;
    }
    let [slide, table, row, column] =
        [slide, table, row, column].map(|step| step.index().map(Position::get));
    let (slide, table, row, column) = (slide?, table?, row?, column?);
    bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets()).ok()?;
    bound(LimitKind::Operations, 1, limits.max_operations()).ok()?;
    file.checked_tree().ok()?;
    let slide_at = file.kids(super::super::At::ROOT).get(slide)?;
    let ordinal = crate::path::navigate::ElementKind::ordinal(super::Kind::Table);
    let table_at = file
        .kids(slide_at)
        .iter()
        .filter(|&child| file.kind_at(child) == ordinal)
        .nth(table)?;
    let mut pending = Pending::default();
    let [slide, _, _, drawable] = file.stand_in(table_at).leaf().map(|value| value as usize);
    let cell = table_cell_value(key, value).ok()?;
    pending.cells.insert(
        (
            slide,
            drawable,
            u32::try_from(row).ok()?,
            u32::try_from(column).ok()?,
        ),
        (cell, 0),
    );
    publish(file, package, pending, 1, limits)
        .ok()
        .map(|(commit, _)| commit)
}

/// The comment-only direct-path setter. Focused comment errors remain
/// terminal instead of falling back to the generic tree route.
fn direct_comment(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Commit>> {
    let mut entries = props.iter();
    let (Some((key, value)), None) = (entries.next(), entries.next()) else {
        return Ok(None);
    };
    if key.as_str() != "comment" {
        return Ok(None);
    }
    let Some(path) = target.path().map(|target| target.path()) else {
        return Ok(None);
    };
    let [slide, table, row, column] = path.steps() else {
        return Ok(None);
    };
    if !path.is_absolute()
        || slide.name() != "slide"
        || table.name() != "table"
        || row.name() != "tr"
        || column.name() != "tc"
    {
        return Ok(None);
    }
    let [slide, table, row, column] =
        [slide, table, row, column].map(|step| step.index().map(Position::get));
    let (Some(slide), Some(table), Some(row), Some(column)) = (slide, table, row, column) else {
        return Ok(None);
    };
    bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    file.checked_tree()?;
    let Some(slide_at) = file.kids(super::super::At::ROOT).get(slide) else {
        return Ok(None);
    };
    let ordinal = crate::path::navigate::ElementKind::ordinal(super::Kind::Table);
    let Some(_table_at) = file
        .kids(slide_at)
        .iter()
        .filter(|&child| file.kind_at(child) == ordinal)
        .nth(table)
    else {
        return Ok(None);
    };
    let mut pending = Pending::default();
    pending.comments.insert(
        (
            slide,
            table,
            u32::try_from(row).map_err(|_| refusal(Code::OutOfRange, "row is too large"))?,
            u32::try_from(column).map_err(|_| refusal(Code::OutOfRange, "column is too large"))?,
        ),
        (super::super::write::text(value, "comment")?, 0),
    );
    publish(file, package, pending, 1, limits).map(|(commit, _)| Some(commit))
}

/// Set properties on every target.
pub(crate) fn set(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    if let Some(commit) = direct_comment(file, package, target, props, limits)? {
        return Ok(commit);
    }
    if let Some(commit) = direct_cell(file, package, target, props, limits) {
        return Ok(commit);
    }
    let nodes = resolve::<Keynote>(file, target, limits.max_expanded_targets())?;
    let mut pending = Pending::default();
    let mut operations = 0_usize;
    for node in &nodes {
        for (key, value) in props.iter() {
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            pending.stage_set(file, node, key.as_str(), value, 0)?;
        }
    }
    publish(file, package, pending, operations, limits).map(|(commit, _)| commit)
}

// ---------------------------------------------------------------------------
// Add
// ---------------------------------------------------------------------------

/// The refusal of an element kind `add` does not create.
fn add_gap(kind: &str) -> Error {
    refusal(
        Code::UnsupportedCapability,
        format!("adding a Keynote `{kind}` is not supported yet"),
    )
    .with_suggestion(match kind {
        "tr" | "tc" | "row" | "cell" => {
            "add a table with `rows`, `cols` and `data` (record 0329); row and column edits of slide tables are a follow-up"
        },
        "chart" => "follow-up: Keynote chart creation in the unified layer (record 0298)",
        "video" | "audio" => "follow-up: wire slide::movie and slide::audio creation (record 0298)",
        "run" | "r" => {
            "add a paragraph, or set a paragraph's text (IWORK_API_REQUESTS.md request 10)"
        },
        _ => "follow-up: record 0298 (Keynote unified structure)",
    })
}

/// The zero-based position a new element takes among `count` siblings.
pub(super) fn insertion(
    count: usize,
    placement: Option<&Placement<'_>>,
    anchor: impl Fn(&PathArg<'_>) -> Result<usize>,
) -> Result<usize> {
    let position = match placement {
        None | Some(Placement::End | Placement::Into(_, None)) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position) | Placement::Into(_, Some(position))) => position.get(),
        Some(Placement::Before(path)) => anchor(path)?,
        Some(Placement::After(path)) => anchor(path)? + 1,
    };
    if position > count {
        return Err(refusal(
            Code::OutOfRange,
            "the insertion position is outside the element's siblings",
        ));
    }
    Ok(position)
}

/// The slide a drawable add targets and its new-drawable request.
pub(super) struct DrawableRequest {
    pub(super) add: Add,
    /// The element kind the new drawable reads as.
    pub(super) element: &'static str,
}

/// The default frame of a new drawable, in points.
const DEFAULT_FRAME: Geometry = Geometry::new(100.0, 100.0, 400.0, 100.0);
const DEFAULT_PICTURE: Geometry = Geometry::new(100.0, 100.0, 200.0, 200.0);

/// Read a new drawable's request from `node`'s properties.
pub(super) fn drawable_request(
    slide: usize,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<DrawableRequest> {
    let kind = node.kind().to_ascii_lowercase();
    let picture = matches!(kind.as_str(), "picture" | "image" | "img" | "pic");
    let text_box = kind == "textbox";
    let mut geometry = if picture {
        DEFAULT_PICTURE
    } else {
        DEFAULT_FRAME
    };
    let mut text = String::new();
    let mut preset = None;
    let mut src = None;
    let mut name = None;
    for (key, value) in node.props().iter() {
        match (key.as_str(), picture) {
            ("x", _) => geometry.x = points(value, "x", false)?,
            ("y", _) => geometry.y = points(value, "y", false)?,
            ("width", _) => geometry.width = points(value, "width", true)?,
            ("height", _) => geometry.height = points(value, "height", true)?,
            ("rotation", _) => geometry.rotation = Some(degrees(value)?),
            ("text", false) => {
                text = super::super::write::text(value, "text")?.unwrap_or_default();
            },
            ("preset", false) if !text_box => {
                let token = super::super::write::text(value, "preset")?.unwrap_or_default();
                preset = Some(match token.to_ascii_lowercase().as_str() {
                    "rect" | "rectangle" => Preset::Rectangle,
                    "roundrect" | "roundedrectangle" => Preset::RoundedRectangle { radius: 0.0 },
                    "ellipse" | "oval" => Preset::Ellipse,
                    _ => {
                        return Err(refusal(
                            Code::InvalidValue,
                            "a Keynote shape preset is rect, roundRect or ellipse",
                        )
                        .with_property("preset".to_owned()));
                    },
                });
            },
            ("src" | "path", true) => {
                src = super::super::write::text(value, "src")?;
            },
            ("name", true) => name = super::super::write::text(value, "name")?,
            (other, true) => return Err(unwritable("picture", other)),
            (other, false) => return Err(unwritable("shape", other)),
        }
    }
    if let Some(Preset::RoundedRectangle { radius }) = &mut preset {
        // A tenth of the shorter side, as Keynote's default corner.
        *radius = geometry.width.min(geometry.height) / 10.0;
    }
    if picture {
        let src = src.ok_or_else(|| {
            refusal(Code::MissingProperty, "pictures require a src").with_property("src".to_owned())
        })?;
        let bytes = super::source::read(&src, limits.max_staged_bytes())?;
        let name = name
            .or_else(|| super::source::file_name(&src))
            .unwrap_or_else(|| super::source::default_name(&bytes).to_owned());
        return Ok(DrawableRequest {
            add: Add::Image {
                slide,
                bytes: Arc::from(bytes),
                name,
                geometry,
            },
            element: "picture",
        });
    }
    let preset = if text_box {
        None
    } else {
        Some(preset.unwrap_or(Preset::Rectangle))
    };
    Ok(DrawableRequest {
        add: Add::Text {
            slide,
            preset,
            geometry,
            text,
        },
        element: "shape",
    })
}

/// The default frame of a new table, in points.
const DEFAULT_TABLE_FRAME: Geometry = Geometry::new(100.0, 100.0, 600.0, 200.0);

/// Read a new table's request from `node`'s properties: the table keys
/// (`rows`, `cols`, `name`, `data`) and the frame (`x`, `y`, `width`,
/// `height`; record 0329).
pub(super) fn table_request(slide: usize, node: &NewNode<'_>) -> Result<DrawableRequest> {
    let mut request = super::super::write::TableRequest::default();
    let mut geometry = DEFAULT_TABLE_FRAME;
    for (key, value) in node.props().iter() {
        match key.as_str() {
            "x" => geometry.x = points(value, "x", false)?,
            "y" => geometry.y = points(value, "y", false)?,
            "width" => geometry.width = points(value, "width", true)?,
            "height" => geometry.height = points(value, "height", true)?,
            other => {
                if !request.key(other, value)? {
                    return Err(unwritable("table", other));
                }
            },
        }
    }
    let (rows, columns, cells) = request.shape()?;
    let mut table = NewTable::new(rows, columns, geometry);
    if let Some(name) = request.name.take() {
        table = table.with_name(name);
    }
    for (row, column, value) in cells {
        table = table.with_cell(row, column, TableCellValue::Text(value));
    }
    Ok(DrawableRequest {
        add: Add::Table { slide, table },
        element: "table",
    })
}

/// What a new slide copies or starts from.
#[derive(Debug, Clone)]
pub(super) enum SlideSource {
    /// A theme layout by name or position (`None`: the theme's default).
    Layout(Option<LayoutChoice>),
    /// A copy of the slide at this position of the base snapshot.
    Duplicate(usize),
}

/// A layout chosen by name or by its one-based number.
#[derive(Debug, Clone)]
pub(super) enum LayoutChoice {
    Name(String),
    Position(usize),
}

/// Read a new slide's `layout` or `from` property. `from` is returned as
/// path text for the caller to resolve.
pub(super) fn slide_request(node: &NewNode<'_>) -> Result<(Option<LayoutChoice>, Option<String>)> {
    let mut layout = None;
    let mut from = None;
    for (key, value) in node.props().iter() {
        match key.as_str() {
            "layout" => {
                let name = super::super::write::text(value, "layout")?.unwrap_or_default();
                layout = Some(match name.parse::<usize>() {
                    Ok(number) if number > 0 => LayoutChoice::Position(number - 1),
                    _ => LayoutChoice::Name(name),
                });
            },
            "from" => from = super::super::write::text(value, "from")?,
            other => return Err(unwritable("slide", other)),
        }
    }
    if layout.is_some() && from.is_some() {
        return Err(refusal(
            Code::InvalidArgument,
            "a new slide takes either a layout or a slide to copy (`from`), not both",
        ));
    }
    Ok((layout, from))
}

/// Resolve `from` to a slide position of `file`.
pub(super) fn slide_position(file: &Iwork, from: &str) -> Result<usize> {
    let path = from
        .into_path()
        .map_err(|error| error.with_property("from"))?;
    slide_at(file, &path).map_err(|error| error.with_property("from"))
}

/// The position of the slide at `path` in `file`.
pub(super) fn slide_at(file: &Iwork, path: &PathArg<'_>) -> Result<usize> {
    let node = node_at::<Keynote>(file, path)?;
    if node.item().type_name() != "slide" {
        return Err(refusal(
            Code::InvalidArgument,
            "the path does not name a slide",
        ));
    }
    Ok(node.item().leaf()[0] as usize)
}

/// Commit one slide-creation edit of `slides` (each a source and a
/// destination among the evolving slide list).
pub(super) fn create_slides(
    chain: &mut Chain<Package>,
    slides: &[(SlideSource, usize, usize)],
) -> std::result::Result<(), (Error, Option<usize>)> {
    let error =
        |index: usize| move |error| (failure(LeafError::KeynoteCreation(error)), Some(index));
    let default = chain
        .package
        .default_slide_layout()
        .map_err(error(slides.first().map_or(0, |(_, _, index)| *index)))?;
    let commit = {
        let mut edit = chain.package.edit_slide_creation();
        for (source, destination, index) in slides {
            let destination = Position::new(*destination);
            match source {
                SlideSource::Duplicate(slide) => {
                    edit.duplicate_slide(SlideSelector::index(*slide), destination)
                        .map_err(error(*index))?;
                },
                SlideSource::Layout(choice) => {
                    let selector = match choice {
                        Some(LayoutChoice::Name(name)) => LayoutSelector::Name(name),
                        Some(LayoutChoice::Position(position)) => {
                            LayoutSelector::Position(Position::new(*position))
                        },
                        None => default
                            .as_ref()
                            .map_or(LayoutSelector::Position(Position::new(0)), |layout| {
                                layout.selector()
                            }),
                    };
                    edit.add_slide(selector, destination)
                        .map_err(error(*index))?;
                },
            }
        }
        edit.commit()
            .map_err(error(slides.first().map_or(0, |(_, _, index)| *index)))?
    };
    let (package, patch) = commit.into_parts();
    push(chain, package, Patch::Creation(patch));
    Ok(())
}

/// Add one element.
pub(crate) fn add(
    file: &Iwork,
    package: &Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    file.checked_tree()?;
    let parent_node = node_at::<Keynote>(file, parent)?;
    let kind = node.kind().to_ascii_lowercase();
    match (parent_node.item().type_name(), kind.as_str()) {
        ("presentation", "slide") => {
            let (layout, from) = slide_request(node)?;
            let source = match from {
                Some(from) => SlideSource::Duplicate(slide_position(file, &from)?),
                None => SlideSource::Layout(layout),
            };
            let count = file.kids(super::super::At::ROOT).len();
            let destination = insertion(count, node.placement(), |anchor| slide_at(file, anchor))?;
            let mut chain = Chain::new(package.clone());
            create_slides(&mut chain, &[(source, destination, 0)]).map_err(|(error, _)| error)?;
            let commit = finish(
                file,
                chain.changed,
                Owner::Keynote(chain.package),
                chain.patches,
                1,
                limits,
            )?;
            let path = Path::parse(&format!("/slide[{}]", destination + 1))?.into_owned();
            Ok((commit, path))
        },
        ("slide", "shape" | "textbox" | "picture" | "image" | "img" | "pic" | "table" | "tbl") => {
            if !matches!(node.placement(), None | Some(Placement::End)) {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "a new Keynote drawable goes in front of the slide's drawables",
                )
                .with_suggestion("omit the placement, then `move` it in the z-order"));
            }
            let slide = parent_node.item().leaf()[0] as usize;
            let request = if matches!(kind.as_str(), "table" | "tbl") {
                table_request(slide, node)?
            } else {
                drawable_request(slide, node, limits)?
            };
            let element = request.element;
            // A new drawable joins the end of its slide's drawables: its
            // ordinal among the slide's elements of its kind is their count.
            // Tables count by kind, without building their contents.
            let table = crate::path::navigate::ElementKind::ordinal(super::Kind::Table);
            let ordinal = parent_node
                .child_nodes()
                .filter(|child| {
                    if element == "table" {
                        child.kind_ordinal() == table
                    } else {
                        child.kind_ordinal() != table && child.item().type_name() == element
                    }
                })
                .count();
            let mut pending = Pending::default();
            pending.stage_add(request.add, 0);
            let (commit, published) = publish(file, package, pending, 1, limits)?;
            if published.created.is_empty() {
                return Err(refusal(Code::InternalError, "the new drawable is missing"));
            }
            let path = Path::parse(&format!("/slide[{}]/{element}[{}]", slide + 1, ordinal + 1))?
                .into_owned();
            Ok((commit, path))
        },
        ("placeholder" | "notes" | "shape", "p" | "paragraph") => {
            let text = Text::of(&parent_node)
                .ok_or_else(|| refusal(Code::InternalError, "the text is missing"))?;
            let mut value = String::new();
            for (key, prop) in node.props().iter() {
                if key.as_str() == "text" {
                    value = paragraph_text(prop)?;
                } else {
                    return Err(unwritable("paragraph", key.as_str()));
                }
            }
            let mut pending = Pending::default();
            let current = pending.current(file, text)?;
            let mut lines = current.split('\n').collect::<Vec<_>>();
            let line = insertion(lines.len(), node.placement(), |anchor| {
                paragraph_line(file, text, anchor)
            })?;
            lines.insert(line, &value);
            pending.stage_text(text, lines.join("\n"), 0);
            let (commit, _) = publish(file, package, pending, 1, limits)?;
            let container = text
                .item(file)
                .map(|item| crate::node::NodePath::new(item.path()).to_string())
                .ok_or_else(|| refusal(Code::InternalError, "the text is missing"))?;
            let path = Path::parse(&format!("{container}/paragraph[{}]", line + 1))?.into_owned();
            Ok((commit, path))
        },
        (_, "slide") => Err(refusal(
            Code::InvalidArgument,
            "a Keynote slide is added to the presentation (`/`)",
        )),
        (_, "p" | "paragraph") => Err(refusal(
            Code::InvalidArgument,
            "a Keynote paragraph is added to a placeholder, the notes or a shape",
        )),
        ("slide", other) => Err(add_gap(other)),
        _ => Err(element_refusal("add", &parent_node)),
    }
}

/// The line of the paragraph at `anchor`, which must belong to `text`.
fn paragraph_line(file: &Iwork, text: Text, anchor: &PathArg<'_>) -> Result<usize> {
    let anchor = node_at::<Keynote>(file, anchor)?;
    if anchor.item().type_name() == "paragraph" && Text::of(&anchor) == Some(text) {
        Ok(anchor.item().leaf()[2] as usize)
    } else {
        Err(refusal(
            Code::InvalidArgument,
            "a paragraph is placed relative to another paragraph of the same text",
        ))
    }
}

// ---------------------------------------------------------------------------
// Remove, move, swap
// ---------------------------------------------------------------------------

/// The refusal of removing an element the leaf keeps.
fn removal_gap(node: &Node<'_>) -> Error {
    let kind = node.item().type_name();
    refusal(
        Code::UnsupportedCapability,
        format!("removing a Keynote `{kind}` is not supported yet"),
    )
    .with_suggestion(match kind {
        "table" | "chart" | "group" => {
            "the leaf's drawable removal refuses tables, charts and groups (record 0298 follow-up)"
        },
        "placeholder" => "set its text to empty instead",
        "notes" => "set the notes text to empty instead",
        _ => "follow-up: record 0298 (Keynote unified structure)",
    })
}

/// Remove every target: paragraphs and drawables in one staged publication,
/// then slides from the last to the first.
pub(crate) fn remove(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Keynote>(file, target, limits.max_expanded_targets())?;
    let mut slides = Vec::new();
    let mut lines: BTreeMap<Text, Vec<usize>> = BTreeMap::new();
    let mut pending = Pending::default();
    let mut media: Vec<(usize, usize)> = Vec::new();
    for node in &nodes {
        let [slide, _, line, drawable] = node.item().leaf().map(|value| value as usize);
        match node.item().type_name() {
            "slide" => slides.push(slide),
            "video" | "audio" => media.push((slide, drawable)),
            "paragraph" => {
                let text = Text::of(node)
                    .ok_or_else(|| refusal(Code::InternalError, "the text is missing"))?;
                lines.entry(text).or_default().push(line);
            },
            "shape" | "picture" | "connector" => pending.stage_removal(slide, drawable, 0),
            "placeholder" => {
                // A layout placeholder is hidden on its slide, not deleted.
                let [_, slot, _, _] = node.item().leaf().map(|value| value as usize);
                if slot != TITLE && slot != BODY {
                    return Err(removal_gap(node));
                }
                pending.visibility.insert((slide, slot), (true, 0));
            },
            "notes" | "table" | "chart" | "group" => {
                return Err(removal_gap(node));
            },
            _ => return Err(element_refusal("remove", node)),
        }
    }
    let operations = nodes.len();
    bound(LimitKind::Operations, operations, limits.max_operations())?;
    for (text, mut removed) in lines {
        removed.sort_unstable();
        removed.dedup();
        let current = pending.current(file, text)?;
        let kept = current
            .split('\n')
            .enumerate()
            .filter(|(line, _)| removed.binary_search(line).is_err())
            .map(|(_, paragraph)| paragraph)
            .collect::<Vec<_>>();
        if kept.is_empty() {
            return Err(
                refusal(Code::InvalidArgument, "a text keeps at least one paragraph")
                    .with_suggestion("set the text to empty instead"),
            );
        }
        pending.stage_text(text, kept.join("\n"), 0);
    }
    let mut chain = Chain::new(package.clone());
    if !pending.is_empty() {
        pending.flush(&mut chain).map_err(|(error, _)| error)?;
    }
    // Movies and audio, each slide's from its last, so the earlier source
    // positions hold (record 0309).
    media.sort_unstable();
    media.dedup();
    media.retain(|(slide, _)| !slides.contains(slide));
    for (slide, drawable) in media.iter().rev() {
        let list = chain
            .package
            .slide_drawable_list(SlideSelector::index(*slide))
            .map_err(|error| failure(LeafError::KeynoteDrawables(error)))?;
        let ordinal = list
            .iter()
            .take(*drawable)
            .filter(|item| {
                matches!(
                    item.kind(),
                    litchi_keynote::slide::drawable::Kind::Movie
                        | litchi_keynote::slide::drawable::Kind::Audio
                )
            })
            .count();
        let commit = chain
            .package
            .edit_slide_media_lifecycle()
            .remove(
                SlideSelector::index(*slide),
                litchi_keynote::MovieSelector::index(ordinal),
            )
            .and_then(|edit| edit.commit())
            .map_err(|error| failure(LeafError::KeynoteMedia(error)))?;
        let patch = commit.patch().clone();
        push(&mut chain, commit.into_package(), Patch::Media(patch));
    }
    slides.sort_unstable();
    slides.dedup();
    let error = |error| failure(LeafError::KeynoteDeletion(error));
    for slide in slides.iter().rev() {
        let commit = {
            let mut edit = chain.package.edit_slide_deletion();
            edit.remove_slide(SlideSelector::index(*slide))
                .map_err(error)?;
            edit.commit().map_err(error)?
        };
        let patch = commit.patch().clone();
        chain.package = commit.into_package();
        chain
            .patches
            .push(LeafPatch::Keynote(Patch::Deletion(patch)));
        chain.changed = true;
    }
    finish(
        file,
        chain.changed,
        Owner::Keynote(chain.package),
        chain.patches,
        operations,
        limits,
    )
}

/// The slide-order commit moving slide `from` to `to`.
fn move_slide(chain: &mut Chain<Package>, from: usize, to: usize) -> Result<()> {
    let error = |error| failure(LeafError::KeynoteOrder(error));
    let commit = {
        let mut edit = chain.package.edit_slide_order();
        edit.move_slide(SlideSelector::index(from), Position::new(to))
            .map_err(error)?;
        edit.commit().map_err(error)?
    };
    let patch = commit.patch().clone();
    push(chain, commit.into_package(), Patch::Order(patch));
    Ok(())
}

/// Whether `kind` is a drawable that moves in its slide's z-order.
fn in_z_order(kind: &str) -> bool {
    matches!(
        kind,
        "placeholder"
            | "shape"
            | "picture"
            | "connector"
            | "video"
            | "audio"
            | "table"
            | "chart"
            | "group"
    )
}

/// The z-order position of a drawable item.
fn z_of(node: &Node<'_>) -> Result<usize> {
    match node.item().value("zorder") {
        Some(crate::node::Value::Int(position)) => usize::try_from(*position)
            .map_err(|_| refusal(Code::InternalError, "the z-order position is negative")),
        _ => Err(refusal(
            Code::UnsupportedCapability,
            "this drawable is not in its slide's z-order",
        )),
    }
}

/// The number of drawables in `slide`'s z-order.
fn z_count(package: &Package, slide: usize) -> Result<usize> {
    let list = package
        .slide_drawable_list(SlideSelector::index(slide))
        .map_err(|error| failure(LeafError::KeynoteDrawables(error)))?;
    Ok(list
        .iter()
        .filter_map(|drawable| drawable.z_position())
        .map(|position| position.get() + 1)
        .max()
        .unwrap_or(0))
}

/// Move one slide, paragraph (within its text) or drawable (in its slide's
/// z-order).
pub(crate) fn move_to(
    file: &Iwork,
    package: &Package,
    path: &PathArg<'_>,
    placement: &Placement<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let node = node_at::<Keynote>(file, path)?;
    let kind = node.item().type_name();
    if matches!(placement, Placement::Into(..)) {
        return Err(refusal(
            Code::UnsupportedCapability,
            "a Keynote element moves among its siblings",
        ));
    }
    let mut chain = Chain::new(package.clone());
    let [slide, _, line, drawable] = node.item().leaf().map(|value| value as usize);
    if kind == "slide" {
        let slides = file.kids(super::super::At::ROOT).len();
        let to = destination(slide, slides, placement, |anchor| slide_at(file, anchor))?;
        move_slide(&mut chain, slide, to)?;
    } else if kind == "paragraph" {
        let text =
            Text::of(&node).ok_or_else(|| refusal(Code::InternalError, "the text is missing"))?;
        let mut pending = Pending::default();
        let current = pending.current(file, text)?;
        let mut lines = current.split('\n').collect::<Vec<_>>();
        let to = destination(line, lines.len(), placement, |anchor| {
            paragraph_line(file, text, anchor)
        })?;
        let moved = lines.remove(line);
        lines.insert(to, moved);
        pending.stage_text(text, lines.join("\n"), 0);
        pending.flush(&mut chain).map_err(|(error, _)| error)?;
    } else if in_z_order(kind) {
        if drawable == NONE {
            return Err(refusal(
                Code::UnsupportedCapability,
                "this placeholder has no frame on its slide",
            ));
        }
        let from = z_of(&node)?;
        let count = z_count(package, slide)?;
        let to = destination(from, count, placement, |anchor| {
            let anchor = node_at::<Keynote>(file, anchor)?;
            if in_z_order(anchor.item().type_name()) && anchor.item().leaf()[0] as usize == slide {
                z_of(&anchor)
            } else {
                Err(refusal(
                    Code::InvalidArgument,
                    "a drawable moves before or after another drawable of its slide",
                ))
            }
        })?;
        let error = |error| failure(LeafError::KeynoteDrawables(error));
        let commit = {
            let mut edit = chain.package.edit_slide_drawables();
            edit.move_to(
                SlideSelector::index(slide),
                DrawableSelector::index(drawable),
                Position::new(to),
            )
            .map_err(error)?;
            edit.commit().map_err(error)?
        };
        let (package, patch) = commit.into_parts();
        push(&mut chain, package, Patch::Drawables(patch));
    } else {
        return Err(element_refusal("move", &node));
    }
    finish(
        file,
        chain.changed,
        Owner::Keynote(chain.package),
        chain.patches,
        1,
        limits,
    )
}

/// Exchange two slides, two paragraphs of one text, or two drawables of one
/// slide in its z-order.
pub(crate) fn swap(
    file: &Iwork,
    package: &Package,
    left: &PathArg<'_>,
    right: &PathArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let left = node_at::<Keynote>(file, left)?;
    let right = node_at::<Keynote>(file, right)?;
    let (left_kind, right_kind) = (left.item().type_name(), right.item().type_name());
    let left_leaf = left.item().leaf().map(|value| value as usize);
    let right_leaf = right.item().leaf().map(|value| value as usize);
    let mut chain = Chain::new(package.clone());
    if left_kind == "slide" && right_kind == "slide" {
        let (first, second) = (
            left_leaf[0].min(right_leaf[0]),
            left_leaf[0].max(right_leaf[0]),
        );
        if first != second {
            // The first slide takes the second's place, which shifts the
            // second one back by one; that one then takes the first's place.
            move_slide(&mut chain, first, second)?;
            move_slide(&mut chain, second - 1, first)?;
        }
    } else if left_kind == "paragraph" && right_kind == "paragraph" {
        let text = Text::of(&left);
        if text.is_none() || text != Text::of(&right) {
            return Err(refusal(
                Code::InvalidArgument,
                "two paragraphs swap within one text",
            ));
        }
        let text = text.ok_or_else(|| refusal(Code::InternalError, "the text is missing"))?;
        let mut pending = Pending::default();
        let current = pending.current(file, text)?;
        let mut lines = current.split('\n').collect::<Vec<_>>();
        if left_leaf[2] >= lines.len() || right_leaf[2] >= lines.len() {
            return Err(refusal(Code::PathNotFound, "the paragraph is missing"));
        }
        lines.swap(left_leaf[2], right_leaf[2]);
        pending.stage_text(text, lines.join("\n"), 0);
        pending.flush(&mut chain).map_err(|(error, _)| error)?;
    } else if in_z_order(left_kind) && in_z_order(right_kind) {
        if left_leaf[0] != right_leaf[0] {
            return Err(refusal(
                Code::InvalidArgument,
                "two drawables swap within one slide's z-order",
            ));
        }
        if left_leaf[3] == NONE || right_leaf[3] == NONE {
            return Err(refusal(
                Code::UnsupportedCapability,
                "this placeholder has no frame on its slide",
            ));
        }
        let (left_z, right_z) = (z_of(&left)?, z_of(&right)?);
        if left_z != right_z {
            let slide = SlideSelector::index(left_leaf[0]);
            let error = |error| failure(LeafError::KeynoteDrawables(error));
            let commit = {
                let mut edit = chain.package.edit_slide_drawables();
                edit.move_to(
                    slide,
                    DrawableSelector::index(left_leaf[3]),
                    Position::new(right_z),
                )
                .map_err(error)?
                .move_to(
                    slide,
                    DrawableSelector::index(right_leaf[3]),
                    Position::new(left_z),
                )
                .map_err(error)?;
                edit.commit().map_err(error)?
            };
            let (package, patch) = commit.into_parts();
            push(&mut chain, package, Patch::Drawables(patch));
        }
    } else {
        return Err(refusal(
            Code::InvalidArgument,
            "two slides, two paragraphs of one text or two drawables of one slide swap",
        ));
    }
    finish(
        file,
        chain.changed,
        Owner::Keynote(chain.package),
        chain.patches,
        1,
        limits,
    )
}

#[cfg(test)]
mod tests {
    use super::{TextGroup, Touched, cancels_text};

    /// A paragraph group reads, replaces and joins paragraphs like splitting
    /// and joining the whole text, and its hulls are the base's UTF-16
    /// paragraph regions (record 0340).
    #[test]
    fn paragraph_groups_match_split_and_join() {
        let base = "one\n\u{e9}\u{1f600}x\n\nlast";
        let mut group = TextGroup::new(base.to_owned());
        assert_eq!(group.count(), 4);
        let lines: Vec<_> = base.split('\n').collect();
        for (at, line) in lines.iter().enumerate() {
            assert_eq!(group.base_line(at), Some(*line));
        }
        assert_eq!(group.base_line(4), None);
        assert_eq!(
            group.hull(Touched::Whole),
            Some(0..base.encode_utf16().count())
        );
        assert_eq!(group.hull(Touched::Lines(0, 0)), Some(0..3));
        // "é😀x" is 4 UTF-16 units starting at 4.
        assert_eq!(group.hull(Touched::Lines(1, 1)), Some(4..8));
        assert_eq!(group.hull(Touched::Lines(2, 2)), Some(9..9));
        assert_eq!(group.hull(Touched::Lines(1, 3)), Some(4..14));
        assert_eq!(group.hull(Touched::Lines(3, 4)), None);

        group.set_line(2, "new".to_owned());
        group.set_line(0, "one".to_owned());
        assert_eq!(group.differs, 1);
        assert_eq!(group.text(), "one\n\u{e9}\u{1f600}x\nnew\nlast");
        group.set_line(2, String::new());
        assert_eq!(group.differs, 0);
        assert_eq!(group.text(), base);
        group.set_line(3, String::new());
        group.set_line(3, "again".to_owned());
        assert_eq!(group.differs, 1);
        assert_eq!(group.line(3), Some("again"));

        let empty = TextGroup::new(String::new());
        assert_eq!(empty.count(), 1);
        assert_eq!(empty.base_line(0), Some(""));
        assert_eq!(empty.hull(Touched::Whole), Some(0..0));
    }

    #[test]
    fn whole_texts_cancel_back_to_their_base() {
        assert!(cancels_text("a\nb", "a\nc", Touched::Lines(1, 1), "b"));
        assert!(!cancels_text("a\nb", "x\nc", Touched::Lines(1, 1), "b"));
        assert!(!cancels_text("a\nb", "a\nb\nc", Touched::Lines(1, 1), "b"));
        assert!(cancels_text("a", "b", Touched::Whole, "a"));
        assert!(!cancels_text("a", "a", Touched::Whole, "a"));
    }
}
