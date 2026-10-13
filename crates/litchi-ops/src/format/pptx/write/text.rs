//! Rich text formatting on shapes, paragraphs, runs and table cells, staged
//! through the native `set_text_properties` edits of both PPTX owners.
use super::{Commit, Editable, Pptx, Proof, allocation, bound, error, validate_text};
use crate::format::pptx::Nav;
use crate::format::pptx::node::Item;
use crate::format::pptx::spec::{
    CELL_WRITE_TABLE, PARAGRAPH_WRITE_TABLE, RUN_WRITE_TABLE, TextSetter,
};
use crate::registry::{Operation, Table};
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::{Base, Scalar};
use crate::{Code, Props, Result};
use litchi_core::Position;
use litchi_drawingml::text::Alignment;
use litchi_pptx::opened::{Rgb, TextEdit, TextProperties, TextTarget};
use std::collections::HashSet;
use std::sync::Arc;

/// The text element a path or query addresses.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum TextKind {
    Paragraph,
    Run,
    Cell,
}

fn kind_of(name: &str) -> Option<TextKind> {
    if name.eq_ignore_ascii_case("paragraph") || name.eq_ignore_ascii_case("p") {
        Some(TextKind::Paragraph)
    } else if ["run", "r", "linebreak", "br"]
        .iter()
        .any(|alias| name.eq_ignore_ascii_case(alias))
    {
        Some(TextKind::Run)
    } else if name.eq_ignore_ascii_case("tc") || name.eq_ignore_ascii_case("cell") {
        Some(TextKind::Cell)
    } else {
        None
    }
}

/// Classify a mutation target by its last path step or its query subjects,
/// without resolving it.
pub(super) fn classify(target: &TargetArg<'_>) -> Option<TextKind> {
    if let Some(path) = target.path() {
        return path
            .path()
            .steps()
            .last()
            .and_then(|step| kind_of(step.name()));
    }
    let query = target.query()?.query()?;
    let mut kinds = query
        .alternatives()
        .iter()
        .map(|alternative| alternative.subject().element().name().and_then(kind_of));
    let first = kinds.next()??;
    kinds.all(|kind| kind == Some(first)).then_some(first)
}

/// Admitted rich text assignments, owned so they outlive the props borrow.
#[derive(Clone, Default)]
#[allow(
    clippy::option_option,
    reason = "omission, removal and replacement are distinct assignments"
)]
pub(in crate::format::pptx) struct TextValues {
    text: Option<String>,
    bold: Option<Option<bool>>,
    italic: Option<Option<bool>>,
    size: Option<Option<u32>>,
    color: Option<Option<Rgb>>,
    font: Option<Option<String>>,
    latin: Option<Option<String>>,
    align: Option<Option<Alignment>>,
    fill: Option<Option<Rgb>>,
}

impl TextValues {
    pub(super) fn is_empty(&self) -> bool {
        self.text.is_none()
            && self.bold.is_none()
            && self.italic.is_none()
            && self.size.is_none()
            && self.color.is_none()
            && self.font.is_none()
            && self.latin.is_none()
            && self.align.is_none()
            && self.fill.is_none()
    }

    pub(super) fn charge(&self) -> usize {
        let text = |value: &Option<Option<String>>| {
            value
                .as_ref()
                .and_then(Option::as_ref)
                .map_or(0, String::len)
        };
        self.text
            .as_ref()
            .map_or(0, String::len)
            .saturating_add(text(&self.font))
            .saturating_add(text(&self.latin))
            .saturating_add(64)
    }

    pub(super) fn properties(&self) -> TextProperties<'_> {
        let mut properties = TextProperties::default();
        if let Some(text) = &self.text {
            properties = properties.with_text(text);
        }
        if let Some(value) = self.bold {
            properties = properties.with_bold(value);
        }
        if let Some(value) = self.italic {
            properties = properties.with_italic(value);
        }
        if let Some(value) = self.size {
            properties = properties.with_size(value);
        }
        if let Some(value) = self.color {
            properties = properties.with_color(value);
        }
        if let Some(value) = &self.font {
            properties = properties.with_font(value.as_deref());
        }
        if let Some(value) = &self.latin {
            properties = properties.with_latin(value.as_deref());
        }
        if let Some(value) = self.align {
            properties = properties.with_align(value);
        }
        if let Some(value) = self.fill {
            properties = properties.with_fill(value);
        }
        properties
    }

    /// Record one coerced assignment.
    pub(super) fn stage(
        &mut self,
        setter: TextSetter,
        value: Option<Scalar<'_>>,
        limits: &SemanticLimits,
    ) -> Result<()> {
        let owned = |value: &str| -> Result<String> {
            bound(
                LimitKind::StagedBytes,
                value.len(),
                limits.max_staged_bytes(),
            )?;
            validate_text(value)?;
            let mut owned = String::new();
            owned.try_reserve_exact(value.len()).map_err(allocation)?;
            owned.push_str(value);
            Ok(owned)
        };
        match setter {
            TextSetter::Text => {
                let text = match value {
                    Some(Scalar::Text(text)) => text,
                    None => "",
                    _ => unreachable!("text coercion"),
                };
                self.text = Some(owned(text)?);
            },
            TextSetter::Bold | TextSetter::Italic => {
                let flag = match value {
                    Some(Scalar::Bool(flag)) => Some(flag),
                    None => None,
                    _ => unreachable!("boolean coercion"),
                };
                if setter == TextSetter::Bold {
                    self.bold = Some(flag);
                } else {
                    self.italic = Some(flag);
                }
            },
            TextSetter::Size => {
                self.size = Some(match value {
                    Some(Scalar::Length(length)) => Some(font_size(length.points())?),
                    None => None,
                    _ => unreachable!("font size coercion"),
                });
            },
            TextSetter::Color | TextSetter::Fill => {
                let color = match value {
                    Some(Scalar::Color(color)) => Some(rgb(&color)?),
                    None => None,
                    _ => unreachable!("colour coercion"),
                };
                if setter == TextSetter::Color {
                    self.color = Some(color);
                } else {
                    self.fill = Some(color);
                }
            },
            TextSetter::Font | TextSetter::FontLatin => {
                let face = match value {
                    Some(Scalar::Text(face)) => Some(owned(face)?),
                    None => None,
                    _ => unreachable!("typeface coercion"),
                };
                if setter == TextSetter::Font {
                    self.font = Some(face);
                } else {
                    self.latin = Some(face);
                }
            },
            TextSetter::Align => {
                self.align = Some(match value {
                    Some(Scalar::Token(token)) => Some(alignment(token)?),
                    None => None,
                    _ => unreachable!("alignment coercion"),
                });
            },
        }
        Ok(())
    }
}

fn font_size(points: f64) -> Result<u32> {
    let hundredths = (points * 100.0).round();
    if !(f64::from(litchi_pptx::opened::MIN_FONT_SIZE)
        ..=f64::from(litchi_pptx::opened::MAX_FONT_SIZE))
        .contains(&hundredths)
    {
        return Err(error(
            Code::InvalidValue,
            "font size must be between 1pt and 4000pt",
        )
        .with_property("size"));
    }
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "the value was range-checked against the u32 font-size domain"
    )]
    Ok(hundredths as u32)
}

fn rgb(color: &crate::value::Color) -> Result<Rgb> {
    match color.base() {
        Base::Rgb(value) if value.alpha() == 255 && color.transforms().is_empty() => {
            Ok(Rgb([value.red(), value.green(), value.blue()]))
        },
        _ => Err(error(
            Code::UnsupportedCapability,
            "only opaque RGB colours without transforms are writable",
        )
        .with_suggestion("use a hex colour such as 1F4E79")),
    }
}

fn alignment(token: &str) -> Result<Alignment> {
    Ok(match token {
        "left" => Alignment::Left,
        "center" => Alignment::Center,
        "right" => Alignment::Right,
        "justify" => Alignment::Justified,
        "justLow" => Alignment::JustifiedLow,
        "dist" => Alignment::Distributed,
        "thaiDist" => Alignment::ThaiDistributed,
        _ => return Err(error(Code::InvalidValue, "unknown alignment").with_property("align")),
    })
}

fn admit<const N: usize>(
    table: &Table<TextSetter, N, 0>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<TextValues> {
    let mut values = TextValues::default();
    let mut seen = 0_u64;
    for (key, value) in props.iter() {
        let (index, entry) = table.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(litchi_core::FileFormat::Pptx)
                .with_property(key.as_str().to_owned())
        })?;
        let bit = 1_u64 << index;
        if seen & bit != 0 {
            return Err(error(
                Code::InvalidArgument,
                "a text property was assigned more than once",
            )
            .with_property(entry.spec.name));
        }
        seen |= bit;
        entry.prepare(Operation::Set, value, |setter, value| {
            values.stage(setter, value, limits)
        })?;
    }
    if values.is_empty() {
        return Err(error(
            Code::InvalidArgument,
            "no text property was assigned",
        ));
    }
    Ok(values)
}

impl Pptx {
    /// Set rich text properties on paragraphs, runs or table cells.
    pub(super) fn set_text(
        &self,
        target: TargetArg<'_>,
        kind: TextKind,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let values = match kind {
            TextKind::Paragraph => admit(&PARAGRAPH_WRITE_TABLE, props, limits)?,
            TextKind::Run => admit(&RUN_WRITE_TABLE, props, limits)?,
            TextKind::Cell => admit(&CELL_WRITE_TABLE, props, limits)?,
        };
        if kind == TextKind::Run
            && values
                .text
                .as_deref()
                .is_some_and(|text| text.contains(['\n', '\r']))
        {
            return Err(
                error(Code::InvalidValue, "run text cannot contain a line break")
                    .with_property("text"),
            );
        }
        let per_target = values.charge();
        let mut targets: Vec<(usize, TextTarget)> = Vec::new();
        let mut seen = HashSet::new();
        let mut append = |node: crate::format::pptx::Node<'_>| -> Result<()> {
            let (slide, target) = match (kind, node.item()) {
                (TextKind::Paragraph, Item::Paragraph(text)) => (
                    text.shape.container.slide(),
                    TextTarget::Paragraph {
                        shape: text.shape.handle.position(),
                        paragraph: Position::new(text.paragraph as usize),
                    },
                ),
                (TextKind::Run, Item::Run(text) | Item::LineBreak(text)) => (
                    text.shape.container.slide(),
                    TextTarget::Run {
                        shape: text.shape.handle.position(),
                        paragraph: Position::new(text.paragraph as usize),
                        run: Position::new(text.run as usize),
                    },
                ),
                (TextKind::Cell, Item::Cell(cell)) => (
                    cell.table.container.slide(),
                    TextTarget::Cell {
                        shape: cell.table.handle.position(),
                        row: Position::new(cell.row as usize),
                        column: Position::new(cell.column as usize),
                    },
                ),
                _ => {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "the selected element does not take these text properties",
                    ));
                },
            };
            let Some(slide) = slide else {
                return Err(error(
                    Code::UnsupportedCapability,
                    "master and layout text writes are not integrated",
                ));
            };
            let key = (slide, target);
            if seen.contains(&key) {
                return Ok(());
            }
            let count = targets.len().saturating_add(1);
            bound(
                LimitKind::ExpandedTargets,
                count,
                limits.max_expanded_targets(),
            )?;
            bound(
                LimitKind::StagedBytes,
                count.saturating_mul(per_target),
                limits.max_staged_bytes(),
            )?;
            targets.try_reserve(1).map_err(allocation)?;
            seen.try_reserve(1).map_err(allocation)?;
            seen.insert(key);
            targets.push((slide as usize, target));
            Ok(())
        };
        let nav = Nav::new(self);
        if let Some(path) = target.path() {
            append(crate::path::resolve::get(&nav, path)?)?;
        } else if let Some(query) = target.query() {
            let walk = crate::query::run(nav, query.clone())?;
            if !walk.plan().diagnostics().is_empty() {
                return Err(error(
                    Code::InvalidSelector,
                    "mutation selector contains an unsupported predicate",
                ));
            }
            for node in walk {
                append(node?)?;
            }
        } else {
            return Err(error(Code::InvalidArgument, "mutation target is empty"));
        }
        if targets.is_empty() {
            return Err(error(
                Code::ZeroMatches,
                "mutation selector matched no element",
            ));
        }
        targets.sort_by_key(|(slide, _)| *slide);
        let properties = values.properties();
        let mut groups: Vec<(usize, Vec<TextEdit<'_>>)> = Vec::new();
        for (slide, target) in &targets {
            if groups.last().is_none_or(|(current, _)| current != slide) {
                groups.try_reserve(1).map_err(allocation)?;
                groups.push((*slide, Vec::new()));
            }
            let Some((_, edits)) = groups.last_mut() else {
                unreachable!("a group was just pushed")
            };
            edits.try_reserve(1).map_err(allocation)?;
            edits.push(TextEdit::new(*target, properties));
        }
        self.commit_text_edits(&groups)
    }

    /// Publish slide-grouped text edits through one native edit.
    pub(super) fn commit_text_edits(
        &self,
        groups: &[(usize, Vec<TextEdit<'_>>)],
    ) -> Result<Commit> {
        let capacity = litchi_pptx::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS;
        let source = self.editable_owner()?;
        let mut operations = 0;
        let (target, proof, changed) = match &source {
            Editable::Source(package) => {
                let mut edit = package.edit()?;
                for (slide, edits) in groups {
                    for chunk in edits.chunks(capacity) {
                        operations += edit.set_text_properties(Position::new(*slide), chunk)?;
                    }
                }
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (Editable::Source(target), Proof::Source(patch), changed)
            },
            Editable::Eager(package) => {
                let mut edit = package.edit()?;
                for (slide, edits) in groups {
                    for chunk in edits.chunks(capacity) {
                        operations += edit.set_text_properties(*slide, chunk)?;
                    }
                }
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (
                    Editable::Eager(Arc::new(target)),
                    Proof::Eager(patch),
                    changed,
                )
            },
        };
        let after = if changed {
            match &target {
                Editable::Source(package) => self.adopt_source(package.clone())?,
                Editable::Eager(package) => self.adopt_eager(Arc::clone(package)),
            }
        } else {
            self.clone()
        };
        Ok(Commit {
            after,
            source,
            target,
            proof,
            operations,
        })
    }
}
