//! Immutable native shape writes; no XML editing in the facade.
pub(crate) mod coalesce;
mod find;
pub(crate) mod merge;
mod notes;
pub(crate) mod refresh;
mod slides;
pub(super) mod text;
mod transition;
use super::node::Item;
use super::spec::{Kind, SHAPE_WRITE_TABLE, ShapeSetter};
use super::{Editable, Nav, Pptx};
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TxError, LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, Error, IntoProps, Props, Result};
use litchi_core::FileFormat;
use std::collections::HashSet;
use std::sync::Arc;

pub(crate) enum Proof {
    Source(litchi_pptx::source::Patch),
    Eager(litchi_pptx::opened::Patch),
    /// A document-metadata patch (record 0253).
    Metadata(litchi_pptx::metadata::Patch),
}
impl Proof {
    pub(crate) fn retained_bytes(&self) -> usize {
        match self {
            Self::Source(patch) => patch.retained_bytes(),
            Self::Eager(patch) => patch.retained_bytes(),
            Self::Metadata(patch) => patch.changes().iter().fold(0_usize, |bytes, change| {
                bytes
                    .saturating_add(change.before().map_or(0, <[u8]>::len))
                    .saturating_add(change.after().map_or(0, <[u8]>::len))
            }),
        }
    }
}
pub(crate) struct Commit {
    pub(crate) after: Pptx,
    pub(crate) source: Editable,
    pub(crate) target: Editable,
    pub(crate) proof: Proof,
    pub(crate) operations: usize,
}
impl Commit {
    pub(crate) fn into_file(self) -> Pptx {
        self.after
    }
}
impl Editable {
    pub(crate) fn same_content(&self, other: &Self) -> Result<bool> {
        match (self, other) {
            (Self::Source(left), Self::Source(right)) => Ok(left.same_content(right)?),
            (Self::Eager(left), Self::Eager(right)) => Ok(left.same_content(right)?),
            _ => Ok(false),
        }
    }
}

/// The only PPTX shape assignments currently handled by the native write
/// table.  Values are owned so an admitted assignment can live until a
/// coalesced run commits it; the ordinary one-shot path uses the same
/// admission result immediately.
#[derive(Clone)]
pub(super) struct ShapeValues {
    pub(super) text: Option<String>,
    pub(super) geometry: [Option<i64>; 4],
    /// Rich formatting for every paragraph and run of the text body.
    pub(super) format: text::TextValues,
}

impl ShapeValues {
    pub(super) fn has_geometry(&self) -> bool {
        self.geometry.iter().any(Option::is_some)
    }

    pub(super) fn charge(&self) -> usize {
        self.text
            .as_ref()
            .map_or(0, String::len)
            .saturating_add(if self.has_geometry() {
                4 * size_of::<i64>()
            } else {
                0
            })
            .saturating_add(2 * size_of::<usize>())
            .saturating_add(if self.format.is_empty() {
                0
            } else {
                self.format.charge()
            })
    }
}

/// Coerce and validate one shape property payload through the canonical
/// registry.  Keeping this admission in the main write module means the
/// coalescer and ordinary selector writes cannot drift in aliases, duplicate
/// assignment handling, or text/length validation.
pub(super) fn admit_shape_props(props: &Props<'_>, limits: &SemanticLimits) -> Result<ShapeValues> {
    let mut values = ShapeValues {
        text: None,
        geometry: [None; 4],
        format: text::TextValues::default(),
    };
    // The PPTX table currently has fewer than 64 canonical entries.  A bit
    // set keeps duplicate aliases allocation-free on the successful path
    // while still using the table's canonical lookup index.
    let mut seen_properties = 0_u64;
    for (key, value) in props.iter() {
        let (index, entry) = SHAPE_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Pptx)
                .with_property(key.as_str().to_owned())
        })?;
        let Some(bit) = 1_u64.checked_shl(index as u32) else {
            return Err(error(
                Code::InternalError,
                "shape property registry has too many entries",
            ));
        };
        if seen_properties & bit != 0 {
            return Err(error(
                Code::InvalidArgument,
                "a shape property was assigned more than once",
            )
            .with_property(entry.spec.name));
        }
        seen_properties |= bit;
        entry.prepare(Operation::Set, value, |setter, value| {
            match setter {
                ShapeSetter::X | ShapeSetter::Y | ShapeSetter::Width | ShapeSetter::Height => {
                    let Some(Scalar::Length(value)) = value else {
                        unreachable!("shape geometry coercion")
                    };
                    let index = match setter {
                        ShapeSetter::X => 0,
                        ShapeSetter::Y => 1,
                        ShapeSetter::Width => 2,
                        ShapeSetter::Height => 3,
                        ShapeSetter::Text | ShapeSetter::Format(_) => {
                            unreachable!("shape text setter")
                        },
                    };
                    values.geometry[index] = Some(value.emu());
                },
                ShapeSetter::Text => {
                    let value = match value {
                        Some(Scalar::Text(value)) => value,
                        None => "",
                        _ => unreachable!("text coercion"),
                    };
                    bound(
                        LimitKind::StagedBytes,
                        value.len(),
                        limits.max_staged_bytes(),
                    )?;
                    validate_text(value)?;
                    let mut owned = String::new();
                    owned.try_reserve_exact(value.len()).map_err(allocation)?;
                    owned.push_str(value);
                    values.text = Some(owned);
                },
                ShapeSetter::Format(setter) => values.format.stage(setter, value, limits)?,
            }
            Ok(())
        })?;
    }
    Ok(values)
}

impl Pptx {
    pub(crate) fn set<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
    ) -> Result<Commit> {
        self.set_with_limits(target, props, &SemanticLimits::default())
    }
    pub(crate) fn set_with_limits<'p, 'v>(
        &self,
        target: TargetArg<'p>,
        props: impl IntoProps<'v>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.check_source()?;
        let result = (|| self.set_inner(target, Props::from_props(props)?, limits))()
            .map_err(|error| error.with_format(FileFormat::Pptx));
        self.check_source()?;
        result
    }
    fn set_inner(
        &self,
        target: TargetArg<'_>,
        props: Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        if let Some(commit) = self.try_set_metadata(&target, &props)? {
            return Ok(commit);
        }
        if let Some(path) = target.path()
            && path
                .path()
                .steps()
                .last()
                .is_some_and(|step| step.name().eq_ignore_ascii_case("comment"))
        {
            let node = crate::path::resolve::get(&Nav::new(self), path)?;
            let Item::Comment(comment) = node.item() else {
                return Err(error(
                    Code::UnsupportedCapability,
                    "comment writes take one legacy comment path",
                ));
            };
            return self.set_comment(comment.slide, comment.index, &props, limits);
        }
        if let Some(commit) = self.try_set_chart(&target, &props, limits)? {
            return Ok(commit);
        }
        if notes::accepts(&target, &props) {
            return self.set_notes(target, &props, limits);
        }
        if let Some(kind) = text::classify(&target) {
            return self.set_text(target, kind, &props, limits);
        }
        if slides::accepts(&target, &props) {
            return self.set_slide_properties(target, &props, limits);
        }
        // Admit caller payload before allocating a native edit or walking targets.
        let values = admit_shape_props(&props, limits)?;
        let text = values.text.as_deref();
        let geometry = values.geometry;
        let has_geometry = values.has_geometry();
        let per_target = text
            .map_or(0, str::len)
            .saturating_add(if has_geometry {
                4 * size_of::<i64>()
            } else {
                0
            })
            .saturating_add(2 * size_of::<usize>());
        let mut targets = Vec::new();
        let mut seen = HashSet::new();
        let mut append = |node: super::Node<'_>| -> Result<()> {
            let Item::Tree(tree) = node.item() else {
                return Err(error(
                    Code::UnsupportedCapability,
                    "shape text requires a slide shape",
                ));
            };
            if !matches!(tree.kind, Kind::Shape | Kind::Placeholder) {
                return Err(error(
                    Code::UnsupportedCapability,
                    "text writes for this shape kind are not integrated",
                ));
            }
            if has_geometry && tree.kind != Kind::Shape {
                return Err(error(
                    Code::UnsupportedCapability,
                    "shape geometry requires a direct p:sp shape",
                ));
            }
            let Some(slide) = tree.container.slide() else {
                return Err(error(
                    Code::UnsupportedCapability,
                    "master and layout text writes are not integrated",
                ));
            };
            let target = (slide as usize, tree.handle.position().get());
            if seen.contains(&target) {
                return Ok(());
            }
            let bounds = if has_geometry {
                let current = tree.handle.view().bounds().ok_or_else(|| {
                    error(
                        Code::MissingProperty,
                        "shape geometry requires a complete local transform",
                    )
                    .with_property("geometry")
                })?;
                let bounds = (
                    geometry[0].unwrap_or(current.x()),
                    geometry[1].unwrap_or(current.y()),
                    geometry[2].unwrap_or(current.width()),
                    geometry[3].unwrap_or(current.height()),
                );
                if bounds.2 <= 0 {
                    return Err(error(Code::InvalidValue, "shape width must be positive")
                        .with_property("width"));
                }
                if bounds.3 <= 0 {
                    return Err(error(Code::InvalidValue, "shape height must be positive")
                        .with_property("height"));
                }
                Some(bounds)
            } else {
                None
            };
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
            seen.insert(target);
            targets.push(Target {
                slide: target.0,
                position: target.1,
                // An empty body (a new placeholder) has no run slot for the
                // text writer; its text goes through the body edit.
                slot: tree
                    .handle
                    .view()
                    .text()
                    .is_some_and(|text| !text.is_empty()),
                bounds,
            });
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
        targets.sort_unstable_by_key(|target| (target.slide, target.position));
        let body_edits = body_edits(&targets, text, &values.format)?;
        let source = self.editable_owner()?;
        let mut operations = 0;
        let (target, proof, changed) = match &source {
            Editable::Source(package) => {
                let mut edit = package.edit()?;
                if has_geometry {
                    stage_properties(&targets, text, |slide, replacements| {
                        operations += if replacements.len() == 1 {
                            edit.set_shape_properties(slide, replacements[0])?
                        } else {
                            edit.set_shape_property_chunks(slide, replacements)?
                        };
                        Ok(())
                    })?;
                } else if let Some(text) = text {
                    stage(&targets, text, |slide, replacements| {
                        operations += if replacements.len() == 1 {
                            edit.set_shape_texts(slide, replacements[0])?
                        } else {
                            edit.set_shape_text_chunks(slide, replacements)?
                        };
                        Ok(())
                    })?;
                }
                for (slide, edits) in &body_edits {
                    for chunk in
                        edits.chunks(litchi_pptx::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS)
                    {
                        operations += edit.set_text_properties(*slide, chunk)?;
                    }
                }
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (Editable::Source(target), Proof::Source(patch), changed)
            },
            Editable::Eager(package) => {
                let mut edit = package.edit()?;
                if has_geometry {
                    stage_properties(&targets, text, |slide, replacements| {
                        operations += if replacements.len() == 1 {
                            edit.set_shape_properties(slide, replacements[0])?
                        } else {
                            edit.set_shape_property_chunks(slide, replacements)?
                        };
                        Ok(())
                    })?;
                } else if let Some(text) = text {
                    stage(&targets, text, |slide, replacements| {
                        operations += if replacements.len() == 1 {
                            edit.set_shape_texts(slide, replacements[0])?
                        } else {
                            edit.set_shape_text_chunks(slide, replacements)?
                        };
                        Ok(())
                    })?;
                }
                for (slide, edits) in &body_edits {
                    for chunk in
                        edits.chunks(litchi_pptx::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS)
                    {
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
        let after = if !changed {
            self.clone()
        } else {
            match &target {
                Editable::Source(package) => self.adopt_source(package.clone())?,
                Editable::Eager(package) => self.adopt_eager(Arc::clone(package)),
            }
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

#[derive(Clone, Copy)]
struct Target {
    slot: bool,
    slide: usize,
    position: usize,
    bounds: Option<(i64, i64, i64, i64)>,
}

/// Slide-grouped body edits: rich formatting for every target, and the text
/// of targets whose body has no run slot.
fn body_edits<'t>(
    targets: &[Target],
    text: Option<&'t str>,
    format: &'t text::TextValues,
) -> Result<Vec<(usize, Vec<litchi_pptx::opened::TextEdit<'t>>)>> {
    let formatting = format.properties();
    let filled = text.map(|text| formatting.with_text(text));
    let mut groups: Vec<(usize, Vec<litchi_pptx::opened::TextEdit<'t>>)> = Vec::new();
    for target in targets {
        let properties = match filled {
            Some(filled) if !target.slot => filled,
            _ if !format.is_empty() => formatting,
            _ => continue,
        };
        if groups
            .last()
            .is_none_or(|(slide, _)| *slide != target.slide)
        {
            groups.try_reserve(1).map_err(allocation)?;
            groups.push((target.slide, Vec::new()));
        }
        let Some((_, edits)) = groups.last_mut() else {
            unreachable!("a group was just pushed")
        };
        edits.try_reserve(1).map_err(allocation)?;
        edits.push(litchi_pptx::opened::TextEdit::new(
            litchi_pptx::opened::TextTarget::Shape {
                shape: litchi_core::Position::new(target.position),
            },
            properties,
        ));
    }
    Ok(groups)
}

fn stage(
    targets: &[Target],
    text: &str,
    mut apply: impl FnMut(usize, &[&[litchi_pptx::opened::ShapeTextReplacement<'_>]]) -> Result<()>,
) -> Result<()> {
    // All same-slide targets share publication. Each input chunk still
    // observes native target and text-byte limits.
    for group in targets.chunk_by(|left, right| left.slide == right.slide) {
        let mut replacements = Vec::new();
        replacements
            .try_reserve_exact(group.len())
            .map_err(allocation)?;
        for target in group.iter().filter(|target| target.slot) {
            replacements.push(litchi_pptx::opened::ShapeTextReplacement::new(
                target.position.into(),
                text,
            ));
        }
        if replacements.is_empty() {
            continue;
        }
        apply_chunks(group[0].slide, &replacements, text.len(), &mut apply)?;
    }
    Ok(())
}

fn stage_properties(
    targets: &[Target],
    text: Option<&str>,
    mut apply: impl FnMut(
        usize,
        &[&[litchi_pptx::opened::ShapePropertiesReplacement<'_>]],
    ) -> Result<()>,
) -> Result<()> {
    for group in targets.chunk_by(|left, right| left.slide == right.slide) {
        let mut replacements = Vec::new();
        replacements
            .try_reserve_exact(group.len())
            .map_err(allocation)?;
        for target in group {
            let mut replacement =
                litchi_pptx::opened::ShapePropertiesReplacement::new(target.position.into());
            if let Some(text) = text.filter(|_| target.slot) {
                replacement = replacement.with_text(text);
            }
            if let Some(bounds) = target.bounds {
                replacement = replacement.with_bounds(bounds);
            }
            replacements.push(replacement);
        }
        apply_chunks(
            group[0].slide,
            &replacements,
            text.map_or(0, str::len),
            &mut apply,
        )?;
    }
    Ok(())
}

fn apply_chunks<T>(
    slide: usize,
    replacements: &[T],
    text_bytes: usize,
    apply: &mut impl FnMut(usize, &[&[T]]) -> Result<()>,
) -> Result<()> {
    let capacity = litchi_pptx::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS;
    // An individually oversized value reaches the ordinary native typed
    // refusal. Uniform fanout values make the byte boundary exact.
    let chunk_size = litchi_pptx::opened::Limits::default()
        .max_text_bytes()
        .checked_div(text_bytes)
        .unwrap_or(capacity)
        .clamp(1, capacity);
    if replacements.len() <= chunk_size {
        return apply(slide, std::slice::from_ref(&replacements));
    }
    let mut chunks = Vec::new();
    chunks
        .try_reserve_exact(replacements.len().div_ceil(chunk_size))
        .map_err(allocation)?;
    chunks.extend(replacements.chunks(chunk_size));
    apply(slide, &chunks)
}

fn bound(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(TxError::Limit {
            kind,
            observed,
            maximum,
        }
        .into());
    }
    Ok(())
}
fn allocation(source: std::collections::TryReserveError) -> Error {
    TxError::Allocation { source }.into()
}
fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Pptx)
}

pub(super) fn validate_text(value: &str) -> Result<()> {
    if !value.chars().all(|ch| matches!(ch, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')) {
        return Err(error(Code::InvalidValue, "text contains a character that XML cannot represent").with_property("text"));
    }
    Ok(())
}
