//! Unified paragraph and run writes over the native immutable RTF editor.
//!
//! Target resolution is performed against one retained `Rtf` snapshot.  The
//! complete selection is then staged as one native paragraph batch, which
//! keeps selector fan-out atomic and preserves the native source/options
//! policy through publication.

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_rtf::Alignment;
use litchi_rtf::UnderlineStyle;
use litchi_rtf::edit::{
    EditOptions, ParagraphAlignmentUpdate, ParagraphLayoutPatch, ParagraphLayoutUpdate,
    ParagraphTextReplacement, Patch, RunFormatPatch, RunFormatUpdate, RunTextInsertion,
    RunTextReplacement, SemanticLimits as NativeLimits,
};
use std::num::NonZeroU16;

use super::{Nav, Node, Rtf};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::placement::Placement;
use crate::props::{IntoProps, PropValue, Props};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::{LengthSpec, Scalar, ValueType};

pub(crate) mod coalesce;
mod destination;
mod find;
mod marks;
pub(crate) mod removal;

/// The native proof and immutable endpoints of one RTF write.
pub(crate) struct Commit {
    pub(crate) before: Rtf,
    pub(crate) after: Rtf,
    pub(crate) patch: Patch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

/// One value held inline, or a fan-out on the heap. Exact-path writes, the
/// common small edit, resolve and stage without a heap allocation.
pub(super) enum Batch<T> {
    None,
    One(T),
    Many(Vec<T>),
}

impl<T> Batch<T> {
    pub(super) fn as_slice(&self) -> &[T] {
        match self {
            Self::None => &[],
            Self::One(value) => std::slice::from_ref(value),
            Self::Many(values) => values,
        }
    }

    fn is_empty(&self) -> bool {
        self.as_slice().is_empty()
    }

    /// `make` applied to every item, inline for a single item.
    fn map<S>(items: &[S], mut make: impl FnMut(&S) -> Result<T>) -> Result<Self> {
        match items {
            [] => Ok(Self::None),
            [item] => Ok(Self::One(make(item)?)),
            _ => {
                let mut values = Vec::new();
                values
                    .try_reserve_exact(items.len())
                    .map_err(|source| TransactionError::Allocation { source })?;
                for item in items {
                    values.push(make(item)?);
                }
                Ok(Self::Many(values))
            },
        }
    }
}

/// Resolve a path, taking a direct route for exact positional main-story
/// paths (`/body`, `/body/p[N]`, `/body/p[N]/r[M]`). Any other path, and a
/// miss, goes through the generic resolver, so its errors are unchanged.
pub(super) fn resolve_exact<'f>(file: &'f Rtf, path: &PathArg<'_>) -> Result<Node<'f>> {
    if !path.raw
        && path.is_absolute()
        && matches!(
            path.origin(),
            crate::path::arg::Origin::Typed | crate::path::arg::Origin::Text(_)
        )
    {
        let package = &file.0.package;
        match path.path().steps() {
            [body] if coalesce::body_step(body) => return Ok(Node::body(file)),
            steps if steps.len() <= 3 => match coalesce::direct_target(path.path()) {
                Some(coalesce::Target::Paragraph(position)) => {
                    if let Some(paragraph) = package.body().paragraphs().nth(position.get()) {
                        return Ok(Node::paragraph(file, position, paragraph));
                    }
                },
                Some(coalesce::Target::Run { paragraph, run }) => {
                    if let Some(native) = package
                        .body()
                        .paragraphs()
                        .nth(paragraph.get())
                        .and_then(|value| value.runs().nth(run.get()))
                    {
                        return Ok(Node::run(file, paragraph, run, native));
                    }
                },
                _ => {},
            },
            _ => {},
        }
        if let [body, paragraph_step, picture_step] = path.path().steps()
            && coalesce::body_step(body)
            && (paragraph_step.name().eq_ignore_ascii_case("p")
                || paragraph_step.name().eq_ignore_ascii_case("paragraph"))
            && ["picture", "image", "img"]
                .iter()
                .any(|name| picture_step.name().eq_ignore_ascii_case(name))
            && let (Some(position), Some(ordinal)) = (paragraph_step.index(), picture_step.index())
            && let Some(paragraph) = package.body().paragraphs().nth(position.get())
            && let Some(index) =
                super::stories::pictures(file, super::stories::anchor_range(paragraph))
                    .nth(ordinal.get())
            && let Some(value) = package.pictures().get(index)
        {
            return Ok(Node::paragraph(file, position, paragraph).child(
                super::Kind::Picture,
                super::node::Item::Picture { index, value },
                ordinal,
            ));
        }
    }
    crate::path::resolve::get(&Nav::new(file), path)
}

/// Add one ordinary main-story paragraph through the native lifecycle editor.
///
/// The parent and placement are resolved against the unchanged source
/// snapshot.  The native edit owns the structural splice, source-closure
/// proof, and reversible history; this adapter only admits the unified node
/// shape and translates its final insertion position.
pub(crate) fn add_with_limits<'p, 'n>(
    file: &Rtf,
    parent: PathArg<'p>,
    node: NewNode<'n>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    file.admit_edit()?;
    let parent = resolve_exact(file, &parent)?;
    if node.kind().eq_ignore_ascii_case("r") || node.kind().eq_ignore_ascii_case("run") {
        if parent.element_kind() != super::Kind::Paragraph {
            return Err(Error::new(
                Code::UnsupportedType,
                "RTF runs can only be added directly under a paragraph",
            )
            .with_format(FileFormat::Rtf));
        }
        return add_run_with_limits(file, parent, node, limits);
    }
    if node.kind().eq_ignore_ascii_case("tab") {
        let paragraph = parent
            .body_paragraph()
            .and(parent.paragraph_position())
            .filter(|_| parent.element_kind() == super::Kind::Paragraph)
            .ok_or_else(|| {
                Error::new(
                    Code::UnsupportedType,
                    "RTF tab stops can only be added directly under a body paragraph",
                )
                .with_format(FileFormat::Rtf)
            })?;
        // Tab stops are kept in position order, as ODT keeps them; a
        // placement does not reorder them.
        return marks::add_tab(file, paragraph.get(), node.props(), limits);
    }
    if parent.element_kind() != super::Kind::Body {
        return Err(Error::new(
            Code::UnsupportedType,
            "RTF paragraphs can only be added directly under the body",
        )
        .with_format(FileFormat::Rtf));
    }
    if !node.kind().eq_ignore_ascii_case("p") && !node.kind().eq_ignore_ascii_case("paragraph") {
        return Err(Error::new(
            Code::UnsupportedElement,
            "the RTF structural verb supports paragraph nodes only",
        )
        .with_format(FileFormat::Rtf));
    }
    let text = add_text(node.props())?;
    let count = file.0.package.paragraph_count();
    let position = add_position(file, count, node.placement())?;

    bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    let staged_bytes = text
        .len()
        .checked_add(1)
        .and_then(|bytes| bytes.checked_add(size_of::<usize>()))
        .and_then(|bytes| bytes.checked_add(size_of::<String>()))
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    bound(
        LimitKind::StagedBytes,
        staged_bytes,
        limits.max_staged_bytes(),
    )?;

    let native_limits = NativeLimits::builder()
        .max_operations(limits.max_operations())
        .map_err(Error::from)?
        .max_replacement_bytes(limits.max_staged_bytes())
        .map_err(Error::from)?
        .build()
        .map_err(Error::from)?;
    let options = EditOptions::new().semantic_limits(native_limits);
    let mut edit = file.0.package.edit_with_options(options);
    edit.insert_paragraph(Position::new(position), text)
        .map_err(Error::from)?;
    let native = edit.commit().map_err(Error::from)?;

    let changed = native.diagnostics().changed();
    let operation_count = native.diagnostics().operation_count();
    let (package, patch) = native.into_parts();
    let snapshot_bytes = patch_snapshot_bytes(&patch);
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let before = file.clone();
    let after = if changed {
        file.with_package(package)
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok((
        Commit {
            before,
            after,
            patch,
            operations: operation_count,
            snapshot_bytes,
        },
        paragraph_path(position),
    ))
}

/// Add one ordinary character run under a body paragraph through the native
/// source-preserving insertion transaction.
fn add_run_with_limits<'n>(
    file: &Rtf,
    parent: Node<'_>,
    node: NewNode<'n>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let Some(paragraph) = parent.paragraph_position() else {
        return Err(Error::new(
            Code::UnsupportedType,
            "RTF run additions require a body paragraph",
        )
        .with_format(FileFormat::Rtf));
    };
    let Some(source_paragraph) = parent.body_paragraph() else {
        return Err(Error::new(
            Code::UnsupportedType,
            "RTF run additions require a body paragraph",
        )
        .with_format(FileFormat::Rtf));
    };
    let (text, patch) = run_add_properties(node.props())?;
    let count = source_paragraph.runs().count();
    let run = run_add_position(file, paragraph, count, node.placement())?;

    bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    let staged_bytes = text
        .len()
        .checked_add(size_of::<RunTextInsertion>())
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    bound(
        LimitKind::StagedBytes,
        staged_bytes,
        limits.max_staged_bytes(),
    )?;

    let native_limits = NativeLimits::builder()
        .max_operations(limits.max_operations())
        .map_err(Error::from)?
        .max_replacement_bytes(limits.max_staged_bytes())
        .map_err(Error::from)?
        .build()
        .map_err(Error::from)?;
    let options = EditOptions::new().semantic_limits(native_limits);
    let mut edit = file.0.package.edit_with_options(options);
    edit.insert_run(paragraph, Position::new(run), text, patch)
        .map_err(Error::from)?;
    let native = edit.commit().map_err(Error::from)?;

    let changed = native.diagnostics().changed();
    let operation_count = native.diagnostics().operation_count();
    let (package, patch) = native.into_parts();
    let snapshot_bytes = patch_snapshot_bytes(&patch);
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let before = file.clone();
    let after = if changed {
        file.with_package(package)
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok((
        Commit {
            before,
            after,
            patch,
            operations: operation_count,
            snapshot_bytes,
        },
        run_path(paragraph, run),
    ))
}

fn run_add_properties<'a>(props: &'a Props<'_>) -> Result<(&'a str, RunFormatPatch)> {
    let text = run_text(props)?.ok_or_else(|| {
        Error::new(
            Code::MissingProperty,
            "RTF run additions require a `text` property",
        )
        .with_property("text")
        .with_format(FileFormat::Rtf)
    })?;
    Ok((text, run_format_patch(props)?))
}

fn run_add_position(
    file: &Rtf,
    paragraph: Position,
    count: usize,
    placement: Option<&Placement<'_>>,
) -> Result<usize> {
    let nav = Nav::new(file);
    let position = match placement {
        None | Some(Placement::End) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => position.get(),
        Some(Placement::Before(path)) => {
            let node = crate::path::resolve::get(&nav, path)?;
            run_sibling_position(node, paragraph)?
        },
        Some(Placement::After(path)) => {
            run_sibling_position(crate::path::resolve::get(&nav, path)?, paragraph)?
                .saturating_add(1)
        },
        Some(Placement::Into(path, position)) => {
            let node = crate::path::resolve::get(&nav, path)?;
            if node.element_kind() != super::Kind::Paragraph
                || node.paragraph_position() != Some(paragraph)
            {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "RTF run placements must stay within the supplied paragraph",
                )
                .with_format(FileFormat::Rtf));
            }
            position.map_or(count, Position::get)
        },
    };
    if position > count {
        return Err(Error::new(
            Code::OutOfRange,
            "the RTF run insertion position is outside the paragraph",
        )
        .with_format(FileFormat::Rtf));
    }
    Ok(position)
}

fn run_sibling_position(node: Node<'_>, paragraph: Position) -> Result<usize> {
    if node.element_kind() != super::Kind::Run || node.paragraph_position() != Some(paragraph) {
        return Err(Error::new(
            Code::UnsupportedType,
            "RTF run placements require a run in the supplied paragraph",
        )
        .with_format(FileFormat::Rtf));
    }
    node.run_position().map(Position::get).ok_or_else(|| {
        Error::new(
            Code::UnsupportedType,
            "RTF run placement target has no source position",
        )
        .with_format(FileFormat::Rtf)
    })
}

fn add_text<'a>(props: &'a Props<'_>) -> Result<&'a str> {
    if props.is_empty() {
        return Ok("");
    }
    if props.len() != 1 {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "RTF paragraph additions accept exactly one `text` property",
        )
        .with_format(FileFormat::Rtf));
    }
    let Some((key, value)) = props.iter().next() else {
        unreachable!("the length check above admits one property")
    };
    if !key.eq_ignore_ascii_case("text") {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "the requested RTF add property is not supported",
        )
        .with_property(key.as_str().to_owned())
        .with_format(FileFormat::Rtf));
    }
    value.as_text().ok_or_else(|| {
        Error::new(
            Code::InvalidValue,
            "RTF paragraph text requires a text property value",
        )
        .with_property(key.as_str().to_owned())
        .with_format(FileFormat::Rtf)
    })
}

fn add_position(file: &Rtf, count: usize, placement: Option<&Placement<'_>>) -> Result<usize> {
    let nav = Nav::new(file);
    let position = match placement {
        None | Some(Placement::End) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => position.get(),
        Some(Placement::Before(path)) => {
            paragraph_position(crate::path::resolve::get(&nav, path)?)?
        },
        Some(Placement::After(path)) => {
            paragraph_position(crate::path::resolve::get(&nav, path)?)?.saturating_add(1)
        },
        Some(Placement::Into(path, position)) => {
            let parent = crate::path::resolve::get(&nav, path)?;
            if parent.element_kind() != super::Kind::Body {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "RTF paragraphs can only be added within the body",
                )
                .with_format(FileFormat::Rtf));
            }
            position.map_or(count, Position::get)
        },
    };
    if position > count {
        return Err(Error::new(
            Code::OutOfRange,
            "the RTF paragraph insertion position is outside the body",
        )
        .with_format(FileFormat::Rtf));
    }
    Ok(position)
}

fn paragraph_path(position: usize) -> Path<'static> {
    Path::from_steps([
        Step::new("body"),
        Step::new("p").with(Predicate::Index(Position::new(position))),
    ])
}

fn run_path(paragraph: Position, run: usize) -> Path<'static> {
    Path::from_steps([
        Step::new("body"),
        Step::new("p").with(Predicate::Index(paragraph)),
        Step::new("r").with(Predicate::Index(Position::new(run))),
    ])
}

/// Admit paragraph-property assignments and publish all selected paragraphs
/// in one native edit.
pub(crate) fn set_with_limits<'p, 'v>(
    file: &Rtf,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.admit_edit()?;
    let props = Props::from_props(props)?;
    let mut before_text_bytes = 0usize;
    let resolved = resolve_mutation_targets(
        file,
        &target,
        limits.max_expanded_targets(),
        Code::UnsupportedType,
        |node| {
            before_text_bytes = before_text_bytes
                .saturating_add(node.native_run().map_or(0, |run| run.text().len()));
        },
    )?;
    let positions = match resolved {
        MutationTargets::Run(targets) => {
            return publish_run_updates(
                file,
                targets.as_slice(),
                run_format_patch(&props)?,
                run_text(&props)?,
                before_text_bytes,
                limits,
            );
        },
        MutationTargets::Destination(targets) => {
            return destination::publish_texts(file, targets.as_slice(), &props, limits);
        },
        MutationTargets::Picture(_) => {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "RTF picture properties are read-only; replace or remove the picture",
            )
            .with_format(FileFormat::Rtf));
        },
        MutationTargets::Mark(targets) => {
            return marks::set(file, targets.as_slice(), &props, limits);
        },
        MutationTargets::Paragraph(positions) => positions,
    };
    let positions = positions.as_slice();
    let ParagraphProperties {
        text,
        alignment,
        layout,
    } = paragraph_properties(&props)?;

    let property_count = usize::from(text.is_some())
        .checked_add(usize::from(alignment.is_some()))
        .and_then(|count| count.checked_add(usize::from(layout.is_some())))
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    let operations = positions
        .len()
        .checked_mul(property_count)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if operations > limits.max_operations() {
        return Err(TransactionError::Limit {
            kind: LimitKind::Operations,
            observed: operations,
            maximum: limits.max_operations(),
        }
        .into());
    }

    // Count the borrowed source text for every selected paragraph before
    // allocating any owned native replacement strings.
    let replacement_bytes = text
        .map(|text| {
            positions
                .len()
                .checked_mul(text.len())
                .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))
        })
        .transpose()?
        .unwrap_or(0);
    let text_action_bytes = text.map_or(Ok(0), |_| {
        positions
            .len()
            .checked_mul(size_of::<ParagraphTextReplacement>())
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))
    })?;
    let alignment_action_bytes = alignment.map_or(Ok(0), |_| {
        positions
            .len()
            .checked_mul(size_of::<ParagraphAlignmentUpdate>())
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))
    })?;
    let layout_action_bytes = layout.map_or(Ok(0), |_| {
        positions
            .len()
            .checked_mul(size_of::<ParagraphLayoutUpdate>())
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))
    })?;
    let action_bytes = text_action_bytes
        .checked_add(alignment_action_bytes)
        .and_then(|bytes| bytes.checked_add(layout_action_bytes))
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    let staged_bytes = replacement_bytes
        .checked_add(action_bytes)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged_bytes > limits.max_staged_bytes() {
        return Err(staged_limit(staged_bytes, limits.max_staged_bytes()));
    }

    let replacements = match text {
        Some(text) => Batch::map(positions, |position| {
            let mut replacement = String::new();
            replacement
                .try_reserve_exact(text.len())
                .map_err(|source| TransactionError::Allocation { source })?;
            replacement.push_str(text);
            Ok(ParagraphTextReplacement::new(*position, replacement))
        })?,
        None => Batch::None,
    };
    let alignments = match alignment {
        Some(alignment) => Batch::map(positions, |position| {
            Ok(ParagraphAlignmentUpdate::new(*position, alignment))
        })?,
        None => Batch::None,
    };
    let layouts = match layout {
        Some(layout) => Batch::map(positions, |position| {
            Ok(ParagraphLayoutUpdate::new(*position, layout))
        })?,
        None => Batch::None,
    };

    let native_limits = NativeLimits::builder()
        .max_operations(limits.max_operations())
        .map_err(Error::from)?
        .max_replacement_bytes(limits.max_staged_bytes())
        .map_err(Error::from)?
        .build()
        .map_err(Error::from)?;
    let options = EditOptions::new().semantic_limits(native_limits);
    let mut edit = file.0.package.edit_with_options(options);
    if !replacements.is_empty() {
        edit.replace_body_paragraph_texts(replacements.as_slice())
            .map_err(Error::from)?;
    }
    if !alignments.is_empty() {
        edit.set_body_paragraph_alignments(alignments.as_slice())
            .map_err(Error::from)?;
    }
    if !layouts.is_empty() {
        // The native editor deliberately rejects text-plus-layout batches
        // with a typed `ParagraphLayoutTextConflict`; keeping both staged in
        // this one edit preserves that policy instead of silently splitting
        // the caller's request into two publications.
        edit.patch_body_paragraph_layouts(layouts.as_slice())
            .map_err(Error::from)?;
    }
    let native = edit.commit().map_err(Error::from)?;

    let changed = native.diagnostics().changed();
    let operation_count = native.diagnostics().operation_count();
    let (package, patch) = native.into_parts();
    let snapshot_bytes = patch_snapshot_bytes(&patch);
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let before = file.clone();
    let after = if changed {
        file.with_package(package)
    } else {
        // Preserve exact no-op identity, including the native package and
        // the origin sidecar used by atomic save.
        file.clone()
    };
    file.check_source()?;
    Ok(Commit {
        before,
        after,
        patch,
        operations: operation_count,
        snapshot_bytes,
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MutationKind {
    Paragraph,
    Run,
    Destination,
    Picture,
    Mark,
}

enum MutationTargets {
    Paragraph(Batch<Position>),
    Run(Batch<(Position, Position)>),
    /// Retained destination stories: cells, header/footer paragraphs, notes
    /// and comments.
    Destination(Batch<destination::Target>),
    /// Native picture indices.
    Picture(Batch<usize>),
    /// Sections, bookmarks, fields, tables of contents and tab stops
    /// (record 0328).
    Mark(Batch<marks::Target>),
}

/// The mutation family of `node`, or `None` when it is not a mutation target.
fn mutation_kind(node: &Node<'_>) -> Option<MutationKind> {
    match node.element_kind() {
        super::Kind::Paragraph if node.body_paragraph().is_some() => Some(MutationKind::Paragraph),
        super::Kind::Run => Some(MutationKind::Run),
        super::Kind::Picture => Some(MutationKind::Picture),
        super::Kind::Section
        | super::Kind::Bookmark
        | super::Kind::Field
        | super::Kind::Toc
        | super::Kind::Tab => marks::Target::of(node).map(|_| MutationKind::Mark),
        _ => destination::Target::of(node).map(|_| MutationKind::Destination),
    }
}

/// Resolve the complete mutation target in one bounded walk. A query that
/// mixes target families (body paragraphs, runs, retained destination
/// stories, pictures) is rejected while it is being collected, so the
/// selector is traversed once and no partially resolved native batch can be
/// published.
fn resolve_mutation_targets(
    file: &Rtf,
    target: &TargetArg<'_>,
    maximum: usize,
    unsupported_kind: Code,
    mut selected: impl FnMut(Node<'_>),
) -> Result<MutationTargets> {
    let mut kind = None;
    let mut paragraphs = Vec::new();
    let mut runs = Vec::new();
    let mut destinations = Vec::new();
    let mut pictures = Vec::new();
    let mut mark_targets = Vec::new();

    let mut append = |node: Node<'_>| -> Result<()> {
        let Some(current) = mutation_kind(&node) else {
            // Projected native concepts the editor cannot change yet are
            // capability gaps; structural containers are the wrong type.
            let code = match node.element_kind() {
                super::Kind::Document | super::Kind::Body => unsupported_kind,
                _ => Code::UnsupportedCapability,
            };
            return Err(Error::new(
                code,
                "RTF mutations target paragraphs, runs, table cells, header/footer stories, notes, comments or pictures",
            )
            .with_format(FileFormat::Rtf));
        };
        if let Some(previous) = kind {
            if previous != current {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "RTF mutations cannot mix body paragraph, run, destination-story and picture targets",
                )
                .with_format(FileFormat::Rtf));
            }
        } else {
            kind = Some(current);
        }
        match current {
            MutationKind::Paragraph => append_paragraph(&mut paragraphs, node, maximum)?,
            MutationKind::Run => append_run(&mut runs, node, maximum)?,
            MutationKind::Destination => {
                destination::append(&mut destinations, &node, maximum)?;
            },
            MutationKind::Picture => destination::append_picture(&mut pictures, &node, maximum)?,
            MutationKind::Mark => marks::append(&mut mark_targets, &node, maximum)?,
        }
        selected(node);
        Ok(())
    };

    if let Some(path) = target.path() {
        // An exact destination-story path resolves straight to its native
        // story; a miss takes the generic route for its canonical error.
        if !path.raw
            && path.is_absolute()
            && maximum != 0
            && let Some(coalesce::Target::DestinationPath(direct)) =
                coalesce::direct_target(path.path())
            && let Ok(target) = destination::resolve(file, direct)
        {
            file.check_source()?;
            return Ok(MutationTargets::Destination(Batch::One(target)));
        }
        // An exact path names one node: classify it without collecting.
        let node = resolve_exact(file, path)?;
        if maximum == 0 {
            return Err(TransactionError::Limit {
                kind: LimitKind::ExpandedTargets,
                observed: 1,
                maximum,
            }
            .into());
        }
        let single = match mutation_kind(&node) {
            Some(MutationKind::Paragraph) => node
                .paragraph_position()
                .map(|position| MutationTargets::Paragraph(Batch::One(position))),
            Some(MutationKind::Run) => node
                .paragraph_position()
                .zip(node.run_position())
                .map(|target| MutationTargets::Run(Batch::One(target))),
            Some(MutationKind::Destination) => destination::Target::of(&node)
                .map(|target| MutationTargets::Destination(Batch::One(target))),
            Some(MutationKind::Picture) => match node.item {
                super::node::Item::Picture { index, .. } => {
                    Some(MutationTargets::Picture(Batch::One(index)))
                },
                _ => None,
            },
            Some(MutationKind::Mark) => {
                marks::Target::of(&node).map(|target| MutationTargets::Mark(Batch::One(target)))
            },
            None => None,
        };
        if let Some(single) = single {
            selected(node);
            file.check_source()?;
            return Ok(single);
        }
        append(node)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the RTF mutation selector contains an unsupported predicate",
            )
            .with_format(FileFormat::Rtf));
        }
        for node in walk {
            append(node?)?;
        }
    } else {
        return Err(
            Error::new(Code::InvalidArgument, "the RTF mutation target is empty")
                .with_format(FileFormat::Rtf),
        );
    }

    file.check_source()?;
    match kind {
        Some(MutationKind::Paragraph) if !paragraphs.is_empty() => {
            paragraphs.sort_unstable();
            paragraphs.dedup();
            Ok(MutationTargets::Paragraph(Batch::Many(paragraphs)))
        },
        Some(MutationKind::Run) if !runs.is_empty() => {
            runs.sort_unstable();
            runs.dedup();
            Ok(MutationTargets::Run(Batch::Many(runs)))
        },
        Some(MutationKind::Destination) if !destinations.is_empty() => {
            destinations.sort_unstable();
            destinations.dedup();
            Ok(MutationTargets::Destination(Batch::Many(destinations)))
        },
        Some(MutationKind::Picture) if !pictures.is_empty() => {
            pictures.sort_unstable();
            pictures.dedup();
            Ok(MutationTargets::Picture(Batch::Many(pictures)))
        },
        Some(MutationKind::Mark) if !mark_targets.is_empty() => {
            mark_targets.sort_unstable();
            mark_targets.dedup();
            Ok(MutationTargets::Mark(Batch::Many(mark_targets)))
        },
        Some(_) | None => Err(Error::new(
            Code::ZeroMatches,
            "the RTF mutation selector matched no paragraph or run",
        )
        .with_format(FileFormat::Rtf)),
    }
}

/// Admit run text/format properties and publish all selected runs in one native
/// transaction.  `RunFormatPatch` is omission-preserving, so one compound
/// setter keeps every unselected character facet and every unrelated source
/// destination intact.
fn publish_run_updates(
    file: &Rtf,
    targets: &[(Position, Position)],
    patch: RunFormatPatch,
    text: Option<&str>,
    before_text_bytes: usize,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let native_limits = NativeLimits::builder()
        .max_operations(limits.max_operations())
        .map_err(Error::from)?
        .max_replacement_bytes(limits.max_staged_bytes())
        .map_err(Error::from)?
        .build()
        .map_err(Error::from)?;
    let options = EditOptions::new().semantic_limits(native_limits);
    let mut edit = file.0.package.edit_with_options(options);
    let native = if patch.is_empty() && text.is_none() {
        // An empty property map is a validated, exact no-op.  Still resolve
        // the target so a miss or an abusive selector cannot be hidden behind
        // the no-op shortcut. It carries no native operation or staging
        // charge because it creates no native update.
        edit.commit().map_err(Error::from)?
    } else {
        let operations = targets.len();
        if operations > limits.max_operations() {
            return Err(TransactionError::Limit {
                kind: LimitKind::Operations,
                observed: operations,
                maximum: limits.max_operations(),
            }
            .into());
        }
        let per_target = if let Some(text) = text {
            size_of::<RunTextReplacement>().checked_add(text.len())
        } else {
            Some(size_of::<RunFormatUpdate>())
        }
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        let staged_bytes = operations
            .checked_mul(per_target)
            .and_then(|bytes| bytes.checked_add(if text.is_some() { before_text_bytes } else { 0 }))
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        if staged_bytes > limits.max_staged_bytes() {
            return Err(staged_limit(staged_bytes, limits.max_staged_bytes()));
        }

        if let Some(text) = text {
            let updates = Batch::map(targets, |(paragraph, run)| {
                let mut replacement = String::new();
                replacement
                    .try_reserve_exact(text.len())
                    .map_err(|source| TransactionError::Allocation { source })?;
                replacement.push_str(text);
                Ok(RunTextReplacement::new(*paragraph, *run, replacement).with_format(patch))
            })?;
            edit.replace_body_run_texts(updates.as_slice())
                .map_err(Error::from)?;
        } else {
            let updates = Batch::map(targets, |(paragraph, run)| {
                Ok(RunFormatUpdate::new(*paragraph, *run, patch))
            })?;
            edit.patch_body_run_formats(updates.as_slice())
                .map_err(Error::from)?;
        }
        edit.commit().map_err(Error::from)?
    };

    let changed = native.diagnostics().changed();
    let operation_count = native.diagnostics().operation_count();
    let (package, patch) = native.into_parts();
    let snapshot_bytes = patch_snapshot_bytes(&patch);
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let before = file.clone();
    let after = if changed {
        file.with_package(package)
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok(Commit {
        before,
        after,
        patch,
        operations: operation_count,
        snapshot_bytes,
    })
}

fn append_run(
    targets: &mut Vec<(Position, Position)>,
    node: Node<'_>,
    maximum: usize,
) -> Result<()> {
    if node.element_kind() != super::Kind::Run {
        return Err(Error::new(
            Code::UnsupportedType,
            "RTF run-format mutations target run nodes",
        )
        .with_format(FileFormat::Rtf));
    }
    let Some(paragraph) = node.paragraph_position() else {
        return Err(Error::new(
            Code::UnsupportedType,
            "RTF run-format mutations require a body paragraph",
        )
        .with_format(FileFormat::Rtf));
    };
    let Some(run) = node.run_position() else {
        return Err(Error::new(
            Code::UnsupportedType,
            "RTF run-format mutations require a body run",
        )
        .with_format(FileFormat::Rtf));
    };
    if targets.len() >= maximum {
        return Err(TransactionError::Limit {
            kind: LimitKind::ExpandedTargets,
            observed: targets.len().checked_add(1).unwrap_or(usize::MAX),
            maximum,
        }
        .into());
    }
    targets
        .try_reserve(1)
        .map_err(|source| TransactionError::Allocation { source })?;
    targets.push((paragraph, run));
    Ok(())
}

fn run_text<'a>(props: &'a Props<'_>) -> Result<Option<&'a str>> {
    let mut text = None;
    for (key, value) in props.iter() {
        if key.eq_ignore_ascii_case("text") {
            text = Some(value.as_text().ok_or_else(|| {
                Error::new(Code::InvalidValue, "RTF run text requires a text value")
                    .with_property("text")
                    .with_format(FileFormat::Rtf)
            })?);
        }
    }
    Ok(text)
}

fn run_format_patch(props: &Props<'_>) -> Result<RunFormatPatch> {
    if props.is_empty() {
        return Ok(RunFormatPatch::new());
    }
    let mut patch = RunFormatPatch::new();
    for (key, value) in props.iter() {
        let name = key.as_str();
        if name.eq_ignore_ascii_case("text") {
            continue;
        } else if name.eq_ignore_ascii_case("bold") {
            patch = patch.with_bold(run_bool(value, "bold")?);
        } else if name.eq_ignore_ascii_case("italic") {
            patch = patch.with_italic(run_bool(value, "italic")?);
        } else if name.eq_ignore_ascii_case("underline") {
            patch = patch.with_underline(run_underline(value)?);
        } else if name.eq_ignore_ascii_case("size") {
            patch = patch.with_size(run_size(value)?);
        } else if name.eq_ignore_ascii_case("strike") {
            patch = patch.with_strike(run_bool(value, "strike")?);
        } else if name.eq_ignore_ascii_case("dstrike") {
            patch = patch.with_double_strike(run_bool(value, "dstrike")?);
        } else if name.eq_ignore_ascii_case("vanish") {
            patch = patch.with_hidden(run_bool(value, "vanish")?);
        } else if name.eq_ignore_ascii_case("caps") {
            patch = patch.with_all_caps(run_bool(value, "caps")?);
        } else if name.eq_ignore_ascii_case("smallcaps") {
            patch = patch.with_small_caps(run_bool(value, "smallcaps")?);
        } else if name.eq_ignore_ascii_case("outline") {
            patch = patch.with_outline(run_bool(value, "outline")?);
        } else {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the requested RTF run property is not writable",
            )
            .with_property(name.to_owned())
            .with_format(FileFormat::Rtf));
        }
    }
    Ok(patch)
}

fn run_bool(value: &PropValue<'_>, key: &'static str) -> Result<bool> {
    value
        .with_coerced(ValueType::Bool, |value| {
            let Some(Scalar::Bool(value)) = value else {
                return Err(Error::new(
                    Code::InvalidValue,
                    "RTF run character formatting requires a boolean value",
                )
                .with_property(key)
                .with_format(FileFormat::Rtf));
            };
            Ok(value)
        })
        .map_err(|error| error.with_property(key).with_format(FileFormat::Rtf))
}

fn run_size(value: &PropValue<'_>) -> Result<NonZeroU16> {
    value
        .with_coerced(ValueType::Length(LengthSpec::FONT_SIZE), |value| {
            let Some(Scalar::Length(value)) = value else {
                return Err(Error::new(
                    Code::InvalidValue,
                    "RTF run size requires a font-size length",
                )
                .with_property("size")
                .with_format(FileFormat::Rtf));
            };
            let emu = value.emu();
            if emu <= 0 || emu % 6_350 != 0 {
                return Err(Error::new(
                    Code::InvalidValue,
                    "RTF run size must be a positive whole half-point",
                )
                .with_property("size")
                .with_format(FileFormat::Rtf));
            }
            let half_points = emu / 6_350;
            let half_points = u16::try_from(half_points).map_err(|_| {
                Error::new(
                    Code::InvalidValue,
                    "RTF run size is outside the native half-point range",
                )
                .with_property("size")
                .with_format(FileFormat::Rtf)
            })?;
            NonZeroU16::new(half_points).ok_or_else(|| {
                Error::new(Code::InvalidValue, "RTF run size must be positive")
                    .with_property("size")
                    .with_format(FileFormat::Rtf)
            })
        })
        .map_err(|error| error.with_property("size").with_format(FileFormat::Rtf))
}

fn run_underline(value: &PropValue<'_>) -> Result<UnderlineStyle> {
    value
        .with_coerced(ValueType::Token(&super::spec::RUN_UNDERLINE), |value| {
            let Some(Scalar::Token(value)) = value else {
                return Err(Error::new(
                    Code::InvalidValue,
                    "RTF run underline requires a known underline token",
                )
                .with_property("underline")
                .with_format(FileFormat::Rtf));
            };
            let underline = match value {
                "none" => UnderlineStyle::None,
                "single" => UnderlineStyle::Single,
                "double" => UnderlineStyle::Double,
                "dotted" => UnderlineStyle::Dotted,
                "dashed" => UnderlineStyle::Dashed,
                "dashDot" => UnderlineStyle::DashDot,
                "dashDotDot" => UnderlineStyle::DashDotDot,
                "words" => UnderlineStyle::Words,
                "thick" => UnderlineStyle::Thick,
                "wave" => UnderlineStyle::Wave,
                "hairline" => UnderlineStyle::Hairline,
                "thickDotted" => UnderlineStyle::ThickDotted,
                "thickDashed" => UnderlineStyle::ThickDashed,
                "thickDashDot" => UnderlineStyle::ThickDashDot,
                "thickDashDotDot" => UnderlineStyle::ThickDashDotDot,
                "thickLongDash" => UnderlineStyle::ThickLongDash,
                "longDash" => UnderlineStyle::LongDash,
                "heavyWave" => UnderlineStyle::HeavyWave,
                "doubleWave" => UnderlineStyle::DoubleWave,
                _ => {
                    return Err(Error::new(
                        Code::InvalidValue,
                        "RTF run underline token is not supported",
                    )
                    .with_property("underline")
                    .with_format(FileFormat::Rtf));
                },
            };
            Ok(underline)
        })
        .map_err(|error| {
            error
                .with_property("underline")
                .with_format(FileFormat::Rtf)
        })
}

/// Remove selected ordinary body paragraphs or text runs in one native transaction.
pub(crate) fn remove_with_limits<'p>(
    file: &Rtf,
    target: TargetArg<'p>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.admit_edit()?;
    let mut removed_bytes = 0usize;
    let targets = resolve_mutation_targets(
        file,
        &target,
        limits.max_expanded_targets(),
        Code::UnsupportedCapability,
        |node| {
            let bytes = node.native_run().map_or_else(
                || node.body_paragraph().map_or(0, |paragraph| paragraph.len()),
                |run| run.text().len(),
            );
            removed_bytes = removed_bytes.saturating_add(bytes);
        },
    )?;
    let positions = match targets {
        MutationTargets::Picture(indices) => {
            return destination::remove_pictures(file, indices.as_slice(), limits);
        },
        MutationTargets::Mark(targets) => {
            return marks::remove(file, targets.as_slice(), limits);
        },
        MutationTargets::Destination(_) => {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "RTF table cells, header/footer stories, notes and comments cannot be removed; set their text instead",
            )
            .with_format(FileFormat::Rtf));
        },
        MutationTargets::Run(targets) => {
            // Removing text runs preserves their containing paragraphs and all
            // unselected native formatting. The native empty replacement path
            // owns terminal-empty delimiters and exact inverse/history proofs.
            return publish_run_updates(
                file,
                targets.as_slice(),
                RunFormatPatch::new(),
                Some(""),
                removed_bytes,
                limits,
            );
        },
        MutationTargets::Paragraph(positions) => positions,
    };
    let positions = positions.as_slice();
    let operations = positions.len();
    if operations > limits.max_operations() {
        return Err(TransactionError::Limit {
            kind: LimitKind::Operations,
            observed: operations,
            maximum: limits.max_operations(),
        }
        .into());
    }

    // Native removal retains each selected paragraph's text for stale-source
    // checks and reversible history. Charge that retained payload before the
    // native operation allocates its owned copies.
    let action_bytes = operations
        .checked_mul(size_of::<usize>() + size_of::<String>())
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    let staged_bytes = removed_bytes
        .checked_add(action_bytes)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged_bytes > limits.max_staged_bytes() {
        return Err(staged_limit(staged_bytes, limits.max_staged_bytes()));
    }

    let native_limits = NativeLimits::builder()
        .max_operations(limits.max_operations())
        .map_err(Error::from)?
        .max_replacement_bytes(limits.max_staged_bytes())
        .map_err(Error::from)?
        .build()
        .map_err(Error::from)?;
    let options = EditOptions::new().semantic_limits(native_limits);
    let mut edit = file.0.package.edit_with_options(options);
    if let [position] = positions {
        edit.remove_paragraph(*position).map_err(Error::from)?;
    } else {
        edit.remove_body_paragraphs(positions)
            .map_err(Error::from)?;
    }
    let native = edit.commit().map_err(Error::from)?;

    let changed = native.diagnostics().changed();
    let operation_count = native.diagnostics().operation_count();
    let (package, patch) = native.into_parts();
    let snapshot_bytes = patch_snapshot_bytes(&patch);
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let before = file.clone();
    let after = if changed {
        file.with_package(package)
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok(Commit {
        before,
        after,
        patch,
        operations: operation_count,
        snapshot_bytes,
    })
}

/// Move one ordinary main-story paragraph to a final zero-based position.
///
/// The placement is resolved against the unchanged source snapshot and the
/// resulting final position is passed directly to the native lifecycle edit.
/// Native validation therefore remains responsible for source closure,
/// protection, and exact byte publication.
pub(crate) fn move_to_with_limits<'p, 'q>(
    file: &Rtf,
    path: PathArg<'p>,
    placement: Placement<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.admit_edit()?;
    let nav = Nav::new(file);
    let selected = resolve_exact(file, &path)?;
    let from = paragraph_position(selected)?;
    let count = file.0.package.paragraph_count();
    if count == 0 {
        return Err(
            Error::new(Code::PathNotFound, "the RTF body has no paragraphs")
                .with_format(FileFormat::Rtf),
        );
    }
    let to = move_destination(&nav, from, count, &placement)?;

    bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
    let text_bytes = selected
        .body_paragraph()
        .map_or(0, |paragraph| paragraph.len());
    let staged_bytes = text_bytes
        .saturating_add(size_of::<usize>().saturating_mul(2))
        .saturating_add(size_of::<String>());
    bound(
        LimitKind::StagedBytes,
        staged_bytes,
        limits.max_staged_bytes(),
    )?;
    if limits.max_operations() == 0 {
        return Err(TransactionError::Limit {
            kind: LimitKind::Operations,
            observed: 1,
            maximum: limits.max_operations(),
        }
        .into());
    }

    let native_limits = NativeLimits::builder()
        .max_operations(limits.max_operations())
        .map_err(Error::from)?
        .max_replacement_bytes(limits.max_staged_bytes())
        .map_err(Error::from)?
        .build()
        .map_err(Error::from)?;
    let options = EditOptions::new().semantic_limits(native_limits);
    let mut edit = file.0.package.edit_with_options(options);
    edit.move_paragraph(Position::new(from), Position::new(to))
        .map_err(Error::from)?;
    let native = edit.commit().map_err(Error::from)?;

    let changed = native.diagnostics().changed();
    let operation_count = native.diagnostics().operation_count();
    let (package, patch) = native.into_parts();
    let snapshot_bytes = patch_snapshot_bytes(&patch);
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let before = file.clone();
    let after = if changed {
        file.with_package(package)
    } else {
        // Equal source and destination is a native exact no-op. Preserve the
        // original Arc identity and origin sidecar for the caller.
        file.clone()
    };
    file.check_source()?;
    Ok(Commit {
        before,
        after,
        patch,
        operations: operation_count,
        snapshot_bytes,
    })
}

/// Exchange two ordinary main-story paragraphs in one native lifecycle edit.
///
/// Both paths are resolved against the same immutable source snapshot.  The
/// native editor retains the selected paragraph payloads and owns candidate
/// publication, validation, and reversible history; this adapter
/// only charges the unified target and staging budgets and adopts the checked
/// endpoint.
pub(crate) fn swap_with_limits<'p, 'q>(
    file: &Rtf,
    left: PathArg<'p>,
    right: PathArg<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.admit_edit()?;
    let left = resolve_exact(file, &left)?;
    let right = resolve_exact(file, &right)?;
    let left_bytes = left.body_paragraph().map_or(0, |paragraph| paragraph.len());
    let right_bytes = right
        .body_paragraph()
        .map_or(0, |paragraph| paragraph.len());
    let left = paragraph_position(left)?;
    let right = paragraph_position(right)?;

    let target_count = usize::from(left != right).saturating_add(1);
    bound(
        LimitKind::ExpandedTargets,
        target_count,
        limits.max_expanded_targets(),
    )?;

    // Changed-position staging retains both paragraph payloads. Equal positions
    // retain only the operation descriptors because the native edit skips copies.
    let retained_bytes = if left == right {
        Some(0)
    } else {
        left_bytes.checked_add(right_bytes)
    };
    let payload_bytes = retained_bytes
        .and_then(|bytes| bytes.checked_add(size_of::<usize>().saturating_mul(2)))
        .and_then(|bytes| bytes.checked_add(size_of::<String>().saturating_mul(2)))
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    bound(
        LimitKind::StagedBytes,
        payload_bytes,
        limits.max_staged_bytes(),
    )?;
    bound(LimitKind::Operations, 1, limits.max_operations())?;

    let native_limits = NativeLimits::builder()
        .max_operations(limits.max_operations())
        .map_err(Error::from)?
        .max_replacement_bytes(limits.max_staged_bytes())
        .map_err(Error::from)?
        .build()
        .map_err(Error::from)?;
    let options = EditOptions::new().semantic_limits(native_limits);
    let mut edit = file.0.package.edit_with_options(options);
    edit.swap_paragraphs(Position::new(left), Position::new(right))
        .map_err(Error::from)?;
    let native = edit.commit().map_err(Error::from)?;

    let changed = native.diagnostics().changed();
    let operation_count = native.diagnostics().operation_count();
    let (package, patch) = native.into_parts();
    let snapshot_bytes = patch_snapshot_bytes(&patch);
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let before = file.clone();
    let after = if changed {
        file.with_package(package)
    } else {
        // Equal-target swap is a native exact no-op. Preserve the original
        // package identity and origin sidecar for atomic save and replay.
        file.clone()
    };
    file.check_source()?;
    Ok(Commit {
        before,
        after,
        patch,
        operations: operation_count,
        snapshot_bytes,
    })
}

fn move_destination(
    nav: &Nav<'_>,
    from: usize,
    count: usize,
    placement: &Placement<'_>,
) -> Result<usize> {
    let last = count.checked_sub(1).ok_or_else(|| {
        Error::new(Code::PathNotFound, "the RTF body has no paragraphs")
            .with_format(FileFormat::Rtf)
    })?;
    let destination = match placement {
        Placement::Start => 0,
        Placement::End => last,
        Placement::Position(position) => position.get(),
        Placement::Before(path) => {
            let anchor = paragraph_position(crate::path::resolve::get(nav, path)?)?;
            if from == anchor {
                from
            } else if from < anchor {
                anchor.saturating_sub(1)
            } else {
                anchor
            }
        },
        Placement::After(path) => {
            let anchor = paragraph_position(crate::path::resolve::get(nav, path)?)?;
            if from == anchor {
                from
            } else if from < anchor {
                anchor
            } else {
                anchor.saturating_add(1)
            }
        },
        Placement::Into(path, position) => {
            let parent = crate::path::resolve::get(nav, path)?;
            if parent.element_kind() != super::Kind::Body {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "RTF paragraphs can only be moved within the body",
                )
                .with_format(FileFormat::Rtf));
            }
            position.map_or(last, Position::get)
        },
    };
    if destination >= count {
        return Err(Error::new(
            Code::OutOfRange,
            "the RTF paragraph move position is outside the body",
        )
        .with_format(FileFormat::Rtf));
    }
    Ok(destination)
}

fn paragraph_position(node: Node<'_>) -> Result<usize> {
    if node.element_kind() != super::Kind::Paragraph {
        return Err(Error::new(
            Code::UnsupportedType,
            "RTF mutations target ordinary body paragraphs",
        )
        .with_format(FileFormat::Rtf));
    }
    node.paragraph_position().map_or_else(
        || {
            Err(Error::new(
                Code::UnsupportedType,
                "RTF move destinations must select body paragraphs",
            )
            .with_format(FileFormat::Rtf))
        },
        |position| Ok(position.get()),
    )
}

fn bound(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(TransactionError::Limit {
            kind,
            observed,
            maximum,
        }
        .into());
    }
    Ok(())
}

struct ParagraphProperties<'a> {
    text: Option<&'a str>,
    alignment: Option<Alignment>,
    layout: Option<ParagraphLayoutPatch>,
}

fn paragraph_properties<'a>(props: &'a Props<'_>) -> Result<ParagraphProperties<'a>> {
    for (key, _) in props.iter() {
        let supported = key.eq_ignore_ascii_case("text")
            || key.eq_ignore_ascii_case("align")
            || key.eq_ignore_ascii_case("spaceBefore")
            || key.eq_ignore_ascii_case("spaceAfter")
            || key.eq_ignore_ascii_case("indent")
            || key.eq_ignore_ascii_case("rightIndent")
            || key.eq_ignore_ascii_case("firstLineIndent")
            || key.eq_ignore_ascii_case("keepNext")
            || key.eq_ignore_ascii_case("keepLines")
            || key.eq_ignore_ascii_case("pageBreakBefore");
        if !supported {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "the requested RTF property is not writable",
            )
            .with_property(key.as_str().to_owned())
            .with_format(FileFormat::Rtf));
        }
    }

    let text = props.get("text").map(|value| {
        value.as_text().ok_or_else(|| {
            Error::new(
                Code::InvalidValue,
                "RTF paragraph text requires a text property value",
            )
            .with_property("text")
            .with_format(FileFormat::Rtf)
        })
    });
    let text = text.transpose()?;
    let alignment = props.get("align").map(|value| paragraph_alignment(value));
    let alignment = alignment.transpose()?;

    let mut layout = ParagraphLayoutPatch::new();
    let mut has_layout = false;
    if let Some(value) = props.get("spaceBefore") {
        layout = layout.with_space_before(paragraph_twips(value, "spaceBefore", true)?);
        has_layout = true;
    }
    if let Some(value) = props.get("spaceAfter") {
        layout = layout.with_space_after(paragraph_twips(value, "spaceAfter", true)?);
        has_layout = true;
    }
    if let Some(value) = props.get("indent") {
        layout = layout.with_left_indent(paragraph_twips(value, "indent", false)?);
        has_layout = true;
    }
    if let Some(value) = props.get("rightIndent") {
        layout = layout.with_right_indent(paragraph_twips(value, "rightIndent", false)?);
        has_layout = true;
    }
    if let Some(value) = props.get("firstLineIndent") {
        layout = layout.with_first_line_indent(paragraph_twips(value, "firstLineIndent", false)?);
        has_layout = true;
    }
    if let Some(value) = props.get("keepNext") {
        layout = layout.with_keep_with_next(paragraph_bool(value, "keepNext")?);
        has_layout = true;
    }
    if let Some(value) = props.get("keepLines") {
        layout = layout.with_keep_together(paragraph_bool(value, "keepLines")?);
        has_layout = true;
    }
    if let Some(value) = props.get("pageBreakBefore") {
        layout = layout.with_page_break_before(paragraph_bool(value, "pageBreakBefore")?);
        has_layout = true;
    }

    if text.is_none() && alignment.is_none() && !has_layout {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "RTF paragraph writes require a `text`, `align`, or layout property",
        )
        .with_format(FileFormat::Rtf));
    }
    Ok(ParagraphProperties {
        text,
        alignment,
        layout: has_layout.then_some(layout),
    })
}

fn paragraph_twips(value: &PropValue<'_>, key: &'static str, nonnegative: bool) -> Result<i32> {
    let ty = if nonnegative {
        ValueType::Length(super::spec::SPACING)
    } else {
        ValueType::Length(super::spec::SIGNED_POINTS)
    };
    value
        .with_coerced(ty, |value| {
            let Some(Scalar::Length(value)) = value else {
                return Err(Error::new(
                    Code::InvalidValue,
                    "RTF paragraph layout requires a length value",
                )
                .with_property(key)
                .with_format(FileFormat::Rtf));
            };
            let Some(twips) = value.checked_twips() else {
                return Err(Error::new(
                    Code::InvalidValue,
                    "RTF paragraph layout length is outside the native twip range",
                )
                .with_property(key)
                .with_format(FileFormat::Rtf));
            };
            Ok(twips)
        })
        .map_err(|error| error.with_property(key).with_format(FileFormat::Rtf))
}

fn paragraph_bool(value: &PropValue<'_>, key: &'static str) -> Result<bool> {
    value
        .with_coerced(ValueType::Bool, |value| {
            let Some(Scalar::Bool(value)) = value else {
                return Err(Error::new(
                    Code::InvalidValue,
                    "RTF paragraph layout requires a boolean value",
                )
                .with_property(key)
                .with_format(FileFormat::Rtf));
            };
            Ok(value)
        })
        .map_err(|error| error.with_property(key).with_format(FileFormat::Rtf))
}

fn paragraph_alignment(value: &PropValue<'_>) -> Result<Alignment> {
    value
        .with_coerced(ValueType::Token(&super::spec::ALIGNMENT), |value| {
            let Some(Scalar::Token(value)) = value else {
                return Err(Error::new(
                    Code::InvalidValue,
                    "RTF paragraph alignment requires a token value",
                )
                .with_property("align")
                .with_format(FileFormat::Rtf));
            };
            match value {
                "left" => Ok(Alignment::Left),
                "right" => Ok(Alignment::Right),
                "center" => Ok(Alignment::Center),
                "justify" => Ok(Alignment::Justify),
                _ => Err(Error::new(
                    Code::InvalidValue,
                    "RTF paragraph alignment has an unknown token",
                )
                .with_property("align")
                .with_format(FileFormat::Rtf)),
            }
        })
        .map_err(|error| error.with_property("align").with_format(FileFormat::Rtf))
}

fn append_paragraph(positions: &mut Vec<Position>, node: Node<'_>, maximum: usize) -> Result<()> {
    if node.element_kind() != super::Kind::Paragraph {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "RTF mutations target paragraph nodes",
        )
        .with_format(FileFormat::Rtf));
    }
    let Some(position) = node.paragraph_position() else {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "RTF mutations target paragraph nodes",
        )
        .with_format(FileFormat::Rtf));
    };
    if positions.len() >= maximum {
        return Err(TransactionError::Limit {
            kind: LimitKind::ExpandedTargets,
            observed: positions.len().checked_add(1).unwrap_or(usize::MAX),
            maximum,
        }
        .into());
    }
    positions
        .try_reserve(1)
        .map_err(|source| TransactionError::Allocation { source })?;
    positions.push(position);
    Ok(())
}

fn staged_limit(observed: usize, maximum: usize) -> Error {
    TransactionError::Limit {
        kind: LimitKind::StagedBytes,
        observed,
        maximum,
    }
    .into()
}

fn snapshot_limit(observed: usize, maximum: usize) -> Error {
    TransactionError::Limit {
        kind: LimitKind::SnapshotBytes,
        observed,
        maximum,
    }
    .into()
}

/// Native history weight includes the exact endpoints and the native semantic
/// change payload.  Keep this adapter's charge tied to that native accounting
/// so a long source paragraph replaced by a short value remains bounded.
fn patch_snapshot_bytes(patch: &Patch) -> usize {
    usize::try_from(patch.history_weight()).unwrap_or(usize::MAX)
}
