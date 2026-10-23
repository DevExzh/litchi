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
use std::sync::{Arc, OnceLock};

pub(crate) enum Proof {
    Source(litchi_pptx::source::Patch),
    Eager(litchi_pptx::opened::Patch),
    /// A document-metadata patch (record 0253).
    Metadata(litchi_pptx::metadata::Patch),
    /// Proofs from two native edits that make up one facade operation.
    ///
    /// A click-action setter can first edit the ordinary shape properties and
    /// then edit the hyperlink.  Both native patches remain live until the
    /// combined commit is consumed; dropping the first one here used to make
    /// its retained snapshot charge disappear from the facade budget.
    Combined(Box<[Self]>),
}
impl Proof {
    pub(crate) fn retained_bytes(&self) -> usize {
        match self {
            Self::Source(patch) => patch.retained_bytes(),
            Self::Eager(patch) => patch.retained_bytes(),
            Self::Metadata(patch) => patch.retained_bytes(),
            Self::Combined(proofs) => proofs.iter().fold(0, |bytes, proof| {
                bytes.saturating_add(proof.retained_bytes())
            }),
        }
    }

    pub(crate) fn is_noop(&self) -> bool {
        match self {
            Self::Source(patch) => patch.is_noop(),
            Self::Eager(patch) => patch.is_noop(),
            Self::Metadata(patch) => patch.is_noop(),
            Self::Combined(proofs) => proofs.iter().all(Self::is_noop),
        }
    }
}
pub(crate) struct Commit {
    pub(crate) after: Pptx,
    pub(crate) source: Editable,
    pub(crate) target: Editable,
    pub(crate) proof: Proof,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
    /// Retained native patch size is cached per commit and must not be
    /// recomputed by each facade limit check. In particular, direct unsigned
    /// writes should not repeatedly walk the complete patch just to read the
    /// same charge.
    pub(crate) retained_cache: OnceLock<usize>,
}
impl Commit {
    /// Build a commit with its native retained-snapshot charge accounted for.
    /// Native routes retain a proof until the facade step is consumed, so do
    /// the exact proof walk once at this boundary instead of exposing a zero
    /// charge while leaving the accessor to discover it later.
    pub(crate) fn new(
        after: Pptx,
        source: Editable,
        target: Editable,
        proof: Proof,
        operations: usize,
    ) -> Self {
        let snapshot_bytes = proof.retained_bytes();
        Self {
            after,
            source,
            target,
            proof,
            operations,
            snapshot_bytes,
            retained_cache: OnceLock::from(snapshot_bytes),
        }
    }

    pub(crate) fn into_file(self) -> Pptx {
        self.after
    }

    /// Return the already-accounted snapshot charge. The native proof walk is
    /// cached per commit so signed and unsigned writes pay it at most once.
    pub(crate) fn snapshot_bytes(&self) -> usize {
        // Older native families initialize this field to zero while they
        // construct the proof.  The proof is the authoritative retained
        // charge, so expose its cached measurement to the transaction
        // wrapper as well as to the setters that call
        // `measure_snapshot_bytes` directly.  This keeps structural,
        // ordering, comments, media, and other native routes from bypassing
        // the facade snapshot budget merely because they do not have a
        // format-specific preflight hook.
        let measured = *self
            .retained_cache
            .get_or_init(|| self.proof.retained_bytes());
        self.snapshot_bytes.max(measured)
    }

    /// Measure the native proof for a hard snapshot check, caching the result
    /// for the transaction wrapper and any later canonical no-op accounting.
    pub(crate) fn measure_snapshot_bytes(&mut self) -> usize {
        let proof = &self.proof;
        let measured = *self.retained_cache.get_or_init(|| proof.retained_bytes());
        self.snapshot_bytes = self.snapshot_bytes.max(measured);
        self.snapshot_bytes
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

/// The facade's semantic snapshot ceiling is also the native patch ceiling.
/// The native option builders validate the finite, non-zero bound while the
/// native editor performs the exact encoded-patch preflight.
pub(super) fn native_snapshot_budget(limits: &SemanticLimits) -> usize {
    limits.max_snapshot_bytes()
}

pub(super) fn opened_edit_options(max_patch_bytes: usize) -> litchi_pptx::opened::EditOptions {
    litchi_pptx::opened::EditOptions::default()
        .limits(litchi_pptx::opened::Limits::default().with_max_patch_bytes(max_patch_bytes))
}

pub(super) fn source_edit_options(max_patch_bytes: usize) -> litchi_pptx::source::EditOptions {
    litchi_pptx::source::EditOptions::default().max_patch_bytes(max_patch_bytes)
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
    /// Whether the payload assigns a click action (`link`, `tooltip`), which
    /// the hyperlink verbs write (record 0325).
    pub(super) click: bool,
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
        click: false,
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
                        ShapeSetter::Text | ShapeSetter::Format(_) | ShapeSetter::Click(_) => {
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
                ShapeSetter::Click(_) => values.click = true,
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
        if self.signature() == crate::format::signature::Signature::Signed {
            self.check_source_for_replay()?;
        } else {
            self.check_source()?;
        }
        let result = (|| self.set_inner(target, Props::from_props(props)?, limits))()
            .map_err(|error| error.with_format(FileFormat::Pptx));
        result.and_then(|commit| {
            // Every native proof retains source/target endpoints, including
            // unsigned writes.  Charge it once here so the facade's snapshot
            // limit cannot be bypassed merely by omitting a signature.
            let mut commit = commit;
            let snapshot_bytes = commit.measure_snapshot_bytes();
            if snapshot_bytes > limits.max_snapshot_bytes() {
                return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
            }
            if self.signature() == crate::format::signature::Signature::Signed {
                self.check_source_for_replay()?;
                if commit.after.same_snapshot(self) {
                    if commit.proof.is_noop() {
                        // The staged private edit may have retained more of
                        // the source than the canonical identity probe.  A
                        // direct signed no-op still owes the caller the
                        // peak charge of both phases, just like a coalesced
                        // run does.
                        let staged_snapshot_bytes = commit.snapshot_bytes();
                        // The canonical replacement is a separate peak
                        // phase. Release the staged target/proof before
                        // retaining the canonical empty patch.
                        let operations = commit.operations;
                        drop(commit);
                        let mut commit = canonical_signed_noop(
                            self,
                            operations,
                            native_snapshot_budget(limits),
                        )?;
                        commit.snapshot_bytes = commit.snapshot_bytes.max(staged_snapshot_bytes);
                        let snapshot_bytes = commit.snapshot_bytes();
                        if snapshot_bytes > limits.max_snapshot_bytes() {
                            return Err(snapshot_limit(
                                snapshot_bytes,
                                limits.max_snapshot_bytes(),
                            ));
                        }
                        Ok(commit)
                    } else {
                        Ok(commit)
                    }
                } else {
                    Err(Error::signed(FileFormat::Pptx))
                }
            } else {
                self.check_source()?;
                Ok(commit)
            }
        })
    }
    fn set_inner(
        &self,
        target: TargetArg<'_>,
        props: Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        if let Some(commit) = self.try_set_metadata(&target, &props, limits)? {
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
        // Modern comment threads and replies (record 0346).
        if let Some(commit) = self.try_set_modern_comment(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_theme(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_master(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_column(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_row(&target, &props, limits)? {
            return Ok(commit);
        }
        // Table-level sets (record 0342).
        if let Some(commit) = self.try_set_table(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_hyperlink(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_click(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_picture(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_chart(&target, &props, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_set_smartart(&target, &props, limits)? {
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
            let walk = crate::query::run(nav, query)?;
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
        let source = if self.signature() == crate::format::signature::Signature::Signed {
            self.editable_owner_for_write()?
        } else {
            self.editable_owner()?
        };
        let mut operations = 0;
        let native_limit = native_snapshot_budget(limits);
        let (target, proof, changed) = match &source {
            Editable::Source(package) => {
                let mut edit = package
                    .edit_with_options(source_edit_options(native_limit))
                    .map_err(native_error)?;
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
                let mut edit = package
                    .edit_with_options(opened_edit_options(native_limit))
                    .map_err(native_error)?;
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
        // One facade setter remains one authored operation even when the
        // native property writer proves that its final value is unchanged.
        operations = operations.max(1);
        // Defer the retained-patch walk to the cached accessor.  This keeps
        // construction of an unsigned direct commit from doing an eager
        // second scan when the transaction only needs the charge once.
        let snapshot_bytes = 0;
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
            snapshot_bytes,
            retained_cache: OnceLock::new(),
        })
    }
}

fn canonical_signed_noop(base: &Pptx, operations: usize, max_patch_bytes: usize) -> Result<Commit> {
    let source = base.original_owner_for_noop()?;
    match &source {
        Editable::Source(package) => {
            let (target, patch) = package
                .edit_with_options(source_edit_options(max_patch_bytes))
                .map_err(native_error)?
                .commit()
                .map_err(native_error)?
                .into_parts();
            let snapshot_bytes = patch.retained_bytes();
            Ok(Commit {
                after: base.clone(),
                source,
                target: Editable::Source(target),
                proof: Proof::Source(patch),
                operations,
                snapshot_bytes,
                retained_cache: OnceLock::from(snapshot_bytes),
            })
        },
        Editable::Eager(package) => {
            let (target, patch) = package
                .edit_with_options(opened_edit_options(max_patch_bytes))
                .map_err(native_error)?
                .commit()
                .map_err(native_error)?
                .into_parts();
            let snapshot_bytes = patch.retained_bytes();
            Ok(Commit {
                after: base.clone(),
                source,
                target: Editable::Eager(Arc::new(target)),
                proof: Proof::Eager(patch),
                operations,
                snapshot_bytes,
                retained_cache: OnceLock::from(snapshot_bytes),
            })
        },
    }
}

#[derive(Clone, Copy)]
struct Target {
    slot: bool,
    slide: usize,
    position: usize,
    bounds: Option<(i64, i64, i64, i64)>,
}

/// The text of a new notes page (record 0346): its `text`, empty when the
/// call assigns none.
pub(in crate::format::pptx) fn notes_add_text(
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<String> {
    if props.iter().next().is_none() {
        return Ok(String::new());
    }
    notes::admitted_text(props, "text", limits)
}

/// Slide-grouped body edits: rich formatting for every target, and the text
/// of targets whose body has no run slot.
fn body_edits<'t>(
    targets: &[Target],
    text: Option<&'t str>,
    format: &'t text::TextValues,
) -> Result<Vec<(usize, Vec<litchi_pptx::opened::TextAssignment<'t>>)>> {
    let formatting = format.properties();
    let filled = text.map(|text| formatting.with_text(text));
    let mut groups: Vec<(usize, Vec<litchi_pptx::opened::TextAssignment<'t>>)> = Vec::new();
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
        edits.push(litchi_pptx::opened::TextAssignment::new(
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

fn snapshot_limit(observed: usize, maximum: usize) -> Error {
    TxError::Limit {
        kind: LimitKind::SnapshotBytes,
        observed,
        maximum,
    }
    .into()
}

fn allocation(source: std::collections::TryReserveError) -> Error {
    TxError::Allocation { source }.into()
}
fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Pptx)
}

pub(super) fn native_error(error: litchi_pptx::Error) -> Error {
    Error::from(error).with_format(FileFormat::Pptx)
}

pub(super) fn validate_text(value: &str) -> Result<()> {
    if !value.chars().all(|ch| matches!(ch, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')) {
        return Err(error(Code::InvalidValue, "text contains a character that XML cannot represent").with_property("text"));
    }
    Ok(())
}
