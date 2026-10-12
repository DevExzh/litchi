//! Unified paragraph-text writes over the native immutable RTF editor.
//!
//! Target resolution is performed against one retained `Rtf` snapshot.  The
//! complete selection is then staged as one native paragraph batch, which
//! keeps selector fan-out atomic and preserves the native source/options
//! policy through publication.

use litchi_core::FileFormat;
use litchi_rtf::edit::{
    EditOptions, ParagraphTextReplacement, Patch, SemanticLimits as NativeLimits,
};

use super::{Nav, Node, Rtf};
use crate::error::{Code, Error, Result};
use crate::props::{IntoProps, Props};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// The native proof and immutable endpoints of one RTF write.
pub(crate) struct Commit {
    pub(crate) before: Rtf,
    pub(crate) after: Rtf,
    pub(crate) patch: Patch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

/// Admit one paragraph-text assignment and publish all selected paragraphs in
/// one native edit.
pub(crate) fn set_with_limits<'p, 'v>(
    file: &Rtf,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.admit_edit()?;
    let props = Props::from_props(props)?;
    let text = paragraph_text(&props)?;
    let positions = resolve_paragraph_targets(file, &target, limits.max_expanded_targets())?;

    let operations = positions.len();
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
    let replacement_bytes = operations
        .checked_mul(text.len())
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    let action_bytes = operations
        .checked_mul(size_of::<ParagraphTextReplacement>())
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    let staged_bytes = replacement_bytes
        .checked_add(action_bytes)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if staged_bytes > limits.max_staged_bytes() {
        return Err(staged_limit(staged_bytes, limits.max_staged_bytes()));
    }

    let mut replacements = Vec::new();
    replacements
        .try_reserve_exact(operations)
        .map_err(|source| TransactionError::Allocation { source })?;
    for position in positions {
        let mut replacement = String::new();
        replacement
            .try_reserve_exact(text.len())
            .map_err(|source| TransactionError::Allocation { source })?;
        replacement.push_str(text);
        replacements.push(ParagraphTextReplacement::new(position, replacement));
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
    edit.replace_body_paragraph_texts(&replacements)
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

/// Remove every selected ordinary body paragraph in one native transaction.
pub(crate) fn remove_with_limits<'p>(
    file: &Rtf,
    target: TargetArg<'p>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.admit_edit()?;
    let positions = resolve_paragraph_targets(file, &target, limits.max_expanded_targets())?;
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
    let mut removed_bytes = 0usize;
    let mut selected = positions.iter().copied();
    let mut next = selected.next();
    for (position, paragraph) in file.0.package.body().paragraphs().enumerate() {
        if next == Some(position) {
            removed_bytes = removed_bytes
                .checked_add(paragraph.len())
                .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
            next = selected.next();
        }
        if next.is_none() {
            break;
        }
    }
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
    edit.remove_body_paragraphs(&positions)
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
    Ok(Commit {
        before,
        after,
        patch,
        operations: operation_count,
        snapshot_bytes,
    })
}

fn paragraph_text<'a>(props: &'a Props<'_>) -> Result<&'a str> {
    if props.len() != 1 {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "RTF paragraph writes accept exactly one `text` property",
        )
        .with_format(FileFormat::Rtf));
    }
    let Some((key, value)) = props.iter().next() else {
        unreachable!("the length check above admits one property")
    };
    if !key.eq_ignore_ascii_case("text") {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "the requested RTF property is not writable",
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

fn resolve_paragraph_targets(
    file: &Rtf,
    target: &TargetArg<'_>,
    maximum: usize,
) -> Result<Vec<usize>> {
    let mut positions = Vec::new();
    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&Nav::new(file), path)?;
        append_paragraph(&mut positions, node, maximum)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query.clone())?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the RTF mutation selector contains an unsupported predicate",
            )
            .with_format(FileFormat::Rtf));
        }
        for node in walk {
            append_paragraph(&mut positions, node?, maximum)?;
        }
    } else {
        return Err(
            Error::new(Code::InvalidArgument, "the RTF mutation target is empty")
                .with_format(FileFormat::Rtf),
        );
    }
    file.check_source()?;
    if positions.is_empty() {
        return Err(Error::new(
            Code::ZeroMatches,
            "the RTF mutation selector matched no paragraph",
        )
        .with_format(FileFormat::Rtf));
    }
    // Query unions are already document ordered, but normalizing here keeps
    // the native batch contract true even if a future walk source changes its
    // duplicate policy.
    positions.sort_unstable();
    positions.dedup();
    Ok(positions)
}

fn append_paragraph(positions: &mut Vec<usize>, node: Node<'_>, maximum: usize) -> Result<()> {
    let Some(position) = node.paragraph_position() else {
        return Err(Error::new(
            Code::UnsupportedCapability,
            "RTF text writes target paragraph nodes",
        )
        .with_format(FileFormat::Rtf));
    };
    if positions.len() >= maximum {
        return Err(TransactionError::Limit {
            kind: LimitKind::ExpandedTargets,
            observed: positions.len().saturating_add(1),
            maximum,
        }
        .into());
    }
    positions
        .try_reserve(1)
        .map_err(|source| TransactionError::Allocation { source })?;
    positions.push(position.get());
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
