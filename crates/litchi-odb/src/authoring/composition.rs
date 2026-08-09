//! Deterministic ODB sub-edit composition and three-way planning.

use std::{collections::BTreeSet, ops::Range, sync::Arc};

use litchi_core::{Error, Result};

use super::{Change, ChangeKind, Patch};
use crate::Database;

pub use litchi_core::{CompositionLimits, MergeChoice};

/// Exact immutable source lineage used by ODB composition.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Lineage(Arc<[u8]>);

/// One independently prepared ODB patch with canonical semantic effects.
pub type PreparedEdit = litchi_core::SubEdit<Lineage, Patch>;

/// Deterministically ordered, provably disjoint ODB sub-edits.
pub type JoinedEdits = litchi_core::JoinedSubEdits<Lineage, Patch>;

/// Recoverable deterministic join refusal.
pub type JoinError = litchi_core::SubEditJoinError<Lineage, Patch>;

/// Non-applying three-way ODB merge plan.
pub type MergePlan = litchi_core::ThreeWayMergePlan<Lineage, Patch>;

/// Recoverable three-way planning failure.
pub type MergePlanError = litchi_core::ThreeWayMergeError<Lineage, Patch>;

impl Patch {
    /// Wraps this patch as independently prepared work with exact semantic
    /// read/write effects.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid identifier or exceeded effect bound.
    pub fn prepare(
        &self,
        identifier: impl Into<String>,
        limits: CompositionLimits,
    ) -> std::result::Result<PreparedEdit, litchi_core::CompositionError> {
        let writes = effect_keys(&self.changes);
        litchi_core::SubEdit::new(
            lineage(&self.source),
            limits,
            identifier,
            Vec::<String>::new(),
            writes,
            self.clone(),
        )
    }

    /// Materializes joined disjoint work against its exact common source.
    ///
    /// Joined sub-edits are applied in stable identifier order. Publication
    /// performs a full ODB reopen and semantic catalog readback.
    ///
    /// # Errors
    ///
    /// Returns an error if exact lineage changed, physical splice ranges
    /// unexpectedly overlap, rebuilding fails, or full readback fails.
    pub fn compose(joined: JoinedEdits) -> Result<Self> {
        let source = joined
            .sub_edits()
            .next()
            .map(|edit| edit.payload().source.clone())
            .ok_or_else(|| Error::InvalidFormat("ODB composition has no sub-edits".to_string()))?;
        if joined.lineage() != &lineage(&source) {
            return invalid("ODB composition source lineage changed");
        }
        let base = source.content_xml();
        let mut hunks = Vec::new();
        let mut changes = Vec::new();
        for (ordinal, edit) in joined.sub_edits().enumerate() {
            let patch = edit.payload();
            if patch.source.as_bytes() != source.as_bytes() {
                return invalid("ODB composition patch source is stale");
            }
            if has_non_content_changes(&patch.source, &patch.target)? {
                return Err(Error::Unsupported(
                    "ODB composition refuses patches that change package members outside content.xml"
                        .to_string(),
                ));
            }
            if let Some(hunk) = difference(base, patch.target.content_xml(), ordinal) {
                hunks.push(hunk);
            }
            changes
                .try_reserve(patch.changes.len())
                .map_err(|source| Error::Allocation {
                    resource: "ODB composed semantic changes",
                    source,
                })?;
            changes.extend(patch.changes.iter().cloned());
        }
        hunks = merge_shared_owners(hunks)?;
        validate_disjoint_hunks(&hunks)?;
        hunks.sort_by(|left, right| {
            right
                .range
                .start
                .cmp(&left.range.start)
                .then_with(|| right.ordinal.cmp(&left.ordinal))
        });
        let mut content = base.to_owned();
        for hunk in hunks {
            content.replace_range(hunk.range, &hunk.replacement);
        }
        if content == base {
            return Ok(Self {
                source: source.clone(),
                target: source,
                changes: Vec::new(),
                legacy_query: None,
            });
        }
        crate::codec::validate(&content)?;
        let target = Database {
            package: source.package.rebuild_with_content(&content)?,
        };
        target.catalog()?;
        Ok(Self {
            source,
            target,
            changes,
            legacy_query: None,
        })
    }
}

fn has_non_content_changes(source: &Database, target: &Database) -> Result<bool> {
    let source_files = source.files()?.into_iter().collect::<BTreeSet<_>>();
    let target_files = target.files()?.into_iter().collect::<BTreeSet<_>>();
    let paths = source_files.union(&target_files);
    for path in paths {
        if path == "content.xml" {
            continue;
        }
        if !source_files.contains(path) || !target_files.contains(path) {
            return Ok(true);
        }
        if source.package.file(path)? != target.package.file(path)? {
            return Ok(true);
        }
    }
    Ok(false)
}

struct Hunk {
    range: Range<usize>,
    replacement: String,
    ordinal: usize,
}

fn lineage(database: &Database) -> Lineage {
    Lineage(Arc::from(database.as_bytes()))
}

fn effect_keys(changes: &[Change]) -> Vec<String> {
    changes
        .iter()
        .map(|change| match change.kind() {
            ChangeKind::Table => format!("table/{}", change.target()),
            ChangeKind::Column | ChangeKind::Key | ChangeKind::Index => {
                let table = semantic_owner(change.target()).unwrap_or(change.target());
                format!("table/{table}")
            },
            ChangeKind::Query => format!("query/{}", change.target()),
            ChangeKind::Connection => "connection".to_string(),
            ChangeKind::Component => format!("component/{}", change.target()),
            ChangeKind::ProducerExtension => {
                format!("producer-extension/{}", change.target())
            },
        })
        .collect()
}

fn semantic_owner(target: &str) -> Option<&str> {
    let (length, value) = target.split_once(':')?;
    let length = length.parse::<usize>().ok()?;
    value.get(..length)
}

fn difference(base: &str, candidate: &str, ordinal: usize) -> Option<Hunk> {
    if base == candidate {
        return None;
    }
    let common = base
        .as_bytes()
        .iter()
        .zip(candidate.as_bytes())
        .take_while(|(left, right)| left == right)
        .count();
    let mut prefix = common;
    while !base.is_char_boundary(prefix) || !candidate.is_char_boundary(prefix) {
        prefix = prefix.saturating_sub(1);
    }
    let remaining_base = base.len() - prefix;
    let remaining_candidate = candidate.len() - prefix;
    let suffix = base.as_bytes()[prefix..]
        .iter()
        .rev()
        .zip(candidate.as_bytes()[prefix..].iter().rev())
        .take_while(|(left, right)| left == right)
        .count()
        .min(remaining_base)
        .min(remaining_candidate);
    let mut base_end = base.len() - suffix;
    let mut candidate_end = candidate.len() - suffix;
    while !base.is_char_boundary(base_end) || !candidate.is_char_boundary(candidate_end) {
        base_end += 1;
        candidate_end += 1;
    }
    Some(Hunk {
        range: prefix..base_end,
        replacement: candidate[prefix..candidate_end].to_owned(),
        ordinal,
    })
}

fn validate_disjoint_hunks(hunks: &[Hunk]) -> Result<()> {
    for (index, left) in hunks.iter().enumerate() {
        for right in &hunks[index + 1..] {
            let separate =
                left.range.end <= right.range.start || right.range.end <= left.range.start;
            let same_insertion = left.range.is_empty()
                && right.range.is_empty()
                && left.range.start == right.range.start;
            if !separate && !same_insertion {
                return invalid("ODB composed XML splice ranges overlap");
            }
        }
    }
    Ok(())
}

fn merge_shared_owners(hunks: Vec<Hunk>) -> Result<Vec<Hunk>> {
    let mut merged = Vec::<Hunk>::new();
    for hunk in hunks {
        if let Some(existing) = merged.iter_mut().find(|item| item.range == hunk.range) {
            if hunk.range.is_empty() {
                existing.replacement.push_str(&hunk.replacement);
                continue;
            }
            let (left_open, left_inner, left_close) = split_owner(&existing.replacement)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODB shared physical owner cannot be composed".to_string())
                })?;
            let (right_open, right_inner, right_close) = split_owner(&hunk.replacement)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODB shared physical owner cannot be composed".to_string())
                })?;
            if left_open != right_open || left_close != right_close {
                return invalid("ODB shared physical owner replacements conflict");
            }
            existing.replacement = format!("{left_open}{left_inner}{right_inner}{left_close}");
        } else {
            merged.push(hunk);
        }
    }
    Ok(merged)
}

fn split_owner(value: &str) -> Option<(&str, &str, &str)> {
    if value.starts_with("</") || value.ends_with("/>") {
        return None;
    }
    let close_start = value.rfind("</")?;
    let open_end = if value.starts_with('<') {
        value.find('>')?.checked_add(1)?
    } else {
        value.find('<')?
    };
    if close_start < open_end {
        return None;
    }
    Some((
        &value[..open_end],
        &value[open_end..close_start],
        &value[close_start..],
    ))
}

fn invalid<T>(message: &str) -> Result<T> {
    Err(Error::InvalidFormat(message.to_owned()))
}
