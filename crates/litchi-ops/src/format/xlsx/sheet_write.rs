//! Immutable XLSX workbook-tab property writes.
//!
//! This adapter is deliberately separate from cell writes.  A target is
//! resolved against the retained source snapshot first; only a target made
//! entirely of workbook tabs reaches the native tab editor.  Cell and range
//! targets return `Ok(None)` so the cell adapter can admit its own registry
//! without resolving the target twice at the dispatch boundary.

use std::collections::{HashSet, TryReserveError};

use litchi_core::FileFormat;
use litchi_xlsx::{self as xlsx, Name};

use super::node::Node;
use super::spec::{Kind, SHEET_WRITE_TABLE, SheetSetter};
use super::write::{Commit, edit_with_snapshot_limit};
use super::{Nav, Xlsx};
use crate::error::{Code, Error, Result};
use crate::props::Props;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

/// Validated sheet assignments shared by sheet set and structural admission.
///
/// `None` means that the caller did not assign that property.  The native
/// editor remains the authority for active-tab and last-visible-tab policy.
#[derive(Debug, Default)]
pub(crate) struct SheetAssignments {
    pub(crate) name: Option<Name>,
    pub(crate) hidden: Option<bool>,
}

const WRITE_ENTRY_OVERHEAD: usize = 64;

/// Try the XLSX sheet-property route.
///
/// `Ok(None)` is reserved for a target that is entirely cell/range-shaped;
/// the caller then invokes the cell adapter.  Any other target is resolved
/// here so unsupported mixed selectors cannot accidentally fall through to a
/// different mutation family.
pub(crate) fn try_set_sheet(
    file: &Xlsx,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Commit>> {
    // The outer writer fences every route. Avoid redundant source checks on
    // direct cell/range writes that cannot enter the sheet editor.
    if direct_cell_or_range_path(target) {
        return Ok(None);
    }
    file.check_source_for_replay()?;
    let result = (|| {
        let Some(positions) = resolve_sheet_targets(file, target, limits)? else {
            return Ok(None);
        };
        let assignments = admit_sheet_properties(Operation::Set, props)?;
        let charge = per_target_charge(&assignments, limits)?;
        check_target_count(positions.len(), charge, limits)?;
        if assignments.name.is_some() && positions.len() > 1 {
            return Err(Error::new(
                Code::InvalidArgument,
                "renaming a selector with multiple sheet matches is ambiguous",
            )
            .with_format(FileFormat::Xlsx)
            .with_property("name"));
        }

        // The native edit validates signed/encrypted/retained-source policy
        // and keeps rename/visibility changes in one atomic workbook edit.
        let names = file.worksheet_names()?;
        let source = file.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        for position in positions {
            let Some(mut tab) = edit.tab(position).map_err(native_error)? else {
                return Err(Error::new(
                    Code::OutOfRange,
                    "the XLSX sheet position is out of range",
                )
                .with_format(FileFormat::Xlsx));
            };
            if let Some(name) = assignments.name.as_ref()
                && names
                    .get(position)
                    .is_none_or(|current| current.as_str() != name.as_str())
            {
                tab.rename(name.clone()).map_err(native_error)?;
            }
            if let Some(hidden) = assignments.hidden {
                if hidden {
                    tab.hide();
                } else {
                    tab.show();
                }
            }
        }
        let native = edit.commit().map_err(native_error)?;
        Commit::from_native_with_limit(file, source, native, limits.max_snapshot_bytes()).map(Some)
    })();
    file.check_source_for_replay()?;
    result
}

/// Inherent forwarding method for the format root.  Keeping the method here
/// means the dispatch layer only needs the stable sheet route name.
impl Xlsx {
    pub(crate) fn try_set_sheet(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let mut filter = None;
        let mut rest = Props::with_capacity(props.len());
        for (key, value) in props.iter() {
            match super::spec::SHEET_FILTER_TABLE.lookup(key.as_str()) {
                Some((_, entry)) if !direct_cell_or_range_path(target) => {
                    let range = entry.prepare(Operation::Set, value, |_, scalar| {
                        Ok(match scalar {
                            Some(Scalar::Text(text)) => Some(text.to_owned()),
                            _ => None,
                        })
                    })?;
                    if filter.is_some() {
                        return Err(duplicate_property("autoFilter"));
                    }
                    filter = Some(range.filter(|range| {
                        !(range.is_empty()
                            || range.eq_ignore_ascii_case("none")
                            || range.eq_ignore_ascii_case("false"))
                    }));
                },
                _ => rest.push(key.reborrow(), value.reborrow()),
            }
        }
        let Some(filter) = filter else {
            return try_set_sheet(self, target, props, limits);
        };
        // The tab properties first, then the auto-filter on their result,
        // published as one commit of the original snapshot (record 0320).
        let Some(positions) = resolve_sheet_targets(self, target, limits)? else {
            return Ok(None);
        };
        let (middle, source, tab_proof_bytes) = if rest.is_empty() {
            (
                self.clone(),
                self.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?,
                0,
            )
        } else {
            match try_set_sheet(self, target, &rest, limits)? {
                Some(commit) => {
                    let (middle, source, _tab_target, tab_patch) = commit.into_parts();
                    // The workbook patch retains both of its complete
                    // workbook endpoints, so its retained-byte value is the
                    // whole first proof charge. Do not add the endpoints a
                    // second time below.
                    (middle, source, tab_patch.retained_bytes())
                },
                None => return Ok(None),
            }
        };
        let base = middle.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = base.edit_auto_filter_with_options(
            super::feature_write::part_edit_options(&base, limits.max_snapshot_bytes()),
        )?;
        for position in positions {
            let value = filter
                .as_deref()
                .map(|range| {
                    xlsx::auto_filter::Range::new(range.trim().replace('$', ""))
                        .map(|range| xlsx::auto_filter::Definition::new(Some(range)))
                        .map_err(|error| Error::from(error).with_property("autoFilter"))
                })
                .transpose()?;
            edit.set_auto_filter(position, value)
                .map_err(|error| native_property_error(error, "autoFilter"))?;
        }
        let filter_commit = edit.commit()?;
        let filter_patch_bytes = filter_commit.patch().retained_bytes();
        let target = filter_commit.into_package();
        let source_bytes = source.retained_bytes();
        let target_bytes = target.retained_bytes();
        let final_endpoints = source_bytes.saturating_add(target_bytes);
        let filter_proof = if filter_patch_bytes == 0 {
            0
        } else {
            base.retained_bytes()
                .saturating_add(target_bytes)
                .saturating_add(filter_patch_bytes)
        };
        // The two native captures are sequential and their proofs are
        // discarded before the unified commit is published. Retain the peak
        // complete proof charge, not the sum of both payloads; each proof
        // keeps its own repeated source/target endpoints. `from_metadata`
        // already charges the final source and target, so pass only the peak
        // excess as its discarded native payload.
        let peak_proof = final_endpoints.max(tab_proof_bytes).max(filter_proof);
        let native_patch_bytes = peak_proof.saturating_sub(final_endpoints);
        let changed = !source.same_content(&target).map_err(native_error)?;
        let after = if changed {
            middle.adopt_workbook(target.clone())
        } else {
            self.clone()
        };
        Ok(Some(Commit::from_metadata(
            self,
            after,
            source,
            target,
            changed,
            1,
            native_patch_bytes,
            limits.max_snapshot_bytes(),
        )?))
    }
}

/// Resolve a target into unique sheet positions in source document order.
///
/// A direct cell/range path is recognized before `Nav` construction, because
/// resolving a range pins a worksheet grid.  Selector targets are streamed;
/// cells/ranges return `None` only when every match belongs to that family.
pub(super) fn resolve_sheet_targets(
    file: &Xlsx,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Option<Vec<usize>>> {
    if direct_cell_or_range_path(target) {
        return Ok(None);
    }

    let nav = Nav::new(file);
    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&nav, path)?;
        return match node.element() {
            Kind::Sheet => Ok(Some(vec![node.sheet_position()])),
            Kind::Cell | Kind::Range => Ok(None),
            _ => Err(unsupported_target(node)),
        };
    }

    let Some(query) = target.query() else {
        return Err(
            Error::new(Code::InvalidArgument, "the mutation target is empty")
                .with_format(FileFormat::Xlsx),
        );
    };
    let walk = crate::query::run(nav, query)?;
    if !walk.plan().diagnostics().is_empty() {
        return Err(Error::new(
            Code::InvalidSelector,
            "the mutation selector contains an unsupported predicate",
        )
        .with_format(FileFormat::Xlsx));
    }

    let mut positions = Vec::new();
    let mut seen = HashSet::new();
    let mut matched = false;
    let mut saw_sheet = false;
    for node in walk {
        let node = node?;
        matched = true;
        match node.element() {
            Kind::Sheet => {
                saw_sheet = true;
                append_position(&mut positions, &mut seen, node.sheet_position(), limits)?;
            },
            Kind::Cell | Kind::Range => {
                if saw_sheet {
                    return Err(Error::new(
                        Code::UnsupportedElement,
                        "an XLSX mutation selector cannot mix sheets with cells or ranges",
                    )
                    .with_format(FileFormat::Xlsx));
                }
                // The cell adapter owns cell/range admission and will rerun
                // this selector with its own bounded expansion policy.  Do
                // not scan an unbounded cell query merely to prove that no
                // later match happens to be a sheet.
                return Ok(None);
            },
            _ => return Err(unsupported_target(node)),
        }
    }
    file.check_source_for_replay()?;
    if !matched {
        return Err(Error::new(
            Code::ZeroMatches,
            "the mutation selector matched no element",
        )
        .with_format(FileFormat::Xlsx));
    }
    Ok(Some(positions))
}

fn direct_cell_or_range_path(target: &TargetArg<'_>) -> bool {
    target.path().is_some_and(|path| {
        // A root-level `/A1` can be a perfectly valid sheet name.  A
        // cell/range path has a sheet scope before its final A1 step.
        path.path().steps().len() > 1
            && path
                .path()
                .steps()
                .last()
                .is_some_and(|step| step.cell_range().is_some())
    })
}

fn unsupported_target(node: Node<'_>) -> Error {
    Error::new(
        Code::UnsupportedElement,
        "the XLSX sheet mutation target must be a sheet",
    )
    .with_format(FileFormat::Xlsx)
    .with_suggestion(format!("resolved target kind: {}", node.kind()))
}

fn append_position(
    positions: &mut Vec<usize>,
    seen: &mut HashSet<usize>,
    position: usize,
    limits: &SemanticLimits,
) -> Result<()> {
    if seen.contains(&position) {
        return Ok(());
    }
    let observed = positions
        .len()
        .checked_add(1)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if observed > limits.max_expanded_targets() {
        return Err(expanded_limit(observed, limits.max_expanded_targets()));
    }
    positions
        .try_reserve(1)
        .map_err(|source| allocation("XLSX sheet write targets", source))?;
    seen.try_reserve(1)
        .map_err(|source| allocation("XLSX sheet write target identities", source))?;
    seen.insert(position);
    positions.push(position);
    Ok(())
}

fn check_target_count(
    count: usize,
    per_target_charge: usize,
    limits: &SemanticLimits,
) -> Result<()> {
    if count > limits.max_expanded_targets() {
        return Err(expanded_limit(count, limits.max_expanded_targets()));
    }
    let staged = count.saturating_mul(per_target_charge);
    if staged > limits.max_staged_bytes() {
        return Err(staged_limit(staged, limits.max_staged_bytes()));
    }
    Ok(())
}

fn per_target_charge(assignments: &SheetAssignments, limits: &SemanticLimits) -> Result<usize> {
    let payload = assignments
        .name
        .as_ref()
        .map_or(0, |name| name.as_str().len())
        .saturating_add(usize::from(assignments.hidden.is_some()));
    payload
        .checked_add(WRITE_ENTRY_OVERHEAD)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))
}

/// Admit all properties through the sheet write registry for the requested
/// operation.  The structure adapter uses `Operation::Add`; the immutable set
/// route uses `Operation::Set`, so the registry's access flags remain the one
/// source of truth for both verbs.
pub(crate) fn admit_sheet_properties(
    operation: Operation,
    props: &Props<'_>,
) -> Result<SheetAssignments> {
    let mut assignments = SheetAssignments::default();
    for (key, value) in props.iter() {
        let (_, entry) = SHEET_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|error| property_error(error, key))?;
        let entry = *entry;
        entry
            .prepare(operation, value, |setter, value| {
                match setter {
                    SheetSetter::Name => {
                        if assignments.name.is_some() {
                            return Err(duplicate_property(entry.spec.name));
                        }
                        let Some(Scalar::Text(value)) = value else {
                            return Err(invalid_property(
                                entry.spec.name,
                                "sheet name must be text",
                            ));
                        };
                        let name = Name::new(value).map_err(|error| {
                            native_property_error(xlsx::Error::from(error), entry.spec.name)
                        })?;
                        assignments.name = Some(name);
                    },
                    SheetSetter::Hidden => {
                        if assignments.hidden.is_some() {
                            return Err(duplicate_property(entry.spec.name));
                        }
                        let Some(Scalar::Bool(value)) = value else {
                            return Err(invalid_property(
                                entry.spec.name,
                                "sheet hidden must be boolean",
                            ));
                        };
                        assignments.hidden = Some(value);
                    },
                }
                Ok(())
            })
            .map_err(|error| property_error(error, key))?;
    }
    Ok(assignments)
}

fn duplicate_property(key: &'static str) -> Error {
    Error::new(
        Code::InvalidArgument,
        "the same XLSX sheet property was supplied more than once",
    )
    .with_property(key)
}

fn invalid_property(key: &'static str, message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message)
        .with_format(FileFormat::Xlsx)
        .with_property(key)
}

fn native_property_error(error: xlsx::Error, key: &'static str) -> Error {
    sheet_error(error)
        .caused_by_argument(Code::InvalidValue)
        .with_property(key)
}

/// Classify a native tab error at the unified XLSX boundary.
///
/// The central XLSX classifier owns the leaf-specific policy mapping (tab
/// protection, visibility guards, rename dependency blocks and resource
/// limits).  Keeping this named seam lets structural tab operations share the
/// exact same mapping while retaining the typed native cause.
pub(super) fn sheet_error(error: xlsx::Error) -> Error {
    match error {
        error @ xlsx::Error::SheetName(_) => Error::from(error)
            .with_format(FileFormat::Xlsx)
            .caused_by_argument(Code::InvalidValue)
            .with_property("name"),
        error @ xlsx::Error::SheetNameConflict { .. } => Error::from(error)
            .with_format(FileFormat::Xlsx)
            .caused_by_argument(Code::DuplicateName)
            .with_property("name"),
        error => Error::from(error).with_format(FileFormat::Xlsx),
    }
}

fn native_error(error: xlsx::Error) -> Error {
    match error {
        error @ xlsx::Error::TabEditBlocked {
            reason: xlsx::TabEditBlock::LastVisibleTab,
            ..
        } => Error::from(error)
            .with_format(FileFormat::Xlsx)
            .caused_by_argument(Code::InvalidArgument)
            .with_property("hidden"),
        error => sheet_error(error),
    }
}

fn property_error(error: Error, key: &crate::props::Key<'_>) -> Error {
    error
        .with_format(FileFormat::Xlsx)
        .with_property(key.as_str().to_owned())
}

fn expanded_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::ExpandedTargets,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xlsx)
}

fn staged_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::StagedBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Xlsx)
}

fn allocation(resource: &'static str, source: TryReserveError) -> Error {
    Error::from(xlsx::Error::Allocation { resource, source }).with_format(FileFormat::Xlsx)
}

#[cfg(test)]
mod tests {
    use super::admit_sheet_properties;
    use crate::props::Props;
    use crate::registry::Operation;

    #[test]
    fn admits_name_and_hidden_through_one_registry() {
        let props = Props::new().with("NaMe", "Input").with("HIDDEN", true);
        let assignments = admit_sheet_properties(Operation::Set, &props).expect("sheet properties");
        assert_eq!(
            assignments.name.as_ref().map(|name| name.as_str()),
            Some("Input")
        );
        assert_eq!(assignments.hidden, Some(true));
    }

    #[test]
    fn add_admission_uses_the_same_sheet_registry() {
        let props = Props::new().with("name", "Input").with("hidden", false);
        let assignments = admit_sheet_properties(Operation::Add, &props).expect("add properties");
        assert_eq!(
            assignments.name.as_ref().map(|name| name.as_str()),
            Some("Input")
        );
        assert_eq!(assignments.hidden, Some(false));
    }

    #[test]
    fn refuses_duplicate_sheet_assignments() {
        let props = Props::new().with("name", "First").with("NAME", "Second");
        let error = admit_sheet_properties(Operation::Set, &props).expect_err("duplicate name");
        assert_eq!(error.code(), crate::Code::InvalidArgument);
    }
}
