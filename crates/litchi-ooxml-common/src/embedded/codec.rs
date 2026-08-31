//! Bounded traversal and validation of the embedded relationship graph.

use super::{Entry, Kind, Limits, Payload, SourceEntry, SourcePayload, SourceTarget, Target};
use crate::{Error, Result};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, Part, PartView, SourceBackedPackage};
use std::collections::{HashMap, HashSet};

// [MS-XLSB] File Structure, sections 2.1.7.36 and 2.1.7.37.
pub(super) const XLSB_DIALOG_SHEET: &str = "application/vnd.ms-excel.dialogsheet";
pub(super) const XLSB_EXTERNAL_LINK: &str = "application/vnd.ms-excel.externalLink";
pub(super) const XLSB_MACRO_SHEET: &str = "application/vnd.ms-excel.macrosheet";
pub(super) const XLSB_WORKSHEET: &str = "application/vnd.ms-excel.worksheet";

/// Inventory embedded parts with the safe general-purpose resource policy.
#[inline]
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn scan(package: &OpcPackage) -> Result<Vec<Entry<'_>>> {
    scan_with(package, &Limits::default())
}

/// Inventory embedded parts with explicit resource budgets.
///
/// Results are ordered by source part name and then relationship identifier,
/// independent of the package's internal hash-map order. Duplicate internal
/// targets remain distinct entries, while their payload relationship graph is
/// validated and charged exactly once under its canonical package part name.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn scan_with<'a>(package: &'a OpcPackage, limits: &Limits) -> Result<Vec<Entry<'a>>> {
    let limits = limits.validate()?;
    for relationship in package.rels().iter() {
        if kind(relationship.reltype()).is_some() {
            return Err(Error::Relationship(format!(
                "package-level embedded relationship '{}' has no normative source part",
                relationship.r_id()
            )));
        }
    }

    let mut entries = Vec::new();
    let mut relationship_count = 0usize;
    let mut payload_relationship_count = 0usize;
    let mut validated_targets = HashSet::new();

    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            let Some(kind) = kind(relationship.reltype()) else {
                continue;
            };
            charge(
                &mut relationship_count,
                1,
                limits.relationships,
                "embedded relationships",
            )?;
            if !is_allowed_source(kind, source.content_type()) {
                return Err(Error::Relationship(format!(
                    "{} is not a normative source for {kind:?} relationship '{}'",
                    source.partname().as_str(),
                    relationship.r_id()
                )));
            }

            let target = if relationship.is_external() {
                Target::External(relationship.target_ref())
            } else {
                if relationship.target_query().is_some() || relationship.target_fragment().is_some()
                {
                    return Err(Error::Relationship(format!(
                        "internal embedded target from {} relationship '{}' cannot contain a query or fragment",
                        source.partname().as_str(),
                        relationship.r_id()
                    )));
                }
                let target_name = relationship.target_partname().map_err(|error| {
                    Error::Relationship(format!(
                        "invalid embedded target from {} relationship '{}': {error}",
                        source.partname().as_str(),
                        relationship.r_id()
                    ))
                })?;
                let part = package.get_part(&target_name).map_err(|error| {
                    Error::Missing(format!(
                        "embedded target '{}' from {} relationship '{}': {error}",
                        target_name.as_str(),
                        source.partname().as_str(),
                        relationship.r_id()
                    ))
                })?;

                // `get_part` resolves ASCII-case differences to the stored part.
                // Keying on that stored name therefore memoizes the canonical
                // target even when source relationships use different casing.
                if !validated_targets.contains(part.partname()) {
                    reserve_hash_set(
                        &mut validated_targets,
                        "embedded payload target validation set",
                    )?;
                    validated_targets.insert(part.partname());
                    validate_payload_relationships(
                        part,
                        &mut payload_relationship_count,
                        limits.payload_relationships,
                    )?;
                }
                Target::Internal(Payload { part })
            };

            reserve_vec(&mut entries, "embedded inventory entries")?;
            entries.push(Entry {
                source: source.partname(),
                id: relationship.r_id(),
                kind,
                target,
            });
        }
    }

    entries.sort_unstable_by(|left, right| {
        left.source
            .as_str()
            .cmp(right.source.as_str())
            .then_with(|| left.id.cmp(right.id))
    });
    Ok(entries)
}

/// Inventory embedded parts from a source-backed OPC catalog without reading
/// any ordinary part payload.
///
/// The returned entries retain deferred [`SourcePayload`] views. Their
/// payload bytes are read only when an explicit `data` or `stream_to` call is
/// made on such a view.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a
/// configured bound, the source becomes stale, execution is cancelled, or an
/// underlying catalog operation fails.
pub fn scan_source(package: &SourceBackedPackage) -> Result<Vec<SourceEntry<'_>>> {
    scan_source_with(package, &Limits::default())
}

/// Inventory embedded parts from a source-backed OPC catalog with explicit
/// resource budgets.
///
/// This operation examines only the retained part, content-type, and
/// relationship metadata. Internal targets are resolved against the unique
/// canonical source catalog entry and their relationship sets are validated
/// once, while duplicate source relationships remain distinct entries.
/// Results use the same source/name and relationship-ID ordering as
/// [`scan_with`].
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a
/// configured bound, the source becomes stale, execution is cancelled, or an
/// underlying catalog operation fails.
pub fn scan_source_with<'a>(
    package: &'a SourceBackedPackage,
    limits: &Limits,
) -> Result<Vec<SourceEntry<'a>>> {
    let limits = limits.validate()?;
    check_source_state(package)?;
    let mut progress = 0usize;

    let mut catalog = Vec::new();
    let mut catalog_index = HashMap::new();
    for part in package.iter_parts() {
        check_source_progress(package, &mut progress)?;
        let canonical_name = canonical_part_name(part.partname())?;
        reserve_hash_map(&mut catalog_index, "source-backed OPC part catalog index")?;
        if catalog_index.insert(canonical_name, part).is_some() {
            return Err(Error::Relationship(format!(
                "source-backed OPC catalog contains duplicate canonical part name '{}'",
                part.partname().as_str()
            )));
        }
        reserve_vec(&mut catalog, "source-backed OPC part catalog")?;
        catalog.push(part);
    }

    for relationship in package.rels().iter() {
        check_source_progress(package, &mut progress)?;
        if kind(relationship.reltype()).is_some() {
            return Err(Error::Relationship(format!(
                "package-level embedded relationship '{}' has no normative source part",
                relationship.r_id()
            )));
        }
    }

    let mut entries = Vec::new();
    let mut relationship_count = 0usize;
    let mut payload_relationship_count = 0usize;
    let mut validated_targets: HashSet<&litchi_opc::PackURI> = HashSet::new();

    for source in &catalog {
        check_source_progress(package, &mut progress)?;
        for relationship in source.rels().iter() {
            check_source_progress(package, &mut progress)?;
            let Some(kind) = kind(relationship.reltype()) else {
                continue;
            };
            charge(
                &mut relationship_count,
                1,
                limits.relationships,
                "embedded relationships",
            )?;
            if !is_allowed_source(kind, source.content_type()) {
                return Err(Error::Relationship(format!(
                    "{} is not a normative source for {kind:?} relationship '{}'",
                    source.partname().as_str(),
                    relationship.r_id()
                )));
            }

            let target = if relationship.is_external() {
                // External references are inert metadata. In particular, do
                // not resolve or read them as if they named local parts.
                SourceTarget::External(relationship.target_ref())
            } else {
                if relationship.target_query().is_some() || relationship.target_fragment().is_some()
                {
                    return Err(Error::Relationship(format!(
                        "internal embedded target from {} relationship '{}' cannot contain a query or fragment",
                        source.partname().as_str(),
                        relationship.r_id()
                    )));
                }
                let target_name = relationship.target_partname().map_err(|error| {
                    Error::Relationship(format!(
                        "invalid embedded target from {} relationship '{}': {error}",
                        source.partname().as_str(),
                        relationship.r_id()
                    ))
                })?;
                let target_key = canonical_part_name(&target_name)?;
                let part = catalog_index
                    .get(target_key.as_slice())
                    .copied()
                    .ok_or_else(|| {
                        Error::Missing(format!(
                            "embedded target '{}' from {} relationship '{}' does not exist",
                            target_name.as_str(),
                            source.partname().as_str(),
                            relationship.r_id()
                        ))
                    })?;

                // `part` is the unique canonical catalog view, so spelling
                // differences in source relationships share one validation.
                if !validated_targets.contains(part.partname()) {
                    reserve_hash_set(
                        &mut validated_targets,
                        "source-backed embedded payload target validation set",
                    )?;
                    validated_targets.insert(part.partname());
                    validate_payload_relationships_source(
                        &part,
                        &mut payload_relationship_count,
                        limits.payload_relationships,
                    )?;
                }
                SourceTarget::Internal(SourcePayload { part })
            };

            reserve_vec(&mut entries, "source-backed embedded inventory entries")?;
            entries.push(SourceEntry {
                source: source.partname(),
                id: relationship.r_id(),
                kind,
                target,
            });
        }
    }

    entries.sort_unstable_by(|left, right| {
        left.source
            .as_str()
            .cmp(right.source.as_str())
            .then_with(|| left.id.cmp(right.id))
    });
    check_source_state(package)?;
    Ok(entries)
}

fn kind(relationship: &str) -> Option<Kind> {
    match relationship {
        relationship_type::OLE_OBJECT | relationship_type::STRICT_OLE_OBJECT => Some(Kind::Object),
        relationship_type::PACKAGE | relationship_type::STRICT_PACKAGE => Some(Kind::Package),
        _ => None,
    }
}

fn is_allowed_source(kind: Kind, source_content_type: &str) -> bool {
    let common = matches!(
        source_content_type,
        content_type::WML_COMMENTS
            | content_type::WML_ENDNOTES
            | content_type::WML_FOOTER
            | content_type::WML_FOOTNOTES
            | content_type::WML_HEADER
            | content_type::WML_DOCUMENT_MAIN
            | content_type::WML_TEMPLATE_MAIN
            | content_type::WML_DOCUMENT_MACRO_MAIN
            | content_type::WML_TEMPLATE_MACRO_MAIN
            | content_type::SML_WORKSHEET
            | content_type::PML_HANDOUT_MASTER
            | content_type::PML_NOTES_SLIDE
            | content_type::PML_NOTES_MASTER
            | content_type::PML_SLIDE
            | content_type::PML_SLIDE_LAYOUT
            | content_type::PML_SLIDE_MASTER
    );
    if common {
        return true;
    }

    matches!(
        (kind, source_content_type),
        (Kind::Object, XLSB_EXTERNAL_LINK)
            | (
                Kind::Object | Kind::Package,
                XLSB_DIALOG_SHEET | XLSB_MACRO_SHEET | XLSB_WORKSHEET
            )
            | (Kind::Package, content_type::DML_CHART)
    )
}

fn validate_payload_relationships(payload: &dyn Part, count: &mut usize, max: usize) -> Result<()> {
    for relationship in payload.rels().iter() {
        charge(count, 1, max, "embedded payload relationships")?;
        if !matches!(
            relationship.reltype(),
            relationship_type::HYPERLINK | relationship_type::STRICT_HYPERLINK
        ) {
            return Err(Error::Relationship(format!(
                "embedded target '{}' has forbidden relationship '{}' of type '{}'",
                payload.partname().as_str(),
                relationship.r_id(),
                relationship.reltype()
            )));
        }
    }
    Ok(())
}

fn validate_payload_relationships_source(
    payload: &PartView<'_>,
    count: &mut usize,
    max: usize,
) -> Result<()> {
    for relationship in payload.rels().iter() {
        charge(count, 1, max, "embedded payload relationships")?;
        if !matches!(
            relationship.reltype(),
            relationship_type::HYPERLINK | relationship_type::STRICT_HYPERLINK
        ) {
            return Err(Error::Relationship(format!(
                "embedded target '{}' has forbidden relationship '{}' of type '{}'",
                payload.partname().as_str(),
                relationship.r_id(),
                relationship.reltype()
            )));
        }
    }
    Ok(())
}

fn check_source_state(package: &SourceBackedPackage) -> Result<()> {
    package.check_execution()?;
    package.source_version()?;
    Ok(())
}

fn check_source_progress(package: &SourceBackedPackage, progress: &mut usize) -> Result<()> {
    package.check_execution()?;
    *progress = progress.checked_add(1).ok_or(Error::Limit {
        resource: "source-backed embedded scan progress",
        max: usize::MAX,
        actual: usize::MAX,
    })?;
    if *progress % 256 == 0 {
        package.source_version()?;
    }
    Ok(())
}

fn reserve_vec<T>(collection: &mut Vec<T>, resource: &'static str) -> Result<()> {
    collection
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

fn reserve_hash_set<T>(collection: &mut HashSet<T>, resource: &'static str) -> Result<()>
where
    T: Eq + std::hash::Hash,
{
    collection
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

fn reserve_hash_map<K, V>(collection: &mut HashMap<K, V>, resource: &'static str) -> Result<()>
where
    K: Eq + std::hash::Hash,
{
    collection
        .try_reserve(1)
        .map_err(|source| Error::Allocation { resource, source })
}

fn canonical_part_name(partname: &litchi_opc::PackURI) -> Result<Vec<u8>> {
    let source = partname.as_str().as_bytes();
    let mut canonical = Vec::new();
    canonical
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "canonical source-backed OPC part name",
            source,
        })?;
    canonical.extend(source.iter().map(u8::to_ascii_lowercase));
    Ok(canonical)
}

fn charge(total: &mut usize, amount: usize, max: usize, resource: &'static str) -> Result<()> {
    let actual = total.checked_add(amount).ok_or(Error::Limit {
        resource,
        max,
        actual: usize::MAX,
    })?;
    if actual > max {
        return Err(Error::Limit {
            resource,
            max,
            actual,
        });
    }
    *total = actual;
    Ok(())
}
