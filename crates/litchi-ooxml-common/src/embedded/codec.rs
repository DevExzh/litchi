//! Bounded traversal and validation of the embedded relationship graph.

use super::{Entry, Kind, Limits, Payload, Target};
use crate::{Error, Result};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, Part};
use std::collections::HashSet;

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
                if validated_targets.insert(part.partname()) {
                    validate_payload_relationships(
                        part,
                        &mut payload_relationship_count,
                        limits.payload_relationships,
                    )?;
                }
                Target::Internal(Payload { part })
            };

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
