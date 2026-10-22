//! Package-graph plumbing shared by the package-level property owners.
//!
//! The owners inspect the same graph over an owning [`OpcPackage`] and over
//! the metadata-only catalog of a [`SourceBackedPackage`]. [`GraphPart`] is
//! the part projection both expose, so one inspection serves both ingress
//! modes and their strict rules cannot drift apart.

use litchi_opc::{OpcPackage, PackURI, PartMetadata, PartView, Relationships, SourceBackedPackage};

use crate::{Error, Result};

/// The metadata of one ordinary part: its name, content type and
/// relationships. No payload is read.
pub(crate) trait GraphPart {
    fn part_name(&self) -> &PackURI;
    fn part_content_type(&self) -> &str;
    fn part_relationships(&self) -> &Relationships;
}

impl GraphPart for PartMetadata<'_> {
    fn part_name(&self) -> &PackURI {
        self.partname()
    }

    fn part_content_type(&self) -> &str {
        self.content_type()
    }

    fn part_relationships(&self) -> &Relationships {
        self.rels()
    }
}

impl GraphPart for PartView<'_> {
    fn part_name(&self) -> &PackURI {
        self.partname()
    }

    fn part_content_type(&self) -> &str {
        self.content_type()
    }

    fn part_relationships(&self) -> &Relationships {
        self.rels()
    }
}

/// The parts and package-level relationships of either package kind.
pub(crate) trait PackageMetadata {
    type Part<'a>: GraphPart
    where
        Self: 'a;

    fn metadata_parts(&self) -> impl Iterator<Item = Self::Part<'_>>;
    fn package_relationships(&self) -> &Relationships;
}

impl PackageMetadata for OpcPackage {
    type Part<'a> = PartMetadata<'a>;

    fn metadata_parts(&self) -> impl Iterator<Item = Self::Part<'_>> {
        self.iter_parts()
    }

    fn package_relationships(&self) -> &Relationships {
        self.rels()
    }
}

impl PackageMetadata for SourceBackedPackage {
    type Part<'a> = PartView<'a>;

    fn metadata_parts(&self) -> impl Iterator<Item = Self::Part<'_>> {
        self.iter_parts()
    }

    fn package_relationships(&self) -> &Relationships {
        self.rels()
    }
}

/// Whether a part name equals `name` under OPC's ASCII case folding.
pub(crate) fn same_part_name(left: &PackURI, right: &PackURI) -> bool {
    left.as_str().eq_ignore_ascii_case(right.as_str())
}

/// Whether two part names cannot coexist in one package: they are
/// equivalent, or one is a segment prefix of the other (OPC M1.11, M1.12).
pub(crate) fn names_conflict(left: &PackURI, right: &PackURI) -> bool {
    fn derived(candidate: &str, base: &str) -> bool {
        candidate.len() > base.len()
            && candidate
                .get(..base.len())
                .is_some_and(|prefix| prefix.eq_ignore_ascii_case(base))
            && candidate.as_bytes().get(base.len()) == Some(&b'/')
    }
    same_part_name(left, right)
        || derived(left.as_str(), right.as_str())
        || derived(right.as_str(), left.as_str())
}

/// Why removing `target` would leave a relationship dangling: anything but
/// the package-level relationship `owner_id` targets it. `None` when nothing
/// else does.
///
/// # Errors
///
/// Returns [`Error::Relationship`] for an unparseable relationship target.
pub(crate) fn other_inbound<P: PackageMetadata>(
    package: &P,
    target: &PackURI,
    owner_id: &str,
    label: &str,
) -> Result<Option<String>> {
    for relationship in package.package_relationships().iter() {
        if relationship.r_id() == owner_id || relationship.is_external() {
            continue;
        }
        let related = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!("invalid package relationship target: {error}"))
        })?;
        if same_part_name(&related, target) {
            return Ok(Some(format!(
                "{label} part '{}' has another package-level inbound relationship",
                target.as_str()
            )));
        }
    }
    for source in package.metadata_parts() {
        for relationship in source.part_relationships().iter() {
            if relationship.is_external() {
                continue;
            }
            let related = relationship.target_partname().map_err(|error| {
                Error::Relationship(format!(
                    "invalid relationship target from '{}': {error}",
                    source.part_name().as_str()
                ))
            })?;
            if same_part_name(&related, target) {
                return Ok(Some(format!(
                    "{label} part '{}' is shared by '{}'",
                    target.as_str(),
                    source.part_name().as_str()
                )));
            }
        }
    }
    Ok(None)
}

/// Refuse when anything but the package-level relationship `owner_id`
/// targets `target` ([`other_inbound`]).
///
/// # Errors
///
/// Returns [`Error::Relationship`] naming the other owner, or an error for an
/// unparseable relationship target.
pub(crate) fn ensure_only_inbound<P: PackageMetadata>(
    package: &P,
    target: &PackURI,
    owner_id: &str,
    label: &str,
) -> Result<()> {
    match other_inbound(package, target, owner_id, label)? {
        Some(reason) => Err(Error::Relationship(reason)),
        None => Ok(()),
    }
}

/// The first free name among `{stem}.xml`, `{stem}2.xml`, … `{stem}N.xml`,
/// as the core-properties writer allocates them.
///
/// A name is free when no existing part has it or conflicts with it (OPC
/// part names fold ASCII case, and a part name cannot be the prefix segment
/// of another).
///
/// # Errors
///
/// Returns [`Error::Limit`] when every candidate is taken, or an error for a
/// malformed candidate.
pub(crate) fn free_part_name<P: PackageMetadata>(
    package: &P,
    stem: &str,
    resource: crate::LimitKind,
) -> Result<PackURI> {
    const MAX_CANDIDATES: u32 = 10_000;
    for index in 1..=MAX_CANDIDATES {
        let candidate = if index == 1 {
            format!("{stem}.xml")
        } else {
            format!("{stem}{index}.xml")
        };
        let candidate = PackURI::new(&candidate).map_err(Error::Uri)?;
        if package
            .metadata_parts()
            .all(|part| !names_conflict(part.part_name(), &candidate))
        {
            return Ok(candidate);
        }
    }
    Err(Error::Limit {
        kind: resource,
        maximum: MAX_CANDIDATES as usize,
        observed: MAX_CANDIDATES as usize,
    })
}
