//! Document metadata of a workbook: the core, application and custom
//! property families, edited as one exact and reversible transaction.
//!
//! [`Package::metadata`] reads the unified [`Metadata`];
//! [`Package::edit_metadata`] starts an [`Edit`] whose [`Edit::commit`]
//! returns a [`Commit`] holding the new package, the [`Patch`] and the
//! [`Diagnostics`]. [`Package::apply_metadata`] replays a patch onto an
//! equal package. The source package never changes.
//!
//! The shared OOXML engine ([`litchi_ooxml_common::properties`]) routes each
//! field to the part that stores it and keeps the lexical form of every
//! unchanged value. A field no OOXML family stores is refused with
//! [`Error::Refused`]; a changed edit of a signed workbook is refused
//! with [`Error::Signed`]; an exact no-op keeps every byte.
//!
//! ```
//! use litchi_xlsx::Package;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::new()?;
//! let mut edit = package.edit_metadata()?;
//! edit.update(|metadata| metadata.title = Some("Budget".to_owned()))?;
//! let after = edit.commit()?.into_package();
//! assert_eq!(after.metadata()?.title.as_deref(), Some("Budget"));
//! assert_eq!(package.metadata()?.title, None);
//! # Ok(())
//! # }
//! ```

use std::mem::size_of;

use litchi_core::Metadata;
use litchi_ooxml_common::custom::{self, Host};
use litchi_ooxml_common::properties::{self, Family, Snapshot, app};
use litchi_opc::{PartMetadata, Relationship, Relationships};

const STRICT_CORE_PROPERTIES_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/package/relationships/metadata/core-properties";
const STRICT_EXTENDED_PROPERTIES_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties";

pub use litchi_ooxml_common::properties::{Diagnostics, Error, LimitKind, Patch};

use crate::Package;

type Result<T> = std::result::Result<T, Error>;

fn retained_relationship_bytes(relationship: &Relationship) -> usize {
    size_of::<Relationship>()
        .saturating_add(relationship.r_id().len())
        .saturating_add(relationship.reltype().len())
        .saturating_add(relationship.target_ref().len())
        .saturating_add(relationship.base_uri().len())
        .saturating_add(relationship.source_uri().map_or(0, str::len))
}

fn retained_relationships_bytes(relationships: &Relationships) -> usize {
    size_of::<Relationships>().saturating_add(
        relationships
            .iter()
            .map(retained_relationship_bytes)
            .sum::<usize>(),
    )
}

fn retained_part_metadata_bytes(part: PartMetadata<'_>) -> usize {
    size_of::<PartMetadata<'_>>()
        .saturating_add(part.partname().as_str().len())
        .saturating_add(part.content_type().len())
        .saturating_add(retained_relationships_bytes(part.rels()))
}

fn is_family_relationship(family: Family, relationship_type: &str) -> bool {
    match family {
        Family::Core => matches!(
            relationship_type,
            litchi_opc::constants::relationship_type::CORE_PROPERTIES
                | STRICT_CORE_PROPERTIES_RELATIONSHIP
        ),
        Family::App => matches!(
            relationship_type,
            litchi_opc::constants::relationship_type::EXTENDED_PROPERTIES
                | STRICT_EXTENDED_PROPERTIES_RELATIONSHIP
        ),
        Family::Custom => {
            relationship_type == litchi_opc::constants::relationship_type::CUSTOM_PROPERTIES
        },
        _ => false,
    }
}

fn owner_family(
    package: &Package,
    snapshot: &Snapshot,
    family: Family,
) -> Result<properties::FamilyAdmission> {
    let package_relationships = package.opc().rels();
    let (present_part_bytes, present_relationship_bytes) = match snapshot.part_name(family) {
        Some(part_name) => {
            let part = package
                .opc()
                .part_metadata(part_name)
                .ok_or(Error::Internal(
                    "metadata part disappeared during admission",
                ))?;
            let relationship = match snapshot.relationship_id_value(family) {
                Some(relationship_id) => package_relationships.get(relationship_id),
                None => package_relationships.iter().find(|relationship| {
                    is_family_relationship(family, relationship.reltype())
                        && relationship.target_partname().ok().is_some_and(|target| {
                            target.as_str().eq_ignore_ascii_case(part_name.as_str())
                        })
                }),
            }
            .ok_or(Error::Internal(
                "metadata relationship disappeared during admission",
            ))?;
            (
                Some(retained_part_metadata_bytes(part)),
                Some(retained_relationship_bytes(relationship)),
            )
        },
        None => (None, None),
    };
    let new_part = snapshot.new_part_name(family);
    let created_part_bytes = size_of::<PartMetadata<'_>>()
        .saturating_add(new_part.as_str().len())
        .saturating_add(family.content_type().len())
        .saturating_add(size_of::<Relationships>());
    let relationship_type = match family {
        Family::Core => litchi_opc::constants::relationship_type::CORE_PROPERTIES,
        Family::App => litchi_opc::constants::relationship_type::EXTENDED_PROPERTIES,
        Family::Custom => litchi_opc::constants::relationship_type::CUSTOM_PROPERTIES,
        _ => return Err(Error::Internal("unknown OOXML metadata family")),
    };
    let created_relationship_base_bytes = size_of::<Relationship>()
        .saturating_add(relationship_type.len())
        .saturating_add(new_part.as_str().trim_start_matches('/').len())
        .saturating_add(package_relationships.base_uri().len())
        .saturating_add(package_relationships.source_uri().map_or(0, str::len));
    Ok(properties::FamilyAdmission::new(
        present_part_bytes,
        present_relationship_bytes,
        created_part_bytes,
        created_relationship_base_bytes,
    ))
}

fn owner_admission(package: &Package, snapshot: &Snapshot) -> Result<properties::OwnerAdmission> {
    Ok(properties::OwnerAdmission::new(
        package.retained_bytes(),
        package.opc().exact_source_len().unwrap_or(0),
        [
            owner_family(package, snapshot, Family::Core)?,
            owner_family(package, snapshot, Family::App)?,
            owner_family(package, snapshot, Family::Custom)?,
        ],
    ))
}

/// Bounds for the source package, target package, and exact metadata patch.
///
/// Repeated endpoint payloads are charged independently; derived cell and
/// shared-string read caches are outside this semantic retention budget.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct EditOptions {
    max_retained_bytes: Option<usize>,
}

impl EditOptions {
    /// Creates an unbounded metadata edit.
    pub const fn new() -> Self {
        Self {
            max_retained_bytes: None,
        }
    }

    /// Sets the combined source, target, and exact patch retention ceiling.
    pub const fn max_retained_bytes(mut self, maximum: usize) -> Self {
        self.max_retained_bytes = Some(maximum);
        self
    }

    /// Returns the selected semantic retention ceiling.
    #[must_use]
    pub const fn max_retained_bytes_value(self) -> Option<usize> {
        self.max_retained_bytes
    }
}

/// A staged metadata edit of one [`Package`].
///
/// Starting one retains only the source package and options; metadata parts
/// and the semantic draft are materialized by the first getter, mutation or
/// commit. Dropping the edit discards it.
#[derive(Debug)]
pub struct Edit {
    package: Package,
    options: EditOptions,
    inner: Option<properties::Edit>,
}

impl Edit {
    fn ensure_inner(&mut self) -> Result<&mut properties::Edit> {
        if self.inner.is_none() {
            let snapshot = Snapshot::read(self.package.opc(), Host::Excel)?;
            let common_options = self
                .options
                .max_retained_bytes_value()
                .map_or_else(properties::EditOptions::default, |maximum| {
                    properties::EditOptions::new().max_retained_bytes(maximum)
                });
            let common_options = if self.options.max_retained_bytes_value().is_some() {
                common_options.with_owner_admission(owner_admission(&self.package, &snapshot)?)
            } else {
                common_options
            };
            // The snapshot was read from this very package just above.
            self.inner = Some(properties::Edit::new_with_own_snapshot(
                self.package.opc(),
                snapshot,
                common_options,
            ));
        }
        Ok(self.inner.as_mut().expect("metadata edit loaded"))
    }

    /// The staged metadata.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of the application-properties part.
    pub fn metadata(&mut self) -> Result<Metadata> {
        self.ensure_inner()?.metadata()
    }

    /// Stage the complete metadata.
    pub fn set(&mut self, metadata: Metadata) -> Result<&mut Self> {
        self.ensure_inner()?.set_metadata(metadata);
        Ok(self)
    }

    /// Stage a change of the metadata, starting from the staged value.
    ///
    /// # Errors
    ///
    /// Returns the decoding errors of the application-properties part.
    pub fn update(&mut self, update: impl FnOnce(&mut Metadata)) -> Result<&mut Self> {
        self.ensure_inner()?.update_metadata(update)?;
        Ok(self)
    }

    /// Stage the complete application properties; `None` removes the part.
    pub fn set_app_props(&mut self, props: Option<app::Props>) -> Result<&mut Self> {
        self.ensure_inner()?.set_app_props(props);
        Ok(self)
    }

    /// The staged custom properties.
    pub fn custom_props(&mut self) -> Result<&custom::Props> {
        Ok(self.ensure_inner()?.custom_props())
    }

    /// Stage the complete custom properties; an empty collection removes the
    /// part when it changes the semantic collection. Setting an equal empty
    /// collection is an exact no-op, preserving a present raw part.
    pub fn set_custom_props(&mut self, props: custom::Props) -> Result<&mut Self> {
        self.ensure_inner()?.set_custom_props(props);
        Ok(self)
    }

    /// Stage explicit removal of the custom-properties part and relationship.
    ///
    /// This is distinct from setting an equal empty collection: it removes a
    /// present raw part even when its tolerant semantic projection is empty.
    /// An absent graph remains an exact no-op.
    pub fn remove_custom_props(&mut self) -> Result<&mut Self> {
        self.ensure_inner()?.remove_custom_props();
        Ok(self)
    }

    /// Compute the patch and publish it to a new snapshot.
    ///
    /// An exact no-op returns the same snapshot.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for a field no OOXML family stores,
    /// [`Error::Invalid`] or [`Error::Limit`] for a value the families
    /// cannot carry, and [`Error::Signed`] for a changed edit of a signed
    /// workbook.
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_inner()?;
        let inner = self.inner.take().expect("metadata edit loaded");
        let commit = inner.commit()?;
        let diagnostics = commit.diagnostics().clone();
        let (opc, patch) = commit.into_parts();
        let package = if patch.is_noop() {
            self.package
        } else {
            self.package.adopt_metadata(opc)?
        };
        Ok(Commit {
            package,
            patch,
            diagnostics,
        })
    }
}

/// The result of [`Edit::commit`].
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// The package after the edit.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit into the package after the edit.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// The exact, reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// What the edit changed.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Consume the commit into the package and the patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

impl Package {
    fn metadata_snapshot(&self) -> Result<Snapshot> {
        Snapshot::read(self.opc(), Host::Excel)
    }

    /// The document metadata: the core and application properties.
    ///
    /// # Errors
    ///
    /// Returns the strict graph and decoding errors of the property parts.
    pub fn metadata(&self) -> Result<Metadata> {
        self.metadata_snapshot()?.metadata()
    }

    /// The OOXML application properties (`docProps/app.xml`), when present.
    ///
    /// # Errors
    ///
    /// Returns the strict graph and decoding errors of the part.
    pub fn app_props(&self) -> Result<Option<app::Props>> {
        self.metadata_snapshot()?.app_props()
    }

    /// Start a metadata edit.
    ///
    /// # Errors
    ///
    /// No source graph is read at entry. Property graph and decoding errors
    /// are deferred to the first getter, mutation or commit.
    pub fn edit_metadata(&self) -> Result<Edit> {
        self.edit_metadata_with_options(EditOptions::default())
    }

    /// Start a metadata edit under an explicit source, target, and patch ceiling.
    pub fn edit_metadata_with_options(&self, options: EditOptions) -> Result<Edit> {
        Ok(Edit {
            package: self.clone(),
            options,
            inner: None,
        })
    }

    /// Apply a metadata patch to a new package; this one is unchanged.
    ///
    /// An empty patch returns this package, a signed one included. The new
    /// package keeps this one's encryption provenance and repair policy.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Commit> {
        if patch.is_noop() {
            // An exact no-op still authenticates every recorded property
            // family, but it must not clone the complete OPC graph merely to
            // discover that no bytes will change.
            patch.check(self.opc())?;
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: patch.diagnostics(),
            });
        }
        let candidate = patch.apply(self.opc())?.into_package();
        let package = self.adopt_metadata(candidate)?;
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: patch.diagnostics(),
        })
    }

    /// Wrap a package whose document properties changed as a new package.
    fn adopt_metadata(&self, candidate: litchi_opc::OpcPackage) -> Result<Self> {
        self.workbook_ref()
            .adopt_published_package(candidate)
            .map(Self::from_workbook)
            .map_err(|error| Error::Invalid(error.to_string()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn title_commit(package: &Package, title: &str) -> Commit {
        let mut edit = package.edit_metadata().unwrap();
        edit.update(|metadata| metadata.title = Some(title.to_owned()))
            .unwrap();
        edit.commit().unwrap()
    }

    fn warm_package() -> Package {
        let source = Package::new().unwrap();
        let mut edit = source.edit().unwrap();
        edit.sheet("Sheet1")
            .unwrap()
            .unwrap()
            .set("A1", "warm")
            .unwrap();
        edit.commit().unwrap().into_package()
    }

    fn warm_cells(package: &Package) {
        let workbook = package.workbook();
        let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
        assert!(sheet.cell("A1").unwrap().stored().is_some());
    }

    fn case_spelled_core_target() -> Package {
        let source = title_commit(&Package::new().unwrap(), "source").into_package();
        let mut opc = source.opc().clone();
        let relationship_id = opc
            .rels()
            .iter()
            .find(|relationship| {
                relationship.reltype() == litchi_opc::constants::relationship_type::CORE_PROPERTIES
            })
            .expect("core-properties relationship")
            .r_id()
            .to_owned();
        opc.rels_mut()
            .retarget(&relationship_id, "docProps/CORE.xml".to_owned())
            .unwrap();
        Package::from_opc(opc).unwrap()
    }

    #[test]
    fn edits_publish_a_new_snapshot_and_replay_exactly() {
        let package = Package::new().unwrap();
        let original = package.to_bytes().unwrap();
        let mut edit = package.edit_metadata().unwrap();
        edit.update(|metadata| {
            metadata.author = Some("Ada".to_owned());
            metadata.company = Some("Litchi".to_owned());
        })
        .unwrap();
        let commit = edit.commit().unwrap();
        assert!(commit.diagnostics().changed());
        let (after, patch) = commit.into_parts();
        assert_eq!(after.metadata().unwrap().author.as_deref(), Some("Ada"));
        assert_eq!(
            after.app_props().unwrap().unwrap().company.as_deref(),
            Some("Litchi")
        );
        assert!(!package.same_content(&after).unwrap());

        let replayed = package.apply_metadata(&patch).unwrap();
        assert!(replayed.diagnostics().changed());
        let replayed = replayed.into_package();
        assert!(replayed.same_content(&after).unwrap());
        assert_eq!(replayed.to_bytes().unwrap(), after.to_bytes().unwrap());
        assert!(matches!(
            package.apply_metadata(&patch.inverse()),
            Err(Error::PatchConflict)
        ));
        let restored = after
            .apply_metadata(&patch.inverse())
            .unwrap()
            .into_package();
        assert!(restored.same_content(&package).unwrap());
        assert_eq!(restored.to_bytes().unwrap(), original);
    }

    #[test]
    fn no_op_edits_return_the_same_snapshot() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_metadata().unwrap();
        edit.set(package.metadata().unwrap()).unwrap();
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert!(commit.package().same_content(&package).unwrap());
    }

    #[test]
    fn no_op_patch_replay_still_requires_its_recorded_property_scope() {
        let source = Package::new().unwrap();
        let mut edit = source.edit_metadata().unwrap();
        edit.set(source.metadata().unwrap()).unwrap();
        let no_op = edit.commit().unwrap();
        assert!(no_op.patch().is_noop());

        let mut other_edit = source.edit_metadata().unwrap();
        other_edit
            .update(|metadata| metadata.title = Some("Different owner".to_owned()))
            .unwrap();
        let other = other_edit.commit().unwrap().into_package();

        let error = other
            .apply_metadata(no_op.patch())
            .expect_err("a no-op patch must authenticate its recorded base");
        assert!(matches!(error, Error::PatchConflict), "{error:?}");

        let replayed = source.apply_metadata(no_op.patch()).unwrap();
        assert!(replayed.patch().is_noop());
        assert!(replayed.package().same_content(&source).unwrap());
    }

    #[test]
    fn bounded_metadata_admission_accepts_exact_shrinking_endpoint() {
        let source = Package::new().unwrap();
        let base = title_commit(&source, &"long title ".repeat(64)).into_package();
        let expected = title_commit(&base, "short");
        let required = base
            .retained_bytes()
            .saturating_add(expected.package().retained_bytes())
            .saturating_add(expected.patch().retained_bytes());

        let exact = base
            .edit_metadata_with_options(EditOptions::new().max_retained_bytes(required))
            .unwrap();
        let mut exact = exact;
        exact
            .update(|metadata| metadata.title = Some("short".to_owned()))
            .unwrap();
        assert!(exact.commit().is_ok());

        let mut below = base
            .edit_metadata_with_options(EditOptions::new().max_retained_bytes(required - 1))
            .unwrap();
        below
            .update(|metadata| metadata.title = Some("short".to_owned()))
            .unwrap();
        assert!(matches!(
            below.commit(),
            Err(Error::Limit {
                kind: LimitKind::RetainedBytes,
                observed,
                maximum,
            }) if observed > maximum && maximum == required - 1
        ));
    }

    #[test]
    fn bounded_metadata_admission_accepts_exact_expanding_endpoint() {
        let source = Package::new().unwrap();
        let base = title_commit(&source, "short").into_package();
        let expected = title_commit(&base, &"expanded title ".repeat(64));
        let required = base
            .retained_bytes()
            .saturating_add(expected.package().retained_bytes())
            .saturating_add(expected.patch().retained_bytes());

        let mut exact = base
            .edit_metadata_with_options(EditOptions::new().max_retained_bytes(required))
            .unwrap();
        exact
            .update(|metadata| metadata.title = Some("expanded title ".repeat(64)))
            .unwrap();
        assert!(exact.commit().is_ok());
    }

    #[test]
    fn bounded_metadata_admission_charges_an_exact_noop_pair() {
        let package = Package::new().unwrap();
        let expected = package.edit_metadata().unwrap().commit().unwrap();
        let required = package
            .retained_bytes()
            .saturating_add(expected.package().retained_bytes())
            .saturating_add(expected.patch().retained_bytes());

        let exact = package
            .edit_metadata_with_options(EditOptions::new().max_retained_bytes(required))
            .unwrap()
            .commit();
        assert!(exact.is_ok());

        let below = package
            .edit_metadata_with_options(EditOptions::new().max_retained_bytes(required - 1))
            .unwrap()
            .commit();
        assert!(matches!(
            below,
            Err(Error::Limit {
                kind: LimitKind::RetainedBytes,
                observed,
                maximum,
            }) if observed > maximum && maximum == required - 1
        ));
    }

    #[test]
    fn bounded_metadata_admission_is_cache_independent_at_boundaries() {
        let long = title_commit(&warm_package(), &"long title ".repeat(64)).into_package();
        warm_cells(&long);
        let cold_long = Package::from_opc(long.opc().clone()).unwrap();
        assert_eq!(
            long.retained_bytes(),
            cold_long.retained_bytes(),
            "warming a semantic read cache must not change the retained charge"
        );
        let shrinking = title_commit(&long, "short");
        let shrinking_required = long
            .retained_bytes()
            .saturating_add(shrinking.package().retained_bytes())
            .saturating_add(shrinking.patch().retained_bytes());
        let mut shrinking_edit = long
            .edit_metadata_with_options(EditOptions::new().max_retained_bytes(shrinking_required))
            .unwrap();
        shrinking_edit
            .update(|metadata| metadata.title = Some("short".to_owned()))
            .unwrap();
        assert!(shrinking_edit.commit().is_ok());
        let mut shrinking_below = long
            .edit_metadata_with_options(
                EditOptions::new().max_retained_bytes(shrinking_required - 1),
            )
            .unwrap();
        shrinking_below
            .update(|metadata| metadata.title = Some("short".to_owned()))
            .unwrap();
        assert!(matches!(
            shrinking_below.commit(),
            Err(Error::Limit {
                kind: LimitKind::RetainedBytes,
                ..
            })
        ));

        let short = title_commit(&warm_package(), "short").into_package();
        warm_cells(&short);
        let expanding = title_commit(&short, &"expanded title ".repeat(64));
        let expanding_required = short
            .retained_bytes()
            .saturating_add(expanding.package().retained_bytes())
            .saturating_add(expanding.patch().retained_bytes());
        let mut expanding_edit = short
            .edit_metadata_with_options(EditOptions::new().max_retained_bytes(expanding_required))
            .unwrap();
        expanding_edit
            .update(|metadata| metadata.title = Some("expanded title ".repeat(64)))
            .unwrap();
        assert!(expanding_edit.commit().is_ok());
        let mut expanding_below = short
            .edit_metadata_with_options(
                EditOptions::new().max_retained_bytes(expanding_required - 1),
            )
            .unwrap();
        expanding_below
            .update(|metadata| metadata.title = Some("expanded title ".repeat(64)))
            .unwrap();
        assert!(matches!(
            expanding_below.commit(),
            Err(Error::Limit {
                kind: LimitKind::RetainedBytes,
                ..
            })
        ));

        let noop = warm_package();
        warm_cells(&noop);
        let cold_noop = Package::from_opc(noop.opc().clone()).unwrap();
        assert_eq!(noop.retained_bytes(), cold_noop.retained_bytes());
        let expected_noop = noop.edit_metadata().unwrap().commit().unwrap();
        let noop_required = noop
            .retained_bytes()
            .saturating_add(expected_noop.package().retained_bytes())
            .saturating_add(expected_noop.patch().retained_bytes());
        assert!(
            noop.edit_metadata_with_options(EditOptions::new().max_retained_bytes(noop_required))
                .unwrap()
                .commit()
                .is_ok()
        );
        assert!(matches!(
            noop.edit_metadata_with_options(
                EditOptions::new().max_retained_bytes(noop_required - 1)
            )
            .unwrap()
            .commit(),
            Err(Error::Limit {
                kind: LimitKind::RetainedBytes,
                ..
            })
        ));
    }

    #[test]
    fn bounded_admission_preserves_case_spelled_core_relationships() {
        let base = case_spelled_core_target();
        assert_eq!(
            base.opc()
                .rels()
                .iter()
                .find(|relationship| {
                    relationship.reltype()
                        == litchi_opc::constants::relationship_type::CORE_PROPERTIES
                })
                .expect("core-properties relationship")
                .target_ref(),
            "docProps/CORE.xml"
        );

        let unbounded = title_commit(&base, "unbounded");
        assert_eq!(
            unbounded
                .package()
                .opc()
                .rels()
                .iter()
                .find(|relationship| {
                    relationship.reltype()
                        == litchi_opc::constants::relationship_type::CORE_PROPERTIES
                })
                .expect("core-properties relationship")
                .target_ref(),
            "docProps/CORE.xml"
        );

        let expected = title_commit(&base, "bounded");
        let required = base
            .retained_bytes()
            .saturating_add(expected.package().retained_bytes())
            .saturating_add(expected.patch().retained_bytes());
        let mut bounded = base
            .edit_metadata_with_options(EditOptions::new().max_retained_bytes(required))
            .unwrap();
        bounded
            .update(|metadata| metadata.title = Some("bounded".to_owned()))
            .unwrap();
        let bounded = bounded.commit().unwrap();
        assert_eq!(
            bounded
                .package()
                .opc()
                .rels()
                .iter()
                .find(|relationship| {
                    relationship.reltype()
                        == litchi_opc::constants::relationship_type::CORE_PROPERTIES
                })
                .expect("core-properties relationship")
                .target_ref(),
            "docProps/CORE.xml"
        );
    }
}
