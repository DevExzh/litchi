//! Shared exact-source machinery of the structural Numbers transactions
//! (record 0301): [`crate::sheet::lifecycle`], [`crate::table::lifecycle`]
//! and [`crate::table::order`].
//!
//! A structural edit stages its changes on a `Working` set: every
//! component it touches is decompressed and parsed once from the exact
//! source, canonical framing is checked, and the edit mutates the parsed
//! objects in place. Objects the edit does not touch keep their exact
//! encoded bytes, and members the edit does not touch keep their exact
//! compressed bytes. `Working::publish` re-encodes only the touched
//! components, rewrites the package metadata registry when objects were
//! added or removed, drops the stale root previews and reassembles the ZIP
//! once. The candidate is reopened through the ordinary strict ingress
//! before any family accepts it, and the family's patch retains both exact
//! artifacts, so its inverse restores the source byte for byte.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the working set keeps its helpers beside the phase that uses them"
)]

// The working set and its phases are the shared table engine
// (`litchi_iwa_table::structure`, record 0379); Numbers keeps its sheet and
// document helpers, the patch artifacts and the family macro here.
pub(super) use litchi_iwa_table::structure::{
    Fault, FaultLimit, FaultResult, allocation, budget, collect, copy_bytes, exact_source, formula,
    map_archive, map_core, map_wire, metadata, new_object, plan, references, reserve, source_owner,
    strip_fields, unique_message,
};

/// The Numbers working set: the shared engine over a Numbers [`Package`].
pub(super) type Working<'p> = litchi_iwa_table::structure::Working<'p, Package>;
/// One reassembled, reopened Numbers candidate.
pub(super) type Published = litchi_iwa_table::structure::Published<Package>;

impl litchi_iwa_table::structure::TableLeaf for Package {
    const CREATES_FIRST_VOLATILE_TILE: bool = true;

    fn reopen_candidate(&self, candidate: Arc<[u8]>) -> FaultResult<Self> {
        Self::reopen_candidate(self, candidate).map_err(reopen_fault)
    }
}

/// Preserve bounded candidate-ingress failures at the shared structural
/// boundary.  The Numbers package error carries richer semantic locations than
/// the shared table engine needs, so non-resource read failures intentionally
/// become verification while resource failures retain their public category.
fn reopen_fault(error: super::Error) -> Fault {
    match error {
        super::Error::Archive(error) => match error {
            litchi_iwa_archive::Error::Limit {
                observed, maximum, ..
            } => Fault::Limit {
                kind: FaultLimit::PackageBytes,
                observed,
                maximum,
            },
            litchi_iwa_archive::Error::Allocation { amount, .. } => Fault::Allocation { amount },
            _ => Fault::Verification,
        },
        super::Error::Common(error) => match error {
            litchi_iwa_common::Error::LimitExceeded {
                observed, limit, ..
            } => Fault::Limit {
                kind: FaultLimit::PayloadBytes,
                observed: observed as u64,
                maximum: limit as u64,
            },
            litchi_iwa_common::Error::Allocation { amount, .. } => Fault::Allocation { amount },
            _ => Fault::Verification,
        },
        super::Error::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Fault::Limit {
            kind: match kind {
                crate::limits::SemanticLimitKind::Objects => FaultLimit::Identifiers,
                _ => FaultLimit::PayloadBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        super::Error::InputTooLarge { observed, maximum } => Fault::Limit {
            kind: FaultLimit::PackageBytes,
            observed,
            maximum,
        },
        _ => Fault::Verification,
    }
}

use std::sync::Arc;

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_core::ArchiveObject;

use super::{
    DOCUMENT_MESSAGE_TYPE, FORM_BASED_SHEET_MESSAGE_TYPE, LEGACY_TABLE_INFO_MESSAGE_TYPE, Package,
    SHEET_MESSAGE_TYPE, TABLE_INFO_MESSAGE_TYPE,
};

/// The member holding the root document object.
pub(super) const DOCUMENT_COMPONENT: &str = "Index/Document.iwa";
/// `TSK.TreeNode`, the sidebar order nodes.
pub(super) const TREE_NODE_MESSAGE_TYPE: u32 = 205;
/// `TN.DocumentArchive.sheets`.
pub(super) const DOCUMENT_SHEETS_FIELD: u32 = 1;
/// `TN.DocumentArchive.sidebar_order`.
pub(super) const DOCUMENT_SIDEBAR_FIELD: u32 = 5;
/// `TSK.TreeNode.children`.
pub(super) const TREE_CHILDREN_FIELD: u32 = 2;
/// `TSK.TreeNode.object`.
pub(super) const TREE_OBJECT_FIELD: u32 = 3;
/// `TN.SheetArchive.name`.
pub(super) const SHEET_NAME_FIELD: u32 = 1;
/// `TN.SheetArchive.drawable_infos`.
pub(super) const SHEET_DRAWABLES_FIELD: u32 = 2;
/// `TN.FormBasedSheetArchive.super`.
pub(super) const FORM_SHEET_SUPER_FIELD: u32 = 1;

/// The sheet message of a sheet object and the field path of its drawable
/// list (`[2]` for a plain sheet, `[1, 2]` for a form-based sheet).
pub(super) fn sheet_message(object: &ArchiveObject) -> FaultResult<(usize, &'static [u32])> {
    let index = unique_message(object, &[SHEET_MESSAGE_TYPE, FORM_BASED_SHEET_MESSAGE_TYPE])?;
    let path: &'static [u32] = if object.messages[index].type_ == SHEET_MESSAGE_TYPE {
        &[SHEET_DRAWABLES_FIELD]
    } else {
        &[FORM_SHEET_SUPER_FIELD, SHEET_DRAWABLES_FIELD]
    };
    Ok((index, path))
}

/// The table-info message of a table drawable.
pub(super) fn table_info_message(object: &ArchiveObject) -> FaultResult<usize> {
    unique_message(
        object,
        &[TABLE_INFO_MESSAGE_TYPE, LEGACY_TABLE_INFO_MESSAGE_TYPE],
    )
}

/// The document message of the root object.
pub(super) fn document_message(object: &ArchiveObject) -> FaultResult<usize> {
    unique_message(object, &[DOCUMENT_MESSAGE_TYPE])
}

/// The exact source/target artifact pair a structural patch retains.
#[derive(Clone, PartialEq, Eq)]
pub(super) struct Artifacts(pub(super) OwnedExactArtifacts);

impl Artifacts {
    /// The pair of `source` and `target`.
    pub(super) fn new(source: &Package, target: &Package) -> FaultResult<Self> {
        Ok(Self(OwnedExactArtifacts::new(
            source_owner(source)?,
            source_owner(target)?,
        )))
    }

    /// The pair of `source` with itself (an exact no-op).
    pub(super) fn unchanged(source: &Package) -> FaultResult<Self> {
        let owner = source_owner(source)?;
        Ok(Self(OwnedExactArtifacts::new(owner.clone(), owner)))
    }

    pub(super) fn is_noop(&self) -> bool {
        self.0.is_byte_noop()
    }

    #[must_use]
    pub(super) fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    /// Republish the retained target over `package`, which must be the
    /// retained exact source.
    ///
    /// Returns `None` when `package` is not the patch's source.
    pub(super) fn apply(&self, package: &Package) -> FaultResult<Option<Package>> {
        if !self.0.authorizes_owner(&source_owner(package)?) {
            return Ok(None);
        }
        if self.0.is_byte_noop() {
            return Ok(Some(package.clone()));
        }
        Package::from_source_owner_with_options(self.0.target_owner(), package.state.options)
            .map(Some)
            .map_err(reopen_fault)
    }
}

/// The `LimitKind`, `Patch`, `Diagnostics` and `Commit` of one structural
/// family, and the mapping of [`Fault`] onto its `Error` (which must hold
/// the shared variants `UnsupportedSource`, `InvalidSource`,
/// `UnsupportedDependency`, `LimitExceeded`, `Allocation` and
/// `Verification`).
macro_rules! structural_family {
    ($family:literal $(, $wire_variant:ident($wire_kind:ty))?) => {
        /// A finite resource governed by the transaction.
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
        #[non_exhaustive]
        pub enum LimitKind {
            /// Complete source or candidate package bytes.
            PackageBytes,
            /// Bytes in one decoded component.
            PayloadBytes,
            /// Native object identifiers.
            Identifiers,
            /// Temporary graph scratch bytes.
            ScratchBytes,
            /// Fallible graph allocation events.
            AllocationEvents,
            $(
                /// A bounded wire operation performed by this family.
                $wire_variant($wire_kind),
            )?
        }

        impl ::std::fmt::Display for LimitKind {
            fn fmt(&self, formatter: &mut ::std::fmt::Formatter<'_>) -> ::std::fmt::Result {
                formatter.write_str(match self {
                    Self::PackageBytes => "package bytes",
                    Self::PayloadBytes => "payload bytes",
                    Self::Identifiers => "object identifiers",
                    Self::ScratchBytes => "scratch bytes",
                    Self::AllocationEvents => "allocation events",
                    $(Self::$wire_variant(kind) => return ::std::fmt::Display::fmt(kind, formatter),)?
                })
            }
        }

        impl From<$crate::package::structure::Fault> for Error {
            fn from(fault: $crate::package::structure::Fault) -> Self {
                use $crate::package::structure::{Fault, FaultLimit};
                match fault {
                    Fault::Unsupported => Self::Refused(Refusal::UnsupportedSource),
                    Fault::Invalid => Self::InvalidSource,
                    Fault::Dependency => Self::Refused(Refusal::UnsupportedDependency),
                    Fault::Limit {
                        kind,
                        observed,
                        maximum,
                    } => Self::Limit {
                        kind: match kind {
                            FaultLimit::PackageBytes => LimitKind::PackageBytes,
                            FaultLimit::PayloadBytes => LimitKind::PayloadBytes,
                            FaultLimit::Identifiers => LimitKind::Identifiers,
                        },
                        observed,
                        maximum,
                    },
                    Fault::Allocation { amount } => Self::Allocation { amount },
                    Fault::Verification => Self::Verification,
                }
            }
        }

        #[doc = concat!("A reversible, exact-source ", $family, " patch.")]
        ///
        /// The patch retains the exact source and target artifacts; it
        /// authorizes application only to its source, and its inverse
        /// restores the source byte for byte.
        #[derive(Clone, PartialEq, Eq)]
        pub struct Patch {
            artifacts: $crate::package::structure::Artifacts,
        }

        impl ::std::fmt::Debug for Patch {
            fn fmt(&self, formatter: &mut ::std::fmt::Formatter<'_>) -> ::std::fmt::Result {
                formatter.debug_struct("Patch").finish_non_exhaustive()
            }
        }

        impl Patch {
            /// Whether the patch is an exact artifact no-op.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.artifacts.is_noop()
            }

            /// The exact target-to-source inverse patch.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self {
                    artifacts: self.artifacts.inverse(),
                }
            }
        }

        #[doc = concat!("Content-free diagnostics for one ", $family, " publication.")]
        #[derive(Debug, Clone, Copy, PartialEq, Eq)]
        pub struct Diagnostics {
            changed: bool,
            touched_components: usize,
            deleted_previews: usize,
        }

        impl Diagnostics {
            const fn unchanged() -> Self {
                Self {
                    changed: false,
                    touched_components: 0,
                    deleted_previews: 0,
                }
            }

            const fn published(published: &$crate::package::structure::Published) -> Self {
                Self {
                    changed: true,
                    touched_components: published.touched_components,
                    deleted_previews: published.deleted_previews,
                }
            }

            /// Whether the exact package artifact changed.
            #[must_use]
            pub const fn changed(self) -> bool {
                self.changed
            }

            /// Number of rewritten native components.
            #[must_use]
            pub const fn touched_components(self) -> usize {
                self.touched_components
            }

            /// Number of stale root previews deleted.
            #[must_use]
            pub const fn deleted_previews(self) -> usize {
                self.deleted_previews
            }
        }

        #[doc = concat!("One validated ", $family, " publication.")]
        #[derive(Debug)]
        pub struct Commit {
            package: $crate::Package,
            patch: Patch,
            diagnostics: Diagnostics,
        }

        impl Commit {
            fn new(
                source: &$crate::Package,
                published: $crate::package::structure::Published,
            ) -> Result<Self, Error> {
                Ok(Self {
                    patch: Patch {
                        artifacts: $crate::package::structure::Artifacts::new(
                            source,
                            &published.package,
                        )?,
                    },
                    diagnostics: Diagnostics::published(&published),
                    package: published.package,
                })
            }

            #[allow(dead_code, reason = "only the families with an exact no-op use it")]
            fn unchanged(source: &$crate::Package) -> Result<Self, Error> {
                Ok(Self {
                    patch: Patch {
                        artifacts: $crate::package::structure::Artifacts::unchanged(source)?,
                    },
                    diagnostics: Diagnostics::unchanged(),
                    package: source.clone(),
                })
            }

            fn applied(package: $crate::Package, patch: &Patch) -> Self {
                Self {
                    diagnostics: if patch.is_noop() {
                        Diagnostics::unchanged()
                    } else {
                        Diagnostics {
                            changed: true,
                            touched_components: 0,
                            deleted_previews: 0,
                        }
                    },
                    patch: patch.clone(),
                    package,
                }
            }

            /// Borrow the validated package snapshot.
            #[must_use]
            pub const fn package(&self) -> &$crate::Package {
                &self.package
            }

            /// Consume this publication and return its package snapshot.
            #[must_use]
            pub fn into_package(self) -> $crate::Package {
                self.package
            }

            /// Consume this publication and return its package and patch.
            #[must_use]
            pub fn into_parts(self) -> ($crate::Package, Patch) {
                (self.package, self.patch)
            }

            /// Borrow the reversible exact-source patch.
            #[must_use]
            pub const fn patch(&self) -> &Patch {
                &self.patch
            }

            /// Borrow content-free diagnostics.
            #[must_use]
            pub const fn diagnostics(&self) -> &Diagnostics {
                &self.diagnostics
            }
        }

        fn apply_patch(package: &$crate::Package, patch: &Patch) -> Result<Commit, Error> {
            let target = patch
                .artifacts
                .apply(package)?
                .ok_or(Error::PatchConflict)?;
            Ok(Commit::applied(target, patch))
        }
    };
}

pub(super) use structural_family;
