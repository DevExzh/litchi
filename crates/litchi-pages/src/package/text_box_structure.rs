//! The structural toolkit of the Pages text-box families (record 0413).
//!
//! The text-box, annotation, field and comment families port the Numbers
//! sheet-drawable publishers (records 0378 and 0378a). Those run on the
//! shared structural working set `litchi_iwa_table::structure` (record
//! 0379), which every leaf `Package` implements through
//! [`litchi_iwa_table::structure::TableLeaf`]. This module gives the Pages
//! families the same vocabulary the Numbers ones use: the working set over a
//! Pages [`Package`], the exact artifact pair a patch retains, and the macro
//! that declares a family's `LimitKind`, `Patch`, `Diagnostics` and
//! `Commit`.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the toolkit keeps the family macro beside the artifacts it uses"
)]

use std::sync::Arc;

use litchi_iwa_archive::package::{OwnedExactArtifacts, SharedBytes};
use litchi_iwa_core::ArchiveObject;
use litchi_iwa_table::TableSource;

pub(super) use litchi_iwa_table::structure::{
    Fault, FaultLimit, FaultResult, exact_source, map_archive, map_core, map_wire, metadata,
    new_object, references, reserve, unique_message,
};

use litchi_iwa_common::wire::{WireFieldView, WireView};

use super::Package;
use super::body_drawable_graph::{self as graph, Budget, GraphError, ObjectIndex};
use crate::selector::DrawableSelector;

/// The Pages working set: the shared engine over a Pages [`Package`].
pub(super) type Working<'p> = litchi_iwa_table::structure::Working<'p, Package>;
/// One reassembled, reopened Pages candidate.
pub(super) type Published = litchi_iwa_table::structure::Published<Package>;

/// The object `identifier` of `package`.
pub(super) fn object(package: &Package, identifier: u64) -> FaultResult<&ArchiveObject> {
    let resolved = package
        .resolve_object(identifier)
        .map_err(|_| Fault::Invalid)?
        .ok_or(Fault::Invalid)?;
    package
        .catalog()
        .get_index(resolved.component_index)
        .and_then(|component| component.archive().objects.get(resolved.object_index))
        .ok_or(Fault::Invalid)
}

/// The single field `number` of `view`, of wire type `wire_type`.
pub(super) fn unique_field<'a>(
    view: &WireView<'a>,
    number: u32,
    wire_type: u8,
) -> FaultResult<Option<WireFieldView<'a>>> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != wire_type {
            return Err(Fault::Invalid);
        }
        if found.replace(field).is_some() {
            return Err(Fault::Invalid);
        }
    }
    Ok(found)
}

/// A body-drawable proof failure as a structural fault.
pub(super) fn graph_fault(error: GraphError) -> Fault {
    match error {
        GraphError::InvalidSource => Fault::Invalid,
        GraphError::Limit {
            observed, maximum, ..
        } => Fault::Limit {
            kind: FaultLimit::PayloadBytes,
            observed,
            maximum,
        },
        GraphError::Allocation { amount } => Fault::Allocation { amount },
    }
}

/// The text a body text box or shape owns, after the body-drawable
/// ownership proof (`body_drawable_graph::select_drawable`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct TextTarget {
    /// The `TSWP.ShapeInfoArchive` object.
    pub(super) drawable: u64,
    /// Its owned `TSWP.StorageArchive` object.
    pub(super) storage: u64,
}

/// How a selector resolves to owned text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Resolved {
    /// The selected drawable and its storage.
    Text(TextTarget),
    /// No body drawable of the selected kind sits at the position.
    NotFound,
    /// The selected kind owns no text (images, movies and audio).
    NoText,
}

/// Resolve `selector` to the body text box or shape that owns text.
pub(super) fn text_target(package: &Package, selector: DrawableSelector) -> FaultResult<Resolved> {
    if !matches!(
        selector,
        DrawableSelector::Shape(_) | DrawableSelector::TextBox(_)
    ) {
        return Ok(Resolved::NoText);
    }
    let mut budget = Budget::new(package).map_err(graph_fault)?;
    let index = ObjectIndex::new(package, &mut budget).map_err(graph_fault)?;
    let body = graph::resolve_body(&index, &mut budget).map_err(graph_fault)?;
    let Some(selected) =
        graph::select_drawable(&index, &body, selector, &mut budget).map_err(graph_fault)?
    else {
        return Ok(Resolved::NotFound);
    };
    let storage = selected.storage.ok_or(Fault::Invalid)?;
    Ok(Resolved::Text(TextTarget {
        drawable: selected.anchor.drawable.get(),
        storage: storage.get(),
    }))
}

/// The exact source/target artifact pair a structural patch retains.
#[derive(Clone, PartialEq, Eq)]
pub(super) struct Artifacts(OwnedExactArtifacts);

impl Artifacts {
    /// The pair of `source` and `target`.
    pub(super) fn new(source: &Package, target: &Package) -> Self {
        Self(OwnedExactArtifacts::new(owner(source), owner(target)))
    }

    /// The pair of `source` with itself (an exact no-op).
    pub(super) fn unchanged(source: &Package) -> Self {
        let owner = owner(source);
        Self(OwnedExactArtifacts::new(owner.clone(), owner))
    }

    pub(super) fn is_noop(&self) -> bool {
        self.0.is_byte_noop()
    }

    #[must_use]
    pub(super) fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    /// Republish the retained target over `package`, which must be the
    /// retained exact source; `None` when it is not.
    pub(super) fn apply(&self, package: &Package) -> FaultResult<Option<Package>> {
        if !self.0.authorizes_owner(&owner(package)) {
            return Ok(None);
        }
        if self.0.is_byte_noop() {
            return Ok(Some(package.clone()));
        }
        let target = Arc::<[u8]>::from(self.0.target_owner().as_ref());
        litchi_iwa_table::structure::TableLeaf::reopen_candidate(package, target).map(Some)
    }
}

fn owner(package: &Package) -> SharedBytes {
    SharedBytes::from_shared_slice(package.state.source.shared_source())
}

/// The `LimitKind`, `Patch`, `Diagnostics` and `Commit` of one text-box
/// family, and the mapping of [`Fault`] onto its `Error` (which must hold
/// the variants `Refused(Refusal)`, `InvalidSource`, `Limit`, `Allocation`
/// and `Verification`, and a `Refusal` with `UnsupportedSource` and
/// `UnsupportedDependency`).
macro_rules! text_box_family {
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
                    $(Self::$wire_variant(kind) => return ::std::fmt::Display::fmt(kind, formatter),)?
                })
            }
        }

        impl From<$crate::package::text_box_structure::Fault> for Error {
            fn from(fault: $crate::package::text_box_structure::Fault) -> Self {
                use $crate::package::text_box_structure::{Fault, FaultLimit};
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
            artifacts: $crate::package::text_box_structure::Artifacts,
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
        }

        impl Diagnostics {
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
                published: $crate::package::text_box_structure::Published,
            ) -> Self {
                Self {
                    patch: Patch {
                        artifacts: $crate::package::text_box_structure::Artifacts::new(
                            source,
                            &published.package,
                        ),
                    },
                    diagnostics: Diagnostics {
                        changed: true,
                        touched_components: published.touched_components,
                    },
                    package: published.package,
                }
            }

            #[allow(dead_code, reason = "only the families with an exact no-op use it")]
            fn unchanged(source: &$crate::Package) -> Self {
                Self {
                    patch: Patch {
                        artifacts: $crate::package::text_box_structure::Artifacts::unchanged(
                            source,
                        ),
                    },
                    diagnostics: Diagnostics {
                        changed: false,
                        touched_components: 0,
                    },
                    package: source.clone(),
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
            Ok(Commit {
                diagnostics: Diagnostics {
                    changed: !patch.is_noop(),
                    touched_components: 0,
                },
                patch: patch.clone(),
                package: target,
            })
        }
    };
}

pub(super) use text_box_family;
