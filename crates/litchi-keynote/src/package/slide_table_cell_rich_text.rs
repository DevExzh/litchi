//! Shared plumbing of the two slide-table cell rich-text families
//! ([`crate::slide::table::cell_text`] and [`crate::slide::table::cell_list`],
//! record 0412): selector resolution to the table model, the exact-source
//! patch and the publication types. The native graph surgery is the shared
//! `litchi_iwa_table::cell_rich_text` engine that the Numbers cell families
//! run; this module only adapts it to the Keynote selectors and package.

use litchi_iwa_archive::package::SharedBytes;
use litchi_iwa_common::table::coordinate::At;

use super::{Package, slide_table_core as core};
use crate::SlideSelector;
use crate::slide::table::TableSelector;

/// One resolved slide table: its slide and table positions, model and size.
#[derive(Debug, Clone, Copy)]
pub(super) struct Table {
    pub(super) slide: usize,
    pub(super) position: usize,
    pub(super) model: u64,
    pub(super) rows: u32,
    pub(super) columns: u32,
}

impl Table {
    pub(super) const fn contains(self, at: At) -> bool {
        at.row() < self.rows && at.column() < self.columns
    }
}

/// Why a slide-table selector did not resolve.
pub(super) enum Unresolved {
    TableNotFound,
    UnsupportedSource,
    UnsupportedDependency,
    Limit { observed: u64, maximum: u64 },
    Allocation(usize),
    Invalid,
}

fn unresolved(error: core::Error) -> Unresolved {
    match error {
        core::Error::SlideNameNotFound
        | core::Error::SlidePositionNotFound(_)
        | core::Error::TablePositionNotFound(_)
        | core::Error::EmptySlideName
        | core::Error::AmbiguousSelector => Unresolved::TableNotFound,
        core::Error::UnsupportedSource => Unresolved::UnsupportedSource,
        core::Error::UnsupportedDependency | core::Error::UnsupportedTopology => {
            Unresolved::UnsupportedDependency
        },
        core::Error::Limit {
            observed, maximum, ..
        } => Unresolved::Limit { observed, maximum },
        core::Error::Allocation(amount) => Unresolved::Allocation(amount),
        _ => Unresolved::Invalid,
    }
}

/// Resolve the slide and table selectors to the rooted slide table.
pub(super) fn resolve(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
) -> Result<Table, Unresolved> {
    let mut budget = core::Budget::new(package).map_err(unresolved)?;
    let selected = core::select_table(package, slide, table, &mut budget).map_err(unresolved)?;
    Ok(Table {
        slide: selected.slide_position.get(),
        position: selected.table_position.get(),
        model: selected.model.identifier,
        rows: selected.rows,
        columns: selected.columns,
    })
}

/// The exact source owner of `package`; a semantic-only package has none
/// and is represented by an empty owner, which no patch authorizes.
pub(super) fn owner(package: &Package) -> SharedBytes {
    SharedBytes::from_shared_slice(core::physical_catalog(package).map_or_else(
        |_| std::sync::Arc::from(Vec::new()),
        litchi_iwa_archive::SourceCatalog::shared_source,
    ))
}

/// The `LimitKind`, `Error`, `Patch`, `Diagnostics` and `Commit` of one
/// rich-text family, the engine error mapping (through the family's
/// `Refusal::from_engine`) and the selector-error mapping. The family keeps
/// its own `Refusal` and `Edit`.
macro_rules! rich_text_family {
    ($family:literal) => {
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
            /// A bounded wire operation of the shared text engines.
            Wire(litchi_iwa_common::LimitKind),
            /// Resources charged while resolving the slide and table
            /// selectors.
            Selection,
        }

        impl ::std::fmt::Display for LimitKind {
            fn fmt(&self, formatter: &mut ::std::fmt::Formatter<'_>) -> ::std::fmt::Result {
                formatter.write_str(match self {
                    Self::PackageBytes => "package bytes",
                    Self::PayloadBytes => "payload bytes",
                    Self::Identifiers => "object identifiers",
                    Self::Wire(kind) => return ::std::fmt::Display::fmt(kind, formatter),
                    Self::Selection => "table selection",
                })
            }
        }

        #[doc = concat!("A typed, content-redacted ", $family, " failure.")]
        #[derive(Debug, Clone, Copy, PartialEq, Eq, ::thiserror::Error)]
        #[non_exhaustive]
        pub enum Error {
            /// A well-formed request this route refuses; [`Refusal::kind`]
            /// classifies it (record 0351 §2.4).
            #[error(transparent)]
            Refused(Refusal),
            /// The rooted source is malformed or ambiguous.
            #[error("rooted slide-table cell rich-text source is invalid")]
            InvalidSource,
            /// A finite transaction limit was exceeded.
            #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
            Limit {
                /// Resource category.
                kind: LimitKind,
                /// Observed amount.
                observed: u64,
                /// Configured maximum.
                maximum: u64,
            },
            /// A bounded allocation failed.
            #[error("could not allocate {amount} cell rich-text items or bytes")]
            Allocation {
                /// Requested amount.
                amount: usize,
            },
            /// The reopened candidate did not reproduce the requested state.
            #[error("cell rich-text candidate verification failed")]
            Verification,
            /// The patch was applied to an artifact other than its exact
            /// source.
            #[error("cell rich-text patch conflicts with this package")]
            PatchConflict,
        }

        impl From<Refusal> for Error {
            fn from(refusal: Refusal) -> Self {
                Self::Refused(refusal)
            }
        }

        impl From<::litchi_iwa_table::cell_rich_text::Error> for Error {
            fn from(error: ::litchi_iwa_table::cell_rich_text::Error) -> Self {
                use ::litchi_iwa_table::cell_rich_text::{Error as E, LimitKind as K};
                match error {
                    E::Refused(refusal) => Self::Refused(Refusal::from_engine(refusal)),
                    E::Limit {
                        kind,
                        observed,
                        maximum,
                    } => Self::Limit {
                        kind: match kind {
                            K::PackageBytes => LimitKind::PackageBytes,
                            K::Identifiers => LimitKind::Identifiers,
                            K::Wire(kind) => LimitKind::Wire(kind),
                            _ => LimitKind::PayloadBytes,
                        },
                        observed,
                        maximum,
                    },
                    E::Allocation { amount } => Self::Allocation { amount },
                    E::Verification => Self::Verification,
                    _ => Self::InvalidSource,
                }
            }
        }

        impl From<::litchi_iwa_table::structure::Fault> for Error {
            fn from(fault: ::litchi_iwa_table::structure::Fault) -> Self {
                ::litchi_iwa_table::cell_rich_text::Error::from(fault).into()
            }
        }

        impl From<super::slide_table_cell_rich_text::Unresolved> for Error {
            fn from(error: super::slide_table_cell_rich_text::Unresolved) -> Self {
                use super::slide_table_cell_rich_text::Unresolved as U;
                match error {
                    U::TableNotFound => Self::Refused(Refusal::TableNotFound),
                    U::UnsupportedSource => Self::Refused(Refusal::UnsupportedSource),
                    U::UnsupportedDependency => Self::Refused(Refusal::UnsupportedDependency),
                    U::Limit { observed, maximum } => Self::Limit {
                        kind: LimitKind::Selection,
                        observed,
                        maximum,
                    },
                    U::Allocation(amount) => Self::Allocation { amount },
                    U::Invalid => Self::InvalidSource,
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
            artifacts: ::litchi_iwa_archive::package::OwnedExactArtifacts,
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
                self.artifacts.is_byte_noop()
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
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
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
            fn new(source: &$crate::Package, package: $crate::Package, touched: usize) -> Self {
                Self {
                    patch: Patch {
                        artifacts: ::litchi_iwa_archive::package::OwnedExactArtifacts::new(
                            super::slide_table_cell_rich_text::owner(source),
                            super::slide_table_cell_rich_text::owner(&package),
                        ),
                    },
                    diagnostics: Diagnostics {
                        changed: true,
                        touched_components: touched,
                    },
                    package,
                }
            }

            fn unchanged(source: &$crate::Package) -> Self {
                let owner = super::slide_table_cell_rich_text::owner(source);
                Self {
                    patch: Patch {
                        artifacts: ::litchi_iwa_archive::package::OwnedExactArtifacts::new(
                            owner.clone(),
                            owner,
                        ),
                    },
                    diagnostics: Diagnostics::default(),
                    package: source.snapshot(),
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

        /// Republish the patch's retained target over its exact source.
        fn apply_patch(package: &$crate::Package, patch: &Patch) -> Result<Commit, Error> {
            if !patch
                .artifacts
                .authorizes_owner(&super::slide_table_cell_rich_text::owner(package))
            {
                return Err(Error::PatchConflict);
            }
            if patch.is_noop() {
                return Ok(Commit {
                    package: package.snapshot(),
                    patch: patch.clone(),
                    diagnostics: Diagnostics::default(),
                });
            }
            let target = patch.artifacts.target_owner();
            let candidate = super::Package::from_source_with_options(
                ::std::sync::Arc::from(target.as_ref()),
                package.state.options,
            )
            .map_err(|_| Error::Verification)?;
            candidate.validate().map_err(|_| Error::Verification)?;
            Ok(Commit {
                package: candidate,
                patch: patch.clone(),
                diagnostics: Diagnostics {
                    changed: true,
                    touched_components: 0,
                },
            })
        }
    };
}

pub(super) use rich_text_family;
