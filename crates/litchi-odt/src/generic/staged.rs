//! The staged edit family shared by the format-neutral packaged and flat
//! owners and the standalone chart document (record 0358).

/// The owner contract the staged family needs: a fallible working copy, a
/// reopen of exact bytes under the owner's own policies, and its bytes.
pub(crate) trait Staged: Sized {
    /// An independent copy the staged steps mutate.
    fn working_copy(&self) -> litchi_core::Result<Self>;
    /// Reopen `bytes` under this owner's limits and execution policy.
    fn reopen(&self, bytes: &[u8]) -> litchi_core::Result<Self>;
    /// The exact owned bytes.
    fn staged_bytes(&self) -> &[u8];
}

/// Declare `Edit`, `Patch`, `Commit`, `Diagnostics`, `Error` and
/// `LimitKind` for one owner whose crate-private `&mut self` mutators the
/// commit applies, in staging order, to a working copy.
macro_rules! staged_family {
    ($owner:ty, $label:literal) => {
        /// Whether a commit changed the document bytes.
        pub use litchi_core::edit::Diagnostics;
        /// The family error: typed limits, refusals and patch conflicts plus
        /// the common-core payload.
        pub use litchi_odf_common::Error;
        /// The dimension named by [`Error::Limit`] and [`Error::InvalidLimit`].
        pub use litchi_odf_common::LimitKind;

        /// A committed edit: the published document, its patch and
        /// diagnostics.
        pub type Commit = litchi_core::edit::Commit<$owner, Patch>;

        type Step = std::sync::Arc<dyn Fn(&mut $owner) -> litchi_core::Result<()> + Send + Sync>;

        /// A staged, source-bound edit. Each verb records one change; the
        /// changes apply in order when [`Edit::commit`] runs and report their
        /// errors there. Dropping the edit discards them.
        pub struct Edit<'source> {
            source: &'source $owner,
            steps: Vec<Step>,
        }

        impl std::fmt::Debug for Edit<'_> {
            fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                formatter
                    .debug_struct("Edit")
                    .field("steps", &self.steps.len())
                    .finish_non_exhaustive()
            }
        }

        impl<'source> Edit<'source> {
            pub(crate) fn new(source: &'source $owner) -> Self {
                Self {
                    source,
                    steps: Vec::new(),
                }
            }

            /// The document this edit stages against.
            #[must_use]
            pub fn package(&self) -> &'source $owner {
                self.source
            }

            /// Discards every staged change in place.
            pub fn reset(&mut self) {
                self.steps.clear();
            }

            fn stage(
                &mut self,
                step: impl Fn(&mut $owner) -> litchi_core::Result<()> + Send + Sync + 'static,
            ) -> &mut Self {
                self.steps.push(std::sync::Arc::new(step));
                self
            }

            /// Applies every staged change in order to a working copy and
            /// publishes it.
            ///
            /// An edit with no staged change, or whose changes keep the bytes,
            /// commits an unchanged copy: `diagnostics().changed()` is false
            /// and `patch().is_noop()` holds.
            ///
            /// # Errors
            ///
            /// Returns the first staged change's failure; the source is
            /// unchanged.
            pub fn commit(self) -> std::result::Result<Commit, Error> {
                use $crate::generic::staged::Staged as _;
                let mut working = self.source.working_copy()?;
                for step in &self.steps {
                    step(&mut working)?;
                }
                let before: std::sync::Arc<[u8]> = self.source.staged_bytes().into();
                let changed = working.staged_bytes() != &*before;
                let after: std::sync::Arc<[u8]> = if changed {
                    working.staged_bytes().into()
                } else {
                    std::sync::Arc::clone(&before)
                };
                Ok(Commit::new(
                    working,
                    Patch { before, after },
                    Diagnostics::new(changed),
                ))
            }
        }

        /// An exact-source, reversible whole-document patch.
        #[derive(Clone)]
        pub struct Patch {
            before: std::sync::Arc<[u8]>,
            after: std::sync::Arc<[u8]>,
        }

        impl std::fmt::Debug for Patch {
            fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
                formatter
                    .debug_struct("Patch")
                    .field("noop", &self.is_noop())
                    .finish_non_exhaustive()
            }
        }

        impl Patch {
            /// Whether this patch was made from exactly `source`'s bytes.
            #[must_use]
            pub fn is_applicable_to(&self, source: &$owner) -> bool {
                use $crate::generic::staged::Staged as _;
                source.staged_bytes() == &*self.before
            }

            /// Whether replaying this patch keeps the source bytes.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.before == self.after
            }

            /// The patch that restores the source bytes.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self {
                    before: std::sync::Arc::clone(&self.after),
                    after: std::sync::Arc::clone(&self.before),
                }
            }

            pub(crate) fn replay(&self, source: &$owner) -> std::result::Result<Commit, Error> {
                use $crate::generic::staged::Staged as _;
                if !self.is_applicable_to(source) {
                    return Err(Error::patch_conflict(
                        concat!($label, ".apply"),
                        "the document is not the patch source",
                    ));
                }
                let target = if self.is_noop() {
                    source.working_copy()?
                } else {
                    source.reopen(&self.after)?
                };
                Ok(Commit::new(
                    target,
                    self.clone(),
                    Diagnostics::new(!self.is_noop()),
                ))
            }
        }
    };
}

pub(crate) use staged_family;
