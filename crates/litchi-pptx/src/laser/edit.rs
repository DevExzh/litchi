//! Immutable laser-pointer trace edits of a [`Package`](crate::Package) (record 0356a).
//!
//! [`Package::edit_laser_traces`](crate::Package::edit_laser_traces) starts an [`Edit`]
//! over a staged copy of the package; its verbs run on the staged graph.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_laser_traces`](crate::Package::apply_laser_traces) replays a patch
//! onto an equal package. The source package never changes, an exact no-op
//! keeps every byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]

use super::package as raw;
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("laser-pointer trace");

/// A staged laser-pointer trace edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a laser-pointer trace edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_laser_traces(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_laser_traces")?,
        })
    }

    /// Apply a laser-pointer trace patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_laser_traces(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_laser_traces")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// Store one trace in a slide's extension list, preserving its dialect.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn set<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        points: &[TracePoint],
    ) -> Result<()> {
        let slide_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self.staged.laser_limits();
        self.staged
            .update(move |opc| raw::store_slide_trace(opc, &slide_name, points, limits))
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged changes to a new package.
    ///
    /// # Errors
    ///
    /// Returns the part-count limit of the patch.
    pub fn commit(self) -> Result<Commit> {
        Ok(commit_of(self.staged.commit()?))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::laser::Limits;
    use crate::time::Offset;
    use litchi_drawingml::coordinate::Coordinate;

    fn point() -> TracePoint {
        TracePoint::new(
            Offset::ZERO,
            Coordinate::emu(914_400).unwrap(),
            Coordinate::emu(457_200).unwrap(),
        )
    }

    #[test]
    fn owning_edit_retains_custom_laser_profile() {
        let mut builder = crate::builder::Builder::new().unwrap();
        builder.presentation().unwrap().add_slide().unwrap();
        let bytes = builder.build().unwrap().to_bytes().unwrap();
        let default = Limits::DEFAULT;
        let profile = Limits::new(
            default.max_slide_xml_bytes(),
            default.max_total_slide_xml_bytes(),
            default.max_trace_count(),
            1,
            default.max_xml_nodes(),
            default.max_xml_depth(),
            default.max_output_bytes(),
        )
        .unwrap();
        let package = Package::from_bytes_with_options(
            &bytes,
            crate::OpenOptions::default().laser_limits(profile),
        )
        .unwrap();
        let mut edit = package.edit_laser_traces().unwrap();
        let points = [point(), point()];
        assert!(matches!(
            edit.set(0_usize, &points),
            Err(crate::Error::Limit {
                resource: "laser trace-point count",
                limit: 1,
            })
        ));
    }
}
