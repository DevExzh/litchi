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
use litchi_opc::Part;
use litchi_opc::constants::content_type as ct;

family!("laser-pointer trace");

/// A staged laser-pointer trace edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
    /// Aggregate usage for the slides successfully changed by this staged
    /// operation.  Keeping it beside the staged graph makes a failed verb
    /// retry-safe: the ledger advances only when `Staged::update` publishes
    /// its candidate.
    usage: Usage,
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
            usage: Usage::default(),
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
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_laser_traces",
            |package, target, patch| {
                let limits = package
                    .semantic_limits()
                    .laser()
                    .constrained_by_source(target.read_limits());
                let mut usage = Usage::default();
                patch.validate_changed_parts(target, |_, part| {
                    if part.content_type() == ct::PML_SLIDE {
                        load_slide_traces(0, part, &limits, &mut usage).map(|_| ())
                    } else {
                        Ok(())
                    }
                })
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
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
        let mut next_usage = self.usage;
        let result = self.staged.update(|opc| {
            raw::store_slide_trace(opc, &slide_name, points, limits, &mut next_usage)
        });
        if result.is_ok() {
            self.usage = next_usage;
        }
        result
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
        self.usage = Usage::default();
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
    use litchi_opc::{PackURI, Part};

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
            Err(Error::Limit {
                kind: LimitKind::Codec(crate::opened::CodecLimitKind::TracePoints),
                observed: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn replay_enforces_the_recipient_profile_but_noop_skips_payload_validation() {
        let mut builder = crate::builder::Builder::new().unwrap();
        builder.presentation().unwrap().add_slide().unwrap();
        let bytes = builder.build().unwrap().to_bytes().unwrap();

        let source = Package::from_bytes(&bytes).unwrap();
        let mut edit = source.edit_laser_traces().unwrap();
        edit.set(0_usize, &[point()]).unwrap();
        let patch = edit.commit().unwrap().patch().clone();

        let strict = Limits::new(1, 1, 1, 1, 1, 1, 1).unwrap();
        let recipient = Package::from_bytes_with_options(
            &bytes,
            crate::OpenOptions::default().laser_limits(strict),
        )
        .unwrap();
        let before = recipient.to_bytes().unwrap();
        assert!(matches!(
            recipient.apply_laser_traces(&patch),
            Err(Error::Limit {
                kind: LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
                maximum: 1,
                ..
            })
        ));
        assert_eq!(recipient.to_bytes().unwrap(), before);

        let noop = recipient.edit_laser_traces().unwrap().commit().unwrap();
        assert!(noop.patch().is_noop());
        let replayed = recipient
            .apply_laser_traces(noop.patch())
            .unwrap()
            .into_package();
        assert_eq!(replayed.to_bytes().unwrap(), before);
    }

    #[test]
    fn owning_edit_charges_changed_slides_once_and_rolls_back_failed_set() {
        let mut builder = crate::builder::Builder::new().unwrap();
        builder.presentation().unwrap().add_slide().unwrap();
        builder.presentation().unwrap().add_slide().unwrap();
        let bytes = builder.build().unwrap().to_bytes().unwrap();

        // Derive an exact one-slide output bound from the ordinary profile;
        // the bounded edit below must admit the first changed slide and
        // refuse the second without charging or publishing it.
        let source = Package::from_bytes(&bytes).unwrap();
        let mut baseline_edit = source.edit_laser_traces().unwrap();
        baseline_edit.set(0_usize, &[point()]).unwrap();
        let baseline = baseline_edit.commit().unwrap().into_package();
        let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let one_slide_bytes = baseline
            .opc
            .get_part(&slide_name)
            .unwrap()
            .payload_len()
            .unwrap();
        let default = Limits::DEFAULT;
        let profile = Limits::new(
            one_slide_bytes,
            one_slide_bytes,
            2,
            2,
            default.max_xml_nodes(),
            default.max_xml_depth(),
            default.max_output_bytes(),
        )
        .unwrap();
        let limited = Package::from_bytes_with_options(
            &bytes,
            crate::OpenOptions::default().laser_limits(profile),
        )
        .unwrap();

        let mut edit = limited.edit_laser_traces().unwrap();
        edit.set(0_usize, &[point()]).unwrap();
        let charged = edit.usage;
        assert_eq!(charged.total_slide_xml_bytes(), one_slide_bytes);
        assert_eq!(charged.trace_count(), 1);
        assert_eq!(charged.point_count(), 1);

        let error = edit.set(1_usize, &[point()]).unwrap_err();
        let expected_observed = one_slide_bytes.checked_add(one_slide_bytes).unwrap();
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Codec(crate::opened::CodecLimitKind::XmlBytes),
                observed,
                maximum,
            } if observed == expected_observed && maximum == one_slide_bytes
        ));
        assert_eq!(edit.usage, charged);
        assert!(edit.set(1_usize, &[point()]).is_err());

        edit.reset();
        assert_eq!(edit.usage, Usage::default());
        assert!(edit.commit().unwrap().patch().is_noop());

        let mut retry = limited.edit_laser_traces().unwrap();
        retry.set(0_usize, &[point()]).unwrap();
        assert!(!retry.commit().unwrap().patch().is_noop());
    }
}
