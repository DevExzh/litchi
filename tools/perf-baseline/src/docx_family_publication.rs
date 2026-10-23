//! Publication-shaped helpers over the DOCX source-backed families (record
//! 0355a).
//!
//! The library no longer has stream publishers: a family commit returns a
//! new source-backed package, and writing that package is the publication.
//! These harness helpers keep the measured work identical (one commit that
//! validates, then one streaming write of the same splice) while the routes
//! keep their publication-shaped evidence. Nothing here buffers the output.

use std::io::Write;
use std::path::Path;

use litchi_docx::source_backed::paragraph_copy::raw::{
    Error as CopyError, Patch as CopyRecordPatch,
};
use litchi_docx::source_backed::tail_append::{
    CandidateProof as TailAppendCandidateProof, Error as TailAppendError, Plan as TailAppendPlan,
};
use litchi_docx::source_backed::tail_append_stream::patch::ParagraphStreamPatch;
use litchi_docx::source_backed::tail_append_stream::{Error as StreamError, ParagraphStreamPlan};
use litchi_docx::source_backed::{self, paragraph_stream, plain_paragraph_copy, tail_append};
use litchi_opc::SourceArtifactFingerprint;

/// A written plain-paragraph copy commit.
pub(crate) struct CopyPublication {
    commit: plain_paragraph_copy::Commit,
}

impl std::fmt::Debug for CopyPublication {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("CopyPublication")
            .finish_non_exhaustive()
    }
}

impl CopyPublication {
    /// The durable inverse bound to the written artifact.
    pub(crate) fn inverse_patch(&self) -> Result<CopyRecordPatch, CopyError> {
        self.commit.patch().inverse().durable()
    }
}

/// Plain-paragraph copy publication through the family.
pub(crate) trait CopyPublish {
    /// Commit a plain-paragraph copy family commit onto this package and write it.
    fn publish_plain_paragraph_copy_to_stream(
        &self,
        writer: impl Write,
        commit: &plain_paragraph_copy::Commit,
    ) -> Result<CopyPublication, CopyError>;

    /// Replay a durable record patch onto this package and write it.
    fn publish_plain_paragraph_copy_patch_to_stream(
        &self,
        writer: impl Write,
        patch: &CopyRecordPatch,
    ) -> Result<CopyPublication, CopyError>;

    /// Restore the source artifact of `publication` from this package (its
    /// written output, reopened) and write it.
    fn publish_plain_paragraph_copy_inverse_to_stream(
        &self,
        writer: impl Write,
        publication: &CopyPublication,
    ) -> Result<(), CopyError>;
}

impl CopyPublish for source_backed::Package {
    fn publish_plain_paragraph_copy_to_stream(
        &self,
        writer: impl Write,
        commit: &plain_paragraph_copy::Commit,
    ) -> Result<CopyPublication, CopyError> {
        let committed = self.apply_plain_paragraph_copy(commit.patch())?;
        committed.package().write_to(writer)?;
        Ok(CopyPublication { commit: committed })
    }

    fn publish_plain_paragraph_copy_patch_to_stream(
        &self,
        writer: impl Write,
        patch: &CopyRecordPatch,
    ) -> Result<CopyPublication, CopyError> {
        let commit = self.apply_plain_paragraph_copy(&patch.clone().into())?;
        commit.package().write_to(writer)?;
        Ok(CopyPublication { commit })
    }

    fn publish_plain_paragraph_copy_inverse_to_stream(
        &self,
        writer: impl Write,
        publication: &CopyPublication,
    ) -> Result<(), CopyError> {
        self.apply_plain_paragraph_copy(&publication.commit.patch().inverse())?
            .package()
            .write_to(writer)?;
        Ok(())
    }
}

/// A written in-place tail-append commit.
pub(crate) struct TailPublication {
    commit: tail_append::Commit,
    candidate: TailAppendCandidateProof,
}

impl std::fmt::Debug for TailPublication {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("TailPublication")
            .finish_non_exhaustive()
    }
}

impl TailPublication {
    /// Whether the append was the exact no-op.
    pub(crate) fn is_noop(&self) -> bool {
        self.commit.patch().is_noop()
    }

    /// The candidate proof of the committed plan.
    pub(crate) const fn candidate_proof(&self) -> TailAppendCandidateProof {
        self.candidate
    }

    /// Restore the source artifact from `current` (the written output,
    /// reopened, or the commit's own package) and write it.
    pub(crate) fn write_inverse_to_stream(
        &self,
        current: &source_backed::Package,
        writer: impl Write,
    ) -> Result<(), TailAppendError> {
        current
            .apply_tail_append(&self.commit.patch().inverse())?
            .package()
            .write_to(writer)
            .map_err(TailAppendError::Document)
    }
}

/// In-place tail-append publication of a prepared plan.
pub(crate) trait TailPlanPublish {
    /// Commit the plan and write the committed package.
    fn write_to(self, writer: impl Write) -> Result<TailPublication, TailAppendError>;
}

impl TailPlanPublish for TailAppendPlan<'_> {
    fn write_to(self, writer: impl Write) -> Result<TailPublication, TailAppendError> {
        let candidate = self.candidate_proof();
        let commit = self.commit()?;
        commit
            .package()
            .write_to(writer)
            .map_err(TailAppendError::Document)?;
        Ok(TailPublication { commit, candidate })
    }
}

/// A written paragraph-stream commit.
pub(crate) struct StreamPublication {
    commit: paragraph_stream::Commit,
}

impl std::fmt::Debug for StreamPublication {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("StreamPublication")
            .finish_non_exhaustive()
    }
}

impl StreamPublication {
    /// The durable forward patch, with the candidate identity measured by
    /// the write.
    pub(crate) fn durable_patch(&self) -> Result<ParagraphStreamPatch, StreamError> {
        self.commit.patch().durable()
    }

    /// The written archive length.
    pub(crate) fn candidate_artifact_length(&self) -> Result<u64, StreamError> {
        Ok(self.durable_patch()?.candidate().archive.len)
    }

    /// The written archive SHA-256.
    pub(crate) fn candidate_artifact_fingerprint(
        &self,
    ) -> Result<SourceArtifactFingerprint, StreamError> {
        Ok(SourceArtifactFingerprint::from_sha256(
            self.durable_patch()?.candidate().archive.sha256,
        ))
    }

    /// Restore the source artifact from `current` and write it.
    pub(crate) fn write_inverse_to_stream(
        &self,
        current: &source_backed::Package,
        writer: impl Write,
    ) -> Result<(), StreamError> {
        current
            .apply_paragraph_stream(&self.commit.patch().inverse())?
            .package()
            .write_to(writer)?;
        Ok(())
    }
}

/// Paragraph-stream publication of a prepared plan.
pub(crate) trait StreamPlanPublish {
    /// Commit the plan and write the committed package.
    fn write_to(self, writer: impl Write) -> Result<StreamPublication, StreamError>;
    /// Commit the plan and save the committed package atomically.
    fn write_to_path(self, path: impl AsRef<Path>) -> Result<StreamPublication, StreamError>;
}

impl StreamPlanPublish for ParagraphStreamPlan<'_> {
    fn write_to(self, writer: impl Write) -> Result<StreamPublication, StreamError> {
        let commit = self.commit()?;
        commit.package().write_to(writer)?;
        Ok(StreamPublication { commit })
    }

    fn write_to_path(self, path: impl AsRef<Path>) -> Result<StreamPublication, StreamError> {
        let commit = self.commit()?;
        commit.package().save(path)?;
        Ok(StreamPublication { commit })
    }
}
