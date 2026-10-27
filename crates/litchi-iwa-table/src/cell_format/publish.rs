//! Exact-source reassembly of the members a cell-format rewrite returns.
//!
//! The same sequence the Numbers display-format owners run after their
//! native rewrite (record 0379): every changed archive is checked against the
//! preflight bound, Snappy-compressed under the source's retained limits, and
//! spliced into the source ZIP; every other member keeps its exact bytes.
//! The caller reopens the returned bytes with its own package type and
//! verifies the semantic readback.

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_core::SnappyStream;

use super::control::{NativeControlMember, NativePreflight};
use super::{Error, LimitKind, Path, TransactionBudget, map_core_error, physical_source};
use crate::TableSource;

/// Reassemble `members` into a candidate package.
///
/// `members` must name between one and `maximum_members` distinct members.
///
/// # Errors
///
/// [`Error::Limit`] when a member exceeds the preflight bound or the budget,
/// [`Error::Verification`] when reassembly or compression disagrees with its
/// own prediction, and the source's refusal when it was not opened from
/// bytes.
pub fn reassemble(
    source: &dyn TableSource,
    members: &[NativeControlMember],
    maximum_members: usize,
    bound: NativePreflight,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let source_catalog = physical_source(source)?;
    let snappy_limits = source_catalog
        .limits()
        .snappy_limits()
        .map_err(|_| Error::Verification)?;
    let member_count = members.len();
    if !(1..=maximum_members).contains(&member_count) {
        return Err(Error::Verification);
    }
    let mut maximum_compressed_bytes = 0usize;
    let mut compression_allocations = 3usize;
    for member in members {
        if member.archive_bytes.len() > bound.archive_bytes {
            return Err(Error::Limit {
                kind: LimitKind::PayloadBytes,
                observed: u64::try_from(member.archive_bytes.len()).unwrap_or(u64::MAX),
                maximum: u64::try_from(bound.archive_bytes).unwrap_or(u64::MAX),
                path,
            });
        }
        let maximum_compressed =
            SnappyStream::maximum_compressed_len(member.archive_bytes.len(), snappy_limits)
                .map_err(|error| map_core_error(error, path))?;
        if maximum_compressed > bound.compressed_bytes {
            return Err(Error::Limit {
                kind: LimitKind::CompressedBytes,
                observed: u64::try_from(maximum_compressed).unwrap_or(u64::MAX),
                maximum: u64::try_from(bound.compressed_bytes).unwrap_or(u64::MAX),
                path,
            });
        }
        maximum_compressed_bytes = maximum_compressed_bytes
            .checked_add(maximum_compressed)
            .ok_or(Error::Verification)?;
        let frames = member
            .archive_bytes
            .len()
            .checked_add(SnappyStream::WRITE_CHUNK_SIZE - 1)
            .ok_or(Error::Verification)?
            / SnappyStream::WRITE_CHUNK_SIZE;
        compression_allocations = compression_allocations
            .checked_add(frames)
            .and_then(|amount| amount.checked_add(1))
            .ok_or(Error::Verification)?;
    }
    budget.charge_compressed_bytes(maximum_compressed_bytes, path)?;
    budget.charge_allocations(compression_allocations, path)?;
    budget.charge_scratch_bytes(maximum_compressed_bytes, path)?;
    budget.charge_retained_bytes(maximum_compressed_bytes, path)?;
    let compression_work = source_catalog
        .source_bytes()
        .len()
        .checked_mul(2)
        .and_then(|amount| amount.checked_add(maximum_compressed_bytes))
        .ok_or(Error::Verification)?;
    budget.charge_transaction_work(compression_work, path)?;
    let mut compressed_members = Vec::new();
    compressed_members
        .try_reserve_exact(member_count)
        .map_err(|_| Error::Allocation {
            amount: member_count,
            path,
        })?;
    for member in members {
        let compressed = SnappyStream::compress(&member.archive_bytes, snappy_limits)
            .map_err(|error| map_core_error(error, path))?;
        let maximum =
            SnappyStream::maximum_compressed_len(member.archive_bytes.len(), snappy_limits)
                .map_err(|error| map_core_error(error, path))?;
        if compressed.len() > maximum {
            return Err(Error::Verification);
        }
        compressed_members.push(compressed);
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(member_count)
        .map_err(|_| Error::Allocation {
            amount: member_count,
            path,
        })?;
    edits.extend(
        members
            .iter()
            .zip(&compressed_members)
            .map(|(member, bytes)| EntryEdit::new(&member.member_name, bytes)),
    );
    // The reassembly planner owns two small edit-index collections.
    budget.charge_allocations(2, path)?;
    let prepared = source_catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &[], source_catalog.limits())
        .map_err(|_| Error::Verification)?;
    let requirements = prepared.execution_requirements();
    budget.preflight_reassembly(requirements, path)?;
    let publication_work = requirements
        .output_bytes()
        .checked_mul(2)
        .and_then(|amount| {
            source_catalog
                .package()
                .iter()
                .count()
                .checked_mul(1024)
                .and_then(|catalog_work| amount.checked_add(catalog_work))
        })
        .ok_or(Error::Verification)?;
    budget.charge_transaction_work(publication_work, path)?;
    budget.charge_candidate_input_bytes(requirements.output_bytes(), path)?;
    prepared
        .execute(requirements.exact_limits())
        .map_err(|_| Error::Verification)
}
