//! Bounded proofs for the private ZIP copy used by cold filesystem samples.
//!
//! This module is harness-only.  It knows the ZIP32 framing emitted by the
//! fixed OPC/PPTX fixtures and deliberately does not become a package writer
//! or a general ZIP normalization API.  The caller must keep the existing
//! semantic OPC verifier in place; the proofs below add physical framing and
//! raw-member checks around that verifier.

use std::{error::Error, fmt, ops::Range};

use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use soapberry_zip::ZipArchive;

const EOCD_SIGNATURE: &[u8; 4] = b"PK\x05\x06";
const EOCD_FIXED_BYTES: usize = 22;
const EOCD_COMMENT_LENGTH_OFFSET: usize = 20;
const MAX_ZIP32_TAIL_BYTES: usize = u16::MAX as usize + EOCD_FIXED_BYTES;

#[derive(Debug, Clone, PartialEq, Eq)]
struct ProofError(String);

impl fmt::Display for ProofError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl Error for ProofError {}

fn error(message: impl Into<String>) -> ProofError {
    ProofError(message.into())
}

/// Evidence that `aligned` is the exact page-aligned private copy of `base`.
///
/// The transformation is restricted to increasing the existing ZIP EOCD
/// comment length and appending zero bytes to that comment.  All bytes before
/// the two-byte comment-length field, all original comment bytes, the EOCD
/// fixed record, the central directory, and every local member must therefore
/// remain identical.
#[derive(Clone, Debug, Deserialize, Serialize)]
pub(crate) struct AlignmentProof {
    pub page_size_bytes: u64,
    pub eocd_offset: u64,
    pub base_bytes: u64,
    pub aligned_bytes: u64,
    pub padding_bytes: u64,
    pub base_comment_bytes: u64,
    pub aligned_comment_bytes: u64,
    pub base_sha256: String,
    pub aligned_sha256: String,
}

/// Evidence that two cold OPC route oracles have the same ZIP bytes up to the
/// EOCD comment, with the source-backed route retaining the aligned source's
/// comment.  Actual child output hashes are checked against these route
/// oracles by the caller; this proof must never replace those checks.
#[cfg(test)]
#[derive(Clone, Debug, Deserialize, Serialize)]
pub(crate) struct ColdOpcOutputProof {
    pub transform: String,
    pub eocd_offset: u64,
    pub aligned_source_bytes: u64,
    pub eager_output_bytes: u64,
    pub source_output_bytes: u64,
    pub aligned_comment_bytes: u64,
    pub unchanged_member_count: u64,
    pub aligned_source_sha256: String,
    pub eager_output_sha256: String,
    pub source_output_sha256: String,
}

/// Evidence for one cold OPC writer route.  The route-specific output hash is
/// kept separate from `canonical_sha256`: the former is the exact oracle used
/// by `record_sample`, while the latter permits a combined selector to compare
/// eager/source routes after independently proving their EOCD comment policy.
#[derive(Clone, Debug, Deserialize, Serialize)]
pub(crate) struct ColdOpcRouteProof {
    pub route: String,
    pub eocd_offset: u64,
    pub aligned_source_bytes: u64,
    pub output_bytes: u64,
    pub output_comment_bytes: u64,
    pub unchanged_member_count: u64,
    pub aligned_source_sha256: String,
    pub output_sha256: String,
    pub canonical_sha256: String,
}

#[derive(Clone, Debug)]
struct Eocd {
    offset: usize,
    comment: Vec<u8>,
}

#[derive(Clone, Debug)]
struct MemberLayout {
    name: Vec<u8>,
    local_span: Range<usize>,
    central_span: Range<usize>,
}

#[derive(Clone, Debug)]
struct ArchiveLayout {
    eocd: Eocd,
    members: Vec<MemberLayout>,
}

/// Proves the exact ZIP padding transform used by
/// `cold_verified::page_aligned_archive` for the fixed ZIP32 fixtures.
///
/// The function rejects a malformed or ambiguous EOCD tail, ZIP64 framing,
/// changed bytes outside the EOCD comment-length field, a changed original
/// comment, non-zero padding, and a result whose length is not page-aligned.
pub(crate) fn verify_zero_padding(
    base: &[u8],
    aligned: &[u8],
    page_size_bytes: u64,
) -> Result<AlignmentProof, Box<dyn Error>> {
    let page_size = usize::try_from(page_size_bytes)
        .map_err(|_| error("ZIP alignment proof page size does not fit usize"))?;
    if page_size == 0 {
        return Err(error("ZIP alignment proof page size is zero").into());
    }
    if base.is_empty() {
        return Err(error("ZIP alignment proof base archive is empty").into());
    }

    let base_eocd = unique_zip32_eocd(base)?;
    let aligned_eocd = unique_zip32_eocd(aligned)?;
    let expected_padding = (page_size - base.len() % page_size) % page_size;

    if expected_padding == 0 {
        if aligned != base {
            return Err(
                error("page-aligned ZIP with zero required padding changed source bytes").into(),
            );
        }
        return Ok(AlignmentProof {
            page_size_bytes,
            eocd_offset: u64::try_from(base_eocd.offset)?,
            base_bytes: u64::try_from(base.len())?,
            aligned_bytes: u64::try_from(aligned.len())?,
            padding_bytes: 0,
            base_comment_bytes: u64::try_from(base_eocd.comment.len())?,
            aligned_comment_bytes: u64::try_from(aligned_eocd.comment.len())?,
            base_sha256: sha256_hex(base),
            aligned_sha256: sha256_hex(aligned),
        });
    }

    if expected_padding > usize::from(u16::MAX)
        || base_eocd
            .comment
            .len()
            .checked_add(expected_padding)
            .is_none_or(|length| length > usize::from(u16::MAX))
    {
        return Err(
            error("ZIP alignment proof padding does not fit the EOCD comment field").into(),
        );
    }
    let expected_aligned_len = base
        .len()
        .checked_add(expected_padding)
        .ok_or_else(|| error("ZIP alignment proof aligned length overflows usize"))?;
    if aligned.len() != expected_aligned_len || aligned.len() % page_size != 0 {
        return Err(error(
            "ZIP alignment proof aligned length differs from the page padding transform",
        )
        .into());
    }
    if aligned_eocd.offset != base_eocd.offset {
        return Err(error("ZIP alignment proof moved the EOCD record").into());
    }
    let expected_comment_len = base_eocd
        .comment
        .len()
        .checked_add(expected_padding)
        .ok_or_else(|| error("ZIP alignment proof comment length overflows usize"))?;
    if aligned_eocd.comment.len() != expected_comment_len
        || aligned_eocd.comment.get(..base_eocd.comment.len()) != Some(base_eocd.comment.as_slice())
        || !aligned_eocd.comment[base_eocd.comment.len()..]
            .iter()
            .all(|byte| *byte == 0)
    {
        return Err(error(
            "ZIP alignment proof did not preserve the original comment and append zero padding",
        )
        .into());
    }

    let comment_length_offset = base_eocd
        .offset
        .checked_add(EOCD_COMMENT_LENGTH_OFFSET)
        .ok_or_else(|| error("ZIP alignment proof EOCD offset overflows usize"))?;
    let comment_length_end = comment_length_offset
        .checked_add(2)
        .ok_or_else(|| error("ZIP alignment proof EOCD length offset overflows usize"))?;
    if base.get(..comment_length_offset) != aligned.get(..comment_length_offset)
        || base.get(comment_length_end..) != aligned.get(comment_length_end..base.len())
        || aligned
            .get(base.len()..)
            .is_none_or(|suffix| suffix.iter().any(|byte| *byte != 0))
    {
        return Err(error(
            "ZIP alignment proof found bytes changed outside the EOCD comment length or zero suffix",
        )
        .into());
    }

    Ok(AlignmentProof {
        page_size_bytes,
        eocd_offset: u64::try_from(base_eocd.offset)?,
        base_bytes: u64::try_from(base.len())?,
        aligned_bytes: u64::try_from(aligned.len())?,
        padding_bytes: u64::try_from(expected_padding)?,
        base_comment_bytes: u64::try_from(base_eocd.comment.len())?,
        aligned_comment_bytes: u64::try_from(aligned_eocd.comment.len())?,
        base_sha256: sha256_hex(base),
        aligned_sha256: sha256_hex(aligned),
    })
}

/// Proves that the cold eager and source-backed OPC output oracles differ only
/// by the aligned source's EOCD comment.
///
/// `target_name` is excluded only from the raw unchanged-member comparison:
/// the caller's existing semantic `verify_opc_overlay_output` remains
/// responsible for proving the replacement payload and every logical member.
/// Every other member's local header/compressed bytes and central-directory
/// record must remain byte-identical to the aligned source.  The two route
/// outputs must have identical bytes through the EOCD fixed record itself.
#[cfg(test)]
pub(crate) fn verify_cold_opc_outputs(
    aligned_source: &[u8],
    eager_output: &[u8],
    source_output: &[u8],
    target_name: &[u8],
) -> Result<ColdOpcOutputProof, Box<dyn Error>> {
    if target_name.is_empty() {
        return Err(error("cold OPC output proof target name is empty").into());
    }

    let aligned = archive_layout(aligned_source)?;
    let eager = archive_layout(eager_output)?;
    let source = archive_layout(source_output)?;

    if eager.eocd.offset != source.eocd.offset || eager.eocd.offset != aligned.eocd.offset {
        return Err(error("cold OPC output proof moved the known EOCD record").into());
    }
    if !eager.eocd.comment.is_empty() {
        return Err(error("cold OPC eager output unexpectedly retained an EOCD comment").into());
    }
    if source.eocd.comment != aligned.eocd.comment {
        return Err(
            error("cold OPC source output did not preserve the aligned EOCD comment").into(),
        );
    }
    if source_output.len()
        != eager_output
            .len()
            .checked_add(aligned.eocd.comment.len())
            .ok_or_else(|| error("cold OPC output length overflows usize"))?
    {
        return Err(
            error("cold OPC route output lengths differ beyond the aligned EOCD comment").into(),
        );
    }
    let eocd_fixed_prefix_end = eager
        .eocd
        .offset
        .checked_add(EOCD_COMMENT_LENGTH_OFFSET)
        .ok_or_else(|| error("cold OPC EOCD fixed prefix overflows usize"))?;
    if eager_output.get(..eocd_fixed_prefix_end) != source_output.get(..eocd_fixed_prefix_end) {
        return Err(
            error("cold OPC route outputs differ before the EOCD comment length field").into(),
        );
    }

    ensure_same_member_order(
        &aligned.members,
        &source.members,
        "aligned source/source output",
    )?;
    ensure_same_member_order(&eager.members, &source.members, "eager/source output")?;
    let target_count = aligned
        .members
        .iter()
        .filter(|member| member.name.as_slice() == target_name)
        .count();
    if target_count != 1
        || source
            .members
            .iter()
            .filter(|member| member.name.as_slice() == target_name)
            .count()
            != 1
        || eager
            .members
            .iter()
            .filter(|member| member.name.as_slice() == target_name)
            .count()
            != 1
    {
        return Err(
            error("cold OPC output proof requires one target member in every archive").into(),
        );
    }

    let mut unchanged_member_count = 0_u64;
    for ((aligned_member, source_member), eager_member) in aligned
        .members
        .iter()
        .zip(source.members.iter())
        .zip(eager.members.iter())
    {
        if aligned_member.name.as_slice() == target_name {
            continue;
        }
        if aligned_source.get(aligned_member.local_span.clone())
            != source_output.get(source_member.local_span.clone())
            || aligned_source.get(aligned_member.central_span.clone())
                != source_output.get(source_member.central_span.clone())
        {
            return Err(error(format!(
                "cold OPC source output changed raw bytes for unchanged member {:?}",
                String::from_utf8_lossy(&aligned_member.name)
            ))
            .into());
        }
        // The eager/source prefix equality above already covers this route;
        // retaining the explicit span check makes the raw-member invariant
        // visible and guards future callers that relax the prefix check.
        if eager_output.get(eager_member.local_span.clone())
            != source_output.get(source_member.local_span.clone())
            || eager_output.get(eager_member.central_span.clone())
                != source_output.get(source_member.central_span.clone())
        {
            return Err(error(format!(
                "cold OPC route output changed raw bytes for unchanged member {:?}",
                String::from_utf8_lossy(&aligned_member.name)
            ))
            .into());
        }
        unchanged_member_count = unchanged_member_count
            .checked_add(1)
            .ok_or_else(|| error("cold OPC unchanged-member count overflows u64"))?;
    }

    Ok(ColdOpcOutputProof {
        transform: "eocd-comment-only".to_owned(),
        eocd_offset: u64::try_from(aligned.eocd.offset)?,
        aligned_source_bytes: u64::try_from(aligned_source.len())?,
        eager_output_bytes: u64::try_from(eager_output.len())?,
        source_output_bytes: u64::try_from(source_output.len())?,
        aligned_comment_bytes: u64::try_from(aligned.eocd.comment.len())?,
        unchanged_member_count,
        aligned_source_sha256: sha256_hex(aligned_source),
        eager_output_sha256: sha256_hex(eager_output),
        source_output_sha256: sha256_hex(source_output),
    })
}

/// Proves the cold eager writer output against one aligned source artifact.
/// The existing semantic OPC verifier remains responsible for the replacement
/// payload; this helper proves the physical framing and every untouched ZIP
/// member around that semantic check.
pub(crate) fn verify_cold_opc_eager_output(
    aligned_source: &[u8],
    output: &[u8],
    target_name: &[u8],
) -> Result<ColdOpcRouteProof, Box<dyn Error>> {
    verify_cold_opc_route_output(aligned_source, output, target_name, "eager")
}

/// Proves the cold source-backed writer output against one aligned source
/// artifact.  The source-backed route must retain the aligned source's EOCD
/// comment exactly; its route-specific exact hash remains an independent
/// oracle check in the harness.
pub(crate) fn verify_cold_opc_source_output(
    aligned_source: &[u8],
    output: &[u8],
    target_name: &[u8],
) -> Result<ColdOpcRouteProof, Box<dyn Error>> {
    verify_cold_opc_route_output(aligned_source, output, target_name, "source-backed")
}

fn verify_cold_opc_route_output(
    aligned_source: &[u8],
    output: &[u8],
    target_name: &[u8],
    route: &'static str,
) -> Result<ColdOpcRouteProof, Box<dyn Error>> {
    if target_name.is_empty() {
        return Err(error("cold OPC route proof target name is empty").into());
    }
    let aligned = archive_layout(aligned_source)?;
    let output_layout = archive_layout(output)?;
    if output_layout.eocd.offset != aligned.eocd.offset {
        return Err(error("cold OPC route output moved the known EOCD record").into());
    }
    let eocd_fixed_prefix_end = aligned
        .eocd
        .offset
        .checked_add(EOCD_COMMENT_LENGTH_OFFSET)
        .ok_or_else(|| error("cold OPC route EOCD fixed prefix overflows usize"))?;
    if aligned_source.get(aligned.eocd.offset..eocd_fixed_prefix_end)
        != output.get(output_layout.eocd.offset..eocd_fixed_prefix_end)
    {
        return Err(error(
            "cold OPC route output changed EOCD fixed bytes before the comment length",
        )
        .into());
    }
    let source_comment_bytes = aligned.eocd.comment.len();
    match route {
        "eager" => {
            if !output_layout.eocd.comment.is_empty()
                || output.len()
                    != aligned_source
                        .len()
                        .checked_sub(source_comment_bytes)
                        .ok_or_else(|| error("cold OPC eager output comment underflows source"))?
            {
                return Err(error(
                    "cold OPC eager route did not remove exactly the aligned EOCD comment",
                )
                .into());
            }
        },
        "source-backed" => {
            if output_layout.eocd.comment != aligned.eocd.comment
                || output.len() != aligned_source.len()
            {
                return Err(error(
                    "cold OPC source-backed route did not preserve the aligned EOCD comment",
                )
                .into());
            }
        },
        _ => return Err(error("cold OPC route proof has an unknown route").into()),
    }
    ensure_same_member_order(
        &aligned.members,
        &output_layout.members,
        "aligned source/route output",
    )?;
    let target_count = aligned
        .members
        .iter()
        .filter(|member| member.name.as_slice() == target_name)
        .count();
    if target_count != 1
        || output_layout
            .members
            .iter()
            .filter(|member| member.name.as_slice() == target_name)
            .count()
            != 1
    {
        return Err(
            error("cold OPC route proof requires one target member in both archives").into(),
        );
    }

    let mut unchanged_member_count = 0_u64;
    for (aligned_member, output_member) in aligned.members.iter().zip(&output_layout.members) {
        if aligned_member.name.as_slice() == target_name {
            continue;
        }
        if aligned_source.get(aligned_member.local_span.clone())
            != output.get(output_member.local_span.clone())
            || aligned_source.get(aligned_member.central_span.clone())
                != output.get(output_member.central_span.clone())
        {
            return Err(error(format!(
                "cold OPC {route} output changed raw bytes for unchanged member {:?}",
                String::from_utf8_lossy(&aligned_member.name)
            ))
            .into());
        }
        unchanged_member_count = unchanged_member_count
            .checked_add(1)
            .ok_or_else(|| error("cold OPC unchanged-member count overflows u64"))?;
    }

    let canonical = canonical_route_bytes(output, output_layout.eocd.offset)?;
    Ok(ColdOpcRouteProof {
        route: route.to_owned(),
        eocd_offset: u64::try_from(aligned.eocd.offset)?,
        aligned_source_bytes: u64::try_from(aligned_source.len())?,
        output_bytes: u64::try_from(output.len())?,
        output_comment_bytes: u64::try_from(output_layout.eocd.comment.len())?,
        unchanged_member_count,
        aligned_source_sha256: sha256_hex(aligned_source),
        output_sha256: sha256_hex(output),
        canonical_sha256: sha256_hex(&canonical),
    })
}

fn canonical_route_bytes(output: &[u8], eocd_offset: usize) -> Result<Vec<u8>, Box<dyn Error>> {
    let end = eocd_offset
        .checked_add(EOCD_FIXED_BYTES)
        .ok_or_else(|| error("cold OPC canonical EOCD range overflows usize"))?;
    let mut canonical = output
        .get(..end)
        .ok_or_else(|| error("cold OPC canonical EOCD range exceeds output"))?
        .to_vec();
    canonical[eocd_offset + EOCD_COMMENT_LENGTH_OFFSET..eocd_offset + EOCD_FIXED_BYTES]
        .copy_from_slice(&[0, 0]);
    Ok(canonical)
}

fn archive_layout(bytes: &[u8]) -> Result<ArchiveLayout, Box<dyn Error>> {
    let eocd = unique_zip32_eocd(bytes)?;
    let archive =
        ZipArchive::from_slice(bytes).map_err(|zip_error| error(zip_error.to_string()))?;
    let mut members = Vec::new();
    for result in archive.entries() {
        let record = result.map_err(|zip_error| error(zip_error.to_string()))?;
        if record.has_data_descriptor() {
            return Err(
                error("cold ZIP proof does not accept data-descriptor member framing").into(),
            );
        }
        let entry = archive
            .get_entry(record.wayfinder())
            .map_err(|zip_error| error(zip_error.to_string()))?;
        let (compressed_start, compressed_end) = entry.compressed_data_range();
        let local_start = usize::try_from(record.local_header_offset())?;
        let local_end = usize::try_from(compressed_end)?;
        let central_start = usize::try_from(record.central_directory_offset())?;
        let compressed_start = usize::try_from(compressed_start)?;
        if local_start > compressed_start
            || compressed_start > local_end
            || local_end > central_start
        {
            return Err(error("cold ZIP proof found an invalid member span").into());
        }
        members.push(MemberLayout {
            name: record.file_path().as_ref().to_vec(),
            local_span: local_start..local_end,
            central_span: central_start..0,
        });
    }
    for index in 0..members.len() {
        let central_end = members
            .get(index + 1)
            .map_or(eocd.offset, |member| member.central_span.start);
        let central_start = members[index].central_span.start;
        if central_end <= central_start || central_end > eocd.offset {
            return Err(error("cold ZIP proof found an invalid central-directory span").into());
        }
        members[index].central_span = central_start..central_end;
    }
    Ok(ArchiveLayout { eocd, members })
}

fn ensure_same_member_order(
    expected: &[MemberLayout],
    actual: &[MemberLayout],
    label: &str,
) -> Result<(), Box<dyn Error>> {
    if expected.len() != actual.len()
        || expected
            .iter()
            .zip(actual)
            .any(|(left, right)| left.name != right.name)
    {
        return Err(error(format!("{label} changed ZIP member order or names")).into());
    }
    Ok(())
}

fn unique_zip32_eocd(bytes: &[u8]) -> Result<Eocd, Box<dyn Error>> {
    let archive =
        ZipArchive::from_slice(bytes).map_err(|zip_error| error(zip_error.to_string()))?;
    if archive.is_zip64() {
        return Err(error("cold ZIP proof requires ZIP32 EOCD framing").into());
    }
    let offset = usize::try_from(archive.eocd_offset())?;
    let comment = archive.comment().as_bytes().to_vec();
    let comment_start = offset
        .checked_add(EOCD_FIXED_BYTES)
        .ok_or_else(|| error("ZIP EOCD comment offset overflows usize"))?;
    let comment_end = comment_start
        .checked_add(comment.len())
        .ok_or_else(|| error("ZIP EOCD comment end overflows usize"))?;
    if comment_end != bytes.len() || archive.end_offset() != u64::try_from(bytes.len())? {
        return Err(error("ZIP EOCD does not terminate at the supplied archive end").into());
    }
    let search_start = bytes.len().saturating_sub(MAX_ZIP32_TAIL_BYTES);
    let mut candidates = Vec::new();
    for (relative, window) in bytes[search_start..].windows(4).enumerate() {
        if window != EOCD_SIGNATURE {
            continue;
        }
        let candidate = search_start
            .checked_add(relative)
            .ok_or_else(|| error("ZIP EOCD candidate offset overflows usize"))?;
        let length_start = candidate
            .checked_add(EOCD_COMMENT_LENGTH_OFFSET)
            .ok_or_else(|| error("ZIP EOCD candidate length offset overflows usize"))?;
        let length_end = length_start
            .checked_add(2)
            .ok_or_else(|| error("ZIP EOCD candidate length end overflows usize"))?;
        let Some(length_bytes) = bytes.get(length_start..length_end) else {
            continue;
        };
        let length = usize::from(u16::from_le_bytes([length_bytes[0], length_bytes[1]]));
        if length_start
            .checked_add(2)
            .and_then(|start| start.checked_add(length))
            == Some(bytes.len())
        {
            candidates.push(candidate);
        }
    }
    if candidates.len() != 1 || candidates[0] != offset {
        return Err(error(format!(
            "ZIP EOCD layout is ambiguous ({} terminal candidates, parser chose {offset})",
            candidates.len()
        ))
        .into());
    }
    Ok(Eocd { offset, comment })
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

#[cfg(test)]
mod tests {
    use super::{
        verify_cold_opc_eager_output, verify_cold_opc_outputs, verify_cold_opc_source_output,
        verify_zero_padding,
    };

    fn put_u16(output: &mut Vec<u8>, value: usize) {
        output.extend_from_slice(&u16::try_from(value).unwrap().to_le_bytes());
    }

    fn put_u32(output: &mut Vec<u8>, value: usize) {
        output.extend_from_slice(&u32::try_from(value).unwrap().to_le_bytes());
    }

    fn stored_zip(entries: &[(&[u8], &[u8])]) -> Vec<u8> {
        let mut output = Vec::new();
        let mut offsets = Vec::with_capacity(entries.len());
        for (name, payload) in entries {
            offsets.push(output.len());
            output.extend_from_slice(b"PK\x03\x04");
            put_u16(&mut output, 20);
            put_u16(&mut output, 0);
            put_u16(&mut output, 0);
            put_u16(&mut output, 0);
            put_u16(&mut output, 0);
            put_u32(&mut output, 0);
            put_u32(&mut output, payload.len());
            put_u32(&mut output, payload.len());
            put_u16(&mut output, name.len());
            put_u16(&mut output, 0);
            output.extend_from_slice(name);
            output.extend_from_slice(payload);
        }
        let central_offset = output.len();
        for ((name, payload), local_offset) in entries.iter().zip(offsets) {
            output.extend_from_slice(b"PK\x01\x02");
            put_u16(&mut output, 20);
            put_u16(&mut output, 20);
            put_u16(&mut output, 0);
            put_u16(&mut output, 0);
            put_u16(&mut output, 0);
            put_u16(&mut output, 0);
            put_u32(&mut output, 0);
            put_u32(&mut output, payload.len());
            put_u32(&mut output, payload.len());
            put_u16(&mut output, name.len());
            put_u16(&mut output, 0);
            put_u16(&mut output, 0);
            put_u16(&mut output, 0);
            put_u16(&mut output, 0);
            put_u32(&mut output, 0);
            put_u32(&mut output, local_offset);
            output.extend_from_slice(name);
        }
        let central_bytes = output.len() - central_offset;
        output.extend_from_slice(b"PK\x05\x06");
        put_u16(&mut output, 0);
        put_u16(&mut output, 0);
        put_u16(&mut output, entries.len());
        put_u16(&mut output, entries.len());
        put_u32(&mut output, central_bytes);
        put_u32(&mut output, central_offset);
        put_u16(&mut output, 0);
        output
    }

    fn aligned_copy(base: &[u8], page_size: usize) -> Vec<u8> {
        let archive = soapberry_zip::ZipArchive::from_slice(base).unwrap();
        let offset = archive.eocd_offset() as usize;
        let comment_len_offset = offset + 20;
        let old_comment_len = archive.comment().as_bytes().len();
        let padding = (page_size - base.len() % page_size) % page_size;
        let mut aligned = base.to_vec();
        let new_comment_len = old_comment_len + padding;
        aligned[comment_len_offset..comment_len_offset + 2]
            .copy_from_slice(&u16::try_from(new_comment_len).unwrap().to_le_bytes());
        aligned.resize(base.len() + padding, 0);
        aligned
    }

    fn with_comment(mut base: Vec<u8>, comment: &[u8]) -> Vec<u8> {
        let archive = soapberry_zip::ZipArchive::from_slice(&base).unwrap();
        let offset = archive.eocd_offset() as usize;
        base[offset + 20..offset + 22]
            .copy_from_slice(&u16::try_from(comment.len()).unwrap().to_le_bytes());
        base.extend_from_slice(comment);
        base
    }

    fn compressed_payload_offset(archive: &[u8], name: &[u8]) -> usize {
        let archive = soapberry_zip::ZipArchive::from_slice(archive).unwrap();
        let record = archive
            .entries()
            .find_map(|result| {
                let record = result.unwrap();
                (record.file_path().as_ref() == name).then_some(record)
            })
            .unwrap();
        archive
            .get_entry(record.wayfinder())
            .unwrap()
            .compressed_data_range()
            .0 as usize
    }

    fn route_oracles() -> (Vec<u8>, Vec<u8>, Vec<u8>) {
        let aligned_source = aligned_copy(
            &stored_zip(&[
                (b"keep.bin", b"keep"),
                (b"target.bin", b"old"),
                (b"tail.bin", b"tail"),
            ]),
            128,
        );
        let eager_output = stored_zip(&[
            (b"keep.bin", b"keep"),
            (b"target.bin", b"new"),
            (b"tail.bin", b"tail"),
        ]);
        let source_output = aligned_copy(&eager_output, 128);
        (aligned_source, eager_output, source_output)
    }

    #[test]
    fn zero_padding_proof_reports_exact_transform() {
        let base = stored_zip(&[(b"a", b"payload")]);
        let aligned = aligned_copy(&base, 128);
        let proof = verify_zero_padding(&base, &aligned, 128).unwrap();
        assert_eq!(proof.base_bytes as usize, base.len());
        assert_eq!(proof.aligned_bytes as usize, aligned.len());
        assert_eq!(proof.aligned_bytes % 128, 0);
        assert_eq!(proof.eocd_offset as usize, base.len() - 22);
        assert_eq!(proof.padding_bytes as usize, aligned.len() - base.len());
        assert_eq!(proof.base_comment_bytes, 0);
    }

    #[test]
    fn zero_padding_proof_preserves_an_existing_comment() {
        let base = with_comment(stored_zip(&[(b"a", b"payload")]), b"old-comment");
        let aligned = aligned_copy(&base, 128);
        let proof = verify_zero_padding(&base, &aligned, 128).unwrap();
        assert_eq!(proof.base_comment_bytes, 11);
        assert_eq!(
            proof.aligned_comment_bytes as usize,
            aligned.len() - (base.len() - 11)
        );
    }

    #[test]
    fn cold_output_proof_accepts_comment_only_route_difference() {
        let (aligned, eager, source) = route_oracles();
        let proof = verify_cold_opc_outputs(&aligned, &eager, &source, b"target.bin").unwrap();
        let eager_route = verify_cold_opc_eager_output(&aligned, &eager, b"target.bin").unwrap();
        let source_route = verify_cold_opc_source_output(&aligned, &source, b"target.bin").unwrap();
        assert_eq!(proof.transform, "eocd-comment-only");
        assert_eq!(proof.unchanged_member_count, 2);
        assert_eq!(
            proof.source_output_bytes,
            proof.eager_output_bytes + proof.aligned_comment_bytes
        );
        assert_eq!(eager_route.route, "eager");
        assert_eq!(source_route.route, "source-backed");
        assert_eq!(eager_route.canonical_sha256, source_route.canonical_sha256);
        assert_eq!(eager_route.output_comment_bytes, 0);
        assert_eq!(source_route.output_bytes as usize, source.len());
        assert_eq!(
            source_route.output_comment_bytes as usize,
            aligned.len() - eager.len()
        );
    }

    #[test]
    fn zero_padding_rejects_changed_member_byte() {
        let base = stored_zip(&[(b"a", b"payload")]);
        let mut aligned = aligned_copy(&base, 128);
        aligned[31] ^= 1;
        assert!(verify_zero_padding(&base, &aligned, 128).is_err());
    }

    #[test]
    fn zero_padding_rejects_changed_eocd_byte() {
        let base = stored_zip(&[(b"a", b"payload")]);
        let mut aligned = aligned_copy(&base, 128);
        let eocd = base.len() - 22;
        aligned[eocd + 8] ^= 1;
        assert!(verify_zero_padding(&base, &aligned, 128).is_err());
    }

    #[test]
    fn zero_padding_rejects_nonzero_suffix() {
        let base = stored_zip(&[(b"a", b"payload")]);
        let mut aligned = aligned_copy(&base, 128);
        *aligned.last_mut().unwrap() = 1;
        assert!(verify_zero_padding(&base, &aligned, 128).is_err());
    }

    #[test]
    fn cold_output_proof_rejects_changed_route_eocd() {
        let (aligned, eager, mut source) = route_oracles();
        let eocd = eager.len() - 22;
        source[eocd + 8] ^= 1;
        assert!(verify_cold_opc_outputs(&aligned, &eager, &source, b"target.bin").is_err());
        assert!(verify_cold_opc_source_output(&aligned, &source, b"target.bin").is_err());
    }

    #[test]
    fn cold_output_proof_rejects_changed_unchanged_member_payload() {
        let (aligned, eager, mut source) = route_oracles();
        let keep_payload = compressed_payload_offset(&source, b"keep.bin");
        source[keep_payload] ^= 1;
        assert!(verify_cold_opc_outputs(&aligned, &eager, &source, b"target.bin").is_err());
        assert!(verify_cold_opc_source_output(&aligned, &source, b"target.bin").is_err());
    }

    #[test]
    fn cold_output_proof_rejects_changed_aligned_member_payload() {
        let (mut aligned, eager, source) = route_oracles();
        let keep_payload = compressed_payload_offset(&aligned, b"keep.bin");
        aligned[keep_payload] ^= 1;
        assert!(verify_cold_opc_outputs(&aligned, &eager, &source, b"target.bin").is_err());
        assert!(verify_cold_opc_eager_output(&aligned, &eager, b"target.bin").is_err());
    }
}
