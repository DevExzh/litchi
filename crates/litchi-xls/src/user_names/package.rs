//! Source spans and package-level mutation helpers for `User Names`.

use std::collections::HashSet;
use std::sync::Arc;

use super::codec::{
    C_USR_RECORD_TYPE, RecordSpan, encode_cbusr, encode_cusr, frame_record, parse_stream,
};
use super::model::{Limits, UserEntry, UserGuid, UserNames};
use crate::{Error, Result};

#[derive(Debug, Clone)]
pub(crate) struct Package {
    pub(crate) records: Vec<RecordSpan>,
    pub(crate) model: UserNames,
    pub(crate) limits: Limits,
    pub(crate) revision_guids: Option<Arc<[UserGuid]>>,
    /// Exact source bytes of the Revision Log used to prove the GUID
    /// dependency closure.  Sharing this allocation makes stale checks
    /// collision-free without retaining a second copy per snapshot.
    pub(crate) revision_log_source: Option<Arc<[u8]>>,
}

impl Package {
    pub(crate) fn parse(source: &[u8], limits: Limits) -> Result<Self> {
        let (model, records) = parse_stream(source, limits)?;
        Ok(Self {
            records,
            model,
            limits,
            revision_guids: None,
            revision_log_source: None,
        })
    }

    pub(crate) fn parse_with_revision_guids_and_source(
        source: &[u8],
        limits: Limits,
        revision_guids: Arc<[UserGuid]>,
        revision_log_source: Option<Arc<[u8]>>,
    ) -> Result<Self> {
        let (model, records) = parse_stream(source, limits)?;
        validate_revision_guids(&model, &revision_guids, limits)?;
        Ok(Self {
            records,
            model,
            limits,
            revision_guids: Some(revision_guids),
            revision_log_source,
        })
    }

    pub(crate) fn replace_user(
        &self,
        source: &[u8],
        index: usize,
        user: &UserEntry,
    ) -> Result<Vec<u8>> {
        let span = self.user_span(index)?;
        let payload = user.to_payload()?;
        let replacement = frame_record(super::codec::USR_INFO_RECORD_TYPE, &payload)?;
        let mut candidate = replace_range(
            source,
            span.record_start,
            span.payload_end,
            &replacement,
            self.limits,
        )?;
        let payload_len = u16::try_from(payload.len()).map_err(|_error| {
            Error::UnsafeEdit("UsrInfo payload length does not fit CbUsr".to_string())
        })?;
        let mut sizes = self.model.user_record_sizes;
        sizes[index] = payload_len;
        let cbusr = self.records[2];
        replace_in_place(
            &mut candidate,
            cbusr.record_start,
            cbusr.payload_end,
            &frame_record(super::codec::CB_USR_RECORD_TYPE, &encode_cbusr(&sizes))?,
        )?;
        Ok(candidate)
    }

    pub(crate) fn insert_user(
        &self,
        source: &[u8],
        index: usize,
        user: &UserEntry,
    ) -> Result<Vec<u8>> {
        if index > self.model.users.len() {
            return Err(Error::UnsafeEdit(format!(
                "User Names insertion index {index} is outside {} users",
                self.model.users.len()
            )));
        }
        self.ensure_revision_guid(user)?;
        if self
            .model
            .users
            .iter()
            .any(|current| current.user_id() == user.user_id())
        {
            return Err(Error::UnsafeEdit(format!(
                "User Names already contains lUsrId {}",
                user.user_id()
            )));
        }
        if self.model.users.len() >= self.limits.max_users || self.model.users.len() >= 255 {
            return Err(Error::UnsafeEdit(
                "User Names cannot contain another user under the configured CUsr bound"
                    .to_string(),
            ));
        }
        if self.model.user_record_sizes[255] != 0 {
            return Err(Error::UnsafeEdit(
                "adding a User Names entry would discard a nonzero reserved CbUsr slot".to_string(),
            ));
        }
        let payload = user.to_payload()?;
        let replacement = frame_record(super::codec::USR_INFO_RECORD_TYPE, &payload)?;
        let insert_at = if index == self.model.users.len() {
            source.len()
        } else {
            self.user_span(index)?.record_start
        };

        let mut sizes = self.model.user_record_sizes;
        // Shift both active and ignored slots.  The last slot is the only
        // value that cannot move in the fixed 256-entry field, and it was
        // checked above for the normative zero value.
        for position in (index..255).rev() {
            sizes[position + 1] = sizes[position];
        }
        sizes[index] = u16::try_from(payload.len()).map_err(|_error| {
            Error::UnsafeEdit("UsrInfo payload length does not fit CbUsr".to_string())
        })?;
        let count_payload = encode_cusr(self.model.users.len() + 1)?;
        let cbusr_payload = encode_cbusr(&sizes);
        let mut candidate = insert_at_range(source, insert_at, &replacement, self.limits)?;
        // Header records precede the insertion, so their offsets are stable.
        let cusr = self.records[0];
        let cbusr = self.records[2];
        replace_in_place(
            &mut candidate,
            cusr.record_start,
            cusr.payload_end,
            &frame_record(C_USR_RECORD_TYPE, &count_payload)?,
        )?;
        replace_in_place(
            &mut candidate,
            cbusr.record_start,
            cbusr.payload_end,
            &frame_record(super::codec::CB_USR_RECORD_TYPE, &cbusr_payload)?,
        )?;
        Ok(candidate)
    }

    pub(crate) fn remove_user(&self, source: &[u8], index: usize) -> Result<(Vec<u8>, UserEntry)> {
        let removed = self.model.users.get(index).cloned().ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "User Names index {index} is outside the collection"
            ))
        })?;
        let span = self.user_span(index)?;
        let mut sizes = self.model.user_record_sizes;
        for position in index..self.model.users.len().saturating_sub(1) {
            sizes[position] = sizes[position + 1];
        }
        if let Some(last_active) = self.model.users.len().checked_sub(1) {
            // The newly ignored slot has no UsrInfo owner and must be zero.
            // Existing ignored slots remain untouched, preserving any
            // producer bytes from a malformed or forward-version source.
            sizes[last_active] = 0;
        }
        let count_payload = encode_cusr(self.model.users.len().saturating_sub(1))?;
        let cbusr_payload = encode_cbusr(&sizes);
        let mut candidate = replace_range(
            source,
            span.record_start,
            span.payload_end,
            &[],
            self.limits,
        )?;
        let cusr = self.records[0];
        let cbusr = self.records[2];
        replace_in_place(
            &mut candidate,
            cusr.record_start,
            cusr.payload_end,
            &frame_record(C_USR_RECORD_TYPE, &count_payload)?,
        )?;
        replace_in_place(
            &mut candidate,
            cbusr.record_start,
            cbusr.payload_end,
            &frame_record(super::codec::CB_USR_RECORD_TYPE, &cbusr_payload)?,
        )?;
        Ok((candidate, removed))
    }

    pub(crate) fn model(&self) -> &UserNames {
        &self.model
    }

    fn user_span(&self, index: usize) -> Result<RecordSpan> {
        self.records.get(index + 4).copied().ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "User Names index {index} is outside the collection"
            ))
        })
    }

    fn ensure_revision_guid(&self, user: &UserEntry) -> Result<()> {
        let Some(guids) = &self.revision_guids else {
            return Err(Error::UnsafeEdit(
                "adding a User Names entry requires a Revision Log GUID closure".to_string(),
            ));
        };
        if !user.guid_matches_any(guids) {
            return Err(Error::UnsafeEdit(
                "UsrInfo.guid does not identify a Revision Log header".to_string(),
            ));
        }
        Ok(())
    }
}

fn validate_revision_guids(
    model: &UserNames,
    revision_guids: &[UserGuid],
    limits: Limits,
) -> Result<()> {
    if revision_guids.len() > limits.max_revision_guids {
        return Err(Error::UnsafeEdit(format!(
            "Revision Log contains {} GUIDs; maximum is {}",
            revision_guids.len(),
            limits.max_revision_guids
        )));
    }
    let mut seen = HashSet::new();
    seen.try_reserve(revision_guids.len())
        .map_err(|_error| Error::Allocation("indexing Revision Log GUID closure"))?;
    for guid in revision_guids {
        if !seen.insert(*guid) {
            return Err(Error::InvalidData(
                "Revision Log contains duplicate RRDHead GUIDs".to_string(),
            ));
        }
    }
    for user in &model.users {
        if !seen.contains(user.guid()) {
            return Err(Error::InvalidData(format!(
                "UsrInfo user {} references a GUID absent from Revision Log",
                user.user_id()
            )));
        }
    }
    Ok(())
}

fn replace_range(
    source: &[u8],
    start: usize,
    end: usize,
    replacement: &[u8],
    limits: Limits,
) -> Result<Vec<u8>> {
    if start > end || end > source.len() {
        return Err(Error::UnsafeEdit(
            "User Names replacement range is outside the source stream".to_string(),
        ));
    }
    let output_len = source
        .len()
        .checked_sub(end - start)
        .and_then(|size| size.checked_add(replacement.len()))
        .ok_or(Error::Allocation("sizing User Names replacement"))?;
    if output_len > limits.max_stream_bytes {
        return Err(Error::UnsafeEdit(format!(
            "edited User Names stream has {output_len} bytes; maximum is {}",
            limits.max_stream_bytes
        )));
    }
    let mut candidate = Vec::new();
    candidate
        .try_reserve_exact(output_len)
        .map_err(|_error| Error::Allocation("allocating User Names replacement"))?;
    candidate.extend_from_slice(&source[..start]);
    candidate.extend_from_slice(replacement);
    candidate.extend_from_slice(&source[end..]);
    Ok(candidate)
}

fn insert_at_range(source: &[u8], at: usize, bytes: &[u8], limits: Limits) -> Result<Vec<u8>> {
    replace_range(source, at, at, bytes, limits)
}

fn replace_in_place(source: &mut [u8], start: usize, end: usize, replacement: &[u8]) -> Result<()> {
    if start > end || end > source.len() {
        return Err(Error::UnsafeEdit(
            "User Names header replacement range is outside the candidate".to_string(),
        ));
    }
    if end - start != replacement.len() {
        return Err(Error::UnsafeEdit(
            "User Names header replacement changed a fixed record length".to_string(),
        ));
    }
    source[start..end].copy_from_slice(replacement);
    Ok(())
}
