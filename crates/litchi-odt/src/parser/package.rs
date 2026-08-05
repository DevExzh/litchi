//! Package/document boundary for ODT-specific parser composition.
//!
//! Each XML codec remains responsible for namespace-aware decoding. This layer
//! combines the declaration and marker passes that form one document view.

use super::codec;
use super::model::{Comment, Section, TrackChange, TrackedChanges};
use litchi_core::Result;

pub(super) fn parse_track_changes(content: &str) -> Result<Vec<TrackChange>> {
    Ok(parse_tracked_changes(content)?.changes)
}

pub(super) fn parse_tracked_changes(content: &str) -> Result<TrackedChanges> {
    let mut tracked = codec::parse_change_declarations(content)?;
    codec::correlate_change_ranges(content, &mut tracked.changes)?;
    Ok(tracked)
}

pub(super) fn parse_comments(content: &str) -> Result<Vec<Comment>> {
    codec::parse_comments(content)
}

pub(super) fn parse_sections(content: &str) -> Result<Vec<Section>> {
    codec::parse_sections(content)
}
