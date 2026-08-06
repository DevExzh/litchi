//! WebVTT track values, codec, and PresentationML relationship lifecycle.

mod codec;
mod model;
mod package;
mod tracks_info;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{CONTENT_TYPE, RELATIONSHIP_TYPE};
pub use model::{
    Block, Caption, CaptionTarget, Cue, CueSetting, CueSettingKind, DisplayLocation, File, Header,
    MediaKey, MediaMetadata, RegionSetting, RegionSettingKind, Target, Track, TracksInfo,
};
pub use package::{apply_media_commit, apply_media_patch, load, load_media, store};
pub use transaction::{Commit, Patch, Revision, Snapshot, Transaction};

pub use tracks_info::{NARRATION_URI, TRACKS_INFO_URI};
