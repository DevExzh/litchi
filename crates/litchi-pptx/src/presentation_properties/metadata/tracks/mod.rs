//! `WebVTT` track values, codec, and `PresentationML` relationship lifecycle.

mod codec;
mod edit;
mod model;
mod package;
mod tracks_info;
mod transaction;
pub use edit::{Commit, Diagnostics, Edit, EditOptions, Error, LimitKind, Patch};
mod validation;

#[cfg(test)]
mod tests;

pub use codec::{CONTENT_TYPE, Limits, RELATIONSHIP_TYPE};
#[allow(
    clippy::module_name_repetitions,
    reason = "re-exported model types keep their public names"
)]
pub use model::{
    Block, Caption, CaptionTarget, Cue, CueSetting, CueSettingKind, DisplayLocation, File, Header,
    MediaKey, MediaMetadata, RegionSetting, RegionSettingKind, Target, Track, TracksInfo,
};
pub(crate) use package::validate_slide_xml;
pub use package::{load, load_media};
pub use transaction::{Draft, Revision, Snapshot};

pub use tracks_info::{NARRATION_URI, TRACKS_INFO_URI};
