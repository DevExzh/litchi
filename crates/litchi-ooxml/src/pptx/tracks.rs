//! Temporary host boundary for the canonical contextual WebVTT owner.

use crate::error::{OoxmlError, Result};
use litchi_opc::OpcPackage;

pub use litchi_pptx::presentation_properties::metadata::tracks::{
    Block as WebVttBlock, Cue as WebVttCue, CueSetting as WebVttCueSetting,
    CueSettingKind as WebVttCueSettingKind, File as WebVttTrack, Header as WebVttHeader,
    RegionSetting as WebVttRegionSetting, RegionSettingKind as WebVttRegionSettingKind,
    Target as TrackTarget, Track,
};

pub const TRACK_CONTENT_TYPE: &str =
    litchi_pptx::presentation_properties::metadata::tracks::CONTENT_TYPE;
pub const TRACK_RELATIONSHIP_TYPE: &str =
    litchi_pptx::presentation_properties::metadata::tracks::RELATIONSHIP_TYPE;

pub fn load_presentation_tracks(package: &OpcPackage) -> Result<Vec<Track>> {
    litchi_pptx::presentation_properties::metadata::tracks::load(package).map_err(OoxmlError::from)
}

pub fn store_presentation_track(package: &mut OpcPackage, value: &Track) -> Result<()> {
    litchi_pptx::presentation_properties::metadata::tracks::store(package, value)
        .map_err(OoxmlError::from)
}
