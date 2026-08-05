//! WebVTT track values, codec, and PresentationML relationship lifecycle.

mod codec;
mod model;
mod package;

pub use codec::{CONTENT_TYPE, RELATIONSHIP_TYPE};
pub use model::{
    Block, Cue, CueSetting, CueSettingKind, File, Header, RegionSetting, RegionSettingKind, Target,
    Track,
};
pub use package::{load, store};
