//! Images, movies and audio clips anchored in the Pages body.
//!
//! [`MediaSelector`] picks one by kind and text-flow position;
//! [`crate::Package::edit_body_media_properties`] stages frame, property and
//! playback changes through [`properties::Edit`].
//!
//! ```no_run
//! use litchi_pages::drawable::FlipAxis;
//! use litchi_pages::{MediaSelector, Package};
//!
//! # fn edit(source: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
//! let package = Package::from_bytes(source)?;
//! let mut edit = package.edit_body_media_properties(MediaSelector::image(0))?;
//! edit.flip(FlipAxis::Horizontal)?;
//! let commit = edit.commit()?;
//! let restored = commit
//!     .package()
//!     .apply_body_media_properties(&commit.patch().inverse())?;
//! # let _ = restored;
//! # Ok(())
//! # }
//! ```

pub use crate::selector::MediaSelector;
pub use litchi_iwa_common::media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};

/// Exact-source frame, property and playback transactions for body media.
pub mod properties {
    pub use crate::package::body_media_properties::{
        Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal, Snapshot,
    };
}
