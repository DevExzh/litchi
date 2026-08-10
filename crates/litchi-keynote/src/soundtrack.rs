//! Archive-free Keynote soundtrack playback values and exact transactions.
//!
//! [`Mode`] and [`Settings`] are semantic values validated with [`crate::Error`].
//! The transaction [`Error`] below describes exact package selection and
//! publication failures without exposing native identifiers or bytes.

use std::fmt;

use litchi_iwa_archive::package::ExactArtifacts;
use thiserror::Error as ThisError;

use crate::Package;

const PLAY_ONCE_MODE: i32 = 0;
const LOOP_MODE: i32 = 1;
const DO_NOT_PLAY_MODE: i32 = 2;

/// How a presentation plays its existing soundtrack.
///
/// [`Self::Unknown`] preserves a genuinely future mode value so that reading
/// a newer presentation does not discard it. [`Settings::validate`] rejects a
/// known mode represented as `Unknown`, keeping one canonical semantic spelling
/// for each currently recognized mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Mode {
    /// Play the soundtrack once.
    PlayOnce,
    /// Repeat the soundtrack.
    Loop,
    /// Do not play the soundtrack.
    DoNotPlay,
    /// A mode introduced by a future Keynote release.
    ///
    /// Known values must use their named variants; use this only for a value
    /// that is not currently recognized.
    Unknown(i32),
}

impl Mode {
    /// Decode one stored soundtrack-mode value losslessly.
    ///
    /// This is a constant-time conversion and accepts every `i32` value.
    #[must_use]
    pub const fn from_raw(value: i32) -> Self {
        match value {
            PLAY_ONCE_MODE => Self::PlayOnce,
            LOOP_MODE => Self::Loop,
            DO_NOT_PLAY_MODE => Self::DoNotPlay,
            other => Self::Unknown(other),
        }
    }

    /// Return the stored soundtrack-mode value.
    ///
    /// This is a constant-time conversion. It is primarily useful to semantic
    /// adapters; ordinary callers should use the named variants.
    #[must_use]
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::PlayOnce => PLAY_ONCE_MODE,
            Self::Loop => LOOP_MODE,
            Self::DoNotPlay => DO_NOT_PLAY_MODE,
            Self::Unknown(value) => value,
        }
    }

    /// Return whether this value uses a named variant for every known value.
    ///
    /// This is a constant-time validation predicate.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(PLAY_ONCE_MODE | LOOP_MODE | DO_NOT_PLAY_MODE)
        )
    }
}

/// Validated playback settings for an existing presentation soundtrack.
///
/// Each optional field preserves presence: `None` means the setting is absent,
/// not zero volume or a named playback mode. `Some` explicitly supplies a
/// checked value.
///
/// Soundtrack media is deliberately absent from this value. The focused
/// settings transaction changes playback settings only; it preserves the
/// presentation's existing soundtrack media collection and does not create,
/// delete, replace, or expose media resources.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct Settings {
    /// Playback volume in the inclusive range `0.0..=1.0` when present.
    volume: Option<f64>,
    /// Playback mode when present.
    mode: Option<Mode>,
}

impl Settings {
    /// Construct playback settings from optional semantic values.
    ///
    /// Values are checked before the settings are returned, so a safe caller
    /// cannot publish an out-of-range volume or a known mode disguised as an
    /// unknown value through this semantic type. `None` preserves absence.
    ///
    /// # Costs
    ///
    /// Performs constant-time validation and no allocation.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::NonFiniteSoundtrackVolume`] or
    /// [`crate::Error::SoundtrackVolumeOutOfRange`] for an invalid volume, or
    /// [`crate::Error::NonCanonicalMode`] for a known value represented as
    /// [`Mode::Unknown`].
    pub const fn new(volume: Option<f64>, mode: Option<Mode>) -> crate::Result<Self> {
        let settings = Self { volume, mode };
        match settings.validate() {
            Ok(()) => Ok(settings),
            Err(error) => Err(error),
        }
    }

    /// Return the optional playback volume.
    ///
    /// `None` means the volume setting is absent. This is a constant-time
    /// accessor.
    #[must_use]
    pub const fn volume(self) -> Option<f64> {
        self.volume
    }

    /// Replace or clear the playback volume after validating it.
    ///
    /// Passing `None` clears the setting to absence; it does not set the
    /// volume to zero.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::NonFiniteSoundtrackVolume`] or
    /// [`crate::Error::SoundtrackVolumeOutOfRange`] when
    /// `volume` is outside the semantic volume domain.
    ///
    /// # Costs
    ///
    /// Performs constant-time validation and no allocation.
    pub const fn set_volume(&mut self, volume: Option<f64>) -> crate::Result<()> {
        if let Err(error) = validate_volume(volume) {
            return Err(error);
        }
        self.volume = volume;
        Ok(())
    }

    /// Return the optional playback mode.
    ///
    /// `None` means the mode setting is absent. This is a constant-time
    /// accessor.
    #[must_use]
    pub const fn mode(self) -> Option<Mode> {
        self.mode
    }

    /// Replace or clear the playback mode after validating it.
    ///
    /// Passing `None` clears the setting to absence. A genuinely future
    /// [`Mode::Unknown`] value is retained, while a known value must use its
    /// named variant.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::NonCanonicalMode`] when a known value is passed
    /// through [`Mode::Unknown`].
    ///
    /// # Costs
    ///
    /// Performs constant-time validation and no allocation.
    pub const fn set_mode(&mut self, mode: Option<Mode>) -> crate::Result<()> {
        if let Some(candidate_mode) = mode
            && !candidate_mode.is_canonical()
        {
            return Err(crate::Error::NonCanonicalMode);
        }
        self.mode = mode;
        Ok(())
    }

    /// Validate values before they cross into a package transaction.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::NonFiniteSoundtrackVolume`] or
    /// [`crate::Error::SoundtrackVolumeOutOfRange`] for an invalid volume, and
    /// [`crate::Error::NonCanonicalMode`] when a known native
    /// value is wrapped in [`Mode::Unknown`].
    ///
    /// # Costs
    ///
    /// Performs constant-time validation and no allocation.
    pub const fn validate(self) -> crate::Result<()> {
        if let Err(error) = validate_volume(self.volume) {
            return Err(error);
        }
        if let Some(mode) = self.mode
            && !mode.is_canonical()
        {
            return Err(crate::Error::NonCanonicalMode);
        }
        Ok(())
    }
}

/// A finite resource governed by a soundtrack-settings transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package or payload bytes.
    OutputBytes,
    /// Package entries and encoded records.
    Entries,
    /// Bytes in one package entry or encoded record.
    EntryBytes,
    /// Aggregate package bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic relationships.
    References,
    /// Semantic text storage.
    TextStorages,
    /// Semantic text fragments.
    TextFragments,
    /// Aggregate retained semantic text.
    TextBytes,
    /// Bytes in one encoded payload.
    WireBytes,
    /// Parsed encoded fields.
    WireFields,
    /// Encoded-data nesting depth.
    WireNesting,
    /// Aggregate traversal and update work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
        })
    }
}

/// A focused soundtrack-settings transaction failed.
///
/// Errors expose only semantic categories and bounded measurements. Media
/// values and implementation details stay private.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// The source cannot support a changed exact-source transaction.
    #[error("this Keynote source does not support physical soundtrack-settings edits")]
    UnsupportedSource,
    /// The presentation has no existing soundtrack settings to edit.
    #[error("the Keynote presentation has no soundtrack")]
    SoundtrackNotFound,
    /// The selected soundtrack settings cannot safely support the request.
    #[error("the Keynote soundtrack source cannot be edited safely")]
    InvalidSource,
    /// A retained resource ceiling was exceeded.
    #[error(
        "Keynote soundtrack-settings {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote soundtrack-settings transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Candidate readback did not reproduce the requested semantic result.
    #[error("the edited Keynote soundtrack settings failed verification")]
    Verification,
    /// The patch does not belong to this exact immutable package snapshot.
    #[error("the Keynote soundtrack-settings patch does not match the exact source package")]
    PatchConflict,
}

/// A soundtrack-settings value staged against one immutable package.
///
/// The edit owns no media resource. It can change only the optional playback
/// values already represented by [`Settings`].
pub struct Edit<'a> {
    pub(crate) source: &'a Package,
    pub(crate) before: Settings,
    pub(crate) settings: Settings,
    pub(crate) prepared: crate::package::soundtrack_settings::Prepared<'a>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("volume_present", &self.settings.volume().is_some())
            .field("mode_present", &self.settings.mode().is_some())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the value that would be published by this edit.
    ///
    /// This is a constant-time accessor. `None` in either field remains an
    /// absent setting, not an implicit default.
    #[must_use]
    pub const fn settings(&self) -> Settings {
        self.settings
    }

    /// Replace the staged value.
    ///
    /// This consumes the edit, is allocation-free, and accepts only validated
    /// [`Settings`]. It changes no package or media resource until
    /// [`Self::commit`] succeeds.
    #[must_use]
    pub fn set(mut self, settings: Settings) -> Self {
        self.settings = settings;
        self
    }

    /// Publish the staged value after exact candidate verification.
    ///
    /// # Costs
    ///
    /// A semantic no-op returns the existing immutable package snapshot without
    /// publication or candidate reopen. A change updates the focused soundtrack
    /// settings once and reopens the complete candidate once. Soundtrack media
    /// and rendered previews are retained.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnsupportedSource`] when changed exact-source
    /// publication is unavailable, or a typed source, resource, allocation,
    /// or verification error. Failures do not publish a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        crate::package::soundtrack_settings::commit(self)
    }
}

/// An exact-source-checked reversible soundtrack-settings patch.
///
/// The complete source and target package snapshots are retained privately.
/// Clone and inversion are `O(1)` shared-handle operations. The patch is
/// process-local, not a durable or compact serialization format.
#[derive(Clone, PartialEq)]
pub struct Patch {
    pub(crate) artifacts: ExactArtifacts,
    pub(crate) before: Settings,
    pub(crate) after: Settings,
    pub(crate) touched_components: usize,
    pub(crate) source_reopen_work: usize,
    pub(crate) target_reopen_work: usize,
    pub(crate) source_reopen_references: usize,
    pub(crate) target_reopen_references: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before_volume_present", &self.before.volume().is_some())
            .field("before_mode_present", &self.before.mode().is_some())
            .field("after_volume_present", &self.after.volume().is_some())
            .field("after_mode_present", &self.after.mode().is_some())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the semantic value required from the source.
    #[must_use]
    pub const fn before(&self) -> Settings {
        self.before
    }

    /// Return the semantic value produced by the target.
    #[must_use]
    pub const fn after(&self) -> Settings {
        self.after
    }

    /// Return the source package snapshot's diagnostic fingerprint.
    ///
    /// This is diagnostic evidence only; it neither identifies a package
    /// durably nor authorizes patch application.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target package snapshot's diagnostic fingerprint.
    ///
    /// This is diagnostic evidence only; it neither identifies a package
    /// durably nor authorizes patch application.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether semantic settings and retained package snapshots are both unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return an exact target-to-source inverse in `O(1)` shared-handle work.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            before: self.after,
            after: self.before,
            touched_components: self.touched_components,
            source_reopen_work: self.target_reopen_work,
            target_reopen_work: self.source_reopen_work,
            source_reopen_references: self.target_reopen_references,
            target_reopen_references: self.source_reopen_references,
        }
    }
}

/// Compact evidence about one soundtrack-settings publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    pub(crate) changed: bool,
    pub(crate) touched_components: usize,
    pub(crate) full_reparse_performed: bool,
}

impl Diagnostics {
    /// Return whether exact package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten underlying components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether a changed candidate was reopened for verification.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The verified result of one immutable soundtrack-settings transaction.
#[must_use = "a soundtrack-settings commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    pub(crate) package: Package,
    pub(crate) patch: Patch,
    pub(crate) diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the verified immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

const fn validate_volume(candidate: Option<f64>) -> crate::Result<()> {
    let Some(volume) = candidate else {
        return Ok(());
    };
    if !volume.is_finite() {
        return Err(crate::Error::NonFiniteSoundtrackVolume);
    }
    if volume < 0.0 || volume > 1.0 {
        return Err(crate::Error::SoundtrackVolumeOutOfRange);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn modes_map_native_values_losslessly() {
        for raw in [0, 1, 2, 19, -1, i32::MIN, i32::MAX] {
            assert_eq!(Mode::from_raw(raw).as_raw(), raw);
        }
        assert_eq!(Mode::from_raw(0), Mode::PlayOnce);
        assert_eq!(Mode::from_raw(1), Mode::Loop);
        assert_eq!(Mode::from_raw(2), Mode::DoNotPlay);
        assert_eq!(Mode::from_raw(19), Mode::Unknown(19));
    }

    #[test]
    fn known_values_cannot_be_smuggled_as_unknown() {
        for raw in [0, 1, 2] {
            assert!(!Mode::Unknown(raw).is_canonical());
        }
        assert!(Mode::Unknown(-1).is_canonical());
        assert!(Mode::Unknown(i32::MAX).is_canonical());
    }

    #[test]
    fn settings_validate_volume_boundaries_and_modes() {
        for volume in [None, Some(0.0), Some(1.0)] {
            assert!(Settings::new(volume, Some(Mode::PlayOnce)).is_ok());
        }
        for volume in [
            Some(-f64::EPSILON),
            Some(1.0 + f64::EPSILON),
            Some(f64::NAN),
            Some(f64::INFINITY),
            Some(f64::NEG_INFINITY),
        ] {
            assert!(Settings::new(volume, None).is_err());
        }
        assert!(Settings::new(None, Some(Mode::Unknown(19))).is_ok());
        assert_eq!(
            Settings::new(None, Some(Mode::Unknown(1))).map(|_| ()),
            Err(crate::Error::NonCanonicalMode)
        );
    }

    #[test]
    fn setters_validate_before_mutating() {
        let mut settings = Settings::default();
        assert_eq!(settings.set_volume(Some(0.5)), Ok(()));
        assert_eq!(settings.volume(), Some(0.5));
        for volume in [
            Some(-f64::EPSILON),
            Some(1.0 + f64::EPSILON),
            Some(f64::NAN),
            Some(f64::INFINITY),
            Some(f64::NEG_INFINITY),
        ] {
            assert!(settings.set_volume(volume).is_err());
            assert_eq!(settings.volume(), Some(0.5));
        }
        assert_eq!(settings.set_mode(Some(Mode::Loop)), Ok(()));
        assert_eq!(settings.mode(), Some(Mode::Loop));
        assert_eq!(
            settings.set_mode(Some(Mode::Unknown(2))),
            Err(crate::Error::NonCanonicalMode)
        );
        assert_eq!(settings.mode(), Some(Mode::Loop));
    }
}
