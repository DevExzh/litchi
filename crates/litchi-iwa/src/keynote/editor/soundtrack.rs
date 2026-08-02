//! Typed playback settings for a Keynote presentation soundtrack.

use super::soundtrack_wire::{
    decode_soundtrack, patch_soundtrack_wire, read_soundtrack, replace_soundtrack_message,
};
use super::*;

const PLAY_ONCE_MODE: i32 = 0;
const LOOP_MODE: i32 = 1;
const DO_NOT_PLAY_MODE: i32 = 2;

/// How Keynote plays the presentation soundtrack.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteSoundtrackMode {
    PlayOnce,
    Loop,
    DoNotPlay,
    /// A mode introduced by a newer Keynote version.
    Unknown(i32),
}

impl KeynoteSoundtrackMode {
    /// Decode a native `KN.Soundtrack.SoundtrackMode` value losslessly.
    pub const fn from_raw(value: i32) -> Self {
        match value {
            PLAY_ONCE_MODE => Self::PlayOnce,
            LOOP_MODE => Self::Loop,
            DO_NOT_PLAY_MODE => Self::DoNotPlay,
            value => Self::Unknown(value),
        }
    }

    /// Return the native `KN.Soundtrack.SoundtrackMode` value.
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::PlayOnce => PLAY_ONCE_MODE,
            Self::Loop => LOOP_MODE,
            Self::DoNotPlay => DO_NOT_PLAY_MODE,
            Self::Unknown(value) => value,
        }
    }

    const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(PLAY_ONCE_MODE | LOOP_MODE | DO_NOT_PLAY_MODE)
        )
    }
}

/// Writable soundtrack playback state plus its read-only media item count.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteSoundtrackSettings {
    /// Native playback volume in the inclusive range `0.0..=1.0`.
    pub volume: Option<f64>,
    pub mode: Option<KeynoteSoundtrackMode>,
    /// Number of audio items currently assigned to the soundtrack.
    ///
    /// [`KeynoteEditor::set_soundtrack_settings`] preserves this collection and
    /// rejects attempts to change the count. Media item CRUD is independent of
    /// playback settings.
    pub media_item_count: usize,
}

impl KeynoteSoundtrackSettings {
    fn from_native(soundtrack: &kn::Soundtrack) -> Self {
        Self {
            volume: soundtrack.volume,
            mode: soundtrack.mode.map(KeynoteSoundtrackMode::from_raw),
            media_item_count: soundtrack.movie_media.len(),
        }
    }
}

impl KeynoteEditor {
    /// Read soundtrack playback state, or `None` for a legacy show without a soundtrack object.
    pub fn soundtrack_settings(&self) -> Result<Option<KeynoteSoundtrackSettings>> {
        let graph = ObjectGraph::read(self.package())?;
        let Some(record) = read_soundtrack(&graph)? else {
            return Ok(None);
        };
        Ok(Some(KeynoteSoundtrackSettings::from_native(&record.native)))
    }

    /// Replace soundtrack mode and volume without changing its media collection.
    pub fn set_soundtrack_settings(&mut self, settings: KeynoteSoundtrackSettings) -> Result<()> {
        validate_settings(&settings)?;
        let graph = ObjectGraph::read(self.package())?;
        let Some(record) = read_soundtrack(&graph)? else {
            return Err(Error::InvalidFormat(
                "Keynote show has no soundtrack object".to_owned(),
            ));
        };
        if record.native.movie_media.len() != settings.media_item_count {
            return Err(Error::ParseError(format!(
                "Keynote soundtrack contains {} media items, not {}",
                record.native.movie_media.len(),
                settings.media_item_count
            )));
        }
        let current = KeynoteSoundtrackSettings::from_native(&record.native);
        if current == settings {
            return Ok(());
        }

        let data = patch_soundtrack_wire(record.data, &record.native, &settings)?;
        let verified_native = decode_soundtrack(&data)?;
        if KeynoteSoundtrackSettings::from_native(&verified_native) != settings {
            return Err(Error::InvalidFormat(
                "Keynote soundtrack wire patch failed validation".to_owned(),
            ));
        }

        let archive_name = graph.archive_name(record.id)?.to_owned();
        let mut staged = self.package().clone();
        staged.update_archive(&archive_name, |archive| {
            replace_soundtrack_message(archive, record.id, data)
        })?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.soundtrack_settings()? != Some(settings) {
            return Err(Error::InvalidFormat(
                "Keynote soundtrack settings failed round-trip validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }
}

fn validate_settings(settings: &KeynoteSoundtrackSettings) -> Result<()> {
    if settings
        .volume
        .is_some_and(|volume| !volume.is_finite() || !(0.0..=1.0).contains(&volume))
    {
        return Err(Error::ParseError(
            "Keynote soundtrack volume must be finite and between zero and one".to_owned(),
        ));
    }
    if settings.mode.is_some_and(|mode| !mode.is_canonical()) {
        return Err(Error::ParseError(
            "Keynote soundtrack mode must use its named variant for known native values".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn soundtrack_modes_map_native_values_losslessly() {
        for (raw, mode) in [
            (0, KeynoteSoundtrackMode::PlayOnce),
            (1, KeynoteSoundtrackMode::Loop),
            (2, KeynoteSoundtrackMode::DoNotPlay),
            (19, KeynoteSoundtrackMode::Unknown(19)),
            (-1, KeynoteSoundtrackMode::Unknown(-1)),
        ] {
            assert_eq!(KeynoteSoundtrackMode::from_raw(raw), mode);
            assert_eq!(mode.as_raw(), raw);
        }
    }
}
