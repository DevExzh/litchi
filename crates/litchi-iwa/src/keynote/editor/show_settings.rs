//! Presentation-level Keynote show settings.

use super::*;

const DOCUMENT_OBJECT_ID: u64 = 1;
const SHOW_ARCHIVE_MESSAGE_TYPE: u32 = 2;
const SHOW_SIZE_FIELD: u32 = 4;
const SIZE_WIDTH_FIELD: u32 = 1;
const SIZE_HEIGHT_FIELD: u32 = 2;
const SLIDE_NUMBERS_VISIBLE_FIELD: u32 = 6;
const LOOP_PRESENTATION_FIELD: u32 = 8;
const MODE_FIELD: u32 = 9;
const AUTOPLAY_TRANSITION_DELAY_FIELD: u32 = 10;
const AUTOPLAY_BUILD_DELAY_FIELD: u32 = 11;
const IDLE_TIMER_ACTIVE_FIELD: u32 = 15;
const IDLE_TIMER_DELAY_FIELD: u32 = 16;
const AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD: u32 = 18;

const NORMAL_MODE: i32 = 0;
const SELF_PLAYING_MODE: i32 = 1;
const LINKS_ONLY_MODE: i32 = 2;

/// How a Keynote presentation advances and responds to input.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteShowMode {
    /// Slides advance normally through presenter input and configured transitions.
    Normal,
    /// The presentation advances automatically using its playback delays.
    SelfPlaying,
    /// Only hyperlinks can navigate between slides.
    LinksOnly,
    /// A mode introduced by a newer Keynote version.
    Unknown(i32),
}

impl KeynoteShowMode {
    /// Decode a native `KNShowMode` value without discarding unknown values.
    pub const fn from_raw(value: i32) -> Self {
        match value {
            NORMAL_MODE => Self::Normal,
            SELF_PLAYING_MODE => Self::SelfPlaying,
            LINKS_ONLY_MODE => Self::LinksOnly,
            value => Self::Unknown(value),
        }
    }

    /// Return the native `KNShowMode` value stored in the Keynote archive.
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::Normal => NORMAL_MODE,
            Self::SelfPlaying => SELF_PLAYING_MODE,
            Self::LinksOnly => LINKS_ONLY_MODE,
            Self::Unknown(value) => value,
        }
    }

    const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(NORMAL_MODE | SELF_PLAYING_MODE | LINKS_ONLY_MODE)
        )
    }
}

/// Writable presentation-level behavior stored in `KN.ShowArchive`.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteShowSettings {
    pub width: f32,
    pub height: f32,
    pub slide_numbers_visible: Option<bool>,
    pub loop_presentation: Option<bool>,
    pub mode: Option<KeynoteShowMode>,
    pub autoplay_transition_delay: Option<f64>,
    pub autoplay_build_delay: Option<f64>,
    pub idle_timer_active: Option<bool>,
    pub idle_timer_delay: Option<f64>,
    pub automatically_plays_upon_open: Option<bool>,
}

impl From<&kn::ShowArchive> for KeynoteShowSettings {
    fn from(show: &kn::ShowArchive) -> Self {
        Self {
            width: show.size.width,
            height: show.size.height,
            slide_numbers_visible: show.slide_numbers_visible,
            loop_presentation: show.loop_presentation,
            mode: show.mode.map(KeynoteShowMode::from_raw),
            autoplay_transition_delay: show.autoplay_transition_delay,
            autoplay_build_delay: show.autoplay_build_delay,
            idle_timer_active: show.idle_timer_active,
            idle_timer_delay: show.idle_timer_delay,
            automatically_plays_upon_open: show.automatically_plays_upon_open,
        }
    }
}

impl KeynoteEditor {
    /// Read presentation-level dimensions and playback behavior.
    pub fn show_settings(&self) -> Result<KeynoteShowSettings> {
        let graph = ObjectGraph::read(self.text.package())?;
        let document: kn::DocumentArchive =
            graph.decode(DOCUMENT_OBJECT_ID, "KN.DocumentArchive")?;
        let show: kn::ShowArchive = graph.decode(document.show.identifier, "KN.ShowArchive")?;
        Ok(KeynoteShowSettings::from(&show))
    }

    /// Replace presentation-level dimensions and playback behavior transactionally.
    pub fn set_show_settings(&mut self, settings: KeynoteShowSettings) -> Result<()> {
        validate_show_settings(&settings)?;
        let graph = ObjectGraph::read(self.text.package())?;
        let document: kn::DocumentArchive =
            graph.decode(DOCUMENT_OBJECT_ID, "KN.DocumentArchive")?;
        let show_id = document.show.identifier;
        let archive_name = graph.archive_name(show_id)?.to_owned();
        let mut staged = self.text.package().clone();
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(show_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote show object {show_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| {
                    message.type_ == SHOW_ARCHIVE_MESSAGE_TYPE
                        && kn::ShowArchive::decode(message.data.as_slice()).is_ok()
                })
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote show object {show_id} has no ShowArchive payload"
                    ))
                })?;
            let original = object.messages[message_index].data.as_slice();
            let show = kn::ShowArchive::decode(original)?;
            let data = patch_show_settings_wire(original, &show, &settings)?;
            let verified = kn::ShowArchive::decode(data.as_slice())?;
            if KeynoteShowSettings::from(&verified) != settings {
                return Err(Error::InvalidFormat(
                    "Keynote show-settings wire patch failed validation".to_owned(),
                ));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: SHOW_ARCHIVE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified.show_settings()? != settings {
            return Err(Error::InvalidFormat(
                "Keynote show settings failed round-trip validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }
}

fn patch_show_settings_wire(
    original: &[u8],
    show: &kn::ShowArchive,
    settings: &KeynoteShowSettings,
) -> Result<Vec<u8>> {
    let mut data = patch_nested_fixed32_field(
        original,
        &[SHOW_SIZE_FIELD, SIZE_WIDTH_FIELD],
        true,
        Some(settings.width.to_bits()),
    )?;
    data = patch_nested_fixed32_field(
        &data,
        &[SHOW_SIZE_FIELD, SIZE_HEIGHT_FIELD],
        true,
        Some(settings.height.to_bits()),
    )?;
    for (field_number, current, replacement) in [
        (
            SLIDE_NUMBERS_VISIBLE_FIELD,
            show.slide_numbers_visible,
            settings.slide_numbers_visible,
        ),
        (
            LOOP_PRESENTATION_FIELD,
            show.loop_presentation,
            settings.loop_presentation,
        ),
        (
            IDLE_TIMER_ACTIVE_FIELD,
            show.idle_timer_active,
            settings.idle_timer_active,
        ),
        (
            AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD,
            show.automatically_plays_upon_open,
            settings.automatically_plays_upon_open,
        ),
    ] {
        data = patch_varint_field(
            &data,
            field_number,
            current.is_some(),
            replacement.map(u64::from),
        )?;
    }
    data = patch_varint_field(
        &data,
        MODE_FIELD,
        show.mode.is_some(),
        settings.mode.map(|mode| i64::from(mode.as_raw()) as u64),
    )?;
    for (field_number, current, replacement) in [
        (
            AUTOPLAY_TRANSITION_DELAY_FIELD,
            show.autoplay_transition_delay,
            settings.autoplay_transition_delay,
        ),
        (
            AUTOPLAY_BUILD_DELAY_FIELD,
            show.autoplay_build_delay,
            settings.autoplay_build_delay,
        ),
        (
            IDLE_TIMER_DELAY_FIELD,
            show.idle_timer_delay,
            settings.idle_timer_delay,
        ),
    ] {
        data = patch_fixed64_field(
            &data,
            field_number,
            current.is_some(),
            replacement.map(f64::to_bits),
        )?;
    }
    Ok(data)
}

fn validate_show_settings(settings: &KeynoteShowSettings) -> Result<()> {
    if !settings.width.is_finite()
        || settings.width <= 0.0
        || !settings.height.is_finite()
        || settings.height <= 0.0
    {
        return Err(Error::ParseError(
            "Keynote show dimensions must be finite and greater than zero".to_owned(),
        ));
    }
    if settings.mode.is_some_and(|mode| !mode.is_canonical()) {
        return Err(Error::ParseError(
            "Keynote show mode must use its named variant for known native values".to_owned(),
        ));
    }
    for (name, value) in [
        (
            "autoplay transition delay",
            settings.autoplay_transition_delay,
        ),
        ("autoplay build delay", settings.autoplay_build_delay),
        ("idle timer delay", settings.idle_timer_delay),
    ] {
        if value.is_some_and(|value| !value.is_finite() || value < 0.0) {
            return Err(Error::ParseError(format!(
                "Keynote {name} must be finite and non-negative"
            )));
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn show_modes_map_native_values_losslessly() {
        for (raw, mode) in [
            (0, KeynoteShowMode::Normal),
            (1, KeynoteShowMode::SelfPlaying),
            (2, KeynoteShowMode::LinksOnly),
            (19, KeynoteShowMode::Unknown(19)),
            (-1, KeynoteShowMode::Unknown(-1)),
        ] {
            assert_eq!(KeynoteShowMode::from_raw(raw), mode);
            assert_eq!(mode.as_raw(), raw);
        }
    }
}
