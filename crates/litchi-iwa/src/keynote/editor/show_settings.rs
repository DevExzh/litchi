//! Protobuf adaptation for Keynote show settings.
//!
//! The validated semantic model lives in `litchi-keynote::show`; this module
//! only resolves archive objects and patches the native wire representation.

use super::*;
use litchi_keynote::{Mode, Seconds, Settings, Size};

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

fn semantic_error(error: litchi_keynote::Error) -> Error {
    Error::ParseError(format!("invalid Keynote show settings: {error}"))
}

fn settings_from_archive(show: &kn::ShowArchive) -> Result<Settings> {
    let size = Size::new(show.size.width, show.size.height).map_err(semantic_error)?;
    let mut settings = Settings::new(size);
    settings.set_slide_numbers_visible(show.slide_numbers_visible);
    settings.set_loop_presentation(show.loop_presentation);
    settings
        .set_mode(show.mode.map(Mode::from_raw))
        .map_err(semantic_error)?;
    settings.set_autoplay_transition_delay(
        show.autoplay_transition_delay
            .map(Seconds::new)
            .transpose()
            .map_err(semantic_error)?,
    );
    settings.set_autoplay_build_delay(
        show.autoplay_build_delay
            .map(Seconds::new)
            .transpose()
            .map_err(semantic_error)?,
    );
    settings.set_idle_timer_active(show.idle_timer_active);
    settings.set_idle_timer_delay(
        show.idle_timer_delay
            .map(Seconds::new)
            .transpose()
            .map_err(semantic_error)?,
    );
    settings.set_automatically_plays_upon_open(show.automatically_plays_upon_open);
    settings.validate().map_err(semantic_error)?;
    Ok(settings)
}

impl KeynoteEditor {
    /// Read validated presentation dimensions and playback behavior.
    ///
    /// # Errors
    ///
    /// Returns an error when the document/show archive is missing, malformed,
    /// or contains invalid semantic values.
    pub fn show_settings(&self) -> Result<Settings> {
        let graph = ObjectGraph::read(self.text.package())?;
        let document: kn::DocumentArchive =
            graph.decode(DOCUMENT_OBJECT_ID, "KN.DocumentArchive")?;
        let show: kn::ShowArchive = graph.decode(document.show.identifier, "KN.ShowArchive")?;
        settings_from_archive(&show)
    }

    /// Replace presentation-level dimensions and playback behavior transactionally.
    ///
    /// # Errors
    ///
    /// Returns an error when the settings are invalid, the native show archive
    /// cannot be decoded, or the staged wire patch fails verification.
    pub fn set_show_settings(&mut self, settings: Settings) -> Result<()> {
        settings.validate().map_err(semantic_error)?;
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
            if settings_from_archive(&verified)? != settings {
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
    settings: &Settings,
) -> Result<Vec<u8>> {
    let size = settings.size();
    let mut data = patch_nested_fixed32_field(
        original,
        &[SHOW_SIZE_FIELD, SIZE_WIDTH_FIELD],
        true,
        Some(size.width().to_bits()),
    )?;
    data = patch_nested_fixed32_field(
        &data,
        &[SHOW_SIZE_FIELD, SIZE_HEIGHT_FIELD],
        true,
        Some(size.height().to_bits()),
    )?;
    for (field_number, current, replacement) in [
        (
            SLIDE_NUMBERS_VISIBLE_FIELD,
            show.slide_numbers_visible,
            settings.slide_numbers_visible(),
        ),
        (
            LOOP_PRESENTATION_FIELD,
            show.loop_presentation,
            settings.loop_presentation(),
        ),
        (
            IDLE_TIMER_ACTIVE_FIELD,
            show.idle_timer_active,
            settings.idle_timer_active(),
        ),
        (
            AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD,
            show.automatically_plays_upon_open,
            settings.automatically_plays_upon_open(),
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
        settings.mode().map(|mode| i64::from(mode.as_raw()) as u64),
    )?;
    for (field_number, current, replacement) in [
        (
            AUTOPLAY_TRANSITION_DELAY_FIELD,
            show.autoplay_transition_delay,
            settings.autoplay_transition_delay(),
        ),
        (
            AUTOPLAY_BUILD_DELAY_FIELD,
            show.autoplay_build_delay,
            settings.autoplay_build_delay(),
        ),
        (
            IDLE_TIMER_DELAY_FIELD,
            show.idle_timer_delay,
            settings.idle_timer_delay(),
        ),
    ] {
        data = patch_fixed64_field(
            &data,
            field_number,
            current.is_some(),
            replacement.map(Seconds::as_f64).map(f64::to_bits),
        )?;
    }
    Ok(data)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn show_modes_map_native_values_losslessly() -> Result<()> {
        for (raw, mode) in [
            (0, Mode::Normal),
            (1, Mode::SelfPlaying),
            (2, Mode::LinksOnly),
            (19, Mode::Unknown(19)),
            (-1, Mode::Unknown(-1)),
        ] {
            assert_eq!(Mode::from_raw(raw), mode);
            assert_eq!(mode.as_raw(), raw);
        }
        Ok(())
    }
}
