//! Typed playback settings for a Keynote presentation soundtrack.

use super::soundtrack_wire::{
    decode_soundtrack, patch_soundtrack_wire, read_soundtrack, replace_soundtrack_message,
};
use super::*;
use litchi_keynote::soundtrack::Settings;

fn settings_from_native(soundtrack: &kn::Soundtrack) -> Result<Settings> {
    Settings::new(
        soundtrack.volume,
        soundtrack
            .mode
            .map(litchi_keynote::soundtrack::Mode::from_raw),
    )
    .map_err(|error| Error::ParseError(format!("invalid Keynote soundtrack settings: {error}")))
}

impl KeynoteEditor {
    /// Read soundtrack playback state, or `None` for a legacy show without a soundtrack object.
    pub fn soundtrack_settings(&self) -> Result<Option<Settings>> {
        let graph = ObjectGraph::read(self.package())?;
        let Some(record) = read_soundtrack(&graph)? else {
            return Ok(None);
        };
        Ok(Some(settings_from_native(&record.native)?))
    }

    /// Replace soundtrack mode and volume without changing its media collection.
    pub fn set_soundtrack_settings(&mut self, settings: Settings) -> Result<()> {
        settings
            .validate()
            .map_err(|error| Error::ParseError(error.to_string()))?;
        let graph = ObjectGraph::read(self.package())?;
        let Some(record) = read_soundtrack(&graph)? else {
            return Err(Error::InvalidFormat(
                "Keynote show has no soundtrack object".to_owned(),
            ));
        };
        let current = settings_from_native(&record.native)?;
        if current == settings {
            return Ok(());
        }

        let data = patch_soundtrack_wire(record.data, &record.native, &settings)?;
        let verified_native = decode_soundtrack(&data)?;
        if settings_from_native(&verified_native)? != settings {
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
