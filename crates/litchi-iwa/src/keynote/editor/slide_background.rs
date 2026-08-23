//! Compatibility facade for focused Keynote slide-background transactions.
//!
//! The semantic background graph belongs to `litchi-keynote`. This adapter
//! keeps the historical `KeynoteEditor` methods available while routing every
//! operation through the selector-first package transaction. The old native
//! protobuf conversion oracle is retained only for the host crate's legacy
//! differential tests; it is not part of production ownership.

use litchi_keynote::background::Background;

use super::*;

impl KeynoteEditor {
    /// Read the effective background through the focused Keynote package API.
    pub fn slide_background(&self, slide_index: usize) -> Result<Background> {
        focused_slide_background_package(self)?
            .slide_background(litchi_core::Position::new(slide_index))
            .map_err(map_focused_slide_background_error)
    }

    /// Read the direct variation-style background override, if one exists.
    pub fn slide_background_override(&self, slide_index: usize) -> Result<Option<Background>> {
        focused_slide_background_package(self)?
            .slide_background_override(litchi_core::Position::new(slide_index))
            .map_err(map_focused_slide_background_error)
    }

    /// Set a slide background through the focused package transaction.
    pub fn set_slide_background(
        &mut self,
        slide_index: usize,
        background: Background,
    ) -> Result<()> {
        let package = focused_slide_background_package(self)?;
        let commit = package
            .edit_slide_background(litchi_core::Position::new(slide_index))
            .map_err(map_focused_slide_background_error)?
            .set(background)
            .map_err(map_focused_slide_background_error)?
            .commit()
            .map_err(map_focused_slide_background_error)?;
        if commit.patch().is_noop() {
            return Ok(());
        }
        replace_from_focused_slide_background_commit(self, commit.package())
    }

    /// Remove a direct background override and restore style inheritance.
    pub fn reset_slide_background(&mut self, slide_index: usize) -> Result<bool> {
        let package = focused_slide_background_package(self)?;
        let commit = package
            .edit_slide_background(litchi_core::Position::new(slide_index))
            .map_err(map_focused_slide_background_error)?
            .clear()
            .map_err(map_focused_slide_background_error)?
            .commit()
            .map_err(map_focused_slide_background_error)?;
        if commit.patch().is_noop() {
            return Ok(false);
        }
        replace_from_focused_slide_background_commit(self, commit.package())?;
        Ok(true)
    }
}

fn focused_slide_background_package(editor: &KeynoteEditor) -> Result<litchi_keynote::Package> {
    let bytes = editor.to_bytes()?;
    litchi_keynote::Package::from_bytes(&bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote slide background source failed: {error}"
        ))
    })
}

fn replace_from_focused_slide_background_commit(
    editor: &mut KeynoteEditor,
    package: &litchi_keynote::Package,
) -> Result<()> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "focused Keynote slide background write failed: {error}"
        ))
    })?;
    *editor = KeynoteEditor::from_bytes(&bytes)?;
    Ok(())
}

fn map_focused_slide_background_error<E: std::fmt::Display>(error: E) -> Error {
    Error::InvalidFormat(format!(
        "focused Keynote slide background operation failed: {error}"
    ))
}

// This parser remains a test-only generated-Prost differential oracle for the
// old host fixtures. Production reads and writes are owned by the focused
// lazy codec behind `litchi-keynote::Package`.
#[cfg(test)]
pub(super) fn background_from_fill(fill_payload: &[u8]) -> Result<Background> {
    use litchi_keynote::background::Opaque;
    use prost::Message as _;

    let fill = tsd::FillArchive::decode(fill_payload)?;
    if fill_payload.is_empty()
        && fill.color.is_none()
        && fill.gradient.is_none()
        && fill.image.is_none()
    {
        return Ok(Background::None);
    }
    if fill.color.is_none() && fill.gradient.is_some() && fill.image.is_none() {
        return Ok(
            match super::slide_background_gradient_wire::gradient_from_fill(fill_payload)? {
                Some(gradient) => Background::Gradient(gradient),
                None => opaque_background(fill_payload)?,
            },
        );
    }
    let Some(color) = fill.color.as_ref() else {
        return opaque_background(fill_payload);
    };
    if fill.gradient.is_some() || fill.image.is_some() {
        return opaque_background(fill_payload);
    }
    Ok(
        match super::slide_background_color::color_from_native(color) {
            Some(color) => Background::Solid(color),
            None => return opaque_background(fill_payload),
        },
    )
}

#[cfg(test)]
fn opaque_background(fill_payload: &[u8]) -> Result<Background> {
    use litchi_keynote::background::Opaque;

    Opaque::from_slice(fill_payload)
        .map(Background::Opaque)
        .map_err(|error| Error::ParseError(error.to_string()))
}
