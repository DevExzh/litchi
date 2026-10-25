//! Optional system-font discovery and `PowerPoint` EOT publication.

use litchi_fonts::{CollectGlyphs, GlyphMap, Prepared};

use crate::font::{self, Charset, Data, Face, Family, Font, Fonts, Pitch, PitchFamily, Style};
use crate::package::Package;
use crate::{Error, Result};

impl Package {
    pub(crate) fn embed_fonts_for_presentation(
        &mut self,
        presentation: &crate::builder::Presentation,
        preparation_limits: litchi_fonts::PreparationLimits,
        eot_limits: litchi_fonts::embedding::powerpoint::Limits,
    ) -> Result<()> {
        let mode = match self.opc.save_options().fonts {
            litchi_opc::FontEmbedding::None => return Ok(()),
            litchi_opc::FontEmbedding::Full => litchi_fonts::embedding::Mode::Full,
            litchi_opc::FontEmbedding::Subset => litchi_fonts::embedding::Mode::Subset,
        };
        self.embed_fonts_with_glyphs(
            presentation.collect_glyphs(),
            mode,
            preparation_limits,
            eot_limits,
        )
    }

    fn embed_fonts_with_glyphs(
        &mut self,
        glyphs: GlyphMap,
        mode: litchi_fonts::embedding::Mode,
        preparation_limits: litchi_fonts::PreparationLimits,
        eot_limits: litchi_fonts::embedding::powerpoint::Limits,
    ) -> Result<()> {
        let prepared = litchi_fonts::embedding::prepare(glyphs, mode, preparation_limits)?;
        if prepared.is_empty() {
            return Ok(());
        }

        let mut staged = (*self.opc).clone();
        let mut fonts = font::load(&staged)?.unwrap_or_else(Fonts::new);
        let mut changed = false;
        for mut prepared in prepared {
            let data = litchi_fonts::embedding::powerpoint::encode(
                &mut prepared,
                litchi_fonts::embedding::powerpoint::Intent::PreviewPrint,
                eot_limits,
            )?;
            changed |= merge(&mut fonts, prepared, data)?;
        }
        if changed {
            let _ = font::put(&mut staged, fonts)?;
            self.opc = std::sync::Arc::new(staged);
        }
        Ok(())
    }
}

fn merge(fonts: &mut Fonts, prepared: Prepared, data: Vec<u8>) -> Result<bool> {
    let Prepared {
        name,
        style,
        properties,
        ..
    } = prepared;
    let current = match fonts.get(name.as_str()) {
        Ok(font) => Some(font.clone()),
        Err(Error::FontNotFound(_)) => None,
        Err(error) => return Err(error),
    };
    let mut next = current.clone().unwrap_or(Font::new(&name)?);
    let _ = next.set_panose(Some(properties.panose().into_bytes().into()));
    let _ = next.set_pitch_family(Some(PitchFamily::new(
        pitch(properties.pitch()),
        family(properties.family()),
    )));
    let _ = next.set_charset(properties.charset().map(|value| Charset::new(value.code())));

    let style = map_style(style);
    let matches = next.get(style).is_some_and(|face| {
        face.data().format() == font::Format::PowerPoint && face.data().bytes() == data
    });
    if !matches {
        let _ = next.put(Face::new(style, Data::powerpoint(data)?))?;
    }

    if current.as_ref() == Some(&next) {
        return Ok(false);
    }
    match current {
        Some(_) => {
            let _ = fonts.replace(name.as_str(), next)?;
        },
        None => fonts.add(next)?,
    }
    Ok(true)
}

fn map_style(value: litchi_fonts::Style) -> Style {
    match value {
        litchi_fonts::Style::Regular => Style::Regular,
        litchi_fonts::Style::Bold => Style::Bold,
        litchi_fonts::Style::Italic => Style::Italic,
        litchi_fonts::Style::BoldItalic => Style::BoldItalic,
    }
}

fn family(value: litchi_fonts::Family) -> Family {
    match value {
        litchi_fonts::Family::Auto => Family::None,
        litchi_fonts::Family::Roman => Family::Roman,
        litchi_fonts::Family::Swiss => Family::Swiss,
        litchi_fonts::Family::Modern => Family::Modern,
        litchi_fonts::Family::Script => Family::Script,
        litchi_fonts::Family::Decorative => Family::Decorative,
    }
}

fn pitch(value: litchi_fonts::Pitch) -> Pitch {
    match value {
        litchi_fonts::Pitch::Default => Pitch::Default,
        litchi_fonts::Pitch::Fixed => Pitch::Fixed,
        litchi_fonts::Pitch::Variable => Pitch::Variable,
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use litchi_fonts::{Charset as FontCharset, FontProperties, License, Panose, Signature};

    fn prepared(style: litchi_fonts::Style) -> Prepared {
        Prepared {
            name: "Litchi Test".into(),
            style,
            data: Vec::new(),
            properties: FontProperties::new(
                License::new(0).unwrap(),
                Panose::new([2, 11, 6, 4, 2, 2, 2, 2, 2, 4]),
                Some(FontCharset::ANSI),
                litchi_fonts::Family::Roman,
                litchi_fonts::Pitch::Variable,
                Signature::new([0; 4], [0; 2]),
            ),
            subsetted: true,
        }
    }

    fn eot(payload: u8) -> Vec<u8> {
        let mut value = vec![0; 96];
        value[0..4].copy_from_slice(&108_u32.to_le_bytes());
        value[4..8].copy_from_slice(&12_u32.to_le_bytes());
        value[8..12].copy_from_slice(&0x0001_0000_u32.to_le_bytes());
        value[16] = payload;
        value[34..36].copy_from_slice(&0x504C_u16.to_le_bytes());
        value.extend_from_slice(&[0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
        value
    }

    #[test]
    fn merges_a_deterministic_powerpoint_face_without_discovery() {
        let mut fonts = Fonts::new();
        assert!(merge(&mut fonts, prepared(litchi_fonts::Style::Bold), eot(7)).unwrap());
        let font = fonts.get("litchi test").unwrap();
        assert_eq!(font.panose(), Some([2, 11, 6, 4, 2, 2, 2, 2, 2, 4].into()));
        assert_eq!(
            font.pitch_family(),
            Some(PitchFamily::new(Pitch::Variable, Family::Roman))
        );
        assert_eq!(font.charset(), Some(Charset::ANSI));
        assert_eq!(
            font.get(Style::Bold).unwrap().data().format(),
            font::Format::PowerPoint
        );
        assert!(!merge(&mut fonts, prepared(litchi_fonts::Style::Bold), eot(7)).unwrap());
    }
}
