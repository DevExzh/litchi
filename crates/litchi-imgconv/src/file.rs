//! Ergonomic codec operations for move-first OfficeArt image files.

use std::borrow::Cow;

use litchi_core::error::{Error, Result};
use litchi_odraw::image::{File, Kind};

use crate::{Limits, Options, decode_data, to_jpeg, to_png, to_svg, to_webp};

/// Optional codec operations for an OfficeArt [`File`].
///
/// Importing this trait adds short conversion verbs without coupling the
/// format-neutral OfficeArt crate to a codec implementation.
pub trait Convert {
    /// Returns bounded, decompressed native file data.
    fn decode(&self, limits: Limits) -> Result<Cow<'_, [u8]>>;

    /// Converts this file to PNG under explicit sizing and resource limits.
    fn png(&self, options: Options) -> Result<Vec<u8>>;

    /// Converts this file to JPEG under explicit sizing and resource limits.
    fn jpeg(&self, options: Options) -> Result<Vec<u8>>;

    /// Converts this file to WebP under explicit sizing and resource limits.
    fn webp(&self, options: Options) -> Result<Vec<u8>>;

    /// Converts an EMF or WMF file to SVG under explicit limits.
    fn svg(&self, options: Options) -> Result<String>;

    /// Extracts the recommended representation under explicit limits.
    fn extract(&self, options: Options) -> Result<Vec<u8>>;

    /// Returns a filename matching [`Convert::extract`]'s representation.
    fn out_name(&self) -> String;
}

impl Convert for File<'_> {
    fn decode(&self, limits: Limits) -> Result<Cow<'_, [u8]>> {
        let blip = self.blip().map_err(office_art)?;
        decode_data(&blip, &limits)
    }

    fn png(&self, options: Options) -> Result<Vec<u8>> {
        to_png(&self.blip().map_err(office_art)?, options)
    }

    fn jpeg(&self, options: Options) -> Result<Vec<u8>> {
        to_jpeg(&self.blip().map_err(office_art)?, options)
    }

    fn webp(&self, options: Options) -> Result<Vec<u8>> {
        to_webp(&self.blip().map_err(office_art)?, options)
    }

    fn svg(&self, options: Options) -> Result<String> {
        to_svg(&self.blip().map_err(office_art)?, options)
    }

    fn extract(&self, options: Options) -> Result<Vec<u8>> {
        match self.kind() {
            Kind::Emf | Kind::Wmf => self.svg(options).map(String::into_bytes),
            Kind::Pict | Kind::Dib | Kind::Tiff => self.png(options),
            Kind::Jpeg | Kind::CmykJpeg | Kind::Png => {
                let decoded = self.decode(options.limits)?;
                if decoded.len() > options.limits.max_output_bytes {
                    return Err(Error::ParseError(format!(
                        "native image output exceeds limit {} bytes",
                        options.limits.max_output_bytes
                    )));
                }
                Ok(decoded.into_owned())
            },
            Kind::Error | Kind::Unknown | Kind::Other(_) => Err(Error::Unsupported(
                "unknown OfficeArt images cannot be decoded".to_string(),
            )),
        }
    }

    fn out_name(&self) -> String {
        let extension = match self.kind() {
            Kind::Emf | Kind::Wmf => "svg",
            Kind::Pict | Kind::Dib | Kind::Tiff => "png",
            kind => kind.extension(),
        };
        let native = self.filename();
        let stem = native
            .rsplit_once('.')
            .map_or(native.as_str(), |(stem, _)| stem);
        format!("{stem}.{extension}")
    }
}

fn office_art(error: litchi_odraw::Error) -> Error {
    Error::ParseError(error.to_string())
}

#[cfg(test)]
mod tests {
    use std::borrow::Cow;

    use litchi_odraw::image::{Blip, File};

    use super::Convert;
    use crate::{Limits, Options};

    fn png_blip(data: &[u8]) -> Vec<u8> {
        let mut payload = vec![0; 16];
        payload.push(0xff);
        payload.extend_from_slice(data);
        let mut record = Vec::new();
        record.extend_from_slice(&(0x6e0u16 << 4).to_le_bytes());
        record.extend_from_slice(&0xf01eu16.to_le_bytes());
        record.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        record.extend_from_slice(&payload);
        record
    }

    #[test]
    fn native_bitmap_decode_stays_borrowed() {
        let record = png_blip(b"png");
        let file = File::new(Blip::parse(&record).unwrap(), None, 0);
        let decoded = file.decode(Limits::default()).unwrap();
        assert!(matches!(&decoded, Cow::Borrowed(_)));
        assert_eq!(decoded.as_ref(), b"png");
    }

    #[test]
    fn extraction_and_output_name_share_the_recommended_kind() {
        let record = png_blip(b"png");
        let file = File::new(
            Blip::parse(&record).unwrap(),
            Some("../../CON.old".to_string()),
            0,
        );
        assert_eq!(file.extract(Options::default()).unwrap(), b"png");
        assert_eq!(file.out_name(), "CON_.png");
    }

    #[test]
    fn native_extraction_honors_output_limit() {
        let record = png_blip(b"png");
        let file = File::new(Blip::parse(&record).unwrap(), None, 0);
        let limits = Limits {
            max_output_bytes: 2,
            ..Limits::default()
        };
        assert!(file.extract(Options::default().limits(limits)).is_err());
    }
}
