//! Inline picture insertion (record 0251d).
//!
//! `add(<paragraph>, picture, {src, width, height, alt, name})` reads the
//! image from a base64 `data:` URI or a local file path, bounded by the
//! default staged-byte limit, and stages it through the leaf
//! `insert_picture_at`: a media part, its relationship and a run holding
//! the inline drawing, in one commit.

use litchi_core::FileFormat;
use litchi_docx::document::NewPicture;

use super::setters::Setter;
use crate::error::{Code, Error, Result};
use crate::props::Props;
use crate::registry::Operation;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message)
        .with_format(FileFormat::Docx)
        .with_property("src")
}

fn bound(observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(crate::transaction::Error::Limit {
            kind: LimitKind::StagedBytes,
            observed,
            maximum,
        }
        .into());
    }
    Ok(())
}

/// Read a picture source: a base64 `data:` URI or a local file path.
fn read_source(src: &str, maximum: usize) -> Result<Vec<u8>> {
    if let Some(rest) = src.strip_prefix("data:") {
        let (header, payload) = rest.split_once(',').ok_or_else(|| {
            error(
                Code::InvalidValue,
                "a data URI needs a comma before its payload",
            )
        })?;
        if !header.ends_with(";base64") {
            return Err(error(
                Code::UnsupportedCapability,
                "only base64 data URIs are accepted",
            ));
        }
        // Base64 expands by 4/3; refuse before decoding.
        bound(payload.len() / 4 * 3, maximum)?;
        use base64::Engine as _;
        return base64::engine::general_purpose::STANDARD
            .decode(payload.trim())
            .map_err(|_| error(Code::InvalidValue, "the data URI payload is not base64"));
    }
    let length = std::fs::metadata(src)
        .map_err(|source| Error::from(source).with_property("src"))?
        .len();
    bound(usize::try_from(length).unwrap_or(usize::MAX), maximum)?;
    use std::io::Read as _;
    let mut data = Vec::new();
    std::fs::File::open(src)
        .and_then(|file| file.take(maximum as u64 + 1).read_to_end(&mut data))
        .map_err(|source| Error::from(source).with_property("src"))?;
    bound(data.len(), maximum)?;
    Ok(data)
}

/// A new picture from its add properties: `src` is required.
pub(super) fn new_picture(props: &Props<'_>) -> Result<NewPicture> {
    let mut src = None::<String>;
    let mut width = None;
    let mut height = None;
    let mut alt = None::<String>;
    let mut name = None::<String>;
    for (key, value) in props.iter() {
        let (_, entry) = super::spec::PICTURE_REGISTRY
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        entry
            .prepare(Operation::Add, value, |setter, value| {
                let text = |value: Option<Scalar<'_>>| match value {
                    Some(Scalar::Text(text) | Scalar::Token(text)) => text.to_owned(),
                    Some(scalar) => scalar.to_string(),
                    None => String::new(),
                };
                let emu = |value: Option<Scalar<'_>>| match value {
                    Some(Scalar::Length(length)) => Ok(length.emu()),
                    _ => Err(Error::new(
                        Code::InvalidValue,
                        "a picture extent is a length",
                    )),
                };
                match setter {
                    Setter::PictureSrc => src = Some(text(value)),
                    Setter::PictureWidth => width = Some(emu(value)?),
                    Setter::PictureHeight => height = Some(emu(value)?),
                    Setter::PictureAlt => alt = Some(text(value)),
                    Setter::PictureName => name = Some(text(value)),
                    _ => {
                        return Err(Error::new(
                            Code::UnsupportedProperty,
                            "the property does not seed a DOCX picture",
                        ));
                    },
                }
                Ok(())
            })
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
    }
    let Some(src) = src else {
        return Err(Error::new(
            Code::MissingProperty,
            "a new DOCX picture needs its image source",
        )
        .with_property("src")
        .with_format(FileFormat::Docx));
    };
    let data = read_source(&src, SemanticLimits::default().max_staged_bytes())?;
    if litchi_docx::writer::ImageFormat::detect_from_bytes(&data).is_none() {
        return Err(error(
            Code::InvalidValue,
            "the picture source is not a recognized image",
        ));
    }
    let mut picture = NewPicture::new(data).with_size(width, height);
    if let Some(alt) = alt {
        picture = picture.with_description(alt);
    }
    if let Some(name) = name {
        picture = picture.with_name(name);
    }
    Ok(picture)
}
