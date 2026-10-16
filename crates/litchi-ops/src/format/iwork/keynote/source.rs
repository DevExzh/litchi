//! The image bytes of a new Keynote picture: a base64 `data:` URI or a
//! local file path, as the PPTX `src` property reads them.

use crate::Result;
use crate::error::{Code, Error};
use crate::transaction::LimitKind;

use super::super::write::{bound, refusal};

fn invalid(message: &'static str) -> Error {
    refusal(Code::InvalidValue, message).with_property("src".to_owned())
}

/// Read a picture source: a base64 `data:` URI or a local file path, at
/// most `maximum` bytes.
pub(super) fn read(src: &str, maximum: usize) -> Result<Vec<u8>> {
    if let Some(rest) = src.strip_prefix("data:") {
        let (header, payload) = rest
            .split_once(',')
            .ok_or_else(|| invalid("a data URI needs a comma before its payload"))?;
        if !header.ends_with(";base64") {
            return Err(refusal(
                Code::UnsupportedCapability,
                "only base64 data URIs are accepted",
            )
            .with_property("src".to_owned()));
        }
        // Base64 expands by 4/3; refuse before decoding.
        bound(LimitKind::StagedBytes, payload.len() / 4 * 3, maximum)?;
        use base64::Engine as _;
        return base64::engine::general_purpose::STANDARD
            .decode(payload.trim())
            .map_err(|_| invalid("the data URI payload is not base64"));
    }
    let length = std::fs::metadata(src)
        .map_err(|source| Error::from(source).with_property("src"))?
        .len();
    bound(
        LimitKind::StagedBytes,
        usize::try_from(length).unwrap_or(usize::MAX),
        maximum,
    )?;
    use std::io::Read as _;
    let mut data = Vec::new();
    std::fs::File::open(src)
        .and_then(|file| file.take(maximum as u64 + 1).read_to_end(&mut data))
        .map_err(|source| Error::from(source).with_property("src"))?;
    bound(LimitKind::StagedBytes, data.len(), maximum)?;
    Ok(data)
}

/// The file name of a path source; `None` for a data URI.
pub(super) fn file_name(src: &str) -> Option<String> {
    if src.starts_with("data:") {
        return None;
    }
    std::path::Path::new(src)
        .file_name()
        .and_then(|name| name.to_str())
        .map(str::to_owned)
}

/// A file name for image bytes without one, by their signature.
pub(super) fn default_name(bytes: &[u8]) -> &'static str {
    if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
        "image.jpeg"
    } else if bytes.starts_with(b"GIF8") {
        "image.gif"
    } else {
        "image.png"
    }
}
