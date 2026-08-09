//! Contextual audio/video authoring values.
//!
//! The package graph for persisted media pictures lives in
//! [`crate::media_parts`]. This module is the small authoring layer used by
//! slide writers: it detects a format from bounded bytes, retains payloads in
//! the shared immutable media buffer, and emits a self-contained picture
//! fragment when a writer has assigned relationship IDs.

use crate::media_parts::Data;
use crate::{Error, Result};
use litchi_opc::constants::relationship_type as rt;
use std::fmt::Write;

const MAX_DATA_BYTES: usize = 128 * 1024 * 1024;
const MAX_NAME_BYTES: usize = 4 * 1024;

/// Whether an authored media item is audio or video.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Audio,
    Video,
}

impl Kind {
    /// MIME family used by the corresponding resource.
    #[must_use]
    pub const fn mime_prefix(self) -> &'static str {
        match self {
            Self::Audio => "audio",
            Self::Video => "video",
        }
    }

    /// OPC relationship type used for the media payload.
    #[must_use]
    pub const fn relationship_type(self) -> &'static str {
        match self {
            Self::Audio => rt::AUDIO,
            Self::Video => rt::VIDEO,
        }
    }
}

/// The bounded set of media formats recognized by the authoring facade.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Format {
    Mp3,
    Wav,
    Wma,
    M4a,
    Mp4,
    Wmv,
    Avi,
    Mov,
    Unknown,
}

impl Format {
    /// Detect a format from a file-name extension.
    #[must_use]
    pub fn from_extension(extension: &str) -> Self {
        match extension
            .trim_start_matches('.')
            .to_ascii_lowercase()
            .as_str()
        {
            "mp3" => Self::Mp3,
            "wav" => Self::Wav,
            "wma" => Self::Wma,
            "m4a" => Self::M4a,
            "mp4" => Self::Mp4,
            "wmv" => Self::Wmv,
            "avi" => Self::Avi,
            "mov" => Self::Mov,
            _ => Self::Unknown,
        }
    }

    /// Detect a format from bounded magic bytes.
    #[must_use]
    pub fn detect_from_bytes(data: &[u8]) -> Self {
        if data.starts_with(b"ID3") {
            return Self::Mp3;
        }
        if data.len() >= 2 && data[0] == 0xff && (data[1] & 0xe0) == 0xe0 {
            return Self::Mp3;
        }
        if data.len() >= 12 && data.starts_with(b"RIFF") {
            return match &data[8..12] {
                b"WAVE" => Self::Wav,
                b"AVI " => Self::Avi,
                _ => Self::Unknown,
            };
        }
        if data.len() >= 12 && &data[4..8] == b"ftyp" {
            return match &data[8..12] {
                b"M4A " | b"M4B " => Self::M4a,
                b"qt  " | b"moov" => Self::Mov,
                _ => Self::Mp4,
            };
        }
        if data.len() >= 16 && data[..8] == [0x30, 0x26, 0xb2, 0x75, 0x8e, 0x66, 0xcf, 0x11] {
            return Self::Wmv;
        }
        Self::Unknown
    }

    /// MIME type associated with this format.
    #[must_use]
    pub const fn mime_type(self) -> &'static str {
        match self {
            Self::Mp3 => "audio/mpeg",
            Self::Wav => "audio/wav",
            Self::Wma => "audio/x-ms-wma",
            Self::M4a => "audio/mp4",
            Self::Mp4 => "video/mp4",
            Self::Wmv => "video/x-ms-wmv",
            Self::Avi => "video/avi",
            Self::Mov => "video/quicktime",
            Self::Unknown => "application/octet-stream",
        }
    }

    /// Conventional extension used for a generated media part.
    #[must_use]
    pub const fn extension(self) -> &'static str {
        match self {
            Self::Mp3 => "mp3",
            Self::Wav => "wav",
            Self::Wma => "wma",
            Self::M4a => "m4a",
            Self::Mp4 => "mp4",
            Self::Wmv => "wmv",
            Self::Avi => "avi",
            Self::Mov => "mov",
            Self::Unknown => "bin",
        }
    }

    /// Infer the audio/video family represented by this format.
    #[must_use]
    pub const fn kind(self) -> Kind {
        match self {
            Self::Mp3 | Self::Wav | Self::Wma | Self::M4a => Kind::Audio,
            Self::Mp4 | Self::Wmv | Self::Avi | Self::Mov | Self::Unknown => Kind::Video,
        }
    }
}

/// A move-first media item waiting to be attached to a slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Item {
    /// Immutable media bytes shared by clones.
    pub data: Data,
    /// Detected or explicitly selected format.
    pub format: Format,
    /// Horizontal offset in EMUs.
    pub x: i64,
    /// Vertical offset in EMUs.
    pub y: i64,
    /// Width in EMUs.
    pub width: i64,
    /// Height in EMUs.
    pub height: i64,
    /// Optional producer-visible name.
    pub name: Option<String>,
    /// Whether playback loops.
    pub loop_playback: bool,
    /// Whether playback starts automatically.
    pub auto_play: bool,
    /// Whether the media icon is hidden during the show.
    pub hide_during_show: bool,
}

impl Item {
    /// Construct an item with format detection and no optional playback flags.
    #[must_use]
    pub fn new(data: Vec<u8>, x: i64, y: i64, width: i64, height: i64) -> Self {
        let format = Format::detect_from_bytes(&data);
        Self::with_format(data, format, x, y, width, height)
    }

    /// Construct an item with an explicit format.
    #[must_use]
    pub fn with_format(
        data: Vec<u8>,
        format: Format,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> Self {
        Self {
            data: Data::from(data),
            format,
            x,
            y,
            width,
            height,
            name: None,
            loop_playback: false,
            auto_play: false,
            hide_during_show: false,
        }
    }

    /// Validate payload, name, and positive frame dimensions before storage.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn validate(&self) -> Result<()> {
        if self.data.len() > MAX_DATA_BYTES {
            return Err(Error::Limit {
                resource: "presentation media payload bytes",
                limit: MAX_DATA_BYTES,
            });
        }
        if self.width <= 0 || self.height <= 0 {
            return Err(Error::Invalid(
                "presentation media frame extents must be positive".to_string(),
            ));
        }
        if let Some(name) = &self.name
            && (name.is_empty() || name.len() > MAX_NAME_BYTES)
        {
            return Err(Error::Invalid(
                "presentation media name is empty or too long".to_string(),
            ));
        }
        Ok(())
    }

    /// Set the producer-visible name.
    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = Some(name.into());
    }

    /// Enable looping playback.
    #[must_use]
    pub fn with_loop(mut self) -> Self {
        self.loop_playback = true;
        self
    }

    /// Enable automatic playback.
    #[must_use]
    pub fn with_auto_play(mut self) -> Self {
        self.auto_play = true;
        self
    }

    /// Hide the media icon during the show.
    #[must_use]
    pub fn with_hide_during_show(mut self) -> Self {
        self.hide_during_show = true;
        self
    }

    /// Return the audio/video family selected by the format.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.format.kind()
    }

    /// Serialize a bounded, namespace-self-contained media picture fragment.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_shape_xml(
        &self,
        shape_id: u32,
        media_relationship_id: &str,
        poster_relationship_id: Option<&str>,
    ) -> Result<String> {
        self.validate()?;
        if shape_id == 0 || media_relationship_id.is_empty() {
            return Err(Error::Invalid(
                "media shape ID and relationship ID are required".to_string(),
            ));
        }
        let default_name = if self.kind() == Kind::Audio {
            "Audio"
        } else {
            "Video"
        };
        let name = self.name.as_deref().unwrap_or(default_name);
        let mut xml = String::with_capacity(2 * 1024);
        write!(
            xml,
            r#"<p:pic xmlns:p="{p}" xmlns:a="{a}" xmlns:r="{r}"><p:nvPicPr><p:cNvPr id="{shape_id}" name="{}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:{}File r:link="{}"/><p:extLst><p:ext uri="{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="{}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill>"#,
            escape(name),
            if self.kind() == Kind::Audio { "audio" } else { "video" },
            escape(media_relationship_id),
            escape(media_relationship_id),
            p = String::from_utf8_lossy(super::embedded::PML),
            a = String::from_utf8_lossy(super::embedded::DML),
            r = String::from_utf8_lossy(super::embedded::REL),
        )
        .map_err(|_err| Error::Write)?;
        if let Some(relationship_id) = poster_relationship_id {
            write!(xml, r#"<a:blip r:embed="{}"/>"#, escape(relationship_id))
                .map_err(|_err| Error::Write)?;
        } else {
            xml.push_str("<a:blip/>");
        }
        write!(
            xml,
            r#"<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#,
            self.x, self.y, self.width, self.height
        )
        .map_err(|_err| Error::Write)?;
        Ok(xml)
    }
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn detects_formats_without_copying_the_payload() {
        let item = Item::new(b"ID3\x04\x00\x00".to_vec(), 1, 2, 3, 4);
        assert_eq!(item.format, Format::Mp3);
        let clone = item.clone();
        assert!(item.data.shares_with(&clone.data));
    }

    #[test]
    fn emits_a_self_contained_media_picture() {
        let item = Item::with_format(vec![1], Format::Mp4, 1, 2, 3, 4);
        let xml = item.to_shape_xml(2, "rIdMedia", Some("rIdPoster")).unwrap();
        assert!(xml.contains("videoFile"));
        assert!(xml.contains("rIdPoster"));
    }

    #[test]
    fn rejects_unbounded_frames() {
        let item = Item::with_format(vec![], Format::Unknown, 0, 0, 0, 1);
        assert!(item.validate().is_err());
    }
}
