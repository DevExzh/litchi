//! Audio and Video media support for PPTX presentations.
//!
//! This module provides types for embedding and referencing audio and video
//! content in PowerPoint presentations.

use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

/// Media type enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MediaType {
    /// Audio file (mp3, wav, etc.)
    Audio,
    /// Video file (mp4, wmv, etc.)
    Video,
}

impl MediaType {
    /// Get the MIME type prefix for this media type.
    pub fn mime_prefix(&self) -> &'static str {
        match self {
            MediaType::Audio => "audio",
            MediaType::Video => "video",
        }
    }

    /// Get the relationship type URL for this media type.
    pub fn relationship_type(&self) -> &'static str {
        match self {
            MediaType::Audio => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
            },
            MediaType::Video => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
            },
        }
    }
}

/// Audio/Video format enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MediaFormat {
    /// MP3 audio
    Mp3,
    /// WAV audio
    Wav,
    /// WMA audio
    Wma,
    /// M4A audio
    M4a,
    /// MP4 video
    Mp4,
    /// WMV video
    Wmv,
    /// AVI video
    Avi,
    /// MOV video
    Mov,
    /// Unknown format
    Unknown,
}

impl MediaFormat {
    /// Detect format from file extension.
    pub fn from_extension(ext: &str) -> Self {
        match ext.to_lowercase().as_str() {
            "mp3" => MediaFormat::Mp3,
            "wav" => MediaFormat::Wav,
            "wma" => MediaFormat::Wma,
            "m4a" => MediaFormat::M4a,
            "mp4" => MediaFormat::Mp4,
            "wmv" => MediaFormat::Wmv,
            "avi" => MediaFormat::Avi,
            "mov" => MediaFormat::Mov,
            _ => MediaFormat::Unknown,
        }
    }

    /// Detect format from file bytes (magic number detection).
    pub fn detect_from_bytes(data: &[u8]) -> Self {
        if data.len() < 12 {
            return MediaFormat::Unknown;
        }

        // ID3 tag (MP3)
        if data.starts_with(b"ID3") {
            return MediaFormat::Mp3;
        }

        // MP3 frame sync
        if data.len() >= 2 && data[0] == 0xFF && (data[1] & 0xE0) == 0xE0 {
            return MediaFormat::Mp3;
        }

        // RIFF/WAV
        if data.starts_with(b"RIFF") && data.len() >= 12 && &data[8..12] == b"WAVE" {
            return MediaFormat::Wav;
        }

        // RIFF/AVI
        if data.starts_with(b"RIFF") && data.len() >= 12 && &data[8..12] == b"AVI " {
            return MediaFormat::Avi;
        }

        // MP4/M4A/MOV (ftyp box)
        if data.len() >= 8 && &data[4..8] == b"ftyp" {
            // Check brand
            if data.len() >= 12 {
                let brand = &data[8..12];
                if brand == b"M4A " || brand == b"M4B " {
                    return MediaFormat::M4a;
                }
                if brand == b"qt  " || brand == b"moov" {
                    return MediaFormat::Mov;
                }
                // Most other brands are MP4
                return MediaFormat::Mp4;
            }
            return MediaFormat::Mp4;
        }

        // ASF/WMV/WMA header
        if data.len() >= 16
            && data[0..4] == [0x30, 0x26, 0xB2, 0x75]
            && data[4..8] == [0x8E, 0x66, 0xCF, 0x11]
        {
            // Could be WMV or WMA, default to WMV
            return MediaFormat::Wmv;
        }

        MediaFormat::Unknown
    }

    /// Get the MIME type for this format.
    pub fn mime_type(&self) -> &'static str {
        match self {
            MediaFormat::Mp3 => "audio/mpeg",
            MediaFormat::Wav => "audio/wav",
            MediaFormat::Wma => "audio/x-ms-wma",
            MediaFormat::M4a => "audio/mp4",
            MediaFormat::Mp4 => "video/mp4",
            MediaFormat::Wmv => "video/x-ms-wmv",
            MediaFormat::Avi => "video/avi",
            MediaFormat::Mov => "video/quicktime",
            MediaFormat::Unknown => "application/octet-stream",
        }
    }

    /// Get the file extension for this format.
    pub fn extension(&self) -> &'static str {
        match self {
            MediaFormat::Mp3 => "mp3",
            MediaFormat::Wav => "wav",
            MediaFormat::Wma => "wma",
            MediaFormat::M4a => "m4a",
            MediaFormat::Mp4 => "mp4",
            MediaFormat::Wmv => "wmv",
            MediaFormat::Avi => "avi",
            MediaFormat::Mov => "mov",
            MediaFormat::Unknown => "bin",
        }
    }

    /// Get the media type (audio or video) for this format.
    pub fn media_type(&self) -> MediaType {
        match self {
            MediaFormat::Mp3 | MediaFormat::Wav | MediaFormat::Wma | MediaFormat::M4a => {
                MediaType::Audio
            },
            MediaFormat::Mp4 | MediaFormat::Wmv | MediaFormat::Avi | MediaFormat::Mov => {
                MediaType::Video
            },
            MediaFormat::Unknown => MediaType::Video, // Default to video
        }
    }
}

/// A media element (audio or video) that can be embedded in a slide.
#[derive(Debug, Clone)]
pub struct Media {
    /// Media data (file content)
    pub data: Vec<u8>,
    /// Media format
    pub format: MediaFormat,
    /// X position in EMUs
    pub x: i64,
    /// Y position in EMUs
    pub y: i64,
    /// Width in EMUs
    pub width: i64,
    /// Height in EMUs
    pub height: i64,
    /// Optional name/description
    pub name: Option<String>,
    /// Whether to loop playback
    pub loop_playback: bool,
    /// Whether to auto-play
    pub auto_play: bool,
    /// Whether to hide during show (audio background)
    pub hide_during_show: bool,
}

impl Media {
    /// Create a new media element.
    pub fn new(data: Vec<u8>, x: i64, y: i64, width: i64, height: i64) -> Self {
        let format = MediaFormat::detect_from_bytes(&data);
        Self {
            data,
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

    /// Create a new media element with explicit format.
    pub fn with_format(
        data: Vec<u8>,
        format: MediaFormat,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> Self {
        Self {
            data,
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

    /// Set the name/description.
    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = Some(name.into());
    }

    /// Enable loop playback.
    pub fn with_loop(mut self) -> Self {
        self.loop_playback = true;
        self
    }

    /// Enable auto-play.
    pub fn with_auto_play(mut self) -> Self {
        self.auto_play = true;
        self
    }

    /// Hide the media icon during slideshow (for background audio).
    pub fn with_hide_during_show(mut self) -> Self {
        self.hide_during_show = true;
        self
    }

    /// Get the media type (audio or video).
    pub fn media_type(&self) -> MediaType {
        self.format.media_type()
    }

    /// Generate the picture shape XML for video/audio placeholder.
    ///
    /// # Arguments
    /// * `shape_id` - Shape ID
    /// * `media_rel_id` - Relationship ID for the media file
    /// * `image_rel_id` - Relationship ID for the poster image (optional)
    pub fn to_shape_xml(
        &self,
        shape_id: u32,
        media_rel_id: &str,
        image_rel_id: Option<&str>,
    ) -> Result<String> {
        let mut xml = String::with_capacity(2048);

        let name = self
            .name
            .as_deref()
            .unwrap_or(if self.media_type() == MediaType::Audio {
                "Audio"
            } else {
                "Video"
            });

        // For video, use p:pic with video extension
        // For audio, can use p:sp with audio icon or p:pic
        xml.push_str("<p:pic>");

        // Non-visual properties
        xml.push_str("<p:nvPicPr>");
        write!(
            xml,
            r#"<p:cNvPr id="{}" name="{}"/>"#,
            shape_id,
            escape_xml(name)
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;

        xml.push_str("<p:cNvPicPr>");
        xml.push_str(r#"<a:picLocks noChangeAspect="1"/>"#);
        xml.push_str("</p:cNvPicPr>");

        // Non-visual properties with media extension
        xml.push_str("<p:nvPr>");

        // Video/Audio frame
        match self.media_type() {
            MediaType::Video => {
                write!(
                    xml,
                    r#"<a:videoFile xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:link="{}"/>"#,
                    media_rel_id
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            },
            MediaType::Audio => {
                write!(
                    xml,
                    r#"<a:audioFile xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:link="{}"/>"#,
                    media_rel_id
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            },
        }

        // Media extension for playback settings
        xml.push_str("<p:extLst>");
        xml.push_str(r#"<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">"#);
        xml.push_str(r#"<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed=""#);
        xml.push_str(media_rel_id);
        xml.push_str(r#"">"#);

        // Trim element with playback options
        xml.push_str("<p14:trim/>");
        xml.push_str("</p14:media>");
        xml.push_str("</p:ext>");
        xml.push_str("</p:extLst>");

        xml.push_str("</p:nvPr>");
        xml.push_str("</p:nvPicPr>");

        // Blip fill (poster image or placeholder)
        xml.push_str("<p:blipFill>");
        if let Some(img_rid) = image_rel_id {
            write!(xml, r#"<a:blip r:embed="{}"/>"#, img_rid)
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        } else {
            // Use empty blip for placeholder
            xml.push_str(r#"<a:blip/>"#);
        }
        xml.push_str("<a:stretch><a:fillRect/></a:stretch>");
        xml.push_str("</p:blipFill>");

        // Shape properties
        xml.push_str("<p:spPr>");
        xml.push_str("<a:xfrm>");
        write!(xml, r#"<a:off x="{}" y="{}"/>"#, self.x, self.y)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        write!(xml, r#"<a:ext cx="{}" cy="{}"/>"#, self.width, self.height)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str("</a:xfrm>");
        xml.push_str(r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#);
        xml.push_str("</p:spPr>");

        xml.push_str("</p:pic>");

        Ok(xml)
    }
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_format_detection_mp3() {
        // ID3v2 header (needs at least 12 bytes)
        let data = b"ID3\x04\x00\x00\x00\x00\x00\x00\x00\x00";
        assert_eq!(MediaFormat::detect_from_bytes(data), MediaFormat::Mp3);
    }

    #[test]
    fn test_format_detection_wav() {
        let data = b"RIFF\x00\x00\x00\x00WAVEfmt ";
        assert_eq!(MediaFormat::detect_from_bytes(data), MediaFormat::Wav);
    }

    #[test]
    fn test_format_detection_mp4() {
        let data = b"\x00\x00\x00\x1cftypmp42\x00\x00\x00\x00";
        assert_eq!(MediaFormat::detect_from_bytes(data), MediaFormat::Mp4);
    }

    #[test]
    fn test_media_type() {
        assert_eq!(MediaFormat::Mp3.media_type(), MediaType::Audio);
        assert_eq!(MediaFormat::Mp4.media_type(), MediaType::Video);
        assert_eq!(MediaFormat::Wav.media_type(), MediaType::Audio);
        assert_eq!(MediaFormat::Wmv.media_type(), MediaType::Video);
    }

    #[test]
    fn test_mime_type() {
        assert_eq!(MediaFormat::Mp3.mime_type(), "audio/mpeg");
        assert_eq!(MediaFormat::Mp4.mime_type(), "video/mp4");
    }
}
