//! WordprocessingML web-settings support.

use crate::docx::namespace::{is_wordprocessing_namespace, word_attribute_value};
use crate::error::{OoxmlError, Result};
use litchi_opc::part::Part;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

/// Scalar settings from a Word `webSettings.xml` part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WebSettings {
    encoding: Option<String>,
    optimize_for_browser: Option<bool>,
    rely_on_vml: Option<bool>,
    allow_png: Option<bool>,
    do_not_rely_on_css: Option<bool>,
    do_not_save_as_single_file: Option<bool>,
    do_not_organize_in_folder: Option<bool>,
    do_not_use_long_file_names: Option<bool>,
    pixels_per_inch: Option<String>,
    target_screen_size: Option<TargetScreenSize>,
    save_smart_tags_as_xml: Option<bool>,
}

/// A spec-defined target size for generated web pages.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TargetScreenSize {
    Pixels544x376,
    Pixels640x480,
    Pixels720x512,
    Pixels800x600,
    Pixels1024x768,
    Pixels1152x882,
    Pixels1152x900,
    Pixels1280x1024,
    Pixels1600x1200,
    Pixels1800x1440,
    Pixels1920x1200,
}

impl TargetScreenSize {
    fn from_xml(value: &str) -> Option<Self> {
        match value {
            "544x376" => Some(Self::Pixels544x376),
            "640x480" => Some(Self::Pixels640x480),
            "720x512" => Some(Self::Pixels720x512),
            "800x600" => Some(Self::Pixels800x600),
            "1024x768" => Some(Self::Pixels1024x768),
            "1152x882" => Some(Self::Pixels1152x882),
            "1152x900" => Some(Self::Pixels1152x900),
            "1280x1024" => Some(Self::Pixels1280x1024),
            "1600x1200" => Some(Self::Pixels1600x1200),
            "1800x1440" => Some(Self::Pixels1800x1440),
            "1920x1200" => Some(Self::Pixels1920x1200),
            _ => None,
        }
    }

    /// Return the OOXML lexical representation.
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Pixels544x376 => "544x376",
            Self::Pixels640x480 => "640x480",
            Self::Pixels720x512 => "720x512",
            Self::Pixels800x600 => "800x600",
            Self::Pixels1024x768 => "1024x768",
            Self::Pixels1152x882 => "1152x882",
            Self::Pixels1152x900 => "1152x900",
            Self::Pixels1280x1024 => "1280x1024",
            Self::Pixels1600x1200 => "1600x1200",
            Self::Pixels1800x1440 => "1800x1440",
            Self::Pixels1920x1200 => "1920x1200",
        }
    }
}

impl WebSettings {
    /// Return the requested output encoding, if declared.
    pub fn encoding(&self) -> Option<&str> {
        self.encoding.as_deref()
    }

    /// Return the browser-optimization setting, preserving absence.
    pub fn optimize_for_browser(&self) -> Option<bool> {
        self.optimize_for_browser
    }

    /// Return whether web output should rely on VML, preserving absence.
    pub fn rely_on_vml(&self) -> Option<bool> {
        self.rely_on_vml
    }

    /// Return whether PNG images are allowed, preserving absence.
    pub fn allow_png(&self) -> Option<bool> {
        self.allow_png
    }

    /// Return whether web output should avoid CSS, preserving absence.
    pub fn do_not_rely_on_css(&self) -> Option<bool> {
        self.do_not_rely_on_css
    }

    /// Return whether web output should use multiple files, preserving absence.
    pub fn do_not_save_as_single_file(&self) -> Option<bool> {
        self.do_not_save_as_single_file
    }

    /// Return whether supporting files should avoid a folder, preserving absence.
    pub fn do_not_organize_in_folder(&self) -> Option<bool> {
        self.do_not_organize_in_folder
    }

    /// Return whether web output should avoid long file names, preserving absence.
    pub fn do_not_use_long_file_names(&self) -> Option<bool> {
        self.do_not_use_long_file_names
    }

    /// Return the arbitrary-precision XML integer for web-output pixel density.
    pub fn pixels_per_inch(&self) -> Option<&str> {
        self.pixels_per_inch.as_deref()
    }

    /// Return the target screen size, if declared.
    pub fn target_screen_size(&self) -> Option<TargetScreenSize> {
        self.target_screen_size
    }

    /// Return whether smart tags should be saved as XML, preserving absence.
    pub fn save_smart_tags_as_xml(&self) -> Option<bool> {
        self.save_smart_tags_as_xml
    }

    pub(crate) fn extract_from_part(part: &dyn Part) -> Result<Self> {
        Self::extract_from_xml(part.blob())
    }

    fn extract_from_xml(xml: &[u8]) -> Result<Self> {
        let mut reader = NsReader::from_reader(xml);
        let mut settings = Self::default();
        let mut depth = 0usize;
        let mut saw_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "Word web-settings XML nesting is too deep".into(),
                        )
                    })?;
                    if depth == 1 {
                        validate_root(&namespace, &element, saw_root)?;
                        saw_root = true;
                    } else if depth == 2 && saw_root && is_wordprocessing_namespace(&namespace) {
                        parse_setting(&element, decoder, &resolver, &mut settings)?;
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "Word web-settings XML nesting is too deep".into(),
                        )
                    })?;
                    if child_depth == 1 {
                        validate_root(&namespace, &element, saw_root)?;
                        saw_root = true;
                    } else if child_depth == 2
                        && saw_root
                        && is_wordprocessing_namespace(&namespace)
                    {
                        parse_setting(&element, decoder, &resolver, &mut settings)?;
                    }
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OoxmlError::InvalidFormat("invalid Word web-settings XML nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated Word web-settings XML".into(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }

        if !saw_root {
            return Err(OoxmlError::InvalidFormat(
                "web-settings part has no webSettings root".into(),
            ));
        }
        Ok(settings)
    }
}

fn validate_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    saw_root: bool,
) -> Result<()> {
    if saw_root
        || !is_wordprocessing_namespace(namespace)
        || element.local_name().as_ref() != b"webSettings"
    {
        return Err(OoxmlError::InvalidFormat(
            "web-settings part has an invalid or trailing root element".into(),
        ));
    }
    Ok(())
}

fn parse_setting(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    settings: &mut WebSettings,
) -> Result<()> {
    match element.local_name().as_ref() {
        b"encoding" => set_once(
            &mut settings.encoding,
            required_value(element, decoder, resolver, "web encoding")?,
            "encoding",
        ),
        b"optimizeForBrowser" => set_on_off(
            &mut settings.optimize_for_browser,
            element,
            decoder,
            resolver,
            "optimizeForBrowser",
        ),
        b"relyOnVML" => set_on_off(
            &mut settings.rely_on_vml,
            element,
            decoder,
            resolver,
            "relyOnVML",
        ),
        b"allowPNG" => set_on_off(
            &mut settings.allow_png,
            element,
            decoder,
            resolver,
            "allowPNG",
        ),
        b"doNotRelyOnCSS" => set_on_off(
            &mut settings.do_not_rely_on_css,
            element,
            decoder,
            resolver,
            "doNotRelyOnCSS",
        ),
        b"doNotSaveAsSingleFile" => set_on_off(
            &mut settings.do_not_save_as_single_file,
            element,
            decoder,
            resolver,
            "doNotSaveAsSingleFile",
        ),
        b"doNotOrganizeInFolder" => set_on_off(
            &mut settings.do_not_organize_in_folder,
            element,
            decoder,
            resolver,
            "doNotOrganizeInFolder",
        ),
        b"doNotUseLongFileNames" => set_on_off(
            &mut settings.do_not_use_long_file_names,
            element,
            decoder,
            resolver,
            "doNotUseLongFileNames",
        ),
        b"pixelsPerInch" => {
            let value = required_value(element, decoder, resolver, "pixels per inch")?;
            let value = value.trim();
            if !is_xml_integer(value) {
                return Err(OoxmlError::InvalidFormat(format!(
                    "invalid pixels-per-inch value '{value}'"
                )));
            }
            set_once(
                &mut settings.pixels_per_inch,
                value.to_owned(),
                "pixelsPerInch",
            )
        },
        b"targetScreenSz" => {
            let value = required_value(element, decoder, resolver, "target screen size")?;
            let value = TargetScreenSize::from_xml(&value).ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("invalid target-screen-size value '{value}'"))
            })?;
            set_once(&mut settings.target_screen_size, value, "targetScreenSz")
        },
        b"saveSmartTagsAsXml" => set_on_off(
            &mut settings.save_smart_tags_as_xml,
            element,
            decoder,
            resolver,
            "saveSmartTagsAsXml",
        ),
        _ => Ok(()),
    }
}

fn is_xml_integer(value: &str) -> bool {
    let digits = value
        .strip_prefix('+')
        .or_else(|| value.strip_prefix('-'))
        .unwrap_or(value);
    !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit())
}

fn required_value(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<String> {
    word_attribute_value(element, b"val", decoder, resolver)?
        .ok_or_else(|| OoxmlError::InvalidFormat(format!("Word {description} value is required")))
}

fn set_on_off(
    slot: &mut Option<bool>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    description: &str,
) -> Result<()> {
    let value = match word_attribute_value(element, b"val", decoder, resolver)? {
        Some(value) => match value.as_str() {
            "true" | "1" | "on" => true,
            "false" | "0" | "off" => false,
            _ => {
                return Err(OoxmlError::InvalidFormat(format!(
                    "invalid Word on/off value '{value}'"
                )));
            },
        },
        None => true,
    };
    set_once(slot, value, description)
}

fn set_once<T>(slot: &mut Option<T>, value: T, description: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "duplicate Word web setting '{description}'"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_all_scalar_web_settings_with_strict_namespaces() {
        let xml = br#"<s:webSettings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:false="urn:not-wordprocessingml">
            <s:encoding s:val="utf-8"/>
            <s:optimizeForBrowser s:val="on"/>
            <s:relyOnVML s:val="off"/>
            <s:allowPNG/>
            <s:doNotRelyOnCSS s:val="0"/>
            <s:doNotSaveAsSingleFile s:val="1"/>
            <s:doNotOrganizeInFolder s:val="false"/>
            <s:doNotUseLongFileNames s:val="true"/>
            <s:pixelsPerInch s:val=" 123456789012345678901234567890 "/>
            <s:targetScreenSz s:val="1920x1200"/>
            <s:saveSmartTagsAsXml s:val="on"/>
            <false:saveSmartTagsAsXml false:val="off"/>
        </s:webSettings>"#;

        let settings = WebSettings::extract_from_xml(xml).unwrap();
        assert_eq!(settings.encoding(), Some("utf-8"));
        assert_eq!(settings.optimize_for_browser(), Some(true));
        assert_eq!(settings.rely_on_vml(), Some(false));
        assert_eq!(settings.allow_png(), Some(true));
        assert_eq!(settings.do_not_rely_on_css(), Some(false));
        assert_eq!(settings.do_not_save_as_single_file(), Some(true));
        assert_eq!(settings.do_not_organize_in_folder(), Some(false));
        assert_eq!(settings.do_not_use_long_file_names(), Some(true));
        assert_eq!(
            settings.pixels_per_inch(),
            Some("123456789012345678901234567890")
        );
        assert_eq!(
            settings.target_screen_size(),
            Some(TargetScreenSize::Pixels1920x1200)
        );
        assert_eq!(settings.save_smart_tags_as_xml(), Some(true));
    }

    #[test]
    fn rejects_invalid_or_duplicate_scalar_web_settings() {
        let missing_value = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pixelsPerInch/></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(missing_value).is_err());

        let invalid_on_off = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:saveSmartTagsAsXml w:val="maybe"/></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(invalid_on_off).is_err());

        let duplicate = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:allowPNG/><w:allowPNG/></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(duplicate).is_err());

        let invalid_screen = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:targetScreenSz w:val="1366x768"/></w:webSettings>"#;
        assert!(WebSettings::extract_from_xml(invalid_screen).is_err());
    }
}
