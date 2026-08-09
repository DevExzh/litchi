//! `PowerPoint` 10 document metadata and privacy settings.

use super::package::{Error, Result};
use super::records::Record;
use super::slide_round_trip::{EmbeddedXmlPackage, parse_embedded_xml_package};
use crate::consts::RecordType;
use litchi_opc::constants::content_type;

const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";

/// Square-grid spacing stored by a `PowerPoint` 10 document extension.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct GridSpacing {
    /// Horizontal and vertical spacing in `PowerPoint` grid units.
    pub grid_units: i32,
}

/// Arrangement of pictures on photo-album slides.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PhotoAlbumLayout {
    FitToSlide,
    OnePicture,
    TwoPictures,
    FourPictures,
    OnePictureAndTitle,
    TwoPicturesAndTitle,
    FourPicturesAndTitle,
}

/// Shape drawn or cropped around pictures in a photo album.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PhotoAlbumFrameShape {
    Rectangle,
    RoundedRectangle,
    Beveled,
    Oval,
    Octagon,
    Cross,
    Plaque,
}

/// `PowerPoint` 10 photo-album display preferences.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PhotoAlbumSettings {
    /// Whether pictures are displayed as grayscale graphics.
    pub use_grayscale: bool,
    /// Whether each picture has a caption beneath it.
    pub has_captions: bool,
    /// Preferred picture arrangement.
    pub layout: PhotoAlbumLayout,
    /// Undefined byte retained for lossless inspection.
    pub unused: u8,
    /// Preferred picture frame shape.
    pub frame_shape: PhotoAlbumFrameShape,
}

/// Metadata and privacy settings stored in the `___PPT10` document extension.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`DocumentProperties10` is the established public API name; renaming it would break downstream crates"
)]
pub struct DocumentProperties10 {
    /// Optional copyright notice.
    pub copyright: Option<String>,
    /// Optional document keywords.
    pub keywords: Option<String>,
    /// Optional password required for modify access.
    ///
    /// MS-PPT also requires presentations containing this field to be encrypted; this field alone
    /// does not establish that the surrounding encryption was validated.
    pub modify_password: Option<String>,
    /// Whether personally identifiable information is removed when saving.
    ///
    /// `None` means that the optional `FilterPrivacyFlags10Atom` is absent.
    pub remove_personally_identifiable_information: Option<bool>,
    /// Optional square-grid spacing used for alignment and positioning cues.
    pub grid_spacing: Option<GridSpacing>,
    /// Optional photo-album display preferences.
    pub photo_album: Option<PhotoAlbumSettings>,
}

/// Document-level settings stored in the `___PPT12` extension.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`DocumentProperties12` is the established public API name; renaming it would break downstream crates"
)]
pub struct DocumentProperties12 {
    /// Whether pictures are automatically compressed when the presentation is saved.
    ///
    /// `None` means that the optional `RoundTripDocFlags12Atom` is absent.
    pub compress_pictures_on_save: Option<bool>,
    /// Validated embedded custom table styles.
    pub custom_table_styles: Option<CustomTableStyles>,
}

/// `PowerPoint` 12 custom table styles stored directly in the Document container.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomTableStyles {
    /// Record version retained because MS-PPT recommends, but does not require, version zero.
    pub record_version: u16,
    /// Validated package containing the `DrawingML` `tblStyleLst` part.
    pub package: EmbeddedXmlPackage,
}

impl DocumentProperties10 {
    /// Discover and parse document properties from every `___PPT10` programmable tag below `root`.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(root: &Record) -> Result<Self> {
        let mut properties = Self::default();
        for record in root.versioned_binary_tag_records(10)? {
            match (record.record_type, record.instance) {
                (RecordType::CString, 1) => {
                    if properties.copyright.is_some() {
                        return Err(Error::Corrupted(
                            "PowerPoint 10 document extension contains duplicate copyright"
                                .to_string(),
                        ));
                    }
                    properties.copyright =
                        Some(parse_document_string(&record, 1, "CopyrightAtom")?);
                },
                (RecordType::CString, 2) => {
                    if properties.keywords.is_some() {
                        return Err(Error::Corrupted(
                            "PowerPoint 10 document extension contains duplicate keywords"
                                .to_string(),
                        ));
                    }
                    properties.keywords = Some(parse_document_string(&record, 2, "KeywordsAtom")?);
                },
                (RecordType::CString, 3) => {
                    if properties.modify_password.is_some() {
                        return Err(Error::Corrupted(
                            "PowerPoint 10 document extension contains duplicate modify password"
                                .to_string(),
                        ));
                    }
                    properties.modify_password =
                        Some(parse_document_string(&record, 3, "ModifyPasswordAtom")?);
                },
                (RecordType::FilterPrivacyFlags10Atom, _) => {
                    if properties
                        .remove_personally_identifiable_information
                        .is_some()
                    {
                        return Err(Error::Corrupted(
                            "PowerPoint 10 document extension contains duplicate privacy flags"
                                .to_string(),
                        ));
                    }
                    properties.remove_personally_identifiable_information =
                        Some(parse_privacy_flags(&record)?);
                },
                (RecordType::GridSpacing10Atom, _) => {
                    if properties.grid_spacing.is_some() {
                        return Err(Error::Corrupted(
                            "PowerPoint 10 document extension contains duplicate grid spacing"
                                .to_string(),
                        ));
                    }
                    properties.grid_spacing = Some(parse_grid_spacing(&record)?);
                },
                (RecordType::PhotoAlbumInfo10Atom, _) => {
                    if properties.photo_album.is_some() {
                        return Err(Error::Corrupted(
                            "PowerPoint 10 document extension contains duplicate photo album settings"
                                .to_string(),
                        ));
                    }
                    properties.photo_album = Some(parse_photo_album(&record)?);
                },
                _ => {},
            }
        }
        Ok(properties)
    }
}

impl DocumentProperties12 {
    /// Discover and parse document properties from every `___PPT12` programmable tag below `root`.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(root: &Record) -> Result<Self> {
        let mut properties = Self::default();
        for record in &root.children {
            if record.record_type != RecordType::RoundTripCustomTableStyles12Atom {
                continue;
            }
            if properties.custom_table_styles.is_some() {
                return Err(Error::Corrupted(
                    "Document contains duplicate RoundTripCustomTableStyles12Atom records"
                        .to_string(),
                ));
            }
            if record.instance != 0 || record.data_length as usize != record.data.len() {
                return Err(Error::Corrupted(
                    "RoundTripCustomTableStyles12Atom has an invalid record header or size"
                        .to_string(),
                ));
            }
            properties.custom_table_styles = Some(CustomTableStyles {
                record_version: record.version,
                package: parse_embedded_xml_package(
                    &record.data,
                    "RoundTripCustomTableStyles12Atom",
                    content_type::PML_TABLE_STYLES,
                    DRAWINGML_NAMESPACE,
                    b"tblStyleLst",
                )?,
            });
        }
        for record in root.versioned_binary_tag_records(12)? {
            if record.record_type != RecordType::RoundTripDocFlags12Atom {
                continue;
            }
            if properties.compress_pictures_on_save.is_some() {
                return Err(Error::Corrupted(
                    "PowerPoint 12 document extension contains duplicate document flags"
                        .to_string(),
                ));
            }
            if record.version != 0 || record.instance != 0 || record.data.len() != 1 {
                return Err(Error::Corrupted(
                    "RoundTripDocFlags12Atom has an invalid record header or size".to_string(),
                ));
            }
            if record.data[0] & 0xfe != 0 {
                return Err(Error::Corrupted(
                    "RoundTripDocFlags12Atom has nonzero reserved bits".to_string(),
                ));
            }
            properties.compress_pictures_on_save = Some(record.data[0] & 1 != 0);
        }
        Ok(properties)
    }
}

fn parse_document_string(record: &Record, instance: u16, name: &str) -> Result<String> {
    if record.record_type != RecordType::CString
        || record.version != 0
        || record.instance != instance
        || record.data.len() > 510
        || record.data.len() & 1 != 0
    {
        return Err(Error::Corrupted(format!(
            "{name} has an invalid record header or size"
        )));
    }

    let mut units = Vec::with_capacity(record.data.len() / 2);
    for bytes in record.data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if unit == 0 {
            break;
        }
        if matches!(unit, 0x0001..=0x001f | 0x007f..=0x009f) {
            return Err(Error::Corrupted(format!(
                "{name} contains a non-printable character"
            )));
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_err| Error::Corrupted(format!("{name} contains invalid UTF-16")))
}

fn parse_privacy_flags(record: &Record) -> Result<bool> {
    if record.record_type != RecordType::FilterPrivacyFlags10Atom
        || record.version != 0
        || record.instance != 0
        || record.data.len() != 4
    {
        return Err(Error::Corrupted(
            "FilterPrivacyFlags10Atom has an invalid record header or size".to_string(),
        ));
    }
    let flags = u32::from_le_bytes([
        record.data[0],
        record.data[1],
        record.data[2],
        record.data[3],
    ]);
    if flags & !1 != 0 {
        return Err(Error::Corrupted(
            "FilterPrivacyFlags10Atom has nonzero reserved bits".to_string(),
        ));
    }
    Ok(flags & 1 != 0)
}

fn parse_grid_spacing(record: &Record) -> Result<GridSpacing> {
    if record.record_type != RecordType::GridSpacing10Atom
        || record.version != 0
        || record.instance != 0
        || record.data.len() != 8
    {
        return Err(Error::Corrupted(
            "GridSpacing10Atom has an invalid record header or size".to_string(),
        ));
    }
    let x = i32::from_le_bytes(record.data[0..4].try_into().map_err(|_err| {
        Error::Corrupted("GridSpacing10Atom horizontal value is truncated".to_string())
    })?);
    let y = i32::from_le_bytes(record.data[4..8].try_into().map_err(|_err| {
        Error::Corrupted("GridSpacing10Atom vertical value is truncated".to_string())
    })?);
    if !(0x0000_5ab8..=0x0012_0000).contains(&x) || x != y {
        return Err(Error::Corrupted(
            "GridSpacing10Atom values are unequal or out of range".to_string(),
        ));
    }
    Ok(GridSpacing { grid_units: x })
}

fn parse_photo_album(record: &Record) -> Result<PhotoAlbumSettings> {
    if record.record_type != RecordType::PhotoAlbumInfo10Atom
        || record.version != 0
        || record.instance != 0
        || record.data.len() != 6
    {
        return Err(Error::Corrupted(
            "PhotoAlbumInfo10Atom has an invalid record header or size".to_string(),
        ));
    }
    let layout = match record.data[2] {
        0 => PhotoAlbumLayout::FitToSlide,
        1 => PhotoAlbumLayout::OnePicture,
        2 => PhotoAlbumLayout::TwoPictures,
        3 => PhotoAlbumLayout::FourPictures,
        4 => PhotoAlbumLayout::OnePictureAndTitle,
        5 => PhotoAlbumLayout::TwoPicturesAndTitle,
        6 => PhotoAlbumLayout::FourPicturesAndTitle,
        _ => {
            return Err(Error::Corrupted(
                "PhotoAlbumInfo10Atom has an invalid layout".to_string(),
            ));
        },
    };
    let frame_shape = match u16::from_le_bytes([record.data[4], record.data[5]]) {
        0 => PhotoAlbumFrameShape::Rectangle,
        1 => PhotoAlbumFrameShape::RoundedRectangle,
        2 => PhotoAlbumFrameShape::Beveled,
        3 => PhotoAlbumFrameShape::Oval,
        4 => PhotoAlbumFrameShape::Octagon,
        5 => PhotoAlbumFrameShape::Cross,
        6 => PhotoAlbumFrameShape::Plaque,
        _ => {
            return Err(Error::Corrupted(
                "PhotoAlbumInfo10Atom has an invalid frame shape".to_string(),
            ));
        },
    };
    Ok(PhotoAlbumSettings {
        use_grayscale: record.data[0] != 0,
        has_captions: record.data[1] != 0,
        layout,
        unused: record.data[3],
        frame_shape,
    })
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::Package;
    use litchi_opc::{OpcPackage, PackURI, XmlPart};
    use std::io::Cursor;

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn utf16(value: &str) -> Vec<u8> {
        value.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn prog_tags_record(version: u8, blob_payload: &[u8]) -> Record {
        let tag_name = utf16(&format!("___PPT{version}"));
        let name = record_bytes(0, 0, 4026, &tag_name);
        let blob = record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
        Record {
            record_type: RecordType::ProgTags,
            record_type_raw: 0x1388,
            version: 0x0f,
            instance: 0,
            data_length: u32::try_from(tag.len()).unwrap(),
            data: tag,
            children: Vec::new(),
        }
    }

    fn root(children: Vec<Record>) -> Record {
        Record {
            record_type: RecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    fn table_styles_package(content_type: &str, xml: &[u8]) -> Vec<u8> {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(XmlPart::new(
            PackURI::new("/tableStyles.xml").unwrap(),
            content_type.to_string(),
            xml.to_vec(),
        )));
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
                .to_string(),
            "tableStyles.xml".to_string(),
            "rId1".to_string(),
            false,
        );
        let mut output = Cursor::new(Vec::new());
        package.to_stream(&mut output).unwrap();
        output.into_inner()
    }

    fn valid_table_styles_package() -> Vec<u8> {
        table_styles_package(
            content_type::PML_TABLE_STYLES,
            br#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>"#,
        )
    }

    fn table_styles_record(version: u16, instance: u16, data: &[u8]) -> Record {
        Record {
            record_type: RecordType::RoundTripCustomTableStyles12Atom,
            record_type_raw: RecordType::RoundTripCustomTableStyles12Atom.as_u16(),
            version,
            instance,
            data_length: u32::try_from(data.len()).unwrap(),
            data: data.to_vec(),
            children: Vec::new(),
        }
    }

    #[test]
    fn parses_powerpoint10_document_properties() {
        let mut records = record_bytes(0, 1, 4026, &utf16("Copyright © 2026"));
        records.extend_from_slice(&record_bytes(0, 2, 4026, &utf16("rust; ooxml")));
        records.extend_from_slice(&record_bytes(0, 3, 4026, &utf16("edit-only")));
        records.extend_from_slice(&record_bytes(0, 0, 0x36b0, &1u32.to_le_bytes()));
        let mut spacing = 0x0000_5ab8i32.to_le_bytes().to_vec();
        spacing.extend_from_slice(&0x0000_5ab8i32.to_le_bytes());
        records.extend_from_slice(&record_bytes(0, 0, 1037, &spacing));
        records.extend_from_slice(&record_bytes(0, 0, 0x36b2, &[2, 0, 6, 0xff, 6, 0]));

        let properties =
            DocumentProperties10::parse(&root(vec![prog_tags_record(10, &records)])).unwrap();

        assert_eq!(properties.copyright.as_deref(), Some("Copyright © 2026"));
        assert_eq!(properties.keywords.as_deref(), Some("rust; ooxml"));
        assert_eq!(properties.modify_password.as_deref(), Some("edit-only"));
        assert_eq!(
            properties.remove_personally_identifiable_information,
            Some(true)
        );
        assert_eq!(properties.grid_spacing.unwrap().grid_units, 0x0000_5ab8);
        let album = properties.photo_album.unwrap();
        assert!(album.use_grayscale);
        assert!(!album.has_captions);
        assert_eq!(album.layout, PhotoAlbumLayout::FourPicturesAndTitle);
        assert_eq!(album.unused, 0xff);
        assert_eq!(album.frame_shape, PhotoAlbumFrameShape::Plaque);
    }

    #[test]
    fn ignores_other_versions_and_preserves_absent_privacy_flags() {
        let copyright = record_bytes(0, 1, 4026, &utf16("old"));
        let keywords = record_bytes(0, 2, 4026, &utf16("current"));
        let password = record_bytes(0, 3, 4026, &utf16("old-password"));
        let properties = DocumentProperties10::parse(&root(vec![
            prog_tags_record(9, &[copyright, password].concat()),
            prog_tags_record(10, &keywords),
        ]))
        .unwrap();

        assert_eq!(properties.copyright, None);
        assert_eq!(properties.keywords.as_deref(), Some("current"));
        assert_eq!(properties.modify_password, None);
        assert_eq!(properties.remove_personally_identifiable_information, None);
        assert_eq!(properties.grid_spacing, None);
        assert_eq!(properties.photo_album, None);
    }

    #[test]
    fn rejects_malformed_document_properties() {
        let malformed = [
            record_bytes(1, 1, 4026, &utf16("copyright")),
            record_bytes(0, 1, 4026, b"A"),
            record_bytes(0, 2, 4026, &vec![0; 512]),
            record_bytes(0, 2, 4026, &[0x01, 0x00]),
            record_bytes(0, 2, 4026, &[0x00, 0xd8]),
            record_bytes(0, 3, 4026, &vec![0; 512]),
            record_bytes(0, 3, 4026, &[0x01, 0]),
            record_bytes(0, 1, 0x36b0, &0u32.to_le_bytes()),
            record_bytes(0, 0, 0x36b0, &[0, 0, 0]),
            record_bytes(0, 0, 0x36b0, &2u32.to_le_bytes()),
        ];
        for record in malformed {
            let document = root(vec![prog_tags_record(10, &record)]);
            assert!(DocumentProperties10::parse(&document).is_err());
        }
    }

    #[test]
    fn rejects_duplicate_document_properties() {
        for record in [
            record_bytes(0, 1, 4026, &utf16("copyright")),
            record_bytes(0, 2, 4026, &utf16("keywords")),
            record_bytes(0, 3, 4026, &utf16("password")),
            record_bytes(0, 0, 0x36b0, &0u32.to_le_bytes()),
            record_bytes(0, 0, 1037, &[0xb8, 0x5a, 0, 0, 0xb8, 0x5a, 0, 0]),
            record_bytes(0, 0, 0x36b2, &[0, 0, 0, 0, 0, 0]),
        ] {
            let mut duplicate = record.clone();
            duplicate.extend_from_slice(&record);
            let document = root(vec![prog_tags_record(10, &duplicate)]);
            assert!(DocumentProperties10::parse(&document).is_err());
        }
    }

    #[test]
    fn rejects_malformed_grid_and_photo_album_settings() {
        let malformed = [
            record_bytes(1, 0, 1037, &[0; 8]),
            record_bytes(0, 1, 1037, &[0; 8]),
            record_bytes(0, 0, 1037, &[0; 7]),
            record_bytes(0, 0, 1037, &[0xb7, 0x5a, 0, 0, 0xb7, 0x5a, 0, 0]),
            record_bytes(0, 0, 1037, &[0xb8, 0x5a, 0, 0, 0xb9, 0x5a, 0, 0]),
            record_bytes(0, 0, 1037, &[1, 0, 0x12, 0, 1, 0, 0x12, 0]),
            record_bytes(1, 0, 0x36b2, &[0; 6]),
            record_bytes(0, 1, 0x36b2, &[0; 6]),
            record_bytes(0, 0, 0x36b2, &[0; 5]),
            record_bytes(0, 0, 0x36b2, &[0, 0, 7, 0, 0, 0]),
            record_bytes(0, 0, 0x36b2, &[0, 0, 0, 0, 7, 0]),
        ];
        for record in malformed {
            let document = root(vec![prog_tags_record(10, &record)]);
            assert!(DocumentProperties10::parse(&document).is_err());
        }
    }

    #[test]
    fn parses_powerpoint12_document_flags_and_isolates_versions() {
        let flags = record_bytes(0, 0, 0x0425, &[1]);
        let ppt12_properties =
            DocumentProperties12::parse(&root(vec![prog_tags_record(12, &flags)])).unwrap();
        assert_eq!(ppt12_properties.compress_pictures_on_save, Some(true));

        let ppt11_properties =
            DocumentProperties12::parse(&root(vec![prog_tags_record(11, &flags)])).unwrap();
        assert_eq!(ppt11_properties.compress_pictures_on_save, None);
    }

    #[test]
    fn rejects_malformed_or_duplicate_powerpoint12_document_flags() {
        for flags in [
            record_bytes(1, 0, 0x0425, &[0]),
            record_bytes(0, 1, 0x0425, &[0]),
            record_bytes(0, 0, 0x0425, &[]),
            record_bytes(0, 0, 0x0425, &[2]),
        ] {
            let document = root(vec![prog_tags_record(12, &flags)]);
            assert!(DocumentProperties12::parse(&document).is_err());
        }

        let flags = record_bytes(0, 0, 0x0425, &[0]);
        let mut duplicate = flags.clone();
        duplicate.extend_from_slice(&flags);
        let document = root(vec![prog_tags_record(12, &duplicate)]);
        assert!(DocumentProperties12::parse(&document).is_err());
    }

    #[test]
    fn parses_custom_table_styles_and_retains_recommended_version_deviations() {
        let package = valid_table_styles_package();
        for version in [0, 0x0f] {
            let parsed =
                DocumentProperties12::parse(&root(vec![table_styles_record(version, 0, &package)]))
                    .unwrap();
            let styles = parsed.custom_table_styles.unwrap();
            assert_eq!(styles.record_version, version);
            assert_eq!(styles.package.data, package);
            assert_eq!(styles.package.part_count, 1);
            assert_eq!(styles.package.xml_part_name, "/tableStyles.xml");
            assert_eq!(parsed.compress_pictures_on_save, None);
        }
    }

    #[test]
    fn rejects_duplicate_or_malformed_custom_table_styles() {
        let package = valid_table_styles_package();
        for malformed in [
            table_styles_record(0, 1, &package),
            table_styles_record(0, 0, b"not a package"),
            table_styles_record(
                0,
                0,
                &table_styles_package(
                    "application/xml",
                    br#"<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
                ),
            ),
            table_styles_record(
                0,
                0,
                &table_styles_package(
                    content_type::PML_TABLE_STYLES,
                    br#"<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
                ),
            ),
        ] {
            assert!(DocumentProperties12::parse(&root(vec![malformed])).is_err());
        }
        let mut wrong_declared = table_styles_record(0, 0, &package);
        wrong_declared.data_length -= 1;
        assert!(DocumentProperties12::parse(&root(vec![wrong_declared])).is_err());

        let end_document = Record {
            record_type: RecordType::EndDocument,
            record_type_raw: RecordType::EndDocument.as_u16(),
            version: 0,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: Vec::new(),
        };
        assert!(
            DocumentProperties12::parse(&root(vec![
                table_styles_record(0, 0, &package),
                end_document,
                table_styles_record(0x0f, 0, &package),
            ]))
            .is_err()
        );
    }

    #[test]
    fn presentation_exposes_real_custom_table_style_version_variants() {
        for (name, expected_version) in [("SampleShow.ppt", 0), ("text-margins.ppt", 0x0f)] {
            let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ole/ppt")
                .join(name);
            let mut package = Package::open(path).unwrap();
            let presentation = package.presentation().unwrap();
            let properties = presentation.powerpoint12_document_properties().unwrap();
            assert_eq!(
                properties.custom_table_styles.unwrap().record_version,
                expected_version
            );
        }
    }
}
