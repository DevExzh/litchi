//! Typed legacy PowerPoint color-scheme atoms (MS-PPT 2.5.14, 2.5.15).
//!
//! Parsing is limited to bytes already present in caller-supplied PPT
//! records. The colors are inert display metadata: nothing is rendered,
//! resolved against a theme, or applied to shapes.

use super::package::{PptError, Result};
use super::records::PptRecord;

/// `RT_ColorSchemeAtom` record type (MS-PPT 2.13.24).
const COLOR_SCHEME_ATOM_TYPE: u16 = 0x07F0;
/// Number of `ColorStruct` items in a color scheme (MS-PPT 2.5.14).
const COLOR_SCHEME_COLOR_COUNT: usize = 8;
/// Byte length of one `ColorStruct` (MS-PPT 2.12.1).
const COLOR_STRUCT_LEN: usize = 4;
/// Byte length of a color-scheme atom payload.
const COLOR_SCHEME_PAYLOAD_LEN: usize = COLOR_SCHEME_COLOR_COUNT * COLOR_STRUCT_LEN;
/// Record instance of a `SlideSchemeColorSchemeAtom`.
const SLIDE_SCHEME_INSTANCE: u16 = 0x001;
/// Record instance of a `SchemeListElementColorSchemeAtom`.
const SCHEME_LIST_ELEMENT_INSTANCE: u16 = 0x006;

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

/// One sRGB color from a `ColorStruct` (MS-PPT 2.12.1).
///
/// The fourth byte of a `ColorStruct` is undefined by the specification and
/// is ignored while parsing and zeroed while serializing.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct PowerPointSchemeColor {
    pub red: u8,
    pub green: u8,
    pub blue: u8,
}

impl PowerPointSchemeColor {
    pub const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self { red, green, blue }
    }
}

/// Which role an `RT_ColorSchemeAtom` record plays in its container.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PowerPointColorSchemeAtomKind {
    /// `SlideSchemeColorSchemeAtom`: the color scheme used by one slide
    /// (MS-PPT 2.5.14).
    SlideScheme,
    /// `SchemeListElementColorSchemeAtom`: one entry in the main master's
    /// list of available color schemes (MS-PPT 2.5.15).
    SchemeListElement,
}

impl PowerPointColorSchemeAtomKind {
    fn from_instance(instance: u16) -> Result<Self> {
        match instance {
            SLIDE_SCHEME_INSTANCE => Ok(Self::SlideScheme),
            SCHEME_LIST_ELEMENT_INSTANCE => Ok(Self::SchemeListElement),
            _ => corrupted("ColorSchemeAtom has an invalid record instance"),
        }
    }

    fn instance(self) -> u16 {
        match self {
            Self::SlideScheme => SLIDE_SCHEME_INSTANCE,
            Self::SchemeListElement => SCHEME_LIST_ELEMENT_INSTANCE,
        }
    }
}

/// The eight colors of a PowerPoint color scheme, in `rgSchemeColor` order
/// (MS-PPT 2.5.14).
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct PowerPointColorScheme {
    pub background: PowerPointSchemeColor,
    pub text_and_lines: PowerPointSchemeColor,
    pub shadows: PowerPointSchemeColor,
    pub title_text: PowerPointSchemeColor,
    pub fills: PowerPointSchemeColor,
    pub accent: PowerPointSchemeColor,
    pub accent_and_hyperlink: PowerPointSchemeColor,
    pub accent_and_followed_hyperlink: PowerPointSchemeColor,
}

impl PowerPointColorScheme {
    fn as_array(&self) -> [PowerPointSchemeColor; COLOR_SCHEME_COLOR_COUNT] {
        [
            self.background,
            self.text_and_lines,
            self.shadows,
            self.title_text,
            self.fills,
            self.accent,
            self.accent_and_hyperlink,
            self.accent_and_followed_hyperlink,
        ]
    }

    fn from_array(colors: [PowerPointSchemeColor; COLOR_SCHEME_COLOR_COUNT]) -> Self {
        Self {
            background: colors[0],
            text_and_lines: colors[1],
            shadows: colors[2],
            title_text: colors[3],
            fills: colors[4],
            accent: colors[5],
            accent_and_hyperlink: colors[6],
            accent_and_followed_hyperlink: colors[7],
        }
    }
}

/// A validated `SlideSchemeColorSchemeAtom` or
/// `SchemeListElementColorSchemeAtom` record.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PowerPointColorSchemeAtom {
    pub kind: PowerPointColorSchemeAtomKind,
    pub scheme: PowerPointColorScheme,
}

impl PowerPointColorSchemeAtom {
    pub fn new_slide_scheme(scheme: PowerPointColorScheme) -> Self {
        Self {
            kind: PowerPointColorSchemeAtomKind::SlideScheme,
            scheme,
        }
    }

    pub fn new_scheme_list_element(scheme: PowerPointColorScheme) -> Self {
        Self {
            kind: PowerPointColorSchemeAtomKind::SchemeListElement,
            scheme,
        }
    }

    /// Strictly parse one already-materialized `RT_ColorSchemeAtom` record.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0
            || record.record_type_raw != COLOR_SCHEME_ATOM_TYPE
            || record.data.len() != COLOR_SCHEME_PAYLOAD_LEN
            || record.data_length != COLOR_SCHEME_PAYLOAD_LEN as u32
        {
            return corrupted("ColorSchemeAtom has an invalid record header or size");
        }
        let kind = PowerPointColorSchemeAtomKind::from_instance(record.instance)?;
        let mut colors = [PowerPointSchemeColor::default(); COLOR_SCHEME_COLOR_COUNT];
        for (index, color) in colors.iter_mut().enumerate() {
            let start = index * COLOR_STRUCT_LEN;
            *color = PowerPointSchemeColor::new(
                record.data[start],
                record.data[start + 1],
                record.data[start + 2],
            );
        }
        Ok(Self {
            kind,
            scheme: PowerPointColorScheme::from_array(colors),
        })
    }

    /// Collect the color-scheme atoms held directly by a slide, notes, main
    /// master, or handout container.
    ///
    /// A container holds at most one `SlideSchemeColorSchemeAtom`; the main
    /// master may additionally hold any number of scheme list elements.
    pub fn collect(container: &PptRecord) -> Result<Vec<Self>> {
        let mut atoms = Vec::new();
        for child in &container.children {
            if child.record_type_raw != COLOR_SCHEME_ATOM_TYPE {
                continue;
            }
            let atom = Self::parse(child)?;
            if atom.kind == PowerPointColorSchemeAtomKind::SlideScheme
                && atoms.iter().any(|existing: &Self| {
                    existing.kind == PowerPointColorSchemeAtomKind::SlideScheme
                })
            {
                return corrupted("container contains more than one SlideSchemeColorSchemeAtom");
            }
            atoms.push(atom);
        }
        Ok(atoms)
    }

    /// Serialize this atom as one canonical `RT_ColorSchemeAtom` record.
    pub fn to_record(&self) -> Result<PptRecord> {
        let bytes = self.to_record_bytes();
        let (record, end) = PptRecord::parse(&bytes, 0)?;
        if end != bytes.len() {
            return corrupted("canonical ColorSchemeAtom did not consume its bytes");
        }
        Ok(record)
    }

    /// Serialize this atom as canonical `RT_ColorSchemeAtom` record bytes.
    pub fn to_record_bytes(&self) -> [u8; 8 + COLOR_SCHEME_PAYLOAD_LEN] {
        let mut bytes = [0u8; 8 + COLOR_SCHEME_PAYLOAD_LEN];
        bytes[0..2].copy_from_slice(&(self.kind.instance() << 4).to_le_bytes());
        bytes[2..4].copy_from_slice(&COLOR_SCHEME_ATOM_TYPE.to_le_bytes());
        bytes[4..8].copy_from_slice(&(COLOR_SCHEME_PAYLOAD_LEN as u32).to_le_bytes());
        for (index, color) in self.scheme.as_array().iter().enumerate() {
            let start = 8 + index * COLOR_STRUCT_LEN;
            bytes[start] = color.red;
            bytes[start + 1] = color.green;
            bytes[start + 2] = color.blue;
        }
        bytes
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::consts::PptRecordType;

    fn sample_scheme() -> PowerPointColorScheme {
        PowerPointColorScheme {
            background: PowerPointSchemeColor::new(255, 255, 255),
            text_and_lines: PowerPointSchemeColor::new(0, 0, 0),
            shadows: PowerPointSchemeColor::new(238, 236, 225),
            title_text: PowerPointSchemeColor::new(31, 73, 125),
            fills: PowerPointSchemeColor::new(79, 129, 189),
            accent: PowerPointSchemeColor::new(192, 80, 77),
            accent_and_hyperlink: PowerPointSchemeColor::new(0, 0, 255),
            accent_and_followed_hyperlink: PowerPointSchemeColor::new(128, 0, 128),
        }
    }

    #[test]
    fn slide_scheme_and_list_element_roundtrip() {
        let scheme = sample_scheme();
        for atom in [
            PowerPointColorSchemeAtom::new_slide_scheme(scheme),
            PowerPointColorSchemeAtom::new_scheme_list_element(scheme),
        ] {
            let record = atom.to_record().unwrap();
            assert_eq!(record.record_type_raw, COLOR_SCHEME_ATOM_TYPE);
            assert_eq!(record.version, 0);
            let parsed = PowerPointColorSchemeAtom::parse(&record).unwrap();
            assert_eq!(parsed, atom);
        }
    }

    #[test]
    fn rejects_invalid_headers_instances_and_sizes() {
        let atom = PowerPointColorSchemeAtom::new_slide_scheme(sample_scheme());
        let mut bad_instance = atom.to_record_bytes();
        bad_instance[1] = 0x20; // record instance 2
        let record = PptRecord::parse(&bad_instance, 0).unwrap().0;
        assert!(PowerPointColorSchemeAtom::parse(&record).is_err());

        let mut bad_version = atom.to_record_bytes();
        bad_version[0] = 0x0f;
        let record = PptRecord::parse(&bad_version, 0).unwrap().0;
        assert!(PowerPointColorSchemeAtom::parse(&record).is_err());

        let short = PptRecord {
            record_type: PptRecordType::Unknown,
            record_type_raw: COLOR_SCHEME_ATOM_TYPE,
            version: 0,
            instance: 1,
            data_length: 31,
            data: vec![0; 31],
            children: Vec::new(),
        };
        assert!(PowerPointColorSchemeAtom::parse(&short).is_err());
    }

    #[test]
    fn collect_rejects_duplicate_slide_schemes() {
        let scheme = sample_scheme();
        let slide = PowerPointColorSchemeAtom::new_slide_scheme(scheme)
            .to_record()
            .unwrap();
        let element = PowerPointColorSchemeAtom::new_scheme_list_element(scheme)
            .to_record()
            .unwrap();
        let container = |children: Vec<PptRecord>| PptRecord {
            record_type: PptRecordType::MainMaster,
            record_type_raw: PptRecordType::MainMaster.as_u16(),
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children,
        };

        let atoms =
            PowerPointColorSchemeAtom::collect(&container(vec![slide.clone(), element])).unwrap();
        assert_eq!(atoms.len(), 2);
        assert_eq!(atoms[0].kind, PowerPointColorSchemeAtomKind::SlideScheme);

        assert!(
            PowerPointColorSchemeAtom::collect(&container(vec![slide.clone(), slide])).is_err()
        );
        assert!(
            PowerPointColorSchemeAtom::collect(&container(Vec::new()))
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn presentation_exposes_real_color_schemes() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/ppt/SampleShow.ppt");
        let mut package = crate::Package::open(path).unwrap();
        let presentation = package.presentation().unwrap();
        let schemes = presentation.color_schemes().unwrap();

        assert!(
            schemes
                .iter()
                .any(|atom| atom.kind == PowerPointColorSchemeAtomKind::SlideScheme)
        );
        assert!(
            schemes
                .iter()
                .any(|atom| atom.kind == PowerPointColorSchemeAtomKind::SchemeListElement)
        );
        let scheme = schemes[0].scheme;
        assert_eq!(scheme.background, PowerPointSchemeColor::new(255, 255, 255));
        assert_eq!(
            scheme.accent_and_hyperlink,
            PowerPointSchemeColor::new(0, 0, 255)
        );
    }
}
