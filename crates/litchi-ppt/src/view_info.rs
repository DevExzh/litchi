//! Typed slide and notes view-information records.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

const MAX_CONTAINER_DATA: usize = 327;
const SLIDE_VIEW_INFO_TYPE: u16 = 1018;
const GUIDE_ATOM_TYPE: u16 = 1019;
const VIEW_INFO_ATOM_TYPE: u16 = 1021;
const SLIDE_VIEW_INFO_ATOM_TYPE: u16 = 1022;

fn corrupted(message: impl Into<String>) -> PptError {
    PptError::Corrupted(message.into())
}

fn read_i32(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap())
}

fn strict_bool(value: u8, field: &str) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(corrupted(format!("{field} is not a bool1"))),
    }
}

/// Whether a view-information container applies to slides or notes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointViewKind {
    Slide,
    Notes,
}

impl PowerPointViewKind {
    fn from_instance(instance: u16) -> Result<Self> {
        match instance {
            0 => Ok(Self::Slide),
            1 => Ok(Self::Notes),
            _ => Err(corrupted(
                "SlideViewInfo record instance must be zero or one",
            )),
        }
    }

    fn instance(self) -> u16 {
        match self {
            Self::Slide => 0,
            Self::Notes => 1,
        }
    }
}

/// Signed rational number used by PowerPoint view scaling.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointRatio {
    numerator: i32,
    denominator: i32,
}

impl PowerPointRatio {
    pub fn new(numerator: i32, denominator: i32) -> Result<Self> {
        if denominator == 0 {
            return Err(corrupted("PowerPoint ratio denominator must not be zero"));
        }
        Ok(Self {
            numerator,
            denominator,
        })
    }

    pub fn numerator(&self) -> i32 {
        self.numerator
    }
    pub fn denominator(&self) -> i32 {
        self.denominator
    }

    fn normalized(self) -> (i64, i64) {
        let numerator = i64::from(self.numerator);
        let denominator = i64::from(self.denominator);
        if denominator < 0 {
            (-numerator, -denominator)
        } else {
            (numerator, denominator)
        }
    }
}

/// Origin in master units relative to the full view's top-left corner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointViewOrigin {
    x: i32,
    y: i32,
}

impl PowerPointViewOrigin {
    pub const fn new(x: i32, y: i32) -> Self {
        Self { x, y }
    }
    pub fn x(&self) -> i32 {
        self.x
    }
    pub fn y(&self) -> i32 {
        self.y
    }
}

/// `ZoomViewInfoAtom`, including ignored bytes retained for exact identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointZoomViewInfo {
    x_scale: PowerPointRatio,
    y_scale: PowerPointRatio,
    ignored1: [u8; 24],
    origin: PowerPointViewOrigin,
    use_variable_scale: bool,
    draft_mode: bool,
    ignored2: [u8; 2],
}

impl PowerPointZoomViewInfo {
    pub fn new(
        x_scale: PowerPointRatio,
        y_scale: PowerPointRatio,
        origin: PowerPointViewOrigin,
        use_variable_scale: bool,
        draft_mode: bool,
    ) -> Result<Self> {
        validate_zoom_scales(x_scale, y_scale)?;
        Ok(Self {
            x_scale,
            y_scale,
            ignored1: [0; 24],
            origin,
            use_variable_scale,
            draft_mode,
            ignored2: [0; 2],
        })
    }

    pub fn x_scale(&self) -> PowerPointRatio {
        self.x_scale
    }
    pub fn y_scale(&self) -> PowerPointRatio {
        self.y_scale
    }
    pub fn origin(&self) -> PowerPointViewOrigin {
        self.origin
    }
    pub fn uses_variable_scale(&self) -> bool {
        self.use_variable_scale
    }
    pub fn is_draft_mode(&self) -> bool {
        self.draft_mode
    }
    pub fn ignored_bytes(&self) -> (&[u8; 24], &[u8; 2]) {
        (&self.ignored1, &self.ignored2)
    }

    fn parse(record: &PptRecord) -> Result<Self> {
        validate_atom(
            record,
            PptRecordType::ViewInfoAtom,
            VIEW_INFO_ATOM_TYPE,
            52,
            0,
        )?;
        let data = &record.data;
        let x_scale = PowerPointRatio::new(read_i32(data, 0), read_i32(data, 4))?;
        let y_scale = PowerPointRatio::new(read_i32(data, 8), read_i32(data, 12))?;
        validate_zoom_scales(x_scale, y_scale)?;
        Ok(Self {
            x_scale,
            y_scale,
            ignored1: data[16..40].try_into().unwrap(),
            origin: PowerPointViewOrigin::new(read_i32(data, 40), read_i32(data, 44)),
            use_variable_scale: strict_bool(data[48], "ZoomViewInfoAtom.fUseVarScale")?,
            draft_mode: strict_bool(data[49], "ZoomViewInfoAtom.fDraftMode")?,
            ignored2: data[50..52].try_into().unwrap(),
        })
    }

    fn to_bytes(&self) -> Result<Vec<u8>> {
        validate_zoom_scales(self.x_scale, self.y_scale)?;
        let mut data = Vec::with_capacity(52);
        for ratio in [self.x_scale, self.y_scale] {
            data.extend_from_slice(&ratio.numerator.to_le_bytes());
            data.extend_from_slice(&ratio.denominator.to_le_bytes());
        }
        data.extend_from_slice(&self.ignored1);
        data.extend_from_slice(&self.origin.x.to_le_bytes());
        data.extend_from_slice(&self.origin.y.to_le_bytes());
        data.push(u8::from(self.use_variable_scale));
        data.push(u8::from(self.draft_mode));
        data.extend_from_slice(&self.ignored2);
        record_bytes(0, 0, VIEW_INFO_ATOM_TYPE, &data)
    }
}

fn validate_zoom_scales(x: PowerPointRatio, y: PowerPointRatio) -> Result<()> {
    for ratio in [x, y] {
        let (numerator, denominator) = ratio.normalized();
        if numerator <= 0 || numerator * 10 < denominator || numerator > denominator * 4 {
            return Err(corrupted("ZoomViewInfo scale must be between 0.10 and 4.0"));
        }
    }
    if i64::from(x.numerator) * i64::from(y.denominator)
        != i64::from(y.numerator) * i64::from(x.denominator)
    {
        return Err(corrupted("ZoomViewInfo x and y scales must be equal"));
    }
    Ok(())
}

/// Editing preferences from `SlideViewInfoAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointSlideViewPreferences {
    ignored: u8,
    snap_to_grid: bool,
    snap_to_shape: bool,
}

impl PowerPointSlideViewPreferences {
    pub const fn new(snap_to_grid: bool, snap_to_shape: bool) -> Self {
        Self {
            ignored: 0,
            snap_to_grid,
            snap_to_shape,
        }
    }
    pub fn snap_to_grid(&self) -> bool {
        self.snap_to_grid
    }
    pub fn snap_to_shape(&self) -> bool {
        self.snap_to_shape
    }
    pub fn ignored_byte(&self) -> u8 {
        self.ignored
    }

    fn parse(record: &PptRecord) -> Result<Self> {
        validate_atom(
            record,
            PptRecordType::SlideViewInfoAtom,
            SLIDE_VIEW_INFO_ATOM_TYPE,
            3,
            0,
        )?;
        Ok(Self {
            ignored: record.data[0],
            snap_to_grid: strict_bool(record.data[1], "SlideViewInfoAtom.fSnapToGrid")?,
            snap_to_shape: strict_bool(record.data[2], "SlideViewInfoAtom.fSnapToShape")?,
        })
    }

    fn to_bytes(self) -> Result<Vec<u8>> {
        record_bytes(
            0,
            0,
            SLIDE_VIEW_INFO_ATOM_TYPE,
            &[
                self.ignored,
                u8::from(self.snap_to_grid),
                u8::from(self.snap_to_shape),
            ],
        )
    }
}

/// Orientation stored in a `GuideAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PowerPointGuideOrientation {
    Horizontal,
    Vertical,
}

/// One alignment guide in master units.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointGuide {
    orientation: PowerPointGuideOrientation,
    position: i32,
}

impl PowerPointGuide {
    pub fn new(orientation: PowerPointGuideOrientation, position: i32) -> Result<Self> {
        if !(-15_840..=32_255).contains(&position) {
            return Err(corrupted("GuideAtom position is outside -15840..=32255"));
        }
        Ok(Self {
            orientation,
            position,
        })
    }
    pub fn orientation(&self) -> PowerPointGuideOrientation {
        self.orientation
    }
    pub fn position(&self) -> i32 {
        self.position
    }

    fn parse(record: &PptRecord) -> Result<Self> {
        validate_atom(record, PptRecordType::GuideAtom, GUIDE_ATOM_TYPE, 8, 7)?;
        let orientation = match read_u32(&record.data, 0) {
            0 => PowerPointGuideOrientation::Horizontal,
            1 => PowerPointGuideOrientation::Vertical,
            _ => return Err(corrupted("GuideAtom type must be horizontal or vertical")),
        };
        Self::new(orientation, read_i32(&record.data, 4))
    }

    fn to_bytes(self) -> Result<Vec<u8>> {
        let orientation = match self.orientation {
            PowerPointGuideOrientation::Horizontal => 0u32,
            PowerPointGuideOrientation::Vertical => 1u32,
        };
        let mut data = Vec::with_capacity(8);
        data.extend_from_slice(&orientation.to_le_bytes());
        data.extend_from_slice(&self.position.to_le_bytes());
        record_bytes(0, 7, GUIDE_ATOM_TYPE, &data)
    }
}

/// Complete slide or notes `SlideViewInfoContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointSlideViewInfo {
    kind: PowerPointViewKind,
    preferences: PowerPointSlideViewPreferences,
    zoom: Option<PowerPointZoomViewInfo>,
    guides: Vec<PowerPointGuide>,
}

impl PowerPointSlideViewInfo {
    pub fn new(
        kind: PowerPointViewKind,
        preferences: PowerPointSlideViewPreferences,
        zoom: Option<PowerPointZoomViewInfo>,
        guides: Vec<PowerPointGuide>,
    ) -> Result<Self> {
        validate_guides(&guides)?;
        Ok(Self {
            kind,
            preferences,
            zoom,
            guides,
        })
    }
    pub fn kind(&self) -> PowerPointViewKind {
        self.kind
    }
    pub fn preferences(&self) -> PowerPointSlideViewPreferences {
        self.preferences
    }
    pub fn zoom(&self) -> Option<&PowerPointZoomViewInfo> {
        self.zoom.as_ref()
    }
    pub fn guides(&self) -> &[PowerPointGuide] {
        &self.guides
    }

    pub fn parse_record(record: &PptRecord) -> Result<Self> {
        let declared = usize::try_from(record.data_length)
            .map_err(|_| corrupted("SlideViewInfo length does not fit memory"))?;
        if record.record_type != PptRecordType::SlideViewInfo
            || record.record_type_raw != SLIDE_VIEW_INFO_TYPE
            || record.version != 0xF
            || declared != record.data.len()
            || record.data.len() > MAX_CONTAINER_DATA
        {
            return Err(corrupted(
                "SlideViewInfo container has an invalid header or size",
            ));
        }
        let kind = PowerPointViewKind::from_instance(record.instance)?;
        let children = PptRecord::parse_sequence_strict(&record.data, "SlideViewInfo")?;
        let Some(first) = children.first() else {
            return Err(corrupted("SlideViewInfo is missing SlideViewInfoAtom"));
        };
        let preferences = PowerPointSlideViewPreferences::parse(first)?;
        let mut index = 1usize;
        let zoom = if children
            .get(index)
            .is_some_and(|child| child.record_type == PptRecordType::ViewInfoAtom)
        {
            let zoom = PowerPointZoomViewInfo::parse(&children[index])?;
            index += 1;
            Some(zoom)
        } else {
            None
        };
        let mut guides = Vec::with_capacity(children.len().saturating_sub(index));
        for child in &children[index..] {
            if child.record_type != PptRecordType::GuideAtom {
                return Err(corrupted(
                    "SlideViewInfo contains an unexpected or out-of-order child",
                ));
            }
            guides.push(PowerPointGuide::parse(child)?);
        }
        validate_guides(&guides)?;
        Ok(Self {
            kind,
            preferences,
            zoom,
            guides,
        })
    }

    /// Serialize a complete container, preserving every ignored atom byte.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        validate_guides(&self.guides)?;
        let mut data = self.preferences.to_bytes()?;
        if let Some(zoom) = &self.zoom {
            data.extend(zoom.to_bytes()?);
        }
        for guide in &self.guides {
            data.extend(guide.to_bytes()?);
        }
        if data.len() > MAX_CONTAINER_DATA {
            return Err(corrupted(
                "SlideViewInfo exceeds its bounded container size",
            ));
        }
        record_bytes(0xF, self.kind.instance(), SLIDE_VIEW_INFO_TYPE, &data)
    }
}

fn validate_guides(guides: &[PowerPointGuide]) -> Result<()> {
    let horizontal = guides
        .iter()
        .filter(|guide| guide.orientation == PowerPointGuideOrientation::Horizontal)
        .count();
    let vertical = guides.len() - horizontal;
    if horizontal > 8 || vertical > 8 {
        return Err(corrupted(
            "SlideViewInfo contains more than eight guides of one orientation",
        ));
    }
    Ok(())
}

fn validate_atom(
    record: &PptRecord,
    record_type: PptRecordType,
    raw_type: u16,
    length: usize,
    instance: u16,
) -> Result<()> {
    let declared = usize::try_from(record.data_length)
        .map_err(|_| corrupted(format!("PPT atom {raw_type} length does not fit memory")))?;
    if record.record_type != record_type
        || record.record_type_raw != raw_type
        || record.version != 0
        || record.instance != instance
        || declared != length
        || record.data.len() != length
    {
        return Err(corrupted(format!(
            "PPT atom {raw_type} has an invalid header or length"
        )));
    }
    Ok(())
}

fn record_bytes(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Result<Vec<u8>> {
    if version > 0xF || instance > 0x0FFF {
        return Err(corrupted("PPT record version or instance is out of range"));
    }
    let length = u32::try_from(data.len())
        .map_err(|_| corrupted("PPT view record payload exceeds 32-bit length"))?;
    let mut bytes = Vec::with_capacity(8 + data.len());
    bytes.extend_from_slice(&(version | instance << 4).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

/// Slide and notes view information exposed by a presentation.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointSlideViewInformation {
    slide: Option<PowerPointSlideViewInfo>,
    notes: Option<PowerPointSlideViewInfo>,
}

impl PowerPointSlideViewInformation {
    pub fn slide(&self) -> Option<&PowerPointSlideViewInfo> {
        self.slide.as_ref()
    }
    pub fn notes(&self) -> Option<&PowerPointSlideViewInfo> {
        self.notes.as_ref()
    }

    pub(crate) fn parse_records(records: &[&PptRecord]) -> Result<Self> {
        let mut information = Self::default();
        for record in records {
            if record.record_type != PptRecordType::SlideViewInfo {
                continue;
            }
            let view = PowerPointSlideViewInfo::parse_record(record)?;
            let slot = match view.kind {
                PowerPointViewKind::Slide => &mut information.slide,
                PowerPointViewKind::Notes => &mut information.notes,
            };
            if slot.replace(view).is_some() {
                return Err(corrupted(
                    "presentation contains duplicate SlideViewInfo instances",
                ));
            }
        }
        Ok(information)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn decode_hex(hex: &str) -> Vec<u8> {
        hex.as_bytes()
            .chunks_exact(2)
            .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
            .collect()
    }

    fn poi_reference() -> Vec<u8> {
        decode_hex(concat!(
            "0f00fa0367000000",
            "0000fe0303000000000100",
            "0000fd0334000000",
            "56000000640000005600000064000000",
            "d44eca6f0000000010d7e56ed5d6e56ed0a7340008000000",
            "faf9ffffa0ffffff01003400",
            "7000fb03080000000000000070080000",
            "7000fb030800000001000000400b0000"
        ))
    }

    #[test]
    fn parses_and_serializes_poi_view_bytes_exactly() {
        let bytes = poi_reference();
        let (record, consumed) = PptRecord::parse_strict(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        let view = PowerPointSlideViewInfo::parse_record(&record).unwrap();
        assert_eq!(view.kind(), PowerPointViewKind::Slide);
        assert!(view.preferences().snap_to_grid());
        assert!(!view.preferences().snap_to_shape());
        let zoom = view.zoom().unwrap();
        assert_eq!(
            (zoom.x_scale().numerator(), zoom.x_scale().denominator()),
            (86, 100)
        );
        assert_eq!((zoom.origin().x(), zoom.origin().y()), (-1542, -96));
        assert!(zoom.uses_variable_scale());
        assert_eq!(view.guides()[0].position(), 2160);
        assert_eq!(
            view.guides()[1].orientation(),
            PowerPointGuideOrientation::Vertical
        );
        assert_eq!(view.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn rejects_malformed_atoms_scales_guides_and_caps() {
        let mut invalid_bool = poi_reference();
        invalid_bool[17] = 2;
        let (record, _) = PptRecord::parse_strict(&invalid_bool, 0).unwrap();
        assert!(PowerPointSlideViewInfo::parse_record(&record).is_err());

        let mut zero_denominator = poi_reference();
        zero_denominator[31..35].copy_from_slice(&0i32.to_le_bytes());
        let (record, _) = PptRecord::parse_strict(&zero_denominator, 0).unwrap();
        assert!(PowerPointSlideViewInfo::parse_record(&record).is_err());

        let mut mismatched_scale = poi_reference();
        mismatched_scale[35..39].copy_from_slice(&85i32.to_le_bytes());
        let (record, _) = PptRecord::parse_strict(&mismatched_scale, 0).unwrap();
        assert!(PowerPointSlideViewInfo::parse_record(&record).is_err());

        assert!(PowerPointGuide::new(PowerPointGuideOrientation::Horizontal, 32_256).is_err());
        let nine =
            vec![PowerPointGuide::new(PowerPointGuideOrientation::Horizontal, 0).unwrap(); 9];
        assert!(
            PowerPointSlideViewInfo::new(
                PowerPointViewKind::Slide,
                PowerPointSlideViewPreferences::new(false, false),
                None,
                nine,
            )
            .is_err()
        );

        let mut truncated = poi_reference();
        truncated.pop();
        assert!(PptRecord::parse_strict(&truncated, 0).is_err());
    }
}
