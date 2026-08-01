//! PowerPoint projections over format-neutral OfficeArt shapes.
//!
//! `[MS-ODRAW]` deliberately leaves `ClientData` and `ClientTextbox` payloads
//! to the host application.  Keeping their interpretation here prevents the
//! shared drawing crate from acquiring a PowerPoint dependency while still
//! giving PPT callers a concise, typed shape API.

use litchi_odraw::shape::Shape;
use litchi_odraw::{Container, Record, RecordKind};

use super::package::{PptError, Result};

const CLIENT_DATA_RAW_KIND: u16 = 0xf011;
const CLIENT_TEXTBOX_RAW_KIND: u16 = 0xf00d;
const MAX_HOST_RECORDS: u32 = 1_000_000;

/// Placeholder metadata embedded in a PowerPoint shape's client data.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Placeholder {
    /// Placeholder position from the PowerPoint `PlaceholderAtom`.
    pub position: Option<u16>,
    /// Exact PowerPoint placeholder kind.
    pub kind: super::PowerPointPlaceholderKind,
    /// Checked PowerPoint placeholder size.
    pub size: super::PowerPointPlaceholderSize,
}

/// Host-specific meaning of an OfficeArt picture-frame shape.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum FrameKind {
    /// An ordinary picture frame.
    #[default]
    Picture,
    /// A frame associated with an embedded or linked OLE object.
    Object,
    /// A frame associated with audio or video media.
    Media,
}

/// Checked PowerPoint shape bounds projected from either anchor encoding.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Anchor {
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
    width: i32,
    height: i32,
}

impl Anchor {
    fn new(left: i32, top: i32, right: i32, bottom: i32) -> Result<Self> {
        let width = right.checked_sub(left).ok_or_else(|| {
            corrupted("PowerPoint shape anchor width exceeds a signed coordinate")
        })?;
        let height = bottom.checked_sub(top).ok_or_else(|| {
            corrupted("PowerPoint shape anchor height exceeds a signed coordinate")
        })?;
        if width < 0 || height < 0 {
            return Err(corrupted("PowerPoint shape anchor has inverted bounds"));
        }
        Ok(Self {
            left,
            top,
            right,
            bottom,
            width,
            height,
        })
    }

    /// Minimum x-coordinate.
    pub const fn left(self) -> i32 {
        self.left
    }

    /// Minimum y-coordinate.
    pub const fn top(self) -> i32 {
        self.top
    }

    /// Maximum x-coordinate.
    pub const fn right(self) -> i32 {
        self.right
    }

    /// Maximum y-coordinate.
    pub const fn bottom(self) -> i32 {
        self.bottom
    }

    /// Width in PowerPoint master units.
    pub const fn width(self) -> i32 {
        self.width
    }

    /// Height in PowerPoint master units.
    pub const fn height(self) -> i32 {
        self.height
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct Frame {
    kind: FrameKind,
    object_id: Option<u32>,
}

/// PowerPoint-only behavior layered over a neutral OfficeArt shape.
///
/// Import this trait as `_` at PPT call sites.  The resulting method surface
/// stays short without exposing `ClientData` record IDs to API users.
pub trait ShapeExt {
    /// Decodes the shape's PPT textbox payload.
    fn text(&self) -> Result<Option<String>>;

    /// Projects legacy placeholder metadata into checked semantic values.
    fn placeholder(&self) -> Result<Option<Placeholder>>;

    /// Distinguishes pictures, OLE frames, and media frames.
    fn frame_kind(&self) -> Result<FrameKind>;

    /// Returns the PPT external-object reference for an OLE or media frame.
    fn external_object_id(&self) -> Result<Option<u32>>;

    /// Parses click and mouse-over actions with default limits.
    fn interactions(&self) -> Result<Vec<super::PowerPointInteraction>>;

    /// Parses click and mouse-over actions with explicit limits.
    fn interactions_with_limits(
        &self,
        limits: super::PowerPointInteractionLimits,
    ) -> Result<Vec<super::PowerPointInteraction>>;

    /// Parses range-anchored text actions with default limits.
    fn text_interactions(&self) -> Result<Vec<super::PowerPointTextInteraction>>;

    /// Parses range-anchored text actions with explicit limits.
    fn text_interactions_with_limits(
        &self,
        limits: super::PowerPointTextInteractionLimits,
    ) -> Result<Vec<super::PowerPointTextInteraction>>;

    /// Parses a context-validated placeholder atom with default limits.
    fn placeholder_atom(
        &self,
        context: super::PowerPointPlaceholderContext,
    ) -> Result<Option<super::PowerPointPlaceholderAtom>>;

    /// Parses a context-validated placeholder atom with explicit limits.
    fn placeholder_atom_with_limits(
        &self,
        context: super::PowerPointPlaceholderContext,
        limits: super::PowerPointPlaceholderLimits,
    ) -> Result<Option<super::PowerPointPlaceholderAtom>>;

    /// Parses PowerPoint 12 shape round-trip metadata.
    fn powerpoint12_shape_metadata(&self) -> Result<Option<super::PowerPoint12ShapeMetadata>>;

    /// Parses inert shape programmable tags with default limits.
    fn programmable_tags(&self) -> Result<Option<super::PowerPointShapeProgrammableTags>>;

    /// Parses inert shape programmable tags with explicit limits.
    fn programmable_tags_with_limits(
        &self,
        limits: super::PowerPointShapeProgrammableTagLimits,
    ) -> Result<Option<super::PowerPointShapeProgrammableTags>>;

    /// Parses the PPT shape-flag projection with default limits.
    fn ppt_flags(&self) -> Result<Option<super::PowerPointShapeFlagProjection>>;

    /// Parses the PPT shape-flag projection with explicit limits.
    fn ppt_flags_with(
        &self,
        limits: super::PowerPointShapeFlagLimits,
    ) -> Result<Option<super::PowerPointShapeFlagProjection>>;

    /// Parses inert legacy PowerPoint animation metadata.
    fn animation(&self) -> Result<Option<super::animation::AnimationInfo>>;
}

impl ShapeExt for Shape<'_> {
    fn text(&self) -> Result<Option<String>> {
        let Some(textbox) = host_record(self, RecordKind::ClientTextbox)? else {
            return Ok(None);
        };
        validate_host_record(&textbox, CLIENT_TEXTBOX_RAW_KIND, "ClientTextbox")?;
        text_from_textbox(&textbox)
    }

    fn placeholder(&self) -> Result<Option<Placeholder>> {
        placeholder(self)
    }

    fn frame_kind(&self) -> Result<FrameKind> {
        Ok(frame(self)?.kind)
    }

    fn external_object_id(&self) -> Result<Option<u32>> {
        Ok(frame(self)?.object_id)
    }

    fn interactions(&self) -> Result<Vec<super::PowerPointInteraction>> {
        self.interactions_with_limits(super::PowerPointInteractionLimits::default())
    }

    fn interactions_with_limits(
        &self,
        limits: super::PowerPointInteractionLimits,
    ) -> Result<Vec<super::PowerPointInteraction>> {
        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(Vec::new());
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
        super::PowerPointInteraction::parse_client_data_payload(client_data.data(), limits)
    }

    fn text_interactions(&self) -> Result<Vec<super::PowerPointTextInteraction>> {
        self.text_interactions_with_limits(super::PowerPointTextInteractionLimits::default())
    }

    fn text_interactions_with_limits(
        &self,
        limits: super::PowerPointTextInteractionLimits,
    ) -> Result<Vec<super::PowerPointTextInteraction>> {
        let Some(textbox) = host_record(self, RecordKind::ClientTextbox)? else {
            return Ok(Vec::new());
        };
        validate_host_record(&textbox, CLIENT_TEXTBOX_RAW_KIND, "ClientTextbox")?;
        super::EscherTextboxWrapper::parse_text_interactions_with_limits(textbox.data(), limits)
    }

    fn placeholder_atom(
        &self,
        context: super::PowerPointPlaceholderContext,
    ) -> Result<Option<super::PowerPointPlaceholderAtom>> {
        self.placeholder_atom_with_limits(context, super::PowerPointPlaceholderLimits::default())
    }

    fn placeholder_atom_with_limits(
        &self,
        context: super::PowerPointPlaceholderContext,
        limits: super::PowerPointPlaceholderLimits,
    ) -> Result<Option<super::PowerPointPlaceholderAtom>> {
        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(None);
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
        Ok(
            super::PowerPointPlaceholderProjection::parse_client_data_payload(
                client_data.data(),
                context,
                limits,
            )?
            .placeholder,
        )
    }

    fn powerpoint12_shape_metadata(&self) -> Result<Option<super::PowerPoint12ShapeMetadata>> {
        use super::{
            PowerPoint12ShapeMetadata, PowerPointHeaderFooterPlaceholder, PowerPointNewPlaceholder,
            PowerPointShapeChecksums,
        };
        use crate::consts::PptRecordType;

        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(None);
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;

        let mut metadata = PowerPoint12ShapeMetadata::default();
        let mut found = false;
        let mut offset = 0usize;
        let mut records = 0u32;
        while offset < client_data.data().len() {
            visit_host_record(&mut records)?;
            let (record, consumed) =
                super::records::PptRecord::parse_strict(client_data.data(), offset)?;
            match record.record_type {
                PptRecordType::RoundTripHFPlaceholder12Atom => {
                    if metadata.header_footer.is_some() {
                        return Err(corrupted(
                            "Shape contains duplicate RoundTripHFPlaceholder12Atom records",
                        ));
                    }
                    validate_round_trip_atom(&record, "RoundTripHFPlaceholder12Atom", 1)?;
                    metadata.header_footer = Some(match record.data[0] {
                        7 => PowerPointHeaderFooterPlaceholder::Date,
                        8 => PowerPointHeaderFooterPlaceholder::SlideNumber,
                        9 => PowerPointHeaderFooterPlaceholder::Footer,
                        10 => PowerPointHeaderFooterPlaceholder::Header,
                        _ => {
                            return Err(corrupted(
                                "RoundTripHFPlaceholder12Atom has an invalid placeholder ID",
                            ));
                        },
                    });
                    found = true;
                },
                PptRecordType::RoundTripNewPlaceholderId12Atom => {
                    if metadata.new_placeholder.is_some() {
                        return Err(corrupted(
                            "Shape contains duplicate RoundTripNewPlaceholderId12Atom records",
                        ));
                    }
                    validate_round_trip_atom(&record, "RoundTripNewPlaceholderId12Atom", 1)?;
                    metadata.new_placeholder = Some(match record.data[0] {
                        25 => PowerPointNewPlaceholder::VerticalObject,
                        26 => PowerPointNewPlaceholder::Picture,
                        _ => {
                            return Err(corrupted(
                                "RoundTripNewPlaceholderId12Atom has an invalid placeholder ID",
                            ));
                        },
                    });
                    found = true;
                },
                PptRecordType::RoundTripShapeId12Atom => {
                    if metadata.shape_id.is_some() {
                        return Err(corrupted(
                            "Shape contains duplicate RoundTripShapeId12Atom records",
                        ));
                    }
                    validate_round_trip_atom(&record, "RoundTripShapeId12Atom", 4)?;
                    metadata.shape_id = Some(u32::from_le_bytes([
                        record.data[0],
                        record.data[1],
                        record.data[2],
                        record.data[3],
                    ]));
                    found = true;
                },
                PptRecordType::RoundTripShapeCheckSumForCustomLayouts12Atom => {
                    if metadata.custom_layout_checksums.is_some() {
                        return Err(corrupted(
                            "Shape contains duplicate RoundTripShapeCheckSumForCustomLayouts12Atom records",
                        ));
                    }
                    validate_round_trip_atom(
                        &record,
                        "RoundTripShapeCheckSumForCustomLayouts12Atom",
                        8,
                    )?;
                    metadata.custom_layout_checksums = Some(PowerPointShapeChecksums {
                        shape: u32::from_le_bytes([
                            record.data[0],
                            record.data[1],
                            record.data[2],
                            record.data[3],
                        ]),
                        text: u32::from_le_bytes([
                            record.data[4],
                            record.data[5],
                            record.data[6],
                            record.data[7],
                        ]),
                    });
                    found = true;
                },
                _ => {},
            }
            offset = advance(offset, consumed, "OfficeArt client-data")?;
        }

        Ok(found.then_some(metadata))
    }

    fn programmable_tags_with_limits(
        &self,
        limits: super::PowerPointShapeProgrammableTagLimits,
    ) -> Result<Option<super::PowerPointShapeProgrammableTags>> {
        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(None);
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;

        let mut result = None;
        let mut offset = 0usize;
        let mut records = 0u32;
        while offset < client_data.data().len() {
            visit_host_record(&mut records)?;
            let (record, consumed) =
                super::records::PptRecord::parse_strict(client_data.data(), offset)?;
            if record.record_type == crate::consts::PptRecordType::ProgTags {
                let parsed = super::PowerPointShapeProgrammableTags::parse(&record, limits)?;
                if result.replace(parsed).is_some() {
                    return Err(corrupted(
                        "Shape ClientData contains multiple ShapeProgTagsContainer records",
                    ));
                }
            }
            offset = advance(offset, consumed, "OfficeArt client-data")?;
        }
        Ok(result)
    }

    fn programmable_tags(&self) -> Result<Option<super::PowerPointShapeProgrammableTags>> {
        self.programmable_tags_with_limits(super::PowerPointShapeProgrammableTagLimits::default())
    }

    fn ppt_flags(&self) -> Result<Option<super::PowerPointShapeFlagProjection>> {
        self.ppt_flags_with(super::PowerPointShapeFlagLimits::default())
    }

    fn ppt_flags_with(
        &self,
        limits: super::PowerPointShapeFlagLimits,
    ) -> Result<Option<super::PowerPointShapeFlagProjection>> {
        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(None);
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
        let projection = super::PowerPointShapeFlagProjection::parse_client_data_payload(
            client_data.data(),
            limits,
        )?;
        Ok(projection.has_flags().then_some(projection))
    }

    fn animation(&self) -> Result<Option<super::animation::AnimationInfo>> {
        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(None);
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
        let mut offset = 0usize;
        let mut records = 0u32;
        while offset < client_data.data().len() {
            visit_host_record(&mut records)?;
            let (record, consumed) =
                super::records::PptRecord::parse_strict(client_data.data(), offset)?;
            if record.record_type == crate::consts::PptRecordType::AnimationInfo {
                return super::animation::parse_animation_info(&record).map(Some);
            }
            offset = advance(offset, consumed, "OfficeArt client-data")?;
        }
        Ok(None)
    }
}

/// Parse every complete OfficeArt drawing in a PowerPoint host stream.
///
/// A PPT `PPDrawing` payload can concatenate drawing containers. Each root is
/// still parsed exactly: malformed records and trailing partial records are
/// returned as errors instead of being ignored.
pub fn parse(data: &[u8]) -> Result<Vec<Shape<'_>>> {
    let mut shapes = Vec::new();
    let mut offset = 0usize;
    let mut records = 0u32;
    while offset < data.len() {
        visit_host_record(&mut records)?;
        let (_, consumed) = Record::parse(data, offset)?;
        let end = advance(offset, consumed, "PowerPoint OfficeArt drawing stream")?;
        shapes.extend(litchi_odraw::shape::parse(&data[offset..end])?);
        offset = end;
    }
    Ok(shapes)
}

/// Project a shape's group-relative or PPT client anchor into checked bounds.
pub fn anchor(shape: &Shape<'_>) -> Result<Option<Anchor>> {
    if let Some(anchor) = shape.anchor() {
        return Anchor::new(anchor.left, anchor.top, anchor.right, anchor.bottom).map(Some);
    }

    let Some(anchor) = shape.client_anchor() else {
        return Ok(None);
    };
    if anchor.version() != 0
        || anchor.instance() != 0
        || anchor.raw_kind() != RecordKind::ClientAnchor.raw()
        || usize::try_from(anchor.len()).ok() != Some(anchor.data().len())
    {
        return Err(corrupted("Invalid PowerPoint OfficeArtClientAnchor header"));
    }
    let anchor = super::PowerPointClientAnchorData::parse(anchor.data())?;
    Anchor::new(anchor.left(), anchor.top(), anchor.right(), anchor.bottom()).map(Some)
}

/// Extract all PPT textbox text from an OfficeArt drawing.
pub fn text_from_drawing(data: &[u8]) -> Result<String> {
    let mut offset = 0usize;
    let mut records = 0u32;
    let mut text = String::with_capacity(1024);
    while offset < data.len() {
        visit_host_record(&mut records)?;
        let (record, consumed) = Record::parse(data, offset)?;
        let container = Container::try_new(record)?;
        text_from_container(&container, &mut text)?;
        offset = advance(offset, consumed, "PowerPoint OfficeArt drawing stream")?;
    }
    Ok(text)
}

/// Decode the PPT records contained by one OfficeArt ClientTextbox atom.
pub fn text_from_textbox(textbox: &Record<'_>) -> Result<Option<String>> {
    validate_host_record(textbox, CLIENT_TEXTBOX_RAW_KIND, "ClientTextbox")?;
    if textbox.data().is_empty() {
        return Ok(None);
    }
    let mut text = String::with_capacity(256);
    text_from_ppt_records(textbox.data(), &mut text)?;
    let trimmed = text.trim_end().len();
    text.truncate(trimmed);
    Ok((!text.is_empty()).then_some(text))
}

fn host_record<'data>(shape: &Shape<'data>, kind: RecordKind) -> Result<Option<Record<'data>>> {
    match kind {
        RecordKind::ClientData => Ok(shape.client_data().cloned()),
        RecordKind::ClientTextbox => Ok(shape.textbox().cloned()),
        _ => shape.container().find(kind).map_err(Into::into),
    }
}

/// Returns a shape's header-validated PowerPoint textbox record.
pub(crate) fn textbox<'data>(shape: &Shape<'data>) -> Result<Option<Record<'data>>> {
    let Some(textbox) = host_record(shape, RecordKind::ClientTextbox)? else {
        return Ok(None);
    };
    validate_host_record(&textbox, CLIENT_TEXTBOX_RAW_KIND, "ClientTextbox")?;
    Ok(Some(textbox))
}

fn validate_host_record(record: &Record<'_>, raw_kind: u16, name: &str) -> Result<()> {
    if record.version() != 0x0f
        || record.instance() != 0
        || record.raw_kind() != raw_kind
        || usize::try_from(record.len()).ok() != Some(record.data().len())
    {
        return Err(corrupted(&format!(
            "Invalid OfficeArt {name} record header"
        )));
    }
    Ok(())
}

fn placeholder(shape: &Shape<'_>) -> Result<Option<Placeholder>> {
    const PLACEHOLDER_ATOM: u16 = 3011;

    let Some(client_data) = host_record(shape, RecordKind::ClientData)? else {
        return Ok(None);
    };
    validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
    let mut records = RawPptRecords::new(client_data.data());
    let mut found = None;
    for record in &mut records {
        let record = record?;
        if record.kind == PLACEHOLDER_ATOM {
            if record.data.len() < 8 {
                return Err(corrupted("PlaceholderAtom is shorter than eight bytes"));
            }
            if found.is_some() {
                return Err(corrupted(
                    "Shape ClientData contains multiple PlaceholderAtom records",
                ));
            }
            let position: [u8; 4] = record.data[..4]
                .try_into()
                .map_err(|_| corrupted("PlaceholderAtom position is not four bytes"))?;
            let position = i32::from_le_bytes(position);
            let position = match position {
                -1 => None,
                position => Some(u16::try_from(position).map_err(|_| {
                    corrupted("PlaceholderAtom position is outside the supported range")
                })?),
            };
            let kind = super::PowerPointPlaceholderKind::try_from(record.data[4])?;
            let size = super::PowerPointPlaceholderSize::try_from(record.data[5])?;
            found = Some(Placeholder {
                position,
                kind,
                size,
            });
        }
    }
    Ok(found)
}

fn frame(shape: &Shape<'_>) -> Result<Frame> {
    const EX_OBJ_REF_ATOM: u16 = 3009;
    const INTERACTIVE_INFO: u16 = 4082;
    const INTERACTIVE_INFO_ATOM: u16 = 4083;
    const ACTION_OLE: u8 = 5;
    const ACTION_MEDIA: u8 = 6;

    let Some(client_data) = host_record(shape, RecordKind::ClientData)? else {
        return Ok(Frame::default());
    };
    validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
    let mut frame = Frame::default();
    let mut action_kind = None;
    let mut host_records = 0;
    let mut records = RawPptRecords::new(client_data.data());
    for record in &mut records {
        let record = record?;
        visit_host_record(&mut host_records)?;
        match record.kind {
            EX_OBJ_REF_ATOM => {
                if record.data.len() < 4 {
                    return Err(corrupted("ExObjRefAtom is shorter than four bytes"));
                }
                if frame.object_id.is_some() {
                    return Err(corrupted(
                        "Shape ClientData contains multiple ExObjRefAtom records",
                    ));
                }
                let object_id: [u8; 4] = record.data[..4]
                    .try_into()
                    .map_err(|_| corrupted("ExObjRefAtom object ID is not four bytes"))?;
                frame.object_id = Some(u32::from_le_bytes(object_id));
                if frame.kind == FrameKind::Picture {
                    frame.kind = FrameKind::Object;
                }
            },
            INTERACTIVE_INFO => {
                let mut children = RawPptRecords::new(record.data);
                for child in &mut children {
                    let child = child?;
                    visit_host_record(&mut host_records)?;
                    if child.kind != INTERACTIVE_INFO_ATOM {
                        continue;
                    }
                    let action = child.data.get(8).ok_or_else(|| {
                        corrupted("InteractiveInfoAtom is shorter than nine bytes")
                    })?;
                    let kind = match *action {
                        ACTION_OLE => Some(FrameKind::Object),
                        ACTION_MEDIA => Some(FrameKind::Media),
                        _ => None,
                    };
                    if let Some(kind) = kind
                        && action_kind.replace(kind).is_some()
                    {
                        return Err(corrupted(
                            "Shape ClientData contains duplicate frame actions",
                        ));
                    }
                }
            },
            _ => {},
        }
    }
    if let Some(kind) = action_kind {
        frame.kind = kind;
    }
    Ok(frame)
}

fn text_from_container<'data>(container: &Container<'data>, text: &mut String) -> Result<()> {
    const MAX_RECORDS: usize = 1_000_000;
    const MAX_DEPTH: usize = 256;

    let mut pending = Vec::new();
    push_children(container, 1, &mut pending, MAX_RECORDS)?;
    let mut visited = 0usize;

    while let Some((record, depth)) = pending.pop() {
        visited = visited
            .checked_add(1)
            .ok_or_else(|| corrupted("OfficeArt traversal count overflow"))?;
        if visited > MAX_RECORDS {
            return Err(corrupted("OfficeArt drawing exceeds the PPT record limit"));
        }

        if record.kind() == RecordKind::ClientTextbox {
            let before = text.len();
            text_from_ppt_records(record.data(), text)?;
            let trimmed = text.trim_end().len();
            text.truncate(trimmed);
            if text.len() > before && !text.ends_with('\n') {
                text.push('\n');
            }
            continue;
        }

        if record.is_container() {
            if depth >= MAX_DEPTH {
                return Err(corrupted("OfficeArt drawing exceeds the PPT nesting limit"));
            }
            let nested = Container::try_new(record)?;
            push_children(&nested, depth + 1, &mut pending, MAX_RECORDS)?;
        }
    }
    Ok(())
}

fn push_children<'data>(
    container: &Container<'data>,
    depth: usize,
    pending: &mut Vec<(Record<'data>, usize)>,
    max_records: usize,
) -> Result<()> {
    let mut children = Vec::new();
    for child in container.children() {
        if pending.len().saturating_add(children.len()) >= max_records {
            return Err(corrupted("OfficeArt drawing exceeds the PPT record limit"));
        }
        children.push(child?);
    }
    pending.extend(children.into_iter().rev().map(|record| (record, depth)));
    Ok(())
}

fn text_from_ppt_records(data: &[u8], text: &mut String) -> Result<()> {
    const MAX_RECORDS: usize = 1_000_000;
    const MAX_DEPTH: usize = 256;

    let mut pending = vec![RawPptRecords::new(data)];
    let mut visited = 0usize;
    while let Some(records) = pending.last_mut() {
        let Some(record) = records.next() else {
            pending.pop();
            continue;
        };
        let record = record?;
        visited = visited
            .checked_add(1)
            .ok_or_else(|| corrupted("ClientTextbox traversal count overflow"))?;
        if visited > MAX_RECORDS {
            return Err(corrupted("ClientTextbox exceeds the PPT record limit"));
        }

        let decoded = match record.kind {
            4000 | 4026 => Some(super::text::extractor::from_utf16le_lossy(record.data)),
            4008 => Some(super::text::extractor::decode_text_bytes(record.data)),
            kind if is_text_container(kind) => {
                if pending.len() >= MAX_DEPTH {
                    return Err(corrupted("ClientTextbox exceeds the PPT nesting limit"));
                }
                pending.push(RawPptRecords::new(record.data));
                None
            },
            _ => None,
        };
        if let Some(decoded) = decoded {
            let decoded = decoded.trim();
            if decoded.is_empty() {
                continue;
            }
            if !text.is_empty() && !text.ends_with('\n') {
                text.push('\n');
            }
            text.push_str(decoded);
        }
    }
    Ok(())
}

fn is_text_container(kind: u16) -> bool {
    matches!(
        kind,
        1000 | 1006 | 1007 | 1010 | 1016 | 2000 | 3008 | 3009 | 4080 | 4085
    )
}

fn validate_round_trip_atom(
    record: &super::records::PptRecord,
    name: &str,
    expected_len: usize,
) -> Result<()> {
    if record.version != 0
        || record.instance != 0
        || record.data_length as usize != expected_len
        || record.data.len() != expected_len
    {
        return Err(corrupted(&format!(
            "{name} has an invalid record header or size"
        )));
    }
    Ok(())
}

fn advance(offset: usize, consumed: usize, context: &str) -> Result<usize> {
    if consumed == 0 {
        return Err(corrupted(&format!("Zero-length PPT record in {context}")));
    }
    offset
        .checked_add(consumed)
        .ok_or_else(|| corrupted(&format!("{context} offset overflow")))
}

fn visit_host_record(records: &mut u32) -> Result<()> {
    *records = records
        .checked_add(1)
        .ok_or_else(|| corrupted("Host payload record count overflow"))?;
    if *records > MAX_HOST_RECORDS {
        return Err(corrupted("Host payload exceeds the PPT record limit"));
    }
    Ok(())
}

fn corrupted(message: &str) -> PptError {
    PptError::Corrupted(message.to_owned())
}

#[derive(Debug, Clone, Copy)]
struct RawPptRecord<'data> {
    kind: u16,
    data: &'data [u8],
}

/// Strict, allocation-free iterator for records embedded in host payloads.
struct RawPptRecords<'data> {
    data: &'data [u8],
    offset: usize,
    seen: u32,
    failed: bool,
}

impl<'data> RawPptRecords<'data> {
    fn new(data: &'data [u8]) -> Self {
        Self {
            data,
            offset: 0,
            seen: 0,
            failed: false,
        }
    }
}

impl<'data> Iterator for RawPptRecords<'data> {
    type Item = Result<RawPptRecord<'data>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.failed || self.offset == self.data.len() {
            return None;
        }
        self.seen = match self.seen.checked_add(1) {
            Some(seen) if seen <= MAX_HOST_RECORDS => seen,
            _ => {
                self.failed = true;
                return Some(Err(corrupted("Host payload exceeds the PPT record limit")));
            },
        };
        let Some(header_end) = self.offset.checked_add(8) else {
            self.failed = true;
            return Some(Err(corrupted("PPT record header offset overflow")));
        };
        let Some(header) = self.data.get(self.offset..header_end) else {
            self.failed = true;
            return Some(Err(corrupted(
                "Host payload ends with a truncated PPT record header",
            )));
        };
        let kind = u16::from_le_bytes([header[2], header[3]]);
        let len = match usize::try_from(u32::from_le_bytes([
            header[4], header[5], header[6], header[7],
        ])) {
            Ok(len) => len,
            Err(_) => {
                self.failed = true;
                return Some(Err(corrupted(
                    "PPT record length cannot be represented on this platform",
                )));
            },
        };
        let Some(end) = header_end.checked_add(len) else {
            self.failed = true;
            return Some(Err(corrupted("PPT record payload offset overflow")));
        };
        let Some(data) = self.data.get(header_end..end) else {
            self.failed = true;
            return Some(Err(corrupted(
                "Host payload contains a truncated PPT record",
            )));
        };
        self.offset = end;
        Some(Ok(RawPptRecord { kind, data }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    #[test]
    fn ppt_text_atoms_use_their_specified_encodings() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&4000u16.to_le_bytes());
        bytes.extend_from_slice(&4u32.to_le_bytes());
        bytes.extend_from_slice(&[0x3d, 0xd8, 0x00, 0xde]);
        bytes.extend_from_slice(&0u16.to_le_bytes());
        bytes.extend_from_slice(&4008u16.to_le_bytes());
        bytes.extend_from_slice(&2u32.to_le_bytes());
        bytes.extend_from_slice(&[0x80, 0xe9]);

        let mut text = String::new();
        text_from_ppt_records(&bytes, &mut text).unwrap();

        assert_eq!(text, "😀\n\u{80}é");
    }

    #[test]
    fn malformed_embedded_record_stops_without_panicking() {
        let mut text = String::new();
        assert!(text_from_ppt_records(&[0; 7], &mut text).is_err());
        assert!(text.is_empty());
    }

    #[test]
    fn concatenated_drawing_roots_are_complete_and_strict() {
        let dg = record(0, 0, RecordKind::Dg.raw(), &[0; 8]);
        let root = record(0x0f, 0, RecordKind::DgContainer.raw(), &dg);
        let stream = [root.as_slice(), root.as_slice()].concat();

        assert!(parse(&stream).unwrap().is_empty());
        assert_eq!(text_from_drawing(&stream).unwrap(), "");

        let malformed = [stream.as_slice(), &[0]].concat();
        assert!(parse(&malformed).is_err());
        assert!(text_from_drawing(&malformed).is_err());
    }

    #[test]
    fn ppt_client_anchor_projects_small_rect_order() {
        let mut shape_atom = Vec::new();
        shape_atom.extend_from_slice(&42u32.to_le_bytes());
        shape_atom.extend_from_slice(&0x0A00u32.to_le_bytes());
        let mut shape = record(2, 1, RecordKind::Sp.raw(), &shape_atom);
        let anchor_data = [
            20i16.to_le_bytes(),
            10i16.to_le_bytes(),
            110i16.to_le_bytes(),
            70i16.to_le_bytes(),
        ]
        .concat();
        shape.extend(record(0, 0, RecordKind::ClientAnchor.raw(), &anchor_data));
        let shape = record(0x0f, 0, RecordKind::SpContainer.raw(), &shape);
        let mut drawing_children = record(0, 0, RecordKind::Dg.raw(), &[0; 8]);
        drawing_children.extend(shape);
        let drawing = record(0x0f, 0, RecordKind::DgContainer.raw(), &drawing_children);

        let shapes = parse(&drawing).unwrap();
        let anchor = anchor(&shapes[0]).unwrap().unwrap();
        assert_eq!((anchor.left(), anchor.top()), (10, 20));
        assert_eq!((anchor.right(), anchor.bottom()), (110, 70));
        assert_eq!((anchor.width(), anchor.height()), (100, 50));
    }
}
