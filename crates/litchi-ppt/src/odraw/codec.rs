use litchi_odraw::shape::Shape;
use litchi_odraw::{Container, Record, RecordKind};

use crate::embedded::reference::Reference;
use crate::package::Result;

use super::model::{Anchor, Frame, FrameKind, Placeholder, ShapeExt};
use super::package::{
    CLIENT_DATA_RAW_KIND, CLIENT_TEXTBOX_RAW_KIND, RawRecords, advance, corrupted, host_record,
    validate_host_record, visit_host_record,
};

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

    fn interactions(&self) -> Result<Vec<crate::Interaction>> {
        self.interactions_with_limits(crate::InteractionLimits::default())
    }

    fn interactions_with_limits(
        &self,
        limits: crate::InteractionLimits,
    ) -> Result<Vec<crate::Interaction>> {
        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(Vec::new());
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
        crate::Interaction::parse_client_data_payload(client_data.data(), limits)
    }

    fn text_interactions(&self) -> Result<Vec<crate::TextInteraction>> {
        self.text_interactions_with_limits(crate::TextInteractionLimits::default())
    }

    fn text_interactions_with_limits(
        &self,
        limits: crate::TextInteractionLimits,
    ) -> Result<Vec<crate::TextInteraction>> {
        let Some(textbox) = host_record(self, RecordKind::ClientTextbox)? else {
            return Ok(Vec::new());
        };
        validate_host_record(&textbox, CLIENT_TEXTBOX_RAW_KIND, "ClientTextbox")?;
        crate::EscherTextboxWrapper::parse_text_interactions_with_limits(textbox.data(), limits)
    }

    fn placeholder_atom(
        &self,
        context: crate::PlaceholderContext,
    ) -> Result<Option<crate::PlaceholderAtom>> {
        self.placeholder_atom_with_limits(context, crate::PlaceholderLimits::default())
    }

    fn placeholder_atom_with_limits(
        &self,
        context: crate::PlaceholderContext,
        limits: crate::PlaceholderLimits,
    ) -> Result<Option<crate::PlaceholderAtom>> {
        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(None);
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
        Ok(crate::PlaceholderProjection::parse_client_data_payload(
            client_data.data(),
            context,
            limits,
        )?
        .placeholder)
    }

    fn powerpoint12_shape_metadata(&self) -> Result<Option<crate::ShapeMetadata>> {
        use crate::consts::RecordType;
        use crate::{HeaderFooterPlaceholder, NewPlaceholder, ShapeChecksums, ShapeMetadata};

        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(None);
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;

        let mut metadata = ShapeMetadata::default();
        let mut found = false;
        let mut offset = 0usize;
        let mut records = 0u32;
        while offset < client_data.data().len() {
            visit_host_record(&mut records)?;
            let (record, consumed) =
                crate::records::Record::parse_strict(client_data.data(), offset)?;
            match record.record_type {
                RecordType::RoundTripHFPlaceholder12Atom => {
                    if metadata.header_footer.is_some() {
                        return Err(corrupted(
                            "Shape contains duplicate RoundTripHFPlaceholder12Atom records",
                        ));
                    }
                    validate_round_trip_atom(&record, "RoundTripHFPlaceholder12Atom", 1)?;
                    metadata.header_footer = Some(match record.data[0] {
                        7 => HeaderFooterPlaceholder::Date,
                        8 => HeaderFooterPlaceholder::SlideNumber,
                        9 => HeaderFooterPlaceholder::Footer,
                        10 => HeaderFooterPlaceholder::Header,
                        _ => {
                            return Err(corrupted(
                                "RoundTripHFPlaceholder12Atom has an invalid placeholder ID",
                            ));
                        },
                    });
                    found = true;
                },
                RecordType::RoundTripNewPlaceholderId12Atom => {
                    if metadata.new_placeholder.is_some() {
                        return Err(corrupted(
                            "Shape contains duplicate RoundTripNewPlaceholderId12Atom records",
                        ));
                    }
                    validate_round_trip_atom(&record, "RoundTripNewPlaceholderId12Atom", 1)?;
                    metadata.new_placeholder = Some(match record.data[0] {
                        25 => NewPlaceholder::VerticalObject,
                        26 => NewPlaceholder::Picture,
                        _ => {
                            return Err(corrupted(
                                "RoundTripNewPlaceholderId12Atom has an invalid placeholder ID",
                            ));
                        },
                    });
                    found = true;
                },
                RecordType::RoundTripShapeId12Atom => {
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
                RecordType::RoundTripShapeCheckSumForCustomLayouts12Atom => {
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
                    metadata.custom_layout_checksums = Some(ShapeChecksums {
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
        limits: crate::ShapeProgrammableTagLimits,
    ) -> Result<Option<crate::ShapeProgrammableTags>> {
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
                crate::records::Record::parse_strict(client_data.data(), offset)?;
            if record.record_type == crate::consts::RecordType::ProgTags {
                let parsed = crate::ShapeProgrammableTags::parse(&record, limits)?;
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

    fn programmable_tags(&self) -> Result<Option<crate::ShapeProgrammableTags>> {
        self.programmable_tags_with_limits(crate::ShapeProgrammableTagLimits::default())
    }

    fn ppt_flags(&self) -> Result<Option<crate::ShapeFlagProjection>> {
        self.ppt_flags_with(crate::ShapeFlagLimits::default())
    }

    fn ppt_flags_with(
        &self,
        limits: crate::ShapeFlagLimits,
    ) -> Result<Option<crate::ShapeFlagProjection>> {
        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(None);
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
        let projection =
            crate::ShapeFlagProjection::parse_client_data_payload(client_data.data(), limits)?;
        Ok(projection.has_flags().then_some(projection))
    }

    fn animation(&self) -> Result<Option<crate::animation::AnimationInfo>> {
        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(None);
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
        let mut offset = 0usize;
        let mut records = 0u32;
        while offset < client_data.data().len() {
            visit_host_record(&mut records)?;
            let (record, consumed) =
                crate::records::Record::parse_strict(client_data.data(), offset)?;
            if record.record_type == crate::consts::RecordType::AnimationInfo {
                return crate::animation::parse_animation_info(&record).map(Some);
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
    let anchor = crate::ClientAnchorData::parse(anchor.data())?;
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

/// Returns a shape's header-validated PowerPoint textbox record.
pub(crate) fn textbox<'data>(shape: &Shape<'data>) -> Result<Option<Record<'data>>> {
    let Some(textbox) = host_record(shape, RecordKind::ClientTextbox)? else {
        return Ok(None);
    };
    validate_host_record(&textbox, CLIENT_TEXTBOX_RAW_KIND, "ClientTextbox")?;
    Ok(Some(textbox))
}

fn placeholder(shape: &Shape<'_>) -> Result<Option<Placeholder>> {
    const PLACEHOLDER_ATOM: u16 = 3011;

    let Some(client_data) = host_record(shape, RecordKind::ClientData)? else {
        return Ok(None);
    };
    validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
    let mut records = RawRecords::new(client_data.data());
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
            let kind = crate::PlaceholderKind::try_from(record.data[4])?;
            let size = crate::AtomPlaceholderSize::try_from(record.data[5])?;
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
    let mut records = RawRecords::new(client_data.data());
    for record in &mut records {
        let record = record?;
        visit_host_record(&mut host_records)?;
        match record.kind {
            EX_OBJ_REF_ATOM => {
                if frame.object_id.is_some() {
                    return Err(corrupted(
                        "Shape ClientData contains multiple ExObjRefAtom records",
                    ));
                }
                frame.object_id = Some(Reference::parse_payload(record.data)?.id);
                if frame.kind == FrameKind::Picture {
                    frame.kind = FrameKind::Object;
                }
            },
            INTERACTIVE_INFO => {
                let mut children = RawRecords::new(record.data);
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

pub(super) fn text_from_ppt_records(data: &[u8], text: &mut String) -> Result<()> {
    const MAX_RECORDS: usize = 1_000_000;
    const MAX_DEPTH: usize = 256;

    let mut pending = vec![RawRecords::new(data)];
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
            4000 | 4026 => Some(crate::text::extractor::from_utf16le_lossy(record.data)),
            4008 => Some(crate::text::extractor::decode_text_bytes(record.data)),
            kind if is_text_container(kind) => {
                if pending.len() >= MAX_DEPTH {
                    return Err(corrupted("ClientTextbox exceeds the PPT nesting limit"));
                }
                pending.push(RawRecords::new(record.data));
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
    record: &crate::records::Record,
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
