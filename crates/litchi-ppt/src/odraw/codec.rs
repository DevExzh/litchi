use litchi_odraw::shape::Shape;
use litchi_odraw::{Container, Record, RecordKind};

use crate::embedded::reference::Reference;
use crate::package::{Result, SourceLimits};

use super::model::{Anchor, Drawing, Frame, FrameKind, Placeholder, ShapeExt, ShapeLimitsExt};
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

    fn text_interactions(&self) -> Result<Vec<crate::TextInteraction>> {
        self.text_interactions_with_limits(crate::TextInteractionLimits::default())
    }

    fn placeholder_atom(
        &self,
        context: crate::PlaceholderContext,
    ) -> Result<Option<crate::PlaceholderAtom>> {
        self.placeholder_atom_with_limits(context, crate::PlaceholderLimits::default())
    }

    fn powerpoint12_shape_metadata(&self) -> Result<Option<crate::ShapeMetadata>> {
        powerpoint12_shape_metadata(self, None)
    }

    fn programmable_tags(&self) -> Result<Option<crate::ShapeProgrammableTags>> {
        self.programmable_tags_with_limits(crate::ShapeProgrammableTagLimits::default())
    }

    fn ppt_flags(&self) -> Result<Option<crate::ShapeFlagProjection>> {
        self.ppt_flags_with(crate::ShapeFlagLimits::default())
    }

    fn animation(&self) -> Result<Option<crate::animation::AnimationInfo>> {
        let Some(client_data) = host_record(self, RecordKind::ClientData)? else {
            return Ok(None);
        };
        validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;
        let mut offset = 0usize;
        let mut records = 0u32;
        let max_records = self.container().record().limits().max_records;
        while offset < client_data.data().len() {
            visit_host_record(&mut records, max_records)?;
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

impl ShapeLimitsExt for Shape<'_> {
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
        let max_records = self.container().record().limits().max_records;
        while offset < client_data.data().len() {
            visit_host_record(&mut records, max_records)?;
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
}

/// Parse every complete `OfficeArt` drawing in a `PowerPoint` host stream.
///
/// A PPT `PPDrawing` payload can concatenate drawing containers. Each root is
/// still parsed exactly: malformed records and trailing partial records are
/// returned as errors instead of being ignored.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse<'data>(data: &'data [u8], limits: &litchi_odraw::Limits) -> Result<Vec<Shape<'data>>> {
    parse_drawing(data, limits).map(Drawing::into_shapes)
}

/// Derive the neutral OfficeArt traversal policy from the selected PPT
/// source profile. A source-backed slide projection must not reopen its
/// drawing with the neutral crate's defaults.
pub(crate) fn limits_for_source(source: SourceLimits) -> litchi_odraw::Limits {
    // `SourceLimits` governs the outer PPT record tree and deliberately allows
    // deeper nesting than the shape decoder can safely traverse.  Intersect
    // the host profile with OfficeArt's own hard safety ceiling before handing
    // it to the format-neutral parser; otherwise the default PPT profile (256)
    // is rejected by `litchi-odraw` (64) even for ordinary source-backed
    // projections.
    let defaults = litchi_odraw::Limits::default();
    // PPT's record ceiling includes its eight-byte header, while OfficeArt's
    // raw limit is the declared body length. Keep both PPT ceilings active
    // when crossing that wire-format boundary.
    litchi_odraw::Limits {
        max_record_bytes: source
            .max_record_bytes()
            .saturating_sub(8)
            .min(source.max_record_payload_bytes())
            .min(u32::MAX as usize) as u32,
        max_depth: source
            .max_depth()
            .min(usize::from(defaults.max_depth))
            .min(u16::MAX as usize) as u16,
        max_records: source.max_events().min(u32::MAX as usize) as u32,
    }
}

/// Parse a `PowerPoint` `OfficeArt` stream while retaining its validated root
/// records beside the typed shape projection.
///
/// A `PPDrawing` payload may concatenate drawing containers. Every complete
/// root is retained in source order, including direct records that the typed
/// shape view does not model. The existing resource ceilings still apply to
/// root traversal and shape parsing.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_drawing<'data>(
    data: &'data [u8],
    limits: &litchi_odraw::Limits,
) -> Result<Drawing<'data>> {
    let mut roots = Vec::new();
    let mut shapes = Vec::new();
    let mut offset = 0usize;
    let mut visited = 0u32;
    while offset < data.len() {
        visit_host_record(&mut visited, limits.max_records)?;
        let (record, consumed) = Record::parse(data, offset, *limits)?;
        let end = advance(offset, consumed, "PowerPoint OfficeArt drawing stream")?;
        shapes.extend(litchi_odraw::shape::parse(&data[offset..end], *limits)?);
        roots.push(record);
        offset = end;
    }
    Ok(Drawing::new(roots, shapes))
}

/// Project a shape's group-relative or PPT client anchor into checked bounds.
///
/// # Errors
///
/// Returns an error if the operation fails.
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
    let parsed =
        crate::client_anchor::Data::parse(anchor.data(), &crate::client_anchor::Limits::default())?;
    Anchor::new(parsed.left(), parsed.top(), parsed.right(), parsed.bottom()).map(Some)
}

/// Extract all PPT textbox text from an `OfficeArt` drawing.
///
/// Each `ClientTextbox` contributes the text records it holds inline. A
/// text box that only references slide-list text (an `OutlineTextRefAtom`,
/// MS-PPT 2.9.78) contributes nothing here, because this function has no
/// slide to resolve the reference against; [`crate::Slide::text`] resolves
/// it.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn text_from_drawing(data: &[u8], source: SourceLimits) -> Result<String> {
    drawing_text(data, None, limits_for_source(source))
}

/// Extract a presentation slide drawing's text, reading each text box that
/// references slide-list text (an `OutlineTextRefAtom`, MS-PPT 2.9.78) from
/// the body of `outline` it selects (L-PPT-R1).
///
/// Text boxes appear in drawing order. A reference that selects no readable
/// body contributes no text, as POI and LibreOffice show none.
pub(crate) fn text_from_slide_drawing(
    data: &[u8],
    outline: &[crate::OutlineTextBody],
    source: SourceLimits,
) -> Result<String> {
    drawing_text(data, Some(outline), limits_for_source(source))
}

/// `outline` is `None` when no slide context exists: every text box is then
/// read from its inline records, as before outline resolution existed.
fn drawing_text(
    data: &[u8],
    outline: Option<&[crate::OutlineTextBody]>,
    limits: litchi_odraw::Limits,
) -> Result<String> {
    let mut offset = 0usize;
    let mut records = 0u32;
    let mut text = String::with_capacity(1024);
    while offset < data.len() {
        visit_host_record(&mut records, limits.max_records)?;
        let (record, consumed) = Record::parse(data, offset, limits)?;
        let container = Container::try_new(record)?;
        text_from_container(&container, outline, &mut text)?;
        offset = advance(offset, consumed, "PowerPoint OfficeArt drawing stream")?;
    }
    Ok(text)
}

/// Decode the PPT records contained by one `OfficeArt` `ClientTextbox` atom.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn text_from_textbox(textbox: &Record<'_>) -> Result<Option<String>> {
    validate_host_record(textbox, CLIENT_TEXTBOX_RAW_KIND, "ClientTextbox")?;
    text_from_textbox_payload_bounded(textbox.data(), textbox.limits().max_records)
}

/// Decode the text of a `ClientTextbox` payload, or of a slide-list text
/// body framed as one: every character atom's text, trimmed and joined with
/// `\n`, then trimmed at the end; `None` when nothing remains.
pub(crate) fn text_from_textbox_payload(payload: &[u8]) -> Result<Option<String>> {
    text_from_textbox_payload_bounded(payload, super::package::MAX_HOST_RECORDS)
}

pub(crate) fn text_from_textbox_payload_bounded(
    payload: &[u8],
    max_records: u32,
) -> Result<Option<String>> {
    if payload.is_empty() {
        return Ok(None);
    }
    let mut text = String::with_capacity(256);
    text_from_ppt_records(payload, &mut text, max_records)?;
    let trimmed = text.trim_end().len();
    text.truncate(trimmed);
    Ok((!text.is_empty()).then_some(text))
}

/// Where a shape's text records live (MS-PPT 2.9.76
/// `OfficeArtClientTextBoxAtoms`).
#[derive(Debug, Clone, Copy)]
pub(crate) enum TextRecords<'a> {
    /// The `ClientTextbox` holds the text records itself.
    Inline(&'a [u8]),
    /// The `ClientTextbox` holds an `OutlineTextRefAtom` that selects this
    /// slide-list text body.
    Outline {
        /// The reference's zero-based body index.
        index: u32,
        /// The selected body.
        body: &'a crate::OutlineTextBody,
    },
    /// The `ClientTextbox` holds an `OutlineTextRefAtom` that selects no
    /// body: its index is negative or past the slide's last body, or the
    /// atom is not four bytes. The value is the index field reinterpreted as
    /// unsigned, or `u32::MAX` for a malformed atom; it is the context of
    /// [`crate::ShapeDefect::UnresolvedOutlineTextRef`].
    Unresolved(u32),
}

/// Classifies a `ClientTextbox` payload against a slide's slide-list text
/// bodies.
///
/// The first `OutlineTextRefAtom` decides, and inline text records beside it
/// are ignored, as POI (`HSLFTextParagraph.findTextParagraphs`) and
/// LibreOffice (`PPTTextObj`) ignore them; MS-PPT allows either an outline
/// reference or inline text, never both. Without a reference the payload is
/// read inline.
///
/// # Errors
///
/// Returns an error when the payload's record framing is broken, exactly as
/// decoding its inline text would.
pub(crate) fn textbox_text_records<'a>(
    payload: &'a [u8],
    outline: &'a [crate::OutlineTextBody],
    max_records: u32,
) -> Result<TextRecords<'a>> {
    const OUTLINE_TEXT_REF_ATOM: u16 = 0x0f9e;

    for raw in RawRecords::new(payload, max_records) {
        let record = raw?;
        if record.kind != OUTLINE_TEXT_REF_ATOM {
            continue;
        }
        let Ok(bytes) = <[u8; 4]>::try_from(record.data) else {
            return Ok(TextRecords::Unresolved(u32::MAX));
        };
        if record.version != 0 {
            return Ok(TextRecords::Unresolved(u32::MAX));
        }
        let raw_index = u32::from_le_bytes(bytes);
        // A negative signed index selects nothing.
        if i32::from_le_bytes(bytes) < 0 {
            return Ok(TextRecords::Unresolved(raw_index));
        }
        return Ok(usize::try_from(raw_index)
            .ok()
            .and_then(|index| outline.get(index))
            .map_or(TextRecords::Unresolved(raw_index), |body| {
                TextRecords::Outline {
                    index: raw_index,
                    body,
                }
            }));
    }
    Ok(TextRecords::Inline(payload))
}

/// The shape's name and alternative text (L-PPT-R3): the `wzName` (0x0380,
/// MS-ODRAW 2.3.4.1) and `wzDescription` (0x0381, MS-ODRAW 2.3.4.3)
/// properties of its primary property table, where POI
/// (`HSLFShape.getShapeName`) and LibreOffice (`SvxMSDffManager::ImportObj`)
/// read them.
///
/// Each is `None` when the property is absent, not complex, or empty. The
/// table is walked once; the property parser already refuses a repeated
/// property identifier.
pub(crate) fn shape_labels(shape: &Shape<'_>) -> (Option<String>, Option<String>) {
    use litchi_odraw::prop::{Id, Value};

    let mut name = None;
    let mut description = None;
    for prop in shape.props().iter() {
        let slot = match prop.id() {
            Id::GroupName => &mut name,
            Id::GroupDescription => &mut description,
            _ => continue,
        };
        if let Value::Complex(data) = prop.value() {
            *slot = utf16z_string(data);
        }
    }
    (name, description)
}

/// Decodes a null-terminated UTF-16LE complex property, as POI's
/// `StringUtil.getFromUnicodeLE0Terminated` does: characters stop at the
/// first NUL, a trailing odd byte is ignored, and unpaired surrogates read as
/// U+FFFD. `None` when no character remains.
fn utf16z_string(data: &[u8]) -> Option<String> {
    let units = data
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| u16::from_le_bytes(*pair))
        .take_while(|&unit| unit != 0);
    let mut text = String::with_capacity(data.len() / 2);
    text.extend(char::decode_utf16(units).map(|unit| unit.unwrap_or(char::REPLACEMENT_CHARACTER)));
    (!text.is_empty()).then_some(text)
}

/// Checks the record framing of a shape's `PowerPoint` host records: the
/// `ClientTextbox` and `ClientData` headers and the header of every child
/// record they hold.
///
/// Shape-level isolation (finding F5.11) drops a shape whose *content*
/// cannot be read; a shape failing this check is structural corruption and
/// stays fatal for the whole drawing.
pub(crate) fn validate_host_framing(shape: &Shape<'_>) -> Result<()> {
    let max_records = shape.container().record().limits().max_records;
    let mut visited = 0u32;
    for (kind, raw_kind, name) in [
        (
            RecordKind::ClientTextbox,
            CLIENT_TEXTBOX_RAW_KIND,
            "ClientTextbox",
        ),
        (RecordKind::ClientData, CLIENT_DATA_RAW_KIND, "ClientData"),
    ] {
        let Some(record) = host_record(shape, kind)? else {
            continue;
        };
        validate_host_record(&record, raw_kind, name)?;
        for child in RawRecords::new(record.data(), max_records) {
            child?;
            visit_host_record(&mut visited, max_records)?;
        }
    }
    Ok(())
}

/// Returns a shape's header-validated `PowerPoint` textbox record.
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
    let mut records = RawRecords::new(
        client_data.data(),
        shape.container().record().limits().max_records,
    );
    let mut found = None;
    for raw in &mut records {
        let record = raw?;
        if record.kind == PLACEHOLDER_ATOM {
            if record.data.len() < 8 {
                return Err(corrupted("PlaceholderAtom is shorter than eight bytes"));
            }
            if found.is_some() {
                return Err(corrupted(
                    "Shape ClientData contains multiple PlaceholderAtom records",
                ));
            }
            let position_bytes: [u8; 4] = record.data[..4]
                .try_into()
                .map_err(|_err| corrupted("PlaceholderAtom position is not four bytes"))?;
            let position_raw = i32::from_le_bytes(position_bytes);
            let position = match position_raw {
                -1 => None,
                index => Some(u16::try_from(index).map_err(|_err| {
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
    let max_records = shape.container().record().limits().max_records;
    let mut records = RawRecords::new(client_data.data(), max_records);
    for raw in &mut records {
        let record = raw?;
        visit_host_record(
            &mut host_records,
            shape.container().record().limits().max_records,
        )?;
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
                let mut children =
                    RawRecords::new(record.data, shape.container().record().limits().max_records);
                for raw_child in &mut children {
                    let child = raw_child?;
                    visit_host_record(
                        &mut host_records,
                        shape.container().record().limits().max_records,
                    )?;
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
                    if let Some(frame_kind) = kind
                        && action_kind.replace(frame_kind).is_some()
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

fn text_from_container(
    container: &Container<'_>,
    outline: Option<&[crate::OutlineTextBody]>,
    text: &mut String,
) -> Result<()> {
    let limits = container.record().limits();
    let max_records = usize::try_from(limits.max_records).unwrap_or(usize::MAX);
    let max_depth = usize::from(limits.max_depth);

    let mut pending = Vec::new();
    push_children(container, 0, &mut pending, max_records)?;
    let mut visited = 0usize;

    while let Some((record, depth)) = pending.pop() {
        visited = visited
            .checked_add(1)
            .ok_or_else(|| corrupted("OfficeArt traversal count overflow"))?;
        if visited > max_records {
            return Err(corrupted("OfficeArt drawing exceeds the PPT record limit"));
        }

        if record.kind() == RecordKind::ClientTextbox {
            let before = text.len();
            match outline {
                None => text_from_ppt_records(record.data(), text, limits.max_records)?,
                Some(bodies) => {
                    match textbox_text_records(record.data(), bodies, limits.max_records)? {
                        TextRecords::Inline(payload) => {
                            text_from_ppt_records(payload, text, limits.max_records)?
                        },
                        TextRecords::Outline { body, .. } => {
                            outline_text_into(body, text, limits.max_records)?;
                        },
                        TextRecords::Unresolved(_) => {},
                    }
                },
            }
            let trimmed = text.trim_end().len();
            text.truncate(trimmed);
            if text.len() > before && !text.ends_with('\n') {
                text.push('\n');
            }
            continue;
        }

        if record.is_container() {
            if depth >= max_depth {
                return Err(corrupted("OfficeArt drawing exceeds the PPT nesting limit"));
            }
            let nested = Container::try_new(record)?;
            push_children(&nested, depth + 1, &mut pending, max_records)?;
        }
    }
    Ok(())
}

/// Appends the text of a slide-list body that a text box references, joined
/// as the text box's own records would be.
///
/// A body whose records cannot be traversed contributes no text, as a shape
/// that references it shows none (the shape reader records it as
/// [`crate::ShapeDefect::UnresolvedOutlineTextRef`]); resource limits stay
/// fatal.
fn outline_text_into(
    body: &crate::OutlineTextBody,
    text: &mut String,
    max_records: u32,
) -> Result<()> {
    let mark = text.len();
    match text_from_ppt_records(body.records(), text, max_records) {
        Ok(()) => Ok(()),
        Err(error) if crate::slide::types::is_isolatable_shape_error(&error) => {
            text.truncate(mark);
            Ok(())
        },
        Err(error) => Err(error),
    }
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

pub(super) fn text_from_ppt_records(
    data: &[u8],
    text: &mut String,
    max_records: u32,
) -> Result<()> {
    const MAX_DEPTH: usize = 256;
    let max_records_u32 = max_records;
    let max_records = usize::try_from(max_records_u32).unwrap_or(usize::MAX);

    let mut pending = vec![RawRecords::new(data, max_records_u32)];
    let mut visited = 0usize;
    while let Some(records) = pending.last_mut() {
        let Some(raw) = records.next() else {
            pending.pop();
            continue;
        };
        let record = raw?;
        visited = visited
            .checked_add(1)
            .ok_or_else(|| corrupted("ClientTextbox traversal count overflow"))?;
        if visited > max_records {
            return Err(corrupted("ClientTextbox exceeds the PPT record limit"));
        }

        let decoded = match record.kind {
            4000 | 4026 => Some(crate::text::extractor::from_utf16le_lossy(record.data)),
            4008 => Some(crate::text::extractor::decode_text_bytes(record.data)),
            kind if is_text_container(kind) => {
                if pending.len() >= MAX_DEPTH {
                    return Err(corrupted("ClientTextbox exceeds the PPT nesting limit"));
                }
                pending.push(RawRecords::new(record.data, max_records_u32));
                None
            },
            _ => None,
        };
        if let Some(piece) = decoded {
            let trimmed = piece.trim();
            if trimmed.is_empty() {
                continue;
            }
            if !text.is_empty() && !text.ends_with('\n') {
                text.push('\n');
            }
            text.push_str(trimmed);
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

/// Parses `PowerPoint` 12 shape round-trip metadata.
///
/// Without a report every deviation is an error. With a report (the slide
/// and notes read path), a duplicated round-trip atom keeps its first
/// occurrence and an undefined placeholder ID is treated as absent, each
/// recorded as a [`crate::ShapeDefect`] (finding F5.11). Record framing and
/// each atom's header and size stay errors either way.
pub(crate) fn powerpoint12_shape_metadata(
    shape: &Shape<'_>,
    mut report: Option<&mut crate::leniency::ToleranceReport>,
) -> Result<Option<crate::ShapeMetadata>> {
    use crate::consts::RecordType;
    use crate::leniency::ShapeDefect;
    use crate::{HeaderFooterPlaceholder, NewPlaceholder, ShapeChecksums, ShapeMetadata};

    let Some(client_data) = host_record(shape, RecordKind::ClientData)? else {
        return Ok(None);
    };
    validate_host_record(&client_data, CLIENT_DATA_RAW_KIND, "ClientData")?;

    // Tolerates `defect` into the report when one is threaded, else rejects
    // the shape with `message`.
    let mut tolerate =
        |defect: ShapeDefect, context: u32, message: &str| match report.as_deref_mut() {
            Some(report) => {
                report.record(defect, context);
                Ok(())
            },
            None => Err(corrupted(message)),
        };

    let mut metadata = ShapeMetadata::default();
    let mut found = false;
    let mut offset = 0usize;
    let mut records = 0u32;
    let max_records = shape.container().record().limits().max_records;
    while offset < client_data.data().len() {
        visit_host_record(&mut records, max_records)?;
        let (record, consumed) = crate::records::Record::parse_strict(client_data.data(), offset)?;
        offset = advance(offset, consumed, "OfficeArt client-data")?;
        let record_type = u32::from(record.record_type.as_u16());
        #[allow(
            clippy::wildcard_enum_match_arm,
            reason = "`RecordType` spans the full MS-PPT record ID space; only the four round-trip atoms carry shape metadata, every other record is skipped"
        )]
        match record.record_type {
            RecordType::RoundTripHFPlaceholder12Atom => {
                if metadata.header_footer.is_some() {
                    tolerate(
                        ShapeDefect::DuplicateRoundTripAtom,
                        record_type,
                        "Shape contains duplicate RoundTripHFPlaceholder12Atom records",
                    )?;
                    continue;
                }
                validate_round_trip_atom(&record, "RoundTripHFPlaceholder12Atom", 1)?;
                let header_footer = match record.data[0] {
                    7 => HeaderFooterPlaceholder::Date,
                    8 => HeaderFooterPlaceholder::SlideNumber,
                    9 => HeaderFooterPlaceholder::Footer,
                    10 => HeaderFooterPlaceholder::Header,
                    id => {
                        tolerate(
                            ShapeDefect::UnknownRoundTripPlaceholderId,
                            (record_type << 16) | u32::from(id),
                            "RoundTripHFPlaceholder12Atom has an invalid placeholder ID",
                        )?;
                        continue;
                    },
                };
                metadata.header_footer = Some(header_footer);
                found = true;
            },
            RecordType::RoundTripNewPlaceholderId12Atom => {
                if metadata.new_placeholder.is_some() {
                    tolerate(
                        ShapeDefect::DuplicateRoundTripAtom,
                        record_type,
                        "Shape contains duplicate RoundTripNewPlaceholderId12Atom records",
                    )?;
                    continue;
                }
                validate_round_trip_atom(&record, "RoundTripNewPlaceholderId12Atom", 1)?;
                let new_placeholder = match record.data[0] {
                    25 => NewPlaceholder::VerticalObject,
                    26 => NewPlaceholder::Picture,
                    id => {
                        tolerate(
                            ShapeDefect::UnknownRoundTripPlaceholderId,
                            (record_type << 16) | u32::from(id),
                            "RoundTripNewPlaceholderId12Atom has an invalid placeholder ID",
                        )?;
                        continue;
                    },
                };
                metadata.new_placeholder = Some(new_placeholder);
                found = true;
            },
            RecordType::RoundTripShapeId12Atom => {
                if metadata.shape_id.is_some() {
                    tolerate(
                        ShapeDefect::DuplicateRoundTripAtom,
                        record_type,
                        "Shape contains duplicate RoundTripShapeId12Atom records",
                    )?;
                    continue;
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
                    tolerate(
                        ShapeDefect::DuplicateRoundTripAtom,
                        record_type,
                        "Shape contains duplicate RoundTripShapeCheckSumForCustomLayouts12Atom records",
                    )?;
                    continue;
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
    }

    Ok(found.then_some(metadata))
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
