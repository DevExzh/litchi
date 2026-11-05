//! Character runs and text-range hyperlinks of one text body (record 0323).
//!
//! A shape's text body is its `ClientTextbox` records, or, for a title or
//! body placeholder whose textbox holds only an `OutlineTextRefAtom`, the
//! slide-list body the reference selects. Its character formatting is the
//! `StyleTextPropAtom` that follows the text atom ([MS-PPT] 2.9.44): runs of
//! `TextCFException` values over UTF-16 code units, together covering the
//! text and its final paragraph mark. Its text hyperlinks are
//! `InteractiveInfo` containers each followed by a `TextInteractiveInfoAtom`
//! range ([MS-PPT] 2.9.?).
//!
//! An edit rewrites the body's style atom or interaction pairs and re-frames
//! only the records that contain them; every other record keeps its bytes.

use crate::SourceLimits;
use crate::consts::RecordType;
use crate::hyperlink::{Interaction, InteractionLimits};
use crate::package::Error as PackageError;
use crate::text_format_exception::TextCFException;

use super::{
    Error, OFFICEART_CLIENT_TEXTBOX, OFFICEART_SP_CONTAINER, Refusal, Result, TextKind, children,
    decode_utf16, drawing_payload, find_ppt_record, for_each_outline_body, outline_reference,
    rebuild, record_instance, record_type, record_version, rewrite_ppt_record, shape_id_of,
    slide_list_of,
};

/// One character-formatting run of a text body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CharacterRun {
    /// First UTF-16 code unit the run covers.
    pub start: usize,
    /// Number of UTF-16 code units the run covers; the last run of a body
    /// also covers the final paragraph mark.
    pub length: usize,
    /// The run's character formatting; a body without a style atom reads as
    /// one run with no explicit formatting.
    pub format: TextCFException,
}

/// One click or mouse-over hyperlink over a range of a text body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextLink {
    /// First UTF-16 code unit of the range.
    pub start: usize,
    /// One past the last UTF-16 code unit of the range.
    pub end: usize,
    /// The interaction; resolve it with
    /// [`crate::Presentation::resolve_click`].
    pub interaction: Interaction,
}

/// The text, character runs and hyperlinks of one text body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextRuns {
    text: String,
    runs: Vec<CharacterRun>,
    links: Vec<TextLink>,
}

impl TextRuns {
    /// The body's text; paragraphs end at `\r`.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The character runs, in order, covering the text and its final mark.
    #[must_use]
    pub fn runs(&self) -> &[CharacterRun] {
        &self.runs
    }

    /// The text hyperlinks, in record order.
    #[must_use]
    pub fn links(&self) -> &[TextLink] {
        &self.links
    }
}

/// Character formatting to apply to a range; `None` keeps a property.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct RunFormat {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    /// Font size in points, 1..=4000.
    pub size: Option<u16>,
    /// sRGB color.
    pub color: Option<[u8; 3]>,
}

impl RunFormat {
    fn apply(&self, format: &TextCFException) -> TextCFException {
        let mut format = format.clone();
        for (mask, value) in [
            (0x0001, self.bold),
            (0x0002, self.italic),
            (0x0004, self.underline),
        ] {
            if let Some(value) = value {
                format = format.with_style_flag(mask, value);
            }
        }
        if let Some(size) = self.size.and_then(|size| i16::try_from(size).ok()) {
            format = format.with_size(size);
        }
        if let Some(color) = self.color {
            format = format.with_rgb(color);
        }
        format
    }

    /// Whether `format` already shows every property this format sets.
    fn holds(&self, format: &TextCFException) -> bool {
        self.apply(format) == *format
    }
}

/// Parse one body's records.
fn parse_body<'a>(
    records: impl IntoIterator<Item = &'a [u8]>,
    interaction_limits: InteractionLimits,
) -> Result<TextRuns> {
    let records = records.into_iter().collect::<Vec<_>>();
    let mut text = None;
    for record in &records {
        let kind = record_type(record)?;
        let decoded = if kind == RecordType::TextCharsAtom as u16 {
            decode_utf16(drawing_payload(record)?)?
        } else if kind == RecordType::TextBytesAtom as u16 {
            drawing_payload(record)?
                .iter()
                .map(|byte| char::from(*byte))
                .collect()
        } else {
            continue;
        };
        if text.replace(decoded).is_some() {
            return Err(Error::Refused(Refusal::MultipleTextAtoms));
        }
    }
    let text = text.unwrap_or_default();
    let units = text.encode_utf16().count();
    let mut runs = Vec::new();
    for record in &records {
        if record_type(record)? != RecordType::StyleTextPropAtom as u16 {
            continue;
        }
        let (parsed, _) = crate::Record::parse_strict(record, 0)?;
        let style = crate::StyleTextPropAtom::parse_record(&parsed, units)?;
        let mut start = 0usize;
        for run in style.character_runs() {
            let length = run.count() as usize;
            runs.push(CharacterRun {
                start,
                length,
                format: run.character().clone(),
            });
            start += length;
        }
        break;
    }
    if runs.is_empty() {
        runs.push(CharacterRun {
            start: 0,
            length: units + 1,
            format: TextCFException::default(),
        });
    }
    let mut links = Vec::new();
    let mut pending: Option<Interaction> = None;
    for record in &records {
        let kind = record_type(record)?;
        if kind == RecordType::InteractiveInfo as u16 {
            pending = Some(Interaction::parse_bytes(record, &interaction_limits)?);
        } else if kind == RecordType::TextInteractiveInfoAtom as u16 {
            let Some(interaction) = pending.take() else {
                continue;
            };
            let payload = drawing_payload(record)?;
            let field = |offset: usize| -> Result<usize> {
                let bytes: [u8; 4] = payload
                    .get(offset..offset + 4)
                    .and_then(|bytes| bytes.try_into().ok())
                    .ok_or_else(|| {
                        PackageError::Corrupted("TextInteractiveInfoAtom is truncated".into())
                    })?;
                usize::try_from(i32::from_le_bytes(bytes))
                    .map_err(|_err| PackageError::Corrupted("text range is negative".into()).into())
            };
            links.push(TextLink {
                start: field(0)?,
                end: field(4)?,
                interaction,
            });
        } else {
            pending = None;
        }
    }
    Ok(TextRuns { text, runs, links })
}

/// The `ClientTextbox` of the shape with `OfficeArt` ID `shape_id` in a
/// drawing payload, or `None` when the shape has none.
fn textbox_in(record: &[u8], shape_id: u32, depth: usize) -> Result<Option<Option<&[u8]>>> {
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("OfficeArt nesting exceeds the run limit".into()).into(),
        );
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER && shape_id_of(record)? == Some(shape_id) {
        let mut textbox = None;
        for child in children(record)? {
            let child = child?;
            if record_type(child)? == OFFICEART_CLIENT_TEXTBOX && textbox.replace(child).is_some() {
                return Err(Error::Refused(Refusal::AmbiguousTextbox));
            }
        }
        return Ok(Some(textbox));
    }
    if record_version(record)? != 0x0F {
        return Ok(None);
    }
    for child in children(record)? {
        if let Some(found) = textbox_in(child?, shape_id, depth + 1)? {
            return Ok(Some(found));
        }
    }
    Ok(None)
}

/// The outline body a textbox's `OutlineTextRefAtom` selects.
fn outline_index(textbox: &[u8]) -> Result<Option<usize>> {
    let mut found = None;
    for child in children(textbox)? {
        let child = child?;
        let kind = record_type(child)?;
        if kind == RecordType::OutlineTextRefAtom as u16 {
            let payload: [u8; 4] = drawing_payload(child)?
                .try_into()
                .map_err(|_err| PackageError::Corrupted("malformed OutlineTextRefAtom".into()))?;
            found = usize::try_from(i32::from_le_bytes(payload)).ok();
        } else if kind != RecordType::TextRulerAtom as u16 {
            return Ok(None);
        }
    }
    Ok(found)
}

/// Read the runs of the shape with `shape_id` from a slide's drawing payload
/// and its slide-list bodies; `None` for a shape without a text body.
pub(crate) fn read(
    drawing: &[u8],
    outline: &[crate::OutlineTextBody],
    shape_id: u32,
    source_limits: SourceLimits,
) -> Result<Option<TextRuns>> {
    let mut found = None;
    let records = super::ChildRecords {
        payload: drawing,
        offset: 0,
    };
    for record in records {
        if let Some(textbox) = textbox_in(record?, shape_id, 0)? {
            found = Some(textbox);
            break;
        }
    }
    let Some(Some(textbox)) = found else {
        return Ok(None);
    };
    if let Some(index) = outline_index(textbox)? {
        let Some(body) = outline.get(index) else {
            return Ok(None);
        };
        let records = super::ChildRecords {
            payload: body.records(),
            offset: 0,
        }
        .collect::<crate::package::Result<Vec<_>>>()?;
        return parse_body(records, interaction_limits_for_source(source_limits)).map(Some);
    }
    let records = children(textbox)?.collect::<crate::package::Result<Vec<_>>>()?;
    parse_body(records, interaction_limits_for_source(source_limits)).map(Some)
}

/// The text-body rewrite of an element edit: the new body records, or
/// `None` to keep the body.
type BodyEdit<'e> = &'e dyn Fn(&[&[u8]]) -> Result<Option<Vec<u8>>>;

/// Apply `edit` to the text body of the shape with `shape_id` on the slide
/// record `slide` whose persist identifier is `slide_persist_id`; the live
/// document `document` holds outline bodies. Returns the new slide and
/// document records, `None` for an unchanged one.
pub(crate) fn rewrite(
    slide: &[u8],
    document: &[u8],
    slide_persist_id: u32,
    shape_id: u32,
    edit: BodyEdit<'_>,
) -> Result<(Option<Vec<u8>>, Option<Vec<u8>>)> {
    if let Some(body) = outline_reference(slide, shape_id)? {
        return rewrite_outline(document, slide_persist_id, body, edit)
            .map(|document| (None, document));
    }
    let drawing = find_ppt_record(slide, RecordType::PPDrawing as u16, 0)?
        .ok_or(Error::Refused(Refusal::ShapeNotFound))?;
    let mut matches = 0usize;
    let mut changed = false;
    let data = rewrite_officeart(
        drawing_payload(drawing)?,
        shape_id,
        edit,
        0,
        &mut matches,
        &mut changed,
    )?;
    match matches {
        0 => return Err(Error::Refused(Refusal::ShapeNotFound)),
        1 => {},
        _ => return Err(Error::Refused(Refusal::AmbiguousShape)),
    }
    if !changed {
        return Ok((None, None));
    }
    let drawing_record = rebuild(drawing, &data)?;
    let slide = rewrite_ppt_record(slide, 0, &mut |record| {
        if std::ptr::eq(record.as_ptr(), drawing.as_ptr()) && record.len() == drawing.len() {
            return Ok(Some(drawing_record.clone()));
        }
        Ok(None)
    })?;
    Ok((slide, None))
}

fn rewrite_officeart(
    record: &[u8],
    shape_id: u32,
    edit: BodyEdit<'_>,
    depth: usize,
    matches: &mut usize,
    changed: &mut bool,
) -> Result<Vec<u8>> {
    if depth >= 128 {
        return Err(
            PackageError::Corrupted("OfficeArt nesting exceeds the run limit".into()).into(),
        );
    }
    if record_type(record)? == OFFICEART_SP_CONTAINER && shape_id_of(record)? == Some(shape_id) {
        *matches += 1;
        let mut data = Vec::with_capacity(record.len() + 64);
        let mut textboxes = 0usize;
        for child in children(record)? {
            let child = child?;
            if record_type(child)? != OFFICEART_CLIENT_TEXTBOX {
                data.extend_from_slice(child);
                continue;
            }
            textboxes += 1;
            let records = children(child)?.collect::<crate::package::Result<Vec<_>>>()?;
            match edit(&records)? {
                Some(body) => {
                    *changed = true;
                    data.extend_from_slice(&rebuild(child, &body)?);
                },
                None => data.extend_from_slice(child),
            }
        }
        return match textboxes {
            0 => Err(Error::Refused(Refusal::NoTextbox)),
            1 => rebuild(record, &data).map_err(Into::into),
            _ => Err(Error::Refused(Refusal::AmbiguousTextbox)),
        };
    }
    if record_version(record)? != 0x0F {
        return Ok(record.to_vec());
    }
    let mut data = Vec::with_capacity(drawing_payload(record)?.len());
    for child in children(record)? {
        data.extend_from_slice(&rewrite_officeart(
            child?,
            shape_id,
            edit,
            depth + 1,
            matches,
            changed,
        )?);
    }
    rebuild(record, &data).map_err(Into::into)
}

/// Apply `edit` to one slide-list body of the live document.
fn rewrite_outline(
    document: &[u8],
    slide_persist_id: u32,
    body: usize,
    edit: BodyEdit<'_>,
) -> Result<Option<Vec<u8>>> {
    let slide_list = slide_list_of(document)?;
    let mut list = Vec::with_capacity(drawing_payload(slide_list)?.len() + 64);
    let mut selected = 0usize;
    let mut changed = false;
    for_each_outline_body(slide_list, |key, records| {
        if key == Some((slide_persist_id, body)) {
            selected += 1;
            if let Some(rewritten) = edit(records)? {
                changed = true;
                list.extend_from_slice(&rewritten);
                return Ok(());
            }
        }
        for record in records {
            list.extend_from_slice(record);
        }
        Ok(())
    })?;
    match selected {
        0 => return Err(Error::Refused(Refusal::NoTextAtom)),
        1 => {},
        _ => return Err(Error::Refused(Refusal::AmbiguousTextbox)),
    }
    if !changed {
        return Ok(None);
    }
    let list = rebuild(slide_list, &list)?;
    let mut data = Vec::with_capacity(drawing_payload(document)?.len() + 64);
    for child in children(document)? {
        let child = child?;
        if std::ptr::eq(child.as_ptr(), slide_list.as_ptr()) {
            data.extend_from_slice(&list);
        } else {
            data.extend_from_slice(child);
        }
    }
    rebuild(document, &data).map(Some).map_err(Into::into)
}

/// The UTF-16 length of a body's text.
fn text_units(records: &[&[u8]]) -> Result<usize> {
    for record in records {
        let kind = record_type(record)?;
        if kind == RecordType::TextCharsAtom as u16 {
            return super::text_units(TextKind::Chars, drawing_payload(record)?.len());
        }
        if kind == RecordType::TextBytesAtom as u16 {
            return Ok(drawing_payload(record)?.len());
        }
    }
    Err(Error::Refused(Refusal::NoTextAtom))
}

/// The body with `format` applied to the units `start..end`; `None` when
/// every unit already shows it.
pub(crate) fn format_body(
    records: &[&[u8]],
    start: usize,
    end: usize,
    format: &RunFormat,
    interaction_limits: InteractionLimits,
) -> Result<Option<Vec<u8>>> {
    let units = text_units(records)?;
    let mut style = None;
    let mut text_index = None;
    for (index, record) in records.iter().enumerate() {
        let kind = record_type(record)?;
        if kind == RecordType::StyleTextPropAtom as u16 && style.replace(index).is_some() {
            return Err(Error::Refused(Refusal::DependencyClosure));
        }
        if kind == RecordType::TextCharsAtom as u16 || kind == RecordType::TextBytesAtom as u16 {
            text_index = Some(index);
        }
    }
    let text_index = text_index.ok_or(Error::Refused(Refusal::NoTextAtom))?;
    let rewritten = crate::StyleTextPropAtom::rewrite_character_format(
        style.map(|index| records[index]),
        units,
        start,
        end,
        |current| {
            if format.holds(current) {
                current.clone()
            } else {
                format.apply(current)
            }
        },
    )?;
    let Some(rewritten) = rewritten else {
        if style.is_some() {
            // The run already shows the format, or the atom is not
            // reproducible byte for byte.
            let current = parse_body(records.iter().copied(), interaction_limits)?;
            if current
                .runs
                .iter()
                .filter(|run| run.start < end && run.start + run.length > start)
                .all(|run| format.holds(&run.format))
            {
                return Ok(None);
            }
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        return Ok(None);
    };
    let mut data =
        Vec::with_capacity(records.iter().map(|record| record.len()).sum::<usize>() + 32);
    for (index, record) in records.iter().enumerate() {
        if Some(index) == style {
            data.extend_from_slice(&rewritten);
            continue;
        }
        data.extend_from_slice(record);
        if style.is_none() && index == text_index {
            // [MS-PPT] 2.9.76: the style atom follows the text atom.
            data.extend_from_slice(&rewritten);
        }
    }
    Ok(Some(data))
}

/// Derive the interaction family profile from the selected package source
/// profile.  Text-run readback is source-backed, so it must not silently
/// reopen an `InteractiveInfo` child with an unrelated default budget.
pub(crate) fn interaction_limits_for_source(source_limits: SourceLimits) -> InteractionLimits {
    InteractionLimits {
        max_record_bytes: source_limits.max_record_bytes(),
        max_macro_name_bytes: source_limits
            .max_record_payload_bytes()
            .min(InteractionLimits::default().max_macro_name_bytes),
        max_events: source_limits.max_events(),
        max_depth: source_limits.max_depth(),
        max_record_payload_bytes: source_limits.max_record_payload_bytes(),
        max_aggregate_payload_bytes: source_limits.max_aggregate_payload_bytes(),
    }
}

/// The body with its click hyperlink over exactly `start..end` replaced by
/// `click` (an `InteractiveInfo` container) or removed; `None` when nothing
/// changes. A click over an overlapping, different range refuses.
pub(crate) fn link_body(
    records: &[&[u8]],
    start: usize,
    end: usize,
    click: Option<&[u8]>,
) -> Result<Option<Vec<u8>>> {
    let units = text_units(records)?;
    if start >= end || end > units + 1 {
        return Err(Error::Refused(Refusal::DependencyClosure));
    }
    let mut data =
        Vec::with_capacity(records.iter().map(|record| record.len()).sum::<usize>() + 48);
    let mut changed = false;
    let mut index = 0;
    while index < records.len() {
        let record = records[index];
        let is_click_pair = record_type(record)? == RecordType::InteractiveInfo as u16
            && record_instance(record)? == 0
            && records.get(index + 1).is_some_and(|next| {
                record_type(next).ok() == Some(RecordType::TextInteractiveInfoAtom as u16)
                    && record_instance(next).ok() == Some(0)
            });
        if is_click_pair {
            let atom = drawing_payload(records[index + 1])?;
            let range = |offset: usize| -> usize {
                atom.get(offset..offset + 4)
                    .and_then(|bytes| bytes.try_into().ok())
                    .and_then(|bytes: [u8; 4]| usize::try_from(i32::from_le_bytes(bytes)).ok())
                    .unwrap_or(usize::MAX)
            };
            let (pair_start, pair_end) = (range(0), range(4));
            if (pair_start, pair_end) == (start, end) {
                changed = true;
                index += 2;
                continue;
            }
            if pair_start < end && pair_end > start {
                return Err(Error::Refused(Refusal::DependencyClosure));
            }
        }
        data.extend_from_slice(record);
        index += 1;
    }
    if let Some(click) = click {
        let mut atom = Vec::with_capacity(16);
        atom.extend_from_slice(&0u16.to_le_bytes());
        atom.extend_from_slice(&(RecordType::TextInteractiveInfoAtom as u16).to_le_bytes());
        atom.extend_from_slice(&8u32.to_le_bytes());
        let to_i32 = |value: usize| {
            i32::try_from(value).map_err(|_err| Error::Refused(Refusal::DependencyClosure))
        };
        atom.extend_from_slice(&to_i32(start)?.to_le_bytes());
        atom.extend_from_slice(&to_i32(end)?.to_le_bytes());
        data.extend_from_slice(click);
        data.extend_from_slice(&atom);
        changed = true;
    }
    Ok(changed.then_some(data))
}
