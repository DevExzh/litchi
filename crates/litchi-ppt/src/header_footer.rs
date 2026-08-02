//! Strict, inert metadata for binary PowerPoint headers and footers.
//!
//! This module implements [MS-PPT] sections 2.4.15 and 2.5.16. It parses and
//! serializes only the relevant record family; it does not format dates, modify
//! an OLE compound file, or activate presentation content.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

const HEADERS_FOOTERS_RECORD_TYPE: u16 = 0x0FD9;
const HEADERS_FOOTERS_ATOM_RECORD_TYPE: u16 = 0x0FDA;
const CSTRING_RECORD_TYPE: u16 = 0x0FBA;
const CONTAINER_VERSION: u16 = 0x000F;
const ATOM_VERSION: u16 = 0;
const PRESENTATION_SLIDES_INSTANCE: u16 = 3;
const NOTES_AND_HANDOUTS_INSTANCE: u16 = 4;
const LOCAL_INSTANCE: u16 = 0;
const USER_DATE_INSTANCE: u16 = 0;
const HEADER_INSTANCE: u16 = 1;
const FOOTER_INSTANCE: u16 = 2;
const HEADERS_FOOTERS_ATOM_LENGTH: usize = 4;
const USER_DATE_MAX_BYTES: usize = 510;
const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
const MAX_AGGREGATE_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_HEADER_FOOTER_ENTRIES: usize = 65_536;
const MAX_SCANNED_RECORDS: usize = 1_000_000;
const KNOWN_FLAG_MASK: u16 = 0x003F;

/// A validated PowerPoint datetime format identifier.
///
/// Values 0 through 12 are the ordinary locale-dependent formats. Value 13
/// is permitted by `HeadersFootersAtom`, although producers are advised not to
/// emit it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PowerPointDateTimeFormatId(u8);

impl PowerPointDateTimeFormatId {
    /// Lowest valid format identifier.
    pub const MIN: u8 = 0;
    /// Highest valid format identifier.
    pub const MAX: u8 = 13;

    /// Construct a validated format identifier.
    pub fn new(value: u8) -> Result<Self> {
        if value > Self::MAX {
            return Err(corrupted(
                "header/footer datetime format ID is outside 0..=13",
            ));
        }
        Ok(Self(value))
    }

    /// Return the on-disk identifier.
    #[inline]
    pub const fn get(self) -> u8 {
        self.0
    }
}

impl Default for PowerPointDateTimeFormatId {
    fn default() -> Self {
        Self(Self::MIN)
    }
}

impl TryFrom<u8> for PowerPointDateTimeFormatId {
    type Error = PptError;

    fn try_from(value: u8) -> Result<Self> {
        Self::new(value)
    }
}

impl From<PowerPointDateTimeFormatId> for u8 {
    fn from(value: PowerPointDateTimeFormatId) -> Self {
        value.get()
    }
}

/// The direct parent of a local header/footer container.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PowerPointHeaderFooterParent {
    /// A presentation slide.
    Slide,
    /// A main-master slide.
    MainMaster,
}

/// A zero-based ordinal among parents of the same kind in record order.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct PowerPointHeaderFooterParentOrdinal(usize);

impl PowerPointHeaderFooterParentOrdinal {
    /// Construct an ordinal from a zero-based parent index.
    #[inline]
    pub const fn new(value: usize) -> Self {
        Self(value)
    }

    /// Return the zero-based ordinal.
    #[inline]
    pub const fn get(self) -> usize {
        self.0
    }
}

/// The specification-defined scope of a header/footer container.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PowerPointHeaderFooterScope {
    /// Presentation-wide defaults for ordinary slides.
    PresentationSlides,
    /// Presentation-wide defaults for notes pages and handouts.
    NotesAndHandouts,
    /// Overrides or defaults attached directly to one slide or main master.
    Local {
        /// Kind of direct parent.
        parent: PowerPointHeaderFooterParent,
        /// Parent ordinal in PowerPoint record order.
        parent_ordinal: PowerPointHeaderFooterParentOrdinal,
    },
}

impl PowerPointHeaderFooterScope {
    fn record_instance(self) -> u16 {
        match self {
            Self::PresentationSlides => PRESENTATION_SLIDES_INSTANCE,
            Self::NotesAndHandouts => NOTES_AND_HANDOUTS_INSTANCE,
            Self::Local { .. } => LOCAL_INSTANCE,
        }
    }

    fn permits_header_atom(self) -> bool {
        matches!(self, Self::NotesAndHandouts)
    }
}

/// Display options stored by `HeadersFootersAtom`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PowerPointHeaderFooterOptions {
    /// Locale-dependent datetime format identifier.
    pub datetime_format: PowerPointDateTimeFormatId,
    /// Display a date placeholder.
    pub show_date: bool,
    /// Use the current date and time.
    pub use_current_datetime: bool,
    /// Use the custom user-date string.
    pub use_user_date: bool,
    /// Display the slide number.
    pub show_slide_number: bool,
    /// Display a header. This bit is retained even where the specification says
    /// it has no effect.
    pub show_header: bool,
    /// Display a footer.
    pub show_footer: bool,
}

/// Text derived from inert header/footer placeholder shapes.
///
/// Office 2007 can save binary presentations with visible header/footer text
/// in placeholders while leaving the corresponding CString atoms absent. This
/// view is kept separate so record-local serialization remains lossless.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointHeaderFooterDisplayText {
    /// Custom date text visible through a datetime placeholder.
    pub user_date: Option<String>,
    /// Header text visible through a header placeholder.
    pub header: Option<String>,
    /// Footer text visible through a footer placeholder.
    pub footer: Option<String>,
}

/// Placeholder-derived display text associated with a specification scope.
///
/// A scoped display can exist without a corresponding local record because
/// Office 2007 binary presentations can inherit document-level options while
/// storing slide-specific text only in placeholder shapes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointScopedHeaderFooterDisplayText {
    /// Slide, master, or document-level scope of the placeholder text.
    pub scope: PowerPointHeaderFooterScope,
    /// Inert text extracted from placeholder shapes.
    pub text: PowerPointHeaderFooterDisplayText,
}

impl PowerPointHeaderFooterOptions {
    fn from_atom(record: &PptRecord) -> Result<Self> {
        validate_record_header(
            record,
            PptRecordType::HeadersFootersAtom,
            HEADERS_FOOTERS_ATOM_RECORD_TYPE,
            ATOM_VERSION,
            0,
        )?;
        if record.data_length as usize != HEADERS_FOOTERS_ATOM_LENGTH
            || record.data.len() != HEADERS_FOOTERS_ATOM_LENGTH
            || !record.children.is_empty()
        {
            return Err(corrupted(
                "HeadersFootersAtom must have exactly four data bytes",
            ));
        }
        let format_id = i16::from_le_bytes([record.data[0], record.data[1]]);
        if !(0..=i16::from(PowerPointDateTimeFormatId::MAX)).contains(&format_id) {
            return Err(corrupted(
                "header/footer datetime format ID is outside 0..=13",
            ));
        }
        let mask = u16::from_le_bytes([record.data[2], record.data[3]]);
        if mask & !KNOWN_FLAG_MASK != 0 {
            return Err(corrupted(
                "HeadersFootersAtom has nonzero reserved flag bits",
            ));
        }
        Ok(Self {
            datetime_format: PowerPointDateTimeFormatId(format_id as u8),
            show_date: mask & 0x0001 != 0,
            use_current_datetime: mask & 0x0002 != 0,
            use_user_date: mask & 0x0004 != 0,
            show_slide_number: mask & 0x0008 != 0,
            show_header: mask & 0x0010 != 0,
            show_footer: mask & 0x0020 != 0,
        })
    }

    fn mask(self) -> u16 {
        u16::from(self.show_date)
            | (u16::from(self.use_current_datetime) << 1)
            | (u16::from(self.use_user_date) << 2)
            | (u16::from(self.show_slide_number) << 3)
            | (u16::from(self.show_header) << 4)
            | (u16::from(self.show_footer) << 5)
    }
}

/// Typed, inert metadata from one PowerPoint header/footer container.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointHeaderFooter {
    /// Container scope and parent association.
    pub scope: PowerPointHeaderFooterScope,
    /// Display and datetime-format options.
    pub options: PowerPointHeaderFooterOptions,
    /// Optional custom date text.
    pub user_date: Option<String>,
    /// Optional notes/handout header text.
    pub header: Option<String>,
    /// Optional footer text.
    pub footer: Option<String>,
    /// Optional text derived from inert placeholders. This is never serialized
    /// into the record-local CString fields.
    pub placeholder_display: Option<PowerPointHeaderFooterDisplayText>,
}

impl PowerPointHeaderFooter {
    /// Strictly parse one already-materialized `RT_HeadersFooters` record.
    ///
    /// The supplied scope is checked against the record instance. Direct-parent
    /// placement is validated by [`PowerPointHeaderFooters`] when parsing a
    /// complete presentation.
    pub fn parse_record(record: &PptRecord, scope: PowerPointHeaderFooterScope) -> Result<Self> {
        let mut aggregate = 0usize;
        Self::parse_record_bounded(record, scope, &mut aggregate)
    }

    fn parse_record_bounded(
        record: &PptRecord,
        scope: PowerPointHeaderFooterScope,
        aggregate: &mut usize,
    ) -> Result<Self> {
        validate_record_header(
            record,
            PptRecordType::HeadersFooters,
            HEADERS_FOOTERS_RECORD_TYPE,
            CONTAINER_VERSION,
            scope.record_instance(),
        )?;
        if record.data_length as usize != record.data.len() {
            return Err(corrupted("HeadersFooters container payload is truncated"));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "HeadersFooters")?;
        let Some(atom) = children.first() else {
            return Err(corrupted(
                "HeadersFooters container is missing HeadersFootersAtom",
            ));
        };
        let options = PowerPointHeaderFooterOptions::from_atom(atom)?;

        let mut user_date = None;
        let mut header = None;
        let mut footer = None;
        let mut previous_instance = None;
        for child in &children[1..] {
            if child.record_type != PptRecordType::CString
                || child.record_type_raw != CSTRING_RECORD_TYPE
                || child.version != ATOM_VERSION
            {
                return Err(corrupted(
                    "HeadersFooters contains an unexpected child record",
                ));
            }
            if child.data_length as usize != child.data.len() || child.data.len() % 2 != 0 {
                return Err(corrupted(
                    "header/footer CString has an invalid byte length",
                ));
            }
            if child.data.len() > MAX_TEXT_BYTES {
                return Err(corrupted(
                    "header/footer CString exceeds the resource limit",
                ));
            }
            if previous_instance.is_some_and(|previous| child.instance <= previous) {
                return Err(corrupted(
                    "header/footer CString children are duplicated or out of order",
                ));
            }
            previous_instance = Some(child.instance);
            *aggregate = aggregate
                .checked_add(child.data.len())
                .ok_or_else(|| corrupted("header/footer aggregate size overflow"))?;
            if *aggregate > MAX_AGGREGATE_TEXT_BYTES {
                return Err(corrupted(
                    "header/footer strings exceed the aggregate resource limit",
                ));
            }
            let value = decode_printable_unicode(&child.data)?;
            match child.instance {
                USER_DATE_INSTANCE => {
                    if child.data.len() > USER_DATE_MAX_BYTES {
                        return Err(corrupted("UserDateAtom exceeds 510 bytes"));
                    }
                    user_date = Some(value);
                },
                HEADER_INSTANCE if scope.permits_header_atom() => header = Some(value),
                HEADER_INSTANCE => {
                    return Err(corrupted(
                        "HeaderAtom is not permitted in this header/footer scope",
                    ));
                },
                FOOTER_INSTANCE => footer = Some(value),
                _ => return Err(corrupted("header/footer CString has an invalid instance")),
            }
        }

        Ok(Self {
            scope,
            options,
            user_date,
            header,
            footer,
            placeholder_display: None,
        })
    }

    /// Return visible custom-date text, preferring an attached Office 2007
    /// placeholder and otherwise using the stored UserDateAtom.
    pub fn display_user_date(&self) -> Option<&str> {
        self.placeholder_display
            .as_ref()
            .and_then(|display| display.user_date.as_deref())
            .or(self.user_date.as_deref())
    }

    /// Return visible header text, preferring an attached Office 2007
    /// placeholder and otherwise using the stored HeaderAtom.
    pub fn display_header(&self) -> Option<&str> {
        self.placeholder_display
            .as_ref()
            .and_then(|display| display.header.as_deref())
            .or(self.header.as_deref())
    }

    /// Return visible footer text, preferring an attached Office 2007
    /// placeholder and otherwise using the stored FooterAtom.
    pub fn display_footer(&self) -> Option<&str> {
        self.placeholder_display
            .as_ref()
            .and_then(|display| display.footer.as_deref())
            .or(self.footer.as_deref())
    }

    /// Serialize this metadata as one canonical `RT_HeadersFooters` record.
    ///
    /// Serialization is record-local and deterministic. It does not evaluate a
    /// date or modify an OLE persistence directory.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.validate_for_write()?;
        let mut body = Vec::new();

        let atom_data = [
            self.options.datetime_format.get(),
            0,
            self.options.mask() as u8,
            (self.options.mask() >> 8) as u8,
        ];
        append_record(
            &mut body,
            ATOM_VERSION,
            0,
            HEADERS_FOOTERS_ATOM_RECORD_TYPE,
            &atom_data,
        )?;
        if let Some(value) = &self.user_date {
            append_cstring(&mut body, USER_DATE_INSTANCE, value)?;
        }
        if let Some(value) = &self.header {
            append_cstring(&mut body, HEADER_INSTANCE, value)?;
        }
        if let Some(value) = &self.footer {
            append_cstring(&mut body, FOOTER_INSTANCE, value)?;
        }

        let mut output = Vec::with_capacity(body.len().saturating_add(8));
        append_record(
            &mut output,
            CONTAINER_VERSION,
            self.scope.record_instance(),
            HEADERS_FOOTERS_RECORD_TYPE,
            &body,
        )?;
        Ok(output)
    }

    fn validate_for_write(&self) -> Result<()> {
        if self.header.is_some() && !self.scope.permits_header_atom() {
            return Err(corrupted(
                "HeaderAtom is not permitted in this header/footer scope",
            ));
        }
        let mut aggregate = 0usize;
        for (kind, value) in [
            (USER_DATE_INSTANCE, self.user_date.as_deref()),
            (HEADER_INSTANCE, self.header.as_deref()),
            (FOOTER_INSTANCE, self.footer.as_deref()),
        ] {
            let Some(value) = value else { continue };
            let bytes = validated_encoded_len(value)?;
            if kind == USER_DATE_INSTANCE && bytes > USER_DATE_MAX_BYTES {
                return Err(corrupted("UserDateAtom exceeds 510 bytes"));
            }
            aggregate = aggregate
                .checked_add(bytes)
                .ok_or_else(|| corrupted("header/footer aggregate size overflow"))?;
        }
        if aggregate > MAX_AGGREGATE_TEXT_BYTES {
            return Err(corrupted(
                "header/footer strings exceed the aggregate resource limit",
            ));
        }
        Ok(())
    }
}

/// All strictly located header/footer containers in a presentation.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointHeaderFooters {
    entries: Vec<PowerPointHeaderFooter>,
    placeholder_displays: Vec<PowerPointScopedHeaderFooterDisplayText>,
    placeholder_display_bytes: usize,
}

impl PowerPointHeaderFooters {
    /// Return entries in PowerPoint record order.
    #[inline]
    pub fn entries(&self) -> &[PowerPointHeaderFooter] {
        &self.entries
    }

    /// Return placeholder-derived displays in physical PowerPoint record order.
    ///
    /// Unlike [`Self::entries`], these values are not necessarily backed by a
    /// local `RT_HeadersFooters` record and cannot be serialized as one.
    #[inline]
    pub fn placeholder_displays(&self) -> &[PowerPointScopedHeaderFooterDisplayText] {
        &self.placeholder_displays
    }

    /// Return placeholder-derived display text for an exact scope.
    pub fn placeholder_display(
        &self,
        scope: PowerPointHeaderFooterScope,
    ) -> Option<&PowerPointHeaderFooterDisplayText> {
        self.placeholder_displays
            .iter()
            .find(|display| display.scope == scope)
            .map(|display| &display.text)
    }

    /// Return the presentation-wide ordinary-slide defaults, if present.
    pub fn presentation_slides(&self) -> Option<&PowerPointHeaderFooter> {
        self.entries
            .iter()
            .find(|entry| entry.scope == PowerPointHeaderFooterScope::PresentationSlides)
    }

    /// Return the presentation-wide notes/handout defaults, if present.
    pub fn notes_and_handouts(&self) -> Option<&PowerPointHeaderFooter> {
        self.entries
            .iter()
            .find(|entry| entry.scope == PowerPointHeaderFooterScope::NotesAndHandouts)
    }

    pub(crate) fn parse_record_tree(records: &[&PptRecord]) -> Result<Self> {
        if records.len() > MAX_SCANNED_RECORDS {
            return Err(corrupted(
                "PowerPoint record tree exceeds the header/footer scan limit",
            ));
        }
        let document_count = records
            .iter()
            .filter(|record| record.record_type == PptRecordType::Document)
            .count();
        if document_count != 1 {
            return Err(corrupted(
                "PowerPoint must contain exactly one Document container",
            ));
        }

        let total_containers = records
            .iter()
            .filter(|record| record.record_type == PptRecordType::HeadersFooters)
            .count();
        if total_containers > MAX_HEADER_FOOTER_ENTRIES {
            return Err(corrupted("too many PowerPoint header/footer containers"));
        }

        let mut entries = Vec::with_capacity(total_containers);
        let mut located = 0usize;
        let mut slide_ordinal = 0usize;
        let mut master_ordinal = 0usize;
        let mut aggregate = 0usize;

        for parent in records {
            match parent.record_type {
                PptRecordType::Document => {
                    let mut saw_slides = false;
                    let mut saw_notes = false;
                    for child in parent
                        .children
                        .iter()
                        .filter(|child| child.record_type == PptRecordType::HeadersFooters)
                    {
                        located += 1;
                        let scope = match child.instance {
                            PRESENTATION_SLIDES_INSTANCE if !saw_slides => {
                                saw_slides = true;
                                PowerPointHeaderFooterScope::PresentationSlides
                            },
                            NOTES_AND_HANDOUTS_INSTANCE if !saw_notes => {
                                saw_notes = true;
                                PowerPointHeaderFooterScope::NotesAndHandouts
                            },
                            PRESENTATION_SLIDES_INSTANCE | NOTES_AND_HANDOUTS_INSTANCE => {
                                return Err(corrupted(
                                    "duplicate document-level header/footer container",
                                ));
                            },
                            _ => {
                                return Err(corrupted(
                                    "invalid document-level header/footer instance",
                                ));
                            },
                        };
                        entries.push(PowerPointHeaderFooter::parse_record_bounded(
                            child,
                            scope,
                            &mut aggregate,
                        )?);
                    }
                },
                PptRecordType::Slide => {
                    locate_local(
                        parent,
                        PowerPointHeaderFooterParent::Slide,
                        slide_ordinal,
                        &mut located,
                        &mut aggregate,
                        &mut entries,
                    )?;
                    slide_ordinal += 1;
                },
                PptRecordType::MainMaster => {
                    locate_local(
                        parent,
                        PowerPointHeaderFooterParent::MainMaster,
                        master_ordinal,
                        &mut located,
                        &mut aggregate,
                        &mut entries,
                    )?;
                    master_ordinal += 1;
                },
                _ => {},
            }
        }
        if located != total_containers {
            return Err(corrupted(
                "HeadersFooters container has an invalid direct parent",
            ));
        }
        Ok(Self {
            entries,
            placeholder_displays: Vec::new(),
            placeholder_display_bytes: 0,
        })
    }

    pub(crate) fn attach_placeholder_display(
        &mut self,
        scope: PowerPointHeaderFooterScope,
        display: PowerPointHeaderFooterDisplayText,
    ) -> Result<()> {
        if self.placeholder_displays.len() == MAX_HEADER_FOOTER_ENTRIES {
            return Err(corrupted("too many PowerPoint placeholder displays"));
        }
        if self
            .placeholder_displays
            .iter()
            .any(|existing| existing.scope == scope)
        {
            return Err(corrupted("duplicate PowerPoint placeholder display scope"));
        }
        let mut display_bytes = 0usize;
        for value in [
            display.user_date.as_deref(),
            display.header.as_deref(),
            display.footer.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            display_bytes = display_bytes
                .checked_add(validated_encoded_len(value)?)
                .ok_or_else(|| corrupted("placeholder display aggregate size overflow"))?;
        }
        self.placeholder_display_bytes = self
            .placeholder_display_bytes
            .checked_add(display_bytes)
            .ok_or_else(|| corrupted("placeholder display aggregate size overflow"))?;
        if self.placeholder_display_bytes > MAX_AGGREGATE_TEXT_BYTES {
            return Err(corrupted(
                "placeholder display strings exceed the aggregate resource limit",
            ));
        }
        if let Some(entry) = self.entries.iter_mut().find(|entry| entry.scope == scope) {
            entry.placeholder_display = Some(display.clone());
        }
        self.placeholder_displays
            .push(PowerPointScopedHeaderFooterDisplayText {
                scope,
                text: display,
            });
        Ok(())
    }

    pub(crate) fn has_scope(&self, scope: PowerPointHeaderFooterScope) -> bool {
        self.entries.iter().any(|entry| entry.scope == scope)
    }
}

fn locate_local(
    parent_record: &PptRecord,
    parent: PowerPointHeaderFooterParent,
    ordinal: usize,
    located: &mut usize,
    aggregate: &mut usize,
    entries: &mut Vec<PowerPointHeaderFooter>,
) -> Result<()> {
    let mut containers = parent_record
        .children
        .iter()
        .filter(|child| child.record_type == PptRecordType::HeadersFooters);
    let Some(container) = containers.next() else {
        return Ok(());
    };
    if containers.next().is_some() {
        return Err(corrupted(
            "slide or main master has duplicate header/footer containers",
        ));
    }
    if container.instance != LOCAL_INSTANCE {
        return Err(corrupted(
            "local header/footer container has a nonzero instance",
        ));
    }
    *located += 1;
    let scope = PowerPointHeaderFooterScope::Local {
        parent,
        parent_ordinal: PowerPointHeaderFooterParentOrdinal(ordinal),
    };
    entries.push(PowerPointHeaderFooter::parse_record_bounded(
        container, scope, aggregate,
    )?);
    Ok(())
}

fn validate_record_header(
    record: &PptRecord,
    expected_type: PptRecordType,
    expected_raw_type: u16,
    expected_version: u16,
    expected_instance: u16,
) -> Result<()> {
    if record.record_type != expected_type
        || record.record_type_raw != expected_raw_type
        || record.version != expected_version
        || record.instance != expected_instance
    {
        return Err(corrupted(
            "header/footer record header does not match [MS-PPT]",
        ));
    }
    Ok(())
}

fn decode_printable_unicode(data: &[u8]) -> Result<String> {
    let mut units = Vec::with_capacity(data.len() / 2);
    let mut terminated = false;
    for bytes in data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if terminated {
            if unit != 0 {
                return Err(corrupted(
                    "PrintableUnicodeString has data after its terminator",
                ));
            }
            continue;
        }
        if unit == 0 {
            terminated = true;
            continue;
        }
        if is_forbidden_printable_unit(unit) {
            return Err(corrupted(
                "PrintableUnicodeString contains a forbidden control character",
            ));
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_| corrupted("PrintableUnicodeString contains invalid UTF-16"))
}

fn is_forbidden_printable_unit(unit: u16) -> bool {
    matches!(unit, 0x0000..=0x001F | 0x007F..=0x009F)
}

fn validated_encoded_len(value: &str) -> Result<usize> {
    let mut units = 0usize;
    for unit in value.encode_utf16() {
        if is_forbidden_printable_unit(unit) {
            return Err(corrupted(
                "PrintableUnicodeString contains a forbidden control character",
            ));
        }
        units = units
            .checked_add(1)
            .ok_or_else(|| corrupted("header/footer string length overflow"))?;
    }
    let bytes = units
        .checked_mul(2)
        .ok_or_else(|| corrupted("header/footer string length overflow"))?;
    if bytes > MAX_TEXT_BYTES {
        return Err(corrupted(
            "header/footer CString exceeds the resource limit",
        ));
    }
    Ok(bytes)
}

fn append_cstring(output: &mut Vec<u8>, instance: u16, value: &str) -> Result<()> {
    let encoded_len = validated_encoded_len(value)?;
    let mut data = Vec::with_capacity(encoded_len);
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    append_record(output, ATOM_VERSION, instance, CSTRING_RECORD_TYPE, &data)
}

fn append_record(
    output: &mut Vec<u8>,
    version: u16,
    instance: u16,
    record_type: u16,
    data: &[u8],
) -> Result<()> {
    if version > 0x000F || instance > 0x0FFF {
        return Err(corrupted("PowerPoint record header field overflow"));
    }
    let length = u32::try_from(data.len())
        .map_err(|_| corrupted("PowerPoint record payload exceeds u32"))?;
    output
        .try_reserve(8usize.saturating_add(data.len()))
        .map_err(|_| corrupted("unable to reserve header/footer record memory"))?;
    let version_instance = version | (instance << 4);
    output.extend_from_slice(&version_instance.to_le_bytes());
    output.extend_from_slice(&record_type.to_le_bytes());
    output.extend_from_slice(&length.to_le_bytes());
    output.extend_from_slice(data);
    Ok(())
}

fn corrupted(message: impl Into<String>) -> PptError {
    PptError::Corrupted(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const POI_SLIDE_BYTES: &[u8] = &[
        0x3F, 0x00, 0xD9, 0x0F, 0x2E, 0, 0, 0, 0, 0, 0xDA, 0x0F, 4, 0, 0, 0, 0, 0, 0x23, 0, 0x20,
        0, 0xBA, 0x0F, 0x1A, 0, 0, 0, 0x4D, 0, 0x79, 0, 0x20, 0, 0x46, 0, 0x6F, 0, 0x6F, 0, 0x74,
        0, 0x65, 0, 0x72, 0, 0x20, 0, 0x2D, 0, 0x20, 0, 0x31, 0,
    ];
    const POI_NOTES_BYTES: &[u8] = &[
        0x4F, 0, 0xD9, 0x0F, 0x48, 0, 0, 0, 0, 0, 0xDA, 0x0F, 4, 0, 0, 0, 0, 0, 0x3D, 0, 0x10, 0,
        0xBA, 0x0F, 0x16, 0, 0, 0, 0x4E, 0, 0x6F, 0, 0x74, 0, 0x65, 0, 0x20, 0, 0x48, 0, 0x65, 0,
        0x61, 0, 0x64, 0, 0x65, 0, 0x72, 0, 0x20, 0, 0xBA, 0x0F, 0x16, 0, 0, 0, 0x4E, 0, 0x6F, 0,
        0x74, 0, 0x65, 0, 0x20, 0, 0x46, 0, 0x6F, 0, 0x6F, 0, 0x74, 0, 0x65, 0, 0x72, 0,
    ];

    fn parsed(bytes: &[u8], scope: PowerPointHeaderFooterScope) -> PowerPointHeaderFooter {
        let (record, consumed) = PptRecord::parse(bytes, 0).expect("record");
        assert_eq!(consumed, bytes.len());
        PowerPointHeaderFooter::parse_record(&record, scope).expect("header/footer")
    }

    #[test]
    fn poi_record_arrays_are_byte_identical() {
        let slide = parsed(
            POI_SLIDE_BYTES,
            PowerPointHeaderFooterScope::PresentationSlides,
        );
        assert_eq!(slide.footer.as_deref(), Some("My Footer - 1"));
        assert_eq!(slide.to_record_bytes().unwrap(), POI_SLIDE_BYTES);

        let notes = parsed(
            POI_NOTES_BYTES,
            PowerPointHeaderFooterScope::NotesAndHandouts,
        );
        assert_eq!(notes.header.as_deref(), Some("Note Header"));
        assert_eq!(notes.footer.as_deref(), Some("Note Footer"));
        assert_eq!(notes.to_record_bytes().unwrap(), POI_NOTES_BYTES);
    }

    #[test]
    fn all_flags_format_13_empty_and_local_roundtrip() {
        let value = PowerPointHeaderFooter {
            scope: PowerPointHeaderFooterScope::Local {
                parent: PowerPointHeaderFooterParent::Slide,
                parent_ordinal: PowerPointHeaderFooterParentOrdinal(7),
            },
            options: PowerPointHeaderFooterOptions {
                datetime_format: PowerPointDateTimeFormatId::new(13).unwrap(),
                show_date: true,
                use_current_datetime: true,
                use_user_date: true,
                show_slide_number: true,
                show_header: true,
                show_footer: true,
            },
            user_date: Some(String::new()),
            header: None,
            footer: Some(String::new()),
            placeholder_display: None,
        };
        let bytes = value.to_record_bytes().unwrap();
        let reparsed = parsed(&bytes, value.scope);
        assert_eq!(reparsed, value);
    }

    #[test]
    fn malformed_record_matrix_is_rejected() {
        let mut cases = Vec::new();
        for (offset, value) in [
            (0, 0x3Eu8),
            (1, 0x01),
            (2, 0xD8),
            (4, 0xFF),
            (8, 0x01),
            (10, 0xD9),
            (12, 0x05),
            (16, 0x0E),
            (17, 0x80),
            (18, 0x40),
            (24, 0x19),
            (28, 0x01),
        ] {
            let mut bytes = POI_SLIDE_BYTES.to_vec();
            bytes[offset] = value;
            cases.push(bytes);
        }
        let mut invalid_utf16 = POI_SLIDE_BYTES.to_vec();
        invalid_utf16[28] = 0x00;
        invalid_utf16[29] = 0xD8;
        cases.push(invalid_utf16);

        for bytes in cases {
            let rejected = PptRecord::parse(&bytes, 0)
                .and_then(|(record, _)| {
                    PowerPointHeaderFooter::parse_record(
                        &record,
                        PowerPointHeaderFooterScope::PresentationSlides,
                    )
                    .map(|_| (record, 0))
                })
                .is_err();
            assert!(rejected, "malformed bytes were accepted");
        }
    }

    #[test]
    fn illegal_header_controls_and_oversize_user_date_are_rejected() {
        let invalid_header = PowerPointHeaderFooter {
            scope: PowerPointHeaderFooterScope::PresentationSlides,
            options: PowerPointHeaderFooterOptions::default(),
            user_date: None,
            header: Some("not permitted".to_string()),
            footer: None,
            placeholder_display: None,
        };
        assert!(invalid_header.to_record_bytes().is_err());

        let control = PowerPointHeaderFooter {
            scope: PowerPointHeaderFooterScope::NotesAndHandouts,
            options: PowerPointHeaderFooterOptions::default(),
            user_date: None,
            header: None,
            footer: Some("bad\nfooter".to_string()),
            placeholder_display: None,
        };
        assert!(control.to_record_bytes().is_err());

        let user_date = PowerPointHeaderFooter {
            scope: PowerPointHeaderFooterScope::NotesAndHandouts,
            options: PowerPointHeaderFooterOptions::default(),
            user_date: Some("x".repeat(256)),
            header: None,
            footer: None,
            placeholder_display: None,
        };
        assert!(user_date.to_record_bytes().is_err());
    }

    #[test]
    fn placement_duplicate_and_order_violations_are_rejected() {
        let (container, _) = PptRecord::parse(POI_SLIDE_BYTES, 0).unwrap();
        let atom = container.children[0].clone();
        let footer = container.children[1].clone();

        let mut out_of_order = container.clone();
        out_of_order.data.clear();
        out_of_order.children = vec![footer.clone(), atom.clone()];
        out_of_order.data_length = 0;
        assert!(
            PowerPointHeaderFooter::parse_record(
                &out_of_order,
                PowerPointHeaderFooterScope::PresentationSlides,
            )
            .is_err()
        );

        let document = PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: 1000,
            version: 0xF,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: vec![container.clone(), container.clone()],
        };
        let records = vec![&document, &document.children[0], &document.children[1]];
        assert!(PowerPointHeaderFooters::parse_record_tree(&records).is_err());

        let wrong_parent = PptRecord {
            record_type: PptRecordType::Notes,
            record_type_raw: 1008,
            version: 0xF,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: vec![container],
        };
        let empty_document = PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: 1000,
            version: 0xF,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: Vec::new(),
        };
        let records = vec![&empty_document, &wrong_parent, &wrong_parent.children[0]];
        assert!(PowerPointHeaderFooters::parse_record_tree(&records).is_err());
    }
}
