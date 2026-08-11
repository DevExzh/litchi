//! Strict generated-free Buffa projections for Pages section settings.
//!
//! The established pagination decoder still projects only
//! `TP.SectionArchive` fields 20--22. The aggregate decoder independently
//! validates fields 17--22, 26, and 28 before a private lazy view borrows the
//! optional name. Callers retain and rewrite the original payload; this module
//! never owns or re-encodes unrelated section fields.

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_pages_section_generated::LitchiIwaProjection as projection;

/// Finite limits established by the caller's strict wire preflight.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_name_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build an explicit finite profile for one preflighted section payload.
    #[must_use]
    pub const fn new(max_message_bytes: usize, recursion_limit: u32) -> Self {
        Self {
            max_message_bytes,
            max_fields: usize::MAX,
            max_work_bytes: usize::MAX,
            max_name_bytes: max_message_bytes,
            recursion_limit,
        }
    }

    /// Restrict the total number of encoded fields visited by strict routing.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Restrict aggregate strict-routing plus Buffa traversal work.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Restrict bytes in the one optional borrowed section name.
    #[must_use]
    pub const fn with_max_name_bytes(mut self, maximum: usize) -> Self {
        self.max_name_bytes = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Borrow-free scalar result of the private Pages projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct PaginationSnapshot {
    /// Optional native section-start discriminant.
    pub section_start_kind: Option<u32>,
    /// Optional native page-numbering discriminant.
    pub section_page_number_kind: Option<u32>,
    /// Optional native first page number.
    pub section_page_number_start: Option<u32>,
}

/// A finite resource rejected by strict aggregate section routing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// The payload or configured Buffa message ceiling is too large.
    Bytes { observed: usize, maximum: usize },
    /// Encoded field records exceed the finite traversal ceiling.
    Fields { observed: usize, maximum: usize },
    /// Strict plus Buffa traversal work exceeds its finite ceiling.
    Work { observed: usize, maximum: usize },
    /// Configured or traversed protobuf nesting exceeds its finite ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// The optional borrowed section name exceeds its finite byte ceiling.
    NameBytes { observed: usize, maximum: usize },
}

/// Failure from a private Pages section projection decoder.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Pagination(buffa::DecodeError),
    Invalid,
    Limited(DecodeLimit),
}

impl DecodeError {
    /// Return the exact observation for a finite aggregate limit failure.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match &self.kind {
            DecodeErrorKind::Limited(limit) => Some(*limit),
            DecodeErrorKind::Pagination(_) | DecodeErrorKind::Invalid => None,
        }
    }

    const fn invalid() -> Self {
        Self {
            kind: DecodeErrorKind::Invalid,
        }
    }

    const fn limited(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Limited(limit),
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Pagination(error) => error.fmt(formatter),
            DecodeErrorKind::Invalid | DecodeErrorKind::Limited(_) => {
                formatter.write_str("invalid Pages section settings payload")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self {
            kind: DecodeErrorKind::Pagination(error),
        }
    }
}

/// Decode the three pagination scalars from an already-preflighted payload.
///
/// The generated lazy view borrows `source`, retains no repeated-element or
/// unknown-field storage, and is dropped before this borrow-free result is
/// returned.
pub fn decode_pagination(
    source: &[u8],
    options: DecodeOptions,
) -> Result<PaginationSnapshot, DecodeError> {
    let view: projection::PagesSectionPaginationArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    Ok(PaginationSnapshot {
        section_start_kind: view.section_start_kind,
        section_page_number_kind: view.section_page_number_kind,
        section_page_number_start: view.section_page_number_start,
    })
}

/// Borrowed, presence-preserving aggregate settings from one native section.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SectionSettingsSnapshot<'source> {
    inherit_previous_header_footer: Option<bool>,
    section_template_first_page_different: Option<bool>,
    section_template_even_odd_pages_different: Option<bool>,
    section_start_kind: Option<u32>,
    section_page_number_kind: Option<u32>,
    section_page_number_start: Option<u32>,
    name: Option<&'source str>,
    section_template_first_page_hides_header_footer: Option<bool>,
}

impl<'source> SectionSettingsSnapshot<'source> {
    /// Optional native header/footer inheritance flag from field 17.
    #[must_use]
    pub const fn inherit_previous_header_footer(self) -> Option<bool> {
        self.inherit_previous_header_footer
    }

    /// Optional native first-page distinction flag from field 18.
    #[must_use]
    pub const fn section_template_first_page_different(self) -> Option<bool> {
        self.section_template_first_page_different
    }

    /// Optional native even/odd distinction flag from field 19.
    #[must_use]
    pub const fn section_template_even_odd_pages_different(self) -> Option<bool> {
        self.section_template_even_odd_pages_different
    }

    /// Optional native section-start discriminant from field 20.
    #[must_use]
    pub const fn section_start_kind(self) -> Option<u32> {
        self.section_start_kind
    }

    /// Optional native page-numbering discriminant from field 21.
    #[must_use]
    pub const fn section_page_number_kind(self) -> Option<u32> {
        self.section_page_number_kind
    }

    /// Optional native non-zero first page number from field 22.
    #[must_use]
    pub const fn section_page_number_start(self) -> Option<u32> {
        self.section_page_number_start
    }

    /// Optional UTF-8 section name borrowed directly from field 26.
    #[must_use]
    pub const fn name(self) -> Option<&'source str> {
        self.name
    }

    /// Optional native first-page header/footer hiding flag from field 28.
    #[must_use]
    pub const fn section_template_first_page_hides_header_footer(self) -> Option<bool> {
        self.section_template_first_page_hides_header_footer
    }
}

/// Exact finite consumption of one aggregate section-settings projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    name_bytes: usize,
}

impl DecodeReport {
    /// Encoded field records visited by strict routing.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Bytes visited by the strict and Buffa root passes.
    #[must_use]
    pub const fn work_bytes(self) -> usize {
        self.work_bytes
    }

    /// Greatest protobuf message or unknown-group depth reached.
    #[must_use]
    pub const fn max_depth(self) -> u32 {
        self.max_depth
    }

    /// Bytes borrowed by the optional section name.
    #[must_use]
    pub const fn name_bytes(self) -> usize {
        self.name_bytes
    }
}

/// Decode all aggregate section settings without exposing generated values.
pub fn decode_section_settings<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<SectionSettingsSnapshot<'source>, DecodeError> {
    Ok(decode_section_settings_with_report(source, options)?.0)
}

/// Decode aggregate section settings and report exact finite consumption.
pub fn decode_section_settings_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(SectionSettingsSnapshot<'source>, DecodeReport), DecodeError> {
    let mut budget = Budget::new(source, options)?;
    budget.charge_work(source.len())?;
    let strict = strict_section_settings(source, &mut budget)?;

    budget.charge_work(source.len())?;
    let view: projection::PagesSectionSettingsArchiveLazyView<'source> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid())?;
    let projected = SectionSettingsSnapshot {
        inherit_previous_header_footer: view.inherit_previous_header_footer,
        section_template_first_page_different: view.section_template_first_page_different,
        section_template_even_odd_pages_different: view.section_template_even_odd_pages_different,
        section_start_kind: view.section_start_kind,
        section_page_number_kind: view.section_page_number_kind,
        section_page_number_start: view.section_page_number_start,
        name: view.name,
        section_template_first_page_hides_header_footer: view
            .section_template_first_page_hides_header_footer,
    };
    if projected != strict {
        return Err(DecodeError::invalid());
    }
    Ok((strict, budget.report()))
}

const INHERIT_HEADER_FOOTER_FIELD: u32 = 17;
const FIRST_PAGE_DIFFERENT_FIELD: u32 = 18;
const EVEN_ODD_PAGES_DIFFERENT_FIELD: u32 = 19;
const SECTION_START_FIELD: u32 = 20;
const PAGE_NUMBERING_FIELD: u32 = 21;
const STARTING_PAGE_NUMBER_FIELD: u32 = 22;
const SECTION_NAME_FIELD: u32 = 26;
const FIRST_PAGE_HIDES_HEADER_FOOTER_FIELD: u32 = 28;
const MAX_RECURSION: u32 = 64;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;

fn strict_section_settings<'source>(
    source: &'source [u8],
    budget: &mut Budget,
) -> Result<SectionSettingsSnapshot<'source>, DecodeError> {
    let mut snapshot = SectionSettingsSnapshot::default();
    let mut remaining = source;
    while let Some(field) = next_root_field(&mut remaining, budget, 1)? {
        match field.number {
            INHERIT_HEADER_FOOTER_FIELD => {
                if snapshot.inherit_previous_header_footer.is_some() {
                    return Err(DecodeError::invalid());
                }
                snapshot.inherit_previous_header_footer = Some(canonical_bool(field.varint()?)?);
            },
            FIRST_PAGE_DIFFERENT_FIELD => {
                if snapshot.section_template_first_page_different.is_some() {
                    return Err(DecodeError::invalid());
                }
                snapshot.section_template_first_page_different =
                    Some(canonical_bool(field.varint()?)?);
            },
            EVEN_ODD_PAGES_DIFFERENT_FIELD => {
                if snapshot.section_template_even_odd_pages_different.is_some() {
                    return Err(DecodeError::invalid());
                }
                snapshot.section_template_even_odd_pages_different =
                    Some(canonical_bool(field.varint()?)?);
            },
            SECTION_START_FIELD => {
                if snapshot.section_start_kind.is_some() {
                    return Err(DecodeError::invalid());
                }
                snapshot.section_start_kind = Some(canonical_u32(field.varint()?)?);
            },
            PAGE_NUMBERING_FIELD => {
                if snapshot.section_page_number_kind.is_some() {
                    return Err(DecodeError::invalid());
                }
                snapshot.section_page_number_kind = Some(canonical_u32(field.varint()?)?);
            },
            STARTING_PAGE_NUMBER_FIELD => {
                if snapshot.section_page_number_start.is_some() {
                    return Err(DecodeError::invalid());
                }
                let number = canonical_u32(field.varint()?)?;
                if number == 0 {
                    return Err(DecodeError::invalid());
                }
                snapshot.section_page_number_start = Some(number);
            },
            SECTION_NAME_FIELD => {
                if snapshot.name.is_some() {
                    return Err(DecodeError::invalid());
                }
                let bytes = field.length_delimited()?;
                budget.charge_name(bytes.len())?;
                let name = std::str::from_utf8(bytes).map_err(|_error| DecodeError::invalid())?;
                if name.contains('\0') {
                    return Err(DecodeError::invalid());
                }
                snapshot.name = Some(name);
            },
            FIRST_PAGE_HIDES_HEADER_FOOTER_FIELD => {
                if snapshot
                    .section_template_first_page_hides_header_footer
                    .is_some()
                {
                    return Err(DecodeError::invalid());
                }
                snapshot.section_template_first_page_hides_header_footer =
                    Some(canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    Ok(snapshot)
}

#[derive(Clone, Copy)]
struct StrictField<'source> {
    number: u32,
    wire_type: u8,
    value: StrictValue<'source>,
}

impl<'source> StrictField<'source> {
    fn varint(self) -> Result<u64, DecodeError> {
        match self.value {
            StrictValue::Varint(value) if self.wire_type == 0 => Ok(value),
            StrictValue::Varint(_) | StrictValue::LengthDelimited(_) | StrictValue::Other => {
                Err(DecodeError::invalid())
            },
        }
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        match self.value {
            StrictValue::LengthDelimited(value) if self.wire_type == 2 => Ok(value),
            StrictValue::Varint(_) | StrictValue::LengthDelimited(_) | StrictValue::Other => {
                Err(DecodeError::invalid())
            },
        }
    }
}

#[derive(Clone, Copy)]
enum StrictValue<'source> {
    Varint(u64),
    LengthDelimited(&'source [u8]),
    Other,
}

enum ParseItem<'source> {
    Field(StrictField<'source>),
    EndGroup(u32),
}

fn next_root_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_field(source, budget, depth)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(_)) => Err(DecodeError::invalid()),
        None => Ok(None),
    }
}

fn parse_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    budget.observe_depth(depth)?;
    budget.charge_field()?;
    let tag = take_varint(source)?;
    let number = u32::try_from(tag >> 3).map_err(|_conversion| DecodeError::invalid())?;
    let wire_type = u8::try_from(tag & 7).map_err(|_conversion| DecodeError::invalid())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(DecodeError::invalid());
    }
    let value = match wire_type {
        0 => StrictValue::Varint(take_varint(source)?),
        1 => {
            take(source, 8)?;
            StrictValue::Other
        },
        2 => {
            let length = usize::try_from(take_varint(source)?)
                .map_err(|_conversion| DecodeError::invalid())?;
            StrictValue::LengthDelimited(take(source, length)?)
        },
        3 => {
            let child_depth = depth.checked_add(1).ok_or_else(DecodeError::invalid)?;
            skip_group(source, number, budget, child_depth)?;
            StrictValue::Other
        },
        4 => return Ok(Some(ParseItem::EndGroup(number))),
        5 => {
            take(source, 4)?;
            StrictValue::Other
        },
        _ => return Err(DecodeError::invalid()),
    };
    Ok(Some(ParseItem::Field(StrictField {
        number,
        wire_type,
        value,
    })))
}

fn skip_group(
    source: &mut &[u8],
    expected_number: u32,
    budget: &mut Budget,
    depth: u32,
) -> Result<(), DecodeError> {
    loop {
        match parse_field(source, budget, depth)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected_number => return Ok(()),
            Some(ParseItem::EndGroup(_)) | None => return Err(DecodeError::invalid()),
        }
    }
}

fn take<'source>(source: &mut &'source [u8], amount: usize) -> Result<&'source [u8], DecodeError> {
    if source.len() < amount {
        return Err(DecodeError::invalid());
    }
    let (selected, remaining) = source.split_at(amount);
    *source = remaining;
    Ok(selected)
}

fn take_varint(source: &mut &[u8]) -> Result<u64, DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original.get(index).ok_or_else(DecodeError::invalid)?;
        if index == 9 && byte > 1 {
            return Err(DecodeError::invalid());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            if encoded_varint_len(value) != consumed {
                return Err(DecodeError::invalid());
            }
            *source = &original[consumed..];
            return Ok(value);
        }
    }
    Err(DecodeError::invalid())
}

const fn encoded_varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

fn canonical_u32(value: u64) -> Result<u32, DecodeError> {
    u32::try_from(value).map_err(|_conversion| DecodeError::invalid())
}

fn canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::invalid()),
    }
}

struct Budget {
    options: DecodeOptions,
    fields: usize,
    work_bytes: usize,
    max_depth: u32,
    name_bytes: usize,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        let hard_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
            .map_err(|_conversion| DecodeError::invalid())?;
        if options.max_message_bytes > hard_bytes {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: options.max_message_bytes,
                maximum: hard_bytes,
            }));
        }
        if source.len() > options.max_message_bytes {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION,
            }));
        }
        let mut budget = Self {
            options,
            fields: 0,
            work_bytes: 0,
            max_depth: 0,
            name_bytes: 0,
        };
        budget.observe_depth(1)?;
        Ok(budget)
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(1);
        if observed > self.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed,
                maximum: self.options.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.saturating_add(amount);
        if observed > self.options.max_work_bytes {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed,
                maximum: self.options.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn charge_name(&mut self, amount: usize) -> Result<(), DecodeError> {
        if amount > self.options.max_name_bytes {
            return Err(DecodeError::limited(DecodeLimit::NameBytes {
                observed: amount,
                maximum: self.options.max_name_bytes,
            }));
        }
        self.name_bytes = amount;
        Ok(())
    }

    fn observe_depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    const fn report(&self) -> DecodeReport {
        DecodeReport {
            fields: self.fields,
            work_bytes: self.work_bytes,
            max_depth: self.max_depth,
            name_bytes: self.name_bytes,
        }
    }
}

#[cfg(test)]
mod tests {
    use prost::Message as _;

    use super::*;

    fn decode(source: &[u8]) -> Result<PaginationSnapshot, DecodeError> {
        decode_pagination(source, DecodeOptions::new(source.len(), 1))
    }

    #[test]
    fn canonical_prost_section_matches_projection() -> Result<(), Box<dyn std::error::Error>> {
        let source = crate::tp::SectionArchive {
            section_start_kind: Some(2),
            section_page_number_kind: Some(1),
            section_page_number_start: Some(42),
            name: Some("opaque to this projection".to_owned()),
            ..crate::tp::SectionArchive::default()
        }
        .encode_to_vec();
        assert_eq!(
            decode(&source)?,
            PaginationSnapshot {
                section_start_kind: Some(2),
                section_page_number_kind: Some(1),
                section_page_number_start: Some(42),
            }
        );
        Ok(())
    }

    #[test]
    fn absent_scalars_and_opaque_unknown_payload_remain_allocation_free()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = [0xd2, 0x0c, 0x03, 0xff, 0x00, 0xfe];
        assert_eq!(decode(&source)?, PaginationSnapshot::default());
        Ok(())
    }

    #[test]
    fn malformed_selected_scalar_is_rejected() {
        let Err(error) = decode(&[0xa2, 0x01, 0x01, 0x00]) else {
            panic!("field 20 with a length-delimited wire type must fail");
        };
        assert!(!error.to_string().is_empty());
    }

    fn push_varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = u8::try_from(value & 0x7f).expect("varint chunk fits u8");
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return;
            }
        }
    }

    fn push_key(output: &mut Vec<u8>, field: u32, wire_type: u8) {
        push_varint(output, (u64::from(field) << 3) | u64::from(wire_type));
    }

    fn push_varint_field(output: &mut Vec<u8>, field: u32, value: u64) {
        push_key(output, field, 0);
        push_varint(output, value);
    }

    fn push_length_field(output: &mut Vec<u8>, field: u32, value: &[u8]) {
        push_key(output, field, 2);
        push_varint(
            output,
            u64::try_from(value.len()).expect("fixture length fits u64"),
        );
        output.extend_from_slice(value);
    }

    #[derive(Clone, Copy)]
    struct AggregateInput<'name> {
        inherit: Option<bool>,
        first: Option<bool>,
        even_odd: Option<bool>,
        start: Option<u32>,
        numbering: Option<u32>,
        page: Option<u32>,
        name: Option<&'name str>,
        hides: Option<bool>,
    }

    fn aggregate_payload(input: AggregateInput<'_>) -> Vec<u8> {
        let mut source = Vec::new();
        for (field, value) in [
            (INHERIT_HEADER_FOOTER_FIELD, input.inherit),
            (FIRST_PAGE_DIFFERENT_FIELD, input.first),
            (EVEN_ODD_PAGES_DIFFERENT_FIELD, input.even_odd),
        ] {
            if let Some(value) = value {
                push_varint_field(&mut source, field, u64::from(value));
            }
        }
        for (field, value) in [
            (SECTION_START_FIELD, input.start),
            (PAGE_NUMBERING_FIELD, input.numbering),
            (STARTING_PAGE_NUMBER_FIELD, input.page),
        ] {
            if let Some(value) = value {
                push_varint_field(&mut source, field, u64::from(value));
            }
        }
        if let Some(name) = input.name {
            push_length_field(&mut source, SECTION_NAME_FIELD, name.as_bytes());
        }
        if let Some(value) = input.hides {
            push_varint_field(
                &mut source,
                FIRST_PAGE_HIDES_HEADER_FOOTER_FIELD,
                u64::from(value),
            );
        }
        source
    }

    fn generous(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(source.len().max(1), MAX_RECURSION)
            .with_max_fields(usize::MAX)
            .with_max_work_bytes(usize::MAX)
            .with_max_name_bytes(source.len())
    }

    #[test]
    fn aggregate_lazy_view_borrows_every_presence_sensitive_setting() {
        let source = aggregate_payload(AggregateInput {
            inherit: Some(true),
            first: Some(false),
            even_odd: Some(true),
            start: Some(2),
            numbering: Some(7),
            page: Some(u32::MAX),
            name: Some("Borrowed 章節"),
            hides: Some(false),
        });
        let (snapshot, report) = decode_section_settings_with_report(&source, generous(&source))
            .expect("valid settings");
        assert_eq!(snapshot.inherit_previous_header_footer(), Some(true));
        assert_eq!(
            snapshot.section_template_first_page_different(),
            Some(false)
        );
        assert_eq!(
            snapshot.section_template_even_odd_pages_different(),
            Some(true)
        );
        assert_eq!(snapshot.section_start_kind(), Some(2));
        assert_eq!(snapshot.section_page_number_kind(), Some(7));
        assert_eq!(snapshot.section_page_number_start(), Some(u32::MAX));
        assert_eq!(snapshot.name(), Some("Borrowed 章節"));
        assert_eq!(
            snapshot.section_template_first_page_hides_header_footer(),
            Some(false)
        );
        let name = snapshot.name().expect("name present");
        let offset = name.as_ptr() as usize - source.as_ptr() as usize;
        assert_eq!(&source[offset..offset + name.len()], name.as_bytes());
        assert_eq!(report.fields(), 8);
        assert_eq!(report.work_bytes(), source.len() * 2);
        assert_eq!(report.max_depth(), 1);
        assert_eq!(report.name_bytes(), name.len());
    }

    #[test]
    fn all_presence_combinations_remain_distinct() {
        let booleans = [None, Some(false), Some(true)];
        let starts = [None, Some(0), Some(u32::MAX)];
        let pages = [None, Some(1), Some(u32::MAX)];
        let names = [None, Some(""), Some("Named")];
        for inherit in booleans {
            for first in booleans {
                for even_odd in booleans {
                    for hides in booleans {
                        for start in starts {
                            for numbering in starts {
                                for page in pages {
                                    for name in names {
                                        let input = AggregateInput {
                                            inherit,
                                            first,
                                            even_odd,
                                            start,
                                            numbering,
                                            page,
                                            name,
                                            hides,
                                        };
                                        let source = aggregate_payload(input);
                                        let snapshot =
                                            decode_section_settings(&source, generous(&source))
                                                .expect("presence combination");
                                        assert_eq!(
                                            snapshot.inherit_previous_header_footer(),
                                            inherit
                                        );
                                        assert_eq!(
                                            snapshot.section_template_first_page_different(),
                                            first
                                        );
                                        assert_eq!(
                                            snapshot.section_template_even_odd_pages_different(),
                                            even_odd
                                        );
                                        assert_eq!(snapshot.section_start_kind(), start);
                                        assert_eq!(snapshot.section_page_number_kind(), numbering);
                                        assert_eq!(snapshot.section_page_number_start(), page);
                                        assert_eq!(snapshot.name(), name);
                                        assert_eq!(
                                            snapshot
                                                .section_template_first_page_hides_header_footer(),
                                            hides
                                        );
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    #[test]
    fn pagination_projection_remains_isolated_from_new_selected_fields() {
        let mut source = Vec::new();
        push_varint_field(&mut source, INHERIT_HEADER_FOOTER_FIELD, 2);
        push_length_field(&mut source, SECTION_NAME_FIELD, &[0xff]);
        push_varint_field(&mut source, SECTION_START_FIELD, 2);
        assert_eq!(
            decode_pagination(&source, DecodeOptions::new(source.len(), 1))
                .expect("non-pagination fields stay opaque"),
            PaginationSnapshot {
                section_start_kind: Some(2),
                section_page_number_kind: None,
                section_page_number_start: None,
            }
        );
        assert!(decode_section_settings(&source, generous(&source)).is_err());
    }

    #[test]
    fn every_selected_field_rejects_duplicates_and_wrong_wire_types() {
        let selected = [
            INHERIT_HEADER_FOOTER_FIELD,
            FIRST_PAGE_DIFFERENT_FIELD,
            EVEN_ODD_PAGES_DIFFERENT_FIELD,
            SECTION_START_FIELD,
            PAGE_NUMBERING_FIELD,
            STARTING_PAGE_NUMBER_FIELD,
            SECTION_NAME_FIELD,
            FIRST_PAGE_HIDES_HEADER_FOOTER_FIELD,
        ];
        for field in selected {
            let mut duplicate = Vec::new();
            if field == SECTION_NAME_FIELD {
                push_length_field(&mut duplicate, field, b"one");
                push_length_field(&mut duplicate, field, b"two");
            } else {
                push_varint_field(&mut duplicate, field, 1);
                push_varint_field(&mut duplicate, field, 1);
            }
            assert!(
                decode_section_settings(&duplicate, generous(&duplicate)).is_err(),
                "duplicate field {field}"
            );

            let mut wrong_wire = Vec::new();
            if field == SECTION_NAME_FIELD {
                push_varint_field(&mut wrong_wire, field, 1);
            } else {
                push_length_field(&mut wrong_wire, field, &[1]);
            }
            assert!(
                decode_section_settings(&wrong_wire, generous(&wrong_wire)).is_err(),
                "wrong wire field {field}"
            );
        }
    }

    #[test]
    fn selected_scalars_and_name_reject_noncanonical_or_invalid_values() {
        for field in [
            INHERIT_HEADER_FOOTER_FIELD,
            FIRST_PAGE_DIFFERENT_FIELD,
            EVEN_ODD_PAGES_DIFFERENT_FIELD,
            FIRST_PAGE_HIDES_HEADER_FOOTER_FIELD,
        ] {
            let mut source = Vec::new();
            push_varint_field(&mut source, field, 2);
            assert!(decode_section_settings(&source, generous(&source)).is_err());
        }

        for field in [
            SECTION_START_FIELD,
            PAGE_NUMBERING_FIELD,
            STARTING_PAGE_NUMBER_FIELD,
        ] {
            let mut source = Vec::new();
            push_varint_field(&mut source, field, u64::from(u32::MAX) + 1);
            assert!(decode_section_settings(&source, generous(&source)).is_err());
        }

        let mut page_zero = Vec::new();
        push_varint_field(&mut page_zero, STARTING_PAGE_NUMBER_FIELD, 0);
        assert!(decode_section_settings(&page_zero, generous(&page_zero)).is_err());

        for name in [&[0xff][..], b"bad\0name"] {
            let mut source = Vec::new();
            push_length_field(&mut source, SECTION_NAME_FIELD, name);
            assert!(decode_section_settings(&source, generous(&source)).is_err());
        }

        assert!(decode_section_settings(&[0], DecodeOptions::new(1, 1)).is_err());
        let overlong_key = [0x88, 0x81, 0x00, 0x00];
        assert!(
            decode_section_settings(&overlong_key, DecodeOptions::new(overlong_key.len(), 1))
                .is_err()
        );
        let overlong_value = [0x88, 0x01, 0x80, 0x00];
        assert!(
            decode_section_settings(&overlong_value, DecodeOptions::new(overlong_value.len(), 1))
                .is_err()
        );
        let overlong_length = [0xd2, 0x01, 0x80, 0x00];
        assert!(
            decode_section_settings(
                &overlong_length,
                DecodeOptions::new(overlong_length.len(), 1)
            )
            .is_err()
        );
    }

    #[test]
    fn truncated_and_reserved_wire_forms_fail_without_panicking() {
        let malformed = [
            &[0x09, 0x00][..],
            &[0x12, 0x02, 0x00][..],
            &[0x0d, 0x00][..],
            &[0x0e][..],
            &[0x0f][..],
            &[0x08, 0x80][..],
            &[
                0x08, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x02,
            ][..],
        ];
        for source in malformed {
            assert!(decode_section_settings(source, DecodeOptions::new(source.len(), 2)).is_err());
        }
    }

    #[test]
    fn exact_aggregate_limits_are_inclusive_and_typed() {
        let source = aggregate_payload(AggregateInput {
            inherit: Some(true),
            first: Some(false),
            even_odd: Some(true),
            start: Some(2),
            numbering: Some(1),
            page: Some(42),
            name: Some("Limits"),
            hides: Some(false),
        });
        let (_, exact) =
            decode_section_settings_with_report(&source, generous(&source)).expect("baseline");
        let exact_options = DecodeOptions::new(source.len(), exact.max_depth())
            .with_max_fields(exact.fields())
            .with_max_work_bytes(exact.work_bytes())
            .with_max_name_bytes(exact.name_bytes());
        assert!(decode_section_settings(&source, exact_options).is_ok());

        let cases = [
            (
                DecodeOptions::new(source.len() - 1, exact.max_depth())
                    .with_max_fields(exact.fields())
                    .with_max_work_bytes(exact.work_bytes())
                    .with_max_name_bytes(exact.name_bytes()),
                DecodeLimit::Bytes {
                    observed: source.len(),
                    maximum: source.len() - 1,
                },
            ),
            (
                DecodeOptions::new(source.len(), exact.max_depth())
                    .with_max_fields(exact.fields() - 1)
                    .with_max_work_bytes(exact.work_bytes())
                    .with_max_name_bytes(exact.name_bytes()),
                DecodeLimit::Fields {
                    observed: exact.fields(),
                    maximum: exact.fields() - 1,
                },
            ),
            (
                DecodeOptions::new(source.len(), exact.max_depth())
                    .with_max_fields(exact.fields())
                    .with_max_work_bytes(exact.work_bytes() - 1)
                    .with_max_name_bytes(exact.name_bytes()),
                DecodeLimit::Work {
                    observed: exact.work_bytes(),
                    maximum: exact.work_bytes() - 1,
                },
            ),
            (
                DecodeOptions::new(source.len(), exact.max_depth())
                    .with_max_fields(exact.fields())
                    .with_max_work_bytes(exact.work_bytes())
                    .with_max_name_bytes(exact.name_bytes() - 1),
                DecodeLimit::NameBytes {
                    observed: exact.name_bytes(),
                    maximum: exact.name_bytes() - 1,
                },
            ),
        ];
        for (options, expected) in cases {
            let error = decode_section_settings(&source, options).expect_err("limit must fail");
            assert_eq!(error.resource_limit(), Some(expected));
        }

        let zero_nesting =
            decode_section_settings(&[], DecodeOptions::new(1, 0)).expect_err("zero nesting");
        assert_eq!(
            zero_nesting.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 0,
                maximum: MAX_RECURSION,
            })
        );
        let hard = usize::try_from(buffa::MAX_MESSAGE_BYTES).expect("hard limit fits usize");
        let excessive = decode_section_settings(&[], DecodeOptions::new(hard + 1, 1))
            .expect_err("hard byte limit");
        assert_eq!(
            excessive.resource_limit(),
            Some(DecodeLimit::Bytes {
                observed: hard + 1,
                maximum: hard,
            })
        );
    }

    #[test]
    fn canonical_unknown_records_and_groups_are_bounded_but_opaque() {
        let mut source = Vec::new();
        push_varint_field(&mut source, 100, 7);
        push_key(&mut source, 101, 1);
        source.extend_from_slice(&u64::MAX.to_le_bytes());
        push_length_field(&mut source, 102, &[0xff, 0x00]);
        push_key(&mut source, 103, 5);
        source.extend_from_slice(&[1, 2, 3, 4]);
        push_key(&mut source, 104, 3);
        push_varint_field(&mut source, 105, 9);
        push_key(&mut source, 104, 4);
        push_varint_field(&mut source, INHERIT_HEADER_FOOTER_FIELD, 1);
        let (snapshot, report) =
            decode_section_settings_with_report(&source, generous(&source)).expect("opaque fields");
        assert_eq!(snapshot.inherit_previous_header_footer(), Some(true));
        assert_eq!(report.fields(), 8);
        assert_eq!(report.max_depth(), 2);
        assert_eq!(report.work_bytes(), source.len() * 2);

        let mut unterminated = Vec::new();
        push_key(&mut unterminated, 104, 3);
        push_varint_field(&mut unterminated, 105, 9);
        assert!(decode_section_settings(&unterminated, generous(&unterminated)).is_err());
        let mut mismatched = Vec::new();
        push_key(&mut mismatched, 104, 3);
        push_key(&mut mismatched, 105, 4);
        assert!(decode_section_settings(&mismatched, generous(&mismatched)).is_err());
    }

    #[test]
    fn four_thousand_to_eight_thousand_field_routing_scales_linearly() {
        fn wide(fields: usize) -> Vec<u8> {
            let mut source = Vec::new();
            for index in 0..fields {
                push_varint_field(
                    &mut source,
                    100,
                    u64::try_from(index).expect("field index fits u64"),
                );
            }
            source
        }

        const SMALL: usize = 4_096;
        const LARGE: usize = 8_192;
        let small_source = wide(SMALL);
        let large_source = wide(LARGE);
        let (_, small) = decode_section_settings_with_report(
            &small_source,
            DecodeOptions::new(small_source.len(), 1)
                .with_max_fields(SMALL)
                .with_max_work_bytes(usize::MAX),
        )
        .expect("small routing");
        let (_, large) = decode_section_settings_with_report(
            &large_source,
            DecodeOptions::new(large_source.len(), 1)
                .with_max_fields(LARGE)
                .with_max_work_bytes(usize::MAX),
        )
        .expect("large routing");
        assert_eq!(small.fields(), SMALL);
        assert_eq!(large.fields(), LARGE);
        assert_eq!(small.work_bytes(), small_source.len() * 2);
        assert_eq!(large.work_bytes(), large_source.len() * 2);
        assert!(large_source.len() * 10 <= small_source.len() * 23);
        assert!(large.fields() * 10 <= small.fields() * 23);
        assert!(large.work_bytes() * 10 <= small.work_bytes() * 23);

        let field_error = decode_section_settings(
            &large_source,
            DecodeOptions::new(large_source.len(), 1)
                .with_max_fields(LARGE - 1)
                .with_max_work_bytes(large.work_bytes()),
        )
        .expect_err("field ceiling");
        assert_eq!(
            field_error.resource_limit(),
            Some(DecodeLimit::Fields {
                observed: LARGE,
                maximum: LARGE - 1,
            })
        );
    }
}
