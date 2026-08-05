//! Strict MS-PPT text-range interaction records.
//!
//! Text actions use the ordinary inert `InteractiveInfo` model followed by a
//! `TextInteractiveInfoAtom`. Offsets are UTF-16 code-unit positions and no
//! target is resolved or activated here.

use super::hyperlink::{Interaction, InteractionLimits, InteractionTrigger, encode_record};
use super::package::{Error, Result};
use super::records::Record;
use crate::consts::RecordType;

/// Resource limits for text-range interaction parsing and authoring.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextInteractionLimits {
    /// Limits for each paired `InteractiveInfo` record.
    pub interaction: InteractionLimits,
    /// Maximum number of action/range pairs in one text body.
    pub max_interactions: usize,
    /// Maximum number of UTF-16 code units in the corresponding text.
    pub max_text_units: u32,
}

impl Default for TextInteractionLimits {
    fn default() -> Self {
        Self {
            interaction: InteractionLimits::default(),
            max_interactions: 4096,
            max_text_units: 16 * 1024 * 1024,
        }
    }
}

/// MS-PPT `TextTypeEnum` stored by `TextHeaderAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextType {
    Title,
    Body,
    Notes,
    Other,
    CenterBody,
    CenterTitle,
    HalfBody,
    QuarterBody,
}

impl TextType {
    pub(crate) fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Title),
            1 => Ok(Self::Body),
            2 => Ok(Self::Notes),
            4 => Ok(Self::Other),
            5 => Ok(Self::CenterBody),
            6 => Ok(Self::CenterTitle),
            7 => Ok(Self::HalfBody),
            8 => Ok(Self::QuarterBody),
            _ => corrupted("TextHeaderAtom has an invalid TextTypeEnum value"),
        }
    }

    /// Numeric value used by the binary record.
    pub const fn value(self) -> u32 {
        match self {
            Self::Title => 0,
            Self::Body => 1,
            Self::Notes => 2,
            Self::Other => 4,
            Self::CenterBody => 5,
            Self::CenterTitle => 6,
            Self::HalfBody => 7,
            Self::QuarterBody => 8,
        }
    }
}

/// Non-empty half-open text range measured in UTF-16 code units.
///
/// Legacy PPT logical text includes one implicit final paragraph mark after
/// the units serialized by `TextCharsAtom` or `TextBytesAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub struct TextRange {
    begin: u32,
    end: u32,
}

impl TextRange {
    /// Construct a representable, non-empty range.
    pub fn new(begin: u32, end: u32) -> Result<Self> {
        let value = Self { begin, end };
        value.validate_shape()?;
        Ok(value)
    }

    /// Number of UTF-16 code units covered by this range.
    pub const fn len(self) -> u32 {
        self.end - self.begin
    }

    pub const fn begin(self) -> u32 {
        self.begin
    }

    pub const fn end(self) -> u32 {
        self.end
    }

    /// Text ranges are always non-empty.
    pub const fn is_empty(self) -> bool {
        false
    }

    /// Validate this range against serialized text and its final paragraph mark.
    pub fn validate_for_text(self, serialized_text_units: u32) -> Result<()> {
        self.validate_shape()?;
        let logical_text_units = serialized_text_units.saturating_add(1);
        if self.end > logical_text_units {
            return corrupted(format!(
                "Text interaction range [{}, {}) extends beyond {serialized_text_units} UTF-16 text units and the final paragraph mark",
                self.begin, self.end
            ));
        }
        Ok(())
    }

    fn validate_shape(self) -> Result<()> {
        if self.begin >= self.end {
            return corrupted("Text interaction range must be non-empty and increasing");
        }
        if self.end > i32::MAX as u32 {
            return corrupted("Text interaction range exceeds the signed TextPosition domain");
        }
        Ok(())
    }
}

/// One inert click or mouse-over action anchored to a text range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextInteraction {
    pub range: TextRange,
    pub interaction: Interaction,
}

impl TextInteraction {
    /// Construct and validate a range/action pair without a text-length bound.
    pub fn new(range: TextRange, interaction: Interaction) -> Result<Self> {
        range.validate_shape()?;
        Ok(Self { range, interaction })
    }

    /// Validate the complete pair against resource and corresponding-text bounds.
    pub fn validate_for_text(&self, text_units: u32, limits: TextInteractionLimits) -> Result<()> {
        if text_units > limits.max_text_units {
            return corrupted("Corresponding text exceeds the configured interaction limit");
        }
        self.range.validate_for_text(text_units)?;
        self.interaction.validate_with_limits(limits.interaction)?;
        Ok(())
    }

    /// Serialize the canonical `InteractiveInfo` and matching range atom.
    pub fn to_bytes_for_text(
        &self,
        text_units: u32,
        limits: TextInteractionLimits,
    ) -> Result<Vec<u8>> {
        self.validate_for_text(text_units, limits)?;
        let mut bytes = self.interaction.to_bytes_with_limits(limits.interaction)?;
        let mut range = [0u8; 8];
        range[0..4].copy_from_slice(&self.range.begin.to_le_bytes());
        range[4..8].copy_from_slice(&self.range.end.to_le_bytes());
        bytes.extend_from_slice(&encode_record(
            0,
            trigger_instance(self.interaction.trigger),
            RecordType::TextInteractiveInfoAtom as u16,
            &range,
        )?);
        Ok(bytes)
    }

    /// Parse every paired interaction in one corresponding text body.
    pub(crate) fn parse_records<'a>(
        records: impl IntoIterator<Item = &'a Record>,
        text_units: u32,
        limits: TextInteractionLimits,
    ) -> Result<Vec<Self>> {
        if text_units > limits.max_text_units {
            return corrupted("Corresponding text exceeds the configured interaction limit");
        }
        let mut records = records.into_iter().peekable();
        let mut result = Vec::new();
        let mut interaction_section = false;
        let mut terminal_text_records = false;
        while let Some(record) = records.next() {
            if record.record_type == RecordType::TextInteractiveInfoAtom {
                return corrupted("TextInteractiveInfoAtom has no preceding InteractiveInfo");
            }
            if record.record_type != RecordType::InteractiveInfo {
                if interaction_section {
                    if !matches!(
                        record.record_type,
                        RecordType::TextRulerAtom
                            | RecordType::MasterTextPropAtom
                            // Seen after the interaction pair in established producer output,
                            // despite its earlier position in the normative ABNF.
                            | RecordType::TextSpecInfoAtom
                    ) {
                        return corrupted(format!(
                            "Text body has a nonterminal {:?} record after interactive information",
                            record.record_type
                        ));
                    }
                    terminal_text_records = true;
                }
                continue;
            }
            if terminal_text_records {
                return corrupted("Text interaction appears after terminal text records");
            }
            interaction_section = true;
            if result.len() >= limits.max_interactions {
                return corrupted("Text body exceeds the configured interaction count");
            }
            let interaction = Interaction::parse_with_limits(record, limits.interaction)?;
            let anchor = records.next().ok_or_else(|| {
                Error::Corrupted(
                    "Text InteractiveInfo has no following TextInteractiveInfoAtom".to_string(),
                )
            })?;
            let range = parse_anchor(anchor, interaction.trigger, text_units)?;
            result.push(Self { range, interaction });
        }
        Ok(result)
    }
}

/// One text body and its range-anchored actions.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextBodyInteractions {
    /// `TextHeaderAtom.recInstance`, identifying the body within its slide set.
    pub text_header_instance: u16,
    pub text_type: TextType,
    pub text: String,
    pub interactions: Vec<TextInteraction>,
    /// Header/footer metacharacter placeholders in this body, in record order.
    pub metachars: Vec<crate::text_metachar::TextMetachar>,
}

/// Text interactions attached to one OfficeArt shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeTextInteractionEntry {
    pub shape_id: u32,
    pub interactions: Vec<TextInteraction>,
}

pub(crate) fn parse_text_bodies(
    records: &[&Record],
    limits: TextInteractionLimits,
) -> Result<Vec<TextBodyInteractions>> {
    let mut result = Vec::new();
    let mut offset = 0;
    while offset < records.len() {
        if records[offset].record_type != RecordType::TextHeaderAtom {
            if matches!(
                records[offset].record_type,
                RecordType::InteractiveInfo | RecordType::TextInteractiveInfoAtom
            ) {
                return corrupted("Text interaction appears before a TextHeaderAtom");
            }
            offset += 1;
            continue;
        }
        let header = records[offset];
        if header.version != 0
            || header.data_length != 4
            || header.data.len() != 4
            || header.record_type_raw != RecordType::TextHeaderAtom as u16
        {
            return corrupted("TextHeaderAtom has an invalid header or length");
        }
        let text_type = TextType::parse(u32::from_le_bytes(header.data[..4].try_into().unwrap()))?;
        let end = records[offset + 1..]
            .iter()
            .position(|record| record.record_type == RecordType::TextHeaderAtom)
            .map_or(records.len(), |relative| offset + 1 + relative);
        let body = &records[offset + 1..end];
        let has_interactions = body.iter().any(|record| {
            matches!(
                record.record_type,
                RecordType::InteractiveInfo | RecordType::TextInteractiveInfoAtom
            )
        });
        if has_interactions {
            let text_units = text_units_from_records(body.iter().copied())?;
            let interactions =
                TextInteraction::parse_records(body.iter().copied(), text_units, limits)?;
            let text = exact_text_from_records(body.iter().copied())?;
            let metachars = crate::text_metachar::metachars_from_records(body.iter().copied())?;
            result.push(TextBodyInteractions {
                text_header_instance: header.instance,
                text_type,
                text,
                interactions,
                metachars,
            });
        }
        offset = end;
    }
    Ok(result)
}

pub(crate) fn text_units_from_records<'a>(
    records: impl IntoIterator<Item = &'a Record>,
) -> Result<u32> {
    let mut text_units = None;
    for record in records {
        let units = match record.record_type {
            RecordType::TextCharsAtom => {
                if record.version != 0
                    || record.instance != 0
                    || record.data.len() % 2 != 0
                    || usize::try_from(record.data_length).ok() != Some(record.data.len())
                {
                    return corrupted("TextCharsAtom has an invalid header or UTF-16 length");
                }
                u32::try_from(record.data.len() / 2)
                    .map_err(|_| Error::Corrupted("TextCharsAtom is too large".to_string()))?
            },
            RecordType::TextBytesAtom => {
                if record.version != 0
                    || record.instance != 0
                    || usize::try_from(record.data_length).ok() != Some(record.data.len())
                {
                    return corrupted("TextBytesAtom has an invalid header or length");
                }
                u32::try_from(record.data.len())
                    .map_err(|_| Error::Corrupted("TextBytesAtom is too large".to_string()))?
            },
            _ => continue,
        };
        if text_units.replace(units).is_some() {
            return corrupted("Text body has multiple character atoms");
        }
    }
    Ok(text_units.unwrap_or(0))
}

fn exact_text_from_records<'a>(records: impl IntoIterator<Item = &'a Record>) -> Result<String> {
    let mut text = None;
    for record in records {
        let value = match record.record_type {
            RecordType::TextCharsAtom => {
                if record.data.len() % 2 != 0 {
                    return corrupted("TextCharsAtom has an odd UTF-16 byte length");
                }
                let units = record
                    .data
                    .chunks_exact(2)
                    .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                    .collect::<Vec<_>>();
                String::from_utf16(&units).map_err(|_| {
                    Error::Corrupted("TextCharsAtom contains invalid UTF-16".to_string())
                })?
            },
            RecordType::TextBytesAtom => record.data.iter().map(|byte| char::from(*byte)).collect(),
            _ => continue,
        };
        if text.replace(value).is_some() {
            return corrupted("Text body has multiple character atoms");
        }
    }
    Ok(text.unwrap_or_default())
}

fn parse_anchor(
    record: &Record,
    trigger: InteractionTrigger,
    text_units: u32,
) -> Result<TextRange> {
    if record.record_type != RecordType::TextInteractiveInfoAtom
        || record.record_type_raw != RecordType::TextInteractiveInfoAtom as u16
        || record.version != 0
        || record.instance != trigger_instance(trigger)
        || record.data_length != 8
        || record.data.len() != 8
    {
        return corrupted(
            "TextInteractiveInfoAtom has an invalid header, trigger instance, or length",
        );
    }
    let begin = i32::from_le_bytes(record.data[0..4].try_into().unwrap());
    let end = i32::from_le_bytes(record.data[4..8].try_into().unwrap());
    let begin = u32::try_from(begin)
        .map_err(|_| Error::Corrupted("Text interaction begin position is negative".into()))?;
    let end = u32::try_from(end)
        .map_err(|_| Error::Corrupted("Text interaction end position is negative".into()))?;
    let range = TextRange::new(begin, end)?;
    range.validate_for_text(text_units)?;
    Ok(range)
}

const fn trigger_instance(trigger: InteractionTrigger) -> u16 {
    match trigger {
        InteractionTrigger::Click => 0,
        InteractionTrigger::MouseOver => 1,
    }
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{InteractionAction, InteractionLinkTarget};

    fn record(version: u16, instance: u16, kind: RecordType, data: &[u8]) -> Record {
        let bytes = encode_record(version, instance, kind as u16, data).unwrap();
        let (record, consumed) = Record::parse_strict(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        record
    }

    fn pair(trigger: InteractionTrigger, range: TextRange) -> TextInteraction {
        TextInteraction::new(
            range,
            Interaction::new(
                trigger,
                InteractionAction::Hyperlink,
                InteractionLinkTarget::Url,
            ),
        )
        .unwrap()
    }

    #[test]
    fn paired_click_and_hover_round_trip_exactly() {
        let values = [
            pair(InteractionTrigger::Click, TextRange::new(1, 3).unwrap()),
            pair(InteractionTrigger::MouseOver, TextRange::new(3, 5).unwrap()),
        ];
        let bytes = values
            .iter()
            .flat_map(|value| {
                value
                    .to_bytes_for_text(5, TextInteractionLimits::default())
                    .unwrap()
            })
            .collect::<Vec<_>>();
        let records = Record::parse_sequence_strict(&bytes, "text interactions").unwrap();

        assert_eq!(
            TextInteraction::parse_records(&records, 5, TextInteractionLimits::default()).unwrap(),
            values
        );
    }

    #[test]
    fn rejects_orphans_trigger_mismatch_ranges_and_limits() {
        let click = pair(InteractionTrigger::Click, TextRange::new(1, 3).unwrap());
        let action = click.interaction.to_record().unwrap();
        let hover_anchor = record(
            0,
            1,
            RecordType::TextInteractiveInfoAtom,
            &[1, 0, 0, 0, 3, 0, 0, 0],
        );
        assert!(
            TextInteraction::parse_records([&action, &hover_anchor], 3, Default::default())
                .is_err()
        );
        assert!(TextInteraction::parse_records([&hover_anchor], 3, Default::default()).is_err());
        assert!(click.validate_for_text(2, Default::default()).is_ok());
        assert!(click.validate_for_text(1, Default::default()).is_err());
        assert!(
            click
                .validate_for_text(
                    3,
                    TextInteractionLimits {
                        max_text_units: 2,
                        ..Default::default()
                    }
                )
                .is_err()
        );
        assert!(
            TextInteraction::parse_records(
                [
                    &action,
                    &record(
                        0,
                        0,
                        RecordType::TextInteractiveInfoAtom,
                        &[1, 0, 0, 0, 3, 0, 0, 0],
                    )
                ],
                3,
                TextInteractionLimits {
                    max_interactions: 0,
                    ..Default::default()
                }
            )
            .is_err()
        );

        let click_anchor = record(
            0,
            0,
            RecordType::TextInteractiveInfoAtom,
            &[1, 0, 0, 0, 3, 0, 0, 0],
        );
        let style = record(0, 0, RecordType::StyleTextPropAtom, &[]);
        assert!(
            TextInteraction::parse_records([&action, &click_anchor, &style], 3, Default::default())
                .is_err()
        );
        let ruler = record(0, 0, RecordType::TextRulerAtom, &[]);
        assert!(
            TextInteraction::parse_records([&action, &click_anchor, &ruler], 3, Default::default())
                .is_ok()
        );
        assert!(
            TextInteraction::parse_records(
                [&action, &click_anchor, &ruler, &action, &click_anchor],
                3,
                Default::default()
            )
            .is_err()
        );
    }

    #[test]
    fn slide_list_bodies_validate_text_type_utf16_and_range_pairing() {
        let header = record(
            0,
            2,
            RecordType::TextHeaderAtom,
            &TextType::Body.value().to_le_bytes(),
        );
        let text = record(
            0,
            0,
            RecordType::TextCharsAtom,
            &[b'A', 0, 0x3D, 0xD8, 0x00, 0xDE],
        );
        let pair = pair(InteractionTrigger::Click, TextRange::new(1, 3).unwrap());
        let pair_bytes = pair.to_bytes_for_text(3, Default::default()).unwrap();
        let pair_records = Record::parse_sequence_strict(&pair_bytes, "text interaction").unwrap();
        let records = [&header, &text, &pair_records[0], &pair_records[1]];

        let parsed = parse_text_bodies(&records, Default::default()).unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].text_header_instance, 2);
        assert_eq!(parsed[0].text_type, TextType::Body);
        assert_eq!(parsed[0].text, "A😀");
        assert_eq!(parsed[0].interactions, [pair]);
    }
}
