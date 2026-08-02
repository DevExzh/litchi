//! Hyperlink definitions and PowerPoint 9 hyperlink extensions.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;
use std::borrow::Cow;

/// Mouse event that triggers an interactive action.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InteractionTrigger {
    /// Mouse click.
    Click,
    /// Mouse pointer moved over the object.
    MouseOver,
}

/// Action stored in an `InteractiveInfoAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InteractionAction {
    NoAction,
    Macro,
    RunProgram,
    Jump,
    Hyperlink,
    Ole,
    Media,
    CustomShow,
}

/// Relative slide-show jump target.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InteractionJump {
    None,
    NextSlide,
    PreviousSlide,
    FirstSlide,
    LastSlide,
    LastSlideViewed,
    EndShow,
}

/// Interpretation of an interactive hyperlink reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InteractionLinkTarget {
    NextSlide,
    PreviousSlide,
    FirstSlide,
    LastSlide,
    CustomShow,
    SlideNumber,
    Url,
    OtherPresentation,
    OtherFile,
    Nil,
}

/// Resource limits for an interactive-information container.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointInteractionLimits {
    /// Maximum complete container size, including its eight-byte header.
    pub max_record_bytes: usize,
    /// Maximum MacroNameAtom UTF-16 payload size.
    pub max_macro_name_bytes: usize,
}

impl Default for PowerPointInteractionLimits {
    fn default() -> Self {
        Self {
            max_record_bytes: 1024 * 1024,
            max_macro_name_bytes: 64 * 1024,
        }
    }
}

/// Typed payload of an MS-PPT §2.6.10 InteractiveInfoAtom.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PowerPointInteractiveInfoAtom {
    pub sound_id: u32,
    pub hyperlink_id: u32,
    pub action: InteractionAction,
    pub ole_verb: u8,
    pub jump: InteractionJump,
    pub animated: bool,
    pub stop_sound: bool,
    pub custom_show_return: bool,
    pub visited: bool,
    pub link_target: InteractionLinkTarget,
    /// Undefined bytes retained without interpretation.
    pub unused: [u8; 3],
}

impl PowerPointInteractiveInfoAtom {
    /// Parse the exact sixteen-byte atom payload.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if data.len() != 16 {
            return corrupted("InteractiveInfoAtom payload must be exactly 16 bytes");
        }
        let flags = data[11];
        if flags & 0xF0 != 0 {
            return corrupted("InteractiveInfoAtom reserved flag bits must be zero");
        }
        Ok(Self {
            sound_id: u32::from_le_bytes(data[0..4].try_into().unwrap()),
            hyperlink_id: u32::from_le_bytes(data[4..8].try_into().unwrap()),
            action: parse_action(data[8])?,
            ole_verb: data[9],
            jump: parse_jump(data[10])?,
            animated: flags & 0x01 != 0,
            stop_sound: flags & 0x02 != 0,
            custom_show_return: flags & 0x04 != 0,
            visited: flags & 0x08 != 0,
            link_target: parse_link_target(data[12])?,
            unused: data[13..16].try_into().unwrap(),
        })
    }

    /// Serialize the exact sixteen-byte payload.
    pub fn to_payload(self) -> [u8; 16] {
        let mut data = [0u8; 16];
        data[0..4].copy_from_slice(&self.sound_id.to_le_bytes());
        data[4..8].copy_from_slice(&self.hyperlink_id.to_le_bytes());
        data[8] = action_value(self.action);
        data[9] = self.ole_verb;
        data[10] = jump_value(self.jump);
        data[11] = u8::from(self.animated)
            | (u8::from(self.stop_sound) << 1)
            | (u8::from(self.custom_show_return) << 2)
            | (u8::from(self.visited) << 3);
        data[12] = link_target_value(self.link_target);
        data[13..16].copy_from_slice(&self.unused);
        data
    }
}

/// Inert MS-PPT §2.6.11 MacroNameAtom data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointMacroNameAtom {
    text: String,
    raw_utf16: Vec<u8>,
}

impl PowerPointMacroNameAtom {
    /// Construct canonical non-terminated printable UTF-16 data.
    pub fn new(text: impl Into<String>) -> Result<Self> {
        let text = text.into();
        validate_printable_text(&text)?;
        let mut raw_utf16 = Vec::with_capacity(text.encode_utf16().count().saturating_mul(2));
        for unit in text.encode_utf16() {
            raw_utf16.extend_from_slice(&unit.to_le_bytes());
        }
        Ok(Self { text, raw_utf16 })
    }

    /// Parse and retain exact UTF-16 bytes. A NULL code unit terminates the exposed text.
    pub fn parse_payload(data: &[u8], max_bytes: usize) -> Result<Self> {
        let text = decode_macro_name(data, max_bytes)?;
        Ok(Self {
            text,
            raw_utf16: data.to_vec(),
        })
    }

    /// Inert macro, file, or named-show text. This accessor never executes it.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Exact original UTF-16 bytes, including a terminator or ignored suffix if present.
    pub fn raw_utf16(&self) -> &[u8] {
        &self.raw_utf16
    }
}

/// One click or mouse-over action attached to a shape or text range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointInteraction {
    pub trigger: InteractionTrigger,
    pub sound_id: u32,
    pub hyperlink_id: u32,
    pub action: InteractionAction,
    pub ole_verb: u8,
    pub jump: InteractionJump,
    pub animated: bool,
    pub stop_sound: bool,
    pub custom_show_return: bool,
    pub visited: bool,
    pub link_target: InteractionLinkTarget,
    pub macro_name: Option<String>,
    /// Undefined atom bytes retained verbatim.
    pub unused: [u8; 3],
    /// Exact inert MacroNameAtom UTF-16 data, if present.
    pub macro_name_data: Option<Vec<u8>>,
}

/// Click and mouse-over actions attached to one slide shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointShapeInteractionEntry {
    /// OfficeArt shape identifier.
    pub shape_id: u32,
    /// At most one action for each [`InteractionTrigger`].
    pub interactions: Vec<PowerPointInteraction>,
}

impl PowerPointInteraction {
    /// Construct a canonical interaction with zero references and unused bytes.
    pub fn new(
        trigger: InteractionTrigger,
        action: InteractionAction,
        link_target: InteractionLinkTarget,
    ) -> Self {
        Self {
            trigger,
            sound_id: 0,
            hyperlink_id: 0,
            action,
            ole_verb: 0,
            jump: InteractionJump::None,
            animated: false,
            stop_sound: false,
            custom_show_return: false,
            visited: false,
            link_target,
            macro_name: None,
            unused: [0; 3],
            macro_name_data: None,
        }
    }

    /// Attach inert macro/file/show name data. No macro activation is performed.
    pub fn with_macro_name(mut self, value: impl Into<String>) -> Result<Self> {
        let atom = PowerPointMacroNameAtom::new(value)?;
        self.macro_name = Some(atom.text().to_string());
        self.macro_name_data = Some(atom.raw_utf16().to_vec());
        Ok(self)
    }

    /// Play one typed built-in sound when the interaction executes.
    ///
    /// The writer resolves this catalog ID to the document's emitted sound
    /// identifier. Audio remains inert and is never played by the library.
    pub fn with_builtin_sound(mut self, sound: crate::animation::BuiltinSound) -> Self {
        self.sound_id = sound.id();
        self.stop_sound = false;
        self
    }

    /// Bind this interaction to an explicitly registered embedded sound.
    pub fn with_sound_reference(mut self, sound_id: std::num::NonZeroU32) -> Self {
        self.sound_id = sound_id.get();
        self.stop_sound = false;
        self
    }

    /// Parse an `InteractiveInfo` click or mouse-over container.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        Self::parse_with_limits(record, PowerPointInteractionLimits::default())
    }

    /// Parse a container with explicit record and MacroNameAtom bounds.
    pub fn parse_with_limits(
        record: &PptRecord,
        limits: PowerPointInteractionLimits,
    ) -> Result<Self> {
        if record.record_type != PptRecordType::InteractiveInfo
            || record.record_type_raw != PptRecordType::InteractiveInfo.as_u16()
            || record.version != 0x0f
            || usize::try_from(record.data_length).ok() != Some(record.data.len())
        {
            return corrupted("InteractiveInfo has an invalid record header or length");
        }
        if record.data.len().saturating_add(8) > limits.max_record_bytes {
            return corrupted("InteractiveInfo exceeds the configured record size limit");
        }
        let trigger = match record.instance {
            0 => InteractionTrigger::Click,
            1 => InteractionTrigger::MouseOver,
            _ => return corrupted("InteractiveInfo has an invalid trigger instance"),
        };
        let children = PptRecord::parse_sequence_strict(&record.data, "interactive information")?;
        if !matches!(children.len(), 1 | 2) {
            return corrupted("InteractiveInfo has an invalid child count");
        }
        let atom = &children[0];
        if atom.record_type != PptRecordType::InteractiveInfoAtom
            || atom.record_type_raw != PptRecordType::InteractiveInfoAtom.as_u16()
            || atom.version != 0
            || atom.instance != 0
            || atom.data_length != 16
            || atom.data.len() != 16
        {
            return corrupted("InteractiveInfoAtom has an invalid header or size");
        }
        let parsed_atom = PowerPointInteractiveInfoAtom::parse_payload(&atom.data)?;
        let macro_atom = if children.len() == 2 {
            let name = &children[1];
            if name.record_type != PptRecordType::CString
                || name.record_type_raw != PptRecordType::CString.as_u16()
                || name.version != 0
                || name.instance != 2
                || usize::try_from(name.data_length).ok() != Some(name.data.len())
            {
                return corrupted("MacroNameAtom has an invalid record header or length");
            }
            Some(PowerPointMacroNameAtom::parse_payload(
                &name.data,
                limits.max_macro_name_bytes,
            )?)
        } else {
            None
        };
        Ok(Self {
            trigger,
            sound_id: parsed_atom.sound_id,
            hyperlink_id: parsed_atom.hyperlink_id,
            action: parsed_atom.action,
            ole_verb: parsed_atom.ole_verb,
            jump: parsed_atom.jump,
            animated: parsed_atom.animated,
            stop_sound: parsed_atom.stop_sound,
            custom_show_return: parsed_atom.custom_show_return,
            visited: parsed_atom.visited,
            link_target: parsed_atom.link_target,
            macro_name: macro_atom.as_ref().map(|atom| atom.text().to_string()),
            unused: parsed_atom.unused,
            macro_name_data: macro_atom.map(|atom| atom.raw_utf16),
        })
    }

    /// Parse one exact complete container from bytes.
    pub fn parse_bytes(bytes: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_limits(bytes, PowerPointInteractionLimits::default())
    }

    /// Parse one exact complete container from bounded bytes.
    pub fn parse_bytes_with_limits(
        bytes: &[u8],
        limits: PowerPointInteractionLimits,
    ) -> Result<Self> {
        if bytes.len() > limits.max_record_bytes {
            return corrupted("InteractiveInfo exceeds the configured record size limit");
        }
        let (record, consumed) = PptRecord::parse_strict(bytes, 0)?;
        if consumed != bytes.len() {
            return corrupted("InteractiveInfo has trailing bytes");
        }
        Self::parse_with_limits(&record, limits)
    }

    /// Return the typed required atom.
    pub fn atom(&self) -> PowerPointInteractiveInfoAtom {
        PowerPointInteractiveInfoAtom {
            sound_id: self.sound_id,
            hyperlink_id: self.hyperlink_id,
            action: self.action,
            ole_verb: self.ole_verb,
            jump: self.jump,
            animated: self.animated,
            stop_sound: self.stop_sound,
            custom_show_return: self.custom_show_return,
            visited: self.visited,
            link_target: self.link_target,
            unused: self.unused,
        }
    }

    /// Return the optional inert name atom with its exact bytes.
    pub fn macro_name_atom(&self) -> Result<Option<PowerPointMacroNameAtom>> {
        match (&self.macro_name, &self.macro_name_data) {
            (None, None) => Ok(None),
            (Some(text), Some(data)) => {
                let atom = PowerPointMacroNameAtom::parse_payload(data, usize::MAX)?;
                if atom.text() != text {
                    return corrupted("MacroNameAtom text and exact data disagree");
                }
                Ok(Some(atom))
            },
            (Some(text), None) => Ok(Some(PowerPointMacroNameAtom::new(text.clone())?)),
            (None, Some(_)) => corrupted("MacroNameAtom exact data is present without text"),
        }
    }

    /// Validate this interaction against explicit resource limits.
    pub fn validate_with_limits(&self, limits: PowerPointInteractionLimits) -> Result<()> {
        let macro_name = self.validated_macro_name_data(limits)?;
        validate_serialized_interaction_size(
            macro_name.as_ref().map(|data| data.len()),
            limits.max_record_bytes,
        )?;
        Ok(())
    }

    /// Serialize canonical headers and exact atom/name payloads with default limits.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_bytes_with_limits(PowerPointInteractionLimits::default())
    }

    /// Serialize canonical headers and exact atom/name payloads with explicit limits.
    pub fn to_bytes_with_limits(&self, limits: PowerPointInteractionLimits) -> Result<Vec<u8>> {
        let macro_name = self.validated_macro_name_data(limits)?;
        validate_serialized_interaction_size(
            macro_name.as_ref().map(|data| data.len()),
            limits.max_record_bytes,
        )?;
        let atom_payload = self.atom().to_payload();
        let mut children = encode_record(
            0,
            0,
            PptRecordType::InteractiveInfoAtom.as_u16(),
            &atom_payload,
        )?;
        if let Some(name) = macro_name {
            children.extend_from_slice(&encode_record(
                0,
                2,
                PptRecordType::CString.as_u16(),
                &name,
            )?);
        }
        let instance = match self.trigger {
            InteractionTrigger::Click => 0,
            InteractionTrigger::MouseOver => 1,
        };
        encode_record(
            0x0F,
            instance,
            PptRecordType::InteractiveInfo.as_u16(),
            &children,
        )
    }

    fn validated_macro_name_data(
        &self,
        limits: PowerPointInteractionLimits,
    ) -> Result<Option<Cow<'_, [u8]>>> {
        match (&self.macro_name, &self.macro_name_data) {
            (None, None) => Ok(None),
            (Some(text), Some(data)) => {
                let parsed = decode_macro_name(data, limits.max_macro_name_bytes)?;
                if &parsed != text {
                    return corrupted("MacroNameAtom text and exact data disagree");
                }
                Ok(Some(Cow::Borrowed(data)))
            },
            (Some(text), None) => {
                let atom = PowerPointMacroNameAtom::new(text.clone())?;
                if atom.raw_utf16().len() > limits.max_macro_name_bytes {
                    return corrupted("MacroNameAtom exceeds the configured size limit");
                }
                Ok(Some(Cow::Owned(atom.raw_utf16)))
            },
            (None, Some(_)) => corrupted("MacroNameAtom exact data is present without text"),
        }
    }

    /// Convert to the generic record model used by existing shape extraction.
    pub fn to_record(&self) -> Result<PptRecord> {
        let bytes = self.to_bytes()?;
        let (record, consumed) = PptRecord::parse_strict(&bytes, 0)?;
        if consumed != bytes.len() {
            return corrupted("serialized InteractiveInfo was only partially parsed");
        }
        Ok(record)
    }

    /// Resolve this action's hyperlink reference.
    pub fn hyperlink<'a>(
        &self,
        hyperlinks: &'a PowerPointHyperlinks,
    ) -> Option<&'a PowerPointHyperlink> {
        hyperlinks.get(self.hyperlink_id)
    }

    /// Resolve this action's embedded sound reference.
    pub fn sound<'collection, 'data>(
        &self,
        sounds: &'collection crate::PowerPointSoundCollection<'data>,
    ) -> Option<&'collection crate::EmbeddedPowerPointSound<'data>> {
        sounds.get(self.sound_id)
    }

    /// Validate this action's non-null sound reference.
    pub fn validate_sound_collection(
        &self,
        sounds: &crate::PowerPointSoundCollection<'_>,
    ) -> Result<()> {
        if self.sound_id != 0 && sounds.get(self.sound_id).is_none() {
            return corrupted(format!(
                "interaction references missing sound ID {}",
                self.sound_id
            ));
        }
        Ok(())
    }

    pub(crate) fn parse_client_data_payload(
        data: &[u8],
        limits: PowerPointInteractionLimits,
    ) -> Result<Vec<Self>> {
        let mut interactions = Vec::new();
        for record in PptRecord::parse_sequence_strict(data, "shape ClientData")? {
            if record.record_type != PptRecordType::InteractiveInfo {
                continue;
            }
            let interaction = Self::parse_with_limits(&record, limits)?;
            if interactions
                .iter()
                .any(|existing: &Self| existing.trigger == interaction.trigger)
            {
                return corrupted("Shape contains duplicate interactive triggers");
            }
            interactions.push(interaction);
        }
        Ok(interactions)
    }
}

fn validate_serialized_interaction_size(
    macro_name_bytes: Option<usize>,
    max_record_bytes: usize,
) -> Result<()> {
    let macro_record_len = macro_name_bytes
        .map(|size| {
            8usize
                .checked_add(size)
                .ok_or_else(|| PptError::Corrupted("InteractiveInfo length overflows".into()))
        })
        .transpose()?
        .unwrap_or(0);
    let complete_len = 32usize
        .checked_add(macro_record_len)
        .ok_or_else(|| PptError::Corrupted("InteractiveInfo length overflows".into()))?;
    if complete_len > max_record_bytes {
        return corrupted("InteractiveInfo exceeds the configured record size limit");
    }
    Ok(())
}

fn action_value(value: InteractionAction) -> u8 {
    match value {
        InteractionAction::NoAction => 0,
        InteractionAction::Macro => 1,
        InteractionAction::RunProgram => 2,
        InteractionAction::Jump => 3,
        InteractionAction::Hyperlink => 4,
        InteractionAction::Ole => 5,
        InteractionAction::Media => 6,
        InteractionAction::CustomShow => 7,
    }
}

fn jump_value(value: InteractionJump) -> u8 {
    match value {
        InteractionJump::None => 0,
        InteractionJump::NextSlide => 1,
        InteractionJump::PreviousSlide => 2,
        InteractionJump::FirstSlide => 3,
        InteractionJump::LastSlide => 4,
        InteractionJump::LastSlideViewed => 5,
        InteractionJump::EndShow => 6,
    }
}

fn link_target_value(value: InteractionLinkTarget) -> u8 {
    match value {
        InteractionLinkTarget::NextSlide => 0,
        InteractionLinkTarget::PreviousSlide => 1,
        InteractionLinkTarget::FirstSlide => 2,
        InteractionLinkTarget::LastSlide => 3,
        InteractionLinkTarget::CustomShow => 6,
        InteractionLinkTarget::SlideNumber => 7,
        InteractionLinkTarget::Url => 8,
        InteractionLinkTarget::OtherPresentation => 9,
        InteractionLinkTarget::OtherFile => 10,
        InteractionLinkTarget::Nil => 0xFF,
    }
}

fn parse_action(value: u8) -> Result<InteractionAction> {
    match value {
        0 => Ok(InteractionAction::NoAction),
        1 => Ok(InteractionAction::Macro),
        2 => Ok(InteractionAction::RunProgram),
        3 => Ok(InteractionAction::Jump),
        4 => Ok(InteractionAction::Hyperlink),
        5 => Ok(InteractionAction::Ole),
        6 => Ok(InteractionAction::Media),
        7 => Ok(InteractionAction::CustomShow),
        _ => Err(PptError::Corrupted(
            "Invalid interactive action".to_string(),
        )),
    }
}

fn parse_jump(value: u8) -> Result<InteractionJump> {
    match value {
        0 => Ok(InteractionJump::None),
        1 => Ok(InteractionJump::NextSlide),
        2 => Ok(InteractionJump::PreviousSlide),
        3 => Ok(InteractionJump::FirstSlide),
        4 => Ok(InteractionJump::LastSlide),
        5 => Ok(InteractionJump::LastSlideViewed),
        6 => Ok(InteractionJump::EndShow),
        _ => Err(PptError::Corrupted("Invalid interactive jump".to_string())),
    }
}

fn parse_link_target(value: u8) -> Result<InteractionLinkTarget> {
    match value {
        0 => Ok(InteractionLinkTarget::NextSlide),
        1 => Ok(InteractionLinkTarget::PreviousSlide),
        2 => Ok(InteractionLinkTarget::FirstSlide),
        3 => Ok(InteractionLinkTarget::LastSlide),
        6 => Ok(InteractionLinkTarget::CustomShow),
        7 => Ok(InteractionLinkTarget::SlideNumber),
        8 => Ok(InteractionLinkTarget::Url),
        9 => Ok(InteractionLinkTarget::OtherPresentation),
        10 => Ok(InteractionLinkTarget::OtherFile),
        0xff => Ok(InteractionLinkTarget::Nil),
        _ => Err(PptError::Corrupted(
            "Invalid hyperlink target type".to_string(),
        )),
    }
}

fn validate_printable_text(value: &str) -> Result<()> {
    if value
        .chars()
        .any(|ch| matches!(ch as u32, 0x0000..=0x001F | 0x007F..=0x009F))
    {
        return corrupted("MacroNameAtom contains a forbidden control character");
    }
    Ok(())
}

fn decode_macro_name(data: &[u8], max_bytes: usize) -> Result<String> {
    if data.len() > max_bytes {
        return corrupted("MacroNameAtom exceeds the configured size limit");
    }
    if !data.len().is_multiple_of(2) {
        return corrupted("MacroNameAtom length must be even");
    }
    let units = data
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    let visible = units
        .iter()
        .position(|unit| *unit == 0)
        .unwrap_or(units.len());
    let text = String::from_utf16(&units[..visible])
        .map_err(|_| PptError::Corrupted("MacroNameAtom contains invalid UTF-16".into()))?;
    validate_printable_text(&text)?;
    Ok(text)
}

pub(crate) fn encode_record(
    version: u16,
    instance: u16,
    record_type: u16,
    data: &[u8],
) -> Result<Vec<u8>> {
    let length = u32::try_from(data.len())
        .map_err(|_| PptError::Corrupted("PowerPoint record payload exceeds u32".into()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod interaction_protocol_tests {
    use super::*;

    fn record(version: u16, instance: u16, kind: PptRecordType, data: &[u8]) -> Vec<u8> {
        encode_record(version, instance, kind.as_u16(), data).unwrap()
    }

    fn interaction(instance: u16, atom: &[u8], macro_data: Option<&[u8]>) -> Vec<u8> {
        let mut children = record(0, 0, PptRecordType::InteractiveInfoAtom, atom);
        if let Some(data) = macro_data {
            children.extend(record(0, 2, PptRecordType::CString, data));
        }
        record(0x0F, instance, PptRecordType::InteractiveInfo, &children)
    }

    fn atom() -> [u8; 16] {
        PowerPointInteractiveInfoAtom {
            sound_id: 7,
            hyperlink_id: 11,
            action: InteractionAction::Macro,
            ole_verb: 19,
            jump: InteractionJump::LastSlideViewed,
            animated: true,
            stop_sound: true,
            custom_show_return: true,
            visited: true,
            link_target: InteractionLinkTarget::OtherFile,
            unused: [0xAA, 0xBB, 0xCC],
        }
        .to_payload()
    }

    #[test]
    fn exact_round_trip_preserves_trigger_macro_bytes_and_undefined_data() {
        let macro_data = [b'R', 0, b'u', 0, b'n', 0, 0, 0, 1, 0];
        let bytes = interaction(1, &atom(), Some(&macro_data));
        let parsed = PowerPointInteraction::parse_bytes(&bytes).unwrap();
        assert_eq!(parsed.trigger, InteractionTrigger::MouseOver);
        assert_eq!(parsed.macro_name.as_deref(), Some("Run"));
        assert_eq!(
            parsed.macro_name_atom().unwrap().unwrap().raw_utf16(),
            macro_data
        );
        assert_eq!(parsed.unused, [0xAA, 0xBB, 0xCC]);
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn canonical_constructor_and_record_accessor_round_trip() {
        let mut value = PowerPointInteraction::new(
            InteractionTrigger::Click,
            InteractionAction::CustomShow,
            InteractionLinkTarget::CustomShow,
        )
        .with_macro_name("Quarterly show")
        .unwrap();
        value.hyperlink_id = 42;
        value.custom_show_return = true;
        let bytes = value.to_bytes().unwrap();
        assert_eq!(
            PowerPointInteraction::parse(&value.to_record().unwrap()).unwrap(),
            value
        );
        assert_eq!(PowerPointInteraction::parse_bytes(&bytes).unwrap(), value);
    }

    #[test]
    fn preserves_ignored_macro_name_without_activating_it() {
        let bytes = interaction(
            0,
            &PowerPointInteractiveInfoAtom {
                action: InteractionAction::Hyperlink,
                ..PowerPointInteraction::new(
                    InteractionTrigger::Click,
                    InteractionAction::Hyperlink,
                    InteractionLinkTarget::Url,
                )
                .atom()
            }
            .to_payload(),
            Some(&[b'X', 0]),
        );
        let parsed = PowerPointInteraction::parse_bytes(&bytes).unwrap();
        assert_eq!(parsed.macro_name.as_deref(), Some("X"));
        assert_eq!(parsed.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn rejects_bad_container_atom_and_child_order() {
        let valid_atom = atom();
        for instance in [2u16, 0x0FFF] {
            assert!(
                PowerPointInteraction::parse_bytes(&interaction(instance, &valid_atom, None))
                    .is_err()
            );
        }
        assert!(
            PowerPointInteraction::parse_bytes(&record(0, 0, PptRecordType::InteractiveInfo, &[]))
                .is_err()
        );
        assert!(
            PowerPointInteraction::parse_bytes(&interaction(0, &valid_atom[..15], None)).is_err()
        );

        let name = record(0, 2, PptRecordType::CString, &[b'A', 0]);
        let atom_record = record(0, 0, PptRecordType::InteractiveInfoAtom, &valid_atom);
        assert!(
            PowerPointInteraction::parse_bytes(&record(
                0x0F,
                0,
                PptRecordType::InteractiveInfo,
                &[name, atom_record].concat()
            ))
            .is_err()
        );
    }

    #[test]
    fn rejects_enum_reserved_and_printable_string_violations() {
        for (offset, value) in [(8usize, 8u8), (10, 7), (12, 4)] {
            let mut bad = atom();
            bad[offset] = value;
            assert!(PowerPointInteraction::parse_bytes(&interaction(0, &bad, None)).is_err());
        }
        let mut reserved = atom();
        reserved[11] |= 0x10;
        assert!(PowerPointInteraction::parse_bytes(&interaction(0, &reserved, None)).is_err());
        assert!(PowerPointInteraction::parse_bytes(&interaction(0, &atom(), Some(&[1]))).is_err());
        assert!(
            PowerPointInteraction::parse_bytes(&interaction(0, &atom(), Some(&[1, 0]))).is_err()
        );
        assert!(
            PowerPointInteraction::parse_bytes(&interaction(0, &atom(), Some(&[0, 0xD8]))).is_err()
        );
    }

    #[test]
    fn enforces_record_and_macro_limits_and_exact_consumption() {
        let bytes = interaction(0, &atom(), Some(&[b'A', 0, b'B', 0]));
        assert!(
            PowerPointInteraction::parse_bytes_with_limits(
                &bytes,
                PowerPointInteractionLimits {
                    max_record_bytes: bytes.len() - 1,
                    ..Default::default()
                }
            )
            .is_err()
        );
        assert!(
            PowerPointInteraction::parse_bytes_with_limits(
                &bytes,
                PowerPointInteractionLimits {
                    max_macro_name_bytes: 2,
                    ..Default::default()
                }
            )
            .is_err()
        );
        let mut trailing = bytes;
        trailing.push(0);
        assert!(PowerPointInteraction::parse_bytes(&trailing).is_err());

        let value = PowerPointInteraction::new(
            InteractionTrigger::Click,
            InteractionAction::NoAction,
            InteractionLinkTarget::Nil,
        );
        let limits = PowerPointInteractionLimits {
            max_record_bytes: 31,
            ..Default::default()
        };
        assert!(value.validate_with_limits(limits).is_err());
        assert!(value.to_bytes_with_limits(limits).is_err());
    }
}

/// Additional hyperlink data introduced by PowerPoint 9.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointHyperlinkExtension {
    /// Optional text displayed as a hover screen tip.
    pub screen_tip: Option<String>,
    /// Whether the hyperlink was created in the Insert Hyperlink dialog.
    pub inserted_with_dialog: bool,
    /// Whether the base hyperlink location names a custom slide show.
    pub location_is_named_show: bool,
    /// Whether a named show returns to the originating slide.
    pub named_show_returns_to_slide: bool,
}

impl PowerPointHyperlinkExtension {
    /// Parse an `ExHyperlink9Container` record and return its referenced ID.
    pub fn parse(record: &PptRecord) -> Result<(u32, Self)> {
        if record.record_type != PptRecordType::ExternalHyperlink9
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(PptError::Corrupted(
                "ExHyperlink9Container has an invalid record header".to_string(),
            ));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "PowerPoint 9 hyperlink")?;
        if !matches!(children.len(), 2 | 3) {
            return Err(PptError::Corrupted(
                "ExHyperlink9Container has an invalid child count".to_string(),
            ));
        }
        let reference = parse_hyperlink_atom(&children[0])?;
        if reference == 0 {
            return Err(PptError::Corrupted(
                "ExHyperlink9Container has a null hyperlink reference".to_string(),
            ));
        }
        let (screen_tip, flags_index) = if children.len() == 3 {
            let tip = &children[1];
            if tip.record_type != PptRecordType::CString || tip.version != 0 || tip.instance != 0 {
                return Err(PptError::Corrupted(
                    "ScreenTipAtom has an invalid record header".to_string(),
                ));
            }
            (Some(parse_unicode_string(&tip.data)?), 2)
        } else {
            (None, 1)
        };
        let flags = &children[flags_index];
        if flags.record_type != PptRecordType::ExternalHyperlinkFlagsAtom
            || flags.version != 0
            || flags.instance != 0
            || flags.data.len() != 4
        {
            return Err(PptError::Corrupted(
                "ExHyperlinkFlagsAtom has an invalid record header or size".to_string(),
            ));
        }
        let value =
            u32::from_le_bytes([flags.data[0], flags.data[1], flags.data[2], flags.data[3]]);
        if value & !0x07 != 0 {
            return Err(PptError::Corrupted(
                "ExHyperlinkFlagsAtom has nonzero reserved bits".to_string(),
            ));
        }
        Ok((
            reference,
            Self {
                screen_tip,
                inserted_with_dialog: value & 0x01 != 0,
                location_is_named_show: value & 0x02 != 0,
                named_show_returns_to_slide: value & 0x04 != 0,
            },
        ))
    }
}

/// One base PowerPoint hyperlink definition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointHyperlink {
    /// Positive identifier referenced by interactive information records.
    pub id: u32,
    /// Optional user-readable hyperlink name.
    pub friendly_name: Option<String>,
    /// Optional full destination-file path or URL.
    pub target: Option<String>,
    /// Optional location within the destination.
    pub location: Option<String>,
    /// Optional PowerPoint 9 metadata for this hyperlink.
    pub extension: Option<PowerPointHyperlinkExtension>,
}

impl PowerPointHyperlink {
    /// Parse an `ExHyperlinkContainer` record.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::ExternalHyperlink
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(PptError::Corrupted(
                "ExHyperlinkContainer has an invalid record header".to_string(),
            ));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "external hyperlink")?;
        let Some(atom) = children.first() else {
            return Err(PptError::Corrupted(
                "ExHyperlinkContainer is missing ExHyperlinkAtom".to_string(),
            ));
        };
        let id = parse_hyperlink_atom(atom)?;
        if id == 0 {
            return Err(PptError::Corrupted(
                "ExHyperlinkAtom has a zero hyperlink ID".to_string(),
            ));
        }

        let mut friendly_name = None;
        let mut target = None;
        let mut location = None;
        let mut previous_instance = None;
        for child in &children[1..] {
            if child.record_type != PptRecordType::CString || child.version != 0 {
                return Err(PptError::Corrupted(
                    "ExHyperlinkContainer has an unexpected child record".to_string(),
                ));
            }
            if previous_instance.is_some_and(|previous| previous >= child.instance) {
                return Err(PptError::Corrupted(
                    "Hyperlink strings are duplicated or out of order".to_string(),
                ));
            }
            previous_instance = Some(child.instance);
            let value = Some(parse_unicode_string(&child.data)?);
            match child.instance {
                0 => friendly_name = value,
                1 => target = value,
                3 => location = value,
                _ => {
                    return Err(PptError::Corrupted(
                        "Hyperlink CString has an invalid record instance".to_string(),
                    ));
                },
            }
        }
        Ok(Self {
            id,
            friendly_name,
            target,
            location,
            extension: None,
        })
    }
}

/// Hyperlink definitions resolved with their PowerPoint 9 extensions.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointHyperlinks {
    /// Seed used when allocating new external-object or hyperlink identifiers.
    pub id_seed: Option<i32>,
    /// Hyperlinks in base `ExObjListContainer` order.
    pub hyperlinks: Vec<PowerPointHyperlink>,
}

impl PowerPointHyperlinks {
    /// Discover base hyperlinks and merge all `___PPT9` hyperlink extensions.
    pub fn parse(root: &PptRecord) -> Result<Self> {
        let mut lists = Vec::new();
        collect_records(root, PptRecordType::ExObjList, &mut lists);
        if lists.len() > 1 {
            return Err(PptError::Corrupted(
                "Record tree contains multiple external-object lists".to_string(),
            ));
        }
        let mut result = if let Some(list) = lists.first() {
            Self::parse_external_object_list(list)?
        } else {
            Self::default()
        };

        let mut extension_ids = Vec::new();
        for record in root.versioned_binary_tag_records(9)? {
            if record.record_type != PptRecordType::ExternalHyperlink9 {
                continue;
            }
            let (id, extension) = PowerPointHyperlinkExtension::parse(&record)?;
            if extension_ids.contains(&id) {
                return Err(PptError::Corrupted(
                    "PowerPoint 9 contains duplicate hyperlink extensions".to_string(),
                ));
            }
            extension_ids.push(id);
            let hyperlink = result.get_mut(id).ok_or_else(|| {
                PptError::Corrupted(
                    "PowerPoint 9 hyperlink extension references an unknown hyperlink".to_string(),
                )
            })?;
            hyperlink.extension = Some(extension);
        }
        Ok(result)
    }

    /// Resolve a hyperlink identifier.
    pub fn get(&self, id: u32) -> Option<&PowerPointHyperlink> {
        self.hyperlinks.iter().find(|hyperlink| hyperlink.id == id)
    }

    fn get_mut(&mut self, id: u32) -> Option<&mut PowerPointHyperlink> {
        self.hyperlinks
            .iter_mut()
            .find(|hyperlink| hyperlink.id == id)
    }

    fn parse_external_object_list(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::ExObjList
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(PptError::Corrupted(
                "ExObjListContainer has an invalid record header".to_string(),
            ));
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "external-object list")?;
        let Some(atom) = children.first() else {
            return Err(PptError::Corrupted(
                "ExObjListContainer is missing ExObjListAtom".to_string(),
            ));
        };
        if atom.record_type != PptRecordType::ExObjListAtom
            || atom.version != 0
            || atom.instance != 0
            || atom.data.len() != 4
        {
            return Err(PptError::Corrupted(
                "ExObjListAtom has an invalid record header or size".to_string(),
            ));
        }
        let id_seed = i32::from_le_bytes([atom.data[0], atom.data[1], atom.data[2], atom.data[3]]);
        if id_seed < 1 {
            return Err(PptError::Corrupted(
                "ExObjListAtom has an invalid identifier seed".to_string(),
            ));
        }

        let mut hyperlinks = Vec::new();
        for child in &children[1..] {
            if child.record_type != PptRecordType::ExternalHyperlink {
                continue;
            }
            let hyperlink = PowerPointHyperlink::parse(child)?;
            if hyperlinks
                .iter()
                .any(|existing: &PowerPointHyperlink| existing.id == hyperlink.id)
            {
                return Err(PptError::Corrupted(
                    "External-object list has duplicate hyperlink IDs".to_string(),
                ));
            }
            hyperlinks.push(hyperlink);
        }
        if hyperlinks
            .iter()
            .any(|hyperlink| hyperlink.id > id_seed as u32)
        {
            return Err(PptError::Corrupted(
                "External-object identifier seed is below a hyperlink ID".to_string(),
            ));
        }
        Ok(Self {
            id_seed: Some(id_seed),
            hyperlinks,
        })
    }
}

fn parse_hyperlink_atom(record: &PptRecord) -> Result<u32> {
    if record.record_type != PptRecordType::ExternalHyperlinkAtom
        || record.version != 0
        || record.instance != 0
        || record.data.len() != 4
    {
        return Err(PptError::Corrupted(
            "ExHyperlinkAtom has an invalid record header or size".to_string(),
        ));
    }
    Ok(u32::from_le_bytes([
        record.data[0],
        record.data[1],
        record.data[2],
        record.data[3],
    ]))
}

fn parse_unicode_string(data: &[u8]) -> Result<String> {
    if data.len() & 1 != 0 {
        return Err(PptError::Corrupted(
            "Hyperlink string has an odd byte length".to_string(),
        ));
    }
    let mut units = Vec::with_capacity(data.len() / 2);
    for bytes in data.chunks_exact(2) {
        let unit = u16::from_le_bytes([bytes[0], bytes[1]]);
        if unit == 0 {
            break;
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_| PptError::Corrupted("Hyperlink string is invalid UTF-16".to_string()))
}

fn collect_records<'a>(
    record: &'a PptRecord,
    record_type: PptRecordType,
    records: &mut Vec<&'a PptRecord>,
) {
    if record.record_type == record_type {
        records.push(record);
    }
    for child in &record.children {
        collect_records(child, record_type, records);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        data.extend_from_slice(&kind.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(payload);
        data
    }

    fn unicode(value: &str) -> Vec<u8> {
        value.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn hyperlink(id: u32) -> Vec<u8> {
        let mut payload = record_bytes(0, 0, 4051, &id.to_le_bytes());
        payload.extend_from_slice(&record_bytes(0, 0, 4026, &unicode("Example")));
        payload.extend_from_slice(&record_bytes(0, 1, 4026, &unicode("https://example.test")));
        payload.extend_from_slice(&record_bytes(0, 3, 4026, &unicode("section")));
        record_bytes(0x0f, 0, 4055, &payload)
    }

    fn external_object_list(seed: i32, hyperlinks: &[Vec<u8>]) -> PptRecord {
        let mut payload = record_bytes(0, 0, 1034, &seed.to_le_bytes());
        for hyperlink in hyperlinks {
            payload.extend_from_slice(hyperlink);
        }
        PptRecord {
            record_type: PptRecordType::ExObjList,
            record_type_raw: 1033,
            version: 0x0f,
            instance: 0,
            data_length: payload.len() as u32,
            data: payload,
            children: Vec::new(),
        }
    }

    fn hyperlink9(id: u32, screen_tip: Option<&str>, flags: u32) -> Vec<u8> {
        let mut payload = record_bytes(0, 0, 4051, &id.to_le_bytes());
        if let Some(screen_tip) = screen_tip {
            payload.extend_from_slice(&record_bytes(0, 0, 4026, &unicode(screen_tip)));
        }
        payload.extend_from_slice(&record_bytes(0, 0, 4120, &flags.to_le_bytes()));
        record_bytes(0x0f, 0, 4068, &payload)
    }

    fn prog_tags_record(blob_payload: &[u8]) -> PptRecord {
        let tag_name: Vec<u8> = "___PPT9"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let name = record_bytes(0, 0, 4026, &tag_name);
        let blob = record_bytes(0, 0, 0x138b, blob_payload);
        let mut tag_payload = name;
        tag_payload.extend_from_slice(&blob);
        let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
        PptRecord {
            record_type: PptRecordType::ProgTags,
            record_type_raw: 0x1388,
            version: 0x0f,
            instance: 0,
            data_length: tag.len() as u32,
            data: tag,
            children: Vec::new(),
        }
    }

    fn interaction_record(trigger: u16, flags: u8, target: u8) -> PptRecord {
        let mut atom = [0u8; 16];
        atom[4..8].copy_from_slice(&3u32.to_le_bytes());
        atom[8] = 4;
        atom[10] = 0;
        atom[11] = flags;
        atom[12] = target;
        let payload = record_bytes(0, 0, 4083, &atom);
        let bytes = record_bytes(0x0f, trigger, 4082, &payload);
        PptRecord::parse(&bytes, 0).unwrap().0
    }

    fn root(list: Option<PptRecord>, extensions: &[Vec<u8>]) -> PptRecord {
        let mut children = Vec::new();
        if let Some(list) = list {
            children.push(list);
        }
        if !extensions.is_empty() {
            let blob: Vec<u8> = extensions.iter().flatten().copied().collect();
            children.push(prog_tags_record(&blob));
        }
        PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    #[test]
    fn parses_and_merges_powerpoint9_hyperlinks() {
        let root = root(
            Some(external_object_list(7, &[hyperlink(3)])),
            &[hyperlink9(3, Some("Open example"), 7)],
        );
        let hyperlinks = PowerPointHyperlinks::parse(&root).unwrap();
        assert_eq!(hyperlinks.id_seed, Some(7));
        let hyperlink = hyperlinks.get(3).unwrap();
        assert_eq!(hyperlink.friendly_name.as_deref(), Some("Example"));
        assert_eq!(hyperlink.target.as_deref(), Some("https://example.test"));
        assert_eq!(hyperlink.location.as_deref(), Some("section"));
        let extension = hyperlink.extension.as_ref().unwrap();
        assert_eq!(extension.screen_tip.as_deref(), Some("Open example"));
        assert!(extension.inserted_with_dialog);
        assert!(extension.location_is_named_show);
        assert!(extension.named_show_returns_to_slide);
    }

    #[test]
    fn parses_and_resolves_interactive_hyperlinks() {
        let interaction = PowerPointInteraction::parse(&interaction_record(0, 0x09, 8)).unwrap();
        assert_eq!(interaction.trigger, InteractionTrigger::Click);
        assert_eq!(interaction.action, InteractionAction::Hyperlink);
        assert_eq!(interaction.link_target, InteractionLinkTarget::Url);
        assert!(interaction.animated);
        assert!(interaction.visited);

        let hyperlinks =
            PowerPointHyperlinks::parse(&root(Some(external_object_list(3, &[hyperlink(3)])), &[]))
                .unwrap();
        assert_eq!(
            interaction
                .hyperlink(&hyperlinks)
                .unwrap()
                .target
                .as_deref(),
            Some("https://example.test")
        );

        assert!(PowerPointInteraction::parse(&interaction_record(2, 0, 8)).is_err());
        assert!(PowerPointInteraction::parse(&interaction_record(0, 0x10, 8)).is_err());
        assert!(PowerPointInteraction::parse(&interaction_record(0, 0, 4)).is_err());
    }

    #[test]
    fn accepts_optional_base_strings_and_absent_extensions() {
        let atom_only = record_bytes(
            0x0f,
            0,
            4055,
            &record_bytes(0, 0, 4051, &1u32.to_le_bytes()),
        );
        let hyperlinks =
            PowerPointHyperlinks::parse(&root(Some(external_object_list(1, &[atom_only])), &[]))
                .unwrap();
        assert_eq!(hyperlinks.get(1).unwrap().target, None);
    }

    #[test]
    fn rejects_invalid_hyperlink_ids_and_extensions() {
        assert!(
            PowerPointHyperlinks::parse(
                &root(Some(external_object_list(2, &[hyperlink(3)])), &[],)
            )
            .is_err()
        );
        assert!(
            PowerPointHyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3), hyperlink(3)])),
                &[],
            ))
            .is_err()
        );
        assert!(
            PowerPointHyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3)])),
                &[hyperlink9(4, None, 0)],
            ))
            .is_err()
        );
        assert!(
            PowerPointHyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3)])),
                &[hyperlink9(3, None, 8)],
            ))
            .is_err()
        );
        assert!(
            PowerPointHyperlinks::parse(&root(
                Some(external_object_list(3, &[hyperlink(3)])),
                &[hyperlink9(3, None, 0), hyperlink9(3, None, 0)],
            ))
            .is_err()
        );
    }

    #[test]
    fn rejects_malformed_hyperlink_strings_and_child_order() {
        let mut invalid_utf16 = hyperlink(1);
        invalid_utf16[28..30].copy_from_slice(&0xd800u16.to_le_bytes());
        assert!(
            PowerPointHyperlinks::parse(&root(
                Some(external_object_list(1, &[invalid_utf16])),
                &[],
            ))
            .is_err()
        );

        let mut payload = record_bytes(0, 0, 4051, &1u32.to_le_bytes());
        payload.extend_from_slice(&record_bytes(0, 3, 4026, &unicode("late")));
        payload.extend_from_slice(&record_bytes(0, 1, 4026, &unicode("early")));
        let out_of_order = record_bytes(0x0f, 0, 4055, &payload);
        assert!(
            PowerPointHyperlinks::parse(
                &root(Some(external_object_list(1, &[out_of_order])), &[],)
            )
            .is_err()
        );
    }
}
