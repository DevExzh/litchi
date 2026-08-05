//! MS-PPT hyperlink and interactive-record codec.

use super::model::*;
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;
use std::borrow::Cow;

impl InteractiveInfoAtom {
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

impl MacroNameAtom {
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

impl Interaction {
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
        let atom = MacroNameAtom::new(value)?;
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
    pub fn parse(record: &Record) -> Result<Self> {
        Self::parse_with_limits(record, InteractionLimits::default())
    }

    /// Parse a container with explicit record and MacroNameAtom bounds.
    pub fn parse_with_limits(record: &Record, limits: InteractionLimits) -> Result<Self> {
        if record.record_type != RecordType::InteractiveInfo
            || record.record_type_raw != RecordType::InteractiveInfo.as_u16()
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
        let children = Record::parse_sequence_strict(&record.data, "interactive information")?;
        if !matches!(children.len(), 1 | 2) {
            return corrupted("InteractiveInfo has an invalid child count");
        }
        let atom = &children[0];
        if atom.record_type != RecordType::InteractiveInfoAtom
            || atom.record_type_raw != RecordType::InteractiveInfoAtom.as_u16()
            || atom.version != 0
            || atom.instance != 0
            || atom.data_length != 16
            || atom.data.len() != 16
        {
            return corrupted("InteractiveInfoAtom has an invalid header or size");
        }
        let parsed_atom = InteractiveInfoAtom::parse_payload(&atom.data)?;
        let macro_atom = if children.len() == 2 {
            let name = &children[1];
            if name.record_type != RecordType::CString
                || name.record_type_raw != RecordType::CString.as_u16()
                || name.version != 0
                || name.instance != 2
                || usize::try_from(name.data_length).ok() != Some(name.data.len())
            {
                return corrupted("MacroNameAtom has an invalid record header or length");
            }
            Some(MacroNameAtom::parse_payload(
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
        Self::parse_bytes_with_limits(bytes, InteractionLimits::default())
    }

    /// Parse one exact complete container from bounded bytes.
    pub fn parse_bytes_with_limits(bytes: &[u8], limits: InteractionLimits) -> Result<Self> {
        if bytes.len() > limits.max_record_bytes {
            return corrupted("InteractiveInfo exceeds the configured record size limit");
        }
        let (record, consumed) = Record::parse_strict(bytes, 0)?;
        if consumed != bytes.len() {
            return corrupted("InteractiveInfo has trailing bytes");
        }
        Self::parse_with_limits(&record, limits)
    }

    /// Return the typed required atom.
    pub fn atom(&self) -> InteractiveInfoAtom {
        InteractiveInfoAtom {
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
    pub fn macro_name_atom(&self) -> Result<Option<MacroNameAtom>> {
        match (&self.macro_name, &self.macro_name_data) {
            (None, None) => Ok(None),
            (Some(text), Some(data)) => {
                let atom = MacroNameAtom::parse_payload(data, usize::MAX)?;
                if atom.text() != text {
                    return corrupted("MacroNameAtom text and exact data disagree");
                }
                Ok(Some(atom))
            },
            (Some(text), None) => Ok(Some(MacroNameAtom::new(text.clone())?)),
            (None, Some(_)) => corrupted("MacroNameAtom exact data is present without text"),
        }
    }

    /// Validate this interaction against explicit resource limits.
    pub fn validate_with_limits(&self, limits: InteractionLimits) -> Result<()> {
        let macro_name = self.validated_macro_name_data(limits)?;
        validate_serialized_interaction_size(
            macro_name.as_ref().map(|data| data.len()),
            limits.max_record_bytes,
        )?;
        Ok(())
    }

    /// Serialize canonical headers and exact atom/name payloads with default limits.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_bytes_with_limits(InteractionLimits::default())
    }

    /// Serialize canonical headers and exact atom/name payloads with explicit limits.
    pub fn to_bytes_with_limits(&self, limits: InteractionLimits) -> Result<Vec<u8>> {
        let macro_name = self.validated_macro_name_data(limits)?;
        validate_serialized_interaction_size(
            macro_name.as_ref().map(|data| data.len()),
            limits.max_record_bytes,
        )?;
        let atom_payload = self.atom().to_payload();
        let mut children = encode_record(
            0,
            0,
            RecordType::InteractiveInfoAtom.as_u16(),
            &atom_payload,
        )?;
        if let Some(name) = macro_name {
            children.extend_from_slice(&encode_record(0, 2, RecordType::CString.as_u16(), &name)?);
        }
        let instance = match self.trigger {
            InteractionTrigger::Click => 0,
            InteractionTrigger::MouseOver => 1,
        };
        encode_record(
            0x0F,
            instance,
            RecordType::InteractiveInfo.as_u16(),
            &children,
        )
    }

    fn validated_macro_name_data(
        &self,
        limits: InteractionLimits,
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
                let atom = MacroNameAtom::new(text.clone())?;
                if atom.raw_utf16().len() > limits.max_macro_name_bytes {
                    return corrupted("MacroNameAtom exceeds the configured size limit");
                }
                Ok(Some(Cow::Owned(atom.raw_utf16)))
            },
            (None, Some(_)) => corrupted("MacroNameAtom exact data is present without text"),
        }
    }

    /// Convert to the generic record model used by existing shape extraction.
    pub fn to_record(&self) -> Result<Record> {
        let bytes = self.to_bytes()?;
        let (record, consumed) = Record::parse_strict(&bytes, 0)?;
        if consumed != bytes.len() {
            return corrupted("serialized InteractiveInfo was only partially parsed");
        }
        Ok(record)
    }

    /// Resolve this action's hyperlink reference.
    pub fn hyperlink<'a>(&self, hyperlinks: &'a Hyperlinks) -> Option<&'a Hyperlink> {
        hyperlinks.get(self.hyperlink_id)
    }

    /// Resolve this action's embedded sound reference.
    pub fn sound<'collection, 'data>(
        &self,
        sounds: &'collection crate::sound_collection::Collection<'data>,
    ) -> Option<&'collection crate::sound_collection::Sound<'data>> {
        sounds.get(self.sound_id)
    }

    /// Validate this action's non-null sound reference.
    pub fn validate_sound_collection(
        &self,
        sounds: &crate::sound_collection::Collection<'_>,
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
        limits: InteractionLimits,
    ) -> Result<Vec<Self>> {
        let mut interactions = Vec::new();
        for record in Record::parse_sequence_strict(data, "shape ClientData")? {
            if record.record_type != RecordType::InteractiveInfo {
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
                .ok_or_else(|| Error::Corrupted("InteractiveInfo length overflows".into()))
        })
        .transpose()?
        .unwrap_or(0);
    let complete_len = 32usize
        .checked_add(macro_record_len)
        .ok_or_else(|| Error::Corrupted("InteractiveInfo length overflows".into()))?;
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
        _ => Err(Error::Corrupted("Invalid interactive action".to_string())),
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
        _ => Err(Error::Corrupted("Invalid interactive jump".to_string())),
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
        _ => Err(Error::Corrupted(
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
        .map_err(|_| Error::Corrupted("MacroNameAtom contains invalid UTF-16".into()))?;
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
        .map_err(|_| Error::Corrupted("PowerPoint record payload exceeds u32".into()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}

impl HyperlinkExtension {
    /// Parse an `ExHyperlink9Container` record and return its referenced ID.
    pub fn parse(record: &Record) -> Result<(u32, Self)> {
        if record.record_type != RecordType::ExternalHyperlink9
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(Error::Corrupted(
                "ExHyperlink9Container has an invalid record header".to_string(),
            ));
        }
        let children = Record::parse_sequence_strict(&record.data, "PowerPoint 9 hyperlink")?;
        if !matches!(children.len(), 2 | 3) {
            return Err(Error::Corrupted(
                "ExHyperlink9Container has an invalid child count".to_string(),
            ));
        }
        let reference = parse_hyperlink_atom(&children[0])?;
        if reference == 0 {
            return Err(Error::Corrupted(
                "ExHyperlink9Container has a null hyperlink reference".to_string(),
            ));
        }
        let (screen_tip, flags_index) = if children.len() == 3 {
            let tip = &children[1];
            if tip.record_type != RecordType::CString || tip.version != 0 || tip.instance != 0 {
                return Err(Error::Corrupted(
                    "ScreenTipAtom has an invalid record header".to_string(),
                ));
            }
            (Some(parse_unicode_string(&tip.data)?), 2)
        } else {
            (None, 1)
        };
        let flags = &children[flags_index];
        if flags.record_type != RecordType::ExternalHyperlinkFlagsAtom
            || flags.version != 0
            || flags.instance != 0
            || flags.data.len() != 4
        {
            return Err(Error::Corrupted(
                "ExHyperlinkFlagsAtom has an invalid record header or size".to_string(),
            ));
        }
        let value =
            u32::from_le_bytes([flags.data[0], flags.data[1], flags.data[2], flags.data[3]]);
        if value & !0x07 != 0 {
            return Err(Error::Corrupted(
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

impl Hyperlink {
    /// Parse an `ExHyperlinkContainer` record.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::ExternalHyperlink
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(Error::Corrupted(
                "ExHyperlinkContainer has an invalid record header".to_string(),
            ));
        }
        let children = Record::parse_sequence_strict(&record.data, "external hyperlink")?;
        let Some(atom) = children.first() else {
            return Err(Error::Corrupted(
                "ExHyperlinkContainer is missing ExHyperlinkAtom".to_string(),
            ));
        };
        let id = parse_hyperlink_atom(atom)?;
        if id == 0 {
            return Err(Error::Corrupted(
                "ExHyperlinkAtom has a zero hyperlink ID".to_string(),
            ));
        }

        let mut friendly_name = None;
        let mut target = None;
        let mut location = None;
        let mut previous_instance = None;
        for child in &children[1..] {
            if child.record_type != RecordType::CString || child.version != 0 {
                return Err(Error::Corrupted(
                    "ExHyperlinkContainer has an unexpected child record".to_string(),
                ));
            }
            if previous_instance.is_some_and(|previous| previous >= child.instance) {
                return Err(Error::Corrupted(
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
                    return Err(Error::Corrupted(
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

impl Hyperlinks {
    /// Discover base hyperlinks and merge all `___PPT9` hyperlink extensions.
    pub fn parse(root: &Record) -> Result<Self> {
        let mut lists = Vec::new();
        collect_records(root, RecordType::ExObjList, &mut lists);
        if lists.len() > 1 {
            return Err(Error::Corrupted(
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
            if record.record_type != RecordType::ExternalHyperlink9 {
                continue;
            }
            let (id, extension) = HyperlinkExtension::parse(&record)?;
            if extension_ids.contains(&id) {
                return Err(Error::Corrupted(
                    "PowerPoint 9 contains duplicate hyperlink extensions".to_string(),
                ));
            }
            extension_ids.push(id);
            let hyperlink = result.get_mut(id).ok_or_else(|| {
                Error::Corrupted(
                    "PowerPoint 9 hyperlink extension references an unknown hyperlink".to_string(),
                )
            })?;
            hyperlink.extension = Some(extension);
        }
        Ok(result)
    }

    /// Resolve a hyperlink identifier.
    pub fn get(&self, id: u32) -> Option<&Hyperlink> {
        self.hyperlinks.iter().find(|hyperlink| hyperlink.id == id)
    }

    fn get_mut(&mut self, id: u32) -> Option<&mut Hyperlink> {
        self.hyperlinks
            .iter_mut()
            .find(|hyperlink| hyperlink.id == id)
    }

    fn parse_external_object_list(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::ExObjList
            || record.version != 0x0f
            || record.instance != 0
        {
            return Err(Error::Corrupted(
                "ExObjListContainer has an invalid record header".to_string(),
            ));
        }
        let children = Record::parse_sequence_strict(&record.data, "external-object list")?;
        let Some(atom) = children.first() else {
            return Err(Error::Corrupted(
                "ExObjListContainer is missing ExObjListAtom".to_string(),
            ));
        };
        if atom.record_type != RecordType::ExObjListAtom
            || atom.version != 0
            || atom.instance != 0
            || atom.data.len() != 4
        {
            return Err(Error::Corrupted(
                "ExObjListAtom has an invalid record header or size".to_string(),
            ));
        }
        let id_seed = i32::from_le_bytes([atom.data[0], atom.data[1], atom.data[2], atom.data[3]]);
        if id_seed < 1 {
            return Err(Error::Corrupted(
                "ExObjListAtom has an invalid identifier seed".to_string(),
            ));
        }

        let mut hyperlinks = Vec::new();
        for child in &children[1..] {
            if child.record_type != RecordType::ExternalHyperlink {
                continue;
            }
            let hyperlink = Hyperlink::parse(child)?;
            if hyperlinks
                .iter()
                .any(|existing: &Hyperlink| existing.id == hyperlink.id)
            {
                return Err(Error::Corrupted(
                    "External-object list has duplicate hyperlink IDs".to_string(),
                ));
            }
            hyperlinks.push(hyperlink);
        }
        if hyperlinks
            .iter()
            .any(|hyperlink| hyperlink.id > id_seed as u32)
        {
            return Err(Error::Corrupted(
                "External-object identifier seed is below a hyperlink ID".to_string(),
            ));
        }
        Ok(Self {
            id_seed: Some(id_seed),
            hyperlinks,
        })
    }
}

fn parse_hyperlink_atom(record: &Record) -> Result<u32> {
    if record.record_type != RecordType::ExternalHyperlinkAtom
        || record.version != 0
        || record.instance != 0
        || record.data.len() != 4
    {
        return Err(Error::Corrupted(
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
        return Err(Error::Corrupted(
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
        .map_err(|_| Error::Corrupted("Hyperlink string is invalid UTF-16".to_string()))
}

fn collect_records<'a>(record: &'a Record, record_type: RecordType, records: &mut Vec<&'a Record>) {
    if record.record_type == record_type {
        records.push(record);
    }
    for child in &record.children {
        collect_records(child, record_type, records);
    }
}
