//! Strict, inert metadata for legacy PowerPoint OLE objects.
//!
//! This module never loads an embedded storage, invokes COM, starts an OLE
//! server, follows a link, or executes object content.

use super::external_media::PowerPointExternalMediaCollection;
use super::hyperlink::PowerPointHyperlinks;
use super::package::{PptError, Result};
use super::persist::PersistMapping;
use super::records::PptRecord;
use crate::consts::PptRecordType;
use std::collections::HashSet;

const MAX_OLE_NAME_UNITS: usize = 32_768;
const MAX_METAFILE_BYTES: usize = 64 * 1_048_576;
const MAX_OLE_OBJECTS: usize = 4_096;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum PowerPointOleDrawAspect {
    Content = 1,
    Thumbnail = 2,
    Icon = 4,
    DocumentPrint = 8,
}

impl PowerPointOleDrawAspect {
    fn parse(value: u32) -> Result<Self> {
        match value {
            1 => Ok(Self::Content),
            2 => Ok(Self::Thumbnail),
            4 => Ok(Self::Icon),
            8 => Ok(Self::DocumentPrint),
            _ => corrupted("ExOleObjAtom contains an invalid draw aspect"),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum PowerPointOleObjectType {
    Embedded = 0,
    Linked = 1,
    ActiveXControl = 2,
}

impl PowerPointOleObjectType {
    fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Embedded),
            1 => Ok(Self::Linked),
            2 => Ok(Self::ActiveXControl),
            _ => corrupted("ExOleObjAtom contains an invalid object type"),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum PowerPointOleObjectSubtype {
    Default = 0,
    ClipArtGallery = 1,
    WordTable = 2,
    Excel = 3,
    Graph = 4,
    OrganizationChart = 5,
    Equation = 6,
    WordArt = 7,
    Sound = 8,
    Image = 9,
    PowerPointPresentation = 10,
    PowerPointSlide = 11,
    Project = 12,
    NoteIt = 13,
    ExcelChart = 14,
    MediaPlayer = 15,
}

impl PowerPointOleObjectSubtype {
    fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Default),
            1 => Ok(Self::ClipArtGallery),
            2 => Ok(Self::WordTable),
            3 => Ok(Self::Excel),
            4 => Ok(Self::Graph),
            5 => Ok(Self::OrganizationChart),
            6 => Ok(Self::Equation),
            7 => Ok(Self::WordArt),
            8 => Ok(Self::Sound),
            9 => Ok(Self::Image),
            10 => Ok(Self::PowerPointPresentation),
            11 => Ok(Self::PowerPointSlide),
            12 => Ok(Self::Project),
            13 => Ok(Self::NoteIt),
            14 => Ok(Self::ExcelChart),
            15 => Ok(Self::MediaPlayer),
            _ => corrupted("ExOleObjAtom contains an invalid subtype"),
        }
    }
}

/// The exact 24-byte payload of an `ExOleObjAtom`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PowerPointOleObjectMetadata {
    pub draw_aspect: PowerPointOleDrawAspect,
    pub object_type: PowerPointOleObjectType,
    pub id: u32,
    pub subtype: PowerPointOleObjectSubtype,
    pub persist_id: u32,
    pub unused: [u8; 4],
}

impl PowerPointOleObjectMetadata {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        require_atom(
            record,
            1,
            0,
            PptRecordType::ExternalOleObjectAtom,
            24,
            "ExOleObjAtom",
        )?;
        let value = Self {
            draw_aspect: PowerPointOleDrawAspect::parse(u32_at(&record.data, 0))?,
            object_type: PowerPointOleObjectType::parse(u32_at(&record.data, 4))?,
            id: u32_at(&record.data, 8),
            subtype: PowerPointOleObjectSubtype::parse(u32_at(&record.data, 12))?,
            persist_id: u32_at(&record.data, 16),
            unused: record.data[20..24].try_into().expect("fixed slice"),
        };
        value.validate()?;
        Ok(value)
    }

    fn validate(&self) -> Result<()> {
        if self.id == 0 {
            return corrupted("ExOleObjAtom object ID must be positive");
        }
        if self.persist_id == 0 {
            return corrupted("ExOleObjAtom persist ID must be positive");
        }
        Ok(())
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.validate()?;
        let mut data = [0; 24];
        data[0..4].copy_from_slice(&(self.draw_aspect as u32).to_le_bytes());
        data[4..8].copy_from_slice(&(self.object_type as u32).to_le_bytes());
        data[8..12].copy_from_slice(&self.id.to_le_bytes());
        data[12..16].copy_from_slice(&(self.subtype as u32).to_le_bytes());
        data[16..20].copy_from_slice(&self.persist_id.to_le_bytes());
        data[20..24].copy_from_slice(&self.unused);
        record_bytes(1, 0, PptRecordType::ExternalOleObjectAtom, &data)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum PowerPointOleColorFollow {
    None = 0,
    EntireScheme = 1,
    TextAndBackground = 2,
}

impl PowerPointOleColorFollow {
    fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::EntireScheme),
            2 => Ok(Self::TextAndBackground),
            _ => corrupted("ExOleEmbedAtom contains an invalid color-follow value"),
        }
    }
}

/// The recommendation-level dimension policy preserves producer-defined bytes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PowerPointOleDimensionPolicy {
    Send,
    Omit,
    ProducerDefined(u8),
}

impl PowerPointOleDimensionPolicy {
    fn parse(value: u8) -> Self {
        match value {
            0 => Self::Send,
            1 => Self::Omit,
            value => Self::ProducerDefined(value),
        }
    }

    fn value(self) -> u8 {
        match self {
            Self::Send => 0,
            Self::Omit => 1,
            Self::ProducerDefined(value) => value,
        }
    }
}

/// The exact eight-byte payload of an `ExOleEmbedAtom`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PowerPointOleEmbedPreferences {
    pub color_follow: PowerPointOleColorFollow,
    pub cannot_lock_server: bool,
    pub dimension_policy: PowerPointOleDimensionPolicy,
    pub is_word_table: bool,
    pub unused: u8,
}

impl PowerPointOleEmbedPreferences {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        require_atom(
            record,
            0,
            0,
            PptRecordType::ExternalOleEmbedAtom,
            8,
            "ExOleEmbedAtom",
        )?;
        Ok(Self {
            color_follow: PowerPointOleColorFollow::parse(u32_at(&record.data, 0))?,
            cannot_lock_server: parse_bool(record.data[4], "fCantLockServer")?,
            dimension_policy: PowerPointOleDimensionPolicy::parse(record.data[5]),
            is_word_table: parse_bool(record.data[6], "fIsTable")?,
            unused: record.data[7],
        })
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let mut data = [0; 8];
        data[0..4].copy_from_slice(&(self.color_follow as u32).to_le_bytes());
        data[4] = self.cannot_lock_server as u8;
        data[5] = self.dimension_policy.value();
        data[6] = self.is_word_table as u8;
        data[7] = self.unused;
        record_bytes(0, 0, PptRecordType::ExternalOleEmbedAtom, &data)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum PowerPointOleUpdateMode {
    Always = 0,
    OnCall = 1,
}

impl PowerPointOleUpdateMode {
    fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Always),
            1 => Ok(Self::OnCall),
            _ => corrupted("ExOleLinkAtom contains an invalid update mode"),
        }
    }
}

/// Inert link metadata. No link is followed by this type.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PowerPointOleLinkInfo {
    pub slide_id: Option<u32>,
    pub update_mode: PowerPointOleUpdateMode,
    pub unused: [u8; 4],
}

impl PowerPointOleLinkInfo {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        require_atom(
            record,
            0,
            0,
            PptRecordType::ExternalOleLinkAtom,
            12,
            "ExOleLinkAtom",
        )?;
        let slide_id = u32_at(&record.data, 0);
        Ok(Self {
            slide_id: (slide_id != 0).then_some(slide_id),
            update_mode: PowerPointOleUpdateMode::parse(u32_at(&record.data, 4))?,
            unused: record.data[8..12].try_into().expect("fixed slice"),
        })
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let mut data = [0; 12];
        data[0..4].copy_from_slice(&self.slide_id.unwrap_or(0).to_le_bytes());
        data[4..8].copy_from_slice(&(self.update_mode as u32).to_le_bytes());
        data[8..12].copy_from_slice(&self.unused);
        record_bytes(0, 0, PptRecordType::ExternalOleLinkAtom, &data)
    }
}

/// Container-specific metadata preceding the shared OLE object atom.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PowerPointOleContainerKind {
    Embedded(PowerPointOleEmbedPreferences),
    Linked(PowerPointOleLinkInfo),
}

/// A strict, inert `ExOleEmbedContainer` or `ExOleLinkContainer`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PowerPointOleObjectDefinition {
    pub kind: PowerPointOleContainerKind,
    pub object: PowerPointOleObjectMetadata,
    pub menu_name: Option<String>,
    pub program_id: Option<String>,
    pub clipboard_name: Option<String>,
    /// Opaque icon bytes. They are retained but never decoded or rendered here.
    pub metafile: Option<Vec<u8>>,
}

/// Inert metadata for an `ExControlContainer` ActiveX definition.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PowerPointOleControl {
    pub slide_id: Option<u32>,
    pub object: PowerPointOleObjectMetadata,
    pub menu_name: Option<String>,
    pub program_id: Option<String>,
    pub clipboard_name: Option<String>,
    /// Opaque icon bytes. Control storage is not loaded or executed.
    pub metafile: Option<Vec<u8>>,
}

impl PowerPointOleControl {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0x0f
            || record.instance != 0
            || record.record_type != PptRecordType::ExternalOleControl
        {
            return corrupted("ExControlContainer has an invalid header");
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "ExControlContainer")?;
        if !(2..=6).contains(&children.len()) {
            return corrupted("ExControlContainer has an invalid child count");
        }
        require_atom(
            &children[0],
            0,
            0,
            PptRecordType::ExternalOleControlAtom,
            4,
            "ExControlAtom",
        )?;
        let slide_id = u32_at(&children[0].data, 0);
        let object = PowerPointOleObjectMetadata::parse(&children[1])?;
        if object.object_type != PowerPointOleObjectType::ActiveXControl {
            return corrupted("ExControlContainer requires an ActiveX ExOleObjAtom");
        }
        let (menu_name, program_id, clipboard_name, metafile) =
            parse_optional_ole_children(&children[2..])?;
        Ok(Self {
            slide_id: (slide_id != 0).then_some(slide_id),
            object,
            menu_name,
            program_id,
            clipboard_name,
            metafile,
        })
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        if self.object.object_type != PowerPointOleObjectType::ActiveXControl {
            return corrupted("ExControlContainer requires an ActiveX ExOleObjAtom");
        }
        let mut children = record_bytes(
            0,
            0,
            PptRecordType::ExternalOleControlAtom,
            &self.slide_id.unwrap_or(0).to_le_bytes(),
        )?;
        children.extend_from_slice(&self.object.to_record_bytes()?);
        append_optional_ole_children(
            &mut children,
            self.menu_name.as_deref(),
            self.program_id.as_deref(),
            self.clipboard_name.as_deref(),
            self.metafile.as_deref(),
        )?;
        record_bytes(0x0f, 0, PptRecordType::ExternalOleControl, &children)
    }
}

impl PowerPointOleObjectDefinition {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0x0f || record.instance != 0 {
            return corrupted("OLE object container has an invalid header");
        }
        let expected_type = match record.record_type {
            PptRecordType::ExternalOleEmbed => PowerPointOleObjectType::Embedded,
            PptRecordType::ExternalOleLink => PowerPointOleObjectType::Linked,
            _ => return corrupted("OLE object container has an invalid record type"),
        };
        let children = PptRecord::parse_sequence_strict(&record.data, "OLE object container")?;
        if !(2..=6).contains(&children.len()) {
            return corrupted("OLE object container has an invalid child count");
        }
        let kind = match expected_type {
            PowerPointOleObjectType::Embedded => PowerPointOleContainerKind::Embedded(
                PowerPointOleEmbedPreferences::parse(&children[0])?,
            ),
            PowerPointOleObjectType::Linked => {
                PowerPointOleContainerKind::Linked(PowerPointOleLinkInfo::parse(&children[0])?)
            },
            PowerPointOleObjectType::ActiveXControl => unreachable!("container type is bounded"),
        };
        let object = PowerPointOleObjectMetadata::parse(&children[1])?;
        if object.object_type != expected_type {
            return corrupted("OLE container type disagrees with ExOleObjAtom");
        }

        let mut menu_name = None;
        let mut program_id = None;
        let mut clipboard_name = None;
        let mut metafile = None;
        let mut last_string_instance = 0u16;
        for child in &children[2..] {
            if child.record_type == PptRecordType::CString {
                if metafile.is_some()
                    || !(1..=3).contains(&child.instance)
                    || child.instance <= last_string_instance
                {
                    return corrupted("OLE object string atoms are duplicated or out of order");
                }
                last_string_instance = child.instance;
                let value = parse_ole_string(child, child.instance != 1)?;
                match child.instance {
                    1 => menu_name = Some(value),
                    2 => program_id = Some(value),
                    3 => clipboard_name = Some(value),
                    _ => unreachable!("instance was bounded"),
                }
            } else if child.record_type == PptRecordType::MetaFile {
                if metafile.is_some()
                    || child.version != 0
                    || child.instance != 0
                    || child.data.len() > MAX_METAFILE_BYTES
                    || usize::try_from(child.data_length).ok() != Some(child.data.len())
                {
                    return corrupted("MetafileBlob has an invalid header, size, or placement");
                }
                metafile = Some(child.data.clone());
            } else {
                return corrupted("OLE object container contains an unexpected child record");
            }
        }
        Ok(Self {
            kind,
            object,
            menu_name,
            program_id,
            clipboard_name,
            metafile,
        })
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let (container_type, expected_type, first) = match self.kind {
            PowerPointOleContainerKind::Embedded(value) => (
                PptRecordType::ExternalOleEmbed,
                PowerPointOleObjectType::Embedded,
                value.to_record_bytes()?,
            ),
            PowerPointOleContainerKind::Linked(value) => (
                PptRecordType::ExternalOleLink,
                PowerPointOleObjectType::Linked,
                value.to_record_bytes()?,
            ),
        };
        if self.object.object_type != expected_type {
            return corrupted("OLE container type disagrees with ExOleObjAtom");
        }
        let mut children = first;
        children.extend_from_slice(&self.object.to_record_bytes()?);
        for (instance, value, printable) in [
            (1, self.menu_name.as_deref(), false),
            (2, self.program_id.as_deref(), true),
            (3, self.clipboard_name.as_deref(), true),
        ] {
            if let Some(value) = value {
                children.extend_from_slice(&record_bytes(
                    0,
                    instance,
                    PptRecordType::CString,
                    &encode_ole_string(value, printable)?,
                )?);
            }
        }
        if let Some(metafile) = &self.metafile {
            if metafile.len() > MAX_METAFILE_BYTES {
                return corrupted("MetafileBlob exceeds 64 MiB");
            }
            children.extend_from_slice(&record_bytes(0, 0, PptRecordType::MetaFile, metafile)?);
        }
        record_bytes(0x0f, 0, container_type, &children)
    }
}

/// Strict embedded and linked OLE definitions in document order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PowerPointOleObjectCollection {
    pub id_seed: u32,
    pub objects: Vec<PowerPointOleExternalObject>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PowerPointOleExternalObject {
    Object(PowerPointOleObjectDefinition),
    ActiveXControl(PowerPointOleControl),
}

impl PowerPointOleExternalObject {
    pub fn id(&self) -> u32 {
        match self {
            Self::Object(value) => value.object.id,
            Self::ActiveXControl(value) => value.object.id,
        }
    }

    pub fn persist_id(&self) -> u32 {
        match self {
            Self::Object(value) => value.object.persist_id,
            Self::ActiveXControl(value) => value.object.persist_id,
        }
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        match self {
            Self::Object(value) => value.to_record_bytes(),
            Self::ActiveXControl(value) => value.to_record_bytes(),
        }
    }
}

impl PowerPointOleObjectCollection {
    pub fn parse(root: &PptRecord) -> Result<Option<Self>> {
        let mut lists = Vec::new();
        collect_external_object_lists(root, &mut lists);
        if lists.len() > 1 {
            return corrupted("record tree contains multiple external-object lists");
        }
        let Some(list) = lists.first() else {
            return Ok(None);
        };
        if list.version != 0x0f
            || list.instance != 0
            || list.record_type != PptRecordType::ExObjList
        {
            return corrupted("ExObjListContainer has an invalid header");
        }
        let children = PptRecord::parse_sequence_strict(&list.data, "ExObjListContainer")?;
        let Some(atom) = children.first() else {
            return corrupted("ExObjListContainer is missing ExObjListAtom");
        };
        require_atom(atom, 0, 0, PptRecordType::ExObjListAtom, 4, "ExObjListAtom")?;
        let signed_seed = i32::from_le_bytes(atom.data[..4].try_into().expect("fixed slice"));
        if signed_seed < 1 {
            return corrupted("ExObjListAtom identifier seed must be positive");
        }
        let id_seed = signed_seed as u32;
        let mut ids = HashSet::new();
        let mut objects = Vec::new();
        for child in &children[1..] {
            if !matches!(
                child.record_type,
                PptRecordType::ExternalOleEmbed
                    | PptRecordType::ExternalOleLink
                    | PptRecordType::ExternalOleControl
            ) {
                continue;
            }
            if objects.len() >= MAX_OLE_OBJECTS {
                return corrupted(format!(
                    "external-object list exceeds {MAX_OLE_OBJECTS} OLE objects"
                ));
            }
            let object = match child.record_type {
                PptRecordType::ExternalOleControl => {
                    PowerPointOleExternalObject::ActiveXControl(PowerPointOleControl::parse(child)?)
                },
                _ => PowerPointOleExternalObject::Object(PowerPointOleObjectDefinition::parse(
                    child,
                )?),
            };
            let id = object.id();
            if id > id_seed {
                return corrupted(format!(
                    "OLE object ID {id} exceeds ExObjList seed {id_seed}"
                ));
            }
            if !ids.insert(id) {
                return corrupted(format!(
                    "external-object list contains duplicate OLE object ID {id}"
                ));
            }
            objects.push(object);
        }

        if let Some(media) = PowerPointExternalMediaCollection::parse(root)? {
            if objects
                .iter()
                .any(|object| media.get(object.id()).is_some())
            {
                return corrupted("external-object list reuses an ID for OLE and media objects");
            }
        }
        let hyperlinks = PowerPointHyperlinks::parse(root)?;
        if objects.iter().any(|object| {
            hyperlinks
                .hyperlinks
                .iter()
                .any(|hyperlink| hyperlink.id == object.id())
        }) {
            return corrupted("external-object list reuses an ID for OLE objects and hyperlinks");
        }
        Ok(Some(Self { id_seed, objects }))
    }

    pub fn get(&self, id: u32) -> Option<&PowerPointOleExternalObject> {
        self.objects.iter().find(|object| object.id() == id)
    }

    pub fn find(&self, id: u32) -> Option<&PowerPointOleExternalObject> {
        self.get(id)
    }

    pub fn add(&mut self, object: PowerPointOleExternalObject) -> Result<()> {
        let mut candidate = self.objects.clone();
        candidate.push(object);
        validate_collection(self.id_seed, &candidate)?;
        self.objects = candidate;
        Ok(())
    }

    pub fn update<F>(&mut self, id: u32, edit: F) -> Result<()>
    where
        F: FnOnce(&mut PowerPointOleExternalObject) -> Result<()>,
    {
        let mut candidate = self.objects.clone();
        let object = candidate
            .iter_mut()
            .find(|object| object.id() == id)
            .ok_or_else(|| PptError::Corrupted(format!("OLE object ID {id} was not found")))?;
        edit(object)?;
        validate_collection(self.id_seed, &candidate)?;
        self.objects = candidate;
        Ok(())
    }

    pub fn replace(
        &mut self,
        id: u32,
        replacement: PowerPointOleExternalObject,
    ) -> Result<PowerPointOleExternalObject> {
        let mut candidate = self.objects.clone();
        let index = candidate
            .iter()
            .position(|object| object.id() == id)
            .ok_or_else(|| PptError::Corrupted(format!("OLE object ID {id} was not found")))?;
        let previous = std::mem::replace(&mut candidate[index], replacement);
        validate_collection(self.id_seed, &candidate)?;
        self.objects = candidate;
        Ok(previous)
    }

    pub fn remove(&mut self, id: u32) -> Result<PowerPointOleExternalObject> {
        let index = self
            .objects
            .iter()
            .position(|object| object.id() == id)
            .ok_or_else(|| PptError::Corrupted(format!("OLE object ID {id} was not found")))?;
        Ok(self.objects.remove(index))
    }

    pub fn reorder(&mut self, ids: &[u32]) -> Result<()> {
        if ids.len() != self.objects.len() {
            return corrupted("OLE reorder must contain every object exactly once");
        }
        let mut remaining = self.objects.clone();
        let mut candidate = Vec::with_capacity(ids.len());
        for id in ids {
            let index = remaining
                .iter()
                .position(|object| object.id() == *id)
                .ok_or_else(|| {
                    PptError::Corrupted(format!("unknown or repeated OLE object ID {id}"))
                })?;
            candidate.push(remaining.remove(index));
        }
        validate_collection(self.id_seed, &candidate)?;
        self.objects = candidate;
        Ok(())
    }

    pub fn validate(&self) -> Result<()> {
        validate_collection(self.id_seed, &self.objects)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.validate()?;
        let seed = i32::try_from(self.id_seed)
            .map_err(|_| PptError::Corrupted("ExObjList identifier seed exceeds i32".into()))?;
        let mut children = record_bytes(
            0,
            0,
            PptRecordType::ExObjListAtom,
            &seed.to_le_bytes(),
        )?;
        for object in &self.objects {
            children.extend_from_slice(&object.to_record_bytes()?);
        }
        record_bytes(0x0f, 0, PptRecordType::ExObjList, &children)
    }

    pub fn validate_persist_mapping(&self, mapping: &PersistMapping) -> Result<()> {
        for object in &self.objects {
            let id = object.persist_id();
            if mapping.get_offset(id).is_none() {
                return corrupted(format!("OLE object references missing persist ID {id}"));
            }
        }
        Ok(())
    }
}

fn validate_collection(id_seed: u32, objects: &[PowerPointOleExternalObject]) -> Result<()> {
    if id_seed == 0 || id_seed > i32::MAX as u32 {
        return corrupted("ExObjList identifier seed must fit a positive signed integer");
    }
    if objects.len() > MAX_OLE_OBJECTS {
        return corrupted(format!(
            "external-object list exceeds {MAX_OLE_OBJECTS} OLE objects"
        ));
    }
    let mut ids = HashSet::new();
    for object in objects {
        let id = object.id();
        if id == 0 || id > id_seed {
            return corrupted(format!(
                "OLE object ID {id} is zero or exceeds ExObjList seed {id_seed}"
            ));
        }
        if object.persist_id() == 0 {
            return corrupted(format!("OLE object ID {id} has zero persist ID"));
        }
        if !ids.insert(id) {
            return corrupted(format!(
                "external-object list contains duplicate OLE object ID {id}"
            ));
        }
    }
    Ok(())
}

fn parse_optional_ole_children(
    children: &[PptRecord],
) -> Result<(
    Option<String>,
    Option<String>,
    Option<String>,
    Option<Vec<u8>>,
)> {
    let mut menu_name = None;
    let mut program_id = None;
    let mut clipboard_name = None;
    let mut metafile = None;
    let mut last_string_instance = 0u16;
    for child in children {
        if child.record_type == PptRecordType::CString {
            if metafile.is_some()
                || !(1..=3).contains(&child.instance)
                || child.instance <= last_string_instance
            {
                return corrupted("OLE object string atoms are duplicated or out of order");
            }
            last_string_instance = child.instance;
            let value = parse_ole_string(child, child.instance != 1)?;
            match child.instance {
                1 => menu_name = Some(value),
                2 => program_id = Some(value),
                3 => clipboard_name = Some(value),
                _ => unreachable!("instance was bounded"),
            }
        } else if child.record_type == PptRecordType::MetaFile {
            if metafile.is_some()
                || child.version != 0
                || child.instance != 0
                || child.data.len() > MAX_METAFILE_BYTES
                || usize::try_from(child.data_length).ok() != Some(child.data.len())
            {
                return corrupted("MetafileBlob has an invalid header, size, or placement");
            }
            metafile = Some(child.data.clone());
        } else {
            return corrupted("OLE object container contains an unexpected child record");
        }
    }
    Ok((menu_name, program_id, clipboard_name, metafile))
}

fn append_optional_ole_children(
    children: &mut Vec<u8>,
    menu_name: Option<&str>,
    program_id: Option<&str>,
    clipboard_name: Option<&str>,
    metafile: Option<&[u8]>,
) -> Result<()> {
    for (instance, value, printable) in [
        (1, menu_name, false),
        (2, program_id, true),
        (3, clipboard_name, true),
    ] {
        if let Some(value) = value {
            children.extend_from_slice(&record_bytes(
                0,
                instance,
                PptRecordType::CString,
                &encode_ole_string(value, printable)?,
            )?);
        }
    }
    if let Some(metafile) = metafile {
        if metafile.len() > MAX_METAFILE_BYTES {
            return corrupted("MetafileBlob exceeds 64 MiB");
        }
        children.extend_from_slice(&record_bytes(0, 0, PptRecordType::MetaFile, metafile)?);
    }
    Ok(())
}

fn collect_external_object_lists<'a>(record: &'a PptRecord, lists: &mut Vec<&'a PptRecord>) {
    if record.record_type == PptRecordType::ExObjList {
        lists.push(record);
    }
    for child in &record.children {
        collect_external_object_lists(child, lists);
    }
}

fn parse_ole_string(record: &PptRecord, printable: bool) -> Result<String> {
    if record.version != 0
        || record.record_type != PptRecordType::CString
        || record.data.len() % 2 != 0
        || record.data.len() / 2 > MAX_OLE_NAME_UNITS
        || usize::try_from(record.data_length).ok() != Some(record.data.len())
    {
        return corrupted("OLE object string atom has an invalid header or size");
    }
    let units = record
        .data
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect::<Vec<_>>();
    if units.contains(&0) {
        return corrupted("OLE object string contains an embedded null");
    }
    let value = String::from_utf16(&units)
        .map_err(|_| PptError::Corrupted("OLE object string contains invalid UTF-16".into()))?;
    if printable && value.chars().any(char::is_control) {
        return corrupted("OLE object printable string contains a control character");
    }
    Ok(value)
}

fn encode_ole_string(value: &str, printable: bool) -> Result<Vec<u8>> {
    if value.contains('\0') {
        return corrupted("OLE object string contains an embedded null");
    }
    if printable && value.chars().any(char::is_control) {
        return corrupted("OLE object printable string contains a control character");
    }
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.len() > MAX_OLE_NAME_UNITS {
        return corrupted(format!(
            "OLE object string exceeds {MAX_OLE_NAME_UNITS} UTF-16 units"
        ));
    }
    Ok(units.into_iter().flat_map(u16::to_le_bytes).collect())
}

fn parse_bool(value: u8, field: &str) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => corrupted(format!("ExOleEmbedAtom {field} is not a bool1")),
    }
}

fn u32_at(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().expect("fixed slice"))
}

fn require_atom(
    record: &PptRecord,
    version: u16,
    instance: u16,
    kind: PptRecordType,
    length: usize,
    context: &str,
) -> Result<()> {
    if record.version != version
        || record.instance != instance
        || record.record_type_raw != kind.as_u16()
        || record.data.len() != length
        || usize::try_from(record.data_length).ok() != Some(length)
    {
        return corrupted(format!("{context} has an invalid header or size"));
    }
    Ok(())
}

fn record_bytes(version: u16, instance: u16, kind: PptRecordType, data: &[u8]) -> Result<Vec<u8>> {
    let length = u32::try_from(data.len())
        .map_err(|_| PptError::Corrupted("PowerPoint record payload exceeds u32".into()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&kind.as_u16().to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn metadata() -> PowerPointOleObjectMetadata {
        PowerPointOleObjectMetadata {
            draw_aspect: PowerPointOleDrawAspect::Icon,
            object_type: PowerPointOleObjectType::Embedded,
            id: 17,
            subtype: PowerPointOleObjectSubtype::ExcelChart,
            persist_id: 9,
            unused: [1, 2, 3, 4],
        }
    }

    #[test]
    fn ole_object_metadata_roundtrips_exactly() {
        let expected = metadata();
        let parsed = PowerPointOleObjectMetadata::parse(&expected.to_record().unwrap()).unwrap();
        assert_eq!(parsed, expected);
        assert_eq!(
            parsed.to_record_bytes().unwrap(),
            expected.to_record_bytes().unwrap()
        );
    }

    #[test]
    fn ole_object_metadata_rejects_invalid_domains_and_ids() {
        let mut bytes = metadata().to_record_bytes().unwrap();
        bytes[8..12].copy_from_slice(&99u32.to_le_bytes());
        assert!(
            PowerPointOleObjectMetadata::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err()
        );
        let mut value = metadata();
        value.persist_id = 0;
        assert!(value.to_record_bytes().is_err());
    }

    #[test]
    fn embed_preferences_preserve_recommendation_and_unused_bytes() {
        let expected = PowerPointOleEmbedPreferences {
            color_follow: PowerPointOleColorFollow::TextAndBackground,
            cannot_lock_server: true,
            dimension_policy: PowerPointOleDimensionPolicy::ProducerDefined(7),
            is_word_table: false,
            unused: 0xa5,
        };
        let parsed = PowerPointOleEmbedPreferences::parse(&expected.to_record().unwrap()).unwrap();
        assert_eq!(parsed, expected);
    }

    #[test]
    fn link_info_roundtrips_nullable_slide_and_rejects_update_domain() {
        let expected = PowerPointOleLinkInfo {
            slide_id: None,
            update_mode: PowerPointOleUpdateMode::OnCall,
            unused: [0xde, 0xad, 0xbe, 0xef],
        };
        assert_eq!(
            PowerPointOleLinkInfo::parse(&expected.to_record().unwrap()).unwrap(),
            expected
        );
        let mut bytes = expected.to_record_bytes().unwrap();
        bytes[12..16].copy_from_slice(&2u32.to_le_bytes());
        assert!(PowerPointOleLinkInfo::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());
    }

    fn definition(kind: PowerPointOleContainerKind) -> PowerPointOleObjectDefinition {
        let object_type = match kind {
            PowerPointOleContainerKind::Embedded(_) => PowerPointOleObjectType::Embedded,
            PowerPointOleContainerKind::Linked(_) => PowerPointOleObjectType::Linked,
        };
        PowerPointOleObjectDefinition {
            kind,
            object: PowerPointOleObjectMetadata {
                object_type,
                ..metadata()
            },
            menu_name: Some("Worksheet".into()),
            program_id: Some("Excel.Sheet.12".into()),
            clipboard_name: Some("Microsoft Excel Worksheet".into()),
            metafile: Some(vec![0xd7, 0xcd, 0xc6, 0x9a]),
        }
    }

    #[test]
    fn embedded_and_linked_containers_roundtrip_canonically() {
        let embedded = definition(PowerPointOleContainerKind::Embedded(
            PowerPointOleEmbedPreferences {
                color_follow: PowerPointOleColorFollow::EntireScheme,
                cannot_lock_server: true,
                dimension_policy: PowerPointOleDimensionPolicy::Omit,
                is_word_table: false,
                unused: 7,
            },
        ));
        let linked = definition(PowerPointOleContainerKind::Linked(PowerPointOleLinkInfo {
            slide_id: Some(256),
            update_mode: PowerPointOleUpdateMode::OnCall,
            unused: [1, 2, 3, 4],
        }));
        for expected in [embedded, linked] {
            let parsed =
                PowerPointOleObjectDefinition::parse(&expected.to_record().unwrap()).unwrap();
            assert_eq!(parsed, expected);
        }
    }

    #[test]
    fn containers_reject_type_mismatch_and_hostile_strings() {
        let mut value = definition(PowerPointOleContainerKind::Embedded(
            PowerPointOleEmbedPreferences {
                color_follow: PowerPointOleColorFollow::None,
                cannot_lock_server: false,
                dimension_policy: PowerPointOleDimensionPolicy::Send,
                is_word_table: false,
                unused: 0,
            },
        ));
        value.object.object_type = PowerPointOleObjectType::Linked;
        assert!(value.to_record_bytes().is_err());
        value.object.object_type = PowerPointOleObjectType::Embedded;
        value.program_id = Some("bad\nprogram".into());
        assert!(value.to_record_bytes().is_err());
        value.program_id = Some("x".repeat(MAX_OLE_NAME_UNITS + 1));
        assert!(value.to_record_bytes().is_err());
    }

    #[test]
    fn containers_reject_duplicate_or_out_of_order_optional_atoms() {
        let value = definition(PowerPointOleContainerKind::Embedded(
            PowerPointOleEmbedPreferences {
                color_follow: PowerPointOleColorFollow::None,
                cannot_lock_server: false,
                dimension_policy: PowerPointOleDimensionPolicy::Send,
                is_word_table: false,
                unused: 0,
            },
        ));
        let mut children = value.kind_embedded_bytes_for_test();
        children.extend_from_slice(&value.object.to_record_bytes().unwrap());
        children.extend_from_slice(
            &record_bytes(
                0,
                2,
                PptRecordType::CString,
                &encode_ole_string("Prog", true).unwrap(),
            )
            .unwrap(),
        );
        children.extend_from_slice(
            &record_bytes(
                0,
                1,
                PptRecordType::CString,
                &encode_ole_string("Menu", false).unwrap(),
            )
            .unwrap(),
        );
        let bytes = record_bytes(0x0f, 0, PptRecordType::ExternalOleEmbed, &children).unwrap();
        assert!(
            PowerPointOleObjectDefinition::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err()
        );
    }

    impl PowerPointOleObjectDefinition {
        fn kind_embedded_bytes_for_test(&self) -> Vec<u8> {
            match self.kind {
                PowerPointOleContainerKind::Embedded(value) => value.to_record_bytes().unwrap(),
                PowerPointOleContainerKind::Linked(_) => unreachable!(),
            }
        }
    }

    fn external_object_list(seed: i32, objects: &[Vec<u8>]) -> PptRecord {
        let mut children =
            record_bytes(0, 0, PptRecordType::ExObjListAtom, &seed.to_le_bytes()).unwrap();
        for object in objects {
            children.extend_from_slice(object);
        }
        let bytes = record_bytes(0x0f, 0, PptRecordType::ExObjList, &children).unwrap();
        PptRecord::parse(&bytes, 0).unwrap().0
    }

    #[test]
    fn ole_collection_discovers_objects_and_enforces_seed() {
        let mut first = definition(PowerPointOleContainerKind::Embedded(
            PowerPointOleEmbedPreferences {
                color_follow: PowerPointOleColorFollow::None,
                cannot_lock_server: false,
                dimension_policy: PowerPointOleDimensionPolicy::Send,
                is_word_table: false,
                unused: 0,
            },
        ));
        first.object.id = 21;
        let root = external_object_list(21, &[first.to_record_bytes().unwrap()]);
        let parsed = PowerPointOleObjectCollection::parse(&root)
            .unwrap()
            .unwrap();
        assert_eq!(parsed.id_seed, 21);
        assert!(parsed.get(21).is_some());
        let root = external_object_list(20, &[first.to_record_bytes().unwrap()]);
        assert!(PowerPointOleObjectCollection::parse(&root).is_err());
    }

    #[test]
    fn ole_collection_rejects_duplicate_ids() {
        let first = definition(PowerPointOleContainerKind::Embedded(
            PowerPointOleEmbedPreferences {
                color_follow: PowerPointOleColorFollow::None,
                cannot_lock_server: false,
                dimension_policy: PowerPointOleDimensionPolicy::Send,
                is_word_table: false,
                unused: 0,
            },
        ));
        let mut second = first.clone();
        second.object.persist_id += 1;
        let root = external_object_list(
            first.object.id as i32,
            &[
                first.to_record_bytes().unwrap(),
                second.to_record_bytes().unwrap(),
            ],
        );
        assert!(PowerPointOleObjectCollection::parse(&root).is_err());
    }

    #[test]
    fn activex_control_roundtrips_as_inert_metadata() {
        let expected = PowerPointOleControl {
            slide_id: Some(512),
            object: PowerPointOleObjectMetadata {
                object_type: PowerPointOleObjectType::ActiveXControl,
                ..metadata()
            },
            menu_name: Some("Calendar".into()),
            program_id: Some("MSCAL.Calendar.7".into()),
            clipboard_name: None,
            metafile: Some(vec![1, 2, 3]),
        };
        let parsed = PowerPointOleControl::parse(&expected.to_record().unwrap()).unwrap();
        assert_eq!(parsed, expected);
    }
}
