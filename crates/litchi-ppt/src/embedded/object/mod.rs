//! Strict, inert metadata for legacy PowerPoint OLE objects.
//!
//! This module never loads an embedded storage, invokes COM, starts an OLE
//! server, follows a link, or executes object content.

use crate::consts::PptRecordType;
use crate::external_media::Collection as MediaCollection;
use crate::hyperlink::Hyperlinks;
use crate::package::{PptError, Result};
use crate::persist::PersistMapping;
use crate::records::PptRecord;
use std::collections::HashSet;

pub mod editor;

pub use editor::Editor;

const MAX_OLE_NAME_UNITS: usize = 32_768;
const MAX_METAFILE_BYTES: usize = 64 * 1_048_576;
const MAX_OLE_OBJECTS: usize = 4_096;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum DrawAspect {
    Content = 1,
    Thumbnail = 2,
    Icon = 4,
    DocumentPrint = 8,
}

impl DrawAspect {
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
pub enum ObjectType {
    Embedded = 0,
    Linked = 1,
    ActiveXControl = 2,
}

impl ObjectType {
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
pub enum ObjectSubtype {
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

impl ObjectSubtype {
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
pub struct Metadata {
    pub draw_aspect: DrawAspect,
    pub object_type: ObjectType,
    pub id: u32,
    pub subtype: ObjectSubtype,
    pub persist_id: u32,
    pub unused: [u8; 4],
}

impl Metadata {
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
            draw_aspect: DrawAspect::parse(u32_at(&record.data, 0))?,
            object_type: ObjectType::parse(u32_at(&record.data, 4))?,
            id: u32_at(&record.data, 8),
            subtype: ObjectSubtype::parse(u32_at(&record.data, 12))?,
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
pub enum ColorFollow {
    None = 0,
    EntireScheme = 1,
    TextAndBackground = 2,
}

impl ColorFollow {
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
pub enum DimensionPolicy {
    Send,
    Omit,
    ProducerDefined(u8),
}

impl DimensionPolicy {
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
pub struct EmbedPreferences {
    pub color_follow: ColorFollow,
    pub cannot_lock_server: bool,
    pub dimension_policy: DimensionPolicy,
    pub is_word_table: bool,
    pub unused: u8,
}

impl EmbedPreferences {
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
            color_follow: ColorFollow::parse(u32_at(&record.data, 0))?,
            cannot_lock_server: parse_bool(record.data[4], "fCantLockServer")?,
            dimension_policy: DimensionPolicy::parse(record.data[5]),
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
pub enum UpdateMode {
    Always = 0,
    OnCall = 1,
}

impl UpdateMode {
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
pub struct LinkInfo {
    pub slide_id: Option<u32>,
    pub update_mode: UpdateMode,
    pub unused: [u8; 4],
}

impl LinkInfo {
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
            update_mode: UpdateMode::parse(u32_at(&record.data, 4))?,
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
pub enum ContainerKind {
    Embedded(EmbedPreferences),
    Linked(LinkInfo),
}

/// A strict, inert `ExOleEmbedContainer` or `ExOleLinkContainer`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Definition {
    pub kind: ContainerKind,
    pub object: Metadata,
    pub menu_name: Option<String>,
    pub program_id: Option<String>,
    pub clipboard_name: Option<String>,
    /// Opaque icon bytes. They are retained but never decoded or rendered here.
    pub metafile: Option<Vec<u8>>,
}

/// Inert metadata for an `ExControlContainer` ActiveX definition.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Control {
    pub slide_id: Option<u32>,
    pub object: Metadata,
    pub menu_name: Option<String>,
    pub program_id: Option<String>,
    pub clipboard_name: Option<String>,
    /// Opaque icon bytes. Control storage is not loaded or executed.
    pub metafile: Option<Vec<u8>>,
}

impl Control {
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
        let object = Metadata::parse(&children[1])?;
        if object.object_type != ObjectType::ActiveXControl {
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
        if self.object.object_type != ObjectType::ActiveXControl {
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

impl Definition {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0x0f || record.instance != 0 {
            return corrupted("OLE object container has an invalid header");
        }
        let expected_type = match record.record_type {
            PptRecordType::ExternalOleEmbed => ObjectType::Embedded,
            PptRecordType::ExternalOleLink => ObjectType::Linked,
            _ => return corrupted("OLE object container has an invalid record type"),
        };
        let children = PptRecord::parse_sequence_strict(&record.data, "OLE object container")?;
        if !(2..=6).contains(&children.len()) {
            return corrupted("OLE object container has an invalid child count");
        }
        let kind = match expected_type {
            ObjectType::Embedded => ContainerKind::Embedded(EmbedPreferences::parse(&children[0])?),
            ObjectType::Linked => ContainerKind::Linked(LinkInfo::parse(&children[0])?),
            ObjectType::ActiveXControl => unreachable!("container type is bounded"),
        };
        let object = Metadata::parse(&children[1])?;
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
            ContainerKind::Embedded(value) => (
                PptRecordType::ExternalOleEmbed,
                ObjectType::Embedded,
                value.to_record_bytes()?,
            ),
            ContainerKind::Linked(value) => (
                PptRecordType::ExternalOleLink,
                ObjectType::Linked,
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
pub struct Collection {
    pub id_seed: u32,
    pub objects: Vec<ExternalObject>,
    unknown_records: Vec<UnknownRecord>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ExternalObject {
    Object(Definition),
    ActiveXControl(Control),
}

impl ExternalObject {
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

/// A bounded, lossless child of `ExObjList` that this crate does not model.
///
/// The record header and payload are retained so a typed OLE edit does not
/// discard unrelated media, hyperlink, or future-version records. The record
/// is exposed through borrowed accessors; callers never need to clone its
/// payload merely to inspect it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct UnknownRecord {
    record: PptRecord,
    object_index: usize,
}

impl UnknownRecord {
    /// The original raw record type value.
    pub fn record_type(&self) -> u16 {
        self.record.record_type_raw
    }

    /// The original record version.
    pub fn version(&self) -> u16 {
        self.record.version
    }

    /// The original record instance.
    pub fn instance(&self) -> u16 {
        self.record.instance
    }

    /// The original record payload, borrowed from this collection snapshot.
    pub fn data(&self) -> &[u8] {
        &self.record.data
    }

    /// Reconstructs the exact header and payload retained by this record.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        record_bytes_raw(
            self.record.version,
            self.record.instance,
            self.record.record_type_raw,
            &self.record.data,
        )
    }

    fn from_record(record: &PptRecord, object_index: usize) -> Self {
        Self {
            record: record.clone(),
            object_index,
        }
    }

    fn validate_for(&self, object_count: usize) -> Result<()> {
        if self.object_index > object_count {
            return corrupted("unknown ExObjList record has an invalid source slot");
        }
        let expected_length = usize::try_from(self.record.data_length)
            .map_err(|_| PptError::Corrupted("unknown ExObjList record size overflows".into()))?;
        if expected_length != self.record.data.len() {
            return corrupted("unknown ExObjList record has inconsistent payload length");
        }
        self.to_record_bytes().map(|_| ())
    }
}

impl Collection {
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
        let mut unknown_records = Vec::new();
        for child in &children[1..] {
            if !matches!(
                child.record_type,
                PptRecordType::ExternalOleEmbed
                    | PptRecordType::ExternalOleLink
                    | PptRecordType::ExternalOleControl
            ) {
                unknown_records.push(UnknownRecord::from_record(child, objects.len()));
                continue;
            }
            if objects.len() >= MAX_OLE_OBJECTS {
                return corrupted(format!(
                    "external-object list exceeds {MAX_OLE_OBJECTS} OLE objects"
                ));
            }
            let object = match child.record_type {
                PptRecordType::ExternalOleControl => {
                    ExternalObject::ActiveXControl(Control::parse(child)?)
                },
                _ => ExternalObject::Object(Definition::parse(child)?),
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

        if let Some(media) = MediaCollection::parse(root)?
            && objects
                .iter()
                .any(|object| media.get(object.id()).is_some())
        {
            return corrupted("external-object list reuses an ID for OLE and media objects");
        }
        let hyperlinks = Hyperlinks::parse(root)?;
        if objects.iter().any(|object| {
            hyperlinks
                .hyperlinks
                .iter()
                .any(|hyperlink| hyperlink.id == object.id())
        }) {
            return corrupted("external-object list reuses an ID for OLE objects and hyperlinks");
        }
        Ok(Some(Self {
            id_seed,
            objects,
            unknown_records,
        }))
    }

    pub fn get(&self, id: u32) -> Option<&ExternalObject> {
        self.objects.iter().find(|object| object.id() == id)
    }

    pub fn find(&self, id: u32) -> Option<&ExternalObject> {
        self.get(id)
    }

    /// Unmodeled `ExObjList` children retained in source order.
    pub fn unknown_records(&self) -> &[UnknownRecord] {
        &self.unknown_records
    }

    pub fn add(&mut self, object: ExternalObject) -> Result<()> {
        let mut candidate = self.clone();
        let insertion_index = candidate.objects.len();
        candidate.objects.push(object);
        for record in &mut candidate.unknown_records {
            if record.object_index >= insertion_index {
                record.object_index += 1;
            }
        }
        validate_collection(candidate.id_seed, &candidate.objects)?;
        *self = candidate;
        Ok(())
    }

    pub fn update<F>(&mut self, id: u32, edit: F) -> Result<()>
    where
        F: FnOnce(&mut ExternalObject) -> Result<()>,
    {
        let mut candidate = self.clone();
        let object = candidate
            .objects
            .iter_mut()
            .find(|object| object.id() == id)
            .ok_or_else(|| PptError::Corrupted(format!("OLE object ID {id} was not found")))?;
        edit(object)?;
        validate_collection(candidate.id_seed, &candidate.objects)?;
        *self = candidate;
        Ok(())
    }

    pub fn replace(&mut self, id: u32, replacement: ExternalObject) -> Result<ExternalObject> {
        let mut candidate = self.clone();
        let index = candidate
            .objects
            .iter()
            .position(|object| object.id() == id)
            .ok_or_else(|| PptError::Corrupted(format!("OLE object ID {id} was not found")))?;
        let previous = std::mem::replace(&mut candidate.objects[index], replacement);
        validate_collection(candidate.id_seed, &candidate.objects)?;
        *self = candidate;
        Ok(previous)
    }

    pub fn remove(&mut self, id: u32) -> Result<ExternalObject> {
        let mut candidate = self.clone();
        let index = candidate
            .objects
            .iter()
            .position(|object| object.id() == id)
            .ok_or_else(|| PptError::Corrupted(format!("OLE object ID {id} was not found")))?;
        let removed = candidate.objects.remove(index);
        for record in &mut candidate.unknown_records {
            if record.object_index > index {
                record.object_index -= 1;
            }
        }
        *self = candidate;
        Ok(removed)
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
        validate_collection(self.id_seed, &self.objects)?;
        self.validate_unknown_records()
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.validate()?;
        let seed = i32::try_from(self.id_seed)
            .map_err(|_| PptError::Corrupted("ExObjList identifier seed exceeds i32".into()))?;
        let mut children = record_bytes(0, 0, PptRecordType::ExObjListAtom, &seed.to_le_bytes())?;
        for object_index in 0..=self.objects.len() {
            for record in self
                .unknown_records
                .iter()
                .filter(|record| record.object_index == object_index)
            {
                children.extend_from_slice(&record.to_record_bytes()?);
            }
            if let Some(object) = self.objects.get(object_index) {
                children.extend_from_slice(&object.to_record_bytes()?);
            }
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

    fn validate_unknown_records(&self) -> Result<()> {
        for record in &self.unknown_records {
            record.validate_for(self.objects.len())?;
        }
        Ok(())
    }
}

fn validate_collection(id_seed: u32, objects: &[ExternalObject]) -> Result<()> {
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

#[allow(clippy::type_complexity)]
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
        || !record.data.len().is_multiple_of(2)
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
    record_bytes_raw(version, instance, kind.as_u16(), data)
}

fn record_bytes_raw(version: u16, instance: u16, kind: u16, data: &[u8]) -> Result<Vec<u8>> {
    if version > 0x000f || instance > 0x0fff {
        return corrupted("PowerPoint record header exceeds its encoded domain");
    }
    let length = u32::try_from(data.len())
        .map_err(|_| PptError::Corrupted("PowerPoint record payload exceeds u32".into()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
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

    fn metadata() -> Metadata {
        Metadata {
            draw_aspect: DrawAspect::Icon,
            object_type: ObjectType::Embedded,
            id: 17,
            subtype: ObjectSubtype::ExcelChart,
            persist_id: 9,
            unused: [1, 2, 3, 4],
        }
    }

    #[test]
    fn ole_object_metadata_roundtrips_exactly() {
        let expected = metadata();
        let parsed = Metadata::parse(&expected.to_record().unwrap()).unwrap();
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
        assert!(Metadata::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());
        let mut value = metadata();
        value.persist_id = 0;
        assert!(value.to_record_bytes().is_err());
    }

    #[test]
    fn embed_preferences_preserve_recommendation_and_unused_bytes() {
        let expected = EmbedPreferences {
            color_follow: ColorFollow::TextAndBackground,
            cannot_lock_server: true,
            dimension_policy: DimensionPolicy::ProducerDefined(7),
            is_word_table: false,
            unused: 0xa5,
        };
        let parsed = EmbedPreferences::parse(&expected.to_record().unwrap()).unwrap();
        assert_eq!(parsed, expected);
    }

    #[test]
    fn link_info_roundtrips_nullable_slide_and_rejects_update_domain() {
        let expected = LinkInfo {
            slide_id: None,
            update_mode: UpdateMode::OnCall,
            unused: [0xde, 0xad, 0xbe, 0xef],
        };
        assert_eq!(
            LinkInfo::parse(&expected.to_record().unwrap()).unwrap(),
            expected
        );
        let mut bytes = expected.to_record_bytes().unwrap();
        bytes[12..16].copy_from_slice(&2u32.to_le_bytes());
        assert!(LinkInfo::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());
    }

    fn definition(kind: ContainerKind) -> Definition {
        let object_type = match kind {
            ContainerKind::Embedded(_) => ObjectType::Embedded,
            ContainerKind::Linked(_) => ObjectType::Linked,
        };
        Definition {
            kind,
            object: Metadata {
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
        let embedded = definition(ContainerKind::Embedded(EmbedPreferences {
            color_follow: ColorFollow::EntireScheme,
            cannot_lock_server: true,
            dimension_policy: DimensionPolicy::Omit,
            is_word_table: false,
            unused: 7,
        }));
        let linked = definition(ContainerKind::Linked(LinkInfo {
            slide_id: Some(256),
            update_mode: UpdateMode::OnCall,
            unused: [1, 2, 3, 4],
        }));
        for expected in [embedded, linked] {
            let parsed = Definition::parse(&expected.to_record().unwrap()).unwrap();
            assert_eq!(parsed, expected);
        }
    }

    #[test]
    fn containers_reject_type_mismatch_and_hostile_strings() {
        let mut value = definition(ContainerKind::Embedded(EmbedPreferences {
            color_follow: ColorFollow::None,
            cannot_lock_server: false,
            dimension_policy: DimensionPolicy::Send,
            is_word_table: false,
            unused: 0,
        }));
        value.object.object_type = ObjectType::Linked;
        assert!(value.to_record_bytes().is_err());
        value.object.object_type = ObjectType::Embedded;
        value.program_id = Some("bad\nprogram".into());
        assert!(value.to_record_bytes().is_err());
        value.program_id = Some("x".repeat(MAX_OLE_NAME_UNITS + 1));
        assert!(value.to_record_bytes().is_err());
    }

    #[test]
    fn containers_reject_duplicate_or_out_of_order_optional_atoms() {
        let value = definition(ContainerKind::Embedded(EmbedPreferences {
            color_follow: ColorFollow::None,
            cannot_lock_server: false,
            dimension_policy: DimensionPolicy::Send,
            is_word_table: false,
            unused: 0,
        }));
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
        assert!(Definition::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());
    }

    impl Definition {
        fn kind_embedded_bytes_for_test(&self) -> Vec<u8> {
            match self.kind {
                ContainerKind::Embedded(value) => value.to_record_bytes().unwrap(),
                ContainerKind::Linked(_) => unreachable!(),
            }
        }
    }

    fn external_object_list(seed: i32, objects: &[Vec<u8>]) -> PptRecord {
        external_object_list_with_children(seed, objects).0
    }

    fn external_object_list_with_children(seed: i32, children: &[Vec<u8>]) -> (PptRecord, Vec<u8>) {
        let mut child_bytes =
            record_bytes(0, 0, PptRecordType::ExObjListAtom, &seed.to_le_bytes()).unwrap();
        for child in children {
            child_bytes.extend_from_slice(child);
        }
        let bytes = record_bytes(0x0f, 0, PptRecordType::ExObjList, &child_bytes).unwrap();
        (PptRecord::parse(&bytes, 0).unwrap().0, bytes)
    }

    #[test]
    fn ole_collection_preserves_unknown_children_and_source_slots() {
        let value = definition(ContainerKind::Embedded(EmbedPreferences {
            color_follow: ColorFollow::None,
            cannot_lock_server: false,
            dimension_policy: DimensionPolicy::Send,
            is_word_table: false,
            unused: 0,
        }));
        let first_unknown = record_bytes_raw(0, 7, 0x7777, b"before").unwrap();
        let second_unknown = record_bytes_raw(0, 9, 0x8888, b"after").unwrap();
        let (root, original) = external_object_list_with_children(
            value.object.id as i32,
            &[
                first_unknown.clone(),
                value.to_record_bytes().unwrap(),
                second_unknown.clone(),
            ],
        );

        let collection = Collection::parse(&root).unwrap().unwrap();
        assert_eq!(collection.unknown_records().len(), 2);
        assert_eq!(collection.unknown_records()[0].record_type(), 0x7777);
        assert_eq!(collection.unknown_records()[0].data(), b"before");
        assert_eq!(collection.unknown_records()[1].record_type(), 0x8888);
        assert_eq!(collection.unknown_records()[1].data(), b"after");
        assert_eq!(collection.to_record_bytes().unwrap(), original);
        assert_eq!(
            collection.unknown_records()[0].to_record_bytes().unwrap(),
            first_unknown
        );
        assert_eq!(
            collection.unknown_records()[1].to_record_bytes().unwrap(),
            second_unknown
        );
    }

    #[test]
    fn ole_collection_reorders_typed_objects_without_losing_unknown_slots() {
        let kind = ContainerKind::Embedded(EmbedPreferences {
            color_follow: ColorFollow::None,
            cannot_lock_server: false,
            dimension_policy: DimensionPolicy::Send,
            is_word_table: false,
            unused: 0,
        });
        let mut first = definition(kind);
        first.object.id = 1;
        let mut second = first.clone();
        second.object.id = 2;
        second.object.persist_id = 10;
        let first_unknown = record_bytes_raw(0, 1, 0x7777, b"slot-0").unwrap();
        let middle_unknown = record_bytes_raw(0, 2, 0x8888, b"slot-1").unwrap();
        let last_unknown = record_bytes_raw(0, 3, 0x9999, b"slot-2").unwrap();
        let (root, _) = external_object_list_with_children(
            2,
            &[
                first_unknown.clone(),
                first.to_record_bytes().unwrap(),
                middle_unknown.clone(),
                second.to_record_bytes().unwrap(),
                last_unknown.clone(),
            ],
        );

        let mut collection = Collection::parse(&root).unwrap().unwrap();
        collection.reorder(&[2, 1]).unwrap();
        let (_, reordered) = external_object_list_with_children(
            2,
            &[
                first_unknown,
                second.to_record_bytes().unwrap(),
                middle_unknown,
                first.to_record_bytes().unwrap(),
                last_unknown,
            ],
        );
        assert_eq!(collection.to_record_bytes().unwrap(), reordered);
    }

    #[test]
    fn ole_collection_discovers_objects_and_enforces_seed() {
        let mut first = definition(ContainerKind::Embedded(EmbedPreferences {
            color_follow: ColorFollow::None,
            cannot_lock_server: false,
            dimension_policy: DimensionPolicy::Send,
            is_word_table: false,
            unused: 0,
        }));
        first.object.id = 21;
        let root = external_object_list(21, &[first.to_record_bytes().unwrap()]);
        let parsed = Collection::parse(&root).unwrap().unwrap();
        assert_eq!(parsed.id_seed, 21);
        assert!(parsed.get(21).is_some());
        let root = external_object_list(20, &[first.to_record_bytes().unwrap()]);
        assert!(Collection::parse(&root).is_err());
    }

    #[test]
    fn ole_collection_rejects_duplicate_ids() {
        let first = definition(ContainerKind::Embedded(EmbedPreferences {
            color_follow: ColorFollow::None,
            cannot_lock_server: false,
            dimension_policy: DimensionPolicy::Send,
            is_word_table: false,
            unused: 0,
        }));
        let mut second = first.clone();
        second.object.persist_id += 1;
        let root = external_object_list(
            first.object.id as i32,
            &[
                first.to_record_bytes().unwrap(),
                second.to_record_bytes().unwrap(),
            ],
        );
        assert!(Collection::parse(&root).is_err());
    }

    #[test]
    fn activex_control_roundtrips_as_inert_metadata() {
        let expected = Control {
            slide_id: Some(512),
            object: Metadata {
                object_type: ObjectType::ActiveXControl,
                ..metadata()
            },
            menu_name: Some("Calendar".into()),
            program_id: Some("MSCAL.Calendar.7".into()),
            clipboard_name: None,
            metafile: Some(vec![1, 2, 3]),
        };
        let parsed = Control::parse(&expected.to_record().unwrap()).unwrap();
        assert_eq!(parsed, expected);
    }
}
