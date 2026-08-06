//! Binary record codecs for the typed OLE object models.

use super::model::*;
use crate::consts::RecordType;
use crate::external_media::Collection as MediaCollection;
use crate::hyperlink::Hyperlinks;
use crate::package::{Error, Result};
use crate::records::Record;
use std::collections::HashSet;

impl DrawAspect {
    pub(crate) fn parse(value: u32) -> Result<Self> {
        match value {
            1 => Ok(Self::Content),
            2 => Ok(Self::Thumbnail),
            4 => Ok(Self::Icon),
            8 => Ok(Self::DocumentPrint),
            _ => corrupted("ExOleObjAtom contains an invalid draw aspect"),
        }
    }
}

impl ObjectType {
    pub(crate) fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Embedded),
            1 => Ok(Self::Linked),
            2 => Ok(Self::ActiveXControl),
            _ => corrupted("ExOleObjAtom contains an invalid object type"),
        }
    }
}

impl ObjectSubtype {
    pub(crate) fn parse(value: u32) -> Result<Self> {
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
            10 => Ok(Self::Presentation),
            11 => Ok(Self::Slide),
            12 => Ok(Self::Project),
            13 => Ok(Self::NoteIt),
            14 => Ok(Self::ExcelChart),
            15 => Ok(Self::MediaPlayer),
            _ => corrupted("ExOleObjAtom contains an invalid subtype"),
        }
    }
}

impl Metadata {
    pub fn parse(record: &Record) -> Result<Self> {
        require_atom(
            record,
            1,
            0,
            RecordType::ExternalOleObjectAtom,
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

    pub(crate) fn validate(&self) -> Result<()> {
        if self.id == 0 {
            return corrupted("ExOleObjAtom object ID must be positive");
        }
        if self.persist_id == 0 {
            return corrupted("ExOleObjAtom persist ID must be positive");
        }
        Ok(())
    }

    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
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
        record_bytes(1, 0, RecordType::ExternalOleObjectAtom, &data)
    }
}

impl ColorFollow {
    pub(crate) fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::EntireScheme),
            2 => Ok(Self::TextAndBackground),
            _ => corrupted("ExOleEmbedAtom contains an invalid color-follow value"),
        }
    }
}

impl DimensionPolicy {
    pub(crate) fn parse(value: u8) -> Self {
        match value {
            0 => Self::Send,
            1 => Self::Omit,
            value => Self::ProducerDefined(value),
        }
    }

    pub(crate) fn value(self) -> u8 {
        match self {
            Self::Send => 0,
            Self::Omit => 1,
            Self::ProducerDefined(value) => value,
        }
    }
}

impl EmbedPreferences {
    pub fn parse(record: &Record) -> Result<Self> {
        require_atom(
            record,
            0,
            0,
            RecordType::ExternalOleEmbedAtom,
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

    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let mut data = [0; 8];
        data[0..4].copy_from_slice(&(self.color_follow as u32).to_le_bytes());
        data[4] = self.cannot_lock_server as u8;
        data[5] = self.dimension_policy.value();
        data[6] = self.is_word_table as u8;
        data[7] = self.unused;
        record_bytes(0, 0, RecordType::ExternalOleEmbedAtom, &data)
    }
}

impl UpdateMode {
    pub(crate) fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Always),
            1 => Ok(Self::OnCall),
            _ => corrupted("ExOleLinkAtom contains an invalid update mode"),
        }
    }
}

impl LinkInfo {
    pub fn parse(record: &Record) -> Result<Self> {
        require_atom(
            record,
            0,
            0,
            RecordType::ExternalOleLinkAtom,
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

    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let mut data = [0; 12];
        data[0..4].copy_from_slice(&self.slide_id.unwrap_or(0).to_le_bytes());
        data[4..8].copy_from_slice(&(self.update_mode as u32).to_le_bytes());
        data[8..12].copy_from_slice(&self.unused);
        record_bytes(0, 0, RecordType::ExternalOleLinkAtom, &data)
    }
}

impl Control {
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != 0x0f
            || record.instance != 0
            || record.record_type != RecordType::ExternalOleControl
        {
            return corrupted("ExControlContainer has an invalid header");
        }
        let children = Record::parse_sequence_strict(&record.data, "ExControlContainer")?;
        if !(2..=6).contains(&children.len()) {
            return corrupted("ExControlContainer has an invalid child count");
        }
        require_atom(
            &children[0],
            0,
            0,
            RecordType::ExternalOleControlAtom,
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

    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        if self.object.object_type != ObjectType::ActiveXControl {
            return corrupted("ExControlContainer requires an ActiveX ExOleObjAtom");
        }
        let mut children = record_bytes(
            0,
            0,
            RecordType::ExternalOleControlAtom,
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
        record_bytes(0x0f, 0, RecordType::ExternalOleControl, &children)
    }
}

impl Definition {
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != 0x0f || record.instance != 0 {
            return corrupted("OLE object container has an invalid header");
        }
        let expected_type = match record.record_type {
            RecordType::ExternalOleEmbed => ObjectType::Embedded,
            RecordType::ExternalOleLink => ObjectType::Linked,
            _ => return corrupted("OLE object container has an invalid record type"),
        };
        let children = Record::parse_sequence_strict(&record.data, "OLE object container")?;
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
        let (menu_name, program_id, clipboard_name, metafile) =
            parse_optional_ole_children(&children[2..])?;
        Ok(Self {
            kind,
            object,
            menu_name,
            program_id,
            clipboard_name,
            metafile,
        })
    }

    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let (container_type, expected_type, first) = match self.kind {
            ContainerKind::Embedded(value) => (
                RecordType::ExternalOleEmbed,
                ObjectType::Embedded,
                value.to_record_bytes()?,
            ),
            ContainerKind::Linked(value) => (
                RecordType::ExternalOleLink,
                ObjectType::Linked,
                value.to_record_bytes()?,
            ),
        };
        if self.object.object_type != expected_type {
            return corrupted("OLE container type disagrees with ExOleObjAtom");
        }
        let mut children = first;
        children.extend_from_slice(&self.object.to_record_bytes()?);
        append_optional_ole_children(
            &mut children,
            self.menu_name.as_deref(),
            self.program_id.as_deref(),
            self.clipboard_name.as_deref(),
            self.metafile.as_deref(),
        )?;
        record_bytes(0x0f, 0, container_type, &children)
    }
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

impl UnknownRecord {
    pub fn record_type(&self) -> u16 {
        self.record.record_type_raw
    }

    pub fn version(&self) -> u16 {
        self.record.version
    }

    pub fn instance(&self) -> u16 {
        self.record.instance
    }

    pub fn data(&self) -> &[u8] {
        &self.record.data
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        record_bytes_raw(
            self.record.version,
            self.record.instance,
            self.record.record_type_raw,
            &self.record.data,
        )
    }

    pub(crate) fn from_record(record: &Record, object_index: usize) -> Self {
        Self {
            record: record.clone(),
            object_index,
        }
    }
}

impl Collection {
    pub fn parse(root: &Record) -> Result<Option<Self>> {
        let mut lists = Vec::new();
        collect_external_object_lists(root, &mut lists);
        if lists.len() > 1 {
            return corrupted("record tree contains multiple external-object lists");
        }
        let Some(list) = lists.first() else {
            return Ok(None);
        };
        if list.version != 0x0f || list.instance != 0 || list.record_type != RecordType::ExObjList {
            return corrupted("ExObjListContainer has an invalid header");
        }
        let children = Record::parse_sequence_strict(&list.data, "ExObjListContainer")?;
        let Some(atom) = children.first() else {
            return corrupted("ExObjListContainer is missing ExObjListAtom");
        };
        require_atom(atom, 0, 0, RecordType::ExObjListAtom, 4, "ExObjListAtom")?;
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
                RecordType::ExternalOleEmbed
                    | RecordType::ExternalOleLink
                    | RecordType::ExternalOleControl
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
                RecordType::ExternalOleControl => {
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

    pub fn unknown_records(&self) -> &[UnknownRecord] {
        &self.unknown_records
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.validate()?;
        let seed = i32::try_from(self.id_seed)
            .map_err(|_| Error::Corrupted("ExObjList identifier seed exceeds i32".into()))?;
        let mut children = record_bytes(0, 0, RecordType::ExObjListAtom, &seed.to_le_bytes())?;
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
        record_bytes(0x0f, 0, RecordType::ExObjList, &children)
    }
}

#[allow(clippy::type_complexity)]
pub(crate) fn parse_optional_ole_children(
    children: &[Record],
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
        if child.record_type == RecordType::CString {
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
        } else if child.record_type == RecordType::MetaFile {
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

pub(crate) fn append_optional_ole_children(
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
                RecordType::CString,
                &encode_ole_string(value, printable)?,
            )?);
        }
    }
    if let Some(metafile) = metafile {
        if metafile.len() > MAX_METAFILE_BYTES {
            return corrupted("MetafileBlob exceeds 64 MiB");
        }
        children.extend_from_slice(&record_bytes(0, 0, RecordType::MetaFile, metafile)?);
    }
    Ok(())
}

fn collect_external_object_lists<'a>(record: &'a Record, lists: &mut Vec<&'a Record>) {
    if record.record_type == RecordType::ExObjList {
        lists.push(record);
    }
    for child in &record.children {
        collect_external_object_lists(child, lists);
    }
}

fn parse_ole_string(record: &Record, printable: bool) -> Result<String> {
    if record.version != 0
        || record.record_type != RecordType::CString
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
        .map_err(|_| Error::Corrupted("OLE object string contains invalid UTF-16".into()))?;
    if printable && value.chars().any(char::is_control) {
        return corrupted("OLE object printable string contains a control character");
    }
    Ok(value)
}

pub(crate) fn encode_ole_string(value: &str, printable: bool) -> Result<Vec<u8>> {
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

pub(crate) fn require_atom(
    record: &Record,
    version: u16,
    instance: u16,
    kind: RecordType,
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

pub(crate) fn record_bytes(
    version: u16,
    instance: u16,
    kind: RecordType,
    data: &[u8],
) -> Result<Vec<u8>> {
    record_bytes_raw(version, instance, kind.as_u16(), data)
}

pub(crate) fn record_bytes_raw(
    version: u16,
    instance: u16,
    kind: u16,
    data: &[u8],
) -> Result<Vec<u8>> {
    if version > 0x000f || instance > 0x0fff {
        return corrupted("PowerPoint record header exceeds its encoded domain");
    }
    let length = u32::try_from(data.len())
        .map_err(|_| Error::Corrupted("PowerPoint record payload exceeds u32".into()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

pub(crate) fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
