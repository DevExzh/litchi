//! Binary codecs for embedded, linked, and `ActiveX` object containers.

use super::super::model::{
    ContainerKind, Control, Definition, EmbedPreferences, ExternalObject, LinkInfo, Metadata,
    ObjectType,
};
use super::strings::{append_optional_ole_children, parse_optional_ole_children};
use super::wire::{corrupted, record_bytes, require_atom, u32_at};
use crate::consts::RecordType;
use crate::package::Result;
use crate::records::Record;

impl Control {
    /// Parse an `ExControlContainer` record into an `ActiveX` control
    /// definition.
    ///
    /// # Errors
    ///
    /// Returns an error if the container header, child count, or any child
    /// record is invalid, or if the carried `ExOleObjAtom` does not describe
    /// an `ActiveX` control.
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

    /// Serialize as a fully parsed `ExControlContainer` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the control fails to serialize or if the encoded
    /// record cannot be re-parsed.
    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    /// Serialize to the raw bytes of a complete `ExControlContainer` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the carried `ExOleObjAtom` does not describe an
    /// `ActiveX` control or if any child record fails to serialize.
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
    /// Parse an embedded or linked OLE object container record.
    ///
    /// # Errors
    ///
    /// Returns an error if the container header, record type, child count, or
    /// any child record is invalid, or if the container type disagrees with
    /// the carried `ExOleObjAtom`.
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != 0x0f || record.instance != 0 {
            return corrupted("OLE object container has an invalid header");
        }
        let expected_type = if record.record_type == RecordType::ExternalOleEmbed {
            ObjectType::Embedded
        } else if record.record_type == RecordType::ExternalOleLink {
            ObjectType::Linked
        } else {
            return corrupted("OLE object container has an invalid record type");
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

    /// Serialize as a fully parsed OLE object container record.
    ///
    /// # Errors
    ///
    /// Returns an error if the definition fails to serialize or if the
    /// encoded record cannot be re-parsed.
    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    /// Serialize to the raw bytes of a complete OLE object container record.
    ///
    /// # Errors
    ///
    /// Returns an error if the container type disagrees with the carried
    /// `ExOleObjAtom` or if any child record fails to serialize.
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
    #[must_use]
    pub fn id(&self) -> u32 {
        match self {
            Self::Object(value) => value.object.id,
            Self::ActiveXControl(value) => value.object.id,
        }
    }

    #[must_use]
    pub fn persist_id(&self) -> u32 {
        match self {
            Self::Object(value) => value.object.persist_id,
            Self::ActiveXControl(value) => value.object.persist_id,
        }
    }

    /// Serialize to the raw bytes of the object's container record.
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying object definition fails to
    /// serialize.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        match self {
            Self::Object(value) => value.to_record_bytes(),
            Self::ActiveXControl(value) => value.to_record_bytes(),
        }
    }
}
