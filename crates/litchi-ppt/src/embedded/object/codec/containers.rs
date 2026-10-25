//! Binary codecs for embedded, linked, and `ActiveX` object containers.

use super::super::model::{
    ContainerKind, Control, Definition, EmbedPreferences, ExternalObject, LinkInfo, Metadata,
    ObjectType,
};
use super::strings::{
    append_optional_ole_children, optional_ole_children_len, parse_optional_ole_children,
};
use super::wire::{
    corrupted, ensure_output_total, ensure_record_size, record_bytes, record_bytes_with_limits,
    require_atom, u32_at,
};
use crate::consts::RecordType;
use crate::package::{Error, Result, SourceLimits};
use crate::records::{Record, RecordParseSession};
use std::borrow::Cow;

/// Use the child tree that the owning source parse already materialized.
///
/// OLE containers are real PPT containers, so `Record::parse` now visits
/// them under the source parse budget.  Falling back to a strict sequence is
/// retained for synthetic records assembled by callers, but normal source
/// projections must not restart a fresh event/depth/payload budget for every
/// OLE object.
fn children<'a>(
    record: &'a Record,
    context: &str,
    session: &mut RecordParseSession,
    logical_depth: usize,
) -> Result<Cow<'a, [Record]>> {
    if !record.children.is_empty() || record.data.is_empty() {
        Ok(Cow::Borrowed(&record.children))
    } else {
        let child_depth = logical_depth
            .checked_add(1)
            .ok_or_else(|| Error::Corrupted("OLE container depth overflows".into()))?;
        Ok(Cow::Owned(session.parse_sequence(
            &record.data,
            context,
            child_depth,
        )?))
    }
}

impl Control {
    pub(crate) fn record_len(&self, limits: &SourceLimits) -> Result<usize> {
        if self.object.object_type != ObjectType::ActiveXControl {
            return corrupted("ExControlContainer requires an ActiveX ExOleObjAtom");
        }
        let optional_len = optional_ole_children_len(
            self.menu_name.as_deref(),
            self.program_id.as_deref(),
            self.clipboard_name.as_deref(),
            self.metafile.as_deref(),
            limits,
        )?;
        let payload_len = 12usize
            .checked_add(32)
            .and_then(|value| value.checked_add(optional_len))
            .ok_or_else(|| Error::Corrupted("ExControlContainer length overflows".into()))?;
        ensure_record_size(payload_len, limits, "ExControlContainer")?;
        ensure_output_total(payload_len, limits, "ExControlContainer")?;
        payload_len
            .checked_add(8)
            .ok_or_else(|| Error::Corrupted("ExControlContainer length overflows".into()))
    }

    /// Parse an `ExControlContainer` record into an `ActiveX` control
    /// definition.
    ///
    /// # Errors
    ///
    /// Returns an error if the container header, child count, or any child
    /// record is invalid, or if the carried `ExOleObjAtom` does not describe
    /// an `ActiveX` control.
    pub fn parse(record: &Record, limits: &SourceLimits) -> Result<Self> {
        let mut session = RecordParseSession::new(*limits, record.data.len())?;
        session.account_existing_tree(record, 0)?;
        Self::parse_with_session(record, limits, &mut session, 0)
    }

    pub(crate) fn parse_with_session(
        record: &Record,
        limits: &SourceLimits,
        session: &mut RecordParseSession,
        logical_depth: usize,
    ) -> Result<Self> {
        if record.version != 0x0f
            || record.instance != 0
            || record.record_type != RecordType::ExternalOleControl
        {
            return corrupted("ExControlContainer has an invalid header");
        }
        ensure_record_size(record.data.len(), limits, "ExControlContainer")?;
        let children = children(record, "ExControlContainer", session, logical_depth)?;
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
            parse_optional_ole_children(&children[2..], limits)?;
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
    pub fn to_record(&self, limits: &SourceLimits) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes(limits)?, 0, limits)?.0)
    }

    /// Serialize to the raw bytes of a complete `ExControlContainer` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the carried `ExOleObjAtom` does not describe an
    /// `ActiveX` control or if any child record fails to serialize.
    pub fn to_record_bytes(&self, limits: &SourceLimits) -> Result<Vec<u8>> {
        if self.object.object_type != ObjectType::ActiveXControl {
            return corrupted("ExControlContainer requires an ActiveX ExOleObjAtom");
        }
        let optional_len = optional_ole_children_len(
            self.menu_name.as_deref(),
            self.program_id.as_deref(),
            self.clipboard_name.as_deref(),
            self.metafile.as_deref(),
            limits,
        )?;
        let children_len = 12usize
            .checked_add(32)
            .and_then(|value| value.checked_add(optional_len))
            .ok_or_else(|| Error::Corrupted("ExControlContainer length overflows".into()))?;
        ensure_record_size(children_len, limits, "ExControlContainer")?;
        ensure_output_total(children_len, limits, "ExControlContainer")?;
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
            limits,
        )?;
        record_bytes_with_limits(
            0x0f,
            0,
            RecordType::ExternalOleControl,
            &children,
            limits,
            "ExControlContainer",
        )
    }
}

impl Definition {
    pub(crate) fn record_len(&self, limits: &SourceLimits) -> Result<usize> {
        let fixed_first_len: usize = match self.kind {
            ContainerKind::Embedded(_) => 16,
            ContainerKind::Linked(_) => 20,
        };
        let optional_len = optional_ole_children_len(
            self.menu_name.as_deref(),
            self.program_id.as_deref(),
            self.clipboard_name.as_deref(),
            self.metafile.as_deref(),
            limits,
        )?;
        let payload_len = fixed_first_len
            .checked_add(32)
            .and_then(|value| value.checked_add(optional_len))
            .ok_or_else(|| Error::Corrupted("OLE object container length overflows".into()))?;
        ensure_record_size(payload_len, limits, "OLE object container")?;
        ensure_output_total(payload_len, limits, "OLE object container")?;
        payload_len
            .checked_add(8)
            .ok_or_else(|| Error::Corrupted("OLE object container length overflows".into()))
    }

    /// Parse an embedded or linked OLE object container record.
    ///
    /// # Errors
    ///
    /// Returns an error if the container header, record type, child count, or
    /// any child record is invalid, or if the container type disagrees with
    /// the carried `ExOleObjAtom`.
    pub fn parse(record: &Record, limits: &SourceLimits) -> Result<Self> {
        let mut session = RecordParseSession::new(*limits, record.data.len())?;
        session.account_existing_tree(record, 0)?;
        Self::parse_with_session(record, limits, &mut session, 0)
    }

    pub(crate) fn parse_with_session(
        record: &Record,
        limits: &SourceLimits,
        session: &mut RecordParseSession,
        logical_depth: usize,
    ) -> Result<Self> {
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
        ensure_record_size(record.data.len(), limits, "OLE object container")?;
        let children = children(record, "OLE object container", session, logical_depth)?;
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
            parse_optional_ole_children(&children[2..], limits)?;
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
    pub fn to_record(&self, limits: &SourceLimits) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes(limits)?, 0, limits)?.0)
    }

    /// Serialize to the raw bytes of a complete OLE object container record.
    ///
    /// # Errors
    ///
    /// Returns an error if the container type disagrees with the carried
    /// `ExOleObjAtom` or if any child record fails to serialize.
    pub fn to_record_bytes(&self, limits: &SourceLimits) -> Result<Vec<u8>> {
        let fixed_first_len: usize = match self.kind {
            ContainerKind::Embedded(_) => 16,
            ContainerKind::Linked(_) => 20,
        };
        let optional_len = optional_ole_children_len(
            self.menu_name.as_deref(),
            self.program_id.as_deref(),
            self.clipboard_name.as_deref(),
            self.metafile.as_deref(),
            limits,
        )?;
        let children_len = fixed_first_len
            .checked_add(32)
            .and_then(|value| value.checked_add(optional_len))
            .ok_or_else(|| Error::Corrupted("OLE object container length overflows".into()))?;
        ensure_record_size(children_len, limits, "OLE object container")?;
        ensure_output_total(children_len, limits, "OLE object container")?;
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
            limits,
        )?;
        record_bytes_with_limits(
            0x0f,
            0,
            container_type,
            &children,
            limits,
            "OLE object container",
        )
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
    pub fn to_record_bytes(&self, limits: &SourceLimits) -> Result<Vec<u8>> {
        match self {
            Self::Object(value) => value.to_record_bytes(limits),
            Self::ActiveXControl(value) => value.to_record_bytes(limits),
        }
    }
}
