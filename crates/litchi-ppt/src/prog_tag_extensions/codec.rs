use crate::consts::PptRecordType;
use crate::records::PptRecord;

use super::super::package::{PptError, Result};
use super::super::prog_tags::{
    PowerPointProgBinaryTag, PowerPointProgBinaryTagVersion, PowerPointProgTag,
    PowerPointProgTagScope, PowerPointProgTags,
};
use super::model::{
    PowerPoint9DocBinaryTagExtension, PowerPoint9SlideBinaryTagExtension,
    PowerPoint10DocBinaryTagExtension, PowerPoint10SlideBinaryTagExtension,
    PowerPoint11DocBinaryTagExtension, PowerPoint12DocBinaryTagExtension,
    PowerPoint12SlideBinaryTagExtension, PowerPointDocBinaryTagExtension,
    PowerPointDocumentTagExtensions, PowerPointSlideBinaryTagExtension,
    PowerPointSlideTagExtensions,
};

/// `RT_PresentationAdvisorFlags9Atom` (MS-PPT 2.13.24).
pub(super) const PRES_ADVISOR_FLAGS_9_ATOM: u16 = 0x177a;
/// `RT_HtmlDocInfo9Atom` (MS-PPT 2.13.24).
pub(super) const HTML_DOC_INFO_9_ATOM: u16 = 0x177b;
/// `RT_HtmlPublishInfo9` (MS-PPT 2.13.24).
pub(super) const HTML_PUBLISH_INFO_9: u16 = 0x177d;
/// `RT_BroadcastDocInfo9` (MS-PPT 2.13.24).
pub(super) const BROADCAST_DOC_INFO_9: u16 = 0x177e;
/// `RT_EnvelopeFlags9Atom` (MS-PPT 2.13.24).
pub(super) const ENVELOPE_FLAGS_9_ATOM: u16 = 0x1784;
/// `RT_EnvelopeData9Atom` (MS-PPT 2.13.24).
pub(super) const ENVELOPE_DATA_9_ATOM: u16 = 0x1785;
/// `RT_Comment10` (MS-PPT 2.13.24).
pub(super) const COMMENT_10: u16 = 0x2ee0;

/// `CopyrightAtom` record instance (MS-PPT 2.4.22.1).
pub(super) const COPYRIGHT_INSTANCE: u16 = 0x001;
/// `KeywordsAtom` record instance (MS-PPT 2.4.22.2).
pub(super) const KEYWORDS_INSTANCE: u16 = 0x002;
/// `ModifyPasswordAtom` record instance (MS-PPT 2.4.7).
pub(super) const MODIFY_PASSWORD_INSTANCE: u16 = 0x003;

/// Container record version nibble.
pub(super) const CONTAINER_VERSION: u16 = 0x0f;
/// Atom record version nibble.
pub(super) const ATOM_VERSION: u16 = 0x00;

/// Implement `parse_records`/`to_payload` for an extension grammar.
//
// Grammars are an ordered sequence of greedy arrays (consumed while the
// record type matches), optional single records, and required single
// records, exactly as the spec tables list them. Slots are declared in
// grammar order: `array(label, field, type, version)` for a record-type
// array, `opt(label, field, type, instance, version)` for an optional
// record, and `req(label, field, type, version)` for a required record.
macro_rules! extension_struct {
    (
        $name:ident, $context:literal,
        $($kind:ident($label:literal, $field:ident $(, $args:expr)*)),* $(,)?
    ) => {
        impl $name {
            /// Parse and validate the ordered record sequence of the extension
            /// payload. Records are consumed, not copied.
            pub fn parse_records(records: Vec<PptRecord>) -> Result<Self> {
                let mut cursor = RecordCursor::new(records, $context);
                $(extension_struct!(@parse cursor, $kind($label, $field $(, $args)*));)*
                cursor.finish()?;
                Ok(Self { $($field,)* })
            }

            /// Serialize the extension payload byte-for-byte.
            ///
            /// The encoded payload is reparsed before returning so public-field
            /// mutations cannot serialize an invalid grammar.
            pub fn to_payload(&self) -> Result<Vec<u8>> {
                let mut payload = Vec::new();
                $(extension_struct!(@encode self, payload, $kind, $field);)*
                Self::parse_records(PptRecord::parse_sequence_strict(&payload, $context)?)?;
                Ok(payload)
            }
        }
    };
    (@parse $cursor:ident, array($label:literal, $field:ident, $ty:expr, $version:expr)) => {
        let $field = $cursor.take_array($ty, $version, $label)?;
    };
    (@parse $cursor:ident, opt($label:literal, $field:ident, $ty:expr, $instance:expr, $version:expr)) => {
        let $field = $cursor.take_optional($ty, $instance, $version, $label)?;
    };
    (@parse $cursor:ident, req($label:literal, $field:ident, $ty:expr, $version:expr)) => {
        let $field = Some($cursor.take_required($ty, $version, $label)?);
    };
    (@encode $this:ident, $payload:ident, array, $field:ident) => {
        for record in &$this.$field {
            $payload.extend_from_slice(&encode_record(record)?);
        }
    };
    (@encode $this:ident, $payload:ident, opt, $field:ident) => {
        if let Some(record) = &$this.$field {
            $payload.extend_from_slice(&encode_record(record)?);
        }
    };
    (@encode $this:ident, $payload:ident, req, $field:ident) => {
        if let Some(record) = &$this.$field {
            $payload.extend_from_slice(&encode_record(record)?);
        }
    };
}

extension_struct! {
    PowerPoint9DocBinaryTagExtension, "PP9DocBinaryTagExtension",
    array("TextMasterStyle9Atom", text_master_styles, PptRecordType::TextMasterStyle9Atom.as_u16(), ATOM_VERSION),
    opt("BlipCollection9Container", blip_collection, PptRecordType::BlipCollection9.as_u16(), None, CONTAINER_VERSION),
    opt("TextDefaults9Atom", text_defaults, PptRecordType::TextDefaults9Atom.as_u16(), None, ATOM_VERSION),
    opt("Kinsoku9Container", kinsoku, PptRecordType::Kinsoku.as_u16(), None, CONTAINER_VERSION),
    array("ExHyperlink9Container", external_hyperlinks, PptRecordType::ExternalHyperlink9.as_u16(), CONTAINER_VERSION),
    opt("PresAdvisorFlags9Atom", advisor_flags, PRES_ADVISOR_FLAGS_9_ATOM, None, ATOM_VERSION),
    opt("EnvelopeData9Atom", envelope_data, ENVELOPE_DATA_9_ATOM, None, ATOM_VERSION),
    opt("EnvelopeFlags9Atom", envelope_flags, ENVELOPE_FLAGS_9_ATOM, None, ATOM_VERSION),
    opt("HTMLDocInfo9Atom", html_doc_info, HTML_DOC_INFO_9_ATOM, None, ATOM_VERSION),
    opt("HTMLPublishInfo9Container", html_publish_info, HTML_PUBLISH_INFO_9, None, CONTAINER_VERSION),
    array("BroadcastDocInfo9Container", broadcasts, BROADCAST_DOC_INFO_9, CONTAINER_VERSION),
    opt("OutlineTextProps9Container", outline_text_props, PptRecordType::OutlineTextProps9.as_u16(), None, CONTAINER_VERSION),
}

extension_struct! {
    PowerPoint10DocBinaryTagExtension, "PP10DocBinaryTagExtension",
    opt("FontCollection10Container", font_collection, PptRecordType::FontCollection10.as_u16(), None, CONTAINER_VERSION),
    array("TextMasterStyle10Atom", text_master_styles, PptRecordType::TextMasterStyle10Atom.as_u16(), ATOM_VERSION),
    opt("TextDefaults10Atom", text_defaults, PptRecordType::TextDefaults10Atom.as_u16(), None, ATOM_VERSION),
    req("GridSpacing10Atom", grid_spacing, PptRecordType::GridSpacing10Atom.as_u16(), ATOM_VERSION),
    array("CommentIndex10Container", comment_indices, PptRecordType::CommentIndex10.as_u16(), CONTAINER_VERSION),
    opt("FontEmbedFlags10Atom", font_embed_flags, PptRecordType::FontEmbedFlags10Atom.as_u16(), None, ATOM_VERSION),
    opt("CopyrightAtom", copyright, PptRecordType::CString.as_u16(), Some(COPYRIGHT_INSTANCE), ATOM_VERSION),
    opt("KeywordsAtom", keywords, PptRecordType::CString.as_u16(), Some(KEYWORDS_INSTANCE), ATOM_VERSION),
    opt("FilterPrivacyFlags10Atom", filter_privacy_flags, PptRecordType::FilterPrivacyFlags10Atom.as_u16(), None, ATOM_VERSION),
    opt("OutlineTextProps10Container", outline_text_props, PptRecordType::OutlineTextProps10.as_u16(), None, CONTAINER_VERSION),
    opt("DocToolbarStates10Atom", toolbar_states, PptRecordType::DocToolbarStates10Atom.as_u16(), None, ATOM_VERSION),
    opt("SlideListTable10Container", slide_list_table, PptRecordType::SlideListTable10.as_u16(), None, CONTAINER_VERSION),
    array("DiffTree10Container", diff_trees, PptRecordType::DiffTree10.as_u16(), CONTAINER_VERSION),
    opt("ModifyPasswordAtom", modify_password, PptRecordType::CString.as_u16(), Some(MODIFY_PASSWORD_INSTANCE), ATOM_VERSION),
    opt("PhotoAlbumInfo10Atom", photo_album_info, PptRecordType::PhotoAlbumInfo10Atom.as_u16(), None, ATOM_VERSION),
}

extension_struct! {
    PowerPoint11DocBinaryTagExtension, "PP11DocBinaryTagExtension",
    opt("SmartTagStore11Container", smart_tag_store, PptRecordType::SmartTagStore11.as_u16(), None, CONTAINER_VERSION),
    opt("OutlineTextProps11Container", outline_text_props, PptRecordType::OutlineTextProps11.as_u16(), None, CONTAINER_VERSION),
}

extension_struct! {
    PowerPoint12DocBinaryTagExtension, "PP12DocBinaryTagExtension",
    opt("RoundTripDocFlags12Atom", doc_flags, PptRecordType::RoundTripDocFlags12Atom.as_u16(), None, ATOM_VERSION),
}

extension_struct! {
    PowerPoint9SlideBinaryTagExtension, "PP9SlideBinaryTagExtension",
    array("TextMasterStyle9Atom", text_master_styles, PptRecordType::TextMasterStyle9Atom.as_u16(), ATOM_VERSION),
}

extension_struct! {
    PowerPoint12SlideBinaryTagExtension, "PP12SlideBinaryTagExtension",
    opt("RoundTripHeaderFooterDefaults12Atom", header_footer_defaults, PptRecordType::RoundTripHeaderFooterDefaults12Atom.as_u16(), None, ATOM_VERSION),
}

impl PowerPoint10SlideBinaryTagExtension {
    /// Parse and validate the ordered record sequence of the extension payload.
    pub fn parse_records(records: Vec<PptRecord>) -> Result<Self> {
        const CONTEXT: &str = "PP10SlideBinaryTagExtension";
        let mut cursor = RecordCursor::new(records, CONTEXT);
        let text_master_styles = cursor.take_array(
            PptRecordType::TextMasterStyle10Atom.as_u16(),
            ATOM_VERSION,
            "TextMasterStyle10Atom",
        )?;
        let comments = cursor.take_array(COMMENT_10, CONTAINER_VERSION, "Comment10Container")?;
        let linked_slide = cursor.take_optional(
            PptRecordType::LinkedSlide10Atom.as_u16(),
            None,
            ATOM_VERSION,
            "LinkedSlide10Atom",
        )?;
        let linked_shapes = cursor.take_array(
            PptRecordType::LinkedShape10Atom.as_u16(),
            ATOM_VERSION,
            "LinkedShape10Atom",
        )?;
        // Section 2.5.24: rgLinkedShape10Atom is counted by
        // linkedSlideAtom.cLinkedShapes and cannot appear without the atom.
        match &linked_slide {
            Some(atom) => {
                let data: [u8; 8] = atom.data.as_slice().try_into().map_err(|_| {
                    PptError::Corrupted("LinkedSlide10Atom payload must be 8 bytes".into())
                })?;
                let count = i32::from_le_bytes([data[4], data[5], data[6], data[7]]);
                if count < 0 || linked_shapes.len() != count as usize {
                    return corrupted(
                        "LinkedShape10Atom count does not match LinkedSlide10Atom.cLinkedShapes",
                    );
                }
            },
            None if !linked_shapes.is_empty() => {
                return corrupted("LinkedShape10Atom array requires a LinkedSlide10Atom");
            },
            None => {},
        }
        let slide_flags = cursor.take_optional(
            PptRecordType::SlideFlags10Atom.as_u16(),
            None,
            ATOM_VERSION,
            "SlideFlags10Atom",
        )?;
        let slide_time = cursor.take_optional(
            PptRecordType::SlideTime10Atom.as_u16(),
            None,
            ATOM_VERSION,
            "SlideTime10Atom",
        )?;
        let hash_code = cursor.take_optional(
            PptRecordType::HashCode10Atom.as_u16(),
            None,
            ATOM_VERSION,
            "HashCode10Atom",
        )?;
        let timing = cursor.take_optional(
            PptRecordType::ExtTimeNode.as_u16(),
            None,
            CONTAINER_VERSION,
            "ExtTimeNodeContainer",
        )?;
        let build_list = cursor.take_optional(
            PptRecordType::BuildList.as_u16(),
            None,
            CONTAINER_VERSION,
            "BuildListContainer",
        )?;
        cursor.finish()?;
        Ok(Self {
            text_master_styles,
            comments,
            linked_slide,
            linked_shapes,
            slide_flags,
            slide_time,
            hash_code,
            timing,
            build_list,
        })
    }

    /// Serialize the extension payload byte-for-byte.
    ///
    /// The encoded payload is reparsed before returning so public-field
    /// mutations cannot serialize an invalid grammar.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        let mut payload = Vec::new();
        for record in &self.text_master_styles {
            payload.extend_from_slice(&encode_record(record)?);
        }
        for record in &self.comments {
            payload.extend_from_slice(&encode_record(record)?);
        }
        if let Some(record) = &self.linked_slide {
            payload.extend_from_slice(&encode_record(record)?);
        }
        for record in &self.linked_shapes {
            payload.extend_from_slice(&encode_record(record)?);
        }
        for record in [
            &self.slide_flags,
            &self.slide_time,
            &self.hash_code,
            &self.timing,
            &self.build_list,
        ]
        .into_iter()
        .flatten()
        {
            payload.extend_from_slice(&encode_record(record)?);
        }
        Self::parse_records(PptRecord::parse_sequence_strict(
            &payload,
            "PP10SlideBinaryTagExtension",
        )?)?;
        Ok(payload)
    }
}

impl PowerPointDocBinaryTagExtension {
    /// Decode a versioned document-scope binary tag payload. Returns `Ok(None)`
    /// for unassigned (unknown) tags.
    pub fn parse(
        version: PowerPointProgBinaryTagVersion,
        records: Vec<PptRecord>,
    ) -> Result<Option<Self>> {
        match version {
            PowerPointProgBinaryTagVersion::PowerPoint9 => Ok(Some(Self::PowerPoint9(Box::new(
                PowerPoint9DocBinaryTagExtension::parse_records(records)?,
            )))),
            PowerPointProgBinaryTagVersion::PowerPoint10 => Ok(Some(Self::PowerPoint10(Box::new(
                PowerPoint10DocBinaryTagExtension::parse_records(records)?,
            )))),
            PowerPointProgBinaryTagVersion::PowerPoint11 => Ok(Some(Self::PowerPoint11(Box::new(
                PowerPoint11DocBinaryTagExtension::parse_records(records)?,
            )))),
            PowerPointProgBinaryTagVersion::PowerPoint12 => Ok(Some(Self::PowerPoint12(Box::new(
                PowerPoint12DocBinaryTagExtension::parse_records(records)?,
            )))),
            PowerPointProgBinaryTagVersion::Unknown => Ok(None),
        }
    }

    /// Serialize the tag payload byte-for-byte.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        match self {
            Self::PowerPoint9(extension) => extension.to_payload(),
            Self::PowerPoint10(extension) => extension.to_payload(),
            Self::PowerPoint11(extension) => extension.to_payload(),
            Self::PowerPoint12(extension) => extension.to_payload(),
        }
    }
}

impl PowerPointSlideBinaryTagExtension {
    /// Decode a versioned slide-scope binary tag payload. Returns `Ok(None)`
    /// for unassigned (unknown) tags.
    pub fn parse(
        version: PowerPointProgBinaryTagVersion,
        records: Vec<PptRecord>,
    ) -> Result<Option<Self>> {
        match version {
            PowerPointProgBinaryTagVersion::PowerPoint9 => Ok(Some(Self::PowerPoint9(Box::new(
                PowerPoint9SlideBinaryTagExtension::parse_records(records)?,
            )))),
            PowerPointProgBinaryTagVersion::PowerPoint10 => Ok(Some(Self::PowerPoint10(Box::new(
                PowerPoint10SlideBinaryTagExtension::parse_records(records)?,
            )))),
            PowerPointProgBinaryTagVersion::PowerPoint12 => Ok(Some(Self::PowerPoint12(Box::new(
                PowerPoint12SlideBinaryTagExtension::parse_records(records)?,
            )))),
            PowerPointProgBinaryTagVersion::PowerPoint11 => {
                corrupted("___PPT11 is not an assigned slide-scope binary tag")
            },
            PowerPointProgBinaryTagVersion::Unknown => Ok(None),
        }
    }

    /// Serialize the tag payload byte-for-byte.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        match self {
            Self::PowerPoint9(extension) => extension.to_payload(),
            Self::PowerPoint10(extension) => extension.to_payload(),
            Self::PowerPoint12(extension) => extension.to_payload(),
        }
    }
}

impl PowerPointProgBinaryTag {
    /// Decode this tag's payload as a versioned document extension.
    ///
    /// Returns `Ok(None)` for unassigned (unknown) tags, whose payloads are
    /// preserved without interpretation.
    pub fn doc_extension(&self) -> Result<Option<PowerPointDocBinaryTagExtension>> {
        PowerPointDocBinaryTagExtension::parse(self.version, self.records()?)
    }

    /// Decode this tag's payload as a versioned slide extension.
    ///
    /// Returns `Ok(None)` for unassigned (unknown) tags. `___PPT11` is not
    /// assigned at slide scope (MS-PPT 2.5.22), so decoding it as a slide
    /// extension is an error.
    pub fn slide_extension(&self) -> Result<Option<PowerPointSlideBinaryTagExtension>> {
        PowerPointSlideBinaryTagExtension::parse(self.version, self.records()?)
    }
}

impl PowerPointProgTags {
    /// Decode every assigned versioned document extension in this container.
    ///
    /// Unknown tags are skipped; their payloads remain available through
    /// [`PowerPointProgTags::binary_tag`].
    pub fn document_extensions(&self) -> Result<PowerPointDocumentTagExtensions> {
        if self.scope != PowerPointProgTagScope::Document {
            return corrupted("slide-scope ProgTags cannot hold document extensions");
        }
        let mut extensions = PowerPointDocumentTagExtensions::default();
        for tag in &self.tags {
            let PowerPointProgTag::Binary(tag) = tag else {
                continue;
            };
            match tag.doc_extension()? {
                Some(PowerPointDocBinaryTagExtension::PowerPoint9(extension)) => {
                    extensions.powerpoint9 = Some(*extension);
                },
                Some(PowerPointDocBinaryTagExtension::PowerPoint10(extension)) => {
                    extensions.powerpoint10 = Some(*extension);
                },
                Some(PowerPointDocBinaryTagExtension::PowerPoint11(extension)) => {
                    extensions.powerpoint11 = Some(*extension);
                },
                Some(PowerPointDocBinaryTagExtension::PowerPoint12(extension)) => {
                    extensions.powerpoint12 = Some(*extension);
                },
                None => {},
            }
        }
        Ok(extensions)
    }

    /// Decode every assigned versioned slide extension in this container.
    ///
    /// Unknown tags are skipped; their payloads remain available through
    /// [`PowerPointProgTags::binary_tag`].
    pub fn slide_extensions(&self) -> Result<PowerPointSlideTagExtensions> {
        if self.scope != PowerPointProgTagScope::Slide {
            return corrupted("document-scope ProgTags cannot hold slide extensions");
        }
        let mut extensions = PowerPointSlideTagExtensions::default();
        for tag in &self.tags {
            let PowerPointProgTag::Binary(tag) = tag else {
                continue;
            };
            match tag.slide_extension()? {
                Some(PowerPointSlideBinaryTagExtension::PowerPoint9(extension)) => {
                    extensions.powerpoint9 = Some(*extension);
                },
                Some(PowerPointSlideBinaryTagExtension::PowerPoint10(extension)) => {
                    extensions.powerpoint10 = Some(*extension);
                },
                Some(PowerPointSlideBinaryTagExtension::PowerPoint12(extension)) => {
                    extensions.powerpoint12 = Some(*extension);
                },
                None => {},
            }
        }
        Ok(extensions)
    }
}

/// Owning cursor over an extension record sequence.
struct RecordCursor {
    records: std::iter::Peekable<std::vec::IntoIter<PptRecord>>,
    context: &'static str,
}

impl RecordCursor {
    fn new(records: Vec<PptRecord>, context: &'static str) -> Self {
        Self {
            records: records.into_iter().peekable(),
            context,
        }
    }

    /// Consume records while they match the array element type, validating the
    /// version nibble of each element.
    fn take_array(&mut self, kind: u16, version: u16, label: &str) -> Result<Vec<PptRecord>> {
        let mut result = Vec::new();
        while let Some(record) = self
            .records
            .next_if(|record| record.record_type_raw == kind)
        {
            if record.version != version {
                return corrupted(format!(
                    "{label} in {} has an invalid record header",
                    self.context
                ));
            }
            result.push(record);
        }
        Ok(result)
    }

    /// Consume the next record when it matches the optional slot. A record of
    /// the slot's type with the wrong instance or version is an error, since
    /// per the grammar it can only belong to this slot.
    fn take_optional(
        &mut self,
        kind: u16,
        instance: Option<u16>,
        version: u16,
        label: &str,
    ) -> Result<Option<PptRecord>> {
        let Some(record) = self
            .records
            .next_if(|record| record.record_type_raw == kind)
        else {
            return Ok(None);
        };
        if instance.is_some_and(|expected| record.instance != expected) || record.version != version
        {
            return corrupted(format!(
                "{label} in {} has an invalid record header",
                self.context
            ));
        }
        Ok(Some(record))
    }

    fn take_required(&mut self, kind: u16, version: u16, label: &str) -> Result<PptRecord> {
        match self.take_optional(kind, None, version, label)? {
            Some(record) => Ok(record),
            None => corrupted(format!("{} is missing its required {label}", self.context)),
        }
    }

    fn finish(&mut self) -> Result<()> {
        if self.records.next().is_some() {
            return corrupted(format!(
                "{} contains a record outside its grammar",
                self.context
            ));
        }
        Ok(())
    }
}

/// Re-encode one parsed record byte-for-byte from its header fields and payload.
fn encode_record(record: &PptRecord) -> Result<Vec<u8>> {
    if record.version > 0x0f || record.instance > 0x0fff {
        return corrupted("PPT record version or instance exceeds its bit field");
    }
    let declared = usize::try_from(record.data_length)
        .map_err(|_| PptError::Corrupted("PPT record length overflow".into()))?;
    if declared != record.data.len() {
        return corrupted("PPT record length does not match its payload");
    }
    let length = u32::try_from(record.data.len())
        .map_err(|_| PptError::Corrupted("PPT record payload exceeds u32".into()))?;
    let mut result = Vec::with_capacity(8usize.saturating_add(record.data.len()));
    result.extend_from_slice(&((record.instance << 4) | record.version).to_le_bytes());
    result.extend_from_slice(&record.record_type_raw.to_le_bytes());
    result.extend_from_slice(&length.to_le_bytes());
    result.extend_from_slice(&record.data);
    Ok(result)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}
