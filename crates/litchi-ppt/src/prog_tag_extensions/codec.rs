use crate::consts::RecordType;
use crate::records::Record;

use super::super::package::{Error, Result};
use super::super::prog_tags::{
    ProgBinaryTag, ProgBinaryTagVersion, ProgTag, ProgTagScope, ProgTags,
};
use super::model::{
    DocBinaryTagExtension, DocBinaryTagExtension9, DocBinaryTagExtension10,
    DocBinaryTagExtension11, DocBinaryTagExtension12, DocumentTagExtensions,
    SlideBinaryTagExtension, SlideBinaryTagExtension9, SlideBinaryTagExtension10,
    SlideBinaryTagExtension12, SlideTagExtensions,
};

/// Implement `parse_records`/`to_payload` for an extension grammar.
//
// Grammars are an ordered sequence of greedy arrays (consumed while the
// record type matches), optional single records, and required single
// records, exactly as the spec tables list them. Slots are declared in
// grammar order: `array(label, field, type, version)` for a record-type
// array, `opt(label, field, type, instance, version)` for an optional
// record.
#[allow(
    unused_macro_rules,
    reason = "schema variants use different required/optional field combinations"
)]
macro_rules! extension_struct {
    (
        $name:ident, $context:literal,
        $($kind:ident($label:literal, $field:ident $(, $args:expr)*)),* $(,)?
    ) => {
        impl $name {
            /// Parse and validate the ordered record sequence of the extension
            /// payload. Records are consumed, not copied.
            ///
            /// # Errors
            ///
            /// Returns an error if the operation fails.
            pub fn parse_records(records: Vec<Record>) -> Result<Self> {
                let mut cursor = RecordCursor::new(records, $context);
                $(extension_struct!(@parse cursor, $kind($label, $field $(, $args)*));)*
                cursor.finish()?;
                Ok(Self { $($field,)* })
            }

            /// Serialize the extension payload byte-for-byte.
            ///
            /// The encoded payload is reparsed before returning so public-field
            /// mutations cannot serialize an invalid grammar.
            ///
            /// # Errors
            ///
            /// Returns an error if the operation fails.
            pub fn to_payload(&self) -> Result<Vec<u8>> {
                let mut payload = Vec::new();
                $(extension_struct!(@encode self, payload, $kind, $field);)*
                Self::parse_records(Record::parse_sequence_strict(&payload, $context)?)?;
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
}

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

extension_struct! {
    DocBinaryTagExtension9, "PP9DocBinaryTagExtension",
    array("TextMasterStyle9Atom", text_master_styles, RecordType::TextMasterStyle9Atom.as_u16(), ATOM_VERSION),
    opt("BlipCollection9Container", blip_collection, RecordType::BlipCollection9.as_u16(), None, CONTAINER_VERSION),
    opt("TextDefaults9Atom", text_defaults, RecordType::TextDefaults9Atom.as_u16(), None, ATOM_VERSION),
    opt("Kinsoku9Container", kinsoku, RecordType::Kinsoku.as_u16(), None, CONTAINER_VERSION),
    array("ExHyperlink9Container", external_hyperlinks, RecordType::ExternalHyperlink9.as_u16(), CONTAINER_VERSION),
    opt("PresAdvisorFlags9Atom", advisor_flags, PRES_ADVISOR_FLAGS_9_ATOM, None, ATOM_VERSION),
    opt("EnvelopeData9Atom", envelope_data, ENVELOPE_DATA_9_ATOM, None, ATOM_VERSION),
    opt("EnvelopeFlags9Atom", envelope_flags, ENVELOPE_FLAGS_9_ATOM, None, ATOM_VERSION),
    opt("HTMLDocInfo9Atom", html_doc_info, HTML_DOC_INFO_9_ATOM, None, ATOM_VERSION),
    opt("HTMLPublishInfo9Container", html_publish_info, HTML_PUBLISH_INFO_9, None, CONTAINER_VERSION),
    array("BroadcastDocInfo9Container", broadcasts, BROADCAST_DOC_INFO_9, CONTAINER_VERSION),
    opt("OutlineTextProps9Container", outline_text_props, RecordType::OutlineTextProps9.as_u16(), None, CONTAINER_VERSION),
}

extension_struct! {
    DocBinaryTagExtension10, "PP10DocBinaryTagExtension",
    opt("FontCollection10Container", font_collection, RecordType::FontCollection10.as_u16(), None, CONTAINER_VERSION),
    array("TextMasterStyle10Atom", text_master_styles, RecordType::TextMasterStyle10Atom.as_u16(), ATOM_VERSION),
    opt("TextDefaults10Atom", text_defaults, RecordType::TextDefaults10Atom.as_u16(), None, ATOM_VERSION),
    opt("GridSpacing10Atom", grid_spacing, RecordType::GridSpacing10Atom.as_u16(), None, ATOM_VERSION),
    array("CommentIndex10Container", comment_indices, RecordType::CommentIndex10.as_u16(), CONTAINER_VERSION),
    opt("FontEmbedFlags10Atom", font_embed_flags, RecordType::FontEmbedFlags10Atom.as_u16(), None, ATOM_VERSION),
    opt("CopyrightAtom", copyright, RecordType::CString.as_u16(), Some(COPYRIGHT_INSTANCE), ATOM_VERSION),
    opt("KeywordsAtom", keywords, RecordType::CString.as_u16(), Some(KEYWORDS_INSTANCE), ATOM_VERSION),
    opt("FilterPrivacyFlags10Atom", filter_privacy_flags, RecordType::FilterPrivacyFlags10Atom.as_u16(), None, ATOM_VERSION),
    opt("OutlineTextProps10Container", outline_text_props, RecordType::OutlineTextProps10.as_u16(), None, CONTAINER_VERSION),
    opt("DocToolbarStates10Atom", toolbar_states, RecordType::DocToolbarStates10Atom.as_u16(), None, ATOM_VERSION),
    opt("SlideListTable10Container", slide_list_table, RecordType::SlideListTable10.as_u16(), None, CONTAINER_VERSION),
    array("DiffTree10Container", diff_trees, RecordType::DiffTree10.as_u16(), CONTAINER_VERSION),
    opt("ModifyPasswordAtom", modify_password, RecordType::CString.as_u16(), Some(MODIFY_PASSWORD_INSTANCE), ATOM_VERSION),
    opt("PhotoAlbumInfo10Atom", photo_album_info, RecordType::PhotoAlbumInfo10Atom.as_u16(), None, ATOM_VERSION),
}

extension_struct! {
    DocBinaryTagExtension11, "PP11DocBinaryTagExtension",
    opt("SmartTagStore11Container", smart_tag_store, RecordType::SmartTagStore11.as_u16(), None, CONTAINER_VERSION),
    opt("OutlineTextProps11Container", outline_text_props, RecordType::OutlineTextProps11.as_u16(), None, CONTAINER_VERSION),
}

extension_struct! {
    DocBinaryTagExtension12, "PP12DocBinaryTagExtension",
    opt("RoundTripDocFlags12Atom", doc_flags, RecordType::RoundTripDocFlags12Atom.as_u16(), None, ATOM_VERSION),
}

extension_struct! {
    SlideBinaryTagExtension9, "PP9SlideBinaryTagExtension",
    array("TextMasterStyle9Atom", text_master_styles, RecordType::TextMasterStyle9Atom.as_u16(), ATOM_VERSION),
}

extension_struct! {
    SlideBinaryTagExtension12, "PP12SlideBinaryTagExtension",
    opt("RoundTripHeaderFooterDefaults12Atom", header_footer_defaults, RecordType::RoundTripHeaderFooterDefaults12Atom.as_u16(), None, ATOM_VERSION),
}

impl SlideBinaryTagExtension10 {
    /// Parse and validate the ordered record sequence of the extension payload.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_records(records: Vec<Record>) -> Result<Self> {
        const CONTEXT: &str = "PP10SlideBinaryTagExtension";
        let mut cursor = RecordCursor::new(records, CONTEXT);
        let text_master_styles = cursor.take_array(
            RecordType::TextMasterStyle10Atom.as_u16(),
            ATOM_VERSION,
            "TextMasterStyle10Atom",
        )?;
        let comments = cursor.take_array(COMMENT_10, CONTAINER_VERSION, "Comment10Container")?;
        let linked_slide = cursor.take_optional(
            RecordType::LinkedSlide10Atom.as_u16(),
            None,
            ATOM_VERSION,
            "LinkedSlide10Atom",
        )?;
        let linked_shapes = cursor.take_array(
            RecordType::LinkedShape10Atom.as_u16(),
            ATOM_VERSION,
            "LinkedShape10Atom",
        )?;
        // Section 2.5.24: rgLinkedShape10Atom is counted by
        // linkedSlideAtom.cLinkedShapes and cannot appear without the atom.
        match &linked_slide {
            Some(atom) => {
                let data: [u8; 8] = atom.data.as_slice().try_into().map_err(|_err| {
                    Error::Corrupted("LinkedSlide10Atom payload must be 8 bytes".into())
                })?;
                let count = i32::from_le_bytes([data[4], data[5], data[6], data[7]]);
                if usize::try_from(count) != Ok(linked_shapes.len()) {
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
            RecordType::SlideFlags10Atom.as_u16(),
            None,
            ATOM_VERSION,
            "SlideFlags10Atom",
        )?;
        let slide_time = cursor.take_optional(
            RecordType::SlideTime10Atom.as_u16(),
            None,
            ATOM_VERSION,
            "SlideTime10Atom",
        )?;
        let hash_code = cursor.take_optional(
            RecordType::HashCode10Atom.as_u16(),
            None,
            ATOM_VERSION,
            "HashCode10Atom",
        )?;
        let timing = cursor.take_optional(
            RecordType::ExtTimeNode.as_u16(),
            None,
            CONTAINER_VERSION,
            "ExtTimeNodeContainer",
        )?;
        let build_list = cursor.take_optional(
            RecordType::BuildList.as_u16(),
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
        Self::parse_records(Record::parse_sequence_strict(
            &payload,
            "PP10SlideBinaryTagExtension",
        )?)?;
        Ok(payload)
    }
}

impl DocBinaryTagExtension {
    /// Decode a versioned document-scope binary tag payload. Returns `Ok(None)`
    /// for unassigned (unknown) tags.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(version: ProgBinaryTagVersion, records: Vec<Record>) -> Result<Option<Self>> {
        match version {
            ProgBinaryTagVersion::PowerPoint9 => Ok(Some(Self::PowerPoint9(Box::new(
                DocBinaryTagExtension9::parse_records(records)?,
            )))),
            ProgBinaryTagVersion::PowerPoint10 => Ok(Some(Self::PowerPoint10(Box::new(
                DocBinaryTagExtension10::parse_records(records)?,
            )))),
            ProgBinaryTagVersion::PowerPoint11 => Ok(Some(Self::PowerPoint11(Box::new(
                DocBinaryTagExtension11::parse_records(records)?,
            )))),
            ProgBinaryTagVersion::PowerPoint12 => Ok(Some(Self::PowerPoint12(Box::new(
                DocBinaryTagExtension12::parse_records(records)?,
            )))),
            ProgBinaryTagVersion::Unknown => Ok(None),
        }
    }

    /// Serialize the tag payload byte-for-byte.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        match self {
            Self::PowerPoint9(extension) => extension.to_payload(),
            Self::PowerPoint10(extension) => extension.to_payload(),
            Self::PowerPoint11(extension) => extension.to_payload(),
            Self::PowerPoint12(extension) => extension.to_payload(),
        }
    }
}

impl SlideBinaryTagExtension {
    /// Decode a versioned slide-scope binary tag payload. Returns `Ok(None)`
    /// for unassigned (unknown) tags.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(version: ProgBinaryTagVersion, records: Vec<Record>) -> Result<Option<Self>> {
        match version {
            ProgBinaryTagVersion::PowerPoint9 => Ok(Some(Self::PowerPoint9(Box::new(
                SlideBinaryTagExtension9::parse_records(records)?,
            )))),
            ProgBinaryTagVersion::PowerPoint10 => Ok(Some(Self::PowerPoint10(Box::new(
                SlideBinaryTagExtension10::parse_records(records)?,
            )))),
            ProgBinaryTagVersion::PowerPoint12 => Ok(Some(Self::PowerPoint12(Box::new(
                SlideBinaryTagExtension12::parse_records(records)?,
            )))),
            ProgBinaryTagVersion::PowerPoint11 => {
                corrupted("___PPT11 is not an assigned slide-scope binary tag")
            },
            ProgBinaryTagVersion::Unknown => Ok(None),
        }
    }

    /// Serialize the tag payload byte-for-byte.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_payload(&self) -> Result<Vec<u8>> {
        match self {
            Self::PowerPoint9(extension) => extension.to_payload(),
            Self::PowerPoint10(extension) => extension.to_payload(),
            Self::PowerPoint12(extension) => extension.to_payload(),
        }
    }
}

impl ProgBinaryTag {
    /// Decode this tag's payload as a versioned document extension.
    ///
    /// Returns `Ok(None)` for unassigned (unknown) tags, whose payloads are
    /// preserved without interpretation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn doc_extension(&self) -> Result<Option<DocBinaryTagExtension>> {
        DocBinaryTagExtension::parse(self.version, self.records()?)
    }

    /// Decode this tag's payload as a versioned slide extension.
    ///
    /// Returns `Ok(None)` for unassigned (unknown) tags. `___PPT11` is not
    /// assigned at slide scope (MS-PPT 2.5.22), so decoding it as a slide
    /// extension is an error.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_extension(&self) -> Result<Option<SlideBinaryTagExtension>> {
        SlideBinaryTagExtension::parse(self.version, self.records()?)
    }
}

impl ProgTags {
    /// Decode every assigned versioned document extension in this container.
    ///
    /// Unknown tags are skipped; their payloads remain available through
    /// [`ProgTags::binary_tag`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn document_extensions(&self) -> Result<DocumentTagExtensions> {
        if self.scope != ProgTagScope::Document {
            return corrupted("slide-scope ProgTags cannot hold document extensions");
        }
        let mut extensions = DocumentTagExtensions::default();
        for tag in &self.tags {
            let ProgTag::Binary(binary_tag) = tag else {
                continue;
            };
            match binary_tag.doc_extension()? {
                Some(DocBinaryTagExtension::PowerPoint9(extension)) => {
                    extensions.powerpoint9 = Some(*extension);
                },
                Some(DocBinaryTagExtension::PowerPoint10(extension)) => {
                    extensions.powerpoint10 = Some(*extension);
                },
                Some(DocBinaryTagExtension::PowerPoint11(extension)) => {
                    extensions.powerpoint11 = Some(*extension);
                },
                Some(DocBinaryTagExtension::PowerPoint12(extension)) => {
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
    /// [`ProgTags::binary_tag`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_extensions(&self) -> Result<SlideTagExtensions> {
        if self.scope != ProgTagScope::Slide {
            return corrupted("document-scope ProgTags cannot hold slide extensions");
        }
        let mut extensions = SlideTagExtensions::default();
        for tag in &self.tags {
            let ProgTag::Binary(binary_tag) = tag else {
                continue;
            };
            match binary_tag.slide_extension()? {
                Some(SlideBinaryTagExtension::PowerPoint9(extension)) => {
                    extensions.powerpoint9 = Some(*extension);
                },
                Some(SlideBinaryTagExtension::PowerPoint10(extension)) => {
                    extensions.powerpoint10 = Some(*extension);
                },
                Some(SlideBinaryTagExtension::PowerPoint12(extension)) => {
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
    records: std::iter::Peekable<std::vec::IntoIter<Record>>,
    context: &'static str,
}

impl RecordCursor {
    fn new(records: Vec<Record>, context: &'static str) -> Self {
        Self {
            records: records.into_iter().peekable(),
            context,
        }
    }

    /// Consume records while they match the array element type, validating the
    /// version nibble of each element.
    fn take_array(&mut self, kind: u16, version: u16, label: &str) -> Result<Vec<Record>> {
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
    ) -> Result<Option<Record>> {
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
fn encode_record(record: &Record) -> Result<Vec<u8>> {
    if record.version > 0x0f || record.instance > 0x0fff {
        return corrupted("PPT record version or instance exceeds its bit field");
    }
    let declared = usize::try_from(record.data_length)
        .map_err(|_err| Error::Corrupted("PPT record length overflow".into()))?;
    if declared != record.data.len() {
        return corrupted("PPT record length does not match its payload");
    }
    let length = u32::try_from(record.data.len())
        .map_err(|_err| Error::Corrupted("PPT record payload exceeds u32".into()))?;
    let mut result = Vec::with_capacity(8usize.saturating_add(record.data.len()));
    result.extend_from_slice(&((record.instance << 4) | record.version).to_le_bytes());
    result.extend_from_slice(&record.record_type_raw.to_le_bytes());
    result.extend_from_slice(&length.to_le_bytes());
    result.extend_from_slice(&record.data);
    Ok(result)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
