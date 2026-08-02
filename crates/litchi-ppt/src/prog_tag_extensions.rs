//! Typed inner payloads of versioned document/slide binary programmable tags.
//!
//! This module implements the extension record grammars referenced by
//! [`super::prog_tags`]: `PP9DocBinaryTagExtension` through
//! `PP12DocBinaryTagExtension` (MS-PPT sections 2.4.23.5 through 2.4.23.8) and
//! `PP9SlideBinaryTagExtension`, `PP10SlideBinaryTagExtension`, and
//! `PP12SlideBinaryTagExtension` (sections 2.5.23, 2.5.24, and 2.5.34).
//!
//! Every grammar slot retains its raw [`PptRecord`]: parsing is strictly
//! ordered per the spec but completely inert, and serialization is byte-exact.
//! Deeper field decoding deliberately stays with the dedicated piecemeal
//! readers (`kinsoku.rs`, `broadcast.rs`, `html_publish.rs`,
//! `presentation_advisor.rs`, `envelope_data.rs`, and friends), which consume
//! the same records through `PptRecord::versioned_binary_tag_records`; this
//! module only assigns each record to its grammar slot.

use crate::consts::PptRecordType;

use super::package::{PptError, Result};
use super::prog_tags::{
    PowerPointProgBinaryTag, PowerPointProgBinaryTagVersion, PowerPointProgTag,
    PowerPointProgTagScope, PowerPointProgTags,
};
use super::records::PptRecord;

/// `RT_PresentationAdvisorFlags9Atom` (MS-PPT 2.13.24).
const PRES_ADVISOR_FLAGS_9_ATOM: u16 = 0x177a;
/// `RT_HtmlDocInfo9Atom` (MS-PPT 2.13.24).
const HTML_DOC_INFO_9_ATOM: u16 = 0x177b;
/// `RT_HtmlPublishInfo9` (MS-PPT 2.13.24).
const HTML_PUBLISH_INFO_9: u16 = 0x177d;
/// `RT_BroadcastDocInfo9` (MS-PPT 2.13.24).
const BROADCAST_DOC_INFO_9: u16 = 0x177e;
/// `RT_EnvelopeFlags9Atom` (MS-PPT 2.13.24).
const ENVELOPE_FLAGS_9_ATOM: u16 = 0x1784;
/// `RT_EnvelopeData9Atom` (MS-PPT 2.13.24).
const ENVELOPE_DATA_9_ATOM: u16 = 0x1785;
/// `RT_Comment10` (MS-PPT 2.13.24).
const COMMENT_10: u16 = 0x2ee0;

/// `CopyrightAtom` record instance (MS-PPT 2.4.22.1).
const COPYRIGHT_INSTANCE: u16 = 0x001;
/// `KeywordsAtom` record instance (MS-PPT 2.4.22.2).
const KEYWORDS_INSTANCE: u16 = 0x002;
/// `ModifyPasswordAtom` record instance (MS-PPT 2.4.7).
const MODIFY_PASSWORD_INSTANCE: u16 = 0x003;

/// Container record version nibble.
const CONTAINER_VERSION: u16 = 0x0f;
/// Atom record version nibble.
const ATOM_VERSION: u16 = 0x00;

/// A `PP9DocBinaryTagExtension` payload (MS-PPT 2.4.23.5).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPoint9DocBinaryTagExtension {
    /// `rgTextMasterStyle9`: `TextMasterStyle9Atom` records.
    pub text_master_styles: Vec<PptRecord>,
    /// `blipCollectionContainer`: optional `BlipCollection9Container`.
    pub blip_collection: Option<PptRecord>,
    /// `textDefaultsAtom`: optional `TextDefaults9Atom`.
    pub text_defaults: Option<PptRecord>,
    /// `kinsokuContainer`: optional `Kinsoku9Container`.
    pub kinsoku: Option<PptRecord>,
    /// `rgExternalHyperlink9`: `ExHyperlink9Container` records.
    pub external_hyperlinks: Vec<PptRecord>,
    /// `presAdvisorFlagsAtom`: optional `PresAdvisorFlags9Atom`.
    pub advisor_flags: Option<PptRecord>,
    /// `envelopeDataAtom`: optional `EnvelopeData9Atom`.
    pub envelope_data: Option<PptRecord>,
    /// `envelopeFlagsAtom`: optional `EnvelopeFlags9Atom`.
    pub envelope_flags: Option<PptRecord>,
    /// `htmlDocInfoAtom`: optional `HTMLDocInfo9Atom`.
    pub html_doc_info: Option<PptRecord>,
    /// `htmlPublishInfoAtom`: optional `HTMLPublishInfo9Container`.
    pub html_publish_info: Option<PptRecord>,
    /// `rgBroadcastDocInfo9`: `BroadcastDocInfo9Container` records.
    pub broadcasts: Vec<PptRecord>,
    /// `outlineTextPropsContainer`: optional `OutlineTextProps9Container`.
    pub outline_text_props: Option<PptRecord>,
}

/// A `PP10DocBinaryTagExtension` payload (MS-PPT 2.4.23.6).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPoint10DocBinaryTagExtension {
    /// `fontCollectionContainer`: optional `FontCollection10Container`.
    pub font_collection: Option<PptRecord>,
    /// `rgTextMasterStyle10`: `TextMasterStyle10Atom` records.
    pub text_master_styles: Vec<PptRecord>,
    /// `textDefaultsAtom`: optional `TextDefaults10Atom`.
    pub text_defaults: Option<PptRecord>,
    /// `gridSpacingAtom`: the `GridSpacing10Atom`. It is grammatically required,
    /// so [`Self::parse_records`] always yields `Some`; the field is an `Option`
    /// only so the struct stays constructible, and [`Self::to_payload`]
    /// rejects a missing value.
    pub grid_spacing: Option<PptRecord>,
    /// `rgCommentIndex10`: `CommentIndex10Container` records.
    pub comment_indices: Vec<PptRecord>,
    /// `fontEmbedFlagsAtom`: optional `FontEmbedFlags10Atom`.
    pub font_embed_flags: Option<PptRecord>,
    /// `copyrightAtom`: optional `CopyrightAtom` (CString instance 0x001).
    pub copyright: Option<PptRecord>,
    /// `keywordsAtom`: optional `KeywordsAtom` (CString instance 0x002).
    pub keywords: Option<PptRecord>,
    /// `filterPrivacyFlagsAtom`: optional `FilterPrivacyFlags10Atom`.
    pub filter_privacy_flags: Option<PptRecord>,
    /// `outlineTextPropsContainer`: optional `OutlineTextProps10Container`.
    pub outline_text_props: Option<PptRecord>,
    /// `docToolbarStatesAtom`: optional `DocToolbarStates10Atom`.
    pub toolbar_states: Option<PptRecord>,
    /// `slideListTableContainer`: optional `SlideListTable10Container`.
    pub slide_list_table: Option<PptRecord>,
    /// `rgDiffTree10Container`: `DiffTree10Container` records.
    pub diff_trees: Vec<PptRecord>,
    /// `modifyPasswordAtom`: optional `ModifyPasswordAtom` (CString instance 0x003).
    pub modify_password: Option<PptRecord>,
    /// `photoAlbumInfoAtom`: optional `PhotoAlbumInfo10Atom`.
    pub photo_album_info: Option<PptRecord>,
}

/// A `PP11DocBinaryTagExtension` payload (MS-PPT 2.4.23.7).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPoint11DocBinaryTagExtension {
    /// `smartTagStore11`: optional `SmartTagStore11Container`.
    pub smart_tag_store: Option<PptRecord>,
    /// `outlineTextProps`: optional `OutlineTextProps11Container`.
    pub outline_text_props: Option<PptRecord>,
}

/// A `PP12DocBinaryTagExtension` payload (MS-PPT 2.4.23.8).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPoint12DocBinaryTagExtension {
    /// `rtDocFlagsAtom`: optional `RoundTripDocFlags12Atom`.
    pub doc_flags: Option<PptRecord>,
}

/// A `PP9SlideBinaryTagExtension` payload (MS-PPT 2.5.23).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPoint9SlideBinaryTagExtension {
    /// `rgTextMasterStyleAtom`: `TextMasterStyle9Atom` records. The spec bounds
    /// the array by `rhData.recLen`, so every record in the payload must be a
    /// `TextMasterStyle9Atom`.
    pub text_master_styles: Vec<PptRecord>,
}

/// A `PP10SlideBinaryTagExtension` payload (MS-PPT 2.5.24).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPoint10SlideBinaryTagExtension {
    /// `rgTextMasterStyleAtom`: `TextMasterStyle10Atom` records.
    pub text_master_styles: Vec<PptRecord>,
    /// `rgComment10Container`: `Comment10Container` records.
    pub comments: Vec<PptRecord>,
    /// `linkedSlideAtom`: optional `LinkedSlide10Atom`.
    pub linked_slide: Option<PptRecord>,
    /// `rgLinkedShape10Atom`: `LinkedShape10Atom` records. The count MUST match
    /// `linkedSlideAtom.cLinkedShapes` when the atom is present, and the array
    /// cannot appear without it.
    pub linked_shapes: Vec<PptRecord>,
    /// `slideFlagsAtom`: optional `SlideFlags10Atom`.
    pub slide_flags: Option<PptRecord>,
    /// `slideTimeAtom`: optional `SlideTime10Atom`.
    pub slide_time: Option<PptRecord>,
    /// `hashCodeAtom`: optional `HashCode10Atom`.
    pub hash_code: Option<PptRecord>,
    /// `extTimeNodeContainer`: optional `ExtTimeNodeContainer`.
    pub timing: Option<PptRecord>,
    /// `buildListContainer`: optional `BuildListContainer`.
    pub build_list: Option<PptRecord>,
}

/// A `PP12SlideBinaryTagExtension` payload (MS-PPT 2.5.34).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPoint12SlideBinaryTagExtension {
    /// `roundTripHeaderFooterDefaultsAtom`: optional
    /// `RoundTripHeaderFooterDefaults12Atom`.
    pub header_footer_defaults: Option<PptRecord>,
}

/// Any versioned `DocProgBinaryTagSubContainerOrAtom` payload (MS-PPT 2.4.23.4).
///
/// Variants are boxed to keep the enum compact regardless of grammar size.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PowerPointDocBinaryTagExtension {
    /// `PP9DocBinaryTagExtension` (section 2.4.23.5).
    PowerPoint9(Box<PowerPoint9DocBinaryTagExtension>),
    /// `PP10DocBinaryTagExtension` (section 2.4.23.6).
    PowerPoint10(Box<PowerPoint10DocBinaryTagExtension>),
    /// `PP11DocBinaryTagExtension` (section 2.4.23.7).
    PowerPoint11(Box<PowerPoint11DocBinaryTagExtension>),
    /// `PP12DocBinaryTagExtension` (section 2.4.23.8).
    PowerPoint12(Box<PowerPoint12DocBinaryTagExtension>),
}

/// Any versioned `SlideProgBinaryTagSubContainerOrAtom` payload (MS-PPT 2.5.22).
///
/// Variants are boxed to keep the enum compact regardless of grammar size.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PowerPointSlideBinaryTagExtension {
    /// `PP9SlideBinaryTagExtension` (section 2.5.23).
    PowerPoint9(Box<PowerPoint9SlideBinaryTagExtension>),
    /// `PP10SlideBinaryTagExtension` (section 2.5.24).
    PowerPoint10(Box<PowerPoint10SlideBinaryTagExtension>),
    /// `PP12SlideBinaryTagExtension` (section 2.5.34).
    PowerPoint12(Box<PowerPoint12SlideBinaryTagExtension>),
}

/// Decoded versioned extensions of one document-level `ProgTags` container.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointDocumentTagExtensions {
    /// Decoded `___PPT9` tag, when present.
    pub powerpoint9: Option<PowerPoint9DocBinaryTagExtension>,
    /// Decoded `___PPT10` tag, when present.
    pub powerpoint10: Option<PowerPoint10DocBinaryTagExtension>,
    /// Decoded `___PPT11` tag, when present.
    pub powerpoint11: Option<PowerPoint11DocBinaryTagExtension>,
    /// Decoded `___PPT12` tag, when present.
    pub powerpoint12: Option<PowerPoint12DocBinaryTagExtension>,
}

/// Decoded versioned extensions of one slide-level `ProgTags` container.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PowerPointSlideTagExtensions {
    /// Decoded `___PPT9` tag, when present.
    pub powerpoint9: Option<PowerPoint9SlideBinaryTagExtension>,
    /// Decoded `___PPT10` tag, when present.
    pub powerpoint10: Option<PowerPoint10SlideBinaryTagExtension>,
    /// Decoded `___PPT12` tag, when present.
    pub powerpoint12: Option<PowerPoint12SlideBinaryTagExtension>,
}

/// Implement `parse_records`/`to_payload` for an extension grammar.
///
/// Grammars are an ordered sequence of greedy arrays (consumed while the
/// record type matches), optional single records, and required single
/// records, exactly as the spec tables list them. Slots are declared in
/// grammar order: `array(label, field, type, version)` for a record-type
/// array, `opt(label, field, type, instance, version)` for an optional
/// record, and `req(label, field, type, version)` for a required record.
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::prog_tags::{PowerPointProgTagLimits, PowerPointProgTags};

    fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + payload.len());
        bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    fn parse_payload(bytes: &[u8]) -> Vec<PptRecord> {
        PptRecord::parse_sequence_strict(bytes, "test payload").unwrap()
    }

    fn atom(instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
        record_bytes(ATOM_VERSION, instance, kind, payload)
    }

    fn container(kind: u16, payload: &[u8]) -> Vec<u8> {
        record_bytes(CONTAINER_VERSION, 0, kind, payload)
    }

    fn cstring(instance: u16, text: &str) -> Vec<u8> {
        let data: Vec<u8> = text.encode_utf16().flat_map(u16::to_le_bytes).collect();
        atom(instance, PptRecordType::CString.as_u16(), &data)
    }

    /// Encode the full PP9 document grammar, in spec order.
    fn pp9_doc_payload() -> Vec<u8> {
        [
            atom(1, PptRecordType::TextMasterStyle9Atom.as_u16(), &[0; 20]),
            atom(2, PptRecordType::TextMasterStyle9Atom.as_u16(), &[0; 20]),
            container(PptRecordType::BlipCollection9.as_u16(), &[]),
            atom(0, PptRecordType::TextDefaults9Atom.as_u16(), &[0; 8]),
            container(PptRecordType::Kinsoku.as_u16(), &[]),
            container(PptRecordType::ExternalHyperlink9.as_u16(), &[]),
            atom(0, PRES_ADVISOR_FLAGS_9_ATOM, &[0; 4]),
            atom(0, ENVELOPE_DATA_9_ATOM, &[1, 2, 3]),
            atom(0, ENVELOPE_FLAGS_9_ATOM, &[0; 4]),
            atom(0, HTML_DOC_INFO_9_ATOM, &[0; 16]),
            container(HTML_PUBLISH_INFO_9, &[]),
            container(BROADCAST_DOC_INFO_9, &[]),
            container(BROADCAST_DOC_INFO_9, &[]),
            container(PptRecordType::OutlineTextProps9.as_u16(), &[]),
        ]
        .concat()
    }

    /// Encode the full PP10 document grammar, in spec order.
    fn pp10_doc_payload() -> Vec<u8> {
        [
            container(PptRecordType::FontCollection10.as_u16(), &[]),
            atom(0, PptRecordType::TextMasterStyle10Atom.as_u16(), &[0; 12]),
            atom(0, PptRecordType::TextDefaults10Atom.as_u16(), &[0; 8]),
            atom(0, PptRecordType::GridSpacing10Atom.as_u16(), &[0; 8]),
            container(PptRecordType::CommentIndex10.as_u16(), &[]),
            atom(0, PptRecordType::FontEmbedFlags10Atom.as_u16(), &[0; 4]),
            cstring(COPYRIGHT_INSTANCE, "(c) Ada"),
            cstring(KEYWORDS_INSTANCE, "ppt,test"),
            atom(0, PptRecordType::FilterPrivacyFlags10Atom.as_u16(), &[0; 4]),
            container(PptRecordType::OutlineTextProps10.as_u16(), &[]),
            atom(0, PptRecordType::DocToolbarStates10Atom.as_u16(), &[0]),
            container(PptRecordType::SlideListTable10.as_u16(), &[]),
            container(PptRecordType::DiffTree10.as_u16(), &[]),
            cstring(MODIFY_PASSWORD_INSTANCE, "secret"),
            atom(0, PptRecordType::PhotoAlbumInfo10Atom.as_u16(), &[0; 6]),
        ]
        .concat()
    }

    /// Encode the full PP10 slide grammar, in spec order.
    fn pp10_slide_payload() -> Vec<u8> {
        let mut linked_slide = Vec::new();
        linked_slide.extend_from_slice(&42u32.to_le_bytes());
        linked_slide.extend_from_slice(&2i32.to_le_bytes());
        [
            atom(0, PptRecordType::TextMasterStyle10Atom.as_u16(), &[0; 12]),
            container(COMMENT_10, &[]),
            atom(0, PptRecordType::LinkedSlide10Atom.as_u16(), &linked_slide),
            atom(0, PptRecordType::LinkedShape10Atom.as_u16(), &[0; 8]),
            atom(0, PptRecordType::LinkedShape10Atom.as_u16(), &[0; 8]),
            atom(0, PptRecordType::SlideFlags10Atom.as_u16(), &[0; 4]),
            atom(0, PptRecordType::SlideTime10Atom.as_u16(), &[0; 8]),
            atom(0, PptRecordType::HashCode10Atom.as_u16(), &[0; 4]),
            container(PptRecordType::ExtTimeNode.as_u16(), &[]),
            container(PptRecordType::BuildList.as_u16(), &[]),
        ]
        .concat()
    }

    /// Wrap an extension payload in a `ProgBinaryTag`/`ProgTags` record pair.
    fn prog_tags_record(tag_name: &str, extension_payload: &[u8]) -> (Vec<u8>, PptRecord) {
        let name: Vec<u8> = tag_name.encode_utf16().flat_map(u16::to_le_bytes).collect();
        let binary_tag = record_bytes(
            CONTAINER_VERSION,
            0,
            PptRecordType::ProgBinaryTag.as_u16(),
            &[
                atom(0, PptRecordType::CString.as_u16(), &name),
                atom(0, PptRecordType::BinaryTagData.as_u16(), extension_payload),
            ]
            .concat(),
        );
        let bytes = record_bytes(
            CONTAINER_VERSION,
            0,
            PptRecordType::ProgTags.as_u16(),
            &binary_tag,
        );
        let (record, consumed) = PptRecord::parse_strict(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        (bytes, record)
    }

    #[test]
    fn pp9_doc_extension_assigns_every_slot_and_round_trips_exactly() {
        let payload = pp9_doc_payload();
        let extension =
            PowerPoint9DocBinaryTagExtension::parse_records(parse_payload(&payload)).unwrap();

        assert_eq!(extension.text_master_styles.len(), 2);
        assert!(extension.blip_collection.is_some());
        assert!(extension.text_defaults.is_some());
        assert!(extension.kinsoku.is_some());
        assert_eq!(extension.external_hyperlinks.len(), 1);
        assert!(extension.advisor_flags.is_some());
        assert!(extension.envelope_data.is_some());
        assert!(extension.envelope_flags.is_some());
        assert!(extension.html_doc_info.is_some());
        assert!(extension.html_publish_info.is_some());
        assert_eq!(extension.broadcasts.len(), 2);
        assert!(extension.outline_text_props.is_some());
        assert_eq!(extension.to_payload().unwrap(), payload);
    }

    #[test]
    fn pp10_doc_extension_assigns_every_slot_and_round_trips_exactly() {
        let payload = pp10_doc_payload();
        let extension =
            PowerPoint10DocBinaryTagExtension::parse_records(parse_payload(&payload)).unwrap();

        assert!(extension.font_collection.is_some());
        assert_eq!(extension.text_master_styles.len(), 1);
        assert!(extension.text_defaults.is_some());
        assert!(extension.grid_spacing.is_some());
        assert_eq!(extension.comment_indices.len(), 1);
        assert!(extension.font_embed_flags.is_some());
        assert!(extension.copyright.is_some());
        assert!(extension.keywords.is_some());
        assert!(extension.filter_privacy_flags.is_some());
        assert!(extension.outline_text_props.is_some());
        assert!(extension.toolbar_states.is_some());
        assert!(extension.slide_list_table.is_some());
        assert_eq!(extension.diff_trees.len(), 1);
        assert!(extension.modify_password.is_some());
        assert!(extension.photo_album_info.is_some());
        assert_eq!(extension.to_payload().unwrap(), payload);
    }

    #[test]
    fn pp10_doc_extension_allows_minimal_grammar() {
        // Only the required GridSpacing10Atom.
        let payload = atom(0, PptRecordType::GridSpacing10Atom.as_u16(), &[0; 8]);
        let extension =
            PowerPoint10DocBinaryTagExtension::parse_records(parse_payload(&payload)).unwrap();
        assert!(extension.grid_spacing.is_some());
        assert!(extension.font_collection.is_none());
        assert!(extension.modify_password.is_none());
        assert_eq!(extension.to_payload().unwrap(), payload);
    }

    #[test]
    fn pp11_and_pp12_doc_extensions_round_trip() {
        let pp11_payload = [
            container(PptRecordType::SmartTagStore11.as_u16(), &[]),
            container(PptRecordType::OutlineTextProps11.as_u16(), &[]),
        ]
        .concat();
        let pp11 =
            PowerPoint11DocBinaryTagExtension::parse_records(parse_payload(&pp11_payload)).unwrap();
        assert!(pp11.smart_tag_store.is_some());
        assert!(pp11.outline_text_props.is_some());
        assert_eq!(pp11.to_payload().unwrap(), pp11_payload);

        let pp12_payload = atom(0, PptRecordType::RoundTripDocFlags12Atom.as_u16(), &[0]);
        let pp12 =
            PowerPoint12DocBinaryTagExtension::parse_records(parse_payload(&pp12_payload)).unwrap();
        assert!(pp12.doc_flags.is_some());
        assert_eq!(pp12.to_payload().unwrap(), pp12_payload);

        // PP12 with all-optional grammar accepts an empty payload.
        let empty = PowerPoint12DocBinaryTagExtension::parse_records(Vec::new()).unwrap();
        assert!(empty.doc_flags.is_none());
        assert_eq!(empty.to_payload().unwrap(), Vec::<u8>::new());
    }

    #[test]
    fn pp9_and_pp12_slide_extensions_round_trip() {
        let pp9_payload = [
            atom(0, PptRecordType::TextMasterStyle9Atom.as_u16(), &[0; 20]),
            atom(3, PptRecordType::TextMasterStyle9Atom.as_u16(), &[0; 20]),
        ]
        .concat();
        let pp9 =
            PowerPoint9SlideBinaryTagExtension::parse_records(parse_payload(&pp9_payload)).unwrap();
        assert_eq!(pp9.text_master_styles.len(), 2);
        assert_eq!(pp9.to_payload().unwrap(), pp9_payload);

        let pp12_payload = atom(
            0,
            PptRecordType::RoundTripHeaderFooterDefaults12Atom.as_u16(),
            &[0],
        );
        let pp12 = PowerPoint12SlideBinaryTagExtension::parse_records(parse_payload(&pp12_payload))
            .unwrap();
        assert!(pp12.header_footer_defaults.is_some());
        assert_eq!(pp12.to_payload().unwrap(), pp12_payload);
    }

    #[test]
    fn pp10_slide_extension_assigns_slots_and_round_trips_exactly() {
        let payload = pp10_slide_payload();
        let extension =
            PowerPoint10SlideBinaryTagExtension::parse_records(parse_payload(&payload)).unwrap();

        assert_eq!(extension.text_master_styles.len(), 1);
        assert_eq!(extension.comments.len(), 1);
        assert!(extension.linked_slide.is_some());
        assert_eq!(extension.linked_shapes.len(), 2);
        assert!(extension.slide_flags.is_some());
        assert!(extension.slide_time.is_some());
        assert!(extension.hash_code.is_some());
        assert!(extension.timing.is_some());
        assert!(extension.build_list.is_some());
        assert_eq!(extension.to_payload().unwrap(), payload);
    }

    #[test]
    fn pp10_slide_extension_validates_linked_shape_count() {
        // Count says 1, array holds 2.
        let mut linked_slide = Vec::new();
        linked_slide.extend_from_slice(&42u32.to_le_bytes());
        linked_slide.extend_from_slice(&1i32.to_le_bytes());
        let mismatched = [
            atom(0, PptRecordType::LinkedSlide10Atom.as_u16(), &linked_slide),
            atom(0, PptRecordType::LinkedShape10Atom.as_u16(), &[0; 8]),
            atom(0, PptRecordType::LinkedShape10Atom.as_u16(), &[0; 8]),
        ]
        .concat();
        assert!(
            PowerPoint10SlideBinaryTagExtension::parse_records(parse_payload(&mismatched)).is_err()
        );

        // Shapes without the linked-slide atom.
        let orphan = atom(0, PptRecordType::LinkedShape10Atom.as_u16(), &[0; 8]);
        assert!(
            PowerPoint10SlideBinaryTagExtension::parse_records(parse_payload(&orphan)).is_err()
        );

        // Truncated linked-slide atom.
        let truncated = atom(0, PptRecordType::LinkedSlide10Atom.as_u16(), &[0; 4]);
        assert!(
            PowerPoint10SlideBinaryTagExtension::parse_records(parse_payload(&truncated)).is_err()
        );
    }

    #[test]
    fn grammars_reject_out_of_order_missing_required_and_trailing_records() {
        // PP9 doc: OutlineTextProps9 before the broadcast array is out of order.
        let out_of_order = [
            container(PptRecordType::OutlineTextProps9.as_u16(), &[]),
            container(BROADCAST_DOC_INFO_9, &[]),
        ]
        .concat();
        assert!(
            PowerPoint9DocBinaryTagExtension::parse_records(parse_payload(&out_of_order)).is_err()
        );

        // PP10 doc: the required GridSpacing10Atom is missing.
        let missing_required = atom(0, PptRecordType::TextMasterStyle10Atom.as_u16(), &[0; 12]);
        assert!(
            PowerPoint10DocBinaryTagExtension::parse_records(parse_payload(&missing_required))
                .is_err()
        );

        // PP10 doc: a ModifyPasswordAtom where the CopyrightAtom belongs.
        let wrong_instance = [
            atom(0, PptRecordType::GridSpacing10Atom.as_u16(), &[0; 8]),
            cstring(MODIFY_PASSWORD_INSTANCE, "secret"),
        ]
        .concat();
        assert!(
            PowerPoint10DocBinaryTagExtension::parse_records(parse_payload(&wrong_instance))
                .is_err()
        );

        // PP9 slide: any non-TextMasterStyle9Atom record is outside the grammar.
        let foreign = atom(0, PptRecordType::TextDefaults9Atom.as_u16(), &[0; 8]);
        assert!(
            PowerPoint9SlideBinaryTagExtension::parse_records(parse_payload(&foreign)).is_err()
        );

        // Wrong version nibble on an array element.
        let bad_version = record_bytes(
            CONTAINER_VERSION,
            0,
            PptRecordType::TextMasterStyle9Atom.as_u16(),
            &[0; 20],
        );
        assert!(
            PowerPoint9SlideBinaryTagExtension::parse_records(parse_payload(&bad_version)).is_err()
        );
    }

    #[test]
    fn tag_and_container_level_dispatch_decode_extensions() {
        let limits = PowerPointProgTagLimits::default();
        let (bytes, record) = prog_tags_record("___PPT9", &pp9_doc_payload());
        let tags =
            PowerPointProgTags::parse(&record, PowerPointProgTagScope::Document, limits).unwrap();

        let extensions = tags.document_extensions().unwrap();
        let pp9 = extensions.powerpoint9.as_ref().unwrap();
        assert_eq!(pp9.text_master_styles.len(), 2);
        assert_eq!(pp9.broadcasts.len(), 2);
        assert!(extensions.powerpoint10.is_none());

        let tag = tags
            .binary_tag(PowerPointProgBinaryTagVersion::PowerPoint9)
            .unwrap();
        let extension = tag.doc_extension().unwrap().unwrap();
        assert_eq!(extension.to_payload().unwrap(), pp9_doc_payload());
        // Container-level bytes are unaffected by extension decoding.
        assert_eq!(tags.to_bytes(limits).unwrap(), bytes);

        // Document tags cannot decode slide extensions and vice versa.
        assert!(tags.slide_extensions().is_err());
        // The doc-scoped ___PPT9 payload is not a valid PP9 slide grammar.
        assert!(tag.slide_extension().is_err());
    }

    #[test]
    fn slide_scope_dispatch_decodes_pp10_slide_extension() {
        let limits = PowerPointProgTagLimits::default();
        let (_, record) = prog_tags_record("___PPT10", &pp10_slide_payload());
        let tags =
            PowerPointProgTags::parse(&record, PowerPointProgTagScope::Slide, limits).unwrap();

        let extensions = tags.slide_extensions().unwrap();
        let pp10 = extensions.powerpoint10.as_ref().unwrap();
        assert_eq!(pp10.linked_shapes.len(), 2);
        assert!(pp10.build_list.is_some());
        assert!(extensions.powerpoint9.is_none());
        assert!(tags.document_extensions().is_err());

        let tag = tags
            .binary_tag(PowerPointProgBinaryTagVersion::PowerPoint10)
            .unwrap();
        assert_eq!(
            tag.slide_extension()
                .unwrap()
                .unwrap()
                .to_payload()
                .unwrap(),
            pp10_slide_payload()
        );
    }

    #[test]
    fn ppt11_versioned_tag_rejects_slide_extension_decode() {
        let limits = PowerPointProgTagLimits::default();
        let payload = container(PptRecordType::SmartTagStore11.as_u16(), &[]);
        let (_, record) = prog_tags_record("___PPT11", &payload);
        let tags =
            PowerPointProgTags::parse(&record, PowerPointProgTagScope::Document, limits).unwrap();
        let tag = tags
            .binary_tag(PowerPointProgBinaryTagVersion::PowerPoint11)
            .unwrap();
        assert!(tag.slide_extension().is_err());
        assert!(tag.doc_extension().unwrap().is_some());
    }
}
