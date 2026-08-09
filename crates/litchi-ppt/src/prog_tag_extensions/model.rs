use crate::records::Record;

/// A `PP9DocBinaryTagExtension` payload (MS-PPT 2.4.23.5).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocBinaryTagExtension9 {
    /// `rgTextMasterStyle9`: `TextMasterStyle9Atom` records.
    pub text_master_styles: Vec<Record>,
    /// `blipCollectionContainer`: optional `BlipCollection9Container`.
    pub blip_collection: Option<Record>,
    /// `textDefaultsAtom`: optional `TextDefaults9Atom`.
    pub text_defaults: Option<Record>,
    /// `kinsokuContainer`: optional `Kinsoku9Container`.
    pub kinsoku: Option<Record>,
    /// `rgExternalHyperlink9`: `ExHyperlink9Container` records.
    pub external_hyperlinks: Vec<Record>,
    /// `presAdvisorFlagsAtom`: optional `PresAdvisorFlags9Atom`.
    pub advisor_flags: Option<Record>,
    /// `envelopeDataAtom`: optional `EnvelopeData9Atom`.
    pub envelope_data: Option<Record>,
    /// `envelopeFlagsAtom`: optional `EnvelopeFlags9Atom`.
    pub envelope_flags: Option<Record>,
    /// `htmlDocInfoAtom`: optional `HTMLDocInfo9Atom`.
    pub html_doc_info: Option<Record>,
    /// `htmlPublishInfoAtom`: optional `HTMLPublishInfo9Container`.
    pub html_publish_info: Option<Record>,
    /// `rgBroadcastDocInfo9`: `BroadcastDocInfo9Container` records.
    pub broadcasts: Vec<Record>,
    /// `outlineTextPropsContainer`: optional `OutlineTextProps9Container`.
    pub outline_text_props: Option<Record>,
}

/// A `PP10DocBinaryTagExtension` payload (MS-PPT 2.4.23.6).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocBinaryTagExtension10 {
    /// `fontCollectionContainer`: optional `FontCollection10Container`.
    pub font_collection: Option<Record>,
    /// `rgTextMasterStyle10`: `TextMasterStyle10Atom` records.
    pub text_master_styles: Vec<Record>,
    /// `textDefaultsAtom`: optional `TextDefaults10Atom`.
    pub text_defaults: Option<Record>,
    /// `gridSpacingAtom`: optional `GridSpacing10Atom`. Real producers omit it
    /// when grid preferences were never materialized.
    pub grid_spacing: Option<Record>,
    /// `rgCommentIndex10`: `CommentIndex10Container` records.
    pub comment_indices: Vec<Record>,
    /// `fontEmbedFlagsAtom`: optional `FontEmbedFlags10Atom`.
    pub font_embed_flags: Option<Record>,
    /// `copyrightAtom`: optional `CopyrightAtom` (`CString` instance 0x001).
    pub copyright: Option<Record>,
    /// `keywordsAtom`: optional `KeywordsAtom` (`CString` instance 0x002).
    pub keywords: Option<Record>,
    /// `filterPrivacyFlagsAtom`: optional `FilterPrivacyFlags10Atom`.
    pub filter_privacy_flags: Option<Record>,
    /// `outlineTextPropsContainer`: optional `OutlineTextProps10Container`.
    pub outline_text_props: Option<Record>,
    /// `docToolbarStatesAtom`: optional `DocToolbarStates10Atom`.
    pub toolbar_states: Option<Record>,
    /// `slideListTableContainer`: optional `SlideListTable10Container`.
    pub slide_list_table: Option<Record>,
    /// `rgDiffTree10Container`: `DiffTree10Container` records.
    pub diff_trees: Vec<Record>,
    /// `modifyPasswordAtom`: optional `ModifyPasswordAtom` (`CString` instance 0x003).
    pub modify_password: Option<Record>,
    /// `photoAlbumInfoAtom`: optional `PhotoAlbumInfo10Atom`.
    pub photo_album_info: Option<Record>,
}

/// A `PP11DocBinaryTagExtension` payload (MS-PPT 2.4.23.7).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocBinaryTagExtension11 {
    /// `smartTagStore11`: optional `SmartTagStore11Container`.
    pub smart_tag_store: Option<Record>,
    /// `outlineTextProps`: optional `OutlineTextProps11Container`.
    pub outline_text_props: Option<Record>,
}

/// A `PP12DocBinaryTagExtension` payload (MS-PPT 2.4.23.8).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocBinaryTagExtension12 {
    /// `rtDocFlagsAtom`: optional `RoundTripDocFlags12Atom`.
    pub doc_flags: Option<Record>,
}

/// A `PP9SlideBinaryTagExtension` payload (MS-PPT 2.5.23).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SlideBinaryTagExtension9 {
    /// `rgTextMasterStyleAtom`: `TextMasterStyle9Atom` records. The spec bounds
    /// the array by `rhData.recLen`, so every record in the payload must be a
    /// `TextMasterStyle9Atom`.
    pub text_master_styles: Vec<Record>,
}

/// A `PP10SlideBinaryTagExtension` payload (MS-PPT 2.5.24).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SlideBinaryTagExtension10 {
    /// `rgTextMasterStyleAtom`: `TextMasterStyle10Atom` records.
    pub text_master_styles: Vec<Record>,
    /// `rgComment10Container`: `Comment10Container` records.
    pub comments: Vec<Record>,
    /// `linkedSlideAtom`: optional `LinkedSlide10Atom`.
    pub linked_slide: Option<Record>,
    /// `rgLinkedShape10Atom`: `LinkedShape10Atom` records. The count MUST match
    /// `linkedSlideAtom.cLinkedShapes` when the atom is present, and the array
    /// cannot appear without it.
    pub linked_shapes: Vec<Record>,
    /// `slideFlagsAtom`: optional `SlideFlags10Atom`.
    pub slide_flags: Option<Record>,
    /// `slideTimeAtom`: optional `SlideTime10Atom`.
    pub slide_time: Option<Record>,
    /// `hashCodeAtom`: optional `HashCode10Atom`.
    pub hash_code: Option<Record>,
    /// `extTimeNodeContainer`: optional `ExtTimeNodeContainer`.
    pub timing: Option<Record>,
    /// `buildListContainer`: optional `BuildListContainer`.
    pub build_list: Option<Record>,
}

/// A `PP12SlideBinaryTagExtension` payload (MS-PPT 2.5.34).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SlideBinaryTagExtension12 {
    /// `roundTripHeaderFooterDefaultsAtom`: optional
    /// `RoundTripHeaderFooterDefaults12Atom`.
    pub header_footer_defaults: Option<Record>,
}

/// Any versioned `DocProgBinaryTagSubContainerOrAtom` payload (MS-PPT 2.4.23.4).
///
/// Variants are boxed to keep the enum compact regardless of grammar size.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DocBinaryTagExtension {
    /// `PP9DocBinaryTagExtension` (section 2.4.23.5).
    PowerPoint9(Box<DocBinaryTagExtension9>),
    /// `PP10DocBinaryTagExtension` (section 2.4.23.6).
    PowerPoint10(Box<DocBinaryTagExtension10>),
    /// `PP11DocBinaryTagExtension` (section 2.4.23.7).
    PowerPoint11(Box<DocBinaryTagExtension11>),
    /// `PP12DocBinaryTagExtension` (section 2.4.23.8).
    PowerPoint12(Box<DocBinaryTagExtension12>),
}

/// Any versioned `SlideProgBinaryTagSubContainerOrAtom` payload (MS-PPT 2.5.22).
///
/// Variants are boxed to keep the enum compact regardless of grammar size.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SlideBinaryTagExtension {
    /// `PP9SlideBinaryTagExtension` (section 2.5.23).
    PowerPoint9(Box<SlideBinaryTagExtension9>),
    /// `PP10SlideBinaryTagExtension` (section 2.5.24).
    PowerPoint10(Box<SlideBinaryTagExtension10>),
    /// `PP12SlideBinaryTagExtension` (section 2.5.34).
    PowerPoint12(Box<SlideBinaryTagExtension12>),
}

/// Decoded versioned extensions of one document-level `ProgTags` container.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentTagExtensions {
    /// Decoded `___PPT9` tag, when present.
    pub powerpoint9: Option<DocBinaryTagExtension9>,
    /// Decoded `___PPT10` tag, when present.
    pub powerpoint10: Option<DocBinaryTagExtension10>,
    /// Decoded `___PPT11` tag, when present.
    pub powerpoint11: Option<DocBinaryTagExtension11>,
    /// Decoded `___PPT12` tag, when present.
    pub powerpoint12: Option<DocBinaryTagExtension12>,
}

/// Decoded versioned extensions of one slide-level `ProgTags` container.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SlideTagExtensions {
    /// Decoded `___PPT9` tag, when present.
    pub powerpoint9: Option<SlideBinaryTagExtension9>,
    /// Decoded `___PPT10` tag, when present.
    pub powerpoint10: Option<SlideBinaryTagExtension10>,
    /// Decoded `___PPT12` tag, when present.
    pub powerpoint12: Option<SlideBinaryTagExtension12>,
}
