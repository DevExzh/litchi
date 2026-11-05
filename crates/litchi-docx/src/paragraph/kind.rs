//! What kind of paragraph child an [`super::OpaqueInline`] is.

use quick_xml::name::{Namespace, QName, ResolveResult};

use super::codec::is_fragment_word_name;

/// The kind of a direct paragraph child kept as inert XML
/// ([`super::OpaqueInline`]).
///
/// The kind is classified by namespace and local name when the paragraph's
/// children are read, so asking for it parses nothing. A WordprocessingML
/// element must be in the transitional or strict namespace (or written with
/// the paragraph's own prefix when the fragment does not declare it); Office
/// Math and markup-compatibility elements are recognized by namespace, or by
/// their conventional `m` and `mc` prefixes when the fragment does not
/// declare them. Anything else, including future WordprocessingML elements
/// and elements of other namespaces, is [`InlineKind::Other`].
///
/// ```
/// use litchi_docx::{Inline, Paragraph};
/// use litchi_docx::paragraph::InlineKind;
///
/// let paragraph = Paragraph::new(
///     br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple><w:ins w:id="1" w:author="A"><w:r><w:t>new</w:t></w:r></w:ins></w:p>"#.to_vec(),
/// );
/// let kinds: Vec<InlineKind> = paragraph
///     .inlines()?
///     .iter()
///     .filter_map(|inline| match inline {
///         Inline::Unknown(opaque) => Some(opaque.kind()),
///         _ => None,
///     })
///     .collect();
/// assert_eq!(kinds, [InlineKind::SimpleField, InlineKind::Insertion]);
/// assert_eq!(InlineKind::SimpleField.local_name(), Some("fldSimple"));
/// # Ok::<(), litchi_docx::Error>(())
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum InlineKind {
    /// `w:hyperlink`. [`crate::Paragraph::inlines_with_relationships`]
    /// resolves these to [`super::Inline::Hyperlink`].
    Hyperlink,
    /// `w:fldSimple`: a simple field and its cached result.
    SimpleField,
    /// `w:subDoc`: a subdocument anchor.
    SubDocument,
    /// `w:customXml`: run-level custom XML markup.
    CustomXml,
    /// `w:smartTag`
    SmartTag,
    /// `w:sdt`: an inline content control.
    ContentControl,
    /// `w:dir`: bidirectional embedding.
    BidiEmbedding,
    /// `w:bdo`: bidirectional override.
    BidiOverride,
    /// `w:proofErr`
    ProofError,
    /// `w:permStart`
    PermissionStart,
    /// `w:permEnd`
    PermissionEnd,
    /// `w:bookmarkStart` without a readable identifier (a readable one is
    /// [`super::Inline::BookmarkStart`]).
    BookmarkStart,
    /// `w:bookmarkEnd` without a readable identifier (a readable one is
    /// [`super::Inline::BookmarkEnd`]).
    BookmarkEnd,
    /// `w:commentRangeStart`
    CommentRangeStart,
    /// `w:commentRangeEnd`
    CommentRangeEnd,
    /// `w:moveFromRangeStart`
    MoveFromRangeStart,
    /// `w:moveFromRangeEnd`
    MoveFromRangeEnd,
    /// `w:moveToRangeStart`
    MoveToRangeStart,
    /// `w:moveToRangeEnd`
    MoveToRangeEnd,
    /// `w:customXmlInsRangeStart`
    CustomXmlInsertionRangeStart,
    /// `w:customXmlInsRangeEnd`
    CustomXmlInsertionRangeEnd,
    /// `w:customXmlDelRangeStart`
    CustomXmlDeletionRangeStart,
    /// `w:customXmlDelRangeEnd`
    CustomXmlDeletionRangeEnd,
    /// `w:customXmlMoveFromRangeStart`
    CustomXmlMoveFromRangeStart,
    /// `w:customXmlMoveFromRangeEnd`
    CustomXmlMoveFromRangeEnd,
    /// `w:customXmlMoveToRangeStart`
    CustomXmlMoveToRangeStart,
    /// `w:customXmlMoveToRangeEnd`
    CustomXmlMoveToRangeEnd,
    /// `w:ins`: a tracked insertion.
    Insertion,
    /// `w:del`: a tracked deletion.
    Deletion,
    /// `w:moveFrom`: the source of a tracked move.
    MoveFrom,
    /// `w:moveTo`: the destination of a tracked move.
    MoveTo,
    /// `m:oMath`: inline Office Math.
    Math,
    /// `m:oMathPara`: display Office Math.
    MathParagraph,
    /// `mc:AlternateContent`, present only in markup that has not been
    /// processed for markup compatibility.
    AlternateContent,
    /// Any other element.
    Other,
}

/// WordprocessingML paragraph children and their kinds.
const WORD_KINDS: [(&[u8], InlineKind); 31] = [
    (b"hyperlink", InlineKind::Hyperlink),
    (b"fldSimple", InlineKind::SimpleField),
    (b"subDoc", InlineKind::SubDocument),
    (b"customXml", InlineKind::CustomXml),
    (b"smartTag", InlineKind::SmartTag),
    (b"sdt", InlineKind::ContentControl),
    (b"dir", InlineKind::BidiEmbedding),
    (b"bdo", InlineKind::BidiOverride),
    (b"proofErr", InlineKind::ProofError),
    (b"permStart", InlineKind::PermissionStart),
    (b"permEnd", InlineKind::PermissionEnd),
    (b"bookmarkStart", InlineKind::BookmarkStart),
    (b"bookmarkEnd", InlineKind::BookmarkEnd),
    (b"commentRangeStart", InlineKind::CommentRangeStart),
    (b"commentRangeEnd", InlineKind::CommentRangeEnd),
    (b"moveFromRangeStart", InlineKind::MoveFromRangeStart),
    (b"moveFromRangeEnd", InlineKind::MoveFromRangeEnd),
    (b"moveToRangeStart", InlineKind::MoveToRangeStart),
    (b"moveToRangeEnd", InlineKind::MoveToRangeEnd),
    (
        b"customXmlInsRangeStart",
        InlineKind::CustomXmlInsertionRangeStart,
    ),
    (
        b"customXmlInsRangeEnd",
        InlineKind::CustomXmlInsertionRangeEnd,
    ),
    (
        b"customXmlDelRangeStart",
        InlineKind::CustomXmlDeletionRangeStart,
    ),
    (
        b"customXmlDelRangeEnd",
        InlineKind::CustomXmlDeletionRangeEnd,
    ),
    (
        b"customXmlMoveFromRangeStart",
        InlineKind::CustomXmlMoveFromRangeStart,
    ),
    (
        b"customXmlMoveFromRangeEnd",
        InlineKind::CustomXmlMoveFromRangeEnd,
    ),
    (
        b"customXmlMoveToRangeStart",
        InlineKind::CustomXmlMoveToRangeStart,
    ),
    (
        b"customXmlMoveToRangeEnd",
        InlineKind::CustomXmlMoveToRangeEnd,
    ),
    (b"ins", InlineKind::Insertion),
    (b"del", InlineKind::Deletion),
    (b"moveFrom", InlineKind::MoveFrom),
    (b"moveTo", InlineKind::MoveTo),
];

const MCE_NAMESPACE: &[u8] = litchi_ooxml_common::mce::NAMESPACE.as_bytes();

impl InlineKind {
    /// The element's local name (`"fldSimple"`, `"oMath"`,
    /// `"AlternateContent"`), or `None` for [`InlineKind::Other`].
    #[must_use]
    pub const fn local_name(self) -> Option<&'static str> {
        Some(match self {
            Self::Hyperlink => "hyperlink",
            Self::SimpleField => "fldSimple",
            Self::SubDocument => "subDoc",
            Self::CustomXml => "customXml",
            Self::SmartTag => "smartTag",
            Self::ContentControl => "sdt",
            Self::BidiEmbedding => "dir",
            Self::BidiOverride => "bdo",
            Self::ProofError => "proofErr",
            Self::PermissionStart => "permStart",
            Self::PermissionEnd => "permEnd",
            Self::BookmarkStart => "bookmarkStart",
            Self::BookmarkEnd => "bookmarkEnd",
            Self::CommentRangeStart => "commentRangeStart",
            Self::CommentRangeEnd => "commentRangeEnd",
            Self::MoveFromRangeStart => "moveFromRangeStart",
            Self::MoveFromRangeEnd => "moveFromRangeEnd",
            Self::MoveToRangeStart => "moveToRangeStart",
            Self::MoveToRangeEnd => "moveToRangeEnd",
            Self::CustomXmlInsertionRangeStart => "customXmlInsRangeStart",
            Self::CustomXmlInsertionRangeEnd => "customXmlInsRangeEnd",
            Self::CustomXmlDeletionRangeStart => "customXmlDelRangeStart",
            Self::CustomXmlDeletionRangeEnd => "customXmlDelRangeEnd",
            Self::CustomXmlMoveFromRangeStart => "customXmlMoveFromRangeStart",
            Self::CustomXmlMoveFromRangeEnd => "customXmlMoveFromRangeEnd",
            Self::CustomXmlMoveToRangeStart => "customXmlMoveToRangeStart",
            Self::CustomXmlMoveToRangeEnd => "customXmlMoveToRangeEnd",
            Self::Insertion => "ins",
            Self::Deletion => "del",
            Self::MoveFrom => "moveFrom",
            Self::MoveTo => "moveTo",
            Self::Math => "oMath",
            Self::MathParagraph => "oMathPara",
            Self::AlternateContent => "AlternateContent",
            Self::Other => return None,
        })
    }

    /// Classify a resolved paragraph child name.
    #[allow(
        clippy::ref_option,
        reason = "the fragment prefix is the shared tri-state of the paragraph scanners"
    )]
    pub(super) fn classify(
        namespace: &ResolveResult<'_>,
        name: QName<'_>,
        fragment_prefix: &Option<Option<Vec<u8>>>,
    ) -> Self {
        if let Some((_, kind)) = WORD_KINDS
            .iter()
            .find(|(word, _)| is_fragment_word_name(namespace, name, word, fragment_prefix))
        {
            return *kind;
        }
        if litchi_ooxml_common::xml::is_omml_name(namespace, name, b"oMath") {
            return Self::Math;
        }
        if litchi_ooxml_common::xml::is_omml_name(namespace, name, b"oMathPara") {
            return Self::MathParagraph;
        }
        let markup_compatibility = match namespace {
            ResolveResult::Bound(Namespace(value)) => *value == MCE_NAMESPACE,
            ResolveResult::Unknown(prefix) => prefix.as_slice() == b"mc",
            ResolveResult::Unbound => false,
        };
        if markup_compatibility && name.local_name().as_ref() == b"AlternateContent" {
            return Self::AlternateContent;
        }
        Self::Other
    }
}
