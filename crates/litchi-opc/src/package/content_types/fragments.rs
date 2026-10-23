//! Exact Default/Override transport for durable metadata edits.
use super::super::{
    OwnedContentTypes,
    metadata_fragments::{self, Fragment, FragmentEdit, FragmentPlan, Fragments, Kind, Parsed},
};
use crate::{OpcError, Result, SourceLimits};
use std::sync::Arc;

/// Native content-type mapping identity. Default extensions and override part
/// names occupy distinct key spaces; both match ASCII case-insensitively.
/// Values returned by fragments use normalized lowercase spelling. Raw XML
/// spelling remains available from [`ContentTypeFragment::bytes`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ContentTypeFragmentKey<'a> {
    /// A Default mapping's XML-normalized extension (without a leading dot).
    Default(&'a str),
    /// An Override mapping's XML-normalized absolute OPC part name.
    Override(&'a str),
}
impl<'a> ContentTypeFragmentKey<'a> {
    fn from_inner(value: &'a str) -> Self {
        if let Some(extension) = value.strip_prefix('D') {
            Self::Default(extension)
        } else {
            Self::Override(value.strip_prefix('O').unwrap_or(value))
        }
    }
    fn encoded(self) -> Result<String> {
        match self {
            Self::Default(value) => metadata_fragments::content_key(true, value),
            Self::Override(value) => metadata_fragments::content_key(false, value),
        }
    }
}

/// Bounded lexical content-type mappings. Decoding confers no mutation authority;
/// applying edits still requires exact destination guards and package audits.
///
/// ```
/// use litchi_opc::{ContentTypeFragmentKey, ContentTypeFragments, SourceLimits};
/// use std::sync::Arc;
/// let xml = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="PNG" ContentType="image/png"/></Types>"#;
/// let entries = ContentTypeFragments::from_bytes(
///     Arc::new(xml.to_vec()), SourceLimits::default())?;
/// let image = entries.get(ContentTypeFragmentKey::Default("png"))?.unwrap();
/// assert!(image.bytes().starts_with(b"<Default"));
/// # Ok::<(), litchi_opc::OpcError>(())
/// ```
#[derive(Debug)]
pub struct ContentTypeFragments {
    inner: Fragments,
}
impl ContentTypeFragments {
    /// Decode untrusted XML under explicit byte, mapping, event and XML limits.
    /// Duplicate native keys and nonstandard roots cannot be edited.
    pub fn from_bytes(bytes: Arc<Vec<u8>>, limits: SourceLimits) -> Result<Self> {
        Ok(Self {
            inner: Fragments::from_bytes(Kind::ContentTypes, bytes, limits)?,
        })
    }
    /// Borrow exact mappings in source order, without copying their XML.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = ContentTypeFragment<'_>> {
        self.inner.iter().map(|inner| ContentTypeFragment { inner })
    }
    /// Resolve a native mapping key with ASCII-insensitive matching.
    /// Whitespace is not collapsed beyond XML attribute normalization at decode.
    pub fn get(&self, key: ContentTypeFragmentKey<'_>) -> Result<Option<ContentTypeFragment<'_>>> {
        Ok(self
            .inner
            .get(&key.encoded()?)
            .map(|inner| ContentTypeFragment { inner }))
    }
}
/// Borrowed exact Default or Override element and its retained gap anchors.
#[derive(Clone, Copy, Debug)]
pub struct ContentTypeFragment<'a> {
    inner: Fragment<'a>,
}
impl<'a> ContentTypeFragment<'a> {
    /// Native mapping identity, with normalized lowercase spelling.
    pub fn key(&self) -> ContentTypeFragmentKey<'a> {
        ContentTypeFragmentKey::from_inner(self.inner.id())
    }
    /// Exact element range in the catalog's original XML, including its BOM offset.
    pub fn byte_range(&self) -> std::ops::Range<usize> {
        self.inner.byte_range()
    }
    /// Exact element bytes, including attribute spelling and paired tags.
    pub fn bytes(&self) -> &'a [u8] {
        self.inner.bytes()
    }
    /// Next mapping in source order, if present.
    pub fn following_key(&self) -> Option<ContentTypeFragmentKey<'a>> {
        self.inner
            .following_id()
            .map(ContentTypeFragmentKey::from_inner)
    }
    /// Exact bytes between this element and its next mapping or root close.
    pub fn trailing_bytes(&self) -> &'a [u8] {
        self.inner.trailing_bytes()
    }
}
/// Guarded lexical delta. Before and after mappings must have the same native key.
#[derive(Clone, Copy, Debug)]
pub struct ContentTypeFragmentEdit<'a> {
    /// Exact expected destination element, or absence for insertion.
    pub before: Option<ContentTypeFragment<'a>>,
    /// Exact replacement element and anchors, or absence for removal.
    pub after: Option<ContentTypeFragment<'a>>,
}
/// Source-bound exact-byte mapping edits admitted before output allocation.
/// Package-wide reservations and format dependency closure remain caller-owned.
pub struct ContentTypeFragmentPlan<'a> {
    source: &'a OwnedContentTypes,
    inner: FragmentPlan<'a>,
}
impl std::fmt::Debug for ContentTypeFragmentPlan<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("ContentTypeFragmentPlan")
            .field("plan", &self.inner)
            .finish_non_exhaustive()
    }
}
impl OwnedContentTypes {
    /// Capture exact Default and Override entries for a durable metadata delta.
    pub fn fragments(&self, limits: SourceLimits) -> Result<ContentTypeFragments> {
        ContentTypeFragments::from_bytes(Arc::clone(&self.xml.bytes), self.xml.limits().min(limits))
    }
    /// Plan exact mapping deltas while retaining unrelated destination XML.
    ///
    /// Touched elements and inherited XML context must match. Replacements stay
    /// in destination slots; insertions prove neighboring keys and exact gaps.
    /// Unprovable, cyclic or overlapping anchors fail before output allocation.
    pub fn plan_fragment_edits<'a>(
        &'a self,
        edits: &[ContentTypeFragmentEdit<'a>],
        limits: SourceLimits,
    ) -> Result<ContentTypeFragmentPlan<'a>> {
        let limits = self.xml.limits().min(limits);
        Kind::ContentTypes.admit_edit_count(edits.len(), limits)?;
        let mut inner_edits = Vec::new();
        inner_edits
            .try_reserve_exact(edits.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC content-type fragment edits",
                source,
            })?;
        inner_edits.extend(edits.iter().map(|edit| FragmentEdit {
            before: edit.before.map(|value| value.inner),
            after: edit.after.map(|value| value.inner),
        }));
        let inner = FragmentPlan::new(&self.xml, Kind::ContentTypes, &inner_edits, limits)?;
        Ok(ContentTypeFragmentPlan {
            source: self,
            inner,
        })
    }
}
impl ContentTypeFragmentPlan<'_> {
    /// Exact final manifest byte length.
    pub const fn final_len(&self) -> usize {
        self.inner.final_len()
    }
    /// Exact final number of Default and Override mappings.
    pub const fn mapping_count(&self) -> usize {
        self.inner.entry_count()
    }
    /// Conservative XML event bound, including EOF. Deletions may merge text.
    pub const fn event_upper_bound(&self) -> usize {
        self.inner.event_upper_bound()
    }
    /// Whether materialization returns the exact source image.
    pub const fn is_noop(&self) -> bool {
        self.inner.is_noop()
    }
    /// Materialize and validate the native binding and XML, even for no-ops.
    pub fn materialize(self, limits: SourceLimits) -> Result<OwnedContentTypes> {
        let noop = self.inner.is_noop();
        let (xml, parsed) = self.inner.materialize(limits)?;
        let Parsed::ContentTypes(binding) = parsed else {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-type fragment binding kind differs".into(),
            ));
        };
        Ok(OwnedContentTypes {
            xml,
            binding: if noop {
                Arc::clone(&self.source.binding)
            } else {
                Arc::new(binding)
            },
        })
    }
}
