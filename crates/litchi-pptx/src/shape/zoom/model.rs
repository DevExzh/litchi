//! Semantic `PresentationML` zoom values and the lossless shape owner.

use std::slice;

use crate::{Error, Result};

use super::codec;

/// A bounded percentage in the `ST_Percentage` wire domain.
///
/// `PresentationML` stores percentages in thousandths of a percent. The whole
/// signed 32-bit range is the schema domain, so every value representable by
/// this type is safe to write without a late string/range escape hatch.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Percentage(i32);

impl Percentage {
    /// Zero percent.
    pub const ZERO: Self = Self(0);

    /// One hundred percent, represented in thousandths of a percent.
    pub const HUNDRED: Self = Self(100_000);

    /// Construct a schema-valid percentage in thousandths of a percent.
    #[inline]
    #[must_use]
    pub const fn new(value: i32) -> Self {
        Self(value)
    }

    /// Return the thousandths-of-a-percent value.
    #[inline]
    #[must_use]
    pub const fn value(self) -> i32 {
        self.0
    }
}

/// Whether a zoom uses the destination preview or a custom cover image.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum ImageType {
    /// Use the destination slide or section preview.
    #[default]
    Preview,
    /// Use the image relationship carried by `a:blipFill`.
    Cover,
}

/// Whether a `DrawingML` image reference is embedded or externally linked.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Link {
    /// `a:blip@r:embed`; the relationship must be internal.
    Embed,
    /// `a:blip@r:link`; the relationship must be external.
    External,
}

/// A zoom-owned image relationship and its resolved package target.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    pub(super) id: String,
    pub(super) link: Link,
    pub(super) target: Option<Target>,
}

impl Relationship {
    /// Construct a relationship reference without resolving a package target.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(id: impl Into<String>, link: Link) -> Result<Self> {
        let id = id.into();
        if id.is_empty() || id.len() > codec::MAX_RELATIONSHIP_ID_BYTES {
            return Err(Error::Invalid(
                "zoom image relationship ID is invalid".into(),
            ));
        }
        Ok(Self {
            id,
            link,
            target: None,
        })
    }

    /// Relationship ID used by `a:blip`.
    #[inline]
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Whether the reference is embedded or external.
    #[inline]
    #[must_use]
    pub const fn link(&self) -> Link {
        self.link
    }

    /// Resolved target metadata when this owner came from a package context.
    #[inline]
    #[must_use]
    pub fn target(&self) -> Option<&Target> {
        self.target.as_ref()
    }
}

/// A resolved OPC relationship target.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Target {
    /// An internal part with its validated package content type.
    Internal {
        /// Canonical target part name.
        part_name: String,
        /// Target part content type.
        content_type: String,
    },
    /// An external URI retained without opening or following it.
    External { uri: String },
}

impl Target {
    /// Return the internal OPC part name, if this is an embedded target.
    #[must_use]
    pub fn part_name(&self) -> Option<&str> {
        match self {
            Self::Internal { part_name, .. } => Some(part_name),
            Self::External { .. } => None,
        }
    }

    /// Return the validated content type, if this is an embedded target.
    #[must_use]
    pub fn content_type(&self) -> Option<&str> {
        match self {
            Self::Internal { content_type, .. } => Some(content_type),
            Self::External { .. } => None,
        }
    }

    /// Return the external URI, if this is an externally linked target.
    #[must_use]
    pub fn uri(&self) -> Option<&str> {
        match self {
            Self::External { uri } => Some(uri),
            Self::Internal { .. } => None,
        }
    }
}

/// Shared typed properties of every section, slide, and summary item.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Properties {
    pub(super) id: String,
    pub(super) return_to_parent: bool,
    pub(super) image_type: ImageType,
    pub(super) transition: Option<crate::time::Offset>,
    pub(super) show_background: bool,
    pub(super) blip_fill_xml: Vec<u8>,
    pub(super) shape_properties_xml: Vec<u8>,
    pub(super) image: Option<Relationship>,
}

impl Properties {
    /// Create typed zoom properties from the required `DrawingML` child XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(
        id: impl Into<String>,
        blip_fill_xml: impl Into<Vec<u8>>,
        shape_properties_xml: impl Into<Vec<u8>>,
    ) -> Result<Self> {
        let id = id.into();
        codec::validate_guid(&id)?;
        let blip_fill_xml = blip_fill_xml.into();
        let shape_properties_xml = shape_properties_xml.into();
        if blip_fill_xml.is_empty() || shape_properties_xml.is_empty() {
            return Err(Error::Invalid(
                "zoom properties require blipFill and spPr XML".into(),
            ));
        }
        let image = codec::parse_blip_relationship(&blip_fill_xml)?;
        Ok(Self {
            id,
            return_to_parent: true,
            image_type: ImageType::Preview,
            transition: None,
            show_background: true,
            blip_fill_xml,
            shape_properties_xml,
            image,
        })
    }

    /// Stable GUID identifying the zoom object.
    #[inline]
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Whether slideshow navigation returns to the parent slide.
    #[inline]
    #[must_use]
    pub const fn return_to_parent(&self) -> bool {
        self.return_to_parent
    }

    /// Set slideshow return-to-parent behavior.
    #[inline]
    pub fn set_return_to_parent(&mut self, value: bool) {
        self.return_to_parent = value;
    }

    /// Select preview or cover-image behavior.
    #[inline]
    #[must_use]
    pub const fn image_type(&self) -> ImageType {
        self.image_type
    }

    /// Set preview or cover-image behavior.
    #[inline]
    pub fn set_image_type(&mut self, value: ImageType) {
        self.image_type = value;
    }

    /// Optional transition duration.
    #[inline]
    #[must_use]
    pub fn transition(&self) -> Option<&crate::time::Offset> {
        self.transition.as_ref()
    }

    /// Set or clear the zoom transition duration.
    #[inline]
    pub fn set_transition(&mut self, value: Option<crate::time::Offset>) {
        self.transition = value;
    }

    /// Whether the destination background is shown during the zoom.
    #[inline]
    #[must_use]
    pub const fn show_background(&self) -> bool {
        self.show_background
    }

    /// Set destination-background behavior.
    #[inline]
    pub fn set_show_background(&mut self, value: bool) {
        self.show_background = value;
    }

    /// Borrow the preserved `a:blipFill` payload.
    #[inline]
    #[must_use]
    pub fn blip_fill_xml(&self) -> &[u8] {
        &self.blip_fill_xml
    }

    /// Replace the `a:blipFill` payload and refresh its typed relationship.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_blip_fill_xml(&mut self, value: impl Into<Vec<u8>>) -> Result<()> {
        let value = value.into();
        if value.is_empty() {
            return Err(Error::Invalid("zoom blipFill XML cannot be empty".into()));
        }
        let image = codec::parse_blip_relationship(&value)?;
        self.blip_fill_xml = value;
        self.image = image;
        Ok(())
    }

    /// Borrow the preserved `a:spPr` payload.
    #[inline]
    #[must_use]
    pub fn shape_properties_xml(&self) -> &[u8] {
        &self.shape_properties_xml
    }

    /// Replace the preserved `a:spPr` payload.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_shape_properties_xml(&mut self, value: impl Into<Vec<u8>>) -> Result<()> {
        let value = value.into();
        if value.is_empty() {
            return Err(Error::Invalid(
                "zoom shape-properties XML cannot be empty".into(),
            ));
        }
        self.shape_properties_xml = value;
        Ok(())
    }

    /// Borrow the image relationship metadata, if `a:blip` declares one.
    #[inline]
    #[must_use]
    pub fn image_relationship(&self) -> Option<&Relationship> {
        self.image.as_ref()
    }

    pub(super) fn image_relationship_mut(&mut self) -> Option<&mut Relationship> {
        self.image.as_mut()
    }
}

/// A section zoom object.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Section {
    pub(super) section_id: String,
    pub(super) properties: Properties,
    pub(super) fallback_xml: Vec<u8>,
    pub(super) extension_xml: Option<Vec<u8>>,
    pub(super) unknown_xml: Vec<Vec<u8>>,
}

impl Section {
    /// Create a section zoom with a preserved or authored `p:pic` fallback.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(
        section_id: impl Into<String>,
        properties: Properties,
        fallback_xml: impl Into<Vec<u8>>,
    ) -> Result<Self> {
        let section_id = section_id.into();
        codec::validate_guid(&section_id)?;
        let fallback_xml = fallback_xml.into();
        if fallback_xml.is_empty() {
            return Err(Error::Invalid(
                "section zoom fallback cannot be empty".into(),
            ));
        }
        Ok(Self {
            section_id,
            properties,
            fallback_xml,
            extension_xml: None,
            unknown_xml: Vec::new(),
        })
    }

    /// Stable section GUID targeted by this object.
    #[inline]
    #[must_use]
    pub fn section_id(&self) -> &str {
        &self.section_id
    }

    /// Shared zoom properties.
    #[inline]
    #[must_use]
    pub fn properties(&self) -> &Properties {
        &self.properties
    }

    /// Mutable shared zoom properties.
    #[inline]
    pub fn properties_mut(&mut self) -> &mut Properties {
        &mut self.properties
    }

    /// Borrow the exact fallback shape XML.
    #[inline]
    #[must_use]
    pub fn fallback_xml(&self) -> &[u8] {
        &self.fallback_xml
    }

    /// Borrow the optional preserved object extension list.
    #[inline]
    #[must_use]
    pub fn extension_xml(&self) -> Option<&[u8]> {
        self.extension_xml.as_deref()
    }

    /// Borrow unsupported choice payloads retained beside the typed choice.
    #[inline]
    #[must_use]
    pub fn unknown_xml(&self) -> &[Vec<u8>] {
        &self.unknown_xml
    }
}

/// A slide zoom object.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Slide {
    pub(super) slide_id: u32,
    pub(super) creation_id: Option<u32>,
    pub(super) properties: Properties,
    pub(super) fallback_xml: Vec<u8>,
    pub(super) extension_xml: Option<Vec<u8>>,
    pub(super) unknown_xml: Vec<Vec<u8>>,
}

impl Slide {
    /// Create a slide zoom with a preserved or authored `p:pic` fallback.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(
        slide_id: u32,
        properties: Properties,
        fallback_xml: impl Into<Vec<u8>>,
    ) -> Result<Self> {
        codec::validate_slide_id(slide_id)?;
        let fallback_xml = fallback_xml.into();
        if fallback_xml.is_empty() {
            return Err(Error::Invalid("slide zoom fallback cannot be empty".into()));
        }
        Ok(Self {
            slide_id,
            creation_id: None,
            properties,
            fallback_xml,
            extension_xml: None,
            unknown_xml: Vec::new(),
        })
    }

    /// Stable `PresentationML` slide ID targeted by this object.
    #[inline]
    #[must_use]
    pub const fn slide_id(&self) -> u32 {
        self.slide_id
    }

    /// Optional creation ID associated with the slide zoom.
    #[inline]
    #[must_use]
    pub const fn creation_id(&self) -> Option<u32> {
        self.creation_id
    }

    /// Set or clear the optional creation ID.
    #[inline]
    pub fn set_creation_id(&mut self, value: Option<u32>) {
        self.creation_id = value;
    }

    /// Shared zoom properties.
    #[inline]
    #[must_use]
    pub fn properties(&self) -> &Properties {
        &self.properties
    }

    /// Mutable shared zoom properties.
    #[inline]
    pub fn properties_mut(&mut self) -> &mut Properties {
        &mut self.properties
    }

    /// Borrow the exact fallback shape XML.
    #[inline]
    #[must_use]
    pub fn fallback_xml(&self) -> &[u8] {
        &self.fallback_xml
    }

    /// Borrow the optional preserved object extension list.
    #[inline]
    #[must_use]
    pub fn extension_xml(&self) -> Option<&[u8]> {
        self.extension_xml.as_deref()
    }

    /// Borrow unsupported choice payloads retained beside the typed choice.
    #[inline]
    #[must_use]
    pub fn unknown_xml(&self) -> &[Vec<u8>] {
        &self.unknown_xml
    }
}

/// Grid or fixed layout for a summary zoom.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Layout {
    /// PowerPoint-managed grid layout.
    Grid,
    /// User-positioned fixed layout.
    Fixed,
}

/// One section item inside a summary zoom.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Item {
    pub(super) section_id: String,
    pub(super) title: String,
    pub(super) description: String,
    pub(super) offset_x: Percentage,
    pub(super) offset_y: Percentage,
    pub(super) scale_x: Percentage,
    pub(super) scale_y: Percentage,
    pub(super) properties: Properties,
    pub(super) extension_xml: Option<Vec<u8>>,
    pub(super) unknown_xml: Vec<Vec<u8>>,
}

impl Item {
    /// Create a summary item with schema defaults for text, offset, and scale.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(section_id: impl Into<String>, properties: Properties) -> Result<Self> {
        let section_id = section_id.into();
        codec::validate_guid(&section_id)?;
        Ok(Self {
            section_id,
            title: String::new(),
            description: String::new(),
            offset_x: Percentage::ZERO,
            offset_y: Percentage::ZERO,
            scale_x: Percentage::HUNDRED,
            scale_y: Percentage::HUNDRED,
            properties,
            extension_xml: None,
            unknown_xml: Vec::new(),
        })
    }

    /// Stable section GUID targeted by this summary item.
    #[inline]
    #[must_use]
    pub fn section_id(&self) -> &str {
        &self.section_id
    }

    /// Alternative-text title.
    #[inline]
    #[must_use]
    pub fn title(&self) -> &str {
        &self.title
    }

    /// Set alternative-text title.
    #[inline]
    pub fn set_title(&mut self, value: impl Into<String>) {
        self.title = value.into();
    }

    /// Alternative-text description.
    #[inline]
    #[must_use]
    pub fn description(&self) -> &str {
        &self.description
    }

    /// Set alternative-text description.
    #[inline]
    pub fn set_description(&mut self, value: impl Into<String>) {
        self.description = value.into();
    }

    /// Horizontal layout offset.
    #[inline]
    #[must_use]
    pub const fn offset_x(&self) -> Percentage {
        self.offset_x
    }

    /// Vertical layout offset.
    #[inline]
    #[must_use]
    pub const fn offset_y(&self) -> Percentage {
        self.offset_y
    }

    /// Set both layout offsets.
    #[inline]
    pub fn set_offsets(&mut self, x: Percentage, y: Percentage) {
        self.offset_x = x;
        self.offset_y = y;
    }

    /// Horizontal layout scale.
    #[inline]
    #[must_use]
    pub const fn scale_x(&self) -> Percentage {
        self.scale_x
    }

    /// Vertical layout scale.
    #[inline]
    #[must_use]
    pub const fn scale_y(&self) -> Percentage {
        self.scale_y
    }

    /// Set both layout scales.
    #[inline]
    pub fn set_scales(&mut self, x: Percentage, y: Percentage) {
        self.scale_x = x;
        self.scale_y = y;
    }

    /// Shared zoom properties.
    #[inline]
    #[must_use]
    pub fn properties(&self) -> &Properties {
        &self.properties
    }

    /// Mutable shared zoom properties.
    #[inline]
    pub fn properties_mut(&mut self) -> &mut Properties {
        &mut self.properties
    }

    /// Borrow the optional preserved object extension list.
    #[inline]
    #[must_use]
    pub fn extension_xml(&self) -> Option<&[u8]> {
        self.extension_xml.as_deref()
    }

    /// Borrow unsupported choice payloads retained beside the typed choice.
    #[inline]
    #[must_use]
    pub fn unknown_xml(&self) -> &[Vec<u8>] {
        &self.unknown_xml
    }
}

/// A summary zoom container containing zero or more section items.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Summary {
    pub(super) items: Vec<Item>,
    pub(super) layout: Layout,
    pub(super) fallback_xml: Vec<u8>,
    pub(super) extension_xml: Option<Vec<u8>>,
    pub(super) unknown_xml: Vec<Vec<u8>>,
}

impl Summary {
    /// Create a summary zoom with a preserved or authored `p:grpSp` fallback.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(layout: Layout, fallback_xml: impl Into<Vec<u8>>) -> Result<Self> {
        let fallback_xml = fallback_xml.into();
        if fallback_xml.is_empty() {
            return Err(Error::Invalid(
                "summary zoom fallback cannot be empty".into(),
            ));
        }
        Ok(Self {
            items: Vec::new(),
            layout,
            fallback_xml,
            extension_xml: None,
            unknown_xml: Vec::new(),
        })
    }

    /// Layout mode used by the summary container.
    #[inline]
    #[must_use]
    pub const fn layout(&self) -> Layout {
        self.layout
    }

    /// Set the summary layout mode.
    #[inline]
    pub fn set_layout(&mut self, value: Layout) {
        self.layout = value;
    }

    /// Borrow summary items in source order.
    #[inline]
    #[must_use]
    pub fn items(&self) -> &[Item] {
        &self.items
    }

    /// Add one summary item, rejecting duplicate section or object IDs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_item(&mut self, item: Item) -> Result<()> {
        if self.items.iter().any(|value| {
            value.section_id == item.section_id || value.properties.id == item.properties.id
        }) {
            return Err(Error::Invalid(
                "summary zoom contains a duplicate section or object ID".into(),
            ));
        }
        self.items.push(item);
        Ok(())
    }

    /// Remove a summary item by zero-based position.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_item(&mut self, index: usize) -> Result<Item> {
        if index >= self.items.len() {
            return Err(Error::Invalid(format!(
                "summary zoom item index {index} is outside {} items",
                self.items.len()
            )));
        }
        Ok(self.items.remove(index))
    }

    /// Borrow the exact fallback group XML.
    #[inline]
    #[must_use]
    pub fn fallback_xml(&self) -> &[u8] {
        &self.fallback_xml
    }

    /// Borrow the optional preserved container extension list.
    #[inline]
    #[must_use]
    pub fn extension_xml(&self) -> Option<&[u8]> {
        self.extension_xml.as_deref()
    }

    /// Borrow unsupported choice payloads retained beside the typed choice.
    #[inline]
    #[must_use]
    pub fn unknown_xml(&self) -> &[Vec<u8>] {
        &self.unknown_xml
    }
}

/// An unsupported or future zoom alternate-content payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Unknown {
    pub(super) xml: Vec<u8>,
}

impl Unknown {
    /// Preserve an opaque `mc:AlternateContent` element.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        if xml.is_empty() {
            return Err(Error::Invalid("unknown zoom XML cannot be empty".into()));
        }
        Ok(Self { xml })
    }

    /// Borrow the exact opaque alternate-content XML.
    #[inline]
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }
}

/// One typed or opaque zoom shape metadata entry.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Zoom {
    /// Section-targeting zoom.
    Section(Section),
    /// Slide-targeting zoom.
    Slide(Slide),
    /// Summary container zoom.
    Summary(Summary),
    /// Unsupported future zoom markup retained losslessly.
    Unknown(Unknown),
}

impl Zoom {
    /// Shared properties for a section or slide zoom.
    #[must_use]
    pub fn properties(&self) -> Option<&Properties> {
        match self {
            Self::Section(value) => Some(&value.properties),
            Self::Slide(value) => Some(&value.properties),
            Self::Summary(_) | Self::Unknown(_) => None,
        }
    }

    /// Mutably borrow shared properties for a section or slide zoom.
    pub fn properties_mut(&mut self) -> Option<&mut Properties> {
        match self {
            Self::Section(value) => Some(&mut value.properties),
            Self::Slide(value) => Some(&mut value.properties),
            Self::Summary(_) | Self::Unknown(_) => None,
        }
    }
}

/// Lossless semantic owner for the zoom alternate-content entries of one
/// slide XML part.
///
/// The owner retains the complete current slide XML. CRUD patches only the
/// individual `mc:AlternateContent` spans, so unrelated shapes, fallbacks,
/// namespace aliases, comments, and unknown extension markup remain intact.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Owner {
    pub(super) xml: Vec<u8>,
    pub(super) base_xml: Vec<u8>,
    pub(super) entries: Vec<Zoom>,
    pub(super) spans: Vec<(usize, usize)>,
    pub(super) insert_at: usize,
    pub(super) insert_owner: Option<(usize, usize, bool)>,
    pub(super) namespaces: Vec<(String, String)>,
    pub(super) pml_namespace: String,
    pub(super) dml_namespace: String,
    pub(super) relationship_namespace: String,
}

impl Owner {
    /// Read zoom entries from a complete slide or shape-tree XML owner.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn read(xml: &[u8]) -> Result<Self> {
        codec::read_owner(xml)
    }

    /// Alias for [`Self::read`] used by package-facing code.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    #[inline]
    pub fn from_xml(xml: &[u8]) -> Result<Self> {
        Self::read(xml)
    }

    /// Current complete owner XML.
    #[inline]
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }

    /// Serialize the current owner without normalizing unrelated markup.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    #[inline]
    pub fn to_xml(&self) -> Result<Vec<u8>> {
        Ok(self.xml.clone())
    }

    /// Number of zoom alternate-content entries, including unknown entries.
    #[inline]
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Whether the owner contains no zoom entries.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Borrow one zoom by checked source order.
    #[inline]
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&Zoom> {
        self.entries.get(index)
    }

    /// Iterate zoom entries in source order.
    #[inline]
    pub fn iter(&self) -> slice::Iter<'_, Zoom> {
        self.entries.iter()
    }

    /// Append one zoom to the first shape tree in the owner.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add(&mut self, value: Zoom) -> Result<usize> {
        let fragment = codec::write_zoom(&value, self)?;
        let mut xml = self.xml.clone();
        let index = self.entries.len();
        let Some((start, end, empty)) = self.insert_owner else {
            return Err(Error::Invalid(
                "zoom owner has no PresentationML spTree or grpSp insertion site".into(),
            ));
        };
        if empty {
            let local = self
                .xml
                .get(start..end)
                .ok_or_else(|| Error::Invalid("zoom insertion site is outside owner XML".into()))?;
            let close = local
                .iter()
                .position(|byte| *byte == b'>')
                .ok_or_else(|| Error::Invalid("empty shape owner has no end tag".into()))?;
            let name_end = local
                .get(..close)
                .and_then(|bytes| bytes.iter().rposition(|byte| *byte == b'<'))
                .ok_or_else(|| Error::Invalid("empty shape owner has no start tag".into()))?;
            let name = &local[name_end + 1..close];
            let name = name
                .strip_suffix(b"/")
                .ok_or_else(|| Error::Invalid("shape owner is not self-closing".into()))?;
            let mut replacement = Vec::with_capacity(local.len() + fragment.len() + name.len() + 3);
            replacement.extend_from_slice(&local[..close]);
            replacement.push(b'>');
            replacement.extend_from_slice(&fragment);
            replacement.extend_from_slice(b"</");
            replacement.extend_from_slice(name);
            replacement.push(b'>');
            replace_range(&mut xml, start, end, &replacement)?;
        } else {
            let insertion = self.insert_at;
            if insertion > xml.len() {
                return Err(Error::Invalid(
                    "zoom insertion offset is outside owner XML".into(),
                ));
            }
            xml.splice(insertion..insertion, fragment);
        }
        self.reparse(xml)?;
        Ok(index)
    }

    /// Replace one zoom and return the previous semantic value.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace(&mut self, index: usize, value: Zoom) -> Result<Zoom> {
        let old =
            self.entries.get(index).cloned().ok_or_else(|| {
                Error::Invalid(format!("zoom index {index} is outside the owner"))
            })?;
        let fragment = codec::write_zoom(&value, self)?;
        let (start, end) = self
            .spans
            .get(index)
            .copied()
            .ok_or_else(|| Error::Invalid("zoom span disappeared during replacement".into()))?;
        let mut xml = self.xml.clone();
        replace_range(&mut xml, start, end, &fragment)?;
        self.reparse(xml)?;
        Ok(old)
    }

    /// Remove one zoom and return the previous semantic value.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove(&mut self, index: usize) -> Result<Zoom> {
        let old =
            self.entries.get(index).cloned().ok_or_else(|| {
                Error::Invalid(format!("zoom index {index} is outside the owner"))
            })?;
        let (start, end) = self
            .spans
            .get(index)
            .copied()
            .ok_or_else(|| Error::Invalid("zoom span disappeared during removal".into()))?;
        let mut xml = self.xml.clone();
        replace_range(&mut xml, start, end, &[])?;
        self.reparse(xml)?;
        Ok(old)
    }

    /// Remove every zoom entry while retaining all unrelated slide markup.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn clear(&mut self) -> Result<()> {
        let mut xml = self.xml.clone();
        for &(start, end) in self.spans.iter().rev() {
            replace_range(&mut xml, start, end, &[])?;
        }
        self.reparse(xml)
    }

    pub(super) fn reparse(&mut self, xml: Vec<u8>) -> Result<()> {
        let base = self.base_xml.clone();
        let mut parsed = codec::read_owner(&xml)?;
        parsed.base_xml = base;
        *self = parsed;
        Ok(())
    }

    pub(crate) fn base_xml(&self) -> &[u8] {
        &self.base_xml
    }

    pub(crate) fn validate_in_package(
        &mut self,
        package: &litchi_opc::OpcPackage,
        owner: &dyn litchi_opc::Part,
    ) -> Result<()> {
        codec::validate_in_package(self, package, owner)
    }
}

impl<'a> IntoIterator for &'a Owner {
    type Item = &'a Zoom;
    type IntoIter = slice::Iter<'a, Zoom>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

fn replace_range(xml: &mut Vec<u8>, start: usize, end: usize, replacement: &[u8]) -> Result<()> {
    if start > end || end > xml.len() {
        return Err(Error::Invalid(
            "zoom patch span is outside owner XML".into(),
        ));
    }
    xml.splice(start..end, replacement.iter().copied());
    Ok(())
}
