/// Presentation part - the main part in a .pptx package.
///
/// Corresponds to `/ppt/presentation.xml` in the package.
use crate::common::xml::unqualified_attribute_value;
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::{is_presentationml_name, relationship_attribute_value};
use litchi_opc::part::Part;
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::sync::Arc;

const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";

/// The presentation slide surface dimensions and declared size type.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SlideSize {
    width: i64,
    height: i64,
    size_type: Option<String>,
}

impl SlideSize {
    /// Return the slide width in EMUs.
    #[inline]
    pub const fn width(&self) -> i64 {
        self.width
    }

    /// Return the slide height in EMUs.
    #[inline]
    pub const fn height(&self) -> i64 {
        self.height
    }

    /// Return the size type declared on the presentation, if present.
    ///
    /// The raw value is preserved so documents with future size types remain
    /// inspectable.
    #[inline]
    pub fn size_type(&self) -> Option<&str> {
        self.size_type.as_deref()
    }
}

/// The notes and handout surface dimensions.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct NotesSize {
    width: i64,
    height: i64,
}

impl NotesSize {
    /// Return the notes surface width in EMUs.
    #[inline]
    pub const fn width(&self) -> i64 {
        self.width
    }

    /// Return the notes surface height in EMUs.
    #[inline]
    pub const fn height(&self) -> i64 {
        self.height
    }
}

/// The effective conformance class for a presentation root.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PresentationConformance {
    /// The standard transitional conformance class.
    Transitional,
    /// The strict conformance class.
    Strict,
}

impl Default for PresentationConformance {
    fn default() -> Self {
        Self::Transitional
    }
}

/// Root-level presentation behavior and document settings with schema defaults applied.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PresentationMetadata {
    server_zoom: i32,
    first_slide_number: i32,
    show_special_placeholders_on_title_slide: bool,
    right_to_left: bool,
    remove_personal_info_on_save: bool,
    compatibility_mode: bool,
    strict_first_and_last_chars: bool,
    embed_true_type_fonts: bool,
    save_subset_fonts: bool,
    auto_compress_pictures: bool,
    bookmark_id_seed: u32,
    conformance: PresentationConformance,
}

impl Default for PresentationMetadata {
    fn default() -> Self {
        Self {
            server_zoom: 50_000,
            first_slide_number: 1,
            show_special_placeholders_on_title_slide: true,
            right_to_left: false,
            remove_personal_info_on_save: false,
            compatibility_mode: false,
            strict_first_and_last_chars: true,
            embed_true_type_fonts: false,
            save_subset_fonts: false,
            auto_compress_pictures: true,
            bookmark_id_seed: 1,
            conformance: PresentationConformance::Transitional,
        }
    }
}

impl PresentationMetadata {
    /// Return the server zoom in DrawingML thousandths of a percent.
    #[inline]
    pub const fn server_zoom(&self) -> i32 {
        self.server_zoom
    }

    /// Return the effective first slide number.
    #[inline]
    pub const fn first_slide_number(&self) -> i32 {
        self.first_slide_number
    }

    /// Whether title slides show their special header and footer placeholders.
    #[inline]
    pub const fn shows_special_placeholders_on_title_slide(&self) -> bool {
        self.show_special_placeholders_on_title_slide
    }

    /// Whether the presentation user interface is right-to-left.
    #[inline]
    pub const fn is_right_to_left(&self) -> bool {
        self.right_to_left
    }

    /// Whether personal information is removed when the document is saved.
    #[inline]
    pub const fn removes_personal_info_on_save(&self) -> bool {
        self.remove_personal_info_on_save
    }

    /// Whether compatibility mode is enabled.
    #[inline]
    pub const fn is_compatibility_mode(&self) -> bool {
        self.compatibility_mode
    }

    /// Whether strict first and last Japanese line characters are used.
    #[inline]
    pub const fn uses_strict_first_and_last_chars(&self) -> bool {
        self.strict_first_and_last_chars
    }

    /// Whether TrueType fonts are embedded automatically.
    #[inline]
    pub const fn embeds_true_type_fonts(&self) -> bool {
        self.embed_true_type_fonts
    }

    /// Whether only used font glyphs are saved for embedded fonts.
    #[inline]
    pub const fn saves_subset_fonts(&self) -> bool {
        self.save_subset_fonts
    }

    /// Whether pictures are compressed automatically.
    #[inline]
    pub const fn automatically_compresses_pictures(&self) -> bool {
        self.auto_compress_pictures
    }

    /// Return the next bookmark identifier seed.
    #[inline]
    pub const fn bookmark_id_seed(&self) -> u32 {
        self.bookmark_id_seed
    }

    /// Return the document conformance class.
    #[inline]
    pub const fn conformance(&self) -> PresentationConformance {
        self.conformance
    }

    fn from_root_element(element: &BytesStart<'_>, decoder: Decoder) -> Result<Self> {
        let mut metadata = Self::default();
        metadata.server_zoom = optional_i32_attribute(
            element,
            b"serverZoom",
            decoder,
            "presentation server zoom",
            metadata.server_zoom,
        )?;
        metadata.first_slide_number = optional_i32_attribute(
            element,
            b"firstSlideNum",
            decoder,
            "first slide number",
            metadata.first_slide_number,
        )?;
        metadata.show_special_placeholders_on_title_slide = optional_boolean_attribute(
            element,
            b"showSpecialPlsOnTitleSld",
            decoder,
            "showSpecialPlsOnTitleSld",
            metadata.show_special_placeholders_on_title_slide,
        )?;
        metadata.right_to_left =
            optional_boolean_attribute(element, b"rtl", decoder, "rtl", metadata.right_to_left)?;
        metadata.remove_personal_info_on_save = optional_boolean_attribute(
            element,
            b"removePersonalInfoOnSave",
            decoder,
            "removePersonalInfoOnSave",
            metadata.remove_personal_info_on_save,
        )?;
        metadata.compatibility_mode = optional_boolean_attribute(
            element,
            b"compatMode",
            decoder,
            "compatMode",
            metadata.compatibility_mode,
        )?;
        metadata.strict_first_and_last_chars = optional_boolean_attribute(
            element,
            b"strictFirstAndLastChars",
            decoder,
            "strictFirstAndLastChars",
            metadata.strict_first_and_last_chars,
        )?;
        metadata.embed_true_type_fonts = optional_boolean_attribute(
            element,
            b"embedTrueTypeFonts",
            decoder,
            "embedTrueTypeFonts",
            metadata.embed_true_type_fonts,
        )?;
        metadata.save_subset_fonts = optional_boolean_attribute(
            element,
            b"saveSubsetFonts",
            decoder,
            "saveSubsetFonts",
            metadata.save_subset_fonts,
        )?;
        metadata.auto_compress_pictures = optional_boolean_attribute(
            element,
            b"autoCompressPictures",
            decoder,
            "autoCompressPictures",
            metadata.auto_compress_pictures,
        )?;
        metadata.bookmark_id_seed =
            optional_bookmark_id_seed(element, decoder, metadata.bookmark_id_seed)?;
        metadata.conformance = optional_conformance(element, decoder, metadata.conformance)?;
        Ok(metadata)
    }
}

/// The declared paragraph-level inventory for a presentation default text style.
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct PresentationDefaultTextStyle {
    has_default_paragraph_properties: bool,
    levels: Vec<u8>,
}

impl PresentationDefaultTextStyle {
    /// Whether the style declares default paragraph properties.
    #[inline]
    pub const fn has_default_paragraph_properties(&self) -> bool {
        self.has_default_paragraph_properties
    }

    /// Return the declared paragraph levels in document order.
    #[inline]
    pub fn levels(&self) -> &[u8] {
        &self.levels
    }

    /// Whether the style declares the requested paragraph level.
    #[inline]
    pub fn has_level(&self, level: u8) -> bool {
        self.levels.contains(&level)
    }
}

/// The main presentation part.
///
/// This part contains the presentation-level properties and references to slides,
/// slide masters, and other presentation resources.
///
/// # Example
///
/// ```rust,ignore
/// let pres_part = PresentationPart::from_part(opc_part)?;
/// let slide_count = pres_part.slide_count()?;
/// ```
pub struct PresentationPart<'a> {
    /// The underlying OPC part
    part: &'a dyn Part,
    xml: Arc<Vec<u8>>,
}

impl<'a> PresentationPart<'a> {
    /// Create a PresentationPart from an OPC Part.
    ///
    /// # Arguments
    ///
    /// * `part` - The underlying OPC part
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let pres_part = PresentationPart::from_part(opc_part)?;
    /// ```
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        let xml = match crate::common::mce::process_ooxml(part.blob())? {
            std::borrow::Cow::Borrowed(_) => part.blob_arc(),
            std::borrow::Cow::Owned(v) => Arc::new(v),
        };
        Ok(Self { part, xml })
    }

    /// Get the XML bytes of the presentation.
    #[inline]
    fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Get the number of slides in the presentation.
    ///
    /// This counts the `<p:sldId>` elements in the presentation.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let count = pres_part.slide_count()?;
    /// println!("Presentation has {} slides", count);
    /// ```
    pub fn slide_count(&self) -> Result<usize> {
        Ok(PresentationInfo::parse(self.xml_bytes())?.slides.len())
    }

    /// Get the slide width in EMUs (English Metric Units).
    ///
    /// Returns None if the slide size is not defined.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// if let Some(width) = pres_part.slide_width()? {
    ///     println!("Slide width: {} EMUs", width);
    /// }
    /// ```
    pub fn slide_width(&self) -> Result<Option<i64>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?
            .slide_size
            .map(|size| size.width()))
    }

    /// Get the slide height in EMUs (English Metric Units).
    ///
    /// Returns None if the slide size is not defined.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// if let Some(height) = pres_part.slide_height()? {
    ///     println!("Slide height: {} EMUs", height);
    /// }
    /// ```
    pub fn slide_height(&self) -> Result<Option<i64>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?
            .slide_size
            .map(|size| size.height()))
    }

    /// Get the presentation slide surface dimensions and declared size type.
    ///
    /// Returns None if the slide size is not defined.
    pub fn slide_size(&self) -> Result<Option<SlideSize>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?.slide_size)
    }

    /// Get the notes and handout surface dimensions.
    ///
    /// Returns None if the notes size is not defined.
    pub fn notes_size(&self) -> Result<Option<NotesSize>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?.notes_size)
    }

    /// Get the root-level presentation behavior and document settings.
    pub fn metadata(&self) -> Result<PresentationMetadata> {
        Ok(PresentationInfo::parse(self.xml_bytes())?.metadata)
    }

    /// Get the presentation-wide default text-style inventory.
    pub fn default_text_style(&self) -> Result<Option<PresentationDefaultTextStyle>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?.default_text_style)
    }

    /// Get the relationship ID of the declared handout master.
    ///
    /// Returns None when the presentation does not declare a handout master.
    pub fn handout_master_relationship_id(&self) -> Result<Option<String>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?.handout_master_relationship_id)
    }

    /// Get the relationship IDs of all slides in presentation order.
    ///
    /// Returns a vector of relationship IDs that can be used to access
    /// the actual slide parts.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let slide_rids = pres_part.slide_rids()?;
    /// for rid in slide_rids {
    ///     // Use rid to get slide part
    /// }
    /// ```
    pub fn slide_rids(&self) -> Result<Vec<String>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?
            .slides
            .into_iter()
            .map(|(_, relationship_id)| relationship_id)
            .collect())
    }

    /// Get the relationship IDs of all slide masters.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let master_rids = pres_part.slide_master_rids()?;
    /// ```
    pub fn slide_master_rids(&self) -> Result<Vec<String>> {
        Ok(PresentationInfo::parse(self.xml_bytes())?
            .masters
            .into_iter()
            .map(|(_, relationship_id)| relationship_id)
            .collect())
    }

    /// Get PowerPoint 2010 slide sections and validate their slide membership.
    ///
    /// Section metadata contains slide IDs, not relationships. Unknown section
    /// extensions remain inert and no targets are opened by this accessor.
    pub fn sections(&self) -> Result<crate::pptx::sections::SectionList> {
        let info = PresentationInfo::parse(self.xml_bytes())?;
        let sections = crate::pptx::sections::SectionList::from_xml(self.xml_bytes())?;
        for section in sections.sections() {
            for slide_id in &section.slide_ids {
                if !info.slides.iter().any(|(id, _)| id == slide_id) {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "PowerPoint section references undeclared slide ID {slide_id}"
                    )));
                }
            }
        }
        Ok(sections)
    }

    /// Get typed PowerPoint 2013 slide and notes guide extensions.
    ///
    /// Guide metadata has no relationships. Unknown guide extensions remain
    /// inert and this accessor never opens package or external targets.
    pub fn extended_guides(
        &self,
    ) -> Result<crate::pptx::extended_guides::PresentationExtendedGuides> {
        crate::pptx::extended_guides::PresentationExtendedGuides::from_xml(self.xml_bytes())
    }

    /// Get the underlying OPC part.
    #[inline]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum PresentationContext {
    Presentation,
    SlideList,
    MasterList,
    HandoutMasterList,
    DefaultTextStyle,
    Other,
}

#[derive(Default)]
struct PresentationInfo {
    slides: Vec<(u32, String)>,
    masters: Vec<(u32, String)>,
    slide_size: Option<SlideSize>,
    notes_size: Option<NotesSize>,
    metadata: PresentationMetadata,
    default_text_style: Option<PresentationDefaultTextStyle>,
    handout_master_relationship_id: Option<String>,
    seen_slide_list: bool,
    seen_master_list: bool,
    seen_handout_master_list: bool,
}

impl PresentationInfo {
    fn parse(xml: &[u8]) -> Result<Self> {
        let mut reader = NsReader::from_reader(xml);
        let mut info = Self::default();
        let mut stack = Vec::new();
        let mut closed_root = false;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) => {
                    if stack.is_empty() {
                        if closed_root
                            || !is_presentationml_name(&namespace, element.name(), b"presentation")
                        {
                            return Err(OoxmlError::InvalidFormat(
                                "presentation XML must have one PresentationML presentation root"
                                    .to_string(),
                            ));
                        }
                        info.metadata = PresentationMetadata::from_root_element(&element, decoder)?;
                        stack.push(PresentationContext::Presentation);
                        continue;
                    }
                    info.process_element(
                        *stack.last().ok_or_else(|| {
                            OoxmlError::InvalidFormat(
                                "missing PowerPoint presentation context".to_string(),
                            )
                        })?,
                        &namespace,
                        &element,
                        decoder,
                        &resolver,
                    )?;
                    stack.push(info.child_context(
                        *stack.last().ok_or_else(|| {
                            OoxmlError::InvalidFormat(
                                "missing PowerPoint presentation context".to_string(),
                            )
                        })?,
                        &namespace,
                        &element,
                    )?);
                },
                Event::Empty(element) => {
                    let parent = *stack.last().ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "presentation XML has an empty or missing root".to_string(),
                        )
                    })?;
                    info.process_element(parent, &namespace, &element, decoder, &resolver)?;
                    info.observe_empty_container(parent, &namespace, &element)?;
                },
                Event::End(element) => {
                    let context = stack.pop().ok_or_else(|| {
                        OoxmlError::InvalidFormat(
                            "invalid PowerPoint presentation nesting".to_string(),
                        )
                    })?;
                    if stack.is_empty() {
                        if context != PresentationContext::Presentation
                            || !is_presentationml_name(&namespace, element.name(), b"presentation")
                        {
                            return Err(OoxmlError::InvalidFormat(
                                "invalid PowerPoint presentation root closure".to_string(),
                            ));
                        }
                        closed_root = true;
                    }
                },
                Event::Eof if !closed_root || !stack.is_empty() => {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated PowerPoint presentation XML".to_string(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(info)
    }

    fn child_context(
        &mut self,
        parent: PresentationContext,
        namespace: &quick_xml::name::ResolveResult<'_>,
        element: &BytesStart<'_>,
    ) -> Result<PresentationContext> {
        if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"sldIdLst")
        {
            if self.seen_slide_list {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint slide ID list".to_string(),
                ));
            }
            self.seen_slide_list = true;
            Ok(PresentationContext::SlideList)
        } else if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"sldMasterIdLst")
        {
            if self.seen_master_list {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint slide-master ID list".to_string(),
                ));
            }
            self.seen_master_list = true;
            Ok(PresentationContext::MasterList)
        } else if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"handoutMasterIdLst")
        {
            if self.seen_handout_master_list {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint handout-master ID list".to_string(),
                ));
            }
            self.seen_handout_master_list = true;
            Ok(PresentationContext::HandoutMasterList)
        } else if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"defaultTextStyle")
        {
            Ok(PresentationContext::DefaultTextStyle)
        } else {
            Ok(PresentationContext::Other)
        }
    }

    fn observe_empty_container(
        &mut self,
        parent: PresentationContext,
        namespace: &quick_xml::name::ResolveResult<'_>,
        element: &BytesStart<'_>,
    ) -> Result<()> {
        if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"sldIdLst")
        {
            if self.seen_slide_list {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint slide ID list".to_string(),
                ));
            }
            self.seen_slide_list = true;
        } else if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"sldMasterIdLst")
        {
            if self.seen_master_list {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint slide-master ID list".to_string(),
                ));
            }
            self.seen_master_list = true;
        } else if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"handoutMasterIdLst")
        {
            if self.seen_handout_master_list {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint handout-master ID list".to_string(),
                ));
            }
            self.seen_handout_master_list = true;
        }
        Ok(())
    }

    fn process_element(
        &mut self,
        parent: PresentationContext,
        namespace: &quick_xml::name::ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &quick_xml::name::NamespaceResolver,
    ) -> Result<()> {
        if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"sldSz")
        {
            if self.slide_size.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint slide size".to_string(),
                ));
            }
            let width = required_positive_i64(element, b"cx", decoder, "slide width")?;
            let height = required_positive_i64(element, b"cy", decoder, "slide height")?;
            let size_type = unqualified_attribute_value(element, b"type", decoder)?;
            self.slide_size = Some(SlideSize {
                width,
                height,
                size_type,
            });
        } else if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"notesSz")
        {
            if self.notes_size.is_some() {
                return Err(OoxmlError::InvalidFormat(
                    "duplicate PowerPoint notes size".to_string(),
                ));
            }
            let width = required_positive_i64(element, b"cx", decoder, "notes width")?;
            let height = required_positive_i64(element, b"cy", decoder, "notes height")?;
            self.notes_size = Some(NotesSize { width, height });
        } else if parent == PresentationContext::Presentation
            && is_presentationml_name(namespace, element.name(), b"defaultTextStyle")
        {
            self.begin_default_text_style()?;
        } else if parent == PresentationContext::DefaultTextStyle {
            self.observe_default_text_style_child(namespace, element)?;
        } else if parent == PresentationContext::SlideList
            && is_presentationml_name(namespace, element.name(), b"sldId")
        {
            let id = required_u32(element, b"id", decoder, "slide ID")?;
            if id < 256 {
                return Err(OoxmlError::InvalidFormat(format!(
                    "PowerPoint slide ID {id} is below 256"
                )));
            }
            let relationship_id = required_relationship_id(element, decoder, resolver, "slide")?;
            push_unique_reference(&mut self.slides, id, relationship_id, "slide")?;
        } else if parent == PresentationContext::MasterList
            && is_presentationml_name(namespace, element.name(), b"sldMasterId")
        {
            let id = required_u32(element, b"id", decoder, "slide-master ID")?;
            if id < 2_147_483_648 {
                return Err(OoxmlError::InvalidFormat(format!(
                    "PowerPoint slide-master ID {id} is below 2147483648"
                )));
            }
            let relationship_id =
                required_relationship_id(element, decoder, resolver, "slide master")?;
            push_unique_reference(&mut self.masters, id, relationship_id, "slide master")?;
        } else if parent == PresentationContext::HandoutMasterList
            && is_presentationml_name(namespace, element.name(), b"handoutMasterId")
        {
            let relationship_id =
                required_relationship_id(element, decoder, resolver, "handout master")?;
            if self
                .handout_master_relationship_id
                .replace(relationship_id)
                .is_some()
            {
                return Err(OoxmlError::InvalidFormat(
                    "PowerPoint presentation has multiple handout-master references".to_string(),
                ));
            }
        }
        Ok(())
    }

    fn begin_default_text_style(&mut self) -> Result<()> {
        if self.default_text_style.is_some() {
            return Err(OoxmlError::InvalidFormat(
                "duplicate PowerPoint default text style".to_string(),
            ));
        }
        self.default_text_style = Some(PresentationDefaultTextStyle::default());
        Ok(())
    }

    fn observe_default_text_style_child(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
    ) -> Result<()> {
        let style = self.default_text_style.as_mut().ok_or_else(|| {
            OoxmlError::InvalidFormat("missing PowerPoint default text style".to_string())
        })?;
        if is_drawingml_name(namespace, element, b"defPPr") {
            if style.has_default_paragraph_properties {
                return Err(OoxmlError::InvalidFormat(
                    "PowerPoint default text style has duplicate default paragraph properties"
                        .to_string(),
                ));
            }
            style.has_default_paragraph_properties = true;
        } else if let Some(level) = drawingml_text_style_level(namespace, element) {
            if style.levels.contains(&level) {
                return Err(OoxmlError::InvalidFormat(format!(
                    "PowerPoint default text style has duplicate level {level}"
                )));
            }
            style.levels.push(level);
        }
        Ok(())
    }
}

fn is_drawingml_name(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> bool {
    if element.name().local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
        },
        ResolveResult::Unknown(prefix) => prefix.as_slice() == b"a",
        ResolveResult::Unbound => false,
    }
}

fn drawingml_text_style_level(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> Option<u8> {
    let local_name = element.name().local_name();
    let local_name = local_name.as_ref();
    if local_name.len() != 7 || !local_name.starts_with(b"lvl") || !local_name.ends_with(b"pPr") {
        return None;
    }
    let level = local_name[3];
    if !(b'1'..=b'9').contains(&level) || !is_drawingml_name(namespace, element, local_name) {
        return None;
    }
    Some(level - b'0')
}

fn optional_i32_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
    default: i32,
) -> Result<i32> {
    let Some(value) = unqualified_attribute_value(element, name, decoder)? else {
        return Ok(default);
    };
    value
        .parse::<i32>()
        .map_err(|_| OoxmlError::InvalidFormat(format!("invalid {description} value '{value}'")))
}

fn optional_boolean_attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
    default: bool,
) -> Result<bool> {
    let Some(value) = unqualified_attribute_value(element, name, decoder)? else {
        return Ok(default);
    };
    match value.as_str() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(OoxmlError::InvalidFormat(format!(
            "invalid presentation {description} value '{value}'"
        ))),
    }
}

fn optional_bookmark_id_seed(
    element: &BytesStart<'_>,
    decoder: Decoder,
    default: u32,
) -> Result<u32> {
    let Some(value) = unqualified_attribute_value(element, b"bookmarkIdSeed", decoder)? else {
        return Ok(default);
    };
    let seed = value.parse::<u32>().map_err(|_| {
        OoxmlError::InvalidFormat(format!("invalid bookmark ID seed value '{value}'"))
    })?;
    if !(1..2_147_483_648).contains(&seed) {
        return Err(OoxmlError::InvalidFormat(format!(
            "bookmark ID seed {seed} is outside the valid range"
        )));
    }
    Ok(seed)
}

fn optional_conformance(
    element: &BytesStart<'_>,
    decoder: Decoder,
    default: PresentationConformance,
) -> Result<PresentationConformance> {
    let Some(value) = unqualified_attribute_value(element, b"conformance", decoder)? else {
        return Ok(default);
    };
    match value.as_str() {
        "transitional" => Ok(PresentationConformance::Transitional),
        "strict" => Ok(PresentationConformance::Strict),
        _ => Err(OoxmlError::InvalidFormat(format!(
            "invalid presentation conformance value '{value}'"
        ))),
    }
}

fn required_relationship_id(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
    description: &str,
) -> Result<String> {
    let value =
        relationship_attribute_value(element, b"id", decoder, resolver)?.ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("PowerPoint {description} is missing r:id"))
        })?;
    if value.is_empty() {
        return Err(OoxmlError::InvalidFormat(format!(
            "PowerPoint {description} has an empty relationship ID"
        )));
    }
    Ok(value)
}

fn required_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<u32> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| OoxmlError::InvalidFormat(format!("missing {description} attribute")))?;
    value
        .parse::<u32>()
        .map_err(|_| OoxmlError::InvalidFormat(format!("invalid {description} value '{value}'")))
}

fn required_positive_i64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<i64> {
    let value = unqualified_attribute_value(element, name, decoder)?
        .ok_or_else(|| OoxmlError::InvalidFormat(format!("missing {description} attribute")))?;
    let parsed = value
        .parse::<i64>()
        .map_err(|_| OoxmlError::InvalidFormat(format!("invalid {description} value '{value}'")))?;
    if parsed <= 0 {
        return Err(OoxmlError::InvalidFormat(format!(
            "{description} must be positive"
        )));
    }
    Ok(parsed)
}

fn push_unique_reference(
    references: &mut Vec<(u32, String)>,
    id: u32,
    relationship_id: String,
    description: &str,
) -> Result<()> {
    if references
        .iter()
        .any(|(existing_id, existing_rid)| *existing_id == id || *existing_rid == relationship_id)
    {
        return Err(OoxmlError::InvalidFormat(format!(
            "duplicate PowerPoint {description} ID or relationship"
        )));
    }
    references.push((id, relationship_id));
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn part(xml: impl Into<Vec<u8>>) -> BlobPart {
        BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .to_string(),
            xml.into(),
        )
    }

    #[test]
    fn parses_ordered_references_and_dimensions_by_namespace() {
        let xml = format!(
            r#"<q:presentation xmlns:q="{P}" xmlns:rel="{R}" xmlns:f="urn:foreign">
                <q:sldMasterIdLst><f:sldMasterId id="4294967295" rel:id="spoof"/>
                    <q:sldMasterId id="2147483648" f:id="7" rel:id="master-alpha"/></q:sldMasterIdLst>
                <q:sldIdLst><q:sldId id="256" f:id="1" rel:id="slide-alpha"/>
                    <q:sldId id="257" rel:id="slide-beta"/><f:sldId id="258" rel:id="spoof"/></q:sldIdLst>
                <f:sldSz cx="1" cy="1"/><q:sldSz cx="9144000" cy="5143500" type="screen16x9"/>
                <f:notesSz cx="1" cy="1"/><q:notesSz cx="6858000" cy="9144000"/>
            </q:presentation>"#
        );
        let blob = part(xml);
        let presentation = PresentationPart::from_part(&blob).unwrap();
        assert_eq!(presentation.slide_count().unwrap(), 2);
        assert_eq!(
            presentation.slide_rids().unwrap(),
            ["slide-alpha", "slide-beta"]
        );
        assert_eq!(presentation.slide_master_rids().unwrap(), ["master-alpha"]);
        assert_eq!(presentation.slide_width().unwrap(), Some(9_144_000));
        assert_eq!(presentation.slide_height().unwrap(), Some(5_143_500));
        let slide_size = presentation.slide_size().unwrap().unwrap();
        assert_eq!(slide_size.width(), 9_144_000);
        assert_eq!(slide_size.height(), 5_143_500);
        assert_eq!(slide_size.size_type(), Some("screen16x9"));
        let notes_size = presentation.notes_size().unwrap().unwrap();
        assert_eq!(notes_size.width(), 6_858_000);
        assert_eq!(notes_size.height(), 9_144_000);
    }

    #[test]
    fn accepts_strict_relationship_aliases() {
        let xml = r#"<x:presentation xmlns:x="http://purl.oclc.org/ooxml/presentationml/main"
            xmlns:z="http://purl.oclc.org/ooxml/officeDocument/relationships">
            <x:sldMasterIdLst><x:sldMasterId id="2147483648" z:id="m"/></x:sldMasterIdLst>
            <x:sldIdLst><x:sldId id="256" z:id="s"/></x:sldIdLst>
            <x:sldSz cx="1" cy="2"/><x:notesSz cx="3" cy="4"/></x:presentation>"#;
        let blob = part(xml);
        let presentation = PresentationPart::from_part(&blob).unwrap();
        assert_eq!(presentation.slide_rids().unwrap(), ["s"]);
        assert_eq!(presentation.slide_master_rids().unwrap(), ["m"]);
        assert_eq!(
            presentation.slide_size().unwrap().unwrap().size_type(),
            None
        );
        let notes_size = presentation.notes_size().unwrap().unwrap();
        assert_eq!(notes_size.width(), 3);
        assert_eq!(notes_size.height(), 4);
    }

    #[test]
    fn parses_handout_master_reference_by_namespace() {
        let xml = format!(
            r#"<q:presentation xmlns:q="{P}" xmlns:rel="{R}" xmlns:f="urn:foreign">
                <f:handoutMasterIdLst><f:handoutMasterId rel:id="spoof"/></f:handoutMasterIdLst>
                <q:handoutMasterIdLst><f:handoutMasterId rel:id="spoof"/>
                    <q:handoutMasterId f:id="wrong" rel:id="handout-alpha">
                        <q:extLst><q:handoutMasterId rel:id="nested"/></q:extLst>
                    </q:handoutMasterId>
                </q:handoutMasterIdLst>
                <q:extLst><q:handoutMasterIdLst>
                    <q:handoutMasterId rel:id="extension"/>
                </q:handoutMasterIdLst></q:extLst>
            </q:presentation>"#
        );
        let blob = part(xml);
        let presentation = PresentationPart::from_part(&blob).unwrap();
        assert_eq!(
            presentation.handout_master_relationship_id().unwrap(),
            Some("handout-alpha".to_string())
        );

        let strict = r#"<x:presentation xmlns:x="http://purl.oclc.org/ooxml/presentationml/main"
            xmlns:z="http://purl.oclc.org/ooxml/officeDocument/relationships">
            <x:handoutMasterIdLst><x:handoutMasterId z:id="strict-handout"/>
            </x:handoutMasterIdLst></x:presentation>"#;
        let blob = part(strict);
        assert_eq!(
            PresentationPart::from_part(&blob)
                .unwrap()
                .handout_master_relationship_id()
                .unwrap(),
            Some("strict-handout".to_string())
        );

        let absent = format!(r#"<p:presentation xmlns:p="{P}"></p:presentation>"#);
        let blob = part(absent);
        assert_eq!(
            PresentationPart::from_part(&blob)
                .unwrap()
                .handout_master_relationship_id()
                .unwrap(),
            None
        );
    }

    #[test]
    fn ignores_nested_and_foreign_reference_lookalikes() {
        let xml = format!(
            r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}" xmlns:f="urn:foreign">
                <p:sldIdLst><f:wrapper><p:sldId id="256" r:id="nested"/></f:wrapper>
                    <p:sldId id="257" r:id="real"/></p:sldIdLst>
                <p:extLst><p:sldIdLst><p:sldId id="258" r:id="extension"/></p:sldIdLst></p:extLst>
            </p:presentation>"#
        );
        let blob = part(xml);
        let presentation = PresentationPart::from_part(&blob).unwrap();
        assert_eq!(presentation.slide_rids().unwrap(), ["real"]);
    }

    #[test]
    fn parses_presentation_root_metadata_and_defaults() {
        let xml = format!(
            r#"<p:presentation xmlns:p="{P}" serverZoom="87500" firstSlideNum="-3"
                showSpecialPlsOnTitleSld="0" rtl="true" removePersonalInfoOnSave="1"
                compatMode="false" strictFirstAndLastChars="0" embedTrueTypeFonts="true"
                saveSubsetFonts="1" autoCompressPictures="false" bookmarkIdSeed="2147483647"
                conformance="strict"></p:presentation>"#
        );
        let blob = part(xml);
        let metadata = PresentationPart::from_part(&blob).unwrap().metadata().unwrap();
        assert_eq!(metadata.server_zoom(), 87_500);
        assert_eq!(metadata.first_slide_number(), -3);
        assert!(!metadata.shows_special_placeholders_on_title_slide());
        assert!(metadata.is_right_to_left());
        assert!(metadata.removes_personal_info_on_save());
        assert!(!metadata.is_compatibility_mode());
        assert!(!metadata.uses_strict_first_and_last_chars());
        assert!(metadata.embeds_true_type_fonts());
        assert!(metadata.saves_subset_fonts());
        assert!(!metadata.automatically_compresses_pictures());
        assert_eq!(metadata.bookmark_id_seed(), 2_147_483_647);
        assert_eq!(metadata.conformance(), PresentationConformance::Strict);

        let defaults = format!(r#"<p:presentation xmlns:p="{P}"></p:presentation>"#);
        let blob = part(defaults);
        let metadata = PresentationPart::from_part(&blob).unwrap().metadata().unwrap();
        assert_eq!(metadata, PresentationMetadata::default());
        assert_eq!(
            metadata.conformance(),
            PresentationConformance::Transitional
        );
    }

    #[test]
    fn presentation_root_metadata_rejects_invalid_attributes() {
        let cases = [
            format!(r#"<p:presentation xmlns:p="{P}" serverZoom="invalid"></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{P}" firstSlideNum="2147483648"></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{P}" rtl="sometimes"></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{P}" bookmarkIdSeed="0"></p:presentation>"#),
            format!(
                r#"<p:presentation xmlns:p="{P}" bookmarkIdSeed="2147483648"></p:presentation>"#
            ),
            format!(
                r#"<p:presentation xmlns:p="{P}" conformance="future"></p:presentation>"#
            ),
            format!(r#"<p:presentation xmlns:p="{P}" rtl="1" rtl="0"></p:presentation>"#),
        ];
        for xml in cases {
            let blob = part(xml);
            assert!(PresentationPart::from_part(&blob).unwrap().metadata().is_err());
        }
    }

    #[test]
    fn parses_default_text_style_inventory() {
        let xml = format!(
            r#"<p:presentation xmlns:p="{P}" xmlns:a="{A}" xmlns:f="urn:foreign">
                <p:defaultTextStyle><a:defPPr/><a:lvl2pPr/><a:lvl8pPr/>
                    <f:lvl3pPr/><a:extLst><a:lvl4pPr/></a:extLst>
                </p:defaultTextStyle></p:presentation>"#
        );
        let blob = part(xml);
        let style = PresentationPart::from_part(&blob)
            .unwrap()
            .default_text_style()
            .unwrap()
            .unwrap();
        assert!(style.has_default_paragraph_properties());
        assert_eq!(style.levels(), [2, 8]);
        assert!(style.has_level(8));
        assert!(!style.has_level(4));

        let strict = r#"<q:presentation xmlns:q="http://purl.oclc.org/ooxml/presentationml/main"
            xmlns:d="http://purl.oclc.org/ooxml/drawingml/main"><q:defaultTextStyle>
            <d:lvl1pPr/></q:defaultTextStyle></q:presentation>"#;
        let blob = part(strict);
        let style = PresentationPart::from_part(&blob)
            .unwrap()
            .default_text_style()
            .unwrap()
            .unwrap();
        assert!(!style.has_default_paragraph_properties());
        assert_eq!(style.levels(), [1]);

        let absent = format!(r#"<p:presentation xmlns:p="{P}"></p:presentation>"#);
        let blob = part(absent);
        assert_eq!(
            PresentationPart::from_part(&blob)
                .unwrap()
                .default_text_style()
                .unwrap(),
            None
        );
    }

    #[test]
    fn rejects_duplicate_default_text_style_declarations() {
        let cases = [
            format!(
                r#"<p:presentation xmlns:p="{P}"><p:defaultTextStyle/><p:defaultTextStyle/></p:presentation>"#
            ),
            format!(
                r#"<p:presentation xmlns:p="{P}" xmlns:a="{A}"><p:defaultTextStyle><a:defPPr/><a:defPPr/></p:defaultTextStyle></p:presentation>"#
            ),
            format!(
                r#"<p:presentation xmlns:p="{P}" xmlns:a="{A}"><p:defaultTextStyle><a:lvl3pPr/><a:lvl3pPr/></p:defaultTextStyle></p:presentation>"#
            ),
        ];
        for xml in cases {
            let blob = part(xml);
            assert!(
                PresentationPart::from_part(&blob)
                    .unwrap()
                    .default_text_style()
                    .is_err()
            );
        }
    }

    #[test]
    fn rejects_malformed_handout_master_references() {
        let cases = [
            format!(
                r#"<p:presentation xmlns:p="{P}"><p:handoutMasterIdLst><p:handoutMasterId/></p:handoutMasterIdLst></p:presentation>"#
            ),
            format!(
                r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:handoutMasterIdLst><p:handoutMasterId r:id=""/></p:handoutMasterIdLst></p:presentation>"#
            ),
            format!(
                r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:handoutMasterIdLst><p:handoutMasterId r:id="one"/><p:handoutMasterId r:id="two"/></p:handoutMasterIdLst></p:presentation>"#
            ),
            format!(
                r#"<p:presentation xmlns:p="{P}"><p:handoutMasterIdLst/><p:handoutMasterIdLst/></p:presentation>"#
            ),
        ];
        for xml in cases {
            let blob = part(xml);
            assert!(
                PresentationPart::from_part(&blob)
                    .unwrap()
                    .handout_master_relationship_id()
                    .is_err()
            );
        }
    }

    #[test]
    fn rejects_malformed_presentation_metadata() {
        let invalid = [
            format!(
                r#"<p:presentation xmlns:p="{P}"><p:sldIdLst><p:sldId id="255"/></p:sldIdLst></p:presentation>"#
            ),
            format!(
                r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id=""/></p:sldIdLst></p:presentation>"#
            ),
            format!(
                r#"<p:presentation xmlns:p="{P}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id="a"/><p:sldId id="256" r:id="b"/></p:sldIdLst></p:presentation>"#
            ),
            format!(r#"<p:presentation xmlns:p="{P}"><p:sldSz cx="0" cy="1"/></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{P}"><p:sldSz cx="1"/></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{P}"><p:notesSz cx="0" cy="1"/></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{P}"><p:notesSz cx="1"/></p:presentation>"#),
            format!(
                r#"<p:presentation xmlns:p="{P}"><p:notesSz cx="1" cy="2"/><p:notesSz cx="3" cy="4"/></p:presentation>"#
            ),
            format!(r#"<p:presentation xmlns:p="{P}"><p:sldIdLst/><p:sldIdLst/></p:presentation>"#),
            format!(r#"<p:presentation xmlns:p="{P}"><p:sldSz cx="1" cy="2"/>"#),
        ];
        for xml in invalid {
            let blob = part(xml);
            assert!(
                PresentationPart::from_part(&blob)
                    .unwrap()
                    .slide_count()
                    .is_err()
            );
        }
    }
}
