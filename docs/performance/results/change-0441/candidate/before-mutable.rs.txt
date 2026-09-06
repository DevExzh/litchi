//! Private staging engine for source-checked presentation transactions.
//!
//! The attached mutable root is intentionally not exported. Public mutation is
//! available only through [`super::edit::Transaction`].

use crate::codec::content_source::ContentSource;
use crate::core::{OwnedPackage, PackageWriter, Structure};
use crate::model::animation::validate_animation_roots;
use crate::model::legacy_animation::validate_legacy_animation_root;
use crate::model::media::{EmbeddedMedia, embed_media, validate_package_media_path};
use crate::{Presentation, Reference, Shape, Slide};
use litchi_core::{Result, xml::escape_xml};
use litchi_odf_common::package::{is_linked_href, resolve_package_path};
use quick_xml::{
    Decoder, Reader, XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::collections::{BTreeMap, BTreeSet, HashSet};

const MAX_DEPENDENCY_FREE_COPY_PAGE_BYTES: usize = 64 * 1024;
const MAX_DEPENDENCY_FREE_COPY_NAME_BYTES: usize = 4 * 1024;
const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const STYLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const PRESENTATION_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const SVG_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
const ANIMATION_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:animation:1.0";
const SMIL_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";
const DR3D_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const SCRIPT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_DEFAULT: &[u8] = b"xmlns";
const XMLNS_PREFIX: &[u8] = b"xmlns:";
const MAX_NAMESPACE_BINDINGS: usize = 512;
const MAX_MEDIA_CHANGE_PAYLOAD_BYTES: usize = 32 * 1024 * 1024;
const MAX_MEDIA_REFERENCE_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_MEDIA_REFERENCE_TOTAL_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_MEDIA_REFERENCE_XML_DEPTH: usize = 128;
const MAX_MEDIA_REFERENCE_XML_EVENTS: usize = 1_000_000;
const MAX_MEDIA_REFERENCE_TOTAL_XML_EVENTS: usize = 4_000_000;
const MAX_MEDIA_REFERENCE_XML_PARTS: usize = 65_536;
const MAX_MEDIA_REFERENCE_VALUE_BYTES: usize = 1_048_576;
const MAX_MEDIA_REFERENCE_PATHS: usize = 131_072;
const MAX_MEDIA_REFERENCE_PATH_BYTES: usize = 16 * 1024 * 1024;

pub(super) struct MediaChangePlan {
    pub(super) media_files: BTreeMap<String, EmbeddedMedia>,
    pub(super) removed_media_paths: BTreeSet<String>,
    pub(super) changed: usize,
    pub(super) staged_media_bytes: usize,
}

#[derive(Default)]
struct MediaReferenceIndex {
    referenced_paths: HashSet<String>,
}

pub(super) struct DependencyFreeBlankSlideCopy {
    slide: Slide,
    page: String,
    page_metadata: crate::model::page_metadata::Collection,
    source_origin: usize,
}

pub(super) struct ForeignDependencyFreeBlankSlideCopy {
    slide: Slide,
    page: String,
    page_metadata: crate::model::page_metadata::Collection,
}

impl ForeignDependencyFreeBlankSlideCopy {
    pub(super) fn resource_bytes(&self) -> usize {
        self.page.len()
    }
}

pub(super) struct DependencyFreeBlankSlideRemoval {
    index: usize,
    page_metadata: Option<crate::model::page_metadata::Collection>,
}

impl DependencyFreeBlankSlideCopy {
    pub(super) fn resource_bytes(&self) -> usize {
        self.page.len()
    }
}

/// A mutable ODP presentation that supports in-place modifications.
///
/// This struct wraps an ODP presentation and provides methods to modify its content,
/// including adding, updating, and removing slides and shapes.
///
/// # Examples
///
/// ```
/// use litchi_odp::{Builder, edit};
///
/// # fn main() -> litchi_core::Result<()> {
/// // Build a presentation and take an immutable editing snapshot
/// let mut builder = Builder::new();
/// builder.add_slide_with_title("Welcome", "First draft")?;
/// let source = edit::Snapshot::from_bytes(builder.build()?)?;
///
/// // Stage edits in an isolated transaction
/// let mut transaction = source.transaction()?;
/// transaction.add("New Slide", "Slide content")?;
/// transaction.remove(0)?;
///
/// // Publish atomically; the source snapshot is never mutated
/// let commit = transaction.commit()?;
/// assert_eq!(commit.snapshot().slides().len(), 1);
/// # Ok(())
/// # }
/// ```
pub(super) struct MutablePresentation {
    /// Mutable slides
    slides: Vec<Slide>,
    /// Original MIME type
    mimetype: String,
    /// Original styles XML (preserved as-is)
    styles_xml: Option<String>,
    /// Original package retained for copying auxiliary package parts.
    source_package: Option<OwnedPackage>,
    /// Newly embedded package media, keyed by package path.
    media_files: BTreeMap<String, EmbeddedMedia>,
    /// Existing source media members explicitly removed by the transaction.
    removed_media_paths: BTreeSet<String>,
    /// Bounded source XML reference index, built once when a media removal is staged.
    media_reference_index: Option<MediaReferenceIndex>,
    /// Inert slide-show settings and custom shows.
    settings: Option<crate::Settings>,
    /// Inert header/footer/date-time declarations and page bindings.
    declarations: Option<crate::model::declaration::Collection>,
    /// Static page names, IDs, and layout/master references.
    page_metadata: Option<crate::model::page_metadata::Collection>,
    /// Verbatim fragments of the source `content.xml`, when one was opened.
    ///
    /// Slides the caller never touched are re-emitted from these fragments so
    /// constructs outside the model — nested tables, image text alternatives,
    /// automatic styles, font declarations — survive a save.
    content_source: Option<ContentSource>,
    /// Slides exactly as parsed, used to detect which pages are still pristine.
    source_slides: Vec<Slide>,
    /// Exact source-page identity for every staged slide; inserted slides have no source origin.
    origins: Vec<Option<usize>>,
    /// Owned exact page fragments for dependency-free blank-slide copies.
    ///
    /// Ordinary source pages continue to borrow their fragment through
    /// `origins`; a copied page owns only the minimally renamed duplicate.
    page_overrides: Vec<Option<String>>,
    /// A foreign copied page has no local source origin. Its owned fragment is
    /// retained only while the imported semantic slide remains untouched.
    foreign_page_overrides: Vec<bool>,
    /// Declaration state exactly as parsed; changing it rewrites every page.
    source_declarations: Option<crate::model::declaration::Collection>,
    /// Lazy result of the current writer's whole-content publication audit.
    /// Repeated reorders in one transaction remain O(1) after the first scan.
    slide_move_supported: Option<bool>,
}

impl MutablePresentation {
    /// Create a mutable presentation from an existing presentation and the
    /// slide projection already validated for the same immutable package.
    ///
    /// Package-owned auxiliary structures are still parsed independently;
    /// only the duplicate complete slide traversal is skipped.
    ///
    /// # Arguments
    ///
    /// * `presentation` - The presentation to make mutable
    /// * `validated_slides` - Slides parsed from the same source package
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odp::{Builder, Presentation};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let presentation = Presentation::from_bytes(Builder::new().build()?)?;
    /// let snapshot = presentation.snapshot()?;
    /// let transaction = snapshot.transaction()?;
    /// assert!(transaction.slides().is_empty());
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn from_presentation_with_validated_slides(
        presentation: &Presentation,
        validated_slides: &[Slide],
    ) -> Result<Self> {
        // Editing snapshots construct this projection from the same immutable
        // package bytes before a transaction can be opened. Keep staging
        // isolated by cloning it into the draft, but do not parse every slide
        // a second time solely to obtain the same semantic values.
        let slides = validated_slides.to_vec();
        let settings = presentation.settings()?;
        let parsed_declarations = presentation.declarations()?;
        let parsed_page_metadata = presentation.pages()?;
        let mimetype = presentation.owned_package().mimetype()?;

        let styles_xml = presentation.styles_xml().map(str::to_owned);
        let content_source = ContentSource::parse(presentation.content_xml())?;
        let source_package = Some(presentation.owned_package().clone());

        let declarations = (!parsed_declarations.is_empty()).then_some(parsed_declarations);
        let page_metadata = (!parsed_page_metadata.is_empty()).then_some(parsed_page_metadata);
        let source_slides = match &content_source {
            // Only pages the scanner actually recovered can be retained; a
            // mismatch means the model and the markup disagree, so fall back to
            // regenerating every page.
            Some(source) if source.page_count() == slides.len() => slides.clone(),
            _ if slides.is_empty() => Vec::new(),
            _ => {
                return Err(litchi_core::Error::Unsupported(
                    "ODP source slide coverage is incomplete; transaction staging refuses regeneration"
                        .to_string(),
                ));
            },
        };
        let retained_source = content_source;
        let origins = if retained_source.is_some() {
            (0..slides.len()).map(Some).collect()
        } else {
            vec![None; slides.len()]
        };
        let page_overrides = (0..slides.len()).map(|_| None).collect();
        let foreign_page_overrides = (0..slides.len()).map(|_| false).collect();

        Ok(Self {
            slides,
            mimetype,
            styles_xml,
            source_package,
            media_files: BTreeMap::new(),
            removed_media_paths: BTreeSet::new(),
            media_reference_index: None,
            settings,
            source_declarations: declarations.clone(),
            declarations,
            page_metadata,
            content_source: retained_source,
            source_slides,
            origins,
            page_overrides,
            foreign_page_overrides,
            slide_move_supported: None,
        })
    }

    /// Get all slides in the presentation.
    pub(super) fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Return whether rewriting this slide would discard retained source XML.
    pub(super) fn retains_source_slide(&self, index: usize) -> bool {
        if self
            .foreign_page_overrides
            .get(index)
            .copied()
            .unwrap_or(false)
        {
            return self
                .page_overrides
                .get(index)
                .and_then(Option::as_ref)
                .is_some();
        }
        self.slides
            .get(index)
            .is_some_and(|slide| self.retained_page(index, slide).is_some())
    }

    /// Return whether structural edits would need declaration rebinding.
    pub(super) fn has_source_declarations(&self) -> bool {
        self.source_declarations.is_some()
    }

    /// Check whether the current package writer can publish a reordered copy
    /// of every retained source page without reclassifying producer XML as
    /// authored markup.
    pub(super) fn check_slide_move_supported(&mut self) -> Result<()> {
        if self.source_declarations.is_some() || self.settings.is_some() {
            return Err(litchi_core::Error::Unsupported(
                "ODP retained producer declarations or settings cannot yet be reordered losslessly"
                    .to_string(),
            ));
        }
        let supported = match self.slide_move_supported {
            Some(supported) => supported,
            None => {
                let supported = if let Some(package) = &self.source_package {
                    let content = package.get_file("content.xml")?;
                    xml_minifier::audit::verify_authored(
                        &content,
                        xml_minifier::audit::Limits::default(),
                    )
                    .is_ok()
                } else {
                    true
                };
                self.slide_move_supported = Some(supported);
                supported
            },
        };
        if !supported {
            return Err(litchi_core::Error::Unsupported(
                "ODP retained producer pages cannot be reordered losslessly by the current writer"
                    .to_string(),
            ));
        }
        Ok(())
    }

    /// Prepare an append-only exact-fragment copy of a dependency-free blank page.
    ///
    /// The admitted XML shape is deliberately tiny: one self-closing
    /// `draw:page` whose only non-namespace attribute is `draw:name`. This
    /// excludes every style, master, layout, ID/navigation, hyperlink, event,
    /// protection, MCE, script, child-content, and opaque dependency surface.
    pub(super) fn prepare_dependency_free_blank_slide_copy(
        &mut self,
        index: usize,
    ) -> Result<DependencyFreeBlankSlideCopy> {
        self.check_slide_move_supported()?;
        let source_origin = self.origins.get(index).copied().flatten().ok_or_else(|| {
            litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide copy requires an exact retained source page"
                    .to_string(),
            )
        })?;
        if self
            .page_overrides
            .get(index)
            .and_then(Option::as_ref)
            .is_some()
        {
            return Err(litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide copy cannot use an already copied page"
                    .to_string(),
            ));
        }
        let source = self.content_source.as_ref().ok_or_else(|| {
            litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide copy requires retained content.xml fragments"
                    .to_string(),
            )
        })?;
        let page = source.page(source_origin).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "ODP retained page origin is outside content.xml coverage".to_string(),
            )
        })?;
        if page.len() > MAX_DEPENDENCY_FREE_COPY_PAGE_BYTES {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP dependency-free blank page exceeds {MAX_DEPENDENCY_FREE_COPY_PAGE_BYTES} bytes"
            )));
        }
        let name_value = dependency_free_blank_name_value(page)?;
        let metadata = self
            .page_metadata
            .as_ref()
            .and_then(|value| value.page(index))
            .ok_or_else(|| {
                litchi_core::Error::Unsupported(
                    "ODP dependency-free blank-slide copy requires parsed draw:name metadata"
                        .to_string(),
                )
            })?;
        let old_name = metadata.name.as_deref().ok_or_else(|| {
            litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide copy requires draw:name".to_string(),
            )
        })?;
        if old_name.len() > MAX_DEPENDENCY_FREE_COPY_NAME_BYTES {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP dependency-free blank page name exceeds {MAX_DEPENDENCY_FREE_COPY_NAME_BYTES} bytes"
            )));
        }
        let names = crate::model::page_metadata::effective_page_names(
            self.page_metadata.as_ref(),
            self.slides.len(),
        )?;
        let new_name = dependency_free_copy_name(old_name, &names)?;
        let escaped_name = escape_xml(&new_name);
        let new_page_len = page
            .len()
            .checked_sub(name_value.end - name_value.start)
            .and_then(|value| value.checked_add(escaped_name.len()))
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(
                    "ODP dependency-free blank page size overflow".to_string(),
                )
            })?;
        if new_page_len > MAX_DEPENDENCY_FREE_COPY_PAGE_BYTES {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP dependency-free copied page exceeds {MAX_DEPENDENCY_FREE_COPY_PAGE_BYTES} bytes"
            )));
        }
        let mut copied_page = String::new();
        copied_page
            .try_reserve_exact(new_page_len)
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODP dependency-free blank page fragment",
                source,
            })?;
        copied_page.push_str(&page[..name_value.start]);
        copied_page.push_str(&escaped_name);
        copied_page.push_str(&page[name_value.end..]);

        let page_metadata = crate::model::page_metadata::metadata_after_dependency_free_page_copy(
            self.page_metadata.as_ref(),
            self.slides.len(),
            index,
            new_name,
        )?;
        let mut slide = self.slides.get(index).cloned().ok_or_else(|| {
            litchi_core::Error::InvalidFormat(format!(
                "ODP dependency-free blank-slide copy index {index} is out of bounds"
            ))
        })?;
        slide.index = self.slides.len();
        Ok(DependencyFreeBlankSlideCopy {
            slide,
            page: copied_page,
            page_metadata,
            source_origin,
        })
    }

    /// Atomically append a prevalidated dependency-free blank-page copy.
    pub(super) fn apply_dependency_free_blank_slide_copy(
        &mut self,
        copy: DependencyFreeBlankSlideCopy,
    ) -> Result<usize> {
        self.slides
            .try_reserve_exact(1)
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODP copied slide projection",
                source,
            })?;
        self.origins
            .try_reserve_exact(1)
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODP copied slide origins",
                source,
            })?;
        self.page_overrides.try_reserve_exact(1).map_err(|source| {
            litchi_core::Error::Allocation {
                resource: "ODP copied slide fragments",
                source,
            }
        })?;
        self.foreign_page_overrides
            .try_reserve_exact(1)
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODP copied slide provenance",
                source,
            })?;
        let index = self.slides.len();
        self.slides.push(copy.slide);
        self.origins.push(Some(copy.source_origin));
        self.page_overrides.push(Some(copy.page));
        self.foreign_page_overrides.push(false);
        self.page_metadata = Some(copy.page_metadata);
        Ok(index)
    }

    /// Prepare an append-only dependency-free page imported from another
    /// presentation. The donor fragment is admitted only when it is a
    /// canonical self-closing `draw:page` with a single semantic attribute,
    /// `draw:name`; all destination page identity is rebuilt locally.
    pub(super) fn prepare_foreign_dependency_free_blank_slide_copy(
        &mut self,
        source_page: &str,
        source_slide: &Slide,
        source_name: &str,
    ) -> Result<ForeignDependencyFreeBlankSlideCopy> {
        self.check_slide_move_supported()?;
        self.check_foreign_transfer_destination_safety()?;
        if self.content_source.is_none() {
            return Err(litchi_core::Error::Unsupported(
                "ODP foreign slide transfer requires a retained destination presentation body"
                    .to_string(),
            ));
        }
        if source_page.len() > MAX_DEPENDENCY_FREE_COPY_PAGE_BYTES {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP foreign blank page exceeds {MAX_DEPENDENCY_FREE_COPY_PAGE_BYTES} bytes"
            )));
        }
        let name_value = dependency_free_blank_name_value(source_page)?;
        if source_name.is_empty() {
            return Err(litchi_core::Error::Unsupported(
                "ODP foreign blank page requires a non-empty draw:name".to_string(),
            ));
        }
        if source_name.len() > MAX_DEPENDENCY_FREE_COPY_NAME_BYTES {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP foreign blank page name exceeds {MAX_DEPENDENCY_FREE_COPY_NAME_BYTES} bytes"
            )));
        }
        if source_slide.title.is_some()
            || !source_slide.text.is_empty()
            || source_slide.notes.is_some()
            || source_slide.transition.is_some()
            || !source_slide.animations.is_empty()
            || source_slide.legacy_animation.is_some()
            || !source_slide.shapes.is_empty()
        {
            return Err(litchi_core::Error::Unsupported(
                "ODP foreign slide transfer requires a dependency-free blank page".to_string(),
            ));
        }
        let names = crate::model::page_metadata::effective_page_names(
            self.page_metadata.as_ref(),
            self.slides.len(),
        )?;
        let new_name = foreign_copy_name(source_name, &names)?;
        let escaped_name = escape_xml(&new_name);
        let new_page_len = source_page
            .len()
            .checked_sub(name_value.end - name_value.start)
            .and_then(|value| value.checked_add(escaped_name.len()))
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(
                    "ODP foreign blank page size overflow".to_string(),
                )
            })?;
        if new_page_len > MAX_DEPENDENCY_FREE_COPY_PAGE_BYTES {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP foreign copied page exceeds {MAX_DEPENDENCY_FREE_COPY_PAGE_BYTES} bytes"
            )));
        }
        let mut page = String::new();
        page.try_reserve_exact(new_page_len)
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODP foreign blank page fragment",
                source,
            })?;
        page.push_str(&source_page[..name_value.start]);
        page.push_str(&escaped_name);
        page.push_str(&source_page[name_value.end..]);

        let page_metadata =
            crate::model::page_metadata::metadata_after_foreign_dependency_free_page_copy(
                self.page_metadata.as_ref(),
                self.slides.len(),
                new_name,
            )?;
        let slide = Slide {
            title: None,
            text: String::new(),
            index: self.slides.len(),
            notes: None,
            transition: None,
            animations: Vec::new(),
            legacy_animation: None,
            shapes: Vec::new(),
        };
        Ok(ForeignDependencyFreeBlankSlideCopy {
            slide,
            page,
            page_metadata,
        })
    }

    /// Atomically append a prevalidated dependency-free page imported from a
    /// different presentation. The page carries no local source origin and is
    /// retained through its explicit foreign provenance marker.
    pub(super) fn apply_foreign_dependency_free_blank_slide_copy(
        &mut self,
        copy: ForeignDependencyFreeBlankSlideCopy,
    ) -> Result<usize> {
        self.slides
            .try_reserve_exact(1)
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODP foreign copied slide projection",
                source,
            })?;
        self.origins
            .try_reserve_exact(1)
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODP foreign copied slide origins",
                source,
            })?;
        self.page_overrides.try_reserve_exact(1).map_err(|source| {
            litchi_core::Error::Allocation {
                resource: "ODP foreign copied slide fragments",
                source,
            }
        })?;
        self.foreign_page_overrides
            .try_reserve_exact(1)
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODP foreign copied slide provenance",
                source,
            })?;
        let index = self.slides.len();
        self.slides.push(copy.slide);
        self.origins.push(None);
        self.page_overrides.push(Some(copy.page));
        self.foreign_page_overrides.push(true);
        self.page_metadata = Some(copy.page_metadata);
        Ok(index)
    }

    /// Prepare exact removal of one retained dependency-free blank page.
    ///
    /// This admits the same deliberately tiny page grammar as exact copying,
    /// then additionally proves that no other XML attribute in `content.xml`
    /// carries the selected page name. The name scan is intentionally
    /// conservative: even an unrelated containing attribute or fragment
    /// hyperlink blocks removal.
    pub(super) fn prepare_dependency_free_blank_slide_removal(
        &mut self,
        index: usize,
    ) -> Result<DependencyFreeBlankSlideRemoval> {
        if self.slides.len() <= 1 {
            return Err(litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide removal refuses the final slide".to_string(),
            ));
        }
        self.check_slide_move_supported()?;
        self.check_dependency_free_removal_package_owners()?;

        let source_origin = self.origins.get(index).copied().flatten().ok_or_else(|| {
            litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide removal requires an exact retained source page"
                    .to_string(),
            )
        })?;
        if self
            .page_overrides
            .get(index)
            .and_then(Option::as_ref)
            .is_some()
        {
            return Err(litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide removal cannot use a copied page".to_string(),
            ));
        }
        let source = self.content_source.as_ref().ok_or_else(|| {
            litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide removal requires retained content.xml fragments"
                    .to_string(),
            )
        })?;
        let page = source.page(source_origin).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "ODP retained page origin is outside content.xml coverage".to_string(),
            )
        })?;
        if page.len() > MAX_DEPENDENCY_FREE_COPY_PAGE_BYTES {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP dependency-free blank page exceeds {MAX_DEPENDENCY_FREE_COPY_PAGE_BYTES} bytes"
            )));
        }
        dependency_free_blank_name_value(page)?;

        let metadata = self
            .page_metadata
            .as_ref()
            .and_then(|value| value.page(index))
            .ok_or_else(|| {
                litchi_core::Error::Unsupported(
                    "ODP dependency-free blank-slide removal requires parsed draw:name metadata"
                        .to_string(),
                )
            })?;
        let name = metadata.name.as_deref().ok_or_else(|| {
            litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide removal requires draw:name".to_string(),
            )
        })?;
        if name.is_empty() {
            return Err(litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide removal requires a non-empty draw:name"
                    .to_string(),
            ));
        }
        if name.len() > MAX_DEPENDENCY_FREE_COPY_NAME_BYTES {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP dependency-free blank page name exceeds {MAX_DEPENDENCY_FREE_COPY_NAME_BYTES} bytes"
            )));
        }
        self.check_dependency_free_removal_name_ownership(name)?;

        let page_metadata = crate::model::page_metadata::metadata_after_page_remove(
            self.page_metadata.as_ref(),
            self.slides.len(),
            index,
        )?;
        Ok(DependencyFreeBlankSlideRemoval {
            index,
            page_metadata,
        })
    }

    /// Apply a fully prevalidated exact blank-page removal without another
    /// fallible parse or allocation.
    pub(super) fn apply_dependency_free_blank_slide_removal(
        &mut self,
        removal: DependencyFreeBlankSlideRemoval,
    ) -> Slide {
        let slide = self.slides.remove(removal.index);
        self.origins.remove(removal.index);
        self.page_overrides.remove(removal.index);
        self.foreign_page_overrides.remove(removal.index);
        self.page_metadata = removal.page_metadata;
        for (index, slide) in self.slides.iter_mut().enumerate().skip(removal.index) {
            slide.index = index;
        }
        slide
    }

    fn check_dependency_free_removal_package_owners(&self) -> Result<()> {
        let package = self.source_package.as_ref().ok_or_else(|| {
            litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide removal requires a retained source package"
                    .to_string(),
            )
        })?;
        if package.files()?.iter().any(|path| {
            let path = path.trim_start_matches('/');
            path == "Basic"
                || path.starts_with("Basic/")
                || path == "Scripts"
                || path.starts_with("Scripts/")
        }) {
            return Err(litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide removal refuses package macro owners".to_string(),
            ));
        }
        Ok(())
    }

    fn check_foreign_transfer_destination_safety(&self) -> Result<()> {
        let package = self.source_package.as_ref().ok_or_else(|| {
            litchi_core::Error::Unsupported(
                "ODP foreign blank-slide transfer requires a retained source package".to_string(),
            )
        })?;
        ensure_foreign_transfer_package_safety(package, "destination")
    }

    fn check_dependency_free_removal_name_ownership(&self, selected_name: &str) -> Result<()> {
        let package = self.source_package.as_ref().ok_or_else(|| {
            litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide removal requires a retained source package"
                    .to_string(),
            )
        })?;
        let mut selected_name_owners = 0usize;
        let archive = package.package()?;
        for path in archive
            .files()?
            .into_iter()
            .filter(|path| is_xml_owner_part(path, archive.manifest().get_media_type(path)))
        {
            let bytes = archive.get_file(&path)?;
            scan_dependency_owners(&path, &bytes, selected_name, &mut selected_name_owners)?;
        }
        if selected_name_owners != 1 {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP dependency-free blank-slide removal requires exactly one selected page-name owner; found {selected_name_owners}"
            )));
        }
        Ok(())
    }

    /// Add a package-contained audio or video payload.
    ///
    /// Existing source-package paths cannot be replaced implicitly. The
    /// returned inert reference can be attached with [`Shape::with_media`].
    pub(super) fn embed_media(
        &mut self,
        path: String,
        bytes: Vec<u8>,
        media_type: String,
    ) -> Result<Reference> {
        validate_package_media_path(&path)?;
        if let Some(package) = &self.source_package {
            let archive = package.package()?;
            ensure_media_manifest_rewritable(&archive)?;
            if package.has_file(&path)? {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "cannot replace existing ODP package media path '{path}' implicitly"
                )));
            }
        }
        embed_media(&mut self.media_files, path, bytes, media_type)
    }

    /// Preflight a package-media batch without changing the staged state.
    pub(super) fn prepare_media_changes(
        &mut self,
        changes: &[super::edit::MediaChange],
    ) -> Result<MediaChangePlan> {
        for change in changes {
            validate_media_change(change)?;
        }
        let source_package = self.source_package.clone().ok_or_else(|| {
            litchi_core::Error::Unsupported(
                "ODP media changes require a retained source package".to_string(),
            )
        })?;
        let archive = source_package.package()?;
        let mut files = self.media_files.clone();
        let mut removed = self.removed_media_paths.clone();
        let mut touched = BTreeSet::new();
        let mut changed = 0usize;

        for change in changes {
            let path = change.path();
            if !touched.insert(path.to_string()) {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "ODP media change batch selects '{path}' more than once"
                )));
            }
            match change {
                super::edit::MediaChange::Add {
                    payload,
                    media_type,
                    ..
                } => {
                    if archive.has_file(path) || files.contains_key(path) {
                        return Err(litchi_core::Error::InvalidFormat(format!(
                            "ODP media add path '{path}' already exists"
                        )));
                    }
                    files.insert(
                        path.to_string(),
                        EmbeddedMedia {
                            bytes: payload.clone(),
                            media_type: media_type.clone(),
                        },
                    );
                    changed = changed.checked_add(1).ok_or_else(|| {
                        litchi_core::Error::InvalidFormat(
                            "ODP media change count overflow".to_string(),
                        )
                    })?;
                },
                super::edit::MediaChange::Replace {
                    payload,
                    media_type,
                    ..
                } => {
                    let source_exists = archive.has_file(path);
                    let existing_type = archive.manifest().get_media_type(path);
                    let previous = files.get(path);
                    if source_exists && previous.is_none() && existing_type.is_none() {
                        return Err(litchi_core::Error::Unsupported(format!(
                            "ODP media replacement refuses member '{path}' without manifest metadata"
                        )));
                    }
                    if let Some(existing_type) = existing_type
                        && existing_type != media_type.as_str()
                    {
                        return Err(litchi_core::Error::Unsupported(format!(
                            "ODP media replacement cannot change manifest type for '{path}'"
                        )));
                    }
                    if let Some(previous) = previous {
                        if previous.bytes.as_slice() == payload.as_slice()
                            && previous.media_type == media_type.as_str()
                        {
                            continue;
                        }
                    } else if !source_exists {
                        return Err(litchi_core::Error::InvalidFormat(format!(
                            "ODP media replacement path '{path}' does not exist"
                        )));
                    }

                    let source_is_exact = if source_exists && previous.is_none() {
                        archive.get_file(path)?.as_slice() == payload.as_slice()
                            && existing_type == Some(media_type.as_str())
                    } else {
                        false
                    };
                    if source_is_exact && !removed.contains(path) {
                        continue;
                    }
                    if source_is_exact {
                        removed.remove(path);
                        changed = changed.checked_add(1).ok_or_else(|| {
                            litchi_core::Error::InvalidFormat(
                                "ODP media change count overflow".to_string(),
                            )
                        })?;
                        continue;
                    }
                    files.insert(
                        path.to_string(),
                        EmbeddedMedia {
                            bytes: payload.clone(),
                            media_type: media_type.clone(),
                        },
                    );
                    removed.remove(path);
                    changed = changed.checked_add(1).ok_or_else(|| {
                        litchi_core::Error::InvalidFormat(
                            "ODP media change count overflow".to_string(),
                        )
                    })?;
                },
                super::edit::MediaChange::Remove { .. } => {
                    let source_exists = archive.has_file(path);
                    if source_exists && self.source_media_is_referenced(path)? {
                        return Err(litchi_core::Error::Unsupported(format!(
                            "ODP media removal refuses referenced member '{path}'"
                        )));
                    }
                    if files.remove(path).is_some() {
                        if source_exists {
                            removed.insert(path.to_string());
                        } else {
                            removed.remove(path);
                        }
                        changed = changed.checked_add(1).ok_or_else(|| {
                            litchi_core::Error::InvalidFormat(
                                "ODP media change count overflow".to_string(),
                            )
                        })?;
                    } else if source_exists {
                        if !removed.insert(path.to_string()) {
                            return Err(litchi_core::Error::InvalidFormat(format!(
                                "ODP media member '{path}' is already removed"
                            )));
                        }
                        changed = changed.checked_add(1).ok_or_else(|| {
                            litchi_core::Error::InvalidFormat(
                                "ODP media change count overflow".to_string(),
                            )
                        })?;
                    } else {
                        return Err(litchi_core::Error::InvalidFormat(format!(
                            "ODP media removal path '{path}' does not exist"
                        )));
                    }
                },
            }
        }

        if changed > 0 {
            ensure_media_manifest_rewritable(&archive)?;
        }
        let staged_media_bytes = staged_media_bytes_for(&files)?;
        Ok(MediaChangePlan {
            media_files: files,
            removed_media_paths: removed,
            changed,
            staged_media_bytes,
        })
    }

    pub(super) fn apply_media_plan(&mut self, plan: MediaChangePlan) {
        self.media_files = plan.media_files;
        self.removed_media_paths = plan.removed_media_paths;
    }

    pub(super) fn has_staged_media_changes(&self) -> bool {
        !self.media_files.is_empty() || !self.removed_media_paths.is_empty()
    }

    pub(super) fn staged_media_bytes(&self) -> Result<usize> {
        staged_media_bytes_for(&self.media_files)
    }

    fn source_media_is_referenced(&mut self, path: &str) -> Result<bool> {
        if self.media_reference_index.is_none() {
            let source_package = self.source_package.clone().ok_or_else(|| {
                litchi_core::Error::Unsupported(
                    "ODP media references require a retained source package".to_string(),
                )
            })?;
            self.media_reference_index = Some(MediaReferenceIndex::build(&source_package)?);
        }
        Ok(self
            .media_reference_index
            .as_ref()
            .is_some_and(|index| index.referenced_paths.contains(path)))
    }

    /// Verify that every media part staged by this transaction survived package publication.
    ///
    /// This deliberately verifies package bytes and manifest metadata independently of slide
    /// references: callers may stage media before attaching it to a drawing frame.
    pub(super) fn verify_embedded_media(&self, reopened: &OwnedPackage) -> Result<()> {
        let package = reopened.package()?;
        for (path, expected) in &self.media_files {
            if !package.has_file(path) {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "published ODP package is missing staged media path '{path}'"
                )));
            }
            let actual = package.get_file(path)?;
            if actual != expected.bytes {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "published ODP media bytes differ for '{path}'"
                )));
            }
            let Some(entry) = package.manifest().get_entry(path) else {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "published ODP manifest is missing staged media entry '{path}'"
                )));
            };
            if entry.media_type != expected.media_type {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "published ODP media type differs for '{path}'"
                )));
            }
        }
        Ok(())
    }

    pub(super) fn verify_removed_media(&self, reopened: &OwnedPackage) -> Result<()> {
        if self.removed_media_paths.is_empty() {
            return Ok(());
        }
        let references = MediaReferenceIndex::build(reopened)?;
        for path in &self.removed_media_paths {
            if reopened.has_file(path)? {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "published ODP package retained removed media path '{path}'"
                )));
            }
            if references.referenced_paths.contains(path) {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "published ODP package retained a reference to removed media path '{path}'"
                )));
            }
        }
        Ok(())
    }

    /// Insert a slide at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Position to insert at (0-based)
    /// * `title` - Optional title for the slide
    /// * `text` - Text content for the slide
    ///
    /// # Examples
    ///
    /// The public equivalent starts from [`crate::edit::Snapshot::transaction`].
    ///
    /// ```
    /// use litchi_odp::{Builder, edit};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let source = edit::Snapshot::from_bytes(Builder::new().build()?)?;
    /// let mut transaction = source.transaction()?;
    /// transaction.add("First", "Content 1")?;
    /// transaction.add("Third", "Content 3")?;
    /// transaction.add_before(1, "Second", "Content 2")?;
    /// assert_eq!(transaction.slides()[1].title.as_deref(), Some("Second"));
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn insert_slide(&mut self, index: usize, title: &str, text: &str) -> Result<()> {
        if index <= self.slides.len() {
            let candidate_metadata = crate::model::page_metadata::metadata_after_page_insert(
                self.page_metadata.as_ref(),
                self.slides.len(),
                index,
            )?;
            let candidate_names = crate::model::page_metadata::effective_page_names(
                Some(&candidate_metadata),
                self.slides.len() + 1,
            )?;
            crate::model::settings::validate_page_references(
                self.settings.as_ref(),
                &candidate_names,
            )?;
            let slide = Slide {
                title: Some(title.to_string()),
                text: text.to_string(),
                index,
                notes: None,
                transition: None,
                animations: Vec::new(),
                legacy_animation: None,
                shapes: Vec::new(),
            };
            self.slides.insert(index, slide);
            self.origins.insert(index, None);
            self.page_overrides.insert(index, None);
            self.foreign_page_overrides.insert(index, false);
            self.page_metadata = Some(candidate_metadata);

            // Update indices of subsequent slides
            for i in (index + 1)..self.slides.len() {
                self.slides[i].index = i;
            }

            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.slides.len()
            )))
        }
    }

    /// Remove a slide at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the slide to remove (0-based)
    ///
    /// # Examples
    ///
    /// The public equivalent starts from [`crate::edit::Snapshot::transaction`].
    ///
    /// ```
    /// use litchi_odp::{Builder, edit};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let source = edit::Snapshot::from_bytes(Builder::new().build()?)?;
    /// let mut transaction = source.transaction()?;
    /// transaction.add("Slide 1", "Content 1")?;
    /// transaction.add("Slide 2", "Content 2")?;
    /// let removed = transaction.remove(0)?; // Remove first slide
    /// assert_eq!(removed.and_then(|slide| slide.title).as_deref(), Some("Slide 1"));
    /// assert_eq!(transaction.slides().len(), 1);
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn remove_slide(&mut self, index: usize) -> Result<Slide> {
        if index < self.slides.len() {
            let current_names = crate::model::page_metadata::effective_page_names(
                self.page_metadata.as_ref(),
                self.slides.len(),
            )?;
            let removed_name = &current_names[index];
            if crate::model::settings::settings_reference_page(self.settings.as_ref(), removed_name)
            {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "cannot remove presentation page '{removed_name}' because slide-show settings reference it"
                )));
            }
            let candidate_metadata = crate::model::page_metadata::metadata_after_page_remove(
                self.page_metadata.as_ref(),
                self.slides.len(),
                index,
            )?;
            let candidate_names = crate::model::page_metadata::effective_page_names(
                candidate_metadata.as_ref(),
                self.slides.len() - 1,
            )?;
            crate::model::settings::validate_page_references(
                self.settings.as_ref(),
                &candidate_names,
            )?;
            let slide = self.slides.remove(index);
            self.origins.remove(index);
            self.page_overrides.remove(index);
            self.foreign_page_overrides.remove(index);
            self.page_metadata = candidate_metadata;

            // Update indices of subsequent slides
            for i in index..self.slides.len() {
                self.slides[i].index = i;
            }

            Ok(slide)
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.slides.len()
            )))
        }
    }

    /// Move one slide to a final zero-based position without regenerating it.
    ///
    /// All fallible metadata and settings checks complete before the slide,
    /// source-fragment origin, or metadata order is changed.
    pub(super) fn move_slide(&mut self, from: usize, to: usize) -> Result<()> {
        if from >= self.slides.len() || to >= self.slides.len() {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "ODP slide move index is out of bounds (from: {from}, to: {to}, length: {})",
                self.slides.len()
            )));
        }
        if from == to {
            return Ok(());
        }

        let candidate_metadata = crate::model::page_metadata::metadata_after_page_move(
            self.page_metadata.as_ref(),
            self.slides.len(),
            from,
            to,
        )?;
        let candidate_names = crate::model::page_metadata::effective_page_names(
            candidate_metadata.as_ref(),
            self.slides.len(),
        )?;
        crate::model::settings::validate_page_references(self.settings.as_ref(), &candidate_names)?;

        let slide = self.slides.remove(from);
        self.slides.insert(to, slide);
        let origin = self.origins.remove(from);
        self.origins.insert(to, origin);
        let page_override = self.page_overrides.remove(from);
        self.page_overrides.insert(to, page_override);
        let foreign_page_override = self.foreign_page_overrides.remove(from);
        self.foreign_page_overrides
            .insert(to, foreign_page_override);
        self.page_metadata = candidate_metadata;
        for (index, slide) in self.slides.iter_mut().enumerate() {
            slide.index = index;
        }
        Ok(())
    }

    /// Update a slide at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the slide to update (0-based)
    /// * `title` - New title for the slide
    /// * `text` - New text content
    ///
    /// # Examples
    ///
    /// The public equivalent starts from [`crate::edit::Snapshot::transaction`].
    ///
    /// ```
    /// use litchi_odp::{Builder, edit};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let source = edit::Snapshot::from_bytes(Builder::new().build()?)?;
    /// let mut transaction = source.transaction()?;
    /// transaction.add("Old Title", "Old content")?;
    /// transaction.replace(0, "New Title", "New content")?;
    /// assert_eq!(transaction.slides()[0].title.as_deref(), Some("New Title"));
    /// assert_eq!(transaction.slides()[0].text, "New content");
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn update_slide(&mut self, index: usize, title: &str, text: &str) -> Result<()> {
        if index < self.slides.len() {
            self.slides[index].title = Some(title.to_string());
            self.slides[index].text = text.to_string();
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.slides.len()
            )))
        }
    }

    /// Add a shape to a slide.
    ///
    /// # Arguments
    ///
    /// * `slide_index` - Index of the slide to add the shape to
    /// * `shape` - Shape to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odp::{Builder, Shape, edit};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let source = edit::Snapshot::from_bytes(Builder::new().build()?)?;
    /// let mut transaction = source.transaction()?;
    /// transaction.add("Slide 1", "Content")?;
    /// let mut shape = Shape::new();
    /// shape.text = "Shape text".to_string();
    /// transaction.add_shape(0, shape)?;
    /// assert_eq!(transaction.slides()[0].shapes.len(), 1);
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn add_shape(&mut self, slide_index: usize, shape: Shape) -> Result<()> {
        if slide_index < self.slides.len() {
            self.slides[slide_index].shapes.push(shape);
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Slide index {slide_index} out of bounds"
            )))
        }
    }

    /// Remove a shape from a slide.
    ///
    /// # Arguments
    ///
    /// * `slide_index` - Index of the slide
    /// * `shape_index` - Index of the shape to remove
    ///
    /// # Examples
    ///
    /// The public equivalent starts from [`crate::edit::Snapshot::transaction`].
    ///
    /// ```
    /// use litchi_odp::{Builder, Shape, edit};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let source = edit::Snapshot::from_bytes(Builder::new().build()?)?;
    /// let mut transaction = source.transaction()?;
    /// transaction.add("Slide 1", "Content")?;
    /// // Add shape first, then remove it
    /// transaction.add_shape(0, Shape::new())?;
    /// let removed = transaction.remove_shape(0, 0)?;
    /// assert!(removed.is_some());
    /// assert!(transaction.slides()[0].shapes.is_empty());
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn remove_shape(&mut self, slide_index: usize, shape_index: usize) -> Result<Shape> {
        if slide_index < self.slides.len() {
            let slide = &mut self.slides[slide_index];
            if shape_index < slide.shapes.len() {
                Ok(slide.shapes.remove(shape_index))
            } else {
                Err(litchi_core::Error::InvalidFormat(format!(
                    "Shape index {shape_index} out of bounds"
                )))
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Slide index {slide_index} out of bounds"
            )))
        }
    }

    /// Clear all content (text and shapes) from a slide.
    ///
    /// # Examples
    ///
    /// The public equivalent starts from [`crate::edit::Snapshot::transaction`].
    ///
    /// ```
    /// use litchi_odp::{Builder, edit};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let source = edit::Snapshot::from_bytes(Builder::new().build()?)?;
    /// let mut transaction = source.transaction()?;
    /// transaction.add("Slide 1", "Content")?;
    /// transaction.clear(0)?;
    /// assert!(transaction.slides()[0].title.is_none());
    /// assert!(transaction.slides()[0].text.is_empty());
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn clear_slide(&mut self, slide_index: usize) -> Result<()> {
        if slide_index < self.slides.len() {
            self.slides[slide_index].title = None;
            self.slides[slide_index].text.clear();
            self.slides[slide_index].shapes.clear();
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Slide index {slide_index} out of bounds"
            )))
        }
    }

    /// Locate the verbatim source markup for a slide that has not been edited.
    ///
    /// Returns `None` whenever the model cannot prove the slide is byte-identical
    /// in meaning to one of the source slides, in which case the caller
    /// regenerates it from the model.
    fn retained_page(&self, slide_index: usize, slide: &Slide) -> Option<&str> {
        let source = self.content_source.as_ref()?;
        if self.declarations != self.source_declarations {
            return None;
        }
        if self
            .foreign_page_overrides
            .get(slide_index)
            .copied()
            .unwrap_or(false)
        {
            return self
                .page_overrides
                .get(slide_index)
                .and_then(Option::as_deref);
        }
        let source_index = self.origins.get(slide_index).copied().flatten()?;
        self.source_slides
            .get(source_index)
            .filter(|candidate| slide_content_eq(candidate, slide))?;
        if let Some(page) = self
            .page_overrides
            .get(slide_index)
            .and_then(Option::as_deref)
        {
            return Some(page);
        }
        source.page(source_index)
    }

    /// Build the `office:automatic-styles` payload for a retained document.
    ///
    /// Only slides that were regenerated need synthesised drawing-page styles,
    /// and any name the source already defines is left untouched so retained
    /// markup keeps referring to the original definition.
    fn generate_page_styles(&self, regenerated: &[usize]) -> String {
        let defines = |name: &str| {
            self.content_source
                .as_ref()
                .is_some_and(|source| source.defines_style(name))
        };
        let mut output = String::new();
        let needs_default = regenerated.iter().any(|index| {
            self.slides.get(*index).is_some_and(|slide| {
                super::builder::slide_style_name(slide, *index)
                    == super::builder::DEFAULT_DRAWING_PAGE_STYLE_NAME
            })
        });
        if needs_default && !defines(super::builder::DEFAULT_DRAWING_PAGE_STYLE_NAME) {
            output.push_str(super::builder::DEFAULT_DRAWING_PAGE_STYLE);
        }
        for &index in regenerated {
            let Some(slide) = self.slides.get(index) else {
                continue;
            };
            if defines(&super::builder::slide_style_name(slide, index)) {
                continue;
            }
            super::builder::push_transition_style(&mut output, slide, index);
        }
        output
    }

    /// Generate content.xml from the current mutable state.
    fn generate_content_xml(&self) -> Result<String> {
        let mut extension_uris = BTreeSet::new();
        for slide in &self.slides {
            validate_animation_roots(&slide.animations)?;
            for animation in &slide.animations {
                animation.collect_extension_namespaces(&mut extension_uris);
            }
            if let Some(animation) = &slide.legacy_animation {
                validate_legacy_animation_root(animation)?;
                animation.collect_extension_namespaces(&mut extension_uris);
            }
        }
        let extension_namespaces = extension_uris
            .into_iter()
            .enumerate()
            .map(|(index, uri)| (uri, format!("anim-ext{}", index + 1)))
            .collect::<BTreeMap<_, _>>();
        let mut extension_declarations = String::new();
        for (uri, prefix) in &extension_namespaces {
            if uri.is_empty() {
                return Err(litchi_core::Error::InvalidFormat(
                    "animation extension namespace URI cannot be empty".to_string(),
                ));
            }
            extension_declarations.push_str(" xmlns:");
            extension_declarations.push_str(prefix);
            extension_declarations.push_str("=\"");
            extension_declarations.push_str(&escape_xml(uri));
            extension_declarations.push('"');
        }

        let shape_count = self.slides.iter().map(|s| s.shapes.len()).sum::<usize>();
        let mut estimated = 256usize;
        estimated += self.slides.len() * 128;
        estimated += shape_count * 192;
        estimated += self
            .slides
            .iter()
            .map(|slide| slide.text.len() + slide.title.as_ref().map_or(0, String::len))
            .sum::<usize>();
        estimated += self
            .slides
            .iter()
            .flat_map(|s| s.shapes.iter())
            .map(|shape| shape.text.len() + shape.name.as_ref().map_or(0, String::len))
            .sum::<usize>();

        let mut body = String::with_capacity(estimated);

        body.push_str(&crate::model::declaration::write_declaration_elements(
            self.declarations.as_ref(),
            self.slides.len(),
        )?);
        if let Some(source) = &self.content_source {
            source.write_leading_extras(&mut body);
        }

        let page_names = crate::model::page_metadata::effective_page_names(
            self.page_metadata.as_ref(),
            self.slides.len(),
        )?;
        crate::model::settings::validate_page_references(self.settings.as_ref(), &page_names)?;

        let mut regenerated = Vec::new();
        for (i, slide) in self.slides.iter().enumerate() {
            if let Some(page) = self.retained_page(i, slide) {
                body.push_str(page);
                continue;
            }
            regenerated.push(i);
            let page_num = i + 1;
            let slide_style = super::builder::slide_style_name(slide, i);
            if let Some(metadata) = &self.page_metadata {
                metadata.validate_for_slides(self.slides.len())?;
            }
            let page_attributes = crate::model::page_metadata::write_page_attributes(
                self.page_metadata.as_ref(),
                i,
                &slide_style,
            )?;
            let declaration_attributes = crate::model::declaration::write_binding_attributes(
                self.declarations.as_ref(),
                i,
                crate::model::declaration::Target::Slide,
            );
            let _ = page_num;
            body.push_str("<draw:page");
            body.push_str(&page_attributes);
            body.push_str(&declaration_attributes);
            body.push('>');

            // Add title frame if title exists
            if let Some(ref title) = slide.title {
                let title_paragraphs = super::builder::generate_text_paragraphs(title, Some("P1"));
                body.push_str(&xml_minifier::minified_xml_format!(
                    r#"<draw:frame draw:style-name="gr1" draw:text-style-name="P1" draw:layer="layout" presentation:class="title" svg:width="25.199cm" svg:height="3.506cm" svg:x="1.4cm" svg:y="0.962cm"><draw:text-box>{}</draw:text-box></draw:frame>"#,
                    title_paragraphs
                ));
            }

            // Add text frame
            if !slide.text.is_empty() {
                let y_position = if slide.title.is_some() {
                    "5.0cm"
                } else {
                    "2.0cm"
                };
                let text_paragraphs =
                    super::builder::generate_text_paragraphs(&slide.text, Some("P2"));
                body.push_str(&xml_minifier::minified_xml_format!(
                    r#"<draw:frame draw:style-name="gr2" draw:text-style-name="P2" draw:layer="layout" presentation:class="object" svg:width="25.199cm" svg:height="10cm" svg:x="1.4cm" svg:y="{}"><draw:text-box>{}</draw:text-box></draw:frame>"#,
                    y_position,
                    text_paragraphs
                ));
            }

            // Add shapes
            for (shape_idx, shape) in slide.shapes.iter().enumerate() {
                body.push_str(&super::builder::Builder::generate_shape_xml(
                    shape, shape_idx,
                )?);
            }

            for animation in &slide.animations {
                animation.write_xml(&mut body, &extension_namespaces)?;
            }
            if let Some(animation) = &slide.legacy_animation {
                animation.write_xml(&mut body, &extension_namespaces)?;
            }

            let notes_attributes = crate::model::declaration::write_binding_attributes(
                self.declarations.as_ref(),
                i,
                crate::model::declaration::Target::Notes,
            );
            body.push_str(&crate::model::declaration::apply_notes_binding(
                super::builder::Builder::generate_notes_xml(slide.notes.as_deref()),
                &notes_attributes,
            )?);

            body.push_str("</draw:page>");
        }

        if let Some(source) = &self.content_source {
            source.write_trailing_extras(&mut body);
        }
        body.push_str(&crate::model::settings::write(self.settings.as_ref())?);

        if let Some(source) = &self.content_source {
            let styles = self.generate_page_styles(&regenerated);
            let synthesised = !regenerated.is_empty() || !styles.is_empty();
            let mut output = String::with_capacity(body.len() + estimated);
            source.write_prolog(&mut output);
            output.push_str(&source.root_start_tag(&extension_declarations, synthesised)?);
            source.write_prologue(&mut output, &styles);
            output.push_str(source.body_start_tag());
            output.push_str(source.presentation_start_tag());
            output.push_str(&body);
            output.push_str(&source.close_tags());
            return Ok(output);
        }

        let transition_styles = super::builder::generate_transition_styles(&self.slides);
        Ok(format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"{extension_declarations} office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles>{transition_styles}</office:automatic-styles><office:body><office:presentation>{body}</office:presentation></office:body></office:document-content>"#
        ))
    }

    /// Preserve source metadata exactly or create deterministic minimal metadata.
    fn generate_meta_xml(&self) -> Result<String> {
        if let Some(package) = &self.source_package
            && let Ok(bytes) = package.get_file("meta.xml")
        {
            return String::from_utf8(bytes).map_err(|error| {
                litchi_core::Error::InvalidFormat(format!(
                    "ODP source meta.xml is not UTF-8: {error}"
                ))
            });
        }
        Ok(r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator></office:meta></office:document-meta>"#.to_string())
    }

    /// Convert the presentation to bytes.
    ///
    /// # Examples
    ///
    /// The public equivalent starts from [`crate::edit::Snapshot::transaction`].
    ///
    /// ```
    /// use litchi_odp::{Builder, edit};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let source = edit::Snapshot::from_bytes(Builder::new().build()?)?;
    /// let mut transaction = source.transaction()?;
    /// transaction.add("Slide 1", "Content")?;
    /// let commit = transaction.commit()?;
    /// let bytes = commit.snapshot().bytes();
    /// assert!(!bytes.is_empty());
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn to_bytes_bounded(&self, limit: usize) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new_bounded(limit);
        let source_manifest = self
            .source_package
            .as_ref()
            .map(|package| package.package())
            .transpose()?;

        // Set MIME type
        writer.set_mimetype(&self.mimetype)?;

        // Add content.xml (regenerated from mutable state)
        let content_xml = self.generate_content_xml()?;
        let content_media_type = source_manifest
            .as_ref()
            .and_then(|package| package.manifest().get_media_type("content.xml"))
            .unwrap_or("text/xml");
        writer.add_file_with_media_type(
            "content.xml",
            content_xml.as_bytes(),
            content_media_type,
        )?;

        // Add styles.xml (preserved or default)
        let default_styles = Structure::default_styles_xml();
        let styles_xml = self.styles_xml.as_deref().unwrap_or(&default_styles);
        let styles_media_type = source_manifest
            .as_ref()
            .and_then(|package| package.manifest().get_media_type("styles.xml"))
            .unwrap_or("text/xml");
        writer.add_file_with_media_type("styles.xml", styles_xml.as_bytes(), styles_media_type)?;

        // Add meta.xml (patched from the source or regenerated with current metadata)
        let meta_xml = self.generate_meta_xml()?;
        let meta_media_type = source_manifest
            .as_ref()
            .and_then(|package| package.manifest().get_media_type("meta.xml"))
            .unwrap_or("text/xml");
        writer.add_file_with_media_type("meta.xml", meta_xml.as_bytes(), meta_media_type)?;

        for (path, media) in &self.media_files {
            writer.add_file_with_media_type(path, &media.bytes, &media.media_type)?;
        }

        if let Some(package) = &self.source_package {
            let mut excluded = self.media_files.keys().cloned().collect::<Vec<_>>();
            excluded.extend(self.removed_media_paths.iter().cloned());
            writer.copy_auxiliary_files_from_except(package, &excluded, &[])?;
            // Opt in only after every regenerated and copied member has been
            // staged. PackageWriter decides at finalization whether the
            // source manifest inventory still matches exactly.
            writer.preserve_source_manifest(package)?;
        }

        writer.finish_to_bounded_bytes()
    }
}

pub(super) fn ensure_foreign_transfer_package_safety(
    package: &OwnedPackage,
    side: &str,
) -> Result<()> {
    let archive = package.package()?;
    let files = archive.files()?;
    if files.iter().any(|path| is_macro_owner_path(path)) {
        return Err(litchi_core::Error::Unsupported(format!(
            "ODP foreign blank-slide transfer refuses {side} package macro owners"
        )));
    }
    if files
        .iter()
        .any(|path| crate::core::is_signature_owner_path(path))
    {
        return Err(litchi_core::Error::Unsupported(format!(
            "ODP foreign blank-slide transfer refuses {side} signature owners"
        )));
    }
    let manifest = archive.manifest();
    if manifest.has_encrypted_entries() {
        return Err(litchi_core::Error::Unsupported(format!(
            "ODP foreign blank-slide transfer refuses encrypted {side} package entries"
        )));
    }
    for path in files {
        if !is_xml_owner_part(&path, manifest.get_media_type(&path)) {
            continue;
        }
        let bytes = archive.get_file(&path)?;
        let xml = std::str::from_utf8(&bytes).map_err(|error| {
            litchi_core::Error::InvalidFormat(format!(
                "ODP {side} XML part '{path}' is not UTF-8 during transfer safety validation: {error}"
            ))
        })?;
        ensure_no_macro_xml_owners(xml, side)?;
        ensure_no_protected_xml(xml, side)?;
        if path.eq_ignore_ascii_case("content.xml") {
            ensure_foreign_transfer_namespace_ancestors(xml, side)?;
        }
    }
    Ok(())
}

fn ensure_foreign_transfer_namespace_ancestors(xml: &str, side: &str) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut scopes = Vec::new();
    let mut root_scope = None;
    loop {
        let decoder = reader.decoder();
        let (resolved, event) = reader.read_resolved_event().map_err(|error| {
            litchi_core::Error::InvalidFormat(format!(
                "invalid ODP {side} content.xml during namespace validation: {error}"
            ))
        })?;
        match event {
            Event::Start(element) => {
                let scope = effective_namespace_scope(scopes.last(), &element, decoder, side)?;
                let target = namespace_ancestor_target(&resolved, element.local_name().as_ref());
                if scopes.is_empty() {
                    root_scope = Some(scope.clone());
                }
                if let Some(target) = target {
                    validate_namespace_ancestor(root_scope.as_ref(), &scope, target, false, side)?;
                }
                scopes.push(scope);
            },
            Event::Empty(element) => {
                let scope = effective_namespace_scope(scopes.last(), &element, decoder, side)?;
                let target = namespace_ancestor_target(&resolved, element.local_name().as_ref());
                if scopes.is_empty() {
                    root_scope = Some(scope.clone());
                }
                if let Some(target) = target {
                    validate_namespace_ancestor(root_scope.as_ref(), &scope, target, true, side)?;
                }
            },
            Event::End(_) => {
                scopes.pop().ok_or_else(|| {
                    litchi_core::Error::InvalidFormat(format!(
                        "ODP {side} content.xml namespace scope underflow"
                    ))
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !scopes.is_empty() {
        return Err(litchi_core::Error::InvalidFormat(format!(
            "ODP {side} content.xml namespace scope is unterminated"
        )));
    }
    Ok(())
}

fn effective_namespace_scope(
    parent: Option<&BTreeMap<String, String>>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    side: &str,
) -> Result<BTreeMap<String, String>> {
    let mut scope = parent.cloned().unwrap_or_default();
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute.map_err(|error| {
            litchi_core::Error::InvalidFormat(format!(
                "invalid ODP {side} namespace declaration: {error}"
            ))
        })?;
        let key = attribute.key.as_ref();
        let prefix = if key == XMLNS_DEFAULT {
            String::new()
        } else if let Some(rest) = key.strip_prefix(XMLNS_PREFIX) {
            String::from_utf8(rest.to_vec()).map_err(|error| {
                litchi_core::Error::InvalidFormat(format!(
                    "invalid ODP {side} namespace prefix: {error}"
                ))
            })?
        } else {
            continue;
        };
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| {
                litchi_core::Error::InvalidFormat(format!(
                    "invalid ODP {side} namespace URI: {error}"
                ))
            })?;
        scope.insert(prefix, value.into_owned());
        if scope.len() > MAX_NAMESPACE_BINDINGS {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP foreign blank-slide transfer refuses {side} namespace scope above {MAX_NAMESPACE_BINDINGS} bindings"
            )));
        }
    }
    Ok(scope)
}

fn namespace_ancestor_target<'a>(resolved: &ResolveResult<'_>, local: &[u8]) -> Option<&'a str> {
    let ResolveResult::Bound(Namespace(uri)) = resolved else {
        return None;
    };
    if *uri != OFFICE_NAMESPACE.as_bytes() {
        return None;
    }
    match local {
        b"body" => Some("office:body"),
        b"presentation" => Some("office:presentation"),
        b"automatic-styles" => Some("office:automatic-styles"),
        _ => None,
    }
}

fn validate_namespace_ancestor(
    root: Option<&BTreeMap<String, String>>,
    current: &BTreeMap<String, String>,
    target: &str,
    empty: bool,
    side: &str,
) -> Result<()> {
    if target == "office:presentation" && empty {
        return Err(litchi_core::Error::Unsupported(format!(
            "ODP foreign blank-slide transfer refuses self-closing {side} office:presentation"
        )));
    }
    let Some(root) = root else {
        return Err(litchi_core::Error::InvalidFormat(format!(
            "ODP {side} namespace validation found {target} before document root"
        )));
    };
    if root.get("") != current.get("") {
        return Err(litchi_core::Error::Unsupported(format!(
            "ODP foreign blank-slide transfer refuses {side} default-namespace rebinding at {target}"
        )));
    }
    const REQUIRED: [(&str, &str); 12] = [
        ("office", OFFICE_NAMESPACE),
        ("draw", DRAW_NAMESPACE),
        ("style", STYLE_NAMESPACE),
        ("presentation", PRESENTATION_NAMESPACE),
        ("text", TEXT_NAMESPACE),
        ("svg", SVG_NAMESPACE),
        ("xlink", XLINK_NAMESPACE),
        ("anim", ANIMATION_NAMESPACE),
        ("smil", SMIL_NAMESPACE),
        ("dr3d", DR3D_NAMESPACE),
        ("script", SCRIPT_NAMESPACE),
        ("xml", XML_NAMESPACE),
    ];
    for (prefix, expected) in REQUIRED {
        let root_binding = root.get(prefix);
        let current_binding = current.get(prefix);
        if root_binding != current_binding {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP foreign blank-slide transfer refuses {side} namespace rebinding of '{prefix}' at {target}"
            )));
        }
        if let Some(binding) = current_binding
            && binding != expected
        {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP foreign blank-slide transfer refuses {side} prefix '{prefix}' bound to '{binding}' at {target}"
            )));
        }
    }
    Ok(())
}

fn is_macro_owner_path(path: &str) -> bool {
    path.trim_start_matches('/').split('/').any(|component| {
        component.eq_ignore_ascii_case("basic") || component.eq_ignore_ascii_case("scripts")
    })
}

pub(super) fn ensure_no_macro_xml_owners(content: &str, side: &str) -> Result<()> {
    let mut reader = Reader::from_str(content);
    reader.config_mut().check_end_names = true;
    loop {
        match reader.read_event().map_err(|error| {
            litchi_core::Error::InvalidFormat(format!(
                "invalid ODP {side} content.xml during macro-owner validation: {error}"
            ))
        })? {
            Event::Start(element) => {
                if is_macro_owner_name(element.local_name().as_ref()) {
                    return Err(litchi_core::Error::Unsupported(format!(
                        "ODP foreign blank-slide transfer refuses {side} XML macro owners"
                    )));
                }
            },
            Event::Empty(element) => {
                let local_name = element.local_name();
                if is_macro_owner_name(local_name.as_ref()) && local_name.as_ref() != b"scripts" {
                    return Err(litchi_core::Error::Unsupported(format!(
                        "ODP foreign blank-slide transfer refuses {side} XML macro owners"
                    )));
                }
            },
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(litchi_core::Error::Unsupported(format!(
                    "ODP foreign blank-slide transfer refuses {side} DTD or entity owners"
                )));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

pub(super) fn ensure_no_protected_xml(content: &str, side: &str) -> Result<()> {
    let mut reader = Reader::from_str(content);
    reader.config_mut().check_end_names = true;
    loop {
        match reader.read_event().map_err(|error| {
            litchi_core::Error::InvalidFormat(format!(
                "invalid ODP {side} content.xml during protection validation: {error}"
            ))
        })? {
            Event::Start(element) | Event::Empty(element) => {
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| {
                        litchi_core::Error::InvalidFormat(format!(
                            "invalid ODP {side} protection attribute: {error}"
                        ))
                    })?;
                    let local_name = attribute.key.local_name();
                    if matches!(
                        local_name.as_ref(),
                        b"protect" | b"protected" | b"protection-key" | b"protection-key-digest"
                    ) {
                        return Err(litchi_core::Error::Unsupported(format!(
                            "ODP foreign blank-slide transfer refuses {side} protected XML"
                        )));
                    }
                }
            },
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(litchi_core::Error::Unsupported(format!(
                    "ODP foreign blank-slide transfer refuses {side} DTD or entity owners"
                )));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

fn count_selected_name_owners(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    selected_name: &str,
    selected_name_owners: &mut usize,
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            litchi_core::Error::InvalidFormat(format!(
                "invalid ODP XML attribute during blank-slide ownership validation: {error}"
            ))
        })?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                litchi_core::Error::InvalidFormat(format!(
                    "invalid ODP XML attribute value during blank-slide ownership validation: {error}"
                ))
            })?;
        let key = attribute.key.as_ref();
        let local_key = attribute.key.local_name();
        if matches!(local_key.as_ref(), b"macro-name" | b"language") {
            return Err(litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide removal refuses macro binding attributes"
                    .to_string(),
            ));
        }
        if (key == b"href" || key.ends_with(b":href")) && value.contains('#') {
            return Err(litchi_core::Error::Unsupported(
                "ODP dependency-free blank-slide removal refuses fragment hyperlink owners"
                    .to_string(),
            ));
        }
        if value.contains(selected_name) {
            *selected_name_owners = selected_name_owners.checked_add(1).ok_or_else(|| {
                litchi_core::Error::InvalidFormat(
                    "ODP selected page ownership count overflow".to_string(),
                )
            })?;
        }
    }
    Ok(())
}

fn scan_dependency_owners(
    path: &str,
    bytes: &[u8],
    selected_name: &str,
    selected_name_owners: &mut usize,
) -> Result<()> {
    let xml = std::str::from_utf8(bytes).map_err(|error| {
        litchi_core::Error::InvalidFormat(format!(
            "ODP XML part '{path}' is not UTF-8 during blank-slide ownership validation: {error}"
        ))
    })?;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    loop {
        match reader.read_event().map_err(|error| {
            litchi_core::Error::InvalidFormat(format!(
                "invalid ODP XML part '{path}' during blank-slide ownership validation: {error}"
            ))
        })? {
            Event::Start(element) => {
                let local_name = element.local_name();
                if is_macro_owner_name(local_name.as_ref()) {
                    return Err(litchi_core::Error::Unsupported(
                        "ODP dependency-free blank-slide removal refuses content macro owners"
                            .to_string(),
                    ));
                }
                count_selected_name_owners(&reader, &element, selected_name, selected_name_owners)?;
            },
            Event::Empty(element) => {
                let local_name = element.local_name();
                if is_macro_owner_name(local_name.as_ref()) && local_name.as_ref() != b"scripts" {
                    return Err(litchi_core::Error::Unsupported(
                        "ODP dependency-free blank-slide removal refuses content macro owners"
                            .to_string(),
                    ));
                }
                count_selected_name_owners(&reader, &element, selected_name, selected_name_owners)?;
            },
            Event::DocType(_) => {
                return Err(litchi_core::Error::Unsupported(
                    "ODP dependency-free blank-slide removal refuses DTD/entity ownership"
                        .to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}

fn is_macro_owner_name(local_name: &[u8]) -> bool {
    matches!(local_name, b"scripts" | b"script" | b"event-listener")
}

fn is_xml_owner_part(path: &str, media_type: Option<&str>) -> bool {
    let xml_extension = path.rsplit_once('.').is_some_and(|(_, extension)| {
        extension.eq_ignore_ascii_case("xml")
            || extension.eq_ignore_ascii_case("rdf")
            || extension.eq_ignore_ascii_case("rels")
    });
    let xml_media_type = media_type.is_some_and(|value| {
        let value = value
            .split_once(';')
            .map_or(value, |(media_type, _)| media_type)
            .trim();
        value.eq_ignore_ascii_case("text/xml")
            || value.eq_ignore_ascii_case("application/xml")
            || value
                .get(value.len().saturating_sub(4)..)
                .is_some_and(|suffix| suffix.eq_ignore_ascii_case("+xml"))
    });
    path.eq_ignore_ascii_case("content.xml")
        || path.eq_ignore_ascii_case("styles.xml")
        || xml_extension
        || xml_media_type
}

fn validate_media_change(change: &super::edit::MediaChange) -> Result<()> {
    match change {
        super::edit::MediaChange::Add {
            path,
            payload,
            media_type,
        }
        | super::edit::MediaChange::Replace {
            path,
            payload,
            media_type,
        } => {
            validate_package_media_path(path)?;
            validate_media_payload(payload, media_type)?;
        },
        super::edit::MediaChange::Remove { path } => {
            validate_package_media_path(path)?;
        },
    }
    Ok(())
}

fn validate_media_payload(payload: &[u8], media_type: &str) -> Result<()> {
    crate::model::media::validate_media_type(media_type)?;
    if payload.len() > MAX_MEDIA_CHANGE_PAYLOAD_BYTES {
        return Err(litchi_core::Error::InvalidFormat(format!(
            "ODP media change payload exceeds {MAX_MEDIA_CHANGE_PAYLOAD_BYTES} bytes"
        )));
    }
    Ok(())
}

fn ensure_media_manifest_rewritable(
    archive: &litchi_odf_common::core::package::Package<'_>,
) -> Result<()> {
    if archive
        .manifest()
        .entries
        .values()
        .any(|entry| entry.size.is_some() || entry.encryption.is_some())
    {
        return Err(litchi_core::Error::Unsupported(
            "ODP media changes refuse manifest metadata the writer cannot preserve".to_string(),
        ));
    }
    Ok(())
}

fn staged_media_bytes_for(files: &BTreeMap<String, EmbeddedMedia>) -> Result<usize> {
    files.iter().try_fold(0usize, |total, (path, media)| {
        let size = path
            .len()
            .checked_add(media.media_type.len())
            .and_then(|value| value.checked_add(media.bytes.len()))
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat("ODP staged media size overflow".to_string())
            })?;
        total.checked_add(size).ok_or_else(|| {
            litchi_core::Error::InvalidFormat("ODP staged media total overflow".to_string())
        })
    })
}

impl MediaReferenceIndex {
    fn build(package: &OwnedPackage) -> Result<Self> {
        let archive = package.package()?;
        let files = archive.files()?;
        if files.len() > MAX_MEDIA_REFERENCE_XML_PARTS {
            return Err(litchi_core::Error::ResourceLimit(
                litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Objects,
                    observed: files.len() as u64,
                    limit: MAX_MEDIA_REFERENCE_XML_PARTS as u64,
                    scope: std::sync::Arc::from("ODP media reference XML parts"),
                },
            ));
        }
        let mut index = Self::default();
        let mut path_bytes = 0usize;
        let mut xml_bytes = 0usize;
        let mut events = 0usize;
        for path in files {
            if path.ends_with('/')
                || path.eq_ignore_ascii_case("META-INF/manifest.xml")
                || path.eq_ignore_ascii_case("manifest.xml")
                || crate::core::is_signature_owner_path(&path)
                || !is_xml_owner_part(&path, archive.manifest().get_media_type(&path))
            {
                continue;
            }
            let bytes = archive.get_file(&path)?;
            if bytes.len() > MAX_MEDIA_REFERENCE_XML_BYTES {
                return Err(litchi_core::Error::ResourceLimit(
                    litchi_core::ResourceLimit {
                        resource: litchi_core::Resource::InputBytes,
                        observed: bytes.len() as u64,
                        limit: MAX_MEDIA_REFERENCE_XML_BYTES as u64,
                        scope: std::sync::Arc::from("ODP media reference XML part"),
                    },
                ));
            }
            xml_bytes = xml_bytes.checked_add(bytes.len()).ok_or_else(|| {
                litchi_core::Error::InvalidFormat(
                    "ODP media reference XML byte count overflow".to_string(),
                )
            })?;
            if xml_bytes > MAX_MEDIA_REFERENCE_TOTAL_XML_BYTES {
                return Err(litchi_core::Error::ResourceLimit(
                    litchi_core::ResourceLimit {
                        resource: litchi_core::Resource::InputBytes,
                        observed: xml_bytes as u64,
                        limit: MAX_MEDIA_REFERENCE_TOTAL_XML_BYTES as u64,
                        scope: std::sync::Arc::from("ODP media reference XML parts"),
                    },
                ));
            }
            scan_media_reference_xml(
                &path,
                &bytes,
                &mut index.referenced_paths,
                &mut path_bytes,
                &mut events,
            )?;
        }
        Ok(index)
    }
}

fn scan_media_reference_xml(
    owner_path: &str,
    bytes: &[u8],
    referenced_paths: &mut HashSet<String>,
    referenced_path_bytes: &mut usize,
    events: &mut usize,
) -> Result<()> {
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let relationship_part = owner_path
        .rsplit_once('.')
        .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("rels"));
    let mut part_events = 0usize;
    loop {
        part_events = part_events.checked_add(1).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "ODP media reference XML event count overflow".to_string(),
            )
        })?;
        *events = events.checked_add(1).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "ODP media reference XML event count overflow".to_string(),
            )
        })?;
        if part_events > MAX_MEDIA_REFERENCE_XML_EVENTS {
            return Err(litchi_core::Error::ResourceLimit(
                litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Work,
                    observed: part_events as u64,
                    limit: MAX_MEDIA_REFERENCE_XML_EVENTS as u64,
                    scope: std::sync::Arc::from("ODP media reference XML"),
                },
            ));
        }
        if *events > MAX_MEDIA_REFERENCE_TOTAL_XML_EVENTS {
            return Err(litchi_core::Error::ResourceLimit(
                litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Work,
                    observed: *events as u64,
                    limit: MAX_MEDIA_REFERENCE_XML_EVENTS as u64,
                    scope: std::sync::Arc::from("ODP media reference XML"),
                },
            ));
        }
        let (_namespace, event) =
            reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| {
                    litchi_core::Error::InvalidFormat(format!(
                        "invalid ODP media reference XML part '{owner_path}': {error}"
                    ))
                })?;
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    litchi_core::Error::InvalidFormat(
                        "ODP media reference XML depth overflow".to_string(),
                    )
                })?;
                if depth > MAX_MEDIA_REFERENCE_XML_DEPTH {
                    return Err(litchi_core::Error::ResourceLimit(
                        litchi_core::ResourceLimit {
                            resource: litchi_core::Resource::Depth,
                            observed: depth as u64,
                            limit: MAX_MEDIA_REFERENCE_XML_DEPTH as u64,
                            scope: std::sync::Arc::from("ODP media reference XML"),
                        },
                    ));
                }
                scan_media_reference_attributes(
                    &reader,
                    &element,
                    owner_path,
                    relationship_part,
                    referenced_paths,
                    referenced_path_bytes,
                )?;
            },
            Event::Empty(element) => {
                scan_media_reference_attributes(
                    &reader,
                    &element,
                    owner_path,
                    relationship_part,
                    referenced_paths,
                    referenced_path_bytes,
                )?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    litchi_core::Error::InvalidFormat(format!(
                        "ODP media reference XML part '{owner_path}' has an unmatched end"
                    ))
                })?;
            },
            Event::DocType(_) | Event::GeneralRef(_) => {
                return Err(litchi_core::Error::Unsupported(format!(
                    "ODP media reference XML part '{owner_path}' contains a DTD or entity reference"
                )));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(())
}

fn scan_media_reference_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    owner_path: &str,
    relationship_part: bool,
    referenced_paths: &mut HashSet<String>,
    referenced_path_bytes: &mut usize,
) -> Result<()> {
    let mut candidates = Vec::new();
    let mut external_target = false;
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute.map_err(|error| {
            litchi_core::Error::InvalidFormat(format!(
                "invalid ODP media reference attribute in '{owner_path}': {error}"
            ))
        })?;
        let (_, local) = reader.resolver().resolve_attribute(attribute.key);
        let local = local.as_ref();
        let target_mode = relationship_part && local == b"TargetMode";
        let candidate =
            local == b"href" || local == b"resource" || (relationship_part && local == b"Target");
        if !target_mode && !candidate {
            continue;
        }
        if attribute.value.len() > MAX_MEDIA_REFERENCE_VALUE_BYTES {
            return Err(litchi_core::Error::ResourceLimit(
                litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::InputBytes,
                    observed: attribute.value.len() as u64,
                    limit: MAX_MEDIA_REFERENCE_VALUE_BYTES as u64,
                    scope: std::sync::Arc::from("ODP media reference XML attribute"),
                },
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                litchi_core::Error::InvalidFormat(format!(
                    "invalid ODP media reference attribute in '{owner_path}': {error}"
                ))
            })?
            .into_owned();
        if target_mode {
            external_target = value.eq_ignore_ascii_case("external");
        } else {
            candidates.push(value);
        }
    }
    if external_target {
        return Ok(());
    }
    for candidate in candidates {
        let Some(path) = resolve_owner_relative_reference(owner_path, &candidate)? else {
            continue;
        };
        if referenced_paths.contains(&path) {
            continue;
        }
        if referenced_paths.len() >= MAX_MEDIA_REFERENCE_PATHS {
            return Err(litchi_core::Error::ResourceLimit(
                litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Objects,
                    observed: referenced_paths.len() as u64 + 1,
                    limit: MAX_MEDIA_REFERENCE_PATHS as u64,
                    scope: std::sync::Arc::from("ODP media reference index"),
                },
            ));
        }
        let path_bytes = referenced_path_bytes
            .checked_add(path.len())
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(
                    "ODP media reference path size overflow".to_string(),
                )
            })?;
        if path_bytes > MAX_MEDIA_REFERENCE_PATH_BYTES {
            return Err(litchi_core::Error::ResourceLimit(
                litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::InputBytes,
                    observed: path_bytes as u64,
                    limit: MAX_MEDIA_REFERENCE_PATH_BYTES as u64,
                    scope: std::sync::Arc::from("ODP media reference paths"),
                },
            ));
        }
        referenced_paths
            .try_reserve(1)
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODP media reference index",
                source,
            })?;
        *referenced_path_bytes = path_bytes;
        referenced_paths.insert(path);
    }
    Ok(())
}

fn resolve_owner_relative_reference(owner_path: &str, value: &str) -> Result<Option<String>> {
    let value = value.trim();
    let target_end = value.find(['?', '#']).unwrap_or(value.len());
    let target = &value[..target_end];
    if target.is_empty() || is_linked_href(target) {
        return Ok(None);
    }
    let base = owner_relative_base(owner_path);
    let joined = if base.is_empty() {
        target.to_string()
    } else {
        format!("{base}/{target}")
    };
    let resolved = resolve_package_path(&joined).map_err(|error| {
        litchi_core::Error::Unsupported(format!(
            "ODP media reference '{value}' in '{owner_path}' is not a safe package path: {error}"
        ))
    })?;
    if is_linked_href(&resolved) {
        return Ok(None);
    }
    Ok(Some(resolved))
}

fn owner_relative_base(owner_path: &str) -> String {
    let source_path = if let Some((prefix, rels_name)) = owner_path.rsplit_once("/_rels/") {
        match strip_ascii_suffix(rels_name, ".rels") {
            Some(source_name) if prefix.is_empty() => source_name.to_string(),
            Some(source_name) => format!("{prefix}/{source_name}"),
            None => owner_path.to_string(),
        }
    } else if let Some(rels_name) = owner_path.strip_prefix("_rels/") {
        strip_ascii_suffix(rels_name, ".rels")
            .unwrap_or(rels_name)
            .to_string()
    } else {
        owner_path.to_string()
    };
    source_path
        .rsplit_once('/')
        .map_or_else(String::new, |(directory, _)| directory.to_string())
}

fn strip_ascii_suffix<'a>(value: &'a str, suffix: &str) -> Option<&'a str> {
    value
        .len()
        .checked_sub(suffix.len())
        .filter(|&start| value[start..].eq_ignore_ascii_case(suffix))
        .map(|start| &value[..start])
}

fn dependency_free_copy_name(old_name: &str, names: &[String]) -> Result<String> {
    let used = names.iter().map(String::as_str).collect::<BTreeSet<_>>();
    for ordinal in 1..=65_536usize {
        let suffix = if ordinal == 1 {
            " Copy".to_string()
        } else {
            format!(" Copy {ordinal}")
        };
        let candidate_len = old_name.len().checked_add(suffix.len()).ok_or_else(|| {
            litchi_core::Error::InvalidFormat(
                "ODP dependency-free copied page name size overflow".to_string(),
            )
        })?;
        if candidate_len > MAX_DEPENDENCY_FREE_COPY_NAME_BYTES {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP dependency-free copied page name exceeds {MAX_DEPENDENCY_FREE_COPY_NAME_BYTES} bytes"
            )));
        }
        let mut candidate = String::new();
        candidate
            .try_reserve_exact(candidate_len)
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODP dependency-free copied page name",
                source,
            })?;
        candidate.push_str(old_name);
        candidate.push_str(&suffix);
        if !used.contains(candidate.as_str()) {
            return Ok(candidate);
        }
    }
    Err(litchi_core::Error::InvalidFormat(
        "ODP dependency-free copied page name space is exhausted".to_string(),
    ))
}

fn foreign_copy_name(source_name: &str, names: &[String]) -> Result<String> {
    if !names.iter().any(|name| name == source_name) {
        let mut name = String::new();
        name.try_reserve_exact(source_name.len())
            .map_err(|source| litchi_core::Error::Allocation {
                resource: "ODP foreign copied page name",
                source,
            })?;
        name.push_str(source_name);
        return Ok(name);
    }
    dependency_free_copy_name(source_name, names)
}

fn dependency_free_blank_name_value(page: &str) -> Result<std::ops::Range<usize>> {
    let mut reader = Reader::from_str(page);
    reader.config_mut().check_end_names = true;
    let event = reader.read_event().map_err(|error| {
        litchi_core::Error::InvalidFormat(format!(
            "invalid ODP dependency-free blank page fragment: {error}"
        ))
    })?;
    let Event::Empty(element) = event else {
        return Err(litchi_core::Error::Unsupported(
            "ODP dependency-free blank-slide copy requires a self-closing page with no children"
                .to_string(),
        ));
    };
    if element.name().as_ref() != b"draw:page" {
        return Err(litchi_core::Error::Unsupported(
            "ODP dependency-free blank-slide copy requires the canonical draw:page prefix"
                .to_string(),
        ));
    }
    let mut saw_name = false;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            litchi_core::Error::InvalidFormat(format!(
                "invalid ODP dependency-free blank page attribute: {error}"
            ))
        })?;
        let key = attribute.key.as_ref();
        if key == b"draw:name" {
            if saw_name {
                return Err(litchi_core::Error::InvalidFormat(
                    "duplicate draw:name on ODP dependency-free blank page".to_string(),
                ));
            }
            saw_name = true;
        } else if key != b"xmlns" && !key.starts_with(b"xmlns:") {
            return Err(litchi_core::Error::Unsupported(format!(
                "ODP dependency-free blank-slide copy refuses attribute '{}'",
                String::from_utf8_lossy(key)
            )));
        }
    }
    if !saw_name {
        return Err(litchi_core::Error::Unsupported(
            "ODP dependency-free blank-slide copy requires draw:name".to_string(),
        ));
    }
    if !matches!(reader.read_event(), Ok(Event::Eof)) {
        return Err(litchi_core::Error::Unsupported(
            "ODP dependency-free blank-slide copy refuses content outside the page tag".to_string(),
        ));
    }
    locate_attribute_value(page, "draw:name").ok_or_else(|| {
        litchi_core::Error::InvalidFormat(
            "cannot locate draw:name in ODP dependency-free blank page".to_string(),
        )
    })
}

fn locate_attribute_value(tag: &str, wanted: &str) -> Option<std::ops::Range<usize>> {
    let bytes = tag.as_bytes();
    let mut index = b"<draw:page".len();
    while index < bytes.len() {
        while bytes.get(index).is_some_and(u8::is_ascii_whitespace) {
            index += 1;
        }
        if bytes.get(index..index + 2) == Some(b"/>") {
            return None;
        }
        let name_start = index;
        while bytes.get(index).is_some_and(|byte| {
            !byte.is_ascii_whitespace() && *byte != b'=' && *byte != b'/' && *byte != b'>'
        }) {
            index += 1;
        }
        let name = tag.get(name_start..index)?;
        while bytes.get(index).is_some_and(u8::is_ascii_whitespace) {
            index += 1;
        }
        if bytes.get(index) != Some(&b'=') {
            return None;
        }
        index += 1;
        while bytes.get(index).is_some_and(u8::is_ascii_whitespace) {
            index += 1;
        }
        let quote = *bytes.get(index)?;
        if quote != b'\'' && quote != b'"' {
            return None;
        }
        index += 1;
        let value_start = index;
        while bytes.get(index) != Some(&quote) {
            index += 1;
            if index >= bytes.len() {
                return None;
            }
        }
        let value_end = index;
        index += 1;
        if name == wanted {
            return Some(value_start..value_end);
        }
    }
    None
}

/// Compare two slides ignoring their position in the deck.
fn slide_content_eq(left: &Slide, right: &Slide) -> bool {
    let Slide {
        title,
        text,
        index: _,
        notes,
        transition,
        animations,
        legacy_animation,
        shapes,
    } = left;
    *title == right.title
        && *text == right.text
        && *notes == right.notes
        && *transition == right.transition
        && *animations == right.animations
        && *legacy_animation == right.legacy_animation
        && *shapes == right.shapes
}

#[cfg(test)]
mod tests {
    use super::MutablePresentation;
    use crate::{Builder, Presentation};
    use litchi_core::Result;

    fn presentation_with_two_slides() -> Result<Presentation> {
        let mut builder = Builder::new();
        builder.add_slide_with_title("First", "First body")?;
        builder.add_slide_with_title("Second", "Second body")?;
        Presentation::from_bytes(builder.build()?)
    }

    #[test]
    fn validated_slide_projection_is_cloned_into_isolated_draft_state() -> Result<()> {
        let presentation = presentation_with_two_slides()?;
        let validated_slides = presentation.slides()?;
        let mut draft = MutablePresentation::from_presentation_with_validated_slides(
            &presentation,
            &validated_slides,
        )?;

        assert_eq!(draft.slides, validated_slides);
        assert_eq!(draft.source_slides, validated_slides);
        assert!(draft.retains_source_slide(0));

        draft.update_slide(0, "Changed", "Changed body")?;
        assert_eq!(validated_slides[0].title.as_deref(), Some("First"));
        assert_eq!(draft.source_slides, validated_slides);
        assert!(!draft.retains_source_slide(0));
        Ok(())
    }

    #[test]
    fn validated_slide_projection_refuses_incomplete_source_page_coverage() -> Result<()> {
        let presentation = presentation_with_two_slides()?;
        let validated_slides = presentation.slides()?;
        let error = match MutablePresentation::from_presentation_with_validated_slides(
            &presentation,
            &validated_slides[..1],
        ) {
            Ok(_) => panic!("incomplete source slide coverage was accepted"),
            Err(error) => error,
        };

        assert!(error.to_string().contains("coverage is incomplete"));
        Ok(())
    }
}
