//! Package-owned worksheet drawing and SVG reads.
//!
//! The worksheet facade resolves the direct SpreadsheetML drawing owner,
//! scans its source bytes with [`crate::drawing::SourceDrawing`], and parses
//! the same member through the typed drawing inventory.  A selected picture
//! therefore carries both views without making source ranges or relationship
//! IDs part of the selector API.  Media payloads are read only by the
//! explicit `read_*` methods and are returned as borrowed package bytes.

use std::collections::HashMap;
use std::str;
use std::sync::Arc;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    OpcPackage, PackURI, Part, PartMetadata, Relationship, Relationships, TargetMode,
};

use crate::drawing::source::{
    self, ContentPartProfile, ContentPartSource, PictureSource, SourceDrawing, SvgOwner,
    SvgOwnerState,
};
use crate::drawing::worksheet_source::{WorksheetSourceLimits, WorksheetSourceScan};
use crate::drawing::{
    ContentPartSelector, Drawing, DrawingAnchor, Picture, PictureSelector, UnknownKind,
};
use crate::error::{Error, Result, allocation, invalid};

use super::model::{Inner, Workbook, Worksheet, WorksheetKind};

const MAX_OUTPUT_MEDIA_BYTES: usize = 32 * 1024 * 1024;
const MAX_CONTENT_PART_PAYLOAD_BYTES: usize = 32 * 1024 * 1024;
const MAX_CONTENT_PART_RELATIONSHIPS: usize = 4_096;

const STRICT_CUSTOM_XML_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";

/// Package relationship metadata for one resolved core SpreadsheetDrawing
/// content part.
#[derive(Clone, Debug, PartialEq, Eq)]
#[must_use]
pub struct WorksheetContentPartRelationship {
    relationship_id: Box<str>,
    relationship_type: Box<str>,
    target_ref: Box<str>,
    target_mode: TargetMode,
    target_uri: PackURI,
    content_type: Box<str>,
}

impl WorksheetContentPartRelationship {
    /// Relationship ID authored by the drawing owner, retained for diagnostics.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Exact relationship type URI.
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        &self.relationship_type
    }

    /// Original relationship target reference.
    #[must_use]
    pub fn target_ref(&self) -> &str {
        &self.target_ref
    }

    /// Validated target mode. Core SpreadsheetDrawing content parts are
    /// always internal in this read profile.
    #[must_use]
    pub const fn target_mode(&self) -> TargetMode {
        self.target_mode
    }

    /// Canonical resolved OPC target part name.
    #[must_use]
    pub const fn target_uri(&self) -> &PackURI {
        &self.target_uri
    }

    /// Actual OPC content type declared for the target part.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }
}

/// One source-backed, semantically owned core SpreadsheetDrawing content part.
///
/// The primary target bytes are deferred until [`Self::read_payload`].  The
/// owner still validates the drawing relationship, target part, content type,
/// and bounded outbound relationship count before this view is returned.
pub struct WorksheetContentPart<'a> {
    owner: &'a Inner,
    source: Arc<SourceDrawing<'a>>,
    source_index: usize,
    relationship: WorksheetContentPartRelationship,
}

impl<'a> WorksheetContentPart<'a> {
    /// Source-preserving direct core content-part projection.
    pub fn source(&self) -> &ContentPartSource<'a> {
        self.source
            .content_parts()
            .get(self.source_index)
            .expect("worksheet content-part source index is validated at construction")
    }

    /// Recognized profile. Group extensions are not returned by this read
    /// owner and remain opaque source markup.
    #[must_use]
    pub fn profile(&self) -> ContentPartProfile {
        self.source().profile()
    }

    /// Complete direct anchor geometry containing this owner.
    #[must_use]
    pub fn anchor(&self) -> &DrawingAnchor {
        self.source().anchor()
    }

    /// Validated package relationship metadata.
    pub const fn relationship(&self) -> &WorksheetContentPartRelationship {
        &self.relationship
    }

    /// Borrow the exact owner element bytes from the paired drawing source.
    pub fn owner_bytes(&self) -> Result<&'a [u8]> {
        self.source().owner_bytes(self.source.source())
    }

    /// Read and validate the bounded XML payload while retaining a borrow into
    /// the immutable package part.
    pub fn read_payload(&self) -> Result<WorksheetContentPartPayload<'a>> {
        let part = self
            .owner
            .package
            .get_part(self.relationship.target_uri())?;
        let caller_limit = usize::try_from(self.owner.package.read_limits().max_part_bytes())
            .unwrap_or(usize::MAX);
        let maximum = MAX_CONTENT_PART_PAYLOAD_BYTES.min(caller_limit);
        if part.blob().len() > maximum {
            return Err(invalid(format!(
                "worksheet content-part payload exceeds {maximum} bytes"
            )));
        }
        validate_content_part_xml(part, self.owner.package.read_limits())?;
        let outbound_relationships = copy_outbound_relationships(
            part.rels(),
            self.owner
                .package
                .read_limits()
                .max_relationships_per_part(),
        )?;
        Ok(WorksheetContentPartPayload {
            bytes: part.blob(),
            outbound_relationships,
        })
    }

    /// Alias emphasizing that payload access is deferred and bounded.
    pub fn payload(&self) -> Result<WorksheetContentPartPayload<'a>> {
        self.read_payload()
    }
}

/// Bounded payload borrowed from a resolved core SpreadsheetDrawing content
/// part target.
#[derive(Debug)]
#[must_use]
pub struct WorksheetContentPartPayload<'a> {
    bytes: &'a [u8],
    outbound_relationships: Box<[WorksheetContentPartOutboundRelationship<'a>]>,
}

impl<'a> WorksheetContentPartPayload<'a> {
    /// Borrow exact target XML bytes without copying.
    #[must_use]
    pub const fn bytes(&self) -> &'a [u8] {
        self.bytes
    }

    /// Bounded relationship metadata authored by the target part.
    pub fn outbound_relationships(&self) -> &[WorksheetContentPartOutboundRelationship<'a>] {
        &self.outbound_relationships
    }
}

/// Opaque outbound relationship metadata retained for a content-part target.
/// The read slice records it but does not recursively follow or interpret the
/// target graph.  Its strings remain borrowed from the immutable OPC package.
#[derive(Debug)]
#[must_use]
pub struct WorksheetContentPartOutboundRelationship<'a> {
    relationship: &'a Relationship,
}

impl WorksheetContentPartOutboundRelationship<'_> {
    /// Relationship identifier in the target's `.rels` member.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        self.relationship.r_id()
    }

    /// Exact relationship type URI.
    #[must_use]
    pub fn relationship_type(&self) -> &str {
        self.relationship.reltype()
    }

    /// Original target reference.
    #[must_use]
    pub fn target_ref(&self) -> &str {
        self.relationship.target_ref()
    }

    /// Target mode retained without following the edge.
    #[must_use]
    pub fn target_mode(&self) -> TargetMode {
        self.relationship.target_mode()
    }
}

/// A worksheet drawing with paired source and typed inventories.
///
/// The drawing XML and its relationships remain borrowed from the immutable
/// workbook package.  The typed inventory is bounded semantic metadata; the
/// source inventory remains the authority for source ranges and extension
/// ownership.
pub struct WorksheetDrawing<'a> {
    owner: &'a Inner,
    part: &'a dyn Part,
    relationships: &'a Relationships,
    source: Arc<SourceDrawing<'a>>,
    typed: Arc<Drawing>,
    /// Object indexes for the typed picture projection.  `Drawing::pictures`
    /// is an iterator over a filtered object slice, so retaining these
    /// indexes avoids rescanning all preceding objects for every selector
    /// while keeping the complete typed `Drawing` as the only picture data
    /// store.
    typed_picture_indices: Box<[usize]>,
}

impl<'a> WorksheetDrawing<'a> {
    /// Source-preserving drawing inventory.
    pub fn source(&self) -> &SourceDrawing<'a> {
        self.source.as_ref()
    }

    /// Typed drawing inventory paired by source-order picture position.
    #[must_use]
    pub fn typed(&self) -> &Drawing {
        self.typed.as_ref()
    }

    /// Borrow the exact drawing XML member bytes.
    #[must_use]
    pub fn source_xml(&self) -> &'a [u8] {
        self.part.blob()
    }

    /// Number of direct pictures in this drawing.
    #[must_use]
    pub fn picture_count(&self) -> usize {
        self.source.pictures().len()
    }

    /// Select one direct picture by source order.
    pub fn picture(&self, position: usize) -> Result<WorksheetPicture<'a>> {
        let source = self.source.picture(position)?;
        let object_index = *self
            .typed_picture_indices
            .get(position)
            .ok_or_else(|| invalid("source picture has no matching typed picture"))?;
        let _typed = match self.typed.objects().get(object_index) {
            Some(crate::drawing::Object::Picture(picture)) => picture,
            _ => return Err(invalid("typed picture index does not point to a picture")),
        };
        // Validate the compatibility raster at selection time.  A typed
        // picture without a closed inert raster graph is still retained by
        // the ordinary drawing inventory, but cannot enter this SVG facade.
        let raster = validate_raster(self.owner, self.relationships, source)?;
        Ok(WorksheetPicture {
            owner: self.owner,
            drawing_part: self.part,
            source: Arc::clone(&self.source),
            source_picture_index: position,
            typed: Arc::clone(&self.typed),
            typed_picture_index: object_index,
            raster,
        })
    }

    /// Iterate source-order direct pictures, returning the first validation
    /// failure instead of silently projecting a malformed relationship graph.
    pub fn pictures(&self) -> Result<Vec<WorksheetPicture<'a>>> {
        let mut pictures = Vec::new();
        pictures
            .try_reserve_exact(self.picture_count())
            .map_err(|source| allocation("worksheet drawing pictures", source))?;
        for position in 0..self.picture_count() {
            pictures.push(self.picture(position)?);
        }
        Ok(pictures)
    }

    /// Number of direct core SpreadsheetDrawing content parts in this drawing.
    #[must_use]
    pub fn content_part_count(&self) -> usize {
        self.source.content_parts().len()
    }

    /// Select one direct core content part by source order and validate its
    /// owning drawing relationship before returning the view.
    pub fn content_part(&self, position: usize) -> Result<WorksheetContentPart<'a>> {
        let source = self.source.content_part(position)?;
        let relationship = self
            .relationships
            .get(source.relationship_id())
            .ok_or_else(|| invalid("worksheet content-part relationship is missing"))?;
        let relationship = validate_content_part_relationship(self.owner, source, relationship)?;
        Ok(WorksheetContentPart {
            owner: self.owner,
            source: Arc::clone(&self.source),
            source_index: position,
            relationship,
        })
    }

    /// Resolve every direct core content part in source order.
    pub fn content_parts(&self) -> Result<Vec<WorksheetContentPart<'a>>> {
        if self.content_part_count() == 0 {
            return Ok(Vec::new());
        }
        let targets = ContentPartTargetIndex::build(self.owner)?;
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(self.content_part_count())
            .map_err(|source| allocation("worksheet content parts", source))?;
        for position in 0..self.content_part_count() {
            parts.push(self.content_part_with_targets(position, &targets)?);
        }
        Ok(parts)
    }

    fn content_part_with_targets(
        &self,
        position: usize,
        targets: &ContentPartTargetIndex<'a>,
    ) -> Result<WorksheetContentPart<'a>> {
        let source = self.source.content_part(position)?;
        let relationship = self
            .relationships
            .get(source.relationship_id())
            .ok_or_else(|| invalid("worksheet content-part relationship is missing"))?;
        validate_content_part_relationship_shape(source, relationship)?;
        let target_uri = relationship.target_partname()?;
        let target = targets.get(&target_uri)?;
        let relationship =
            validate_content_part_relationship_target(self.owner, relationship, target)?;
        Ok(WorksheetContentPart {
            owner: self.owner,
            source: Arc::clone(&self.source),
            source_index: position,
            relationship,
        })
    }
}

/// One direct worksheet picture with paired source and typed projections.
pub struct WorksheetPicture<'a> {
    owner: &'a Inner,
    drawing_part: &'a dyn Part,
    source: Arc<SourceDrawing<'a>>,
    source_picture_index: usize,
    typed: Arc<Drawing>,
    typed_picture_index: usize,
    raster: WorksheetImageReference,
}

impl<'a> WorksheetPicture<'a> {
    /// Source-preserving direct-picture projection.
    pub fn source(&self) -> &PictureSource<'a> {
        self.source
            .pictures()
            .get(self.source_picture_index)
            .expect("worksheet picture source index is validated at construction")
    }

    /// Typed direct-picture projection.
    #[must_use]
    pub fn typed(&self) -> &Picture {
        match self.typed.objects().get(self.typed_picture_index) {
            Some(crate::drawing::Object::Picture(picture)) => picture,
            _ => panic!("worksheet typed picture index is validated at construction"),
        }
    }

    /// Complete typed anchor geometry.
    #[must_use]
    pub fn anchor(&self) -> &DrawingAnchor {
        self.typed().drawing_anchor()
    }

    /// Compatibility raster relationship and inert image target.
    pub const fn raster(&self) -> &WorksheetImageReference {
        &self.raster
    }

    /// Source SVG-owner status before package relationship validation.
    #[must_use]
    pub fn svg_owner(&self) -> &SvgOwnerState<'a> {
        self.source().svg_owner()
    }

    /// Return a validated embedded SVG descriptor, when this picture has one.
    ///
    /// Linked, ambiguous, and refused owners remain inert and return a typed
    /// error.  Opaque and absent extensions return `Ok(None)`.
    pub fn svg(&self) -> Result<Option<WorksheetSvgDescriptor<'a>>> {
        match self.source().svg_owner() {
            SvgOwnerState::None | SvgOwnerState::Opaque => Ok(None),
            SvgOwnerState::Embedded(owner) => Ok(Some(validate_svg(
                self.owner,
                self.drawing_part,
                Arc::clone(&self.source),
                self.source_picture_index,
                owner,
            )?)),
            SvgOwnerState::Linked(_) => {
                Err(invalid("linked SVG owners are inert in worksheet reads"))
            },
            SvgOwnerState::Ambiguous => Err(invalid(
                "worksheet picture has more than one admitted SVG owner",
            )),
            SvgOwnerState::Refused => Err(invalid("worksheet picture has a refused SVG owner")),
        }
    }

    /// Alias for [`Self::svg`] emphasizing descriptor-only access.
    pub fn svg_descriptor(&self) -> Result<Option<WorksheetSvgDescriptor<'a>>> {
        self.svg()
    }

    /// Read the selected embedded SVG payload without copying it.
    pub fn read_svg_image(&self) -> Result<WorksheetSvgImage<'a>> {
        let descriptor = self
            .svg()?
            .ok_or_else(|| invalid("worksheet picture has no embedded SVG owner"))?;
        let part = self.owner.package.get_part(&descriptor.part_uri)?;
        if part.blob().len() > MAX_OUTPUT_MEDIA_BYTES {
            return Err(invalid(
                "worksheet SVG payload exceeds the hard media limit",
            ));
        }
        Ok(WorksheetSvgImage {
            descriptor,
            bytes: part.blob(),
        })
    }

    /// Alias for [`Self::read_svg_image`].
    pub fn svg_image(&self) -> Result<WorksheetSvgImage<'a>> {
        self.read_svg_image()
    }
}

/// Relationship and package metadata for an inert raster compatibility image.
#[derive(Clone, Debug, PartialEq, Eq)]
#[must_use]
pub struct WorksheetImageReference {
    relationship_id: Box<str>,
    part_uri: PackURI,
    content_type: Box<str>,
}

impl WorksheetImageReference {
    /// Drawing relationship ID.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Canonical package part URI.
    #[must_use]
    pub const fn part_uri(&self) -> &PackURI {
        &self.part_uri
    }

    /// Declared package content type.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }
}

/// Validated direct SVG owner and its package target metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
#[must_use]
pub struct WorksheetSvgDescriptor<'a> {
    source: Arc<SourceDrawing<'a>>,
    source_picture_index: usize,
    relationship_id: Box<str>,
    part_uri: PackURI,
    content_type: Box<str>,
}

impl<'a> WorksheetSvgDescriptor<'a> {
    /// Source-backed typed SVG extension owner.
    pub fn owner(&self) -> &SvgOwner<'a> {
        match self
            .source
            .pictures()
            .get(self.source_picture_index)
            .expect("worksheet SVG source index is validated at construction")
            .svg_owner()
        {
            SvgOwnerState::Embedded(owner) | SvgOwnerState::Linked(owner) => owner,
            _ => panic!("worksheet SVG owner is validated at construction"),
        }
    }

    /// Shared typed `asvg:svgBlip` metadata.
    pub fn reference(&self) -> &litchi_drawingml::svg_blip::Reference {
        self.owner().value().reference()
    }

    /// SVG image relationship ID.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Canonical package part URI.
    #[must_use]
    pub const fn part_uri(&self) -> &PackURI {
        &self.part_uri
    }

    /// Exact package content type (`image/svg+xml`).
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }
}

/// One selected SVG payload borrowed from the immutable package part.
#[derive(Debug)]
#[must_use]
pub struct WorksheetSvgImage<'a> {
    descriptor: WorksheetSvgDescriptor<'a>,
    bytes: &'a [u8],
}

impl<'a> WorksheetSvgImage<'a> {
    /// Validated SVG descriptor.
    pub const fn descriptor(&self) -> &WorksheetSvgDescriptor<'a> {
        &self.descriptor
    }

    /// Borrow the exact package payload bytes.
    #[must_use]
    pub const fn bytes(&self) -> &'a [u8] {
        self.bytes
    }
}

impl Worksheet {
    /// Read one direct worksheet drawing by source-order drawing ordinal.
    pub fn drawing<'a>(&'a self, ordinal: usize) -> Result<WorksheetDrawing<'a>> {
        read_drawing(
            &self.owner,
            self.kind(),
            self.name(),
            self.part_uri(),
            ordinal,
        )
    }

    /// Read one direct worksheet picture by drawing and picture ordinals.
    pub fn picture<'a>(&'a self, selector: PictureSelector) -> Result<WorksheetPicture<'a>> {
        read_drawing(
            &self.owner,
            self.kind(),
            self.name(),
            self.part_uri(),
            selector.drawing,
        )?
        .picture(selector.picture)
    }

    /// Read one direct core worksheet content part by drawing and content-part
    /// source-order positions.
    pub fn content_part<'a>(
        &'a self,
        selector: ContentPartSelector,
    ) -> Result<WorksheetContentPart<'a>> {
        self.drawing(selector.drawing)?
            .content_part(selector.content_part)
    }

    /// Read every direct core content part in one worksheet drawing.
    pub fn content_parts<'a>(&'a self, drawing: usize) -> Result<Vec<WorksheetContentPart<'a>>> {
        self.drawing(drawing)?.content_parts()
    }

    /// Read one embedded SVG paired with a direct worksheet picture.
    pub fn read_svg_image<'a>(
        &'a self,
        selector: PictureSelector,
    ) -> Result<WorksheetSvgImage<'a>> {
        self.picture(selector)?.read_svg_image()
    }
}

impl Workbook {
    /// Read one direct worksheet picture selected by worksheet and picture
    /// selectors, keeping package ownership below the workbook facade.
    pub fn picture<'a, 'selector>(
        &'a self,
        worksheet: impl Into<super::Selector<'selector>>,
        picture: PictureSelector,
    ) -> Result<WorksheetPicture<'a>> {
        let worksheet = self
            .sheet(worksheet)?
            .ok_or_else(|| invalid("worksheet selector did not resolve"))?;
        read_drawing(
            &self.inner,
            worksheet.kind(),
            worksheet.name(),
            worksheet.part_uri(),
            picture.drawing,
        )?
        .picture(picture.picture)
    }

    /// Read one direct core worksheet content part by worksheet selector and
    /// drawing/content-part source-order positions.
    pub fn content_part<'a, 'selector>(
        &'a self,
        worksheet: impl Into<super::Selector<'selector>>,
        content_part: ContentPartSelector,
    ) -> Result<WorksheetContentPart<'a>> {
        let worksheet = self
            .sheet(worksheet)?
            .ok_or_else(|| invalid("worksheet selector did not resolve"))?;
        read_drawing(
            &self.inner,
            worksheet.kind(),
            worksheet.name(),
            worksheet.part_uri(),
            content_part.drawing,
        )?
        .content_part(content_part.content_part)
    }

    /// Read one embedded SVG paired with a direct worksheet picture.
    pub fn read_svg_image<'a, 'selector>(
        &'a self,
        worksheet: impl Into<super::Selector<'selector>>,
        picture: PictureSelector,
    ) -> Result<WorksheetSvgImage<'a>> {
        self.picture(worksheet, picture)?.read_svg_image()
    }
}

fn read_drawing<'a>(
    owner: &'a Inner,
    kind: WorksheetKind,
    sheet_name: &str,
    sheet_uri: &PackURI,
    ordinal: usize,
) -> Result<WorksheetDrawing<'a>> {
    if kind != WorksheetKind::Worksheet {
        return Err(Error::NotWorksheet {
            sheet: sheet_name.to_owned(),
        });
    }
    let worksheet_part = owner.package.get_part(sheet_uri)?;
    let worksheet_scan = WorksheetSourceScan::scan_with_limits(
        worksheet_part.blob(),
        worksheet_source_limits(owner),
    )?;
    let reference = worksheet_scan.drawing(ordinal)?;
    let relationship = worksheet_part
        .rels()
        .get(reference.relationship_id())
        .ok_or_else(|| invalid("worksheet drawing relationship is missing"))?;
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::DRAWING | rt::STRICT_DRAWING)
    {
        return Err(invalid("worksheet drawing relationship is unsupported"));
    }
    let drawing_uri = relationship.target_partname()?;
    let drawing_part = owner.package.get_part(&drawing_uri)?;
    if drawing_part.content_type() != ct::OFC_DRAWING {
        return Err(invalid(format!(
            "worksheet drawing part has content type '{}', expected '{}'",
            drawing_part.content_type(),
            ct::OFC_DRAWING
        )));
    }
    let source = Arc::new(SourceDrawing::scan_with_limits(
        drawing_part.blob(),
        ordinal,
        drawing_source_limits(owner),
    )?);
    let text = str::from_utf8(drawing_part.blob())
        .map_err(|error| invalid(format!("worksheet drawing XML is not UTF-8: {error}")))?;
    let typed = Arc::new(
        crate::drawing::parse_with_limits(text, &owner.package.read_limits())?
            .ok_or_else(|| invalid("worksheet drawing is missing its wsDr root"))?,
    );
    let mut typed_picture_indices = Vec::new();
    typed_picture_indices
        .try_reserve_exact(typed.pictures().count())
        .map_err(|source| allocation("worksheet typed picture indexes", source))?;
    for (index, object) in typed.objects().iter().enumerate() {
        if matches!(object, crate::drawing::Object::Picture(_)) {
            typed_picture_indices.push(index);
        }
    }
    if typed_picture_indices.len() != source.pictures().len() {
        return Err(invalid(
            "source and typed worksheet drawing picture inventories disagree",
        ));
    }
    for (position, source_picture) in source.pictures().iter().enumerate() {
        let object_index = typed_picture_indices[position];
        let typed_picture = match typed.objects().get(object_index) {
            Some(crate::drawing::Object::Picture(picture)) => picture,
            _ => return Err(invalid("typed picture index does not point to a picture")),
        };
        if typed_picture.relationship_id != source_picture.raster_relationship_id()
            || typed_picture.drawing_anchor() != source_picture.anchor()
        {
            return Err(invalid(
                "source picture does not match the typed drawing inventory",
            ));
        }
    }
    let mut typed_content_part_indices = Vec::new();
    typed_content_part_indices
        .try_reserve_exact(source.content_parts().len())
        .map_err(|source| allocation("worksheet typed content-part indexes", source))?;
    for (index, object) in typed.objects().iter().enumerate() {
        if matches!(
            object,
            crate::drawing::Object::Unknown(unknown)
                if unknown.kind == UnknownKind::ContentPart
        ) {
            typed_content_part_indices.push(index);
        }
    }
    if typed_content_part_indices.len() != source.content_parts().len() {
        return Err(invalid(
            "source and typed worksheet drawing content-part inventories disagree",
        ));
    }
    for (position, source_content_part) in source.content_parts().iter().enumerate() {
        let object_index = typed_content_part_indices[position];
        let typed_content_part = match typed.objects().get(object_index) {
            Some(crate::drawing::Object::Unknown(unknown))
                if unknown.kind == UnknownKind::ContentPart =>
            {
                unknown
            },
            _ => {
                return Err(invalid(
                    "typed content-part index does not point to a content part",
                ));
            },
        };
        if typed_content_part.drawing_anchor() != source_content_part.anchor() {
            return Err(invalid(
                "source content part does not match the typed drawing inventory",
            ));
        }
    }
    let relationships = drawing_part.rels();
    Ok(WorksheetDrawing {
        owner,
        part: drawing_part,
        relationships,
        source,
        typed,
        typed_picture_indices: typed_picture_indices.into_boxed_slice(),
    })
}

fn worksheet_source_limits(owner: &Inner) -> WorksheetSourceLimits {
    WorksheetSourceLimits::from_read_limits(owner.package.read_limits())
}

fn drawing_source_limits(owner: &Inner) -> source::ScanLimits {
    let caller = owner.package.read_limits();
    let defaults = source::ScanLimits::default();
    let part_bytes = usize::try_from(caller.max_part_bytes()).unwrap_or(usize::MAX);
    source::ScanLimits {
        max_xml_bytes: defaults.max_xml_bytes.min(part_bytes),
        max_nodes: defaults.max_nodes.min(caller.max_xml_events()),
        max_depth: defaults.max_depth.min(caller.max_xml_depth()),
        max_pictures: defaults
            .max_pictures
            .min(caller.max_relationships_per_part()),
        max_content_parts: defaults
            .max_content_parts
            .min(caller.max_relationships_per_part()),
        max_relationship_references: defaults
            .max_relationship_references
            .min(caller.max_relationships_per_part()),
        max_fragment_bytes: defaults.max_fragment_bytes.min(part_bytes),
    }
}

fn validate_raster(
    owner: &Inner,
    relationships: &Relationships,
    picture: &PictureSource<'_>,
) -> Result<WorksheetImageReference> {
    let relationship = relationships
        .get(picture.raster_relationship_id())
        .ok_or_else(|| invalid("worksheet picture raster relationship is missing"))?;
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
    {
        return Err(invalid(
            "worksheet picture raster relationship is unsupported",
        ));
    }
    let target_uri = relationship.target_partname()?;
    ensure_media_uri(&target_uri, "raster")?;
    let part = owner.package.get_part(&target_uri)?;
    let part_uri = clone_pack_uri(part.partname(), "worksheet raster part URI")?;
    ensure_media_uri(&part_uri, "raster")?;
    if part.content_type() != ct::PNG {
        return Err(invalid(
            "worksheet picture raster fallback must be image/png",
        ));
    }
    if !part.rels().is_empty() {
        return Err(invalid("worksheet picture raster fallback must be inert"));
    }
    Ok(WorksheetImageReference {
        relationship_id: boxed_str(
            picture.raster_relationship_id(),
            "worksheet raster relationship ID",
        )?,
        part_uri,
        content_type: boxed_str(part.content_type(), "worksheet raster content type")?,
    })
}

fn validate_svg<'a>(
    owner: &Inner,
    drawing_part: &dyn Part,
    source: Arc<SourceDrawing<'a>>,
    source_picture_index: usize,
    svg_owner: &SvgOwner<'a>,
) -> Result<WorksheetSvgDescriptor<'a>> {
    let relationship_id = svg_owner
        .embedded_relationship_id()
        .ok_or_else(|| invalid("worksheet SVG owner is not embedded"))?;
    let relationship = drawing_part
        .rels()
        .get(relationship_id)
        .ok_or_else(|| invalid("worksheet SVG relationship is missing"))?;
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE)
    {
        return Err(invalid("worksheet SVG relationship is unsupported"));
    }
    let target_uri = relationship.target_partname()?;
    ensure_media_uri(&target_uri, "SVG")?;
    let part = owner.package.get_part(&target_uri)?;
    let part_uri = clone_pack_uri(part.partname(), "worksheet SVG part URI")?;
    ensure_media_uri(&part_uri, "SVG")?;
    if part.content_type() != "image/svg+xml" {
        return Err(invalid(
            "worksheet SVG target must have exact content type image/svg+xml",
        ));
    }
    if !part.rels().is_empty() {
        return Err(invalid("worksheet SVG target must be inert"));
    }
    Ok(WorksheetSvgDescriptor {
        source,
        source_picture_index,
        relationship_id: boxed_str(relationship_id, "worksheet SVG relationship ID")?,
        part_uri,
        content_type: boxed_str(part.content_type(), "worksheet SVG content type")?,
    })
}

/// Operation-local case-insensitive OPC part-name index used by a batch
/// content-part read.  `OpcPackage::get_part` intentionally keeps a linear
/// fallback for one-off case variants; this bounded index prevents that
/// fallback from multiplying across many owners in one batch.
///
/// The index holds part metadata only. A lookup resolves the indexed exact
/// name through the fallible accessor, so a batch read surfaces the same
/// decode refusal as the one-owner read (ADR 0030) and still never takes the
/// linear fallback.
struct ContentPartTargetIndex<'a> {
    package: &'a OpcPackage,
    parts: HashMap<String, PartMetadata<'a>>,
}

impl<'a> ContentPartTargetIndex<'a> {
    fn build(owner: &'a Inner) -> Result<Self> {
        let part_count = owner.package.part_count();
        if part_count > owner.package.read_limits().max_parts() {
            return Err(invalid(
                "worksheet content-part target index exceeds the package part limit",
            ));
        }
        let mut parts = HashMap::new();
        parts
            .try_reserve(part_count)
            .map_err(|source| allocation("worksheet content-part target index", source))?;
        for part in owner.package.iter_parts() {
            let key = canonical_part_key(part.partname(), "worksheet content-part target index")?;
            if parts.insert(key, part).is_some() {
                return Err(invalid(
                    "worksheet content-part target index has equivalent part names",
                ));
            }
        }
        Ok(Self {
            package: &owner.package,
            parts,
        })
    }

    fn get(&self, requested: &PackURI) -> Result<&'a dyn Part> {
        let key = canonical_part_key(requested, "worksheet content-part target lookup")?;
        let part = self
            .parts
            .get(&key)
            .copied()
            .ok_or_else(|| invalid("worksheet content-part target part is missing"))?;
        Ok(self.package.get_part(part.partname())?)
    }
}

fn validate_content_part_relationship(
    owner: &Inner,
    owner_source: &ContentPartSource<'_>,
    relationship: &Relationship,
) -> Result<WorksheetContentPartRelationship> {
    validate_content_part_relationship_shape(owner_source, relationship)?;
    let target_uri = relationship.target_partname()?;
    let target = owner.package.get_part(&target_uri)?;
    validate_content_part_relationship_target(owner, relationship, target)
}

fn validate_content_part_relationship_shape(
    owner_source: &ContentPartSource<'_>,
    relationship: &Relationship,
) -> Result<()> {
    let expected_type = match owner_source.relationship_dialect() {
        source::RelationshipDialect::Transitional => rt::CUSTOM_XML,
        source::RelationshipDialect::Strict => STRICT_CUSTOM_XML_RELATIONSHIP,
    };
    if relationship.target_mode() != TargetMode::Internal {
        return Err(invalid(
            "worksheet core contentPart relationship must be internal",
        ));
    }
    if relationship.reltype() != expected_type {
        return Err(invalid(format!(
            "worksheet core contentPart relationship has type '{}', expected '{expected_type}'",
            relationship.reltype()
        )));
    }
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return Err(invalid(
            "worksheet core contentPart relationship cannot contain a query or fragment",
        ));
    }
    if relationship.r_id() != owner_source.relationship_id() {
        return Err(invalid(
            "worksheet contentPart relationship ID does not match its source owner",
        ));
    }
    Ok(())
}

fn validate_content_part_relationship_target(
    owner: &Inner,
    relationship: &Relationship,
    target: &dyn Part,
) -> Result<WorksheetContentPartRelationship> {
    let canonical_target_uri = target.partname();
    if !is_xml_content_type(target.content_type()) {
        return Err(invalid(format!(
            "worksheet contentPart target has unsupported non-XML content type '{}'",
            target.content_type()
        )));
    }
    let max_relationships = owner
        .package
        .read_limits()
        .max_relationships_per_part()
        .min(MAX_CONTENT_PART_RELATIONSHIPS);
    if target.rels().len() > max_relationships {
        return Err(invalid(format!(
            "worksheet contentPart target has {} outbound relationships, exceeding {max_relationships}",
            target.rels().len()
        )));
    }
    Ok(WorksheetContentPartRelationship {
        relationship_id: boxed_str(
            relationship.r_id(),
            "worksheet content-part relationship ID",
        )?,
        relationship_type: boxed_str(
            relationship.reltype(),
            "worksheet content-part relationship type",
        )?,
        target_ref: boxed_str(
            relationship.target_ref(),
            "worksheet content-part target reference",
        )?,
        target_mode: relationship.target_mode(),
        target_uri: clone_pack_uri(canonical_target_uri, "worksheet content-part target URI")?,
        content_type: boxed_str(
            target.content_type(),
            "worksheet content-part target content type",
        )?,
    })
}

fn copy_outbound_relationships<'a>(
    relationships: &'a Relationships,
    caller_limit: usize,
) -> Result<Box<[WorksheetContentPartOutboundRelationship<'a>]>> {
    let maximum = caller_limit.min(MAX_CONTENT_PART_RELATIONSHIPS);
    if relationships.len() > maximum {
        return Err(invalid(format!(
            "worksheet contentPart target has too many outbound relationships ({} > {maximum})",
            relationships.len()
        )));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(relationships.len())
        .map_err(|source| allocation("worksheet content-part outbound relationships", source))?;
    for relationship in relationships.iter() {
        output.push(WorksheetContentPartOutboundRelationship { relationship });
    }
    Ok(output.into_boxed_slice())
}

fn validate_content_part_xml(part: &dyn Part, limits: litchi_opc::ReadLimits) -> Result<()> {
    litchi_opc::validate_source_xml_bytes(part.partname(), part.blob(), limits).map_err(|error| {
        invalid(format!(
            "worksheet contentPart target XML is invalid: {error}"
        ))
    })
}

fn is_xml_content_type(content_type: &str) -> bool {
    let value = content_type.trim();
    value.eq_ignore_ascii_case("text/xml")
        || value.eq_ignore_ascii_case("application/xml")
        || value
            .get(value.len().saturating_sub(4)..)
            .is_some_and(|suffix| suffix.eq_ignore_ascii_case("+xml"))
}

fn ensure_media_uri(uri: &PackURI, kind: &str) -> Result<()> {
    let mut segments = uri.as_str().split('/');
    let valid_prefix = segments.next() == Some("")
        && segments
            .next()
            .is_some_and(|segment| segment.eq_ignore_ascii_case("xl"))
        && segments
            .next()
            .is_some_and(|segment| segment.eq_ignore_ascii_case("media"));
    let has_media_name = segments.next().is_some_and(|segment| !segment.is_empty());
    if !valid_prefix {
        return Err(invalid(format!(
            "worksheet {kind} target is outside canonical /xl/media/"
        )));
    }
    if !has_media_name {
        return Err(invalid(format!(
            "worksheet {kind} target has no media name"
        )));
    }
    Ok(())
}

fn boxed_str(value: &str, resource: &'static str) -> Result<Box<str>> {
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    copy.push_str(value);
    Ok(copy.into_boxed_str())
}

fn canonical_part_key(uri: &PackURI, resource: &'static str) -> Result<String> {
    let value = uri.as_str();
    let mut key = String::new();
    key.try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    for byte in value.bytes() {
        key.push(char::from(byte.to_ascii_lowercase()));
    }
    Ok(key)
}

fn clone_pack_uri(uri: &PackURI, resource: &'static str) -> Result<PackURI> {
    let mut copy = String::new();
    copy.try_reserve_exact(uri.as_str().len())
        .map_err(|source| allocation(resource, source))?;
    copy.push_str(uri.as_str());
    PackURI::new(copy).map_err(invalid)
}
