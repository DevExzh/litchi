//! Typed native archive construction shared by drawable title and caption editors.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use prost::Message;

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::protobuf::{tsa, tsd, tsp, tss, tswp};
use crate::text::style_registry::{insert_private_style, register_private_style};
use crate::wire::{patch_length_delimited_field, transform_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

pub(crate) const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
pub(crate) const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
pub(crate) const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
pub(crate) const STORAGE_MESSAGE_TYPE: u32 = 2_001;
pub(crate) const STANDIN_CAPTION_MESSAGE_TYPE: u32 = 3_097;

const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_CAPTION_MESSAGE_VERSION: [u32; 3] = [10, 1, 0];
const COMPONENTIZED_CAPTION_MESSAGE_VERSION: [u32; 3] = [10, 1, 0];
const CAPTION_DRAWABLE_TEXT_WRAP_FIELD_PATH: [u32; 4] = [1, 1, 1, 12];
const CAPTION_OWNED_STORAGE_FIELD_PATH: [u32; 2] = [1, 6];
const CAPTION_DRAWABLE_LOCKED_FIELD_PATH: [u32; 4] = [1, 1, 1, 13];
const DRAWABLE_TITLE_FIELD: u32 = 10;
const DRAWABLE_CAPTION_FIELD: u32 = 11;
const SHAPE_INFO_SUPER_FIELD: u32 = 1;
const SHAPE_ARCHIVE_DRAWABLE_FIELD: u32 = 1;
const DEFAULT_CAPTION_GEOMETRY_FLAGS: u32 = 1;
const DEFAULT_CAPTION_ROTATION_DEGREES: f32 = 0.0;
const DEFAULT_TEXT_WRAP_MARGIN_POINTS: f32 = 12.0;
const DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD: f32 = 0.5;
const DEFAULT_CAPTION_STYLE_OVERRIDE_COUNT: u32 = 7;
const DEFAULT_CAPTION_STYLE_PADDING_POINTS: f32 = 4.0;
const DEFAULT_STROKE_WIDTH_POINTS: f32 = 1.0;
const DEFAULT_STROKE_MITER_LIMIT: f32 = 4.0;
const DEFAULT_SHADOW_ANGLE_DEGREES: f32 = 315.0;
const DEFAULT_SHADOW_OFFSET_POINTS: f32 = 5.0;
const DEFAULT_SHADOW_RADIUS_POINTS: i32 = 1;
const DEFAULT_SHADOW_OPACITY: f32 = 1.0;
const NORMALIZED_PATH_EXTENT: f32 = 100.0;
const DEFAULT_STROKE_PATTERN_ELEMENT_COUNT: usize = 6;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum DrawableCaptionKind {
    Caption,
    Title,
}

impl DrawableCaptionKind {
    pub(crate) const fn drawable_field(self) -> u32 {
        match self {
            Self::Caption => DRAWABLE_CAPTION_FIELD,
            Self::Title => DRAWABLE_TITLE_FIELD,
        }
    }

    pub(crate) const fn native_kind(self) -> i32 {
        match self {
            Self::Caption => tsd::CaptionOrTitleKind::Caption as i32,
            Self::Title => tsd::CaptionOrTitleKind::Title as i32,
        }
    }

    const fn style_identifier(self) -> &'static str {
        match self {
            Self::Caption => "captions-0-shapestyle-Object Caption",
            Self::Title => "captions-0-shapestyle-Object Title",
        }
    }

    const fn placement(self, profile: CaptionGraphProfile) -> tsa::CaptionPlacementArchive {
        match (self, profile) {
            (Self::Caption, CaptionGraphProfile::Inline) => tsa::CaptionPlacementArchive {
                caption_anchor_location: Some(CaptionAnchorLocation::Bottom as i32),
                drawable_anchor_location: Some(CaptionAnchorLocation::Top as i32),
            },
            (Self::Title, CaptionGraphProfile::Inline) => tsa::CaptionPlacementArchive {
                caption_anchor_location: Some(CaptionAnchorLocation::Top as i32),
                drawable_anchor_location: Some(CaptionAnchorLocation::Bottom as i32),
            },
            (Self::Caption, CaptionGraphProfile::Componentized) => tsa::CaptionPlacementArchive {
                caption_anchor_location: Some(CaptionAnchorLocation::Top as i32),
                drawable_anchor_location: Some(CaptionAnchorLocation::Bottom as i32),
            },
            (Self::Title, CaptionGraphProfile::Componentized) => tsa::CaptionPlacementArchive {
                caption_anchor_location: Some(CaptionAnchorLocation::Bottom as i32),
                drawable_anchor_location: Some(CaptionAnchorLocation::Top as i32),
            },
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CaptionGraphProfile {
    Inline,
    Componentized,
}

/// Existing theme objects needed to materialize a native drawable title or caption.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CaptionThemeStyle {
    pub(crate) stylesheet_id: u64,
    pub(crate) paragraph_style_id: u64,
}

/// Fresh identifiers for one native title or caption graph.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CaptionObjectIds {
    pub(crate) style: u64,
    pub(crate) info: u64,
    pub(crate) storage: u64,
    pub(crate) placement: u64,
}

/// The private object graph behind a drawable's native title or caption.
///
/// Applications use a stand-in object until a label is created. Once present,
/// the reference instead targets a `TSA.CaptionInfoArchive` with private text,
/// placement, and sometimes style objects. A shared style can live outside the
/// drawable component, so it is deliberately omitted from [`Self::object_ids`]
/// in that case.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DrawableCaptionSlot {
    pub(crate) reference_id: u64,
    pub(crate) storage_id: Option<u64>,
    pub(crate) object_ids: Vec<u64>,
}

/// Read-only archive state for one drawable title/caption graph resolution.
///
/// The package cache intentionally retains only one parsed component. Keep a
/// small borrowed object-location index for this operation, while retaining
/// the component that owns the drawable so all private graph checks can borrow
/// its already-parsed [`Archive`] without reparsing it.
struct DrawableCaptionArchives<'package> {
    object_archives: HashMap<u64, &'package str>,
    duplicate_objects: HashSet<u64>,
    drawable_archive_name: &'package str,
    drawable_archive: Arc<Archive>,
}

impl<'package> DrawableCaptionArchives<'package> {
    fn discover(package: &'package IWorkPackage, drawable_object_id: u64) -> Result<Self> {
        let mut object_archives = HashMap::new();
        let mut duplicate_objects = HashSet::new();
        let mut drawable_archive = None;

        for archive_name in package.iwa_entry_names() {
            let archive = package.parsed_archive(archive_name)?;
            if archive.object(drawable_object_id).is_some()
                && drawable_archive
                    .replace((archive_name, Arc::clone(&archive)))
                    .is_some()
            {
                return Err(Error::Archive(format!(
                    "object {drawable_object_id} occurs in multiple IWA components"
                )));
            }
            for object in &archive.objects {
                let Some(identifier) = object.archive_info.identifier else {
                    continue;
                };
                if let Some(previous_archive_name) =
                    object_archives.insert(identifier, archive_name)
                    && previous_archive_name != archive_name
                {
                    duplicate_objects.insert(identifier);
                }
            }
        }

        let (drawable_archive_name, drawable_archive) = drawable_archive.ok_or_else(|| {
            Error::InvalidFormat(format!("object {drawable_object_id} is missing"))
        })?;
        Ok(Self {
            object_archives,
            duplicate_objects,
            drawable_archive_name,
            drawable_archive,
        })
    }

    fn archive_name(&self, identifier: u64) -> Result<&'package str> {
        if self.duplicate_objects.contains(&identifier) {
            return Err(Error::Archive(format!(
                "object {identifier} occurs in multiple IWA components"
            )));
        }
        self.object_archives
            .get(&identifier)
            .copied()
            .ok_or_else(|| Error::InvalidFormat(format!("object {identifier} is missing")))
    }

    fn drawable_archive(&self) -> &Archive {
        self.drawable_archive.as_ref()
    }

    fn require_exact_message_count(
        &self,
        package: &IWorkPackage,
        object_id: u64,
        message_type: u32,
        label: &str,
        drawable_label: &str,
    ) -> Result<()> {
        let archive_name = self.archive_name(object_id)?;
        if archive_name == self.drawable_archive_name {
            return require_exact_caption_message_count_in_archive(
                self.drawable_archive(),
                object_id,
                message_type,
                label,
                drawable_label,
            );
        }
        package.with_parsed_archive(archive_name, |archive| {
            require_exact_caption_message_count_in_archive(
                archive,
                object_id,
                message_type,
                label,
                drawable_label,
            )
        })
    }
}

/// Resolve and validate a drawable title or caption graph without assuming a
/// particular iWork application or drawable type.
#[allow(deprecated)]
pub(crate) fn drawable_caption_slot(
    package: &IWorkPackage,
    drawable_object_id: u64,
    reference: Option<&tsp::Reference>,
    kind: DrawableCaptionKind,
    drawable_label: &str,
) -> Result<DrawableCaptionSlot> {
    let reference_id = reference
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} {drawable_object_id} has no native title/caption reference"
            ))
        })?;
    let archives = DrawableCaptionArchives::discover(package, drawable_object_id)?;
    let drawable_archive_name = archives.archive_name(drawable_object_id)?;
    let archive_name = archives.archive_name(reference_id)?;
    if archive_name != drawable_archive_name {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} title/caption object {reference_id} is outside drawable {drawable_object_id}'s component"
        )));
    }
    let archive = archives.drawable_archive();
    let object = archive.object(reference_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} title/caption object {reference_id} is missing"
        ))
    })?;
    let caption_messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == CAPTION_INFO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if caption_messages.is_empty() {
        archives.require_exact_message_count(
            package,
            reference_id,
            STANDIN_CAPTION_MESSAGE_TYPE,
            "stand-in caption",
            drawable_label,
        )?;
        return Ok(DrawableCaptionSlot {
            reference_id,
            storage_id: None,
            object_ids: vec![reference_id],
        });
    }
    let [message] = caption_messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} title/caption object {reference_id} has multiple CaptionInfo payloads"
        )));
    };
    let info = tsa::CaptionInfoArchive::decode(message.data.as_slice())?;
    if info.child_info_kind != Some(kind.native_kind()) {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} title/caption object {reference_id} has the wrong native kind"
        )));
    }
    if info
        .super_
        .super_
        .super_
        .parent
        .as_ref()
        .map(|parent| parent.identifier)
        != Some(drawable_object_id)
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} title/caption object {reference_id} has the wrong parent drawable"
        )));
    }
    let storage_id = required_caption_reference(
        reference_id,
        info.super_.owned_storage.as_ref(),
        "text storage",
        drawable_label,
    )?;
    if info
        .super_
        .deprecated_storage
        .as_ref()
        .map(|storage| storage.identifier)
        != Some(storage_id)
        || info.super_.is_text_box != Some(true)
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} title/caption object {reference_id} has inconsistent text storage"
        )));
    }
    let style_id = required_caption_reference(
        reference_id,
        info.super_.super_.style.as_ref(),
        "shape style",
        drawable_label,
    )?;
    let placement_id = required_caption_reference(
        reference_id,
        info.placement.as_ref(),
        "placement",
        drawable_label,
    )?;
    for (identifier, message_type, label) in [
        (style_id, SHAPE_STYLE_MESSAGE_TYPE, "shape style"),
        (storage_id, STORAGE_MESSAGE_TYPE, "text storage"),
        (placement_id, CAPTION_PLACEMENT_MESSAGE_TYPE, "placement"),
    ] {
        archives.require_exact_message_count(
            package,
            identifier,
            message_type,
            label,
            drawable_label,
        )?;
    }
    for (identifier, label) in [(storage_id, "text storage"), (placement_id, "placement")] {
        if archives.archive_name(identifier)? != drawable_archive_name {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} title/caption {label} {identifier} is outside drawable {drawable_object_id}'s component"
            )));
        }
    }
    let mut object_ids = vec![reference_id];
    if archives.archive_name(style_id)? == drawable_archive_name {
        object_ids.push(style_id);
    }
    object_ids.extend([storage_id, placement_id]);
    if object_ids.iter().copied().collect::<HashSet<_>>().len() != object_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} title/caption object {reference_id} aliases its private graph"
        )));
    }
    Ok(DrawableCaptionSlot {
        reference_id,
        storage_id: Some(storage_id),
        object_ids,
    })
}

impl CaptionObjectIds {
    pub(crate) fn allocate(first: u64) -> Result<Self> {
        let identifier = |offset: u64| {
            first
                .checked_add(offset)
                .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))
        };
        Ok(Self {
            style: identifier(0)?,
            info: identifier(1)?,
            storage: identifier(2)?,
            placement: identifier(3)?,
        })
    }

    pub(crate) const fn last(self) -> u64 {
        self.placement
    }

    pub(crate) const fn all(self) -> [u64; 4] {
        [self.style, self.info, self.storage, self.placement]
    }

    pub(crate) const fn component_objects(self) -> [u64; 3] {
        [self.info, self.storage, self.placement]
    }
}

/// Create all four native objects that back one drawable title or caption.
#[allow(deprecated)]
pub(crate) fn caption_objects(
    ids: CaptionObjectIds,
    drawable_id: u64,
    drawable_width: f32,
    text: &str,
    kind: DrawableCaptionKind,
    theme: CaptionThemeStyle,
    language: Option<&str>,
) -> Result<[ArchiveObject; 4]> {
    caption_objects_with_profile(
        ids,
        drawable_id,
        drawable_width,
        text,
        kind,
        theme,
        language,
        CaptionGraphProfile::Inline,
    )
}

/// Create the native cross-component graph used by current Numbers files.
pub(crate) fn componentized_caption_objects(
    ids: CaptionObjectIds,
    drawable_id: u64,
    drawable_width: f32,
    text: &str,
    kind: DrawableCaptionKind,
    theme: CaptionThemeStyle,
    language: Option<&str>,
) -> Result<[ArchiveObject; 4]> {
    caption_objects_with_profile(
        ids,
        drawable_id,
        drawable_width,
        text,
        kind,
        theme,
        language,
        CaptionGraphProfile::Componentized,
    )
}

#[allow(deprecated, clippy::too_many_arguments)]
fn caption_objects_with_profile(
    ids: CaptionObjectIds,
    drawable_id: u64,
    drawable_width: f32,
    text: &str,
    kind: DrawableCaptionKind,
    theme: CaptionThemeStyle,
    language: Option<&str>,
    profile: CaptionGraphProfile,
) -> Result<[ArchiveObject; 4]> {
    if !drawable_width.is_finite() || drawable_width <= 0.0 {
        return Err(Error::InvalidFormat(
            "native drawable title/caption width must be finite and greater than zero".to_owned(),
        ));
    }
    let style = caption_shape_style(theme, kind);
    let storage = caption_storage(text, theme, language);
    let info = tsa::CaptionInfoArchive {
        super_: tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    geometry: Some(tsd::GeometryArchive {
                        position: Some(tsp::Point { x: 0.0, y: 0.0 }),
                        size: Some(tsp::Size {
                            width: drawable_width,
                            height: 0.0,
                        }),
                        flags: Some(DEFAULT_CAPTION_GEOMETRY_FLAGS),
                        angle: Some(DEFAULT_CAPTION_ROTATION_DEGREES),
                    }),
                    parent: Some(reference(drawable_id)),
                    exterior_text_wrap: Some(tsd::ExteriorTextWrapArchive {
                        r#type: Some(TextWrapType::Square as u32),
                        direction: Some(TextWrapDirection::BothSides as u32),
                        fit_type: Some(TextWrapFit::Text as u32),
                        margin: Some(DEFAULT_TEXT_WRAP_MARGIN_POINTS),
                        alpha_threshold: Some(DEFAULT_TEXT_WRAP_ALPHA_THRESHOLD),
                        is_html_wrap: Some(false),
                    }),
                    locked: Some(false),
                    aspect_ratio_locked: Some(false),
                    title_hidden: Some(false),
                    caption_hidden: Some(false),
                    ..Default::default()
                },
                style: Some(reference(ids.style)),
                pathsource: Some(caption_path_source(drawable_width)),
                stroke_pattern_offset_distance: Some(0.0),
                ..Default::default()
            },
            deprecated_storage: Some(reference(ids.storage)),
            owned_storage: Some(reference(ids.storage)),
            is_text_box: Some(true),
            ..Default::default()
        },
        placement: Some(reference(ids.placement)),
        child_info_kind: Some(kind.native_kind()),
    };
    Ok([
        archive_object(
            ids.style,
            SHAPE_STYLE_MESSAGE_TYPE,
            style,
            &[theme.paragraph_style_id],
            profile,
        )?,
        archive_object(
            ids.info,
            CAPTION_INFO_MESSAGE_TYPE,
            info,
            &[ids.style, ids.storage, ids.placement],
            profile,
        )?,
        archive_object(
            ids.storage,
            STORAGE_MESSAGE_TYPE,
            storage,
            &[theme.paragraph_style_id],
            profile,
        )?,
        archive_object(
            ids.placement,
            CAPTION_PLACEMENT_MESSAGE_TYPE,
            kind.placement(profile),
            &[],
            profile,
        )?,
    ])
}

/// Insert a componentized title/caption style into the document stylesheet and
/// expose it to the drawable component.
pub(crate) fn insert_componentized_caption_style(
    package: &mut IWorkPackage,
    drawable_archive_name: &str,
    stylesheet_id: u64,
    style: ArchiveObject,
) -> Result<()> {
    let style_id = style
        .archive_info
        .identifier
        .ok_or_else(|| Error::InvalidFormat("native caption style has no identifier".to_owned()))?;
    let stylesheet_archive_name = unique_object_archive_name(package, stylesheet_id)?;
    insert_private_style(
        package,
        &stylesheet_archive_name,
        stylesheet_id,
        style_id,
        style,
    )?;
    register_private_style(
        package,
        drawable_archive_name,
        &stylesheet_archive_name,
        style_id,
    )
}

/// Create a fresh empty stand-in when a native title or caption is removed.
pub(crate) fn standin_caption_object(identifier: u64) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: STANDIN_CAPTION_MESSAGE_TYPE,
            data: tsd::StandinCaptionArchive::default().encode_to_vec(),
        }],
    )?;
    object.archive_info.message_infos[0].versions = STANDIN_CAPTION_MESSAGE_VERSION.to_vec();
    Ok(object)
}

/// Rewrite just one reference inside a drawable while retaining all unknown fields.
pub(crate) fn patch_drawable_caption_reference(
    data: &[u8],
    kind: DrawableCaptionKind,
    replacement: u64,
) -> Result<Vec<u8>> {
    let drawable = tsd::DrawableArchive::decode(data)?;
    let was_present = match kind {
        DrawableCaptionKind::Caption => drawable.caption.is_some(),
        DrawableCaptionKind::Title => drawable.title.is_some(),
    };
    let replacement = reference(replacement).encode_to_vec();
    let data =
        patch_length_delimited_field(data, kind.drawable_field(), was_present, Some(&replacement))?;
    let verified = tsd::DrawableArchive::decode(data.as_slice())?;
    let actual = match kind {
        DrawableCaptionKind::Caption => verified.caption,
        DrawableCaptionKind::Title => verified.title,
    }
    .map(|reference| reference.identifier);
    if actual != Some(tsp::Reference::decode(replacement.as_slice())?.identifier) {
        return Err(Error::InvalidFormat(
            "native drawable title/caption reference patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

/// Rewrite one title or caption reference inside a `TSWP.ShapeInfoArchive`
/// without dropping unknown protobuf fields.
pub(crate) fn patch_shape_info_caption_reference(
    data: &[u8],
    kind: DrawableCaptionKind,
    replacement: u64,
) -> Result<Vec<u8>> {
    let data = transform_length_delimited_field(data, SHAPE_INFO_SUPER_FIELD, |shape| {
        transform_length_delimited_field(shape, SHAPE_ARCHIVE_DRAWABLE_FIELD, |drawable| {
            patch_drawable_caption_reference(drawable, kind, replacement)
        })
    })?;
    let verified = tswp::ShapeInfoArchive::decode(data.as_slice())?;
    let actual = match kind {
        DrawableCaptionKind::Caption => verified.super_.super_.caption,
        DrawableCaptionKind::Title => verified.super_.super_.title,
    }
    .map(|reference| reference.identifier);
    if actual != Some(replacement) {
        return Err(Error::InvalidFormat(
            "native shape title/caption reference patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}

/// Replace one archive metadata reference without perturbing unrelated ordering.
pub(crate) fn replace_object_reference(references: &mut Vec<u64>, old: u64, new: u64) {
    if let Some(index) = references.iter().position(|identifier| *identifier == old) {
        references[index] = new;
        let mut seen_new = false;
        references.retain(|identifier| {
            if *identifier != new {
                return true;
            }
            let keep = !seen_new;
            seen_new = true;
            keep
        });
    } else if !references.contains(&new) {
        references.push(new);
    }
}

fn required_caption_reference(
    reference_id: u64,
    reference: Option<&tsp::Reference>,
    label: &str,
    drawable_label: &str,
) -> Result<u64> {
    reference
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} title/caption object {reference_id} has no {label} reference"
            ))
        })
}

fn require_exact_caption_message_count_in_archive(
    archive: &Archive,
    object_id: u64,
    message_type: u32,
    label: &str,
    drawable_label: &str,
) -> Result<()> {
    let object = archive.object(object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} title/caption {label} {object_id} is missing"
        ))
    })?;
    if object
        .messages
        .iter()
        .filter(|message| message.type_ == message_type)
        .count()
        != 1
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} title/caption {label} {object_id} must have exactly one expected payload"
        )));
    }
    Ok(())
}

fn unique_object_archive_name(package: &IWorkPackage, identifier: u64) -> Result<String> {
    let mut archive_name = None;
    for name in package.iwa_entry_names() {
        if package.archive(name)?.object(identifier).is_none() {
            continue;
        }
        if archive_name.replace(name.to_owned()).is_some() {
            return Err(Error::Archive(format!(
                "object {identifier} occurs in multiple IWA components"
            )));
        }
    }
    archive_name.ok_or_else(|| Error::InvalidFormat(format!("object {identifier} is missing")))
}

fn caption_shape_style(
    theme: CaptionThemeStyle,
    kind: DrawableCaptionKind,
) -> tswp::ShapeStyleArchive {
    tswp::ShapeStyleArchive {
        super_: tsd::ShapeStyleArchive {
            super_: tss::StyleArchive {
                style_identifier: Some(kind.style_identifier().to_owned()),
                stylesheet: Some(reference(theme.stylesheet_id)),
                ..Default::default()
            },
            override_count: Some(DEFAULT_CAPTION_STYLE_OVERRIDE_COUNT),
            shape_properties: Some(tsd::ShapeStylePropertiesArchive {
                fill: Some(transparent_fill()),
                stroke: Some(empty_stroke()),
                opacity: Some(1.0),
                shadow: Some(disabled_shadow()),
                reflection: Some(tsd::ReflectionArchive::default()),
                ..Default::default()
            }),
        },
        override_count: Some(DEFAULT_CAPTION_STYLE_OVERRIDE_COUNT),
        shape_properties: Some(tswp::ShapeStylePropertiesArchive {
            padding: Some(tswp::PaddingArchive {
                left: Some(DEFAULT_CAPTION_STYLE_PADDING_POINTS),
                top: Some(DEFAULT_CAPTION_STYLE_PADDING_POINTS),
                right: Some(DEFAULT_CAPTION_STYLE_PADDING_POINTS),
                bottom: Some(DEFAULT_CAPTION_STYLE_PADDING_POINTS),
            }),
            paragraph_style: Some(reference(theme.paragraph_style_id)),
            ..Default::default()
        }),
    }
}

fn caption_storage(
    text: &str,
    theme: CaptionThemeStyle,
    language: Option<&str>,
) -> tswp::StorageArchive {
    tswp::StorageArchive {
        style_sheet: Some(reference(theme.stylesheet_id)),
        text: vec![text.to_owned()],
        in_document: Some(true),
        table_para_style: Some(object_attribute_table(Some(reference(
            theme.paragraph_style_id,
        )))),
        table_para_data: Some(zero_para_data()),
        table_para_starts: Some(zero_para_data()),
        table_language: language.map(|language| tswp::StringAttributeTable {
            entries: vec![tswp::string_attribute_table::StringAttribute {
                character_index: 0,
                object: Some(language.to_owned()),
            }],
        }),
        table_para_bidi: Some(zero_para_data()),
        table_drop_cap_style: Some(object_attribute_table(None)),
        ..Default::default()
    }
}

fn caption_path_source(width: f32) -> tsd::PathSourceArchive {
    use tsp::path::ElementType;

    let point = |x, y| tsp::Point { x, y };
    let element = |r#type, points| tsp::path::Element {
        r#type: r#type as i32,
        points,
    };
    tsd::PathSourceArchive {
        horizontal_flip: Some(false),
        vertical_flip: Some(false),
        bezier_path_source: Some(tsd::BezierPathSourceArchive {
            natural_size: Some(tsp::Size { width, height: 0.0 }),
            path: Some(tsp::Path {
                elements: vec![
                    element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
                    element(
                        ElementType::LineTo,
                        vec![point(NORMALIZED_PATH_EXTENT, 0.0)],
                    ),
                    element(
                        ElementType::LineTo,
                        vec![point(NORMALIZED_PATH_EXTENT, NORMALIZED_PATH_EXTENT)],
                    ),
                    element(
                        ElementType::LineTo,
                        vec![point(0.0, NORMALIZED_PATH_EXTENT)],
                    ),
                    element(ElementType::CloseSubpath, Vec::new()),
                    element(ElementType::MoveTo, vec![point(0.0, 0.0)]),
                ],
            }),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn transparent_fill() -> tsd::FillArchive {
    tsd::FillArchive {
        color: Some(color(0.0)),
        ..Default::default()
    }
}

fn empty_stroke() -> tsd::StrokeArchive {
    tsd::StrokeArchive {
        color: Some(color(1.0)),
        width: Some(DEFAULT_STROKE_WIDTH_POINTS),
        cap: Some(tsd::stroke_archive::LineCap::ButtCap as i32),
        join: Some(tsd::LineJoin::MiterJoin as i32),
        miter_limit: Some(DEFAULT_STROKE_MITER_LIMIT),
        pattern: Some(tsd::StrokePatternArchive {
            r#type: Some(tsd::stroke_pattern_archive::StrokePatternType::TsdEmptyPattern as i32),
            phase: Some(0.0),
            count: Some(0),
            pattern: vec![0.0; DEFAULT_STROKE_PATTERN_ELEMENT_COUNT],
        }),
        ..Default::default()
    }
}

fn disabled_shadow() -> tsd::ShadowArchive {
    tsd::ShadowArchive {
        color: Some(color(1.0)),
        angle: Some(DEFAULT_SHADOW_ANGLE_DEGREES),
        offset: Some(DEFAULT_SHADOW_OFFSET_POINTS),
        radius: Some(DEFAULT_SHADOW_RADIUS_POINTS),
        opacity: Some(DEFAULT_SHADOW_OPACITY),
        is_enabled: Some(false),
        r#type: Some(tsd::shadow_archive::ShadowType::TsdDropShadow as i32),
        ..Default::default()
    }
}

fn color(alpha: f32) -> tsp::Color {
    tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(0.0),
        g: Some(0.0),
        b: Some(0.0),
        rgbspace: Some(tsp::color::RgbColorSpace::Srgb as i32),
        a: Some(alpha),
        ..Default::default()
    }
}

fn object_attribute_table(object: Option<tsp::Reference>) -> tswp::ObjectAttributeTable {
    tswp::ObjectAttributeTable {
        entries: vec![tswp::object_attribute_table::ObjectAttribute {
            character_index: 0,
            object,
        }],
    }
}

fn zero_para_data() -> tswp::ParaDataAttributeTable {
    tswp::ParaDataAttributeTable {
        entries: vec![tswp::para_data_attribute_table::ParaDataAttribute {
            character_index: 0,
            first: 0,
            second: 0,
        }],
    }
}

fn archive_object(
    identifier: u64,
    message_type: u32,
    message: impl Message,
    object_references: &[u64],
    profile: CaptionGraphProfile,
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: message.encode_to_vec(),
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = if profile == CaptionGraphProfile::Componentized
        && matches!(
            message_type,
            CAPTION_INFO_MESSAGE_TYPE | CAPTION_PLACEMENT_MESSAGE_TYPE
        ) {
        COMPONENTIZED_CAPTION_MESSAGE_VERSION.to_vec()
    } else {
        STANDARD_MESSAGE_VERSION.to_vec()
    };
    info.object_references = object_references.to_vec();
    if profile == CaptionGraphProfile::Componentized && message_type == CAPTION_INFO_MESSAGE_TYPE {
        info.field_infos = [
            CAPTION_DRAWABLE_TEXT_WRAP_FIELD_PATH.as_slice(),
            CAPTION_OWNED_STORAGE_FIELD_PATH.as_slice(),
            CAPTION_DRAWABLE_LOCKED_FIELD_PATH.as_slice(),
        ]
        .into_iter()
        .map(preserved_caption_field)
        .collect();
    }
    Ok(object)
}

fn preserved_caption_field(path: &[u32]) -> tsp::FieldInfo {
    tsp::FieldInfo {
        path: tsp::FieldPath {
            path: path.to_vec(),
        },
        unknown_field_rule: Some(tsp::field_info::UnknownFieldRule::IgnoreAndPreserve as i32),
        ..Default::default()
    }
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum TextWrapType {
    Square = 4,
}

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum TextWrapDirection {
    BothSides = 2,
}

#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum TextWrapFit {
    Text = 1,
}

/// Native `CaptionPlacementArchive` anchors used by title and caption children.
#[derive(Debug, Clone, Copy)]
#[repr(i32)]
enum CaptionAnchorLocation {
    Top = 1,
    Bottom = 7,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn componentized_caption_matches_current_numbers_metadata() {
        let ids = CaptionObjectIds {
            style: 351,
            info: 352,
            storage: 353,
            placement: 354,
        };
        let [style, info, storage, placement] = componentized_caption_objects(
            ids,
            149,
            480.0,
            "Native caption",
            DrawableCaptionKind::Caption,
            CaptionThemeStyle {
                stylesheet_id: 5,
                paragraph_style_id: 41,
            },
            Some("en"),
        )
        .expect("componentized caption graph");

        assert_eq!(ids.component_objects(), [352, 353, 354]);
        assert_eq!(
            style.archive_info.message_infos[0].versions,
            STANDARD_MESSAGE_VERSION
        );
        assert_eq!(
            storage.archive_info.message_infos[0].versions,
            STANDARD_MESSAGE_VERSION
        );

        let info_metadata = &info.archive_info.message_infos[0];
        assert_eq!(
            info_metadata.versions,
            COMPONENTIZED_CAPTION_MESSAGE_VERSION
        );
        assert_eq!(info_metadata.object_references, [351, 353, 354]);
        assert_eq!(
            info_metadata
                .field_infos
                .iter()
                .map(|field| field.path.path.as_slice())
                .collect::<Vec<_>>(),
            [
                CAPTION_DRAWABLE_TEXT_WRAP_FIELD_PATH.as_slice(),
                CAPTION_OWNED_STORAGE_FIELD_PATH.as_slice(),
                CAPTION_DRAWABLE_LOCKED_FIELD_PATH.as_slice(),
            ]
        );
        assert!(info_metadata.field_infos.iter().all(|field| {
            field.unknown_field_rule
                == Some(tsp::field_info::UnknownFieldRule::IgnoreAndPreserve as i32)
        }));

        let placement_metadata = &placement.archive_info.message_infos[0];
        assert_eq!(
            placement_metadata.versions,
            COMPONENTIZED_CAPTION_MESSAGE_VERSION
        );
        let decoded = tsa::CaptionPlacementArchive::decode(placement.messages[0].data.as_slice())
            .expect("caption placement");
        assert_eq!(
            decoded.caption_anchor_location,
            Some(CaptionAnchorLocation::Top as i32)
        );
        assert_eq!(
            decoded.drawable_anchor_location,
            Some(CaptionAnchorLocation::Bottom as i32)
        );
    }
}
