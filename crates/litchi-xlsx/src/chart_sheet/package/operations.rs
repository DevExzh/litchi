//! OPC load/store and relationship validation for chartsheet resources.

use super::super::{
    Chart, Conformance, State, parse_chartsheet, validate_chartsheet, write_chartsheet,
};
use super::codec::{
    DrawingChartKind, DrawingChartReference, Node, add_relationship_checked, add_resource, attr,
    bounded, collect_extension_relationship_ids, drawing_chart_references, escape,
    internal_relationship, leaf, new_uri, optional, parse_document, parse_state,
    require_content_type, require_workbook, required, required_child, root_conformance, staged_uri,
    validate_chart_companion_xml, validate_chart_ex_relationships, validate_chart_user_shapes_xml,
    validate_chart_xml, validate_id, validate_theme_override_xml, xml_error,
};
use super::model::{
    BackgroundImageContentType, BackgroundPicture, ChartCompanionResource,
    ChartEmbeddedPackageContentType, ChartEmbeddedPackageResource, ChartOutboundResource,
    ChartResource, ChartResourceKind, ChartThemeOverrideResource, ChartUserShapesResource,
    DrawingResource, Entry, ExtensionRelationship, ExtensionRelationshipTarget, ImageContentType,
    ImageResource, Package, PrinterSettings, VmlDrawingResource,
};
use super::{
    CHART_COLOR_STYLE_CT, CHART_COLOR_STYLE_REL, CHART_CT, CHART_EX_CT, CHART_EX_REL,
    CHART_STYLE_CT, CHART_STYLE_REL, CHART_USER_SHAPES_CT, CHARTSHEET_CT, CHARTSHEET_REL,
    DRAWING_CT, MAX_BACKGROUND_IMAGE_BYTES, MAX_CHART_BYTES, MAX_CHART_COLOR_STYLE_BYTES,
    MAX_CHART_DIRECT_IMAGES, MAX_CHART_EMBEDDED_PACKAGE_BYTES, MAX_CHART_EX_BYTES,
    MAX_CHART_RELATIONSHIPS, MAX_CHART_STYLE_BYTES, MAX_CHART_STYLE_PARTS, MAX_CHART_THEME_IMAGES,
    MAX_CHART_THEME_OVERRIDE_BYTES, MAX_CHART_USER_SHAPE_IMAGE_BYTES, MAX_CHART_USER_SHAPE_IMAGES,
    MAX_CHART_USER_SHAPES_BYTES, MAX_DRAWING_BYTES, MAX_EXTENSION_PAYLOAD_BYTES,
    MAX_EXTENSION_RELATIONSHIP_STRING_BYTES, MAX_EXTENSION_RELATIONSHIPS, MAX_VML_DRAWING_BYTES,
    MAX_XML_BYTES, STRICT_CHARTSHEET_REL, THEME_OVERRIDE_CT, VML_DRAWING_CT, invalid, limit,
};
use crate::Result;
use crate::package::printer_settings::{
    MAX_SETTINGS_BYTES, PRINTER_CT, PrinterSettingsResource, is_printer_relationship,
    validate_printer_settings_uri, validate_settings_bytes,
};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet, HashSet};

/// Loads one workbook-referenced chartsheet and validates its bounded leaf graph.
pub fn load_chartsheet(
    package: &OpcPackage,
    workbook_name: &PackURI,
    workbook_relationship_id: &str,
) -> Result<Package> {
    if package.rels().iter().any(|rel| {
        matches!(rel.reltype(), CHARTSHEET_REL | STRICT_CHARTSHEET_REL)
            || is_printer_relationship(rel.reltype())
    }) {
        return Err(invalid(
            "package root cannot source a chartsheet or Printer Settings relationship",
        ));
    }
    let workbook = package.get_part(workbook_name)?;
    require_workbook(workbook)?;
    let workbook_root = parse_document(workbook.blob(), MAX_XML_BYTES)?;
    let conformance = root_conformance(&workbook_root, "workbook")?;
    let relationship = internal_relationship(
        workbook,
        workbook_relationship_id,
        conformance.chartsheet_rel(),
    )?;
    let chartsheet_name = relationship.target_partname()?;
    if !chartsheet_name.as_str().starts_with("/xl/chartsheets/") {
        return Err(invalid("chartsheet target is outside /xl/chartsheets"));
    }
    let entry = workbook_entry(
        &workbook_root,
        conformance,
        workbook_relationship_id,
        chartsheet_name.to_string(),
    )?;
    let chartsheet_part = package.get_part(&chartsheet_name)?;
    require_content_type(chartsheet_part, CHARTSHEET_CT, "chartsheet")?;
    let (part_conformance, chartsheet) = parse_chartsheet(chartsheet_part.blob())?;
    if part_conformance != conformance {
        return Err(invalid("workbook and chartsheet conformance differ"));
    }
    let drawing_rel = internal_relationship(
        chartsheet_part,
        &chartsheet.drawing_relationship_id,
        conformance.drawing_rel(),
    )?;
    let drawing_name = drawing_rel.target_partname()?;
    if !drawing_name.as_str().starts_with("/xl/drawings/") {
        return Err(invalid("chartsheet drawing is outside /xl/drawings"));
    }
    let legacy_drawing = chartsheet
        .legacy_drawing_relationship_id
        .as_deref()
        .map(|id| load_vml_resource(package, chartsheet_part, id, conformance))
        .transpose()?;
    let legacy_header_footer_drawing = chartsheet
        .legacy_header_footer_drawing_relationship_id
        .as_deref()
        .map(|id| load_vml_resource(package, chartsheet_part, id, conformance))
        .transpose()?;
    let background_picture = if let Some(id) = &chartsheet.background_picture_relationship_id {
        let rel = internal_relationship(chartsheet_part, id, conformance.image_rel())?;
        let name = rel.target_partname()?;
        if !name.as_str().starts_with("/xl/media/") {
            return Err(invalid("chartsheet background image is outside /xl/media"));
        }
        let part = package.get_part(&name)?;
        let content_type = BackgroundImageContentType::parse(part.content_type())?;
        if part.blob().len() > MAX_BACKGROUND_IMAGE_BYTES {
            return Err(limit("background image bytes"));
        }
        if !part.rels().is_empty() {
            return Err(invalid(
                "chartsheet background image must be a relationship-free leaf",
            ));
        }
        Some(BackgroundPicture {
            relationship_id: id.clone(),
            part_name: name.to_string(),
            content_type,
            data: part.blob().to_vec(),
        })
    } else {
        None
    };
    let printer_settings = if let Some(id) = chartsheet
        .page_setup
        .as_ref()
        .and_then(|setup| setup.printer_settings_relationship_id.as_ref())
    {
        let rel = internal_relationship(chartsheet_part, id, conformance.printer_rel())?;
        let name = rel.target_partname()?;
        validate_printer_settings_uri(&name)?;
        let part = package.get_part(&name)?;
        require_content_type(part, PRINTER_CT, "Printer Settings")?;
        validate_settings_bytes(part.blob())?;
        if !part.rels().is_empty() {
            return Err(invalid(
                "chartsheet Printer Settings must be a relationship-free leaf",
            ));
        }
        Some(PrinterSettings {
            relationship_id: id.clone(),
            resource: PrinterSettingsResource {
                part_name: name.to_string(),
                data: part.blob().to_vec(),
            },
        })
    } else {
        None
    };
    let known_relationships = known_chartsheet_relationship_ids(&chartsheet);
    let extension_ids = extension_relationship_ids(&chartsheet, conformance)?;
    let mut extension_relationships =
        Vec::with_capacity(extension_ids.len().min(MAX_EXTENSION_RELATIONSHIPS));
    for id in extension_ids.difference(&known_relationships) {
        if extension_relationships.len() >= MAX_EXTENSION_RELATIONSHIPS {
            return Err(limit("extension relationship count"));
        }
        let relationship = chartsheet_part
            .rels()
            .get(id)
            .ok_or_else(|| invalid(format!("missing extension relationship '{id}'")))?;
        validate_extension_relationship_string(relationship.reltype(), "type")?;
        let target = if relationship.is_external() {
            validate_extension_relationship_string(relationship.target_ref(), "target")?;
            ExtensionRelationshipTarget::External {
                target: relationship.target_ref().to_owned(),
            }
        } else {
            let name = relationship.target_partname()?;
            ExtensionRelationshipTarget::Internal {
                part_name: name.to_string(),
            }
        };
        extension_relationships.push(ExtensionRelationship {
            relationship_id: id.clone(),
            relationship_type: relationship.reltype().to_owned(),
            target,
        });
    }
    let expected_relationships = known_relationships.len() + extension_relationships.len();
    if chartsheet_part.rels().iter().count() != expected_relationships {
        return Err(invalid(
            "bounded chartsheet has unsupported or unreferenced relationships",
        ));
    }
    let drawing_part = package.get_part(&drawing_name)?;
    require_content_type(drawing_part, DRAWING_CT, "drawing")?;
    if drawing_part.blob().len() > MAX_DRAWING_BYTES {
        return Err(limit("drawing bytes"));
    }
    let chart_references = drawing_chart_references(drawing_part.blob(), conformance)?;
    if drawing_part.rels().iter().count() != chart_references.len() {
        return Err(invalid(
            "bounded chartsheet drawing has unsupported or unreferenced relationships",
        ));
    }
    let mut charts = Vec::with_capacity(chart_references.len());
    let mut total = drawing_part.blob().len();
    for vml in [&legacy_drawing, &legacy_header_footer_drawing]
        .into_iter()
        .flatten()
    {
        add_resource(
            &mut total,
            vml.data.len(),
            MAX_VML_DRAWING_BYTES,
            "VML drawing bytes",
        )?;
    }
    if let Some(picture) = &background_picture {
        add_resource(
            &mut total,
            picture.data.len(),
            MAX_BACKGROUND_IMAGE_BYTES,
            "background image bytes",
        )?;
    }
    if let Some(settings) = &printer_settings {
        add_resource(
            &mut total,
            settings.resource.data.len(),
            MAX_SETTINGS_BYTES,
            "Printer Settings bytes",
        )?;
    }
    for reference in chart_references {
        charts.push(load_chart_resource(
            package,
            drawing_part,
            &reference,
            conformance,
            &mut total,
        )?);
    }
    Ok(Package {
        entry,
        chartsheet,
        drawing: DrawingResource {
            part_name: drawing_name.to_string(),
            content_type: drawing_part.content_type().to_owned(),
            data: drawing_part.blob().to_vec(),
            charts,
        },
        legacy_drawing,
        legacy_header_footer_drawing,
        background_picture,
        printer_settings,
        extension_relationships,
    })
}

pub(super) fn load_vml_resource(
    package: &OpcPackage,
    chartsheet_part: &dyn Part,
    id: &str,
    conformance: Conformance,
) -> Result<VmlDrawingResource> {
    let rel = internal_relationship(chartsheet_part, id, conformance.vml_drawing_rel())?;
    let name = rel.target_partname()?;
    if !name.as_str().starts_with("/xl/drawings/") || !name.as_str().ends_with(".vml") {
        return Err(invalid(
            "chartsheet VML drawing target is outside /xl/drawings or lacks .vml suffix",
        ));
    }
    let part = package.get_part(&name)?;
    require_content_type(part, VML_DRAWING_CT, "VML drawing")?;
    if part.blob().len() > MAX_VML_DRAWING_BYTES {
        return Err(limit("VML drawing bytes"));
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "chartsheet VML drawing must be a relationship-free leaf",
        ));
    }
    Ok(VmlDrawingResource {
        relationship_id: id.to_owned(),
        part_name: name.to_string(),
        content_type: part.content_type().to_owned(),
        data: part.blob().to_vec(),
    })
}

pub(super) fn load_chart_resource(
    package: &OpcPackage,
    drawing_part: &dyn Part,
    reference: &DrawingChartReference,
    conformance: Conformance,
    total: &mut usize,
) -> Result<ChartResource> {
    let relationship_type = match reference.kind {
        DrawingChartKind::Classic => conformance.chart_rel(),
        DrawingChartKind::Extended => CHART_EX_REL,
    };
    let relationship =
        internal_relationship(drawing_part, &reference.relationship_id, relationship_type)?;
    let name = relationship.target_partname()?;
    if !name.as_str().starts_with("/xl/charts/") || !name.as_str().ends_with(".xml") {
        return Err(invalid(
            "chart target is outside /xl/charts or lacks .xml suffix",
        ));
    }
    let part = package.get_part(&name)?;
    let kind = match reference.kind {
        DrawingChartKind::Classic => {
            require_content_type(part, CHART_CT, "chart")?;
            validate_chart_xml(part.blob(), conformance)?;
            if !part.rels().is_empty() {
                return Err(invalid(
                    "bounded classic chart must be a relationship-free leaf",
                ));
            }
            add_resource(total, part.blob().len(), MAX_CHART_BYTES, "chart bytes")?;
            ChartResourceKind::Classic
        },
        DrawingChartKind::Extended => {
            require_content_type(part, CHART_EX_CT, "chartEx")?;
            if part.rels().len() > MAX_CHART_RELATIONSHIPS {
                return Err(limit("chartEx relationship count"));
            }
            let references = validate_chart_ex_relationships(part.blob(), conformance)?;
            add_resource(
                total,
                part.blob().len(),
                MAX_CHART_EX_BYTES,
                "chartEx bytes",
            )?;
            let (mut styles, mut color_styles, mut user_shapes, mut outbound_resources) = (
                Vec::with_capacity(MAX_CHART_STYLE_PARTS.min(part.rels().len())),
                Vec::with_capacity(MAX_CHART_STYLE_PARTS.min(part.rels().len())),
                None,
                Vec::with_capacity(part.rels().len()),
            );
            let (mut theme_seen, mut package_seen) = (false, false);
            for relationship in part.rels().iter() {
                if relationship.reltype() == conformance.chart_user_shapes_rel() {
                    if user_shapes.is_some() {
                        return Err(invalid(
                            "chartEx has multiple chartUserShapes relationships",
                        ));
                    }
                    if relationship.is_external() {
                        return Err(invalid("external chartUserShapes relationship is rejected"));
                    }
                    user_shapes = Some(load_chart_user_shapes_resource(
                        package,
                        relationship,
                        conformance,
                        total,
                    )?);
                    continue;
                }
                if relationship.reltype() == conformance.image_rel() {
                    if outbound_resources
                        .iter()
                        .filter(|value| matches!(value, ChartOutboundResource::Image(_)))
                        .count()
                        >= MAX_CHART_DIRECT_IMAGES
                    {
                        return Err(limit("chartEx direct image count"));
                    }
                    outbound_resources.push(ChartOutboundResource::Image(
                        load_chart_image_resource(
                            package,
                            relationship,
                            total,
                            "chartEx direct image",
                        )?,
                    ));
                    continue;
                }
                if relationship.reltype() == conformance.theme_override_rel() {
                    if theme_seen {
                        return Err(invalid("chartEx has multiple themeOverride relationships"));
                    }
                    theme_seen = true;
                    outbound_resources.push(ChartOutboundResource::ThemeOverride(
                        load_chart_theme_override_resource(
                            package,
                            relationship,
                            conformance,
                            total,
                        )?,
                    ));
                    continue;
                }
                if relationship.reltype() == conformance.package_rel() {
                    if package_seen {
                        return Err(invalid(
                            "chartEx has multiple embedded package relationships",
                        ));
                    }
                    package_seen = true;
                    outbound_resources.push(ChartOutboundResource::EmbeddedPackage(
                        load_chart_embedded_package_resource(package, relationship, total)?,
                    ));
                    continue;
                }
                let (target, root, max_bytes, collection) = match relationship.reltype() {
                    CHART_STYLE_REL => (
                        CHART_STYLE_CT,
                        "chartStyle",
                        MAX_CHART_STYLE_BYTES,
                        &mut styles,
                    ),
                    CHART_COLOR_STYLE_REL => (
                        CHART_COLOR_STYLE_CT,
                        "colorStyle",
                        MAX_CHART_COLOR_STYLE_BYTES,
                        &mut color_styles,
                    ),
                    _ => {
                        return Err(invalid(
                            "chartEx has an unsupported or active outbound relationship",
                        ));
                    },
                };
                if collection.len() >= MAX_CHART_STYLE_PARTS {
                    return Err(limit("chart companion count"));
                }
                if relationship.is_external() {
                    return Err(invalid(
                        "external chartEx companion relationship is rejected",
                    ));
                }
                let companion_name = relationship.target_partname()?;
                if !companion_name.as_str().starts_with("/xl/charts/")
                    || !companion_name.as_str().ends_with(".xml")
                {
                    return Err(invalid(
                        "chart companion target is outside /xl/charts or lacks .xml suffix",
                    ));
                }
                let companion = package.get_part(&companion_name)?;
                require_content_type(companion, target, "chart companion")?;
                validate_chart_companion_xml(companion.blob(), root, max_bytes)?;
                if !companion.rels().is_empty() {
                    return Err(invalid("chart companion must be a relationship-free leaf"));
                }
                add_resource(
                    total,
                    companion.blob().len(),
                    max_bytes,
                    "chart companion bytes",
                )?;
                collection.push(ChartCompanionResource {
                    relationship_id: relationship.r_id().to_owned(),
                    part_name: companion_name.to_string(),
                    content_type: companion.content_type().to_owned(),
                    data: companion.blob().to_vec(),
                });
            }
            let image_ids = outbound_resources
                .iter()
                .filter_map(|value| match value {
                    ChartOutboundResource::Image(image) => Some(image.relationship_id.clone()),
                    _ => None,
                })
                .collect::<BTreeSet<_>>();
            let package_id = outbound_resources.iter().find_map(|value| match value {
                ChartOutboundResource::EmbeddedPackage(package) => {
                    Some(package.relationship_id.clone())
                },
                _ => None,
            });
            if image_ids != references.images || package_id != references.package {
                return Err(invalid(
                    "chartEx XML relationship references do not close over direct images and embedded package",
                ));
            }
            styles.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
            color_styles.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
            outbound_resources
                .sort_by(|left, right| left.relationship_id().cmp(right.relationship_id()));
            ChartResourceKind::Extended {
                styles,
                color_styles,
                user_shapes,
                outbound_resources,
            }
        },
    };
    Ok(ChartResource {
        relationship_id: reference.relationship_id.clone(),
        part_name: name.to_string(),
        content_type: part.content_type().to_owned(),
        data: part.blob().to_vec(),
        kind,
    })
}

pub(super) fn load_chart_user_shapes_resource(
    package: &OpcPackage,
    relationship: &litchi_opc::Relationship,
    conformance: Conformance,
    total: &mut usize,
) -> Result<ChartUserShapesResource> {
    let name = relationship.target_partname()?;
    if !name.as_str().starts_with("/xl/drawings/") || !name.as_str().ends_with(".xml") {
        return Err(invalid(
            "chartUserShapes target is outside /xl/drawings or lacks .xml suffix",
        ));
    }
    let part = package.get_part(&name)?;
    require_content_type(part, CHART_USER_SHAPES_CT, "chartUserShapes")?;
    let referenced = validate_chart_user_shapes_xml(part.blob(), conformance)?;
    add_resource(
        total,
        part.blob().len(),
        MAX_CHART_USER_SHAPES_BYTES,
        "chartUserShapes bytes",
    )?;
    if referenced.len() > MAX_CHART_USER_SHAPE_IMAGES {
        return Err(limit("chart user-shape image count"));
    }
    if part.rels().len() > MAX_CHART_USER_SHAPE_IMAGES {
        return Err(limit("chart user-shape relationship count"));
    }
    if part.rels().len() != referenced.len() {
        return Err(invalid(
            "chartUserShapes image relationships are missing or orphaned",
        ));
    }
    let mut images = Vec::with_capacity(referenced.len());
    for id in referenced {
        let image_relationship = internal_relationship(part, &id, conformance.image_rel())?;
        let image_name = image_relationship.target_partname()?;
        if !image_name.as_str().starts_with("/xl/media/") {
            return Err(invalid("chart user-shape image is outside /xl/media"));
        }
        let image = package.get_part(&image_name)?;
        let content_type = ImageContentType::parse(image.content_type())?;
        if !content_type.validates_part_name(image_name.as_str()) {
            return Err(invalid(
                "chart user-shape image suffix does not match its content type",
            ));
        }
        if !image.rels().is_empty() {
            return Err(invalid(
                "chart user-shape image must be a relationship-free leaf",
            ));
        }
        add_resource(
            total,
            image.blob().len(),
            MAX_CHART_USER_SHAPE_IMAGE_BYTES,
            "chart user-shape image bytes",
        )?;
        images.push(ImageResource {
            relationship_id: id,
            part_name: image_name.to_string(),
            content_type,
            data: image.blob().to_vec(),
        });
    }
    images.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
    Ok(ChartUserShapesResource {
        relationship_id: relationship.r_id().to_owned(),
        part_name: name.to_string(),
        content_type: part.content_type().to_owned(),
        data: part.blob().to_vec(),
        images,
    })
}

pub(super) fn load_chart_image_resource(
    package: &OpcPackage,
    relationship: &litchi_opc::Relationship,
    total: &mut usize,
    label: &str,
) -> Result<ImageResource> {
    if relationship.is_external() {
        return Err(invalid(format!(
            "external {label} relationship is rejected"
        )));
    }
    let name = relationship.target_partname()?;
    if !name.as_str().starts_with("/xl/media/") {
        return Err(invalid(format!("{label} is outside /xl/media")));
    }
    let part = package.get_part(&name)?;
    let content_type = ImageContentType::parse(part.content_type())?;
    if !content_type.validates_part_name(name.as_str()) {
        return Err(invalid(format!(
            "{label} suffix does not match its content type"
        )));
    }
    if !part.rels().is_empty() {
        return Err(invalid(format!("{label} must be a relationship-free leaf")));
    }
    add_resource(
        total,
        part.blob().len(),
        MAX_CHART_USER_SHAPE_IMAGE_BYTES,
        "chart image bytes",
    )?;
    Ok(ImageResource {
        relationship_id: relationship.r_id().to_owned(),
        part_name: name.to_string(),
        content_type,
        data: part.blob().to_vec(),
    })
}

pub(super) fn load_chart_theme_override_resource(
    package: &OpcPackage,
    relationship: &litchi_opc::Relationship,
    conformance: Conformance,
    total: &mut usize,
) -> Result<ChartThemeOverrideResource> {
    if relationship.is_external() {
        return Err(invalid("external themeOverride relationship is rejected"));
    }
    let name = relationship.target_partname()?;
    if !name.as_str().starts_with("/xl/theme/") || !name.as_str().ends_with(".xml") {
        return Err(invalid(
            "themeOverride target is outside /xl/theme or lacks .xml suffix",
        ));
    }
    let part = package.get_part(&name)?;
    require_content_type(part, THEME_OVERRIDE_CT, "themeOverride")?;
    let referenced = validate_theme_override_xml(part.blob(), conformance)?;
    add_resource(
        total,
        part.blob().len(),
        MAX_CHART_THEME_OVERRIDE_BYTES,
        "themeOverride bytes",
    )?;
    if referenced.len() > MAX_CHART_THEME_IMAGES {
        return Err(limit("themeOverride image count"));
    }
    if part.rels().len() > MAX_CHART_THEME_IMAGES {
        return Err(limit("themeOverride relationship count"));
    }
    if part.rels().len() != referenced.len() {
        return Err(invalid(
            "themeOverride image relationships are missing or orphaned",
        ));
    }
    let mut images = Vec::with_capacity(referenced.len());
    for id in referenced {
        let image_relationship = internal_relationship(part, &id, conformance.image_rel())?;
        images.push(load_chart_image_resource(
            package,
            image_relationship,
            total,
            "themeOverride image",
        )?);
    }
    images.sort_by(|left, right| left.relationship_id.cmp(&right.relationship_id));
    Ok(ChartThemeOverrideResource {
        relationship_id: relationship.r_id().to_owned(),
        part_name: name.to_string(),
        content_type: part.content_type().to_owned(),
        data: part.blob().to_vec(),
        images,
    })
}

pub(super) fn load_chart_embedded_package_resource(
    package: &OpcPackage,
    relationship: &litchi_opc::Relationship,
    total: &mut usize,
) -> Result<ChartEmbeddedPackageResource> {
    if relationship.is_external() {
        return Err(invalid(
            "external chartEx embedded-package relationship is rejected",
        ));
    }
    let name = relationship.target_partname()?;
    if !name.as_str().starts_with("/xl/embeddings/") {
        return Err(invalid(
            "chartEx embedded package is outside /xl/embeddings",
        ));
    }
    let part = package.get_part(&name)?;
    let content_type = ChartEmbeddedPackageContentType::parse(part.content_type())?;
    if !content_type.validates_part_name(name.as_str()) {
        return Err(invalid(
            "chartEx embedded-package suffix does not match its content type",
        ));
    }
    if !part.rels().is_empty() {
        return Err(invalid(
            "chartEx embedded package must be a relationship-free opaque leaf",
        ));
    }
    add_resource(
        total,
        part.blob().len(),
        MAX_CHART_EMBEDDED_PACKAGE_BYTES,
        "chartEx embedded-package bytes",
    )?;
    Ok(ChartEmbeddedPackageResource {
        relationship_id: relationship.r_id().to_owned(),
        part_name: name.to_string(),
        content_type,
        data: part.blob().to_vec(),
    })
}

/// Adds a preflighted chartsheet package graph and workbook sheet entry.
pub fn store_chartsheet(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: &Package,
    conformance: Conformance,
) -> Result<()> {
    validate_package_value(value, conformance)?;
    let mut staged = package.clone();
    store_chartsheet_inner(&mut staged, workbook_name, value, conformance)?;
    *package = staged;
    Ok(())
}

/// Validates a typed chartsheet package graph without mutating an OPC package.
///
/// This is the in-memory counterpart to [`store_chartsheet`]. It checks the
/// `CT_Chartsheet` shape and the relationship/part closure required by the
/// `SpreadsheetML` chartsheet profile (ISO/IEC 29500-1, §18.3.1 and §12.3),
/// including the drawing, chart, media, VML, Printer Settings, and extension
/// resources represented by [`Package`].
///
/// # Errors
///
/// Returns [`crate::Error::Invalid`] when the typed graph violates the
/// chartsheet schema, relationship vocabulary, resource bounds, or reference
/// closure rules.
pub fn validate_package(value: &Package, conformance: Conformance) -> Result<()> {
    validate_package_value(value, conformance)
}

pub(super) fn store_chartsheet_inner(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    value: &Package,
    conformance: Conformance,
) -> Result<()> {
    let workbook = package.get_part(workbook_name)?;
    require_workbook(workbook)?;
    let workbook_root = parse_document(workbook.blob(), MAX_XML_BYTES)?;
    if root_conformance(&workbook_root, "workbook")? != conformance {
        return Err(invalid("requested conformance does not match workbook"));
    }
    validate_new_entry(&workbook_root, conformance, &value.entry)?;
    if workbook
        .rels()
        .get(&value.entry.workbook_relationship_id)
        .is_some()
    {
        return Err(invalid("workbook relationship ID already exists"));
    }
    let chartsheet_uri = new_uri(package, &value.entry.part_name, "/xl/chartsheets/")?;
    let drawing_uri = new_uri(package, &value.drawing.part_name, "/xl/drawings/")?;
    let legacy_uri = value
        .legacy_drawing
        .as_ref()
        .map(|resource| new_uri(package, &resource.part_name, "/xl/drawings/"))
        .transpose()?;
    let legacy_hf_uri = value
        .legacy_header_footer_drawing
        .as_ref()
        .map(|resource| new_uri(package, &resource.part_name, "/xl/drawings/"))
        .transpose()?;
    let picture_uri = value
        .background_picture
        .as_ref()
        .map(|picture| new_uri(package, &picture.part_name, "/xl/media/"))
        .transpose()?;
    let printer_uri = value
        .printer_settings
        .as_ref()
        .map(|settings| -> Result<PackURI> {
            let uri = PackURI::new(&settings.resource.part_name).map_err(invalid)?;
            validate_printer_settings_uri(&uri)?;
            package.validate_new_part_name(&uri)?;
            Ok(uri)
        })
        .transpose()?;
    let mut chart_uris = BTreeMap::new();
    let mut companion_uris = BTreeMap::new();
    let mut user_shape_uris = BTreeMap::new();
    let mut user_shape_image_uris = BTreeMap::new();
    let mut outbound_uris = BTreeMap::new();
    let mut theme_image_uris = BTreeMap::new();
    for chart in &value.drawing.charts {
        chart_uris.insert(
            chart.relationship_id.clone(),
            new_uri(package, &chart.part_name, "/xl/charts/")?,
        );
        if let ChartResourceKind::Extended {
            styles,
            color_styles,
            user_shapes,
            outbound_resources,
        } = &chart.kind
        {
            for companion in styles.iter().chain(color_styles) {
                companion_uris.insert(
                    (
                        chart.relationship_id.clone(),
                        companion.relationship_id.clone(),
                    ),
                    new_uri(package, &companion.part_name, "/xl/charts/")?,
                );
            }
            if let Some(user_shapes) = user_shapes {
                user_shape_uris.insert(
                    chart.relationship_id.clone(),
                    new_uri(package, &user_shapes.part_name, "/xl/drawings/")?,
                );
                for image in &user_shapes.images {
                    user_shape_image_uris.insert(
                        (chart.relationship_id.clone(), image.relationship_id.clone()),
                        new_uri(package, &image.part_name, "/xl/media/")?,
                    );
                }
            }
            for resource in outbound_resources {
                let relationship_id = resource.relationship_id().to_owned();
                let (prefix, part_name) = match resource {
                    ChartOutboundResource::Image(image) => ("/xl/media/", image.part_name.as_str()),
                    ChartOutboundResource::ThemeOverride(theme) => {
                        ("/xl/theme/", theme.part_name.as_str())
                    },
                    ChartOutboundResource::EmbeddedPackage(embedded) => {
                        ("/xl/embeddings/", embedded.part_name.as_str())
                    },
                };
                outbound_uris.insert(
                    (chart.relationship_id.clone(), relationship_id.clone()),
                    new_uri(package, part_name, prefix)?,
                );
                if let ChartOutboundResource::ThemeOverride(theme) = resource {
                    for image in &theme.images {
                        theme_image_uris.insert(
                            (
                                chart.relationship_id.clone(),
                                relationship_id.clone(),
                                image.relationship_id.clone(),
                            ),
                            new_uri(package, &image.part_name, "/xl/media/")?,
                        );
                    }
                }
            }
        }
    }
    let updated_workbook = insert_workbook_entry(workbook.blob(), &value.entry, conformance)?;
    let chartsheet_xml = write_chartsheet(&value.chartsheet, conformance)?;
    package
        .get_part_mut(workbook_name)?
        .set_blob(updated_workbook);
    package.try_add_part(Box::new(BlobPart::new(
        chartsheet_uri.clone(),
        CHARTSHEET_CT.into(),
        chartsheet_xml,
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        drawing_uri.clone(),
        value.drawing.content_type.clone(),
        value.drawing.data.clone(),
    )))?;
    if let (Some(resource), Some(uri)) = (&value.legacy_drawing, &legacy_uri) {
        package.try_add_part(Box::new(BlobPart::new(
            uri.clone(),
            resource.content_type.clone(),
            resource.data.clone(),
        )))?;
    }
    if let (Some(resource), Some(uri)) = (&value.legacy_header_footer_drawing, &legacy_hf_uri) {
        package.try_add_part(Box::new(BlobPart::new(
            uri.clone(),
            resource.content_type.clone(),
            resource.data.clone(),
        )))?;
    }
    if let (Some(picture), Some(uri)) = (&value.background_picture, &picture_uri) {
        package.try_add_part(Box::new(BlobPart::new(
            uri.clone(),
            picture.content_type.as_str().into(),
            picture.data.clone(),
        )))?;
    }
    if let (Some(settings), Some(uri)) = (&value.printer_settings, &printer_uri) {
        package.try_add_part(Box::new(BlobPart::new(
            uri.clone(),
            PRINTER_CT.into(),
            settings.resource.data.clone(),
        )))?;
    }
    for chart in &value.drawing.charts {
        let chart_uri = staged_uri(&chart_uris, &chart.relationship_id, "chart")?;
        package.try_add_part(Box::new(BlobPart::new(
            chart_uri,
            chart.content_type.clone(),
            chart.data.clone(),
        )))?;
        if let ChartResourceKind::Extended {
            styles,
            color_styles,
            user_shapes,
            outbound_resources,
        } = &chart.kind
        {
            for companion in styles.iter().chain(color_styles) {
                let companion_uri = staged_uri(
                    &companion_uris,
                    &(
                        chart.relationship_id.clone(),
                        companion.relationship_id.clone(),
                    ),
                    "chart companion",
                )?;
                package.try_add_part(Box::new(BlobPart::new(
                    companion_uri,
                    companion.content_type.clone(),
                    companion.data.clone(),
                )))?;
            }
            if let Some(user_shapes) = user_shapes {
                let user_shape_uri =
                    staged_uri(&user_shape_uris, &chart.relationship_id, "chart user-shape")?;
                package.try_add_part(Box::new(BlobPart::new(
                    user_shape_uri,
                    user_shapes.content_type.clone(),
                    user_shapes.data.clone(),
                )))?;
                for image in &user_shapes.images {
                    let image_uri = staged_uri(
                        &user_shape_image_uris,
                        &(chart.relationship_id.clone(), image.relationship_id.clone()),
                        "chart user-shape image",
                    )?;
                    package.try_add_part(Box::new(BlobPart::new(
                        image_uri,
                        image.content_type.as_str().into(),
                        image.data.clone(),
                    )))?;
                }
            }
            for resource in outbound_resources {
                let key = (
                    chart.relationship_id.clone(),
                    resource.relationship_id().to_owned(),
                );
                let uri = staged_uri(&outbound_uris, &key, "chart outbound")?;
                match resource {
                    ChartOutboundResource::Image(image) => package.try_add_part(Box::new(
                        BlobPart::new(uri, image.content_type.as_str().into(), image.data.clone()),
                    ))?,
                    ChartOutboundResource::ThemeOverride(theme) => {
                        package.try_add_part(Box::new(BlobPart::new(
                            uri.clone(),
                            theme.content_type.clone(),
                            theme.data.clone(),
                        )))?;
                        for image in &theme.images {
                            let image_uri = staged_uri(
                                &theme_image_uris,
                                &(
                                    chart.relationship_id.clone(),
                                    theme.relationship_id.clone(),
                                    image.relationship_id.clone(),
                                ),
                                "themeOverride image",
                            )?;
                            package.try_add_part(Box::new(BlobPart::new(
                                image_uri,
                                image.content_type.as_str().into(),
                                image.data.clone(),
                            )))?;
                        }
                    },
                    ChartOutboundResource::EmbeddedPackage(embedded) => {
                        package.try_add_part(Box::new(BlobPart::new(
                            uri,
                            embedded.content_type.as_str().into(),
                            embedded.data.clone(),
                        )))?;
                    },
                }
            }
        }
    }
    add_relationship_checked(
        package,
        workbook_name,
        conformance.chartsheet_rel(),
        chartsheet_uri.relative_ref(workbook_name.base_uri()),
        value.entry.workbook_relationship_id.clone(),
        TargetMode::Internal,
    )?;
    add_relationship_checked(
        package,
        &chartsheet_uri,
        conformance.drawing_rel(),
        drawing_uri.relative_ref(chartsheet_uri.base_uri()),
        value.chartsheet.drawing_relationship_id.clone(),
        TargetMode::Internal,
    )?;
    if let (Some(resource), Some(uri)) = (&value.legacy_drawing, &legacy_uri) {
        add_relationship_checked(
            package,
            &chartsheet_uri,
            conformance.vml_drawing_rel(),
            uri.relative_ref(chartsheet_uri.base_uri()),
            resource.relationship_id.clone(),
            TargetMode::Internal,
        )?;
    }
    if let (Some(resource), Some(uri)) = (&value.legacy_header_footer_drawing, &legacy_hf_uri) {
        add_relationship_checked(
            package,
            &chartsheet_uri,
            conformance.vml_drawing_rel(),
            uri.relative_ref(chartsheet_uri.base_uri()),
            resource.relationship_id.clone(),
            TargetMode::Internal,
        )?;
    }
    if let (Some(picture), Some(uri)) = (&value.background_picture, &picture_uri) {
        add_relationship_checked(
            package,
            &chartsheet_uri,
            conformance.image_rel(),
            uri.relative_ref(chartsheet_uri.base_uri()),
            picture.relationship_id.clone(),
            TargetMode::Internal,
        )?;
    }
    if let (Some(settings), Some(uri)) = (&value.printer_settings, &printer_uri) {
        add_relationship_checked(
            package,
            &chartsheet_uri,
            conformance.printer_rel(),
            uri.relative_ref(chartsheet_uri.base_uri()),
            settings.relationship_id.clone(),
            TargetMode::Internal,
        )?;
    }
    for relationship in &value.extension_relationships {
        let (target, external) = match &relationship.target {
            ExtensionRelationshipTarget::Internal { part_name } => (
                PackURI::new(part_name)
                    .map_err(invalid)?
                    .relative_ref(chartsheet_uri.base_uri()),
                false,
            ),
            ExtensionRelationshipTarget::External { target } => (target.clone(), true),
        };
        add_relationship_checked(
            package,
            &chartsheet_uri,
            &relationship.relationship_type,
            target,
            relationship.relationship_id.clone(),
            if external {
                TargetMode::External
            } else {
                TargetMode::Internal
            },
        )?;
    }
    for chart in &value.drawing.charts {
        let relationship_type = match &chart.kind {
            ChartResourceKind::Classic => conformance.chart_rel(),
            ChartResourceKind::Extended { .. } => CHART_EX_REL,
        };
        let chart_uri = staged_uri(&chart_uris, &chart.relationship_id, "chart")?;
        add_relationship_checked(
            package,
            &drawing_uri,
            relationship_type,
            chart_uri.relative_ref(drawing_uri.base_uri()),
            chart.relationship_id.clone(),
            TargetMode::Internal,
        )?;
        if let ChartResourceKind::Extended {
            styles,
            color_styles,
            user_shapes,
            outbound_resources,
        } = &chart.kind
        {
            for (companions, relationship_type) in [
                (styles, CHART_STYLE_REL),
                (color_styles, CHART_COLOR_STYLE_REL),
            ] {
                for companion in companions {
                    let uri = staged_uri(
                        &companion_uris,
                        &(
                            chart.relationship_id.clone(),
                            companion.relationship_id.clone(),
                        ),
                        "chart companion",
                    )?;
                    add_relationship_checked(
                        package,
                        &chart_uri,
                        relationship_type,
                        uri.relative_ref(chart_uri.base_uri()),
                        companion.relationship_id.clone(),
                        TargetMode::Internal,
                    )?;
                }
            }
            if let Some(user_shapes) = user_shapes {
                let uri = staged_uri(&user_shape_uris, &chart.relationship_id, "chart user-shape")?;
                add_relationship_checked(
                    package,
                    &chart_uri,
                    conformance.chart_user_shapes_rel(),
                    uri.relative_ref(chart_uri.base_uri()),
                    user_shapes.relationship_id.clone(),
                    TargetMode::Internal,
                )?;
                for image in &user_shapes.images {
                    let image_uri = staged_uri(
                        &user_shape_image_uris,
                        &(chart.relationship_id.clone(), image.relationship_id.clone()),
                        "chart user-shape image",
                    )?;
                    add_relationship_checked(
                        package,
                        &uri,
                        conformance.image_rel(),
                        image_uri.relative_ref(uri.base_uri()),
                        image.relationship_id.clone(),
                        TargetMode::Internal,
                    )?;
                }
            }
            for resource in outbound_resources {
                let key = (
                    chart.relationship_id.clone(),
                    resource.relationship_id().to_owned(),
                );
                let uri = staged_uri(&outbound_uris, &key, "chart outbound")?;
                let relationship_type = match resource {
                    ChartOutboundResource::Image(_) => conformance.image_rel(),
                    ChartOutboundResource::ThemeOverride(_) => conformance.theme_override_rel(),
                    ChartOutboundResource::EmbeddedPackage(_) => conformance.package_rel(),
                };
                add_relationship_checked(
                    package,
                    &chart_uri,
                    relationship_type,
                    uri.relative_ref(chart_uri.base_uri()),
                    resource.relationship_id().to_owned(),
                    TargetMode::Internal,
                )?;
                if let ChartOutboundResource::ThemeOverride(theme) = resource {
                    for image in &theme.images {
                        let image_uri = staged_uri(
                            &theme_image_uris,
                            &(
                                chart.relationship_id.clone(),
                                theme.relationship_id.clone(),
                                image.relationship_id.clone(),
                            ),
                            "themeOverride image",
                        )?;
                        add_relationship_checked(
                            package,
                            &uri,
                            conformance.image_rel(),
                            image_uri.relative_ref(uri.base_uri()),
                            image.relationship_id.clone(),
                            TargetMode::Internal,
                        )?;
                    }
                }
            }
        }
    }
    Ok(())
}

pub(super) fn validate_package_value(value: &Package, conformance: Conformance) -> Result<()> {
    validate_entry(&value.entry)?;
    validate_chartsheet(&value.chartsheet)?;
    if value.drawing.content_type != DRAWING_CT || value.drawing.data.len() > MAX_DRAWING_BYTES {
        return Err(invalid("invalid or oversized chartsheet drawing resource"));
    }
    let drawing_uri = PackURI::new(&value.drawing.part_name).map_err(invalid)?;
    if !drawing_uri.as_str().starts_with("/xl/drawings/") {
        return Err(invalid("drawing resource is outside /xl/drawings"));
    }
    let references = drawing_chart_references(&value.drawing.data, conformance)?;
    if references.len() != value.drawing.charts.len() {
        return Err(invalid(
            "drawing chart references and chart resources differ",
        ));
    }
    let reference_ids = references
        .iter()
        .map(|reference| reference.relationship_id.as_str())
        .collect::<HashSet<_>>();
    let mut chart_ids = HashSet::with_capacity(value.drawing.charts.len());
    let mut resources = BTreeMap::new();
    let mut total = value.drawing.data.len();
    for chart in &value.drawing.charts {
        if !chart_ids.insert(chart.relationship_id.as_str()) {
            return Err(invalid("duplicate chart resource relationship ID"));
        }
        let reference = references
            .iter()
            .find(|reference| reference.relationship_id == chart.relationship_id)
            .ok_or_else(|| {
                invalid(format!(
                    "drawing does not reference chart relationship '{}'",
                    chart.relationship_id
                ))
            })?;
        validate_chart_resource_value(chart, reference, conformance, &mut total, &mut resources)?;
    }
    if chart_ids != reference_ids {
        return Err(invalid(
            "drawing chart references and chart resources are not a bijection",
        ));
    }
    validate_vml_pair(
        value.chartsheet.legacy_drawing_relationship_id.as_deref(),
        value.legacy_drawing.as_ref(),
        "legacyDrawing",
        &mut total,
        &mut resources,
    )?;
    validate_vml_pair(
        value
            .chartsheet
            .legacy_header_footer_drawing_relationship_id
            .as_deref(),
        value.legacy_header_footer_drawing.as_ref(),
        "legacyDrawingHF",
        &mut total,
        &mut resources,
    )?;
    match (
        &value.chartsheet.background_picture_relationship_id,
        &value.background_picture,
    ) {
        (None, None) => {},
        (Some(id), Some(picture)) => {
            validate_id(id)?;
            validate_id(&picture.relationship_id)?;
            if id != &picture.relationship_id {
                return Err(invalid(
                    "chartsheet picture relationship and resource metadata differ",
                ));
            }
            if id == &value.chartsheet.drawing_relationship_id {
                return Err(invalid(
                    "chartsheet drawing and picture relationship IDs collide",
                ));
            }
            let uri = PackURI::new(&picture.part_name).map_err(invalid)?;
            if !uri.as_str().starts_with("/xl/media/") {
                return Err(invalid("background image resource is outside /xl/media"));
            }
            add_resource(
                &mut total,
                picture.data.len(),
                MAX_BACKGROUND_IMAGE_BYTES,
                "background image bytes",
            )?;
            if resources
                .insert(picture.part_name.clone(), &picture.data)
                .is_some()
            {
                return Err(invalid("duplicate chartsheet resource part name"));
            }
        },
        _ => {
            return Err(invalid(
                "chartsheet picture relationship and resource must either both be present or both be absent",
            ));
        },
    }
    let printer_id = value
        .chartsheet
        .page_setup
        .as_ref()
        .and_then(|setup| setup.printer_settings_relationship_id.as_deref());
    match (printer_id, value.printer_settings.as_ref()) {
        (None, None) => {},
        (Some(id), Some(settings)) => {
            validate_id(id)?;
            validate_id(&settings.relationship_id)?;
            if id != settings.relationship_id {
                return Err(invalid(
                    "chartsheet pageSetup and Printer Settings relationship IDs differ",
                ));
            }
            validate_settings_bytes(&settings.resource.data)?;
            let uri = PackURI::new(&settings.resource.part_name).map_err(invalid)?;
            validate_printer_settings_uri(&uri)?;
            add_resource(
                &mut total,
                settings.resource.data.len(),
                MAX_SETTINGS_BYTES,
                "Printer Settings bytes",
            )?;
            if resources
                .insert(settings.resource.part_name.clone(), &settings.resource.data)
                .is_some()
            {
                return Err(invalid("duplicate chartsheet resource part name"));
            }
        },
        _ => {
            return Err(invalid(
                "chartsheet pageSetup relationship and Printer Settings resource must either both be present or both be absent",
            ));
        },
    }
    validate_extension_relationships(value, conformance)?;
    Ok(())
}

pub(super) fn validate_chart_resource_value<'a>(
    chart: &'a ChartResource,
    reference: &DrawingChartReference,
    conformance: Conformance,
    total: &mut usize,
    resources: &mut BTreeMap<String, &'a Vec<u8>>,
) -> Result<()> {
    validate_id(&chart.relationship_id)?;
    let uri = PackURI::new(&chart.part_name).map_err(invalid)?;
    if !uri.as_str().starts_with("/xl/charts/") || !uri.as_str().ends_with(".xml") {
        return Err(invalid(
            "chart resource is outside /xl/charts or lacks .xml suffix",
        ));
    }
    match (&chart.kind, reference.kind) {
        (ChartResourceKind::Classic, DrawingChartKind::Classic) => {
            if chart.content_type != CHART_CT {
                return Err(invalid("classic chart has invalid content type"));
            }
            validate_chart_xml(&chart.data, conformance)?;
            add_resource(total, chart.data.len(), MAX_CHART_BYTES, "chart bytes")?;
        },
        (
            ChartResourceKind::Extended {
                styles,
                color_styles,
                user_shapes,
                ..
            },
            DrawingChartKind::Extended,
        ) => {
            if chart.content_type != CHART_EX_CT {
                return Err(invalid("chartEx has invalid content type"));
            }
            validate_chart_ex_relationships(&chart.data, conformance)?;
            add_resource(total, chart.data.len(), MAX_CHART_EX_BYTES, "chartEx bytes")?;
            if styles.len() > MAX_CHART_STYLE_PARTS || color_styles.len() > MAX_CHART_STYLE_PARTS {
                return Err(limit("chart companion count"));
            }
            let mut ids = HashSet::new();
            for (companions, content_type, root, max_bytes) in [
                (styles, CHART_STYLE_CT, "chartStyle", MAX_CHART_STYLE_BYTES),
                (
                    color_styles,
                    CHART_COLOR_STYLE_CT,
                    "colorStyle",
                    MAX_CHART_COLOR_STYLE_BYTES,
                ),
            ] {
                for companion in companions {
                    validate_id(&companion.relationship_id)?;
                    if !ids.insert(companion.relationship_id.as_str()) {
                        return Err(invalid("chartEx companion relationship IDs collide"));
                    }
                    if companion.content_type != content_type {
                        return Err(invalid("chart companion has invalid content type"));
                    }
                    let uri = PackURI::new(&companion.part_name).map_err(invalid)?;
                    if !uri.as_str().starts_with("/xl/charts/") || !uri.as_str().ends_with(".xml") {
                        return Err(invalid(
                            "chart companion is outside /xl/charts or lacks .xml suffix",
                        ));
                    }
                    validate_chart_companion_xml(&companion.data, root, max_bytes)?;
                    add_resource(
                        total,
                        companion.data.len(),
                        max_bytes,
                        "chart companion bytes",
                    )?;
                    if resources
                        .insert(companion.part_name.clone(), &companion.data)
                        .is_some()
                    {
                        return Err(invalid("duplicate chartsheet resource part name"));
                    }
                }
            }
            if let Some(user_shapes) = user_shapes {
                validate_id(&user_shapes.relationship_id)?;
                if !ids.insert(user_shapes.relationship_id.as_str()) {
                    return Err(invalid("chartEx outbound relationship IDs collide"));
                }
                if user_shapes.content_type != CHART_USER_SHAPES_CT {
                    return Err(invalid("chartUserShapes has invalid content type"));
                }
                let uri = PackURI::new(&user_shapes.part_name).map_err(invalid)?;
                if !uri.as_str().starts_with("/xl/drawings/") || !uri.as_str().ends_with(".xml") {
                    return Err(invalid(
                        "chartUserShapes is outside /xl/drawings or lacks .xml suffix",
                    ));
                }
                let referenced = validate_chart_user_shapes_xml(&user_shapes.data, conformance)?;
                if referenced.len() != user_shapes.images.len()
                    || user_shapes.images.len() > MAX_CHART_USER_SHAPE_IMAGES
                {
                    return Err(invalid(
                        "chartUserShapes image relationship metadata does not match XML references",
                    ));
                }
                add_resource(
                    total,
                    user_shapes.data.len(),
                    MAX_CHART_USER_SHAPES_BYTES,
                    "chartUserShapes bytes",
                )?;
                if resources
                    .insert(user_shapes.part_name.clone(), &user_shapes.data)
                    .is_some()
                {
                    return Err(invalid("duplicate chartsheet resource part name"));
                }
                let mut image_ids = BTreeSet::new();
                for image in &user_shapes.images {
                    validate_id(&image.relationship_id)?;
                    if !referenced.contains(&image.relationship_id)
                        || !image_ids.insert(image.relationship_id.as_str())
                    {
                        return Err(invalid(
                            "chartUserShapes image metadata is duplicate or unreferenced",
                        ));
                    }
                    let image_uri = PackURI::new(&image.part_name).map_err(invalid)?;
                    if !image_uri.as_str().starts_with("/xl/media/")
                        || !image.content_type.validates_part_name(image_uri.as_str())
                    {
                        return Err(invalid(
                            "invalid chart user-shape image path or content type suffix",
                        ));
                    }
                    add_resource(
                        total,
                        image.data.len(),
                        MAX_CHART_USER_SHAPE_IMAGE_BYTES,
                        "chart user-shape image bytes",
                    )?;
                    if resources
                        .insert(image.part_name.clone(), &image.data)
                        .is_some()
                    {
                        return Err(invalid("duplicate chartsheet resource part name"));
                    }
                }
            }
            validate_chart_outbound_resources(chart, conformance, total, resources)?;
        },
        _ => {
            return Err(invalid(
                "drawing chart reference kind and chart resource kind differ",
            ));
        },
    }
    if resources
        .insert(chart.part_name.clone(), &chart.data)
        .is_some()
    {
        return Err(invalid("duplicate chart resource part name"));
    }
    Ok(())
}

pub(super) fn validate_chart_outbound_resources<'a>(
    chart: &'a ChartResource,
    conformance: Conformance,
    total: &mut usize,
    resources: &mut BTreeMap<String, &'a Vec<u8>>,
) -> Result<()> {
    let ChartResourceKind::Extended {
        styles,
        color_styles,
        user_shapes,
        outbound_resources,
    } = &chart.kind
    else {
        return Ok(());
    };
    let references = validate_chart_ex_relationships(&chart.data, conformance)?;
    let mut source_ids = HashSet::new();
    for companion in styles.iter().chain(color_styles) {
        source_ids.insert(companion.relationship_id.as_str());
    }
    if let Some(user_shapes) = user_shapes {
        source_ids.insert(user_shapes.relationship_id.as_str());
    }
    let mut direct_ids = BTreeSet::new();
    let mut package_id = None;
    let mut theme_count = 0usize;
    let mut package_count = 0usize;
    if outbound_resources
        .iter()
        .filter(|resource| matches!(resource, ChartOutboundResource::Image(_)))
        .count()
        > MAX_CHART_DIRECT_IMAGES
    {
        return Err(limit("chartEx direct image count"));
    }
    for resource in outbound_resources {
        validate_id(resource.relationship_id())?;
        if !source_ids.insert(resource.relationship_id()) {
            return Err(invalid("chartEx outbound relationship IDs collide"));
        }
        match resource {
            ChartOutboundResource::Image(image) => {
                if !direct_ids.insert(image.relationship_id.clone()) {
                    return Err(invalid("chartEx direct image relationship IDs collide"));
                }
                validate_chart_image_value(
                    image,
                    total,
                    resources,
                    MAX_CHART_USER_SHAPE_IMAGE_BYTES,
                    "chartEx direct image bytes",
                )?;
            },
            ChartOutboundResource::ThemeOverride(theme) => {
                theme_count += 1;
                if theme_count > 1 {
                    return Err(invalid("chartEx has multiple themeOverride relationships"));
                }
                if theme.content_type != THEME_OVERRIDE_CT {
                    return Err(invalid("themeOverride has invalid content type"));
                }
                let uri = PackURI::new(&theme.part_name).map_err(invalid)?;
                if !uri.as_str().starts_with("/xl/theme/") || !uri.as_str().ends_with(".xml") {
                    return Err(invalid(
                        "themeOverride is outside /xl/theme or lacks .xml suffix",
                    ));
                }
                let theme_references = validate_theme_override_xml(&theme.data, conformance)?;
                if theme.images.len() > MAX_CHART_THEME_IMAGES
                    || theme.images.len() != theme_references.len()
                {
                    return Err(invalid(
                        "themeOverride image relationship metadata does not match XML references",
                    ));
                }
                add_resource(
                    total,
                    theme.data.len(),
                    MAX_CHART_THEME_OVERRIDE_BYTES,
                    "themeOverride bytes",
                )?;
                if resources
                    .insert(theme.part_name.clone(), &theme.data)
                    .is_some()
                {
                    return Err(invalid("duplicate chartsheet resource part name"));
                }
                let mut image_ids = BTreeSet::new();
                for image in &theme.images {
                    validate_id(&image.relationship_id)?;
                    if !image_ids.insert(image.relationship_id.clone())
                        || !theme_references.contains(&image.relationship_id)
                    {
                        return Err(invalid(
                            "themeOverride image metadata is duplicate or unreferenced",
                        ));
                    }
                    validate_chart_image_value(
                        image,
                        total,
                        resources,
                        MAX_CHART_USER_SHAPE_IMAGE_BYTES,
                        "themeOverride image bytes",
                    )?;
                }
            },
            ChartOutboundResource::EmbeddedPackage(embedded) => {
                package_count += 1;
                if package_count > 1 {
                    return Err(invalid(
                        "chartEx has multiple embedded package relationships",
                    ));
                }
                let uri = PackURI::new(&embedded.part_name).map_err(invalid)?;
                if !uri.as_str().starts_with("/xl/embeddings/")
                    || !embedded.content_type.validates_part_name(uri.as_str())
                {
                    return Err(invalid(
                        "invalid chartEx embedded package path or content type suffix",
                    ));
                }
                add_resource(
                    total,
                    embedded.data.len(),
                    MAX_CHART_EMBEDDED_PACKAGE_BYTES,
                    "chartEx embedded package bytes",
                )?;
                if resources
                    .insert(embedded.part_name.clone(), &embedded.data)
                    .is_some()
                {
                    return Err(invalid("duplicate chartsheet resource part name"));
                }
                package_id = Some(embedded.relationship_id.clone());
            },
        }
    }
    if direct_ids != references.images || package_id != references.package {
        return Err(invalid(
            "chartEx outbound relationship metadata does not match XML references",
        ));
    }
    Ok(())
}

pub(super) fn validate_chart_image_value<'a>(
    image: &'a ImageResource,
    total: &mut usize,
    resources: &mut BTreeMap<String, &'a Vec<u8>>,
    max_bytes: usize,
    label: &str,
) -> Result<()> {
    let uri = PackURI::new(&image.part_name).map_err(invalid)?;
    if !uri.as_str().starts_with("/xl/media/")
        || !image.content_type.validates_part_name(uri.as_str())
    {
        return Err(invalid("invalid chart image path or content type suffix"));
    }
    add_resource(total, image.data.len(), max_bytes, label)?;
    if resources
        .insert(image.part_name.clone(), &image.data)
        .is_some()
    {
        return Err(invalid("duplicate chartsheet resource part name"));
    }
    Ok(())
}

pub(super) fn validate_vml_pair<'a>(
    id: Option<&str>,
    resource: Option<&'a VmlDrawingResource>,
    label: &str,
    total: &mut usize,
    resources: &mut BTreeMap<String, &'a Vec<u8>>,
) -> Result<()> {
    match (id, resource) {
        (None, None) => Ok(()),
        (Some(id), Some(resource)) => {
            validate_id(id)?;
            validate_id(&resource.relationship_id)?;
            if id != resource.relationship_id {
                return Err(invalid(format!(
                    "{label} relationship and resource metadata differ"
                )));
            }
            let uri = PackURI::new(&resource.part_name).map_err(invalid)?;
            if !uri.as_str().starts_with("/xl/drawings/")
                || !uri.as_str().ends_with(".vml")
                || resource.content_type != VML_DRAWING_CT
            {
                return Err(invalid(format!(
                    "invalid {label} VML resource path or content type"
                )));
            }
            add_resource(
                total,
                resource.data.len(),
                MAX_VML_DRAWING_BYTES,
                "VML drawing bytes",
            )?;
            if resources
                .insert(resource.part_name.clone(), &resource.data)
                .is_some()
            {
                return Err(invalid("duplicate chartsheet resource part name"));
            }
            Ok(())
        },
        _ => Err(invalid(format!(
            "{label} relationship and resource must either both be present or both be absent"
        ))),
    }
}

pub(super) fn workbook_entry(
    root: &Node,
    conformance: Conformance,
    relationship_id: &str,
    part_name: String,
) -> Result<Entry> {
    let sheets = required_child(root, conformance.sml(), "sheets")?;
    let mut found = None;
    for sheet in &sheets.children {
        if sheet.namespace == conformance.sml()
            && sheet.name == "sheet"
            && optional(sheet, conformance.rel(), "id") == Some(relationship_id)
        {
            if found.is_some() {
                return Err(invalid(
                    "multiple workbook sheets reference the chartsheet relationship",
                ));
            }
            found = Some(parse_entry(sheet, conformance, part_name.clone())?);
        }
    }
    found.ok_or_else(|| invalid("workbook has no sheet entry for the chartsheet relationship"))
}

pub(super) fn parse_entry(
    node: &Node,
    conformance: Conformance,
    part_name: String,
) -> Result<Entry> {
    leaf(node, "workbook sheet")?;
    let state = optional(node, "", "state")
        .map(parse_state)
        .transpose()?
        .unwrap_or(State::Visible);
    Ok(Entry {
        name: required(node, "", "name")?.to_owned(),
        sheet_id: required(node, "", "sheetId")?
            .parse()
            .map_err(|_| invalid("invalid workbook sheetId"))?,
        state,
        workbook_relationship_id: required(node, conformance.rel(), "id")?.to_owned(),
        part_name,
    })
}

pub(super) fn validate_new_entry(
    root: &Node,
    conformance: Conformance,
    entry: &Entry,
) -> Result<()> {
    let sheets = required_child(root, conformance.sml(), "sheets")?;
    for sheet in &sheets.children {
        if sheet.namespace == conformance.sml() && sheet.name == "sheet" {
            if optional(sheet, "", "name").is_some_and(|v| v.eq_ignore_ascii_case(&entry.name)) {
                return Err(invalid("workbook sheet name already exists"));
            }
            if optional(sheet, "", "sheetId") == Some(entry.sheet_id.to_string().as_str()) {
                return Err(invalid("workbook sheetId already exists"));
            }
        }
    }
    Ok(())
}

pub(super) fn validate_entry(entry: &Entry) -> Result<()> {
    bounded(&entry.name)?;
    if entry.name.is_empty()
        || entry.name.chars().count() > 31
        || entry
            .name
            .chars()
            .any(|c| matches!(c, ':' | '\\' | '/' | '?' | '*' | '[' | ']'))
    {
        return Err(invalid("invalid Excel chartsheet name"));
    }
    if entry.sheet_id == 0 {
        return Err(invalid("chartsheet sheetId must be positive"));
    }
    validate_id(&entry.workbook_relationship_id)?;
    let uri = PackURI::new(&entry.part_name).map_err(invalid)?;
    if !uri.as_str().starts_with("/xl/chartsheets/") {
        return Err(invalid("chartsheet part is outside /xl/chartsheets"));
    }
    Ok(())
}

pub(super) fn insert_workbook_entry(
    xml: &[u8],
    entry: &Entry,
    conformance: Conformance,
) -> Result<Vec<u8>> {
    let mut fragment = Vec::new();
    fragment.extend_from_slice(b"<x:sheet xmlns:x=\"");
    escape(&mut fragment, conformance.sml());
    fragment.extend_from_slice(b"\" xmlns:r=\"");
    escape(&mut fragment, conformance.rel());
    fragment.extend_from_slice(b"\"");
    attr(&mut fragment, "name", &entry.name);
    attr(&mut fragment, "sheetId", &entry.sheet_id.to_string());
    if entry.state != State::Visible {
        attr(
            &mut fragment,
            "state",
            match entry.state {
                State::Visible => "visible",
                State::Hidden => "hidden",
                State::VeryHidden => "veryHidden",
            },
        );
    }
    attr(&mut fragment, "r:id", &entry.workbook_relationship_id);
    fragment.extend_from_slice(b"/>");
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut sheets_depth = None;
    let mut position = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("workbook XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                let core = matches!(namespace, ResolveResult::Bound(Namespace(v)) if v == conformance.sml().as_bytes());
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("workbook XML depth"))?;
                if core
                    && element.local_name().as_ref() == b"sheets"
                    && sheets_depth.replace(depth).is_some()
                {
                    return Err(invalid("workbook has multiple sheets collections"));
                }
            },
            Event::Empty(element) if element.local_name().as_ref() == b"sheets" => {
                return Err(invalid("cannot insert into empty sheets collection"));
            },
            Event::End(element) => {
                if sheets_depth == Some(depth) && element.local_name().as_ref() == b"sheets" {
                    position = Some(start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unexpected workbook closing element"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    let position = position.ok_or_else(|| invalid("workbook is missing sheets collection"))?;
    let size = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| limit("updated workbook bytes"))?;
    if size > MAX_XML_BYTES {
        return Err(limit("updated workbook bytes"));
    }
    let prefix = xml
        .get(..position)
        .ok_or_else(|| invalid("invalid workbook XML insertion offset"))?;
    let suffix = xml
        .get(position..)
        .ok_or_else(|| invalid("invalid workbook XML insertion offset"))?;
    let mut out = Vec::with_capacity(size);
    out.extend_from_slice(prefix);
    out.extend_from_slice(&fragment);
    out.extend_from_slice(suffix);
    Ok(out)
}

pub(super) fn known_chartsheet_relationship_ids(value: &Chart) -> BTreeSet<String> {
    let mut ids = BTreeSet::new();
    ids.insert(value.drawing_relationship_id.clone());
    for id in [
        value.legacy_drawing_relationship_id.as_ref(),
        value.legacy_header_footer_drawing_relationship_id.as_ref(),
        value.background_picture_relationship_id.as_ref(),
        value
            .page_setup
            .as_ref()
            .and_then(|setup| setup.printer_settings_relationship_id.as_ref()),
    ]
    .into_iter()
    .flatten()
    {
        ids.insert(id.clone());
    }
    ids
}
pub(super) fn extension_relationship_ids(
    value: &Chart,
    conformance: Conformance,
) -> Result<BTreeSet<String>> {
    let mut ids = BTreeSet::new();
    if let Some(list) = &value.extension_list {
        for extension in &list.extensions {
            let node = parse_document(&extension.payload_xml, MAX_EXTENSION_PAYLOAD_BYTES)?;
            collect_extension_relationship_ids(
                &node,
                conformance.rel(),
                &mut ids,
                MAX_EXTENSION_RELATIONSHIPS,
            )?;
        }
    }
    Ok(ids)
}
pub(super) fn validate_extension_relationship_string(value: &str, label: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!(
            "extension relationship {label} cannot be empty"
        )));
    }
    if value.len() > MAX_EXTENSION_RELATIONSHIP_STRING_BYTES {
        return Err(limit("extension relationship string bytes"));
    }
    Ok(())
}
pub(super) fn validate_extension_relationships(
    value: &Package,
    conformance: Conformance,
) -> Result<()> {
    if value.extension_relationships.len() > MAX_EXTENSION_RELATIONSHIPS {
        return Err(limit("extension relationship count"));
    }
    let referenced = extension_relationship_ids(&value.chartsheet, conformance)?;
    let known = known_chartsheet_relationship_ids(&value.chartsheet);
    let unknown = referenced
        .difference(&known)
        .cloned()
        .collect::<BTreeSet<_>>();
    if value.extension_relationships.len() != unknown.len() {
        return Err(invalid(
            "extension relationship metadata does not match referenced unknown relationships",
        ));
    }
    let mut seen = BTreeSet::new();
    for relationship in &value.extension_relationships {
        validate_id(&relationship.relationship_id)?;
        if !unknown.contains(&relationship.relationship_id)
            || !seen.insert(relationship.relationship_id.clone())
        {
            return Err(invalid(
                "extension relationship metadata is duplicate or unreferenced",
            ));
        }
        validate_extension_relationship_string(&relationship.relationship_type, "type")?;
        match &relationship.target {
            ExtensionRelationshipTarget::Internal { part_name } => {
                validate_extension_relationship_string(part_name, "target")?;
                let uri = PackURI::new(part_name).map_err(invalid)?;
                if !uri.as_str().starts_with('/') {
                    return Err(invalid(
                        "internal extension relationship target must be an absolute part name",
                    ));
                }
            },
            ExtensionRelationshipTarget::External { target } => {
                validate_extension_relationship_string(target, "target")?;
            },
        }
    }
    Ok(())
}
