//! Chart-space and package-level DrawingML chart records.

use super::super::validation::{invalid_chart_input, validate_chart_style};
use super::super::xml::{write_bool_element, write_fragment, write_text_element};
use super::{
    legend::write_legend,
    plot_area::write_plot_area,
    presentation::{write_marker, write_title, write_view_3d, write_wall_floor},
    series::write_data_label,
};
use crate::chart::model::{
    Chart, ColorMapOverride, ColorMapping, ExternalData, HeaderFooter, PageMargins, PageSetup,
    PivotFormat, PivotSource, PrintSettings, Protection, UserShapes,
};
use litchi_core::xml::escape_xml;
use std::io::Write;

pub(in crate::chart::writer) fn write_chart_space<W: Write>(
    writer: &mut W,
    chart: &Chart,
    external_data_relationship_id: Option<&str>,
    user_shapes_relationship_id: Option<&str>,
) -> std::io::Result<()> {
    write!(
        writer,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#
    )?;
    write!(
        writer,
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#
    )?;
    write!(
        writer,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#
    )?;
    write!(
        writer,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#
    )?;

    write!(
        writer,
        r#"<c:date1904 val="{}"/>"#,
        if chart.date_1904 { "1" } else { "0" }
    )?;
    if let Some(language) = chart.language.as_ref() {
        write!(writer, r#"<c:lang val="{}"/>"#, escape_xml(language))?;
    }
    write!(
        writer,
        r#"<c:roundedCorners val="{}"/>"#,
        if chart.rounded_corners { "1" } else { "0" }
    )?;

    validate_chart_style(chart.style)?;
    if let Some(ref style) = chart.style {
        write!(writer, r#"<c:style val="{}"/>"#, style)?;
    }
    if let Some(color_map) = chart.color_map_override.as_ref() {
        write_color_map_override(writer, color_map)?;
    }

    if let Some(source) = chart.pivot_source.as_ref() {
        write_pivot_source(writer, source)?;
    }
    if let Some(protection) = chart.protection.as_ref() {
        write_chart_protection(writer, protection)?;
    }

    write!(writer, "<c:chart>")?;

    if let Some(ref title) = chart.title {
        write_title(
            writer,
            title,
            chart.title_layout.as_ref(),
            chart.title_overlay,
            chart.title_shape_properties.as_ref(),
            chart.title_text_properties.as_ref(),
            chart.title_extension_list.as_ref(),
        )?;
    }

    write!(
        writer,
        r#"<c:autoTitleDeleted val="{}"/>"#,
        if chart.auto_title_deleted { "1" } else { "0" }
    )?;

    if let Some(formats) = chart.pivot_formats.as_deref() {
        write_pivot_formats(writer, formats)?;
    }

    if let Some(ref view) = chart.view_3d {
        write_view_3d(writer, view)?;
    }

    if let Some(ref floor) = chart.floor {
        write!(writer, "<c:floor>")?;
        write_wall_floor(writer, floor)?;
        write!(writer, "</c:floor>")?;
    }

    if let Some(ref back_wall) = chart.back_wall {
        write!(writer, "<c:backWall>")?;
        write_wall_floor(writer, back_wall)?;
        write!(writer, "</c:backWall>")?;
    }

    if let Some(ref side_wall) = chart.side_wall {
        write!(writer, "<c:sideWall>")?;
        write_wall_floor(writer, side_wall)?;
        write!(writer, "</c:sideWall>")?;
    }

    write_plot_area(writer, &chart.plot_area)?;

    if let Some(ref legend) = chart.legend {
        write_legend(writer, legend)?;
    }

    write!(
        writer,
        r#"<c:plotVisOnly val="{}"/>"#,
        if chart.plot_visible_only { "1" } else { "0" }
    )?;
    write!(
        writer,
        r#"<c:dispBlanksAs val="{}"/>"#,
        chart.display_blanks_as.xml_value()
    )?;

    if chart.show_data_labels_over_max {
        write!(writer, r#"<c:showDLblsOverMax val="1"/>"#)?;
    }
    if let Some(extension_list) = chart.chart_extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }

    write!(writer, "</c:chart>")?;

    if let Some(shape_properties) = chart.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(text_properties) = chart.text_properties.as_ref() {
        write_fragment(writer, text_properties.as_xml())?;
    }

    if let Some(external_data) = chart.external_data.as_ref() {
        write_external_data(writer, external_data, external_data_relationship_id)?;
    } else if external_data_relationship_id.is_some() {
        return Err(invalid_chart_input(
            "chart package supplied external data without chart metadata",
        ));
    }

    if let Some(settings) = chart.print_settings.as_ref() {
        write_print_settings(writer, settings)?;
    }
    if let Some(user_shapes) = chart.user_shapes.as_ref() {
        write_user_shapes(writer, user_shapes, user_shapes_relationship_id)?;
    } else if user_shapes_relationship_id.is_some() {
        return Err(invalid_chart_input(
            "chart package supplied user shapes without chart metadata",
        ));
    }
    if let Some(extension_list) = chart.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }

    write!(writer, "</c:chartSpace>")?;

    Ok(())
}

fn write_user_shapes<W: Write>(
    writer: &mut W,
    user_shapes: &UserShapes,
    relationship_id_override: Option<&str>,
) -> std::io::Result<()> {
    let relationship_id = relationship_id_override
        .or(user_shapes.relationship_id.as_deref())
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid_chart_input("chart user shapes have no relationship ID"))?;
    write!(
        writer,
        r#"<c:userShapes r:id="{}"/>"#,
        escape_xml(relationship_id)
    )?;
    Ok(())
}

fn write_external_data<W: Write>(
    writer: &mut W,
    external_data: &ExternalData,
    relationship_id_override: Option<&str>,
) -> std::io::Result<()> {
    let relationship_id = relationship_id_override
        .or(external_data.relationship_id.as_deref())
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid_chart_input("chart external data has no relationship ID"))?;
    write!(
        writer,
        r#"<c:externalData r:id="{}">"#,
        escape_xml(relationship_id)
    )?;
    if let Some(auto_update) = external_data.auto_update {
        write!(
            writer,
            r#"<c:autoUpdate val="{}"/>"#,
            if auto_update { "1" } else { "0" }
        )?;
    }
    write!(writer, "</c:externalData>")?;
    Ok(())
}

fn write_pivot_source<W: Write>(writer: &mut W, source: &PivotSource) -> std::io::Result<()> {
    write!(
        writer,
        r#"<c:pivotSource><c:name>{}</c:name><c:fmtId val="{}"/></c:pivotSource>"#,
        escape_xml(&source.name),
        source.format_id
    )?;
    Ok(())
}

fn write_color_map_override<W: Write>(
    writer: &mut W,
    color_map: &ColorMapOverride,
) -> std::io::Result<()> {
    match color_map {
        ColorMapOverride::Master => {
            write!(writer, "<c:clrMapOvr><a:masterClrMapping/></c:clrMapOvr>")?;
        },
        ColorMapOverride::Override(mapping) => {
            write!(writer, "<c:clrMapOvr><a:overrideClrMapping")?;
            write_color_mapping_attributes(writer, mapping)?;
            write!(writer, "/></c:clrMapOvr>")?;
        },
    }
    Ok(())
}

fn write_color_mapping_attributes<W: Write>(
    writer: &mut W,
    mapping: &ColorMapping,
) -> std::io::Result<()> {
    for (name, value) in [
        ("bg1", mapping.background1),
        ("tx1", mapping.text1),
        ("bg2", mapping.background2),
        ("tx2", mapping.text2),
        ("accent1", mapping.accent1),
        ("accent2", mapping.accent2),
        ("accent3", mapping.accent3),
        ("accent4", mapping.accent4),
        ("accent5", mapping.accent5),
        ("accent6", mapping.accent6),
        ("hlink", mapping.hyperlink),
        ("folHlink", mapping.followed_hyperlink),
    ] {
        write!(writer, r#" {name}="{}""#, value.as_str())?;
    }
    Ok(())
}

fn write_chart_protection<W: Write>(
    writer: &mut W,
    protection: &Protection,
) -> std::io::Result<()> {
    write!(writer, "<c:protection>")?;
    for (name, value) in [
        ("chartObject", protection.chart_object),
        ("data", protection.data),
        ("formatting", protection.formatting),
        ("selection", protection.selection),
        ("userInterface", protection.user_interface),
    ] {
        if let Some(value) = value {
            write_bool_element(writer, name, value)?;
        }
    }
    write!(writer, "</c:protection>")?;
    Ok(())
}

fn write_pivot_formats<W: Write>(writer: &mut W, formats: &[PivotFormat]) -> std::io::Result<()> {
    let mut indexes = std::collections::HashSet::with_capacity(formats.len());
    write!(writer, "<c:pivotFmts>")?;
    for format in formats {
        if !indexes.insert(format.index) {
            return Err(invalid_chart_input(format!(
                "chart contains duplicate pivot-format index {}",
                format.index
            )));
        }
        write!(writer, r#"<c:pivotFmt><c:idx val="{}"/>"#, format.index)?;
        if let Some(shape_properties) = format.shape_properties.as_ref() {
            write_fragment(writer, shape_properties.as_xml())?;
        }
        if let Some(text_properties) = format.text_properties.as_ref() {
            write_fragment(writer, text_properties.as_xml())?;
        }
        if let Some(marker) = format.marker.as_ref() {
            write_marker(writer, marker, "chart pivot-format")?;
        }
        if let Some(label) = format.data_label.as_ref() {
            write_data_label(writer, label)?;
        }
        if let Some(extension_list) = format.extension_list.as_ref() {
            write_fragment(writer, extension_list.as_xml())?;
        }
        write!(writer, "</c:pivotFmt>")?;
    }
    write!(writer, "</c:pivotFmts>")?;
    Ok(())
}

fn write_print_settings<W: Write>(writer: &mut W, settings: &PrintSettings) -> std::io::Result<()> {
    write!(writer, "<c:printSettings>")?;
    if let Some(header_footer) = settings.header_footer.as_ref() {
        write_chart_header_footer(writer, header_footer)?;
    }
    if let Some(margins) = settings.page_margins.as_ref() {
        write_chart_page_margins(writer, margins)?;
    }
    if let Some(setup) = settings.page_setup.as_ref() {
        write_chart_page_setup(writer, setup)?;
    }
    write!(writer, "</c:printSettings>")?;
    Ok(())
}

fn write_chart_header_footer<W: Write>(
    writer: &mut W,
    header_footer: &HeaderFooter,
) -> std::io::Result<()> {
    write!(
        writer,
        r#"<c:headerFooter alignWithMargins="{}" differentOddEven="{}" differentFirst="{}">"#,
        if header_footer.align_with_margins {
            "1"
        } else {
            "0"
        },
        if header_footer.different_odd_even {
            "1"
        } else {
            "0"
        },
        if header_footer.different_first {
            "1"
        } else {
            "0"
        }
    )?;
    for (name, value) in [
        ("oddHeader", header_footer.odd_header.as_ref()),
        ("oddFooter", header_footer.odd_footer.as_ref()),
        ("evenHeader", header_footer.even_header.as_ref()),
        ("evenFooter", header_footer.even_footer.as_ref()),
        ("firstHeader", header_footer.first_header.as_ref()),
        ("firstFooter", header_footer.first_footer.as_ref()),
    ] {
        if let Some(value) = value {
            write_text_element(writer, name, value)?;
        }
    }
    write!(writer, "</c:headerFooter>")?;
    Ok(())
}

fn write_chart_page_margins<W: Write>(
    writer: &mut W,
    margins: &PageMargins,
) -> std::io::Result<()> {
    for (name, value) in [
        ("left", margins.left),
        ("right", margins.right),
        ("top", margins.top),
        ("bottom", margins.bottom),
        ("header", margins.header),
        ("footer", margins.footer),
    ] {
        if !value.is_finite() {
            return Err(invalid_chart_input(format!(
                "chart {name} page margin must be finite"
            )));
        }
    }
    write!(
        writer,
        r#"<c:pageMargins l="{}" r="{}" t="{}" b="{}" header="{}" footer="{}"/>"#,
        margins.left, margins.right, margins.top, margins.bottom, margins.header, margins.footer
    )?;
    Ok(())
}

fn write_chart_page_setup<W: Write>(writer: &mut W, setup: &PageSetup) -> std::io::Result<()> {
    write!(
        writer,
        r#"<c:pageSetup paperSize="{}" firstPageNumber="{}" orientation="{}" blackAndWhite="{}" draft="{}" useFirstPageNumber="{}" horizontalDpi="{}" verticalDpi="{}" copies="{}"/>"#,
        setup.paper_size,
        setup.first_page_number,
        setup.orientation.xml_value(),
        if setup.black_and_white { "1" } else { "0" },
        if setup.draft { "1" } else { "0" },
        if setup.use_first_page_number {
            "1"
        } else {
            "0"
        },
        setup.horizontal_dpi,
        setup.vertical_dpi,
        setup.copies
    )?;
    Ok(())
}
