//! Legend and legend-entry record families.

use super::super::validation::invalid_chart_input;
use super::super::xml::write_fragment;
use super::presentation::write_layout;
use crate::chart::legend::Legend;
use std::io::Write;

pub(super) fn write_legend<W: Write>(writer: &mut W, legend: &Legend) -> std::io::Result<()> {
    write!(writer, "<c:legend>")?;
    write!(
        writer,
        r#"<c:legendPos val="{}"/>"#,
        legend.position.xml_value()
    )?;
    let mut entry_indexes = std::collections::HashSet::with_capacity(legend.entries.len());
    for entry in &legend.entries {
        if !entry_indexes.insert(entry.index) {
            return Err(invalid_chart_input(format!(
                "chart legend contains duplicate entry index {}",
                entry.index
            )));
        }
        write!(writer, "<c:legendEntry>")?;
        write!(writer, r#"<c:idx val="{}"/>"#, entry.index)?;
        if let Some(text_properties) = entry.text_properties.as_ref() {
            if entry.deleted {
                return Err(invalid_chart_input(
                    "chart legend entry cannot be deleted and have text properties",
                ));
            }
            write_fragment(writer, text_properties.as_xml())?;
        } else {
            write!(
                writer,
                r#"<c:delete val="{}"/>"#,
                if entry.deleted { "1" } else { "0" }
            )?;
        }
        if let Some(extension_list) = entry.extension_list.as_ref() {
            write_fragment(writer, extension_list.as_xml())?;
        }
        write!(writer, "</c:legendEntry>")?;
    }
    if let Some(layout) = legend.layout.as_ref() {
        write_layout(writer, Some(layout))?;
    }
    write!(
        writer,
        r#"<c:overlay val="{}"/>"#,
        if legend.overlay { "1" } else { "0" }
    )?;
    if let Some(shape_properties) = legend.shape_properties.as_ref() {
        write_fragment(writer, shape_properties.as_xml())?;
    }
    if let Some(text_properties) = legend.text_properties.as_ref() {
        write_fragment(writer, text_properties.as_xml())?;
    }
    if let Some(extension_list) = legend.extension_list.as_ref() {
        write_fragment(writer, extension_list.as_xml())?;
    }
    write!(writer, "</c:legend>")?;
    Ok(())
}
