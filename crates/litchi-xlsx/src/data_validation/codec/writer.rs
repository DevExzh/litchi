//! Canonical `SpreadsheetML` data-validation XML serialization.

use super::super::model::{
    Collection, Conformance, ListSource, Source, Validation, ValidationErrorStyle,
    ValidationImeMode, ValidationOperator, ValidationType,
};
use super::super::{EXTENSION_URI, X12AC_URI, X14_URI, XM_URI, XR_URI};
use super::validation::{validate_data_validation_collections, validate_rule};
use super::wire::{BoundedXml, invalid, sqref_text};
use crate::error::Result;
use litchi_core::xml::escape::escape_xml;

pub fn write_data_validation_collections(
    values: &[Collection],
    conformance: Conformance,
) -> Result<String> {
    let core = write_data_validation_core(values, conformance)?;
    let extensions = write_data_validation_extensions(values, conformance)?;
    let mut xml = BoundedXml::new();
    xml.push_str(&core)?;
    xml.push_str(&extensions)?;
    Ok(xml.finish())
}

pub fn write_data_validation_core(
    values: &[Collection],
    conformance: Conformance,
) -> Result<String> {
    validate_data_validation_collections(values)?;
    let mut xml = BoundedXml::new();
    if let Some(collection) = values
        .iter()
        .find(|collection| collection.source == Source::Core)
    {
        xml.write_arguments(format_args!(
            "<dataValidations xmlns=\"{}\" xmlns:xr=\"{}\"",
            conformance.namespace(),
            XR_URI
        ))?;
        write_collection_attributes(&mut xml, collection)?;
        xml.push_str(">")?;
        for rule in &collection.validations {
            write_rule(&mut xml, rule)?;
        }
        xml.push_str("</dataValidations>")?;
    }
    Ok(xml.finish())
}

pub fn write_data_validation_extensions(
    values: &[Collection],
    conformance: Conformance,
) -> Result<String> {
    validate_data_validation_collections(values)?;
    let mut xml = BoundedXml::new();
    if let Some(collection) = values
        .iter()
        .find(|collection| collection.source == Source::Office2010)
    {
        xml.write_arguments(format_args!(
            "<extLst xmlns=\"{}\"><ext uri=\"{}\"><x14:dataValidations xmlns:x14=\"{}\" xmlns:xm=\"{}\" xmlns:x12ac=\"{}\" xmlns:xr=\"{}\"",
            conformance.namespace(),
            EXTENSION_URI,
            X14_URI,
            XM_URI,
            X12AC_URI,
            XR_URI,
        ))?;
        write_collection_attributes(&mut xml, collection)?;
        xml.push_str(">")?;
        for rule in &collection.validations {
            write_rule(&mut xml, rule)?;
        }
        xml.push_str("</x14:dataValidations></ext></extLst>")?;
    }
    Ok(xml.finish())
}

fn write_collection_attributes(xml: &mut BoundedXml, value: &Collection) -> Result<()> {
    if value.disable_prompts {
        xml.push_str(" disablePrompts=\"1\"")?;
    }
    if let Some(value) = value.x_window {
        xml.write_arguments(format_args!(" xWindow=\"{value}\""))?;
    }
    if let Some(value) = value.y_window {
        xml.write_arguments(format_args!(" yWindow=\"{value}\""))?;
    }
    xml.write_arguments(format_args!(" count=\"{}\"", value.validations.len()))?;
    Ok(())
}

fn write_rule(xml: &mut BoundedXml, value: &Validation) -> Result<()> {
    validate_rule(value)?;
    let prefix = if value.source == Source::Office2010 {
        "x14:"
    } else {
        ""
    };
    xml.write_arguments(format_args!("<{prefix}dataValidation"))?;
    if value.validation_type != ValidationType::None {
        xml.write_arguments(format_args!(" type=\"{}\"", value.validation_type.as_str()))?;
    }
    if value.operator != ValidationOperator::Between {
        xml.write_arguments(format_args!(" operator=\"{}\"", value.operator.as_str()))?;
    }
    if value.error_style != ValidationErrorStyle::Stop {
        xml.write_arguments(format_args!(
            " errorStyle=\"{}\"",
            value.error_style.as_str()
        ))?;
    }
    if value.ime_mode != ValidationImeMode::NoControl {
        xml.write_arguments(format_args!(" imeMode=\"{}\"", value.ime_mode.as_str()))?;
    }
    for (name, enabled) in [
        ("allowBlank", value.allow_blank),
        ("showDropDown", value.show_drop_down),
        ("showInputMessage", value.show_input_message),
        ("showErrorMessage", value.show_error_message),
    ] {
        if enabled {
            xml.write_arguments(format_args!(" {name}=\"1\""))?;
        }
    }
    for (name, text) in [
        ("errorTitle", value.error_title.as_deref()),
        ("error", value.error.as_deref()),
        ("promptTitle", value.prompt_title.as_deref()),
        ("prompt", value.prompt.as_deref()),
    ] {
        if let Some(text) = text {
            xml.write_arguments(format_args!(" {name}=\"{}\"", escape_xml(text)))?;
        }
    }
    if let Some(uid) = value.uid.as_deref() {
        xml.write_arguments(format_args!(" xr:uid=\"{}\"", escape_xml(uid)))?;
    }
    if value.source == Source::Core {
        let sqref = sqref_text(&value.sqref)?;
        xml.write_arguments(format_args!(" sqref=\"{}\"", escape_xml(&sqref)))?;
    }
    xml.push_str(">")?;
    write_formula(xml, prefix, 1, value.formula1.as_ref())?;
    if let Some(formula) = value.formula2.as_ref() {
        write_formula_source(xml, prefix, 2, FormulaSource::Formula(&formula.0))?;
    }
    if value.source == Source::Office2010 {
        xml.push_str("<xm:sqref")?;
        for (name, enabled) in [
            ("edited", value.sqref.edited),
            ("split", value.sqref.split),
            ("adjusted", value.sqref.adjusted),
            ("adjust", value.sqref.adjust),
        ] {
            if enabled {
                xml.write_arguments(format_args!(" {name}=\"1\""))?;
            }
        }
        let sqref = sqref_text(&value.sqref)?;
        xml.write_arguments(format_args!(">{}</xm:sqref>", escape_xml(&sqref)))?;
    }
    xml.write_arguments(format_args!("</{prefix}dataValidation>"))?;
    Ok(())
}

enum FormulaSource<'a> {
    Formula(&'a str),
    QuotedList(&'a str),
}

fn write_formula(
    xml: &mut BoundedXml,
    prefix: &str,
    number: u8,
    value: Option<&ListSource>,
) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    let source = match value {
        ListSource::Formula(value) => FormulaSource::Formula(&value.0),
        ListSource::QuotedList(value) => FormulaSource::QuotedList(value),
    };
    write_formula_source(xml, prefix, number, source)
}

fn write_formula_source(
    xml: &mut BoundedXml,
    prefix: &str,
    number: u8,
    value: FormulaSource<'_>,
) -> Result<()> {
    xml.write_arguments(format_args!("<{prefix}formula{number}>"))?;
    match (prefix.is_empty(), value) {
        (true, FormulaSource::Formula(value)) => {
            xml.push_str(&escape_xml(value))?;
        },
        (true, FormulaSource::QuotedList(_)) => {
            return Err(invalid(
                "quoted-list source requires Office 2010 data validation",
            ));
        },
        (false, FormulaSource::Formula(value)) => {
            xml.write_arguments(format_args!("<xm:f>{}</xm:f>", escape_xml(value)))?;
        },
        (false, FormulaSource::QuotedList(value)) => {
            xml.write_arguments(format_args!(
                "<x12ac:list>{}</x12ac:list>",
                escape_xml(value)
            ))?;
        },
    }
    xml.write_arguments(format_args!("</{prefix}formula{number}>"))?;
    Ok(())
}
