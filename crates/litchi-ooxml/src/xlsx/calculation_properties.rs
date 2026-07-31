//! Immutable XLSX workbook calculation-properties read model.

use crate::error::{OoxmlError, Result};
use crate::xlsx::namespace::is_spreadsheetml_name;
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

/// Workbook formula calculation mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum WorkbookCalculationMode {
    Manual,
    #[default]
    Automatic,
    AutomaticExceptTables,
}

impl WorkbookCalculationMode {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "manual" => Ok(Self::Manual),
            "auto" => Ok(Self::Automatic),
            "autoNoTable" => Ok(Self::AutomaticExceptTables),
            _ => Err(invalid(format!("invalid calcPr calcMode '{value}'"))),
        }
    }
}

/// Cell-reference style used by formulas in the workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum WorkbookReferenceMode {
    #[default]
    A1,
    R1C1,
}

impl WorkbookReferenceMode {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "A1" => Ok(Self::A1),
            "R1C1" => Ok(Self::R1C1),
            _ => Err(invalid(format!("invalid calcPr refMode '{value}'"))),
        }
    }
}

/// Effective workbook calculation policy from `calcPr`.
#[derive(Debug, Clone, PartialEq)]
pub struct WorkbookCalculationProperties {
    calculation_id: u32,
    calculation_mode: WorkbookCalculationMode,
    full_calculation_on_load: bool,
    reference_mode: WorkbookReferenceMode,
    iterative_calculation: bool,
    iteration_count: u32,
    iteration_delta: f64,
    full_precision: bool,
    calculation_completed: bool,
    calculate_on_save: bool,
    concurrent_calculation: bool,
    concurrent_manual_count: Option<u32>,
    force_full_calculation: bool,
}

impl WorkbookCalculationProperties {
    /// Calculation-engine identifier. Excel's effective default is zero.
    pub fn calculation_id(&self) -> u32 {
        self.calculation_id
    }
    pub fn calculation_mode(&self) -> WorkbookCalculationMode {
        self.calculation_mode
    }
    pub fn full_calculation_on_load(&self) -> bool {
        self.full_calculation_on_load
    }
    pub fn reference_mode(&self) -> WorkbookReferenceMode {
        self.reference_mode
    }
    pub fn iterative_calculation(&self) -> bool {
        self.iterative_calculation
    }
    pub fn iteration_count(&self) -> u32 {
        self.iteration_count
    }
    pub fn iteration_delta(&self) -> f64 {
        self.iteration_delta
    }
    pub fn full_precision(&self) -> bool {
        self.full_precision
    }
    pub fn calculation_completed(&self) -> bool {
        self.calculation_completed
    }
    pub fn calculate_on_save(&self) -> bool {
        self.calculate_on_save
    }
    pub fn concurrent_calculation(&self) -> bool {
        self.concurrent_calculation
    }
    pub fn concurrent_manual_count(&self) -> Option<u32> {
        self.concurrent_manual_count
    }
    /// Whether Excel should perform a full calculation on the next calculation cycle.
    pub fn force_full_calculation(&self) -> bool {
        self.force_full_calculation
    }
}

#[derive(Default)]
struct Builder {
    calculation_id: Option<u32>,
    calculation_mode: Option<WorkbookCalculationMode>,
    full_calculation_on_load: Option<bool>,
    reference_mode: Option<WorkbookReferenceMode>,
    iterative_calculation: Option<bool>,
    iteration_count: Option<u32>,
    iteration_delta: Option<f64>,
    full_precision: Option<bool>,
    calculation_completed: Option<bool>,
    calculate_on_save: Option<bool>,
    concurrent_calculation: Option<bool>,
    concurrent_manual_count: Option<u32>,
    force_full_calculation: Option<bool>,
}

impl Builder {
    fn finish(self) -> WorkbookCalculationProperties {
        WorkbookCalculationProperties {
            calculation_id: self.calculation_id.unwrap_or(0),
            calculation_mode: self.calculation_mode.unwrap_or_default(),
            full_calculation_on_load: self.full_calculation_on_load.unwrap_or(false),
            reference_mode: self.reference_mode.unwrap_or_default(),
            iterative_calculation: self.iterative_calculation.unwrap_or(false),
            iteration_count: self.iteration_count.unwrap_or(100),
            iteration_delta: self.iteration_delta.unwrap_or(0.001),
            full_precision: self.full_precision.unwrap_or(true),
            calculation_completed: self.calculation_completed.unwrap_or(true),
            calculate_on_save: self.calculate_on_save.unwrap_or(true),
            concurrent_calculation: self.concurrent_calculation.unwrap_or(true),
            concurrent_manual_count: self.concurrent_manual_count,
            force_full_calculation: self.force_full_calculation.unwrap_or(false),
        }
    }
}

/// Parse the workbook's direct `calcPr` child without executing calculations.
pub fn parse_workbook_calculation_properties(
    xml: &[u8],
) -> Result<Option<WorkbookCalculationProperties>> {
    let processed =
        process_markup_compatibility(xml, &MceCapabilities::default(), &MceLimits::default())?;
    let mut reader = NsReader::from_reader(processed.xml.as_ref());
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut leaf_depth = None;
    let mut properties = None;

    loop {
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        reject_unsafe_event(&event)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let local = element.local_name();
                let core = is_spreadsheetml_name(&namespace, element.name(), local.as_ref());
                if depth == 0 {
                    if root_seen || !core || local.as_ref() != b"workbook" {
                        return Err(invalid("calcPr parser requires a workbook root"));
                    }
                    root_seen = true;
                } else if leaf_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("calcPr is a leaf element"));
                } else if depth == 1 && core && local.as_ref() == b"calcPr" {
                    if properties.is_some() {
                        return Err(invalid("duplicate workbook calcPr element"));
                    }
                    properties = Some(parse_attributes(&element, decoder, &resolver)?.finish());
                    leaf_depth = Some(depth + 1);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("XML depth overflow"))?;
            },
            Event::Empty(element) => {
                let local = element.local_name();
                let core = is_spreadsheetml_name(&namespace, element.name(), local.as_ref());
                if depth == 0 {
                    return Err(invalid("workbook root cannot be empty"));
                }
                if leaf_depth.is_some_and(|value| depth >= value) {
                    return Err(invalid("calcPr is a leaf element"));
                }
                if depth == 1 && core && local.as_ref() == b"calcPr" {
                    if properties.is_some() {
                        return Err(invalid("duplicate workbook calcPr element"));
                    }
                    properties = Some(parse_attributes(&element, decoder, &resolver)?.finish());
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("unexpected workbook end element"));
                }
                if leaf_depth == Some(depth) {
                    leaf_depth = None;
                }
                depth -= 1;
            },
            Event::Text(text)
                if leaf_depth.is_some_and(|value| depth >= value)
                    && !text.decode().map_err(xml_error)?.trim().is_empty() =>
            {
                return Err(invalid("calcPr cannot contain text"));
            },
            Event::CData(_) if leaf_depth.is_some_and(|value| depth >= value) => {
                return Err(invalid("calcPr cannot contain CDATA"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || depth != 0 || leaf_depth.is_some() {
        return Err(invalid("unterminated workbook XML"));
    }
    Ok(properties)
}

fn parse_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Builder> {
    let mut builder = Builder::default();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(invalid(format!(
                "unknown namespaced calcPr attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref()),
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        match local.as_ref() {
            b"calcId" => set_once(
                &mut builder.calculation_id,
                parse_u32(&value, "calcId")?,
                "calcId",
            )?,
            b"calcMode" => set_once(
                &mut builder.calculation_mode,
                WorkbookCalculationMode::parse(&value)?,
                "calcMode",
            )?,
            b"fullCalcOnLoad" => set_once(
                &mut builder.full_calculation_on_load,
                parse_bool(&value, "fullCalcOnLoad")?,
                "fullCalcOnLoad",
            )?,
            b"refMode" => set_once(
                &mut builder.reference_mode,
                WorkbookReferenceMode::parse(&value)?,
                "refMode",
            )?,
            b"iterate" => set_once(
                &mut builder.iterative_calculation,
                parse_bool(&value, "iterate")?,
                "iterate",
            )?,
            b"iterateCount" => set_once(
                &mut builder.iteration_count,
                parse_u32(&value, "iterateCount")?,
                "iterateCount",
            )?,
            b"iterateDelta" => set_once(
                &mut builder.iteration_delta,
                parse_delta(&value)?,
                "iterateDelta",
            )?,
            b"fullPrecision" => set_once(
                &mut builder.full_precision,
                parse_bool(&value, "fullPrecision")?,
                "fullPrecision",
            )?,
            b"calcCompleted" => set_once(
                &mut builder.calculation_completed,
                parse_bool(&value, "calcCompleted")?,
                "calcCompleted",
            )?,
            b"calcOnSave" => set_once(
                &mut builder.calculate_on_save,
                parse_bool(&value, "calcOnSave")?,
                "calcOnSave",
            )?,
            b"concurrentCalc" => set_once(
                &mut builder.concurrent_calculation,
                parse_bool(&value, "concurrentCalc")?,
                "concurrentCalc",
            )?,
            b"concurrentManualCount" => set_once(
                &mut builder.concurrent_manual_count,
                parse_u32(&value, "concurrentManualCount")?,
                "concurrentManualCount",
            )?,
            b"forceFullCalc" => set_once(
                &mut builder.force_full_calculation,
                parse_bool(&value, "forceFullCalc")?,
                "forceFullCalc",
            )?,
            name => {
                return Err(invalid(format!(
                    "unknown calcPr attribute '{}'",
                    String::from_utf8_lossy(name),
                )));
            },
        }
    }
    Ok(builder)
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.is_some() {
        return Err(invalid(format!("duplicate {name} attribute")));
    }
    *slot = Some(value);
    Ok(())
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!("invalid calcPr {name} boolean '{value}'"))),
    }
}

fn parse_u32(value: &str, name: &str) -> Result<u32> {
    value
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid calcPr {name} value '{value}'")))
}

fn parse_delta(value: &str) -> Result<f64> {
    let parsed = value
        .parse::<f64>()
        .map_err(|_| invalid(format!("invalid calcPr iterateDelta '{value}'")))?;
    if !parsed.is_finite() || parsed < 0.0 {
        return Err(invalid(
            "calcPr iterateDelta must be finite and non-negative",
        ));
    }
    Ok(parsed)
}

fn reject_unsafe_event(event: &Event<'_>) -> Result<()> {
    if matches!(event, Event::DocType(_) | Event::PI(_)) {
        return Err(invalid("DTD and processing instructions are rejected"));
    }
    if let Event::GeneralRef(reference) = event {
        let name = reference.decode().map_err(xml_error)?;
        if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") && !name.starts_with('#')
        {
            return Err(invalid("custom XML entities are rejected"));
        }
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    invalid(format!("invalid workbook calcPr XML: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse(child: &str) -> Result<Option<WorkbookCalculationProperties>> {
        parse_workbook_calculation_properties(
            format!(r#"<workbook xmlns="{NS}">{child}</workbook>"#).as_bytes(),
        )
    }

    #[test]
    fn parses_all_attributes_and_effective_defaults() {
        let value = parse(concat!(
            r#"<calcPr calcId="42" calcMode="autoNoTable" fullCalcOnLoad="1" "#,
            r#"refMode="R1C1" iterate="true" iterateCount="250" iterateDelta="1E-4" "#,
            r#"fullPrecision="0" calcCompleted="false" calcOnSave="0" concurrentCalc="false" "#,
            r#"concurrentManualCount="6" forceFullCalc="true"/>"#,
        ))
        .unwrap()
        .unwrap();
        assert_eq!(value.calculation_id(), 42);
        assert_eq!(
            value.calculation_mode(),
            WorkbookCalculationMode::AutomaticExceptTables
        );
        assert!(value.full_calculation_on_load());
        assert_eq!(value.reference_mode(), WorkbookReferenceMode::R1C1);
        assert!(value.iterative_calculation());
        assert_eq!(value.iteration_count(), 250);
        assert_eq!(value.iteration_delta(), 0.0001);
        assert!(!value.full_precision());
        assert!(!value.calculation_completed());
        assert!(!value.calculate_on_save());
        assert!(!value.concurrent_calculation());
        assert_eq!(value.concurrent_manual_count(), Some(6));
        assert!(value.force_full_calculation());

        let defaults = parse("<calcPr/>").unwrap().unwrap();
        assert_eq!(defaults.calculation_id(), 0);
        assert_eq!(
            defaults.calculation_mode(),
            WorkbookCalculationMode::Automatic
        );
        assert!(!defaults.full_calculation_on_load());
        assert_eq!(defaults.reference_mode(), WorkbookReferenceMode::A1);
        assert!(!defaults.iterative_calculation());
        assert_eq!(defaults.iteration_count(), 100);
        assert_eq!(defaults.iteration_delta(), 0.001);
        assert!(defaults.full_precision());
        assert!(defaults.calculation_completed());
        assert!(defaults.calculate_on_save());
        assert!(defaults.concurrent_calculation());
        assert_eq!(defaults.concurrent_manual_count(), None);
        assert!(!defaults.force_full_calculation());
    }

    #[test]
    fn supports_strict_namespace_and_mce_fallback() {
        let strict = br#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><calcPr calcMode="manual"/></workbook>"#;
        assert_eq!(
            parse_workbook_calculation_properties(strict)
                .unwrap()
                .unwrap()
                .calculation_mode(),
            WorkbookCalculationMode::Manual
        );
        let mce = format!(
            concat!(
                r#"<workbook xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported">"#,
                r#"<mc:AlternateContent><mc:Choice Requires="x"><x:calcPr/></mc:Choice><mc:Fallback>"#,
                r#"<calcPr refMode="R1C1"/></mc:Fallback></mc:AlternateContent></workbook>"#,
            ),
            NS
        );
        assert_eq!(
            parse_workbook_calculation_properties(mce.as_bytes())
                .unwrap()
                .unwrap()
                .reference_mode(),
            WorkbookReferenceMode::R1C1
        );
    }

    #[test]
    fn rejects_invalid_values_structure_and_attributes() {
        for child in [
            r#"<calcPr calcMode="sometimes"/>"#,
            r#"<calcPr refMode="A2"/>"#,
            r#"<calcPr iterate="yes"/>"#,
            r#"<calcPr iterateCount="-1"/>"#,
            r#"<calcPr iterateDelta="NaN"/>"#,
            r#"<calcPr iterateDelta="-0.1"/>"#,
            r#"<calcPr mystery="1"/>"#,
            r#"<calcPr><child/></calcPr>"#,
            r#"<wrapper><calcPr calcId="1"/></wrapper>"#,
        ] {
            let result = parse(child);
            if child.starts_with("<wrapper>") {
                assert!(result.unwrap().is_none());
            } else {
                assert!(result.is_err(), "expected rejection for {child}");
            }
        }
        assert!(parse("<calcPr/><calcPr/>").is_err());
        assert!(parse(r#"<calcPr calcId="1" calcId="2"/>"#).is_err());
    }

    fn fixture(bytes: &[u8]) -> WorkbookCalculationProperties {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap();
        parse_workbook_calculation_properties(part.blob())
            .unwrap()
            .unwrap()
    }

    #[test]
    fn reads_poi_calculation_fixtures() {
        let iterative = fixture(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/47889.xlsx"
        )));
        assert_eq!(
            iterative.calculation_mode(),
            WorkbookCalculationMode::Automatic
        );
        assert!(iterative.iterative_calculation());
        assert_eq!(iterative.iteration_count(), 100);
        assert_eq!(iterative.iteration_delta(), 0.001);

        let no_save = fixture(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/58106.xlsx"
        )));
        assert!(!no_save.calculate_on_save());

        let recalculate = fixture(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/60289.xlsx"
        )));
        assert!(recalculate.full_calculation_on_load());
    }

    #[test]
    fn reads_libreoffice_calculation_fixtures() {
        let displayed_precision = fixture(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/totalsRowShown.xlsx"
        )));
        assert!(!displayed_precision.full_precision());

        let r1c1 = fixture(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf134455.xlsx"
        )));
        assert_eq!(r1c1.reference_mode(), WorkbookReferenceMode::R1C1);
        assert_eq!(r1c1.iteration_count(), 100);
        assert_eq!(r1c1.iteration_delta(), 0.001);
    }
}
