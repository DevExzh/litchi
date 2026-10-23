//! Raw OPC chart-graph codec boundary.
//!
//! This module keeps the historical graph capability available for callers
//! that work directly with an [`OpcPackage`]. The package owner uses the
//! canonical `chart::Edit` family; raw callers must provide an explicit
//! admission policy.

pub use super::limit::Limits;
use super::model::{
    CHART_CT, COLOR_STYLE_CT, COLOR_STYLE_REL, EmbeddedWorkbookContentType, Graph, STYLE_CT,
    STYLE_REL,
};
use super::package;
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::part::Part;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::Reader;
use quick_xml::events::Event;

pub use super::model::Resource;

/// Load a selected document's complete chart graph under `limits`.
///
/// The document and every chart-owned XML or workbook payload are admitted
/// before the graph model is allocated. The returned graph retains the exact
/// payloads accepted by the OPC codec.
pub fn load(package: &OpcPackage, document_name: &PackURI, limits: &Limits) -> Result<Graph> {
    admit_package(package, document_name, limits)?;
    package::load_with_limits(package, document_name, limits)
}

/// Replace a selected document's complete chart graph under `limits`.
///
/// All caller-owned graph bytes are admitted before the package codec starts
/// its validation and staging pass.
pub fn store(
    package: &mut OpcPackage,
    document_name: &PackURI,
    graph: &Graph,
    limits: &Limits,
) -> Result<()> {
    admit_package(package, document_name, limits)?;
    admit_graph(graph, limits)?;
    package::store_with_limits(package, document_name, graph, limits)
}

fn admit_package(package: &OpcPackage, document_name: &PackURI, limits: &Limits) -> Result<()> {
    let document = package.get_part(document_name)?;
    admit_xml(
        document.blob(),
        limits.document_xml_bytes(),
        limits,
        crate::chart::LimitKind::DocumentXmlBytes,
    )?;
    let (_, references) = super::codec::document_references_with_limits(document.blob(), limits)?;
    if references.len() > limits.charts() {
        return Err(Error::ChartLimit {
            kind: crate::chart::LimitKind::Count,
            actual: references.len(),
            maximum: limits.charts(),
        });
    }
    let mut total = document.blob().len();
    for relationship_id in references {
        let Some(relationship) = document.rels().get(&relationship_id) else {
            continue;
        };
        let Ok(chart_name) = super::codec::relationship_target(document, relationship) else {
            continue;
        };
        let Ok(chart) = package.get_part(&chart_name) else {
            continue;
        };
        if chart.content_type() != CHART_CT {
            continue;
        }
        admit_xml_part(
            &mut total,
            chart,
            limits.chart_xml_bytes(),
            limits,
            crate::chart::LimitKind::XmlBytes,
        )?;
        let mut companions = 0usize;
        for child in chart.rels().iter() {
            let is_companion = matches!(child.reltype(), STYLE_REL | COLOR_STYLE_REL);
            let is_workbook = child.reltype()
                == super::model::Conformance::Transitional.package_rel()
                || child.reltype() == super::model::Conformance::Strict.package_rel();
            if !is_companion && !is_workbook {
                continue;
            }
            let Ok(target) = super::codec::relationship_target(chart, child) else {
                continue;
            };
            let Ok(part) = package.get_part(&target) else {
                continue;
            };
            if is_companion {
                companions = companions.saturating_add(1);
                if companions > limits.companions() {
                    return Err(Error::ChartLimit {
                        kind: crate::chart::LimitKind::CompanionCount,
                        actual: companions,
                        maximum: limits.companions(),
                    });
                }
                if !matches!(part.content_type(), STYLE_CT | COLOR_STYLE_CT) {
                    continue;
                }
                admit_xml_part(
                    &mut total,
                    part,
                    limits.companion_xml_bytes(),
                    limits,
                    crate::chart::LimitKind::CompanionXmlBytes,
                )?;
            } else if EmbeddedWorkbookContentType::parse(part.content_type()).is_some() {
                if part.blob().len() > limits.workbook_bytes() {
                    return Err(Error::ChartLimit {
                        kind: crate::chart::LimitKind::EmbeddedWorkbookBytes,
                        actual: part.blob().len(),
                        maximum: limits.workbook_bytes(),
                    });
                }
                total = total
                    .checked_add(part.blob().len())
                    .ok_or(Error::ChartLimit {
                        kind: crate::chart::LimitKind::AggregateBytes,
                        actual: usize::MAX,
                        maximum: limits.aggregate_bytes(),
                    })?;
            }
        }
    }
    if total > limits.aggregate_bytes() {
        return Err(Error::ChartLimit {
            kind: crate::chart::LimitKind::AggregateBytes,
            actual: total,
            maximum: limits.aggregate_bytes(),
        });
    }
    Ok(())
}

fn admit_xml_part(
    total: &mut usize,
    part: &dyn Part,
    maximum: usize,
    limits: &Limits,
    kind: crate::chart::LimitKind,
) -> Result<()> {
    admit_xml(part.blob(), maximum, limits, kind)?;
    *total = total
        .checked_add(part.blob().len())
        .ok_or(Error::ChartLimit {
            kind: crate::chart::LimitKind::AggregateBytes,
            actual: usize::MAX,
            maximum: limits.aggregate_bytes(),
        })?;
    Ok(())
}

fn admit_graph(graph: &Graph, limits: &Limits) -> Result<()> {
    if graph.charts.len() > limits.charts() {
        return Err(Error::ChartLimit {
            kind: crate::chart::LimitKind::Count,
            actual: graph.charts.len(),
            maximum: limits.charts(),
        });
    }
    let mut total = 0usize;
    for chart in &graph.charts {
        admit_xml(
            &chart.data,
            limits.chart_xml_bytes(),
            limits,
            crate::chart::LimitKind::XmlBytes,
        )?;
        total = total
            .checked_add(chart.data.len())
            .ok_or(Error::ChartLimit {
                kind: crate::chart::LimitKind::AggregateBytes,
                actual: usize::MAX,
                maximum: limits.aggregate_bytes(),
            })?;
        if chart.styles.len() > limits.companions()
            || chart.color_styles.len() > limits.companions()
        {
            let actual = chart.styles.len().max(chart.color_styles.len());
            return Err(Error::ChartLimit {
                kind: crate::chart::LimitKind::CompanionCount,
                actual,
                maximum: limits.companions(),
            });
        }
        for companion in chart.styles.iter().chain(&chart.color_styles) {
            admit_xml(
                &companion.data,
                limits.companion_xml_bytes(),
                limits,
                crate::chart::LimitKind::CompanionXmlBytes,
            )?;
            total = total
                .checked_add(companion.data.len())
                .ok_or(Error::ChartLimit {
                    kind: crate::chart::LimitKind::AggregateBytes,
                    actual: usize::MAX,
                    maximum: limits.aggregate_bytes(),
                })?;
        }
        if let Some(workbook) = &chart.workbook {
            if workbook.data.len() > limits.workbook_bytes() {
                return Err(Error::ChartLimit {
                    kind: crate::chart::LimitKind::EmbeddedWorkbookBytes,
                    actual: workbook.data.len(),
                    maximum: limits.workbook_bytes(),
                });
            }
            total = total
                .checked_add(workbook.data.len())
                .ok_or(Error::ChartLimit {
                    kind: crate::chart::LimitKind::AggregateBytes,
                    actual: usize::MAX,
                    maximum: limits.aggregate_bytes(),
                })?;
        }
    }
    if total > limits.aggregate_bytes() {
        return Err(Error::ChartLimit {
            kind: crate::chart::LimitKind::AggregateBytes,
            actual: total,
            maximum: limits.aggregate_bytes(),
        });
    }
    Ok(())
}

fn admit_xml(
    xml: &[u8],
    maximum: usize,
    limits: &Limits,
    byte_kind: crate::chart::LimitKind,
) -> Result<()> {
    if xml.len() > maximum {
        return Err(Error::ChartLimit {
            kind: byte_kind,
            actual: xml.len(),
            maximum,
        });
    }
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut attributes = 0usize;
    let mut attribute_bytes = 0usize;
    loop {
        match reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(element) => {
                depth = depth.saturating_add(1);
                nodes = nodes.saturating_add(1);
                check_structure(depth, nodes, limits)?;
                check_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
            },
            Event::Empty(element) => {
                nodes = nodes.saturating_add(1);
                check_structure(depth.saturating_add(1), nodes, limits)?;
                check_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("chart XML has an unexpected closing element".to_owned())
                })?
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::Invalid(
                    "chart XML contains a forbidden declaration".to_owned(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::Invalid("chart XML is unterminated".to_owned()));
    }
    Ok(())
}

fn check_structure(depth: usize, nodes: usize, limits: &Limits) -> Result<()> {
    if depth > limits.depth() {
        return Err(Error::ChartLimit {
            kind: crate::chart::LimitKind::XmlStructure,
            actual: depth,
            maximum: limits.depth(),
        });
    }
    if nodes > limits.nodes() {
        return Err(Error::ChartLimit {
            kind: crate::chart::LimitKind::XmlStructure,
            actual: nodes,
            maximum: limits.nodes(),
        });
    }
    Ok(())
}

fn check_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    attributes: &mut usize,
    attribute_bytes: &mut usize,
    limits: &Limits,
) -> Result<()> {
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        *attributes = attributes.saturating_add(1);
        *attribute_bytes = attribute_bytes
            .checked_add(attribute.key.as_ref().len() + attribute.value.len())
            .ok_or(Error::ChartLimit {
                kind: crate::chart::LimitKind::XmlAttributeBytes,
                actual: usize::MAX,
                maximum: limits.attribute_bytes(),
            })?;
        if *attributes > limits.attributes() {
            return Err(Error::ChartLimit {
                kind: crate::chart::LimitKind::XmlAttributeCount,
                actual: *attributes,
                maximum: limits.attributes(),
            });
        }
        if *attribute_bytes > limits.attribute_bytes() {
            return Err(Error::ChartLimit {
                kind: crate::chart::LimitKind::XmlAttributeBytes,
                actual: *attribute_bytes,
                maximum: limits.attribute_bytes(),
            });
        }
    }
    Ok(())
}
