//! Inline chart insertion and chart-part edits staged with a main-document
//! edit (record 0283).
//!
//! A new chart is an authored [`Spec`]: its chart part (a `c:chartSpace`, or
//! for an extended family a `cx:chartSpace` with its chart-style and
//! color-style companions), a deterministic embedded workbook holding the
//! data the chart refers to, and a run holding the inline drawing — all one
//! graph change and one document operation, published by one commit.

use std::sync::Arc;

use litchi_core::Position;
use litchi_drawingml::chart::authoring::{Binding, Spec};
use litchi_drawingml::chart::chartex;
use litchi_opc::PackURI;
use litchi_opc::constants::relationship_type as rt;

use super::super::transaction::{ChartMarkup, GraphChange, ReplacedPart};
use super::super::{ElementPath, Error, Refusal, Result};
use super::Edit;
use crate::Package;
use crate::document::{TransferPart, TransferRelationship};

/// The default extent of a new chart: 15 cm × 10 cm (OfficeCLI's default).
const DEFAULT_WIDTH: i64 = 5_400_000;
const DEFAULT_HEIGHT: i64 = 3_600_000;
/// The largest accepted extent: 22 inches, Word's page-size limit.
const MAX_EXTENT: i64 = 22 * 914_400;
const WORKBOOK_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const STRICT_CHART: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/chart";
const STRICT_PACKAGE: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/package";

/// A chart to insert: its authored spec, an extent in EMUs, and its frame
/// name and description (alternative text).
///
/// ```
/// use litchi_docx::document::NewChart;
/// use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
///
/// let spec = Spec::new(Variant::new(Kind::Column))
///     .with_series(Series::new("East", vec![1.0, 2.0]));
/// let chart = NewChart::new(spec).with_size(Some(5_400_000), None).with_name("Sales");
/// assert_eq!(chart.size(), (5_400_000, 3_600_000));
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct NewChart {
    spec: Spec,
    width: Option<i64>,
    height: Option<i64>,
    name: Option<String>,
    description: String,
}

impl NewChart {
    /// A chart of `spec`, 15 cm × 10 cm.
    #[must_use]
    pub const fn new(spec: Spec) -> Self {
        Self {
            spec,
            width: None,
            height: None,
            name: None,
            description: String::new(),
        }
    }

    /// The extent in EMUs; a dimension left `None` keeps its default.
    #[must_use]
    pub const fn with_size(mut self, width: Option<i64>, height: Option<i64>) -> Self {
        self.width = width;
        self.height = height;
        self
    }

    /// The drawing object's name (`Chart N` by default).
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// The alternative text.
    #[must_use]
    pub fn with_description(mut self, description: impl Into<String>) -> Self {
        self.description = description.into();
        self
    }

    /// The chart's spec.
    #[must_use]
    pub const fn spec(&self) -> &Spec {
        &self.spec
    }

    /// The extent in EMUs, defaults filled in.
    #[must_use]
    pub fn size(&self) -> (i64, i64) {
        (
            self.width.unwrap_or(DEFAULT_WIDTH),
            self.height.unwrap_or(DEFAULT_HEIGHT),
        )
    }
}

/// The relationship ID and part name of one new related part.
struct Planned {
    name: String,
    target: String,
}

fn invalid() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::InvalidPropertyValue,
    }
}

fn drawing_error(error: litchi_drawingml::Error) -> Error {
    match error {
        litchi_drawingml::Error::Invalid(_) | litchi_drawingml::Error::Xml(_) => invalid(),
        error => Error::Document(crate::Error::Drawing(error)),
    }
}

impl Edit<Package> {
    /// Insert `chart` as a new run of the paragraph at `paragraph`, before
    /// its direct run `at` (after its last child for `None`), and return its
    /// position among the paragraph's charts (the order of
    /// [`crate::Paragraph::charts`]).
    ///
    /// The chart part is staged under `charts/` with a `chart` (or, for an
    /// extended family, `chartEx`) relationship of the main part; it refers
    /// to an embedded workbook under `embeddings/` holding the spec's data
    /// on `Sheet1` ([`Spec::workbook_members`]). An extended chart also gets
    /// its chart-style and color-style parts. Every part is generated
    /// deterministically.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::Package;
    /// use litchi_docx::document::{ElementPath, ElementStep, NewChart};
    /// use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Sales: ");
    /// let package = builder.build()?;
    /// let spec = Spec::new(Variant::new(Kind::Column))
    ///     .with_categories(["Q1", "Q2"])
    ///     .with_series(Series::new("East", vec![1.0, 2.0]));
    /// let mut edit = package.edit()?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// edit.insert_chart_at(&paragraph, None, &NewChart::new(spec))?;
    /// let mut next = edit.commit()?.into_package();
    /// let reopened = Package::from_bytes(&next.to_bytes()?)?;
    /// let charts = reopened.document()?.paragraphs()?[0].charts()?;
    /// assert_eq!(charts.len(), 1);
    /// assert!(reopened.chart_xml(charts[0].relationship_id())?.is_some());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidPropertyValue`] for an invalid spec or an
    /// extent outside 1 EMU to 22 inches, the refusals of
    /// [`Self::insert_run_at`], or a resource-limit error.
    pub fn insert_chart_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        chart: &NewChart,
    ) -> Result<Position> {
        let (width, height) = chart.size();
        if !(1..=MAX_EXTENT).contains(&width) || !(1..=MAX_EXTENT).contains(&height) {
            return Err(invalid());
        }
        let spec = chart.spec.clone().with_binding(Binding::sheet("Sheet1"));
        let extended = spec.variant().kind.is_extended();
        let workbook = workbook_bytes(&spec)?;
        let id = self
            .inner_mut()?
            .max_drawing_id()?
            .map_or(Some(1), |maximum| maximum.checked_add(1))
            .ok_or_else(invalid)?;
        let strict = self.is_strict()?;
        let base = self.main_folder()?;
        let used = self.part_names()?;
        let free = |directory: &str, stem: &str, extension: &str, taken: &[String]| {
            (1_usize..)
                .map(|index| {
                    let target = format!("{directory}{stem}{index}{extension}");
                    (format!("{base}{target}"), target)
                })
                .find(|(name, _)| !used.contains(name) && !taken.contains(name))
                .map(|(name, target)| Planned { name, target })
        };
        let chart_part = free(
            "charts/",
            if extended { "chartEx" } else { "chart" },
            ".xml",
            &[],
        )
        .ok_or_else(invalid)?;
        let workbook_part = free(
            "embeddings/",
            "Microsoft_Excel_Worksheet",
            ".xlsx",
            std::slice::from_ref(&chart_part.name),
        )
        .ok_or_else(invalid)?;
        // Targets from the chart part, which sits in `charts/`.
        let from_chart = |planned: &Planned| -> Result<String> {
            let name = PackURI::new(planned.name.clone()).map_err(crate::Error::Invalid)?;
            let chart_name =
                PackURI::new(chart_part.name.clone()).map_err(crate::Error::Invalid)?;
            Ok(name.relative_ref(chart_name.base_uri()))
        };
        let mut chart_relationships = vec![TransferRelationship {
            id: "rId1".to_owned(),
            relationship_type: if strict { STRICT_PACKAGE } else { rt::PACKAGE }.to_owned(),
            target: from_chart(&workbook_part)?,
            external: false,
        }];
        let mut parts = vec![TransferPart {
            name: workbook_part.name.clone(),
            content_type: WORKBOOK_CONTENT_TYPE.to_owned(),
            blob: Arc::new(workbook),
            relationships: Arc::new([]),
        }];
        if extended {
            let mut taken = vec![chart_part.name.clone(), workbook_part.name.clone()];
            for (index, (stem, content_type, relationship_type, bytes)) in [
                (
                    "style",
                    chartex::STYLE_CONTENT_TYPE,
                    chartex::STYLE_RELATIONSHIP_TYPE,
                    chartex::style_xml(),
                ),
                (
                    "colors",
                    chartex::COLORS_CONTENT_TYPE,
                    chartex::COLORS_RELATIONSHIP_TYPE,
                    chartex::colors_xml(),
                ),
            ]
            .into_iter()
            .enumerate()
            {
                let planned = free("charts/", stem, ".xml", &taken).ok_or_else(invalid)?;
                taken.push(planned.name.clone());
                chart_relationships.push(TransferRelationship {
                    id: format!("rId{}", index + 2),
                    relationship_type: relationship_type.to_owned(),
                    target: from_chart(&planned)?,
                    external: false,
                });
                parts.push(TransferPart {
                    name: planned.name,
                    content_type: content_type.to_owned(),
                    blob: Arc::new(bytes),
                    relationships: Arc::new([]),
                });
            }
        }
        let xml = spec.to_xml(Some("rId1")).map_err(|_| invalid())?;
        let (content_type, relationship_type) = if extended {
            (chartex::CONTENT_TYPE, chartex::RELATIONSHIP_TYPE)
        } else {
            (
                crate::chart::inline::CHART_CONTENT_TYPE,
                if strict { STRICT_CHART } else { rt::CHART },
            )
        };
        parts.insert(
            0,
            TransferPart {
                name: chart_part.name.clone(),
                content_type: content_type.to_owned(),
                blob: Arc::new(xml),
                relationships: chart_relationships.into(),
            },
        );
        for part in &parts {
            PackURI::new(part.name.clone()).map_err(crate::Error::Invalid)?;
        }
        let relationship_id = self.next_relationship_id()?;
        let name = chart.name.clone().unwrap_or_else(|| format!("Chart {id}"));
        // The drawing first: a refused target stages nothing.
        let position = self.inner_mut()?.insert_chart_markup_at(
            paragraph,
            at,
            &relationship_id,
            &ChartMarkup {
                width,
                height,
                id,
                name: &name,
                description: &chart.description,
                extended,
            },
        )?;
        self.inner_mut()?.stage_graph_change(
            Arc::new(GraphChange {
                relationships: Arc::new([TransferRelationship {
                    id: relationship_id,
                    relationship_type: relationship_type.to_owned(),
                    target: chart_part.target,
                    external: false,
                }]),
                parts: parts.into(),
                replaced: Arc::new([]),
            }),
            true,
        )?;
        Ok(position)
    }

    /// The chart part the main document relates through `relationship_id`
    /// and, when it has one, its embedded workbook, as the staged changes
    /// leave them; `None` when the relationship names no chart.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error for an unreadable part.
    pub fn chart_parts(&self, relationship_id: &str) -> Result<Option<(Vec<u8>, Option<Vec<u8>>)>> {
        let Some((chart, workbook)) = self.chart_names(relationship_id)? else {
            return Ok(None);
        };
        let xml = self
            .part_payload(&chart)?
            .ok_or_else(|| crate::Error::PartNotFound(chart.clone()))?;
        let workbook = match workbook {
            Some(name) => Some(
                self.part_payload(&name)?
                    .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?
                    .as_ref()
                    .clone(),
            ),
            None => None,
        };
        Ok(Some((xml.as_ref().clone(), workbook)))
    }

    /// Replace the chart part the main document relates through
    /// `relationship_id` and, when given, the embedded workbook it already
    /// refers to. The new part must read back as a chart of the same family
    /// (ordinary or extended). Returns whether anything changed; identical
    /// bytes are an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidPropertyValue`] for XML that is not a chart
    /// of the part's family or a workbook for a chart that has none, and a
    /// package-graph error for a missing chart.
    pub fn set_chart(
        &mut self,
        relationship_id: &str,
        xml: Vec<u8>,
        workbook: Option<Vec<u8>>,
    ) -> Result<bool> {
        let (chart, workbook_name) = self
            .chart_names(relationship_id)?
            .ok_or_else(|| crate::Error::PartNotFound(relationship_id.to_owned()))?;
        let current = self
            .part_payload(&chart)?
            .ok_or_else(|| crate::Error::PartNotFound(chart.clone()))?;
        let extended = chartex::is_chartex(&current);
        if chartex::is_chartex(&xml) != extended {
            return Err(invalid());
        }
        if extended {
            chartex::read(&xml, &litchi_drawingml::chart::Limits::DEFAULT)
                .map_err(drawing_error)?;
        } else {
            litchi_drawingml::chart::reader::read(
                &xml[..],
                &litchi_drawingml::chart::Limits::DEFAULT,
            )
            .map_err(drawing_error)?;
        }
        let mut replaced = Vec::new();
        if current.as_slice() != xml.as_slice() {
            replaced.push(ReplacedPart {
                name: chart,
                before: current,
                after: Arc::new(xml),
            });
        }
        if let Some(bytes) = workbook {
            let name = workbook_name.ok_or_else(invalid)?;
            let before = self
                .part_payload(&name)?
                .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?;
            if before.as_slice() != bytes.as_slice() {
                replaced.push(ReplacedPart {
                    name,
                    before,
                    after: Arc::new(bytes),
                });
            }
        }
        if replaced.is_empty() {
            return Ok(false);
        }
        self.inner_mut()?.stage_graph_change(
            Arc::new(GraphChange {
                relationships: Arc::new([]),
                parts: Arc::new([]),
                replaced: replaced.into(),
            }),
            true,
        )?;
        Ok(true)
    }

    /// Resize the chart at `chart` among the paragraph's charts to
    /// `(width, height)` EMUs. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidPropertyValue`] for an extent outside 1 EMU
    /// to 22 inches and [`Refusal::RunNotFound`] for a missing chart.
    pub fn set_chart_extent_at(
        &mut self,
        paragraph: &ElementPath,
        chart: Position,
        (width, height): (i64, i64),
    ) -> Result<bool> {
        if !(1..=MAX_EXTENT).contains(&width) || !(1..=MAX_EXTENT).contains(&height) {
            return Err(invalid());
        }
        self.inner_mut()?
            .rewrite_chart_at(paragraph, chart, Some((width, height)))
    }

    /// Remove the chart at `chart` among the paragraph's charts: its
    /// drawing, and its run when the run holds nothing else. The chart part
    /// and its relationship stay in the package, unreferenced (the same
    /// policy as a removed picture).
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::RunNotFound`] for a missing chart.
    pub fn remove_chart_at(&mut self, paragraph: &ElementPath, chart: Position) -> Result<()> {
        self.inner_mut()?.rewrite_chart_at(paragraph, chart, None)?;
        Ok(())
    }

    /// The chart part a main relationship names, and its embedded workbook
    /// part when it has one, staged changes included.
    fn chart_names(&self, relationship_id: &str) -> Result<Option<(String, Option<String>)>> {
        let main = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?;
        let base = main.partname().base_uri().to_owned();
        let mut relationship = main.rels().get(relationship_id).map(|relationship| {
            (
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.is_external(),
            )
        });
        self.inner_ref(|inner| {
            for (change, forward) in inner.staged_graph_changes() {
                if !forward {
                    continue;
                }
                for staged in change
                    .relationships
                    .iter()
                    .filter(|staged| staged.id == relationship_id)
                {
                    relationship = Some((
                        staged.relationship_type.clone(),
                        staged.target.clone(),
                        staged.external,
                    ));
                }
            }
        })?;
        let Some((relationship_type, target, external)) = relationship else {
            return Ok(None);
        };
        if crate::chart::inline::chart_content_type(&relationship_type).is_none() {
            return Ok(None);
        }
        if external {
            return Err(invalid());
        }
        let chart = PackURI::from_rel_ref(&base, &target)
            .map_err(crate::Error::Invalid)?
            .as_str()
            .to_owned();
        let chart_uri = PackURI::new(chart.clone()).map_err(crate::Error::Invalid)?;
        let is_package = |reltype: &str| reltype == rt::PACKAGE || reltype == STRICT_PACKAGE;
        let mut workbook = None;
        if let Ok(part) = self.package.opc_package().get_part(&chart_uri) {
            for child in part.rels().iter() {
                if !child.is_external() && is_package(child.reltype()) {
                    workbook = Some(
                        child
                            .target_partname()
                            .map_err(crate::Error::from)?
                            .as_str()
                            .to_owned(),
                    );
                }
            }
        }
        self.inner_ref(|inner| {
            for (change, forward) in inner.staged_graph_changes() {
                if !forward {
                    continue;
                }
                for part in change.parts.iter().filter(|part| part.name == chart) {
                    for child in part.relationships.iter() {
                        if !child.external && is_package(&child.relationship_type) {
                            workbook = Some(
                                PackURI::from_rel_ref(chart_uri.base_uri(), &child.target)
                                    .map_err(crate::Error::Invalid)?
                                    .as_str()
                                    .to_owned(),
                            );
                        }
                    }
                }
            }
            Ok::<_, crate::Error>(())
        })??;
        Ok(Some((chart, workbook)))
    }
}

/// The embedded workbook of an authored chart: the spec's workbook
/// members ([`Spec::workbook_members`]), zipped. [`Edit::insert_chart_at`]
/// embeds this workbook; pass it to [`Edit::set_chart`] after changing a
/// chart's data so the workbook keeps matching the chart.
///
/// ```
/// use litchi_drawingml::chart::authoring::{Binding, Kind, Series, Spec, Variant};
///
/// let spec = Spec::new(Variant::new(Kind::Column))
///     .with_series(Series::new("East", vec![1.0]))
///     .with_binding(Binding::sheet("Sheet1"));
/// let bytes = litchi_docx::document::chart_workbook(&spec)?;
/// assert!(bytes.starts_with(b"PK"));
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
///
/// # Errors
///
/// Returns [`Refusal::InvalidPropertyValue`] for an invalid spec.
pub fn chart_workbook(spec: &Spec) -> Result<Vec<u8>> {
    workbook_bytes(spec)
}

fn workbook_bytes(spec: &Spec) -> Result<Vec<u8>> {
    let members = spec.workbook_members().map_err(|_| invalid())?;
    let mut writer = litchi_opc::phys_pkg::PhysPkgWriter::new();
    for (name, bytes) in members {
        let name = PackURI::new(name).map_err(crate::Error::Invalid)?;
        writer.write(&name, &bytes).map_err(crate::Error::from)?;
    }
    Ok(writer.finish().map_err(crate::Error::from)?)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::document::ElementStep;
    use litchi_drawingml::chart::authoring::{self, Kind, Series, Variant};

    fn spec(kind: Kind) -> Spec {
        Spec::new(Variant::new(kind))
            .with_title("Sales")
            .with_categories(["Q1", "Q2", "Q3"])
            .with_series(Series::new("East", vec![10.0, -4.0, 6.0]))
    }

    fn package() -> Package {
        let mut builder = crate::Builder::new().unwrap();
        builder
            .document()
            .unwrap()
            .add_paragraph_with_text("Charts: ");
        builder.build().unwrap()
    }

    #[test]
    fn charts_insert_read_edit_and_remove() {
        let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
        let package = package();
        let mut edit = package.edit().unwrap();
        let first = edit
            .insert_chart_at(
                &paragraph,
                None,
                &NewChart::new(spec(Kind::Column)).with_name("A"),
            )
            .unwrap();
        let second = edit
            .insert_chart_at(&paragraph, None, &NewChart::new(spec(Kind::Waterfall)))
            .unwrap();
        assert_eq!((first.get(), second.get()), (0, 1));
        // Staged reads see the new parts.
        let charts = {
            let projected = edit.projected().unwrap();
            crate::chart::inline::parse_inline_charts(projected.xml_bytes()).unwrap()
        };
        let (xml, workbook) = edit
            .chart_parts(charts[1].relationship_id())
            .unwrap()
            .unwrap();
        assert!(chartex::is_chartex(&xml));
        assert!(workbook.unwrap().starts_with(b"PK"));
        let next = edit.commit().unwrap().into_package();
        // Deterministic.
        let mut again = package.edit().unwrap();
        again
            .insert_chart_at(
                &paragraph,
                None,
                &NewChart::new(spec(Kind::Column)).with_name("A"),
            )
            .unwrap();
        again
            .insert_chart_at(&paragraph, None, &NewChart::new(spec(Kind::Waterfall)))
            .unwrap();
        let bytes = next.to_bytes().unwrap();
        assert_eq!(
            bytes,
            again.commit().unwrap().into_package().to_bytes().unwrap()
        );

        let reopened = Package::from_bytes(&bytes).unwrap();
        let charts = reopened.document().unwrap().paragraphs().unwrap()[0]
            .charts()
            .unwrap();
        assert_eq!(charts.len(), 2);
        assert_eq!(charts[0].name(), "A");
        assert_eq!(
            (charts[0].width_emu(), charts[0].height_emu()),
            (5_400_000, 3_600_000)
        );
        assert!(charts[1].is_extended());
        let classic = reopened
            .chart_xml(charts[0].relationship_id())
            .unwrap()
            .unwrap();
        let parsed = authoring::read(&classic, &litchi_drawingml::chart::Limits::DEFAULT).unwrap();
        assert_eq!(parsed.spec.variant().kind, Kind::Column);
        assert!(parsed.is_authored(&classic, &litchi_drawingml::chart::Limits::DEFAULT,));
        let source = crate::source_backed::Package::from_bytes(&bytes).unwrap();
        assert_eq!(
            source
                .chart_xml(charts[1].relationship_id())
                .unwrap()
                .unwrap(),
            reopened
                .chart_xml(charts[1].relationship_id())
                .unwrap()
                .unwrap()
        );

        // Replace, resize and remove.
        let mut edit = reopened.edit().unwrap();
        let retitled = parsed
            .spec
            .clone()
            .with_title("Revenue")
            .to_xml(parsed.external_data.as_deref())
            .unwrap();
        assert!(
            !edit
                .set_chart(charts[0].relationship_id(), classic.clone(), None)
                .unwrap()
        );
        assert!(
            edit.set_chart(charts[0].relationship_id(), retitled.clone(), None)
                .unwrap()
        );
        // A chart part keeps its family.
        assert!(
            edit.set_chart(charts[1].relationship_id(), retitled.clone(), None)
                .is_err()
        );
        assert!(
            edit.set_chart_extent_at(&paragraph, Position::new(1), (100, 200))
                .unwrap()
        );
        assert!(
            !edit
                .set_chart_extent_at(&paragraph, Position::new(1), (100, 200))
                .unwrap()
        );
        assert!(
            edit.set_chart_extent_at(&paragraph, Position::new(1), (0, 200))
                .is_err()
        );
        edit.remove_chart_at(&paragraph, Position::new(0)).unwrap();
        assert!(edit.remove_chart_at(&paragraph, Position::new(5)).is_err());
        let after = edit.commit().unwrap().into_package();
        let charts = after.document().unwrap().paragraphs().unwrap()[0]
            .charts()
            .unwrap();
        assert_eq!(charts.len(), 1);
        assert_eq!((charts[0].width_emu(), charts[0].height_emu()), (100, 200));
        assert_eq!(
            after.document().unwrap().paragraphs().unwrap()[0]
                .text()
                .unwrap(),
            "Charts: "
        );
    }

    #[test]
    fn invalid_charts_are_refused() {
        let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
        let package = package();
        let mut edit = package.edit().unwrap();
        let empty = Spec::new(Variant::new(Kind::Column));
        assert!(
            edit.insert_chart_at(&paragraph, None, &NewChart::new(empty))
                .is_err()
        );
        let huge = NewChart::new(spec(Kind::Column)).with_size(Some(MAX_EXTENT + 1), None);
        assert!(edit.insert_chart_at(&paragraph, None, &huge).is_err());
        let missing = ElementPath::new([ElementStep::Paragraph(Position::new(9))]);
        assert!(
            edit.insert_chart_at(&missing, None, &NewChart::new(spec(Kind::Bar)))
                .is_err()
        );
        assert!(edit.commit().unwrap().patch().is_noop());
    }
}
