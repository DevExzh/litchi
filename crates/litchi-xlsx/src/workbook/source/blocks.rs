//! Worksheet reads served from the blocks outside `sheetData`.
//!
//! Data validations, conditional formats, the auto-filter, hyperlinks and
//! sheet views live in the worksheet part beside `sheetData`, which usually
//! carries almost all of the part's bytes. Each of their parsers walks the
//! whole document it is given, and counts every event against a
//! 1,000,000-event ceiling, so handing them the complete worksheet makes
//! every such read cost a pass over the cell data and lets a large sheet
//! refuse a read even when it has none of these blocks.
//!
//! The source-backed worksheet instead scans its part once, locates the
//! content of the direct `sheetData` child of the worksheet root, and
//! retains the part with that content removed: the bytes before it, the
//! `sheetData` start tag, its end tag and everything after it, unchanged.
//! The projection is a weighted clean value of the workbook's cache. Each
//! read then parses the projection with the same parser the owning workbook
//! uses on the complete part, so for a worksheet whose `sheetData` holds only
//! rows the values are identical, and the event ceilings apply to the blocks
//! themselves.
//!
//! The cells are not part of these reads: `sheetData` is validated by the
//! cell reads of [`super::SourceWorksheet`]. Markup nested inside `sheetData`
//! is therefore not seen here, and a payload whose `sheetData` cannot be
//! located (no direct `sheetData` child, or XML the locating scan cannot
//! read) is retained whole, so its reads see exactly the complete part.

use std::ops::Range;
use std::sync::Arc;

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::{OpcError, SourceLimitKind, SourceLimits};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::values::{Axis, Cached};
use super::{SourceInner, SourceWorksheet, WorksheetKind};
use crate::error::{Error, Result, allocation, invalid};
use crate::raw::namespace::{SPREADSHEETML_NAMESPACE, STRICT_SPREADSHEETML_NAMESPACE};

/// The retained projection of one worksheet part.
pub(super) struct Blocks {
    xml: Box<[u8]>,
    /// Whether the content of `sheetData` was removed.
    #[cfg(test)]
    elided: bool,
}

impl Blocks {
    /// Project a worksheet payload onto its bytes outside `sheetData`.
    fn project(xml: &[u8], limits: SourceLimits) -> Result<Self> {
        let content = sheet_data_content(xml, limits)?;
        let kept = content
            .as_ref()
            .map_or(xml.len(), |content| xml.len() - content.len());
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(kept)
            .map_err(|source| allocation("worksheet blocks outside sheetData", source))?;
        match &content {
            Some(content) => {
                bytes.extend_from_slice(&xml[..content.start]);
                bytes.extend_from_slice(&xml[content.end..]);
            },
            None => bytes.extend_from_slice(xml),
        }
        Ok(Self {
            xml: bytes.into_boxed_slice(),
            #[cfg(test)]
            elided: content.is_some(),
        })
    }

    pub(super) fn xml(&self) -> &[u8] {
        &self.xml
    }

    /// Logical retained weight in bytes.
    pub(super) fn weight_bytes(&self) -> usize {
        size_of::<Self>().saturating_add(self.xml.len())
    }
}

/// Locate the content of the worksheet root's direct `sheetData` child.
///
/// The returned range runs from just after the `sheetData` start tag to just
/// before its end tag. `None` means there is nothing to remove: no direct
/// `sheetData` child before the root closes, an empty `<sheetData/>`, or XML
/// this scan cannot read. The whole scan checks the caller's physical XML
/// budgets, including cell events; subsequent block parsing counts only the
/// projected events. Malformed roots and multiple documents remain whole.
fn sheet_data_content(xml: &[u8], limits: SourceLimits) -> Result<Option<Range<usize>>> {
    // Keep malformed or unsupported structure whole for the block parser's
    // authoritative refusal. Physical source budgets include elided cells.
    const BLOCK_DEPTH: usize = 128;
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut roots = 0usize;
    let mut worksheet_root = false;
    let mut content_start = None;
    let mut content = None;
    loop {
        let before = reader.buffer_position() as usize;
        let (namespace, event) = match reader.read_resolved_event() {
            Ok(event) => event,
            Err(_) => return Ok(None),
        };
        if matches!(event, Event::Eof) {
            break;
        }
        events = events.saturating_add(1);
        check_limit(SourceLimitKind::Events, events, limits.max_events())?;
        if let Event::Start(start) | Event::Empty(start) = &event {
            check_limit(
                SourceLimitKind::Depth,
                depth.saturating_add(1),
                limits.max_depth(),
            )?;
            if depth >= BLOCK_DEPTH {
                return Ok(None);
            }
            for attribute in start.checked_attributes() {
                let Ok(attribute) = attribute else {
                    return Ok(None);
                };
                check_limit(
                    SourceLimitKind::XmlAttributeBytes,
                    attribute.value.len(),
                    limits.max_xml_attribute_bytes(),
                )?;
            }
            if depth == 0 {
                roots += 1;
                worksheet_root =
                    start.local_name().as_ref() == b"worksheet" && is_spreadsheetml(&namespace);
            }
        }
        match event {
            Event::Start(start) => {
                if depth == 1
                    && worksheet_root
                    && start.local_name().as_ref() == b"sheetData"
                    && is_spreadsheetml(&namespace)
                {
                    if content.is_some() {
                        return Ok(None);
                    }
                    content_start = Some(reader.buffer_position() as usize);
                }
                depth += 1;
            },
            Event::End(_) => {
                if depth == 2 {
                    if let Some(start) = content_start.take() {
                        content = Some(start..before);
                    }
                }
                let Some(parent) = depth.checked_sub(1) else {
                    return Ok(None);
                };
                depth = parent;
            },
            _ => {},
        }
    }
    Ok(if roots == 1 && worksheet_root && depth == 0 {
        content
    } else {
        None
    })
}

fn check_limit(kind: SourceLimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        Err(Error::Package(OpcError::Limit {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        }))
    } else {
        Ok(())
    }
}

fn is_spreadsheetml(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == SPREADSHEETML_NAMESPACE || *value == STRICT_SPREADSHEETML_NAMESPACE
    )
}

impl SourceInner {
    /// Number of differential formats in the workbook styles part, which
    /// conditional-format rules index. It is parsed once per workbook.
    fn differential_format_count(&self) -> Result<usize> {
        self.execution_check()?;
        let Some(uri) = self.styles_uri.as_ref() else {
            return Ok(0);
        };
        let part = self.package.part(uri)?;
        if let Some(count) = self.dependencies().differential_format_count.get() {
            self.execution_check()?;
            self.package.source_version()?;
            return Ok(*count);
        }
        let data = part.data()?;
        let count =
            crate::conditional_formatting::parse_differential_formats(data.as_bytes())?.len();
        self.execution_check()?;
        self.package.source_version()?;
        let _publish_result = self.dependencies().differential_format_count.set(count);
        Ok(count)
    }
}

impl crate::workbook::source::SourceBackedWorkbook {
    /// The workbook's people, the authors and mentions of threaded comments
    /// (record 0322); `None` without a persons part.
    ///
    /// # Errors
    ///
    /// Returns a typed parse error for a malformed persons part, and the
    /// source, cancellation and limit errors of every source-backed read.
    pub fn people(&self) -> Result<Option<crate::threaded_comments::People>> {
        let inner = &self.view()?.inner;
        inner.execution_check()?;
        let value = crate::threaded_comments::package::source_people(&inner.package)?;
        inner.execution_check()?;
        inner.package.source_version()?;
        Ok(value)
    }

    /// Parse the styles part's differential formats, the formats conditional
    /// formatting rules reference by `dxfId`; empty without a styles part.
    ///
    /// # Errors
    ///
    /// Returns a typed parse error for malformed `dxfs` markup, and the
    /// source, cancellation and limit errors of every source-backed read.
    pub fn differential_formats(&self) -> Result<Vec<crate::conditional_formatting::Differential>> {
        let inner = &self.view()?.inner;
        inner.execution_check()?;
        let Some(uri) = inner.styles_uri.as_ref() else {
            return Ok(Vec::new());
        };
        let part = inner.package.part(uri)?;
        let data = part.data()?;
        let values = crate::conditional_formatting::parse_differential_formats(data.as_bytes())?;
        inner.execution_check()?;
        inner.package.source_version()?;
        Ok(values)
    }
}

impl SourceWorksheet {
    /// Parse this worksheet's data-validation collections.
    ///
    /// The first block read of a worksheet scans its part once and retains
    /// the XML outside `sheetData` in the workbook's semantic cache; later
    /// block reads parse only that projection. Values are those of
    /// [`crate::Worksheet::data_validations`] for the same part.
    ///
    /// # Errors
    ///
    /// [`Error::NotWorksheet`] for a chart, dialog, or other non-grid sheet;
    /// a typed parse error for malformed validation markup; and the source,
    /// cancellation, and limit errors of every source-backed read.
    pub fn data_validations(&self) -> Result<Vec<crate::data_validation::Collection>> {
        let value = self.blocks().and_then(|blocks| {
            crate::data_validation::parse_data_validation_collections(blocks.xml())
        });
        self.finish_result(value)
    }

    /// Parse every conditional-formatting container and associate its
    /// differential formats with the workbook styles part.
    ///
    /// Served from the retained blocks outside `sheetData`, as for
    /// [`Self::data_validations`]. Values are those of
    /// [`crate::Worksheet::conditional_formattings`].
    pub fn conditional_formattings(
        &self,
    ) -> Result<Vec<crate::conditional_formatting::Formatting>> {
        let value = self.owner.differential_format_count().and_then(|count| {
            let blocks = self.blocks()?;
            crate::conditional_formatting::parse_conditional_formattings(blocks.xml(), count)
        });
        self.finish_result(value)
    }

    /// Parse the worksheet's direct auto-filter declaration.
    ///
    /// Served from the retained blocks outside `sheetData`, as for
    /// [`Self::data_validations`]. Values are those of
    /// [`crate::Worksheet::auto_filter`].
    pub fn auto_filter(&self) -> Result<Option<crate::auto_filter::Definition>> {
        let value = self
            .blocks()
            .and_then(|blocks| crate::auto_filter::parse_auto_filter(blocks.xml()));
        self.finish_result(value)
    }

    /// Parse this worksheet's direct page breaks.
    ///
    /// Served from the retained blocks outside `sheetData`, as for
    /// [`Self::data_validations`]. Values are those of
    /// [`crate::Worksheet::page_breaks`].
    pub fn page_breaks(&self) -> Result<crate::page_breaks::PageBreaks> {
        let value = self
            .blocks()
            .and_then(|blocks| crate::page_breaks::parse(blocks.xml()));
        self.finish_result(value)
    }

    /// Parse this worksheet's legacy notes; `None` without a comments part.
    /// Only the worksheet's relationships and its comments part are read.
    pub fn comments(&self) -> Result<Option<crate::workbook::comments::Comments>> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name().to_owned(),
            });
        }
        let value = crate::workbook::comments::package::load_from_source(
            &self.owner.package,
            &self.data.part_uri,
        );
        self.finish_result(value)
    }

    /// Parse this worksheet's threaded comments (record 0322); `None`
    /// without a threaded-comments part. Only the worksheet's relationships
    /// and that part are read.
    pub fn threaded_comments(&self) -> Result<Option<crate::threaded_comments::Comments>> {
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name().to_owned(),
            });
        }
        let value = crate::threaded_comments::package::source_comments(
            &self.owner.package,
            &self.data.part_uri,
        );
        self.finish_result(value)
    }

    /// Parse this worksheet's typed inert hyperlinks.
    ///
    /// Served from the retained blocks outside `sheetData`, with external
    /// targets resolved through the worksheet part's relationships. Values
    /// are those of [`crate::Worksheet::hyperlinks`].
    pub fn hyperlinks(&self) -> Result<Vec<crate::hyperlinks::Hyperlink>> {
        let value = self.blocks().and_then(|blocks| {
            let part = self.owner.package.part(&self.data.part_uri)?;
            crate::hyperlinks::parse(blocks.xml(), part.rels())
        });
        self.finish_result(value)
    }

    /// Parse the worksheet's ordinary sheet-view collection.
    ///
    /// Served from the retained blocks outside `sheetData`, as for
    /// [`Self::data_validations`]. Values are those of
    /// [`crate::Worksheet::views`].
    pub fn views(&self) -> Result<Option<crate::sheet_view::Collection>> {
        let value = self
            .blocks()
            .and_then(|blocks| crate::sheet_view::parse_worksheet_views(blocks.xml()));
        self.finish_result(value)
    }

    /// The retained projection of this worksheet's blocks outside
    /// `sheetData`, built on first use.
    fn blocks(&self) -> Result<Arc<Blocks>> {
        self.owner.execution_check()?;
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name().to_owned(),
            });
        }
        // Resolving the part checks the captured source version even when the
        // projection is already retained.
        let part = self.owner.package.part(&self.data.part_uri)?;
        if let Some(blocks) = self.cached_blocks(true) {
            self.owner.package.check_execution()?;
            self.owner.execution_check()?;
            self.owner.package.source_version()?;
            return Ok(blocks);
        }
        let _build = self.data.lock_build();
        if let Some(blocks) = self.cached_blocks(false) {
            self.owner.package.check_execution()?;
            self.owner.execution_check()?;
            self.owner.package.source_version()?;
            return Ok(blocks);
        }
        let data = part.data()?;
        let blocks = Blocks::project(data.as_bytes(), self.owner.package.read_limits())?;
        drop(data);
        self.owner.execution_check()?;
        self.owner.package.source_version()?;
        let weight = blocks.weight_bytes();
        match self.owner.values().publish(
            self.data.position,
            Cached::Blocks(Arc::new(blocks)),
            weight,
        ) {
            Cached::Blocks(blocks) => Ok(blocks),
            Cached::Grid(_) | Cached::Prefix(_) => Err(invalid(
                "source-backed worksheet block cache returned another value kind",
            )),
        }
    }

    fn cached_blocks(&self, record_miss: bool) -> Option<Arc<Blocks>> {
        match self
            .owner
            .values()
            .get(self.data.position, Axis::Blocks, record_miss)?
        {
            Cached::Blocks(blocks) => Some(blocks),
            Cached::Grid(_) | Cached::Prefix(_) => None,
        }
    }

    /// Whether this worksheet's retained block projection removed the
    /// content of `sheetData`.
    #[cfg(test)]
    pub(super) fn blocks_elided_sheet_data(&self) -> Result<bool> {
        self.blocks().map(|blocks| blocks.elided)
    }
}

#[cfg(test)]
mod tests {
    use super::{Blocks, sheet_data_content};
    use crate::Error;
    use litchi_opc::{OpcError, SourceLimitKind, SourceLimits};

    const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    #[test]
    fn projection_removes_only_direct_sheet_data_content() {
        let xml = format!(
            r#"<?xml version="1.0"?><worksheet xmlns="{SML}"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><autoFilter ref="A1:B2"/></worksheet>"#
        );
        let blocks = Blocks::project(xml.as_bytes(), SourceLimits::default()).unwrap();
        assert!(blocks.elided);
        assert_eq!(
            std::str::from_utf8(blocks.xml()).unwrap(),
            format!(
                r#"<?xml version="1.0"?><worksheet xmlns="{SML}"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData></sheetData><autoFilter ref="A1:B2"/></worksheet>"#
            )
        );
    }

    #[test]
    fn projection_keeps_prefixed_strict_and_nested_names_exact() {
        let prefixed = r#"<x:worksheet xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main"><x:sheetData xmlns:y="urn:y"><x:row r="1"><x:sheetData/></x:row></x:sheetData><x:hyperlinks/></x:worksheet>"#;
        let blocks = Blocks::project(prefixed.as_bytes(), SourceLimits::default()).unwrap();
        assert_eq!(
            std::str::from_utf8(blocks.xml()).unwrap(),
            r#"<x:worksheet xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main"><x:sheetData xmlns:y="urn:y"></x:sheetData><x:hyperlinks/></x:worksheet>"#
        );

        // A foreign `sheetData`, a nested one, and an empty one are not
        // direct SpreadsheetML content to remove.
        for xml in [
            r#"<worksheet xmlns="urn:other"><sheetData><row/></sheetData></worksheet>"#.to_owned(),
            format!(
                r#"<worksheet xmlns="{SML}"><extLst><sheetData><row/></sheetData></extLst></worksheet>"#
            ),
            format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#),
        ] {
            assert_eq!(
                sheet_data_content(xml.as_bytes(), SourceLimits::default()).unwrap(),
                None,
                "{xml}"
            );
            let blocks = Blocks::project(xml.as_bytes(), SourceLimits::default()).unwrap();
            assert!(!blocks.elided);
            assert_eq!(blocks.xml(), xml.as_bytes());
        }
    }

    #[test]
    fn projection_does_not_hide_excessive_cell_markup_depth() {
        let xml = format!(
            r#"<worksheet xmlns="{SML}"><sheetData>{}<v>1</v>{}</sheetData></worksheet>"#,
            "<row>".repeat(128),
            "</row>".repeat(128)
        );
        let blocks = Blocks::project(xml.as_bytes(), SourceLimits::default()).unwrap();
        assert!(!blocks.elided);
        assert_eq!(blocks.xml(), xml.as_bytes());
        assert!(crate::data_validation::parse_data_validation_collections(blocks.xml()).is_err());
    }

    #[test]
    fn source_budgets_apply_before_cell_elision() {
        let xml = format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
        );
        for (limits, kind) in [
            (
                SourceLimits::builder()
                    .max_events(3)
                    .unwrap()
                    .build()
                    .unwrap(),
                SourceLimitKind::Events,
            ),
            (
                SourceLimits::builder()
                    .max_depth(3)
                    .unwrap()
                    .build()
                    .unwrap(),
                SourceLimitKind::Depth,
            ),
            (
                SourceLimits::builder()
                    .max_xml_attribute_bytes(2)
                    .unwrap()
                    .max_relationship_target_bytes(2)
                    .unwrap()
                    .build()
                    .unwrap(),
                SourceLimitKind::XmlAttributeBytes,
            ),
        ] {
            assert!(
                matches!(Blocks::project(xml.as_bytes(), limits), Err(Error::Package(OpcError::Limit { kind: actual, .. })) if actual == kind)
            );
        }
    }

    #[test]
    fn projection_preserves_foreign_roots_and_multiple_documents() {
        for xml in [
            format!(r#"<other xmlns="{SML}"><sheetData><row/></sheetData></other>"#),
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData><row/></sheetData></worksheet><worksheet xmlns="{SML}"/>"#
            ),
        ] {
            let blocks = Blocks::project(xml.as_bytes(), SourceLimits::default()).unwrap();
            assert!(!blocks.elided);
            assert_eq!(blocks.xml(), xml.as_bytes());
        }
    }

    #[test]
    fn unreadable_sheet_data_is_retained_whole() {
        let xml = format!(
            r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"></row></sheetData></worksheet>"#
        );
        assert_eq!(
            sheet_data_content(xml.as_bytes(), SourceLimits::default()).unwrap(),
            None
        );
        let blocks = Blocks::project(xml.as_bytes(), SourceLimits::default()).unwrap();
        assert!(!blocks.elided);
        assert_eq!(blocks.xml(), xml.as_bytes());
    }
}
