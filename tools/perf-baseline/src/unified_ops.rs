//! Paired unified-API and leaf-API routes for the M0a re-measurement
//! (plan 0003 §5.3, step S8b).
//!
//! M0a moves every open and read behind `litchi::open*` and the flat
//! `litchi::File`. The goal's performance bar judges that layer against the
//! leaf APIs it wraps, scenario by scenario, so every §5.3 scenario gets two
//! new opt-in selectors:
//!
//! - a **unified** leg (`<prefix>_unified_<rest>`) that opens through
//!   `litchi::open`, `litchi::open_bytes` or `litchi::open_bytes_with` and
//!   reads through `litchi::File`, its family views (`File::document`,
//!   `File::presentation`, `File::workbook`) and its format handles; and
//! - a **leaf** leg (`<prefix>_leaf_<rest>`) that makes the same leaf calls
//!   directly on the format's leaf crate. The format is known, so the leaf leg
//!   runs no format detection, no signature or encryption classification and
//!   no source-freshness re-check, and builds no handle. Everything else the
//!   unified route does through the leaf (an eager validation pass, a
//!   core-properties read at open) is done here too, so a pair's difference
//!   is the unified layer's own work.
//!
//! Each route mirrors the predecessor selector it re-measures: corpus,
//! ingress, untimed preparation and timer placement. M0a step S9 deleted the
//! 47 predecessors that timed the deleted facade types; the 15 that time the
//! detection functions or the DOCX leaf owner stay registered. Frozen records
//! are never rewritten, and S10 compares a predecessor's pre-M0a record with
//! its unified successor's post-S9 record.
//!
//! The DOCX, PPTX and XLSX filesystem pairs run in fresh child processes
//! (`crate::filesystem`); the other pairs run here, in process. Detection has
//! a leaf leg only where one leaf detector decides the input (ODF bytes and
//! readers); the polyglot and catalog-alias inputs are decided by the
//! coordinator's precedence between two leaf owners, so their unified leg is
//! paired with the facade's public detector instead.
//!
//! No route here makes a performance claim. Every sample is checked outside
//! the timer against an oracle built from the leaf crates and the corpus
//! generators, never from the code under measurement.

use std::{
    error::Error,
    fs,
    io::{Cursor, Write},
    path::{Path, PathBuf},
    sync::Arc,
    time::{Duration, Instant},
};

use litchi::{OpenOptions as UnifiedOpenOptions, SourceMode};
use litchi_core::{Metadata, OwnedSource, Position, ReadAt};
use litchi_ole_common::property_set::PropertySetReader;
use litchi_opc::{OpcPackage, SourceBackedPackage, SourceLimits};
use litchi_ops::detect::SourceLimits as DetectSourceLimits;
use litchi_ops::transaction::SemanticLimits;
use serde::Serialize;

use crate::{
    Case, CaseResult, Corpus, CorpusManifest, DetectionMode, DetectionOdfFamily,
    RtfSemanticVariant, SemanticShape, SourceSummary, iteration_count, record_elapsed,
    result_with_source, sha256_hex,
};

/// Which side of a pair a route measures.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Leg {
    /// `litchi::open*`, `litchi::File`, its views and handles.
    Unified,
    /// The format's leaf crate, called directly.
    Leaf,
}

impl Leg {
    const fn name(self) -> &'static str {
        match self {
            Self::Unified => "unified",
            Self::Leaf => "leaf",
        }
    }
}

/// The ingress and owner state an in-process route mirrors: `Eager` is the
/// in-memory byte route, `Source` the path route.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Ingress {
    Eager,
    Source,
}

/// The six legacy OLE2 scenarios of the predecessor facade routes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Ole2Query {
    DocOpen,
    DocFullText,
    DocOneParagraph,
    PptOpen,
    PptFullText,
    PptOneSlideText,
}

impl Ole2Query {
    const fn is_doc(self) -> bool {
        matches!(
            self,
            Self::DocOpen | Self::DocFullText | Self::DocOneParagraph
        )
    }
}

/// What an in-process route runs. Filesystem routes are planned by
/// `crate::filesystem`.
#[derive(Clone, Copy, Debug)]
pub(crate) enum Plan {
    Filesystem,
    XlsxBytes {
        lifecycle: bool,
    },
    DocxSemantic {
        /// Time only `Paragraph::text` on a prepared paragraph.
        text: bool,
    },
    DocxProducer,
    Ole2(Ole2Query),
    Rtf {
        lifecycle: bool,
    },
    Odt {
        ingress: Ingress,
        lifecycle: bool,
    },
    Odp {
        ingress: Ingress,
        selected_slide: bool,
    },
    Ods {
        ingress: Ingress,
    },
    Detect {
        family: DetectionOdfFamily,
        mode: DetectionMode,
    },
}

/// One registered route of a pair.
#[derive(Clone, Copy, Debug)]
pub(crate) struct Entry {
    pub(crate) case: Case,
    pub(crate) leg: Leg,
    /// The name of the pre-M0a selector whose scenario this route
    /// re-measures. M0a step S9 deleted the predecessors that timed the
    /// facade types; their frozen records keep the name.
    pub(crate) predecessor: &'static str,
    /// The other leg of the pair. `None` only for the detection inputs that no
    /// single leaf detector decides.
    pub(crate) twin: Option<Case>,
    pub(crate) plan: Plan,
}

/// Every paired route, in `Case::name` order.
pub(crate) const CATALOG: [Entry; 118] = [
    Entry {
        case: Case::DocxFileUnifiedEagerOpen,
        leg: Leg::Unified,
        predecessor: "docx_file_eager_open",
        twin: Some(Case::DocxFileLeafEagerOpen),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileUnifiedSourceOpen,
        leg: Leg::Unified,
        predecessor: "docx_file_source_open",
        twin: Some(Case::DocxFileLeafSourceOpen),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileUnifiedEagerParagraphCount,
        leg: Leg::Unified,
        predecessor: "docx_file_eager_paragraph_count",
        twin: Some(Case::DocxFileLeafEagerParagraphCount),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileUnifiedSourceParagraphCount,
        leg: Leg::Unified,
        predecessor: "docx_file_source_paragraph_count",
        twin: Some(Case::DocxFileLeafSourceParagraphCount),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileUnifiedEagerListParagraphs,
        leg: Leg::Unified,
        predecessor: "docx_file_eager_list_paragraphs",
        twin: Some(Case::DocxFileLeafEagerListParagraphs),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileUnifiedSourceListParagraphs,
        leg: Leg::Unified,
        predecessor: "docx_file_source_list_paragraphs",
        twin: Some(Case::DocxFileLeafSourceListParagraphs),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileUnifiedEagerFullText,
        leg: Leg::Unified,
        predecessor: "docx_file_eager_full_text",
        twin: Some(Case::DocxFileLeafEagerFullText),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileUnifiedSourceFullText,
        leg: Leg::Unified,
        predecessor: "docx_file_source_full_text",
        twin: Some(Case::DocxFileLeafSourceFullText),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileUnifiedEagerOpenParagraphCountLifecycle,
        leg: Leg::Unified,
        predecessor: "docx_file_eager_open_paragraph_count_lifecycle",
        twin: Some(Case::DocxFileLeafEagerOpenParagraphCountLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileUnifiedSourceOpenParagraphCountLifecycle,
        leg: Leg::Unified,
        predecessor: "docx_file_source_open_paragraph_count_lifecycle",
        twin: Some(Case::DocxFileLeafSourceOpenParagraphCountLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileUnifiedEagerOpenFullTextLifecycle,
        leg: Leg::Unified,
        predecessor: "docx_file_eager_open_full_text_lifecycle",
        twin: Some(Case::DocxFileLeafEagerOpenFullTextLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileUnifiedSourceOpenFullTextLifecycle,
        leg: Leg::Unified,
        predecessor: "docx_file_source_open_full_text_lifecycle",
        twin: Some(Case::DocxFileLeafSourceOpenFullTextLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafEagerOpen,
        leg: Leg::Leaf,
        predecessor: "docx_file_eager_open",
        twin: Some(Case::DocxFileUnifiedEagerOpen),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafSourceOpen,
        leg: Leg::Leaf,
        predecessor: "docx_file_source_open",
        twin: Some(Case::DocxFileUnifiedSourceOpen),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafEagerParagraphCount,
        leg: Leg::Leaf,
        predecessor: "docx_file_eager_paragraph_count",
        twin: Some(Case::DocxFileUnifiedEagerParagraphCount),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafSourceParagraphCount,
        leg: Leg::Leaf,
        predecessor: "docx_file_source_paragraph_count",
        twin: Some(Case::DocxFileUnifiedSourceParagraphCount),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafEagerListParagraphs,
        leg: Leg::Leaf,
        predecessor: "docx_file_eager_list_paragraphs",
        twin: Some(Case::DocxFileUnifiedEagerListParagraphs),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafSourceListParagraphs,
        leg: Leg::Leaf,
        predecessor: "docx_file_source_list_paragraphs",
        twin: Some(Case::DocxFileUnifiedSourceListParagraphs),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafEagerFullText,
        leg: Leg::Leaf,
        predecessor: "docx_file_eager_full_text",
        twin: Some(Case::DocxFileUnifiedEagerFullText),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafSourceFullText,
        leg: Leg::Leaf,
        predecessor: "docx_file_source_full_text",
        twin: Some(Case::DocxFileUnifiedSourceFullText),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafEagerOpenParagraphCountLifecycle,
        leg: Leg::Leaf,
        predecessor: "docx_file_eager_open_paragraph_count_lifecycle",
        twin: Some(Case::DocxFileUnifiedEagerOpenParagraphCountLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafSourceOpenParagraphCountLifecycle,
        leg: Leg::Leaf,
        predecessor: "docx_file_source_open_paragraph_count_lifecycle",
        twin: Some(Case::DocxFileUnifiedSourceOpenParagraphCountLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafEagerOpenFullTextLifecycle,
        leg: Leg::Leaf,
        predecessor: "docx_file_eager_open_full_text_lifecycle",
        twin: Some(Case::DocxFileUnifiedEagerOpenFullTextLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::DocxFileLeafSourceOpenFullTextLifecycle,
        leg: Leg::Leaf,
        predecessor: "docx_file_source_open_full_text_lifecycle",
        twin: Some(Case::DocxFileUnifiedSourceOpenFullTextLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedEagerOpen,
        leg: Leg::Unified,
        predecessor: "pptx_file_eager_open",
        twin: Some(Case::PptxFileLeafEagerOpen),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedSourceOpen,
        leg: Leg::Unified,
        predecessor: "pptx_file_source_open",
        twin: Some(Case::PptxFileLeafSourceOpen),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedEagerListSlides,
        leg: Leg::Unified,
        predecessor: "pptx_file_eager_list_slides",
        twin: Some(Case::PptxFileLeafEagerListSlides),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedSourceListSlides,
        leg: Leg::Unified,
        predecessor: "pptx_file_source_list_slides",
        twin: Some(Case::PptxFileLeafSourceListSlides),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedEagerSlideCount,
        leg: Leg::Unified,
        predecessor: "pptx_file_eager_slide_count",
        twin: Some(Case::PptxFileLeafEagerSlideCount),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedSourceSlideCount,
        leg: Leg::Unified,
        predecessor: "pptx_file_source_slide_count",
        twin: Some(Case::PptxFileLeafSourceSlideCount),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedEagerSelectedSlide,
        leg: Leg::Unified,
        predecessor: "pptx_file_eager_selected_slide",
        twin: Some(Case::PptxFileLeafEagerSelectedSlide),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedSourceSelectedSlide,
        leg: Leg::Unified,
        predecessor: "pptx_file_source_selected_slide",
        twin: Some(Case::PptxFileLeafSourceSelectedSlide),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedEagerOpenSlideCountLifecycle,
        leg: Leg::Unified,
        predecessor: "pptx_file_eager_open_slide_count_lifecycle",
        twin: Some(Case::PptxFileLeafEagerOpenSlideCountLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedSourceOpenSlideCountLifecycle,
        leg: Leg::Unified,
        predecessor: "pptx_file_source_open_slide_count_lifecycle",
        twin: Some(Case::PptxFileLeafSourceOpenSlideCountLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedEagerOpenSelectedSlideLifecycle,
        leg: Leg::Unified,
        predecessor: "pptx_file_eager_open_selected_slide_lifecycle",
        twin: Some(Case::PptxFileLeafEagerOpenSelectedSlideLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileUnifiedSourceOpenSelectedSlideLifecycle,
        leg: Leg::Unified,
        predecessor: "pptx_file_source_open_selected_slide_lifecycle",
        twin: Some(Case::PptxFileLeafSourceOpenSelectedSlideLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafEagerOpen,
        leg: Leg::Leaf,
        predecessor: "pptx_file_eager_open",
        twin: Some(Case::PptxFileUnifiedEagerOpen),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafSourceOpen,
        leg: Leg::Leaf,
        predecessor: "pptx_file_source_open",
        twin: Some(Case::PptxFileUnifiedSourceOpen),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafEagerListSlides,
        leg: Leg::Leaf,
        predecessor: "pptx_file_eager_list_slides",
        twin: Some(Case::PptxFileUnifiedEagerListSlides),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafSourceListSlides,
        leg: Leg::Leaf,
        predecessor: "pptx_file_source_list_slides",
        twin: Some(Case::PptxFileUnifiedSourceListSlides),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafEagerSlideCount,
        leg: Leg::Leaf,
        predecessor: "pptx_file_eager_slide_count",
        twin: Some(Case::PptxFileUnifiedEagerSlideCount),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafSourceSlideCount,
        leg: Leg::Leaf,
        predecessor: "pptx_file_source_slide_count",
        twin: Some(Case::PptxFileUnifiedSourceSlideCount),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafEagerSelectedSlide,
        leg: Leg::Leaf,
        predecessor: "pptx_file_eager_selected_slide",
        twin: Some(Case::PptxFileUnifiedEagerSelectedSlide),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafSourceSelectedSlide,
        leg: Leg::Leaf,
        predecessor: "pptx_file_source_selected_slide",
        twin: Some(Case::PptxFileUnifiedSourceSelectedSlide),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafEagerOpenSlideCountLifecycle,
        leg: Leg::Leaf,
        predecessor: "pptx_file_eager_open_slide_count_lifecycle",
        twin: Some(Case::PptxFileUnifiedEagerOpenSlideCountLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafSourceOpenSlideCountLifecycle,
        leg: Leg::Leaf,
        predecessor: "pptx_file_source_open_slide_count_lifecycle",
        twin: Some(Case::PptxFileUnifiedSourceOpenSlideCountLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafEagerOpenSelectedSlideLifecycle,
        leg: Leg::Leaf,
        predecessor: "pptx_file_eager_open_selected_slide_lifecycle",
        twin: Some(Case::PptxFileUnifiedEagerOpenSelectedSlideLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::PptxFileLeafSourceOpenSelectedSlideLifecycle,
        leg: Leg::Leaf,
        predecessor: "pptx_file_source_open_selected_slide_lifecycle",
        twin: Some(Case::PptxFileUnifiedSourceOpenSelectedSlideLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::XlsxFileUnifiedOpen,
        leg: Leg::Unified,
        predecessor: "xlsx_file_open",
        twin: Some(Case::XlsxFileLeafOpen),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::XlsxFileUnifiedOpenLifecycle,
        leg: Leg::Unified,
        predecessor: "xlsx_file_open_lifecycle",
        twin: Some(Case::XlsxFileLeafOpenLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::XlsxFileUnifiedSelectedCell,
        leg: Leg::Unified,
        predecessor: "xlsx_file_selected_cell",
        twin: Some(Case::XlsxFileLeafSelectedCell),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::XlsxFileLeafOpen,
        leg: Leg::Leaf,
        predecessor: "xlsx_file_open",
        twin: Some(Case::XlsxFileUnifiedOpen),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::XlsxFileLeafOpenLifecycle,
        leg: Leg::Leaf,
        predecessor: "xlsx_file_open_lifecycle",
        twin: Some(Case::XlsxFileUnifiedOpenLifecycle),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::XlsxFileLeafSelectedCell,
        leg: Leg::Leaf,
        predecessor: "xlsx_file_selected_cell",
        twin: Some(Case::XlsxFileUnifiedSelectedCell),
        plan: Plan::Filesystem,
    },
    Entry {
        case: Case::XlsxBytesUnifiedOpen,
        leg: Leg::Unified,
        predecessor: "xlsx_bytes_open",
        twin: Some(Case::XlsxBytesLeafOpen),
        plan: Plan::XlsxBytes { lifecycle: false },
    },
    Entry {
        case: Case::XlsxBytesUnifiedOpenLifecycle,
        leg: Leg::Unified,
        predecessor: "xlsx_bytes_open_lifecycle",
        twin: Some(Case::XlsxBytesLeafOpenLifecycle),
        plan: Plan::XlsxBytes { lifecycle: true },
    },
    Entry {
        case: Case::XlsxBytesLeafOpen,
        leg: Leg::Leaf,
        predecessor: "xlsx_bytes_open",
        twin: Some(Case::XlsxBytesUnifiedOpen),
        plan: Plan::XlsxBytes { lifecycle: false },
    },
    Entry {
        case: Case::XlsxBytesLeafOpenLifecycle,
        leg: Leg::Leaf,
        predecessor: "xlsx_bytes_open_lifecycle",
        twin: Some(Case::XlsxBytesUnifiedOpenLifecycle),
        plan: Plan::XlsxBytes { lifecycle: true },
    },
    Entry {
        case: Case::DocxSemanticUnifiedOneParagraph,
        leg: Leg::Unified,
        predecessor: "docx_semantic_one_paragraph",
        twin: Some(Case::DocxSemanticLeafOneParagraph),
        plan: Plan::DocxSemantic { text: false },
    },
    Entry {
        case: Case::DocxSemanticUnifiedOneParagraphText,
        leg: Leg::Unified,
        predecessor: "docx_semantic_one_paragraph_text",
        twin: Some(Case::DocxSemanticLeafOneParagraphText),
        plan: Plan::DocxSemantic { text: true },
    },
    Entry {
        case: Case::DocxSemanticLeafOneParagraph,
        leg: Leg::Leaf,
        predecessor: "docx_semantic_one_paragraph",
        twin: Some(Case::DocxSemanticUnifiedOneParagraph),
        plan: Plan::DocxSemantic { text: false },
    },
    Entry {
        case: Case::DocxSemanticLeafOneParagraphText,
        leg: Leg::Leaf,
        predecessor: "docx_semantic_one_paragraph_text",
        twin: Some(Case::DocxSemanticUnifiedOneParagraphText),
        plan: Plan::DocxSemantic { text: true },
    },
    Entry {
        case: Case::DocxProducerUnifiedSourceSelectedParagraph,
        leg: Leg::Unified,
        predecessor: "docx_producer_source_selected_paragraph",
        twin: Some(Case::DocxProducerLeafSourceSelectedParagraph),
        plan: Plan::DocxProducer,
    },
    Entry {
        case: Case::DocxProducerLeafSourceSelectedParagraph,
        leg: Leg::Leaf,
        predecessor: "docx_producer_source_selected_paragraph",
        twin: Some(Case::DocxProducerUnifiedSourceSelectedParagraph),
        plan: Plan::DocxProducer,
    },
    Entry {
        case: Case::DocFileUnifiedOpen,
        leg: Leg::Unified,
        predecessor: "doc_facade_file_open",
        twin: Some(Case::DocFileLeafOpen),
        plan: Plan::Ole2(Ole2Query::DocOpen),
    },
    Entry {
        case: Case::DocFileUnifiedFullText,
        leg: Leg::Unified,
        predecessor: "doc_facade_file_full_text",
        twin: Some(Case::DocFileLeafFullText),
        plan: Plan::Ole2(Ole2Query::DocFullText),
    },
    Entry {
        case: Case::DocFileUnifiedOneParagraph,
        leg: Leg::Unified,
        predecessor: "doc_facade_file_one_paragraph",
        twin: Some(Case::DocFileLeafOneParagraph),
        plan: Plan::Ole2(Ole2Query::DocOneParagraph),
    },
    Entry {
        case: Case::DocFileLeafOpen,
        leg: Leg::Leaf,
        predecessor: "doc_facade_file_open",
        twin: Some(Case::DocFileUnifiedOpen),
        plan: Plan::Ole2(Ole2Query::DocOpen),
    },
    Entry {
        case: Case::DocFileLeafFullText,
        leg: Leg::Leaf,
        predecessor: "doc_facade_file_full_text",
        twin: Some(Case::DocFileUnifiedFullText),
        plan: Plan::Ole2(Ole2Query::DocFullText),
    },
    Entry {
        case: Case::DocFileLeafOneParagraph,
        leg: Leg::Leaf,
        predecessor: "doc_facade_file_one_paragraph",
        twin: Some(Case::DocFileUnifiedOneParagraph),
        plan: Plan::Ole2(Ole2Query::DocOneParagraph),
    },
    Entry {
        case: Case::PptFileUnifiedOpen,
        leg: Leg::Unified,
        predecessor: "ppt_facade_file_open",
        twin: Some(Case::PptFileLeafOpen),
        plan: Plan::Ole2(Ole2Query::PptOpen),
    },
    Entry {
        case: Case::PptFileUnifiedFullText,
        leg: Leg::Unified,
        predecessor: "ppt_facade_file_full_text",
        twin: Some(Case::PptFileLeafFullText),
        plan: Plan::Ole2(Ole2Query::PptFullText),
    },
    Entry {
        case: Case::PptFileUnifiedOneSlideText,
        leg: Leg::Unified,
        predecessor: "ppt_facade_file_one_slide_text",
        twin: Some(Case::PptFileLeafOneSlideText),
        plan: Plan::Ole2(Ole2Query::PptOneSlideText),
    },
    Entry {
        case: Case::PptFileLeafOpen,
        leg: Leg::Leaf,
        predecessor: "ppt_facade_file_open",
        twin: Some(Case::PptFileUnifiedOpen),
        plan: Plan::Ole2(Ole2Query::PptOpen),
    },
    Entry {
        case: Case::PptFileLeafFullText,
        leg: Leg::Leaf,
        predecessor: "ppt_facade_file_full_text",
        twin: Some(Case::PptFileUnifiedFullText),
        plan: Plan::Ole2(Ole2Query::PptFullText),
    },
    Entry {
        case: Case::PptFileLeafOneSlideText,
        leg: Leg::Leaf,
        predecessor: "ppt_facade_file_one_slide_text",
        twin: Some(Case::PptFileUnifiedOneSlideText),
        plan: Plan::Ole2(Ole2Query::PptOneSlideText),
    },
    Entry {
        case: Case::RtfFileUnifiedOpen,
        leg: Leg::Unified,
        predecessor: "rtf_file_open",
        twin: Some(Case::RtfFileLeafOpen),
        plan: Plan::Rtf { lifecycle: false },
    },
    Entry {
        case: Case::RtfFileUnifiedOpenLifecycle,
        leg: Leg::Unified,
        predecessor: "rtf_file_open_lifecycle",
        twin: Some(Case::RtfFileLeafOpenLifecycle),
        plan: Plan::Rtf { lifecycle: true },
    },
    Entry {
        case: Case::RtfFileLeafOpen,
        leg: Leg::Leaf,
        predecessor: "rtf_file_open",
        twin: Some(Case::RtfFileUnifiedOpen),
        plan: Plan::Rtf { lifecycle: false },
    },
    Entry {
        case: Case::RtfFileLeafOpenLifecycle,
        leg: Leg::Leaf,
        predecessor: "rtf_file_open_lifecycle",
        twin: Some(Case::RtfFileUnifiedOpenLifecycle),
        plan: Plan::Rtf { lifecycle: true },
    },
    Entry {
        case: Case::OdtFileUnifiedEagerOpen,
        leg: Leg::Unified,
        predecessor: "odt_file_eager_open",
        twin: Some(Case::OdtFileLeafEagerOpen),
        plan: Plan::Odt {
            ingress: Ingress::Eager,
            lifecycle: false,
        },
    },
    Entry {
        case: Case::OdtFileUnifiedSourceOpen,
        leg: Leg::Unified,
        predecessor: "odt_file_source_open",
        twin: Some(Case::OdtFileLeafSourceOpen),
        plan: Plan::Odt {
            ingress: Ingress::Source,
            lifecycle: false,
        },
    },
    Entry {
        case: Case::OdtFileUnifiedEagerOpenFullTextLifecycle,
        leg: Leg::Unified,
        predecessor: "odt_file_eager_open_full_text_lifecycle",
        twin: Some(Case::OdtFileLeafEagerOpenFullTextLifecycle),
        plan: Plan::Odt {
            ingress: Ingress::Eager,
            lifecycle: true,
        },
    },
    Entry {
        case: Case::OdtFileUnifiedSourceOpenFullTextLifecycle,
        leg: Leg::Unified,
        predecessor: "odt_file_source_open_full_text_lifecycle",
        twin: Some(Case::OdtFileLeafSourceOpenFullTextLifecycle),
        plan: Plan::Odt {
            ingress: Ingress::Source,
            lifecycle: true,
        },
    },
    Entry {
        case: Case::OdtFileLeafEagerOpen,
        leg: Leg::Leaf,
        predecessor: "odt_file_eager_open",
        twin: Some(Case::OdtFileUnifiedEagerOpen),
        plan: Plan::Odt {
            ingress: Ingress::Eager,
            lifecycle: false,
        },
    },
    Entry {
        case: Case::OdtFileLeafSourceOpen,
        leg: Leg::Leaf,
        predecessor: "odt_file_source_open",
        twin: Some(Case::OdtFileUnifiedSourceOpen),
        plan: Plan::Odt {
            ingress: Ingress::Source,
            lifecycle: false,
        },
    },
    Entry {
        case: Case::OdtFileLeafEagerOpenFullTextLifecycle,
        leg: Leg::Leaf,
        predecessor: "odt_file_eager_open_full_text_lifecycle",
        twin: Some(Case::OdtFileUnifiedEagerOpenFullTextLifecycle),
        plan: Plan::Odt {
            ingress: Ingress::Eager,
            lifecycle: true,
        },
    },
    Entry {
        case: Case::OdtFileLeafSourceOpenFullTextLifecycle,
        leg: Leg::Leaf,
        predecessor: "odt_file_source_open_full_text_lifecycle",
        twin: Some(Case::OdtFileUnifiedSourceOpenFullTextLifecycle),
        plan: Plan::Odt {
            ingress: Ingress::Source,
            lifecycle: true,
        },
    },
    Entry {
        case: Case::OdpFileUnifiedEagerOpen,
        leg: Leg::Unified,
        predecessor: "odp_file_eager_open",
        twin: Some(Case::OdpFileLeafEagerOpen),
        plan: Plan::Odp {
            ingress: Ingress::Eager,
            selected_slide: false,
        },
    },
    Entry {
        case: Case::OdpFileUnifiedSourceOpen,
        leg: Leg::Unified,
        predecessor: "odp_file_source_open",
        twin: Some(Case::OdpFileLeafSourceOpen),
        plan: Plan::Odp {
            ingress: Ingress::Source,
            selected_slide: false,
        },
    },
    Entry {
        case: Case::OdpFileUnifiedEagerSelectedSlide,
        leg: Leg::Unified,
        predecessor: "odp_file_eager_selected_slide",
        twin: Some(Case::OdpFileLeafEagerSelectedSlide),
        plan: Plan::Odp {
            ingress: Ingress::Eager,
            selected_slide: true,
        },
    },
    Entry {
        case: Case::OdpFileUnifiedSourceSelectedSlide,
        leg: Leg::Unified,
        predecessor: "odp_file_source_selected_slide",
        twin: Some(Case::OdpFileLeafSourceSelectedSlide),
        plan: Plan::Odp {
            ingress: Ingress::Source,
            selected_slide: true,
        },
    },
    Entry {
        case: Case::OdpFileLeafEagerOpen,
        leg: Leg::Leaf,
        predecessor: "odp_file_eager_open",
        twin: Some(Case::OdpFileUnifiedEagerOpen),
        plan: Plan::Odp {
            ingress: Ingress::Eager,
            selected_slide: false,
        },
    },
    Entry {
        case: Case::OdpFileLeafSourceOpen,
        leg: Leg::Leaf,
        predecessor: "odp_file_source_open",
        twin: Some(Case::OdpFileUnifiedSourceOpen),
        plan: Plan::Odp {
            ingress: Ingress::Source,
            selected_slide: false,
        },
    },
    Entry {
        case: Case::OdpFileLeafEagerSelectedSlide,
        leg: Leg::Leaf,
        predecessor: "odp_file_eager_selected_slide",
        twin: Some(Case::OdpFileUnifiedEagerSelectedSlide),
        plan: Plan::Odp {
            ingress: Ingress::Eager,
            selected_slide: true,
        },
    },
    Entry {
        case: Case::OdpFileLeafSourceSelectedSlide,
        leg: Leg::Leaf,
        predecessor: "odp_file_source_selected_slide",
        twin: Some(Case::OdpFileUnifiedSourceSelectedSlide),
        plan: Plan::Odp {
            ingress: Ingress::Source,
            selected_slide: true,
        },
    },
    Entry {
        case: Case::OdsFileUnifiedEagerOpen,
        leg: Leg::Unified,
        predecessor: "ods_file_eager_open",
        twin: Some(Case::OdsFileLeafEagerOpen),
        plan: Plan::Ods {
            ingress: Ingress::Eager,
        },
    },
    Entry {
        case: Case::OdsFileUnifiedSourceOpen,
        leg: Leg::Unified,
        predecessor: "ods_file_source_open",
        twin: Some(Case::OdsFileLeafSourceOpen),
        plan: Plan::Ods {
            ingress: Ingress::Source,
        },
    },
    Entry {
        case: Case::OdsFileLeafEagerOpen,
        leg: Leg::Leaf,
        predecessor: "ods_file_eager_open",
        twin: Some(Case::OdsFileUnifiedEagerOpen),
        plan: Plan::Ods {
            ingress: Ingress::Eager,
        },
    },
    Entry {
        case: Case::OdsFileLeafSourceOpen,
        leg: Leg::Leaf,
        predecessor: "ods_file_source_open",
        twin: Some(Case::OdsFileUnifiedSourceOpen),
        plan: Plan::Ods {
            ingress: Ingress::Source,
        },
    },
    Entry {
        case: Case::DetectOdtUnifiedBytes,
        leg: Leg::Unified,
        predecessor: "detect_odt_bytes",
        twin: Some(Case::DetectOdtLeafBytes),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odt,
            mode: DetectionMode::Bytes,
        },
    },
    Entry {
        case: Case::DetectOdtUnifiedReader,
        leg: Leg::Unified,
        predecessor: "detect_odt_reader",
        twin: Some(Case::DetectOdtLeafReader),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odt,
            mode: DetectionMode::Reader,
        },
    },
    Entry {
        case: Case::DetectOdtUnifiedPolyglot,
        leg: Leg::Unified,
        predecessor: "detect_odt_polyglot",
        twin: None,
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odt,
            mode: DetectionMode::Polyglot,
        },
    },
    Entry {
        case: Case::DetectOdtUnifiedCatalogAlias,
        leg: Leg::Unified,
        predecessor: "detect_odt_catalog_alias",
        twin: None,
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odt,
            mode: DetectionMode::CatalogAlias,
        },
    },
    Entry {
        case: Case::DetectOdtLeafBytes,
        leg: Leg::Leaf,
        predecessor: "detect_odt_bytes",
        twin: Some(Case::DetectOdtUnifiedBytes),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odt,
            mode: DetectionMode::Bytes,
        },
    },
    Entry {
        case: Case::DetectOdtLeafReader,
        leg: Leg::Leaf,
        predecessor: "detect_odt_reader",
        twin: Some(Case::DetectOdtUnifiedReader),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odt,
            mode: DetectionMode::Reader,
        },
    },
    Entry {
        case: Case::DetectOdsUnifiedBytes,
        leg: Leg::Unified,
        predecessor: "detect_ods_bytes",
        twin: Some(Case::DetectOdsLeafBytes),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Ods,
            mode: DetectionMode::Bytes,
        },
    },
    Entry {
        case: Case::DetectOdsUnifiedReader,
        leg: Leg::Unified,
        predecessor: "detect_ods_reader",
        twin: Some(Case::DetectOdsLeafReader),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Ods,
            mode: DetectionMode::Reader,
        },
    },
    Entry {
        case: Case::DetectOdsUnifiedPolyglot,
        leg: Leg::Unified,
        predecessor: "detect_ods_polyglot",
        twin: None,
        plan: Plan::Detect {
            family: DetectionOdfFamily::Ods,
            mode: DetectionMode::Polyglot,
        },
    },
    Entry {
        case: Case::DetectOdsUnifiedCatalogAlias,
        leg: Leg::Unified,
        predecessor: "detect_ods_catalog_alias",
        twin: None,
        plan: Plan::Detect {
            family: DetectionOdfFamily::Ods,
            mode: DetectionMode::CatalogAlias,
        },
    },
    Entry {
        case: Case::DetectOdsLeafBytes,
        leg: Leg::Leaf,
        predecessor: "detect_ods_bytes",
        twin: Some(Case::DetectOdsUnifiedBytes),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Ods,
            mode: DetectionMode::Bytes,
        },
    },
    Entry {
        case: Case::DetectOdsLeafReader,
        leg: Leg::Leaf,
        predecessor: "detect_ods_reader",
        twin: Some(Case::DetectOdsUnifiedReader),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Ods,
            mode: DetectionMode::Reader,
        },
    },
    Entry {
        case: Case::DetectOdpUnifiedBytes,
        leg: Leg::Unified,
        predecessor: "detect_odp_bytes",
        twin: Some(Case::DetectOdpLeafBytes),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odp,
            mode: DetectionMode::Bytes,
        },
    },
    Entry {
        case: Case::DetectOdpUnifiedReader,
        leg: Leg::Unified,
        predecessor: "detect_odp_reader",
        twin: Some(Case::DetectOdpLeafReader),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odp,
            mode: DetectionMode::Reader,
        },
    },
    Entry {
        case: Case::DetectOdpUnifiedPolyglot,
        leg: Leg::Unified,
        predecessor: "detect_odp_polyglot",
        twin: None,
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odp,
            mode: DetectionMode::Polyglot,
        },
    },
    Entry {
        case: Case::DetectOdpUnifiedCatalogAlias,
        leg: Leg::Unified,
        predecessor: "detect_odp_catalog_alias",
        twin: None,
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odp,
            mode: DetectionMode::CatalogAlias,
        },
    },
    Entry {
        case: Case::DetectOdpLeafBytes,
        leg: Leg::Leaf,
        predecessor: "detect_odp_bytes",
        twin: Some(Case::DetectOdpUnifiedBytes),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odp,
            mode: DetectionMode::Bytes,
        },
    },
    Entry {
        case: Case::DetectOdpLeafReader,
        leg: Leg::Leaf,
        predecessor: "detect_odp_reader",
        twin: Some(Case::DetectOdpUnifiedReader),
        plan: Plan::Detect {
            family: DetectionOdfFamily::Odp,
            mode: DetectionMode::Reader,
        },
    },
];

/// The catalog entry of `case`, if it is a paired route.
pub(crate) fn entry(case: Case) -> Option<&'static Entry> {
    CATALOG.iter().find(|entry| entry.case == case)
}

/// The owner used by the opt-in legacy-PPT write pair.  `Eager` starts from
/// owned bytes; `Positional` opens the same prepared file path and retains its
/// source-version fence through publication.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PptTextBatchOwner {
    Eager,
    Positional,
}

impl PptTextBatchOwner {
    const fn name(self) -> &'static str {
        match self {
            Self::Eager => "eager",
            Self::Positional => "positional",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PptTextBatchLeg {
    Native,
    Unified,
}

impl PptTextBatchLeg {
    const fn name(self) -> &'static str {
        match self {
            Self::Native => "native",
            Self::Unified => "unified",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct PptTextBatchEntry {
    case: Case,
    twin: Case,
    owner: PptTextBatchOwner,
    leg: PptTextBatchLeg,
    count: usize,
}

/// Opt-in, named native/unified legacy-PPT write routes.  These are kept
/// outside the M0a read-pair catalog: their stable identities compare two
/// write implementations rather than a deleted facade predecessor.
const PPT_TEXT_BATCH_CATALOG: [PptTextBatchEntry; 28] = [
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerNative1,
        twin: Case::PptTextAggregateEagerUnified1,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Native,
        count: 1,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerNative10,
        twin: Case::PptTextAggregateEagerUnified10,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Native,
        count: 10,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerNative100,
        twin: Case::PptTextAggregateEagerUnified100,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Native,
        count: 100,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerNative256,
        twin: Case::PptTextAggregateEagerUnified256,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Native,
        count: 256,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerNative1000,
        twin: Case::PptTextAggregateEagerUnified1000,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Native,
        count: 1_000,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerNative4096,
        twin: Case::PptTextAggregateEagerUnified4096,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Native,
        count: 4_096,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerNative10000,
        twin: Case::PptTextAggregateEagerUnified10000,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Native,
        count: 10_000,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerUnified1,
        twin: Case::PptTextAggregateEagerNative1,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Unified,
        count: 1,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerUnified10,
        twin: Case::PptTextAggregateEagerNative10,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Unified,
        count: 10,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerUnified100,
        twin: Case::PptTextAggregateEagerNative100,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Unified,
        count: 100,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerUnified256,
        twin: Case::PptTextAggregateEagerNative256,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Unified,
        count: 256,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerUnified1000,
        twin: Case::PptTextAggregateEagerNative1000,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Unified,
        count: 1_000,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerUnified4096,
        twin: Case::PptTextAggregateEagerNative4096,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Unified,
        count: 4_096,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregateEagerUnified10000,
        twin: Case::PptTextAggregateEagerNative10000,
        owner: PptTextBatchOwner::Eager,
        leg: PptTextBatchLeg::Unified,
        count: 10_000,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalNative1,
        twin: Case::PptTextAggregatePositionalUnified1,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Native,
        count: 1,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalNative10,
        twin: Case::PptTextAggregatePositionalUnified10,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Native,
        count: 10,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalNative100,
        twin: Case::PptTextAggregatePositionalUnified100,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Native,
        count: 100,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalNative256,
        twin: Case::PptTextAggregatePositionalUnified256,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Native,
        count: 256,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalNative1000,
        twin: Case::PptTextAggregatePositionalUnified1000,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Native,
        count: 1_000,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalNative4096,
        twin: Case::PptTextAggregatePositionalUnified4096,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Native,
        count: 4_096,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalNative10000,
        twin: Case::PptTextAggregatePositionalUnified10000,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Native,
        count: 10_000,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalUnified1,
        twin: Case::PptTextAggregatePositionalNative1,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Unified,
        count: 1,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalUnified10,
        twin: Case::PptTextAggregatePositionalNative10,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Unified,
        count: 10,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalUnified100,
        twin: Case::PptTextAggregatePositionalNative100,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Unified,
        count: 100,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalUnified256,
        twin: Case::PptTextAggregatePositionalNative256,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Unified,
        count: 256,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalUnified1000,
        twin: Case::PptTextAggregatePositionalNative1000,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Unified,
        count: 1_000,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalUnified4096,
        twin: Case::PptTextAggregatePositionalNative4096,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Unified,
        count: 4_096,
    },
    PptTextBatchEntry {
        case: Case::PptTextAggregatePositionalUnified10000,
        twin: Case::PptTextAggregatePositionalNative10000,
        owner: PptTextBatchOwner::Positional,
        leg: PptTextBatchLeg::Unified,
        count: 10_000,
    },
];

fn ppt_text_batch_entry(case: Case) -> Option<PptTextBatchEntry> {
    PPT_TEXT_BATCH_CATALOG
        .iter()
        .copied()
        .find(|entry| entry.case == case)
}

const PERFORMANCE_CLAIM: &str = "none: one leg of a paired unified/leaf re-measurement (plan 0003 \
                                 §5.3); comparisons belong to the M0a S10 record";

/// What a paired route measures, restated in every result as
/// `source.unified_ops`.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct RouteSummary {
    /// `unified` or `leaf`.
    pub(crate) leg: &'static str,
    /// The pre-M0a selector whose scenario this route re-measures.
    pub(crate) predecessor: &'static str,
    /// The other leg of the pair.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) twin: Option<&'static str>,
    /// The public calls inside the timed scope.
    pub(crate) timed_calls: &'static str,
    pub(crate) performance_claim: &'static str,
    /// Digest of the oracle projection every retained sample reproduced.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) projection_sha256: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) verified_samples: Option<usize>,
    /// Legacy OLE2 routes: the frozen outcome of this leg, the selected
    /// target, and whether both legs reached the same outcome.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) observation: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) selected_target: Option<usize>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) legs_agree: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) real_file: Option<crate::producer_shape::RealFileProvenance>,
    /// Qualification and native rollover evidence for the opt-in legacy PPT
    /// text-batch routes.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) ppt_text_batch: Option<PptTextBatchSummary>,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct PptTextBatchSummary {
    pub(crate) leg: &'static str,
    pub(crate) owner: &'static str,
    pub(crate) count: usize,
    pub(crate) native_chunk_size: usize,
    pub(crate) native_operation_limit: usize,
    pub(crate) native_inner_chunks: usize,
    pub(crate) native_aggregate_calls: usize,
    pub(crate) native_publications: usize,
    pub(crate) native_commits: usize,
    pub(crate) unified_operations: usize,
    pub(crate) unified_commits: usize,
    pub(crate) semantic_snapshot_limit_bytes: usize,
    pub(crate) source_sha256: String,
    pub(crate) expected_output_sha256: String,
    pub(crate) exact_output_verified: bool,
    pub(crate) unchanged_source_verified: bool,
    pub(crate) timing_scope: &'static str,
}

impl RouteSummary {
    pub(crate) fn for_case(case: Case, timed_calls: &'static str) -> Result<Self, Box<dyn Error>> {
        let entry = entry(case).ok_or("case is not a paired unified/leaf route")?;
        Ok(Self {
            leg: entry.leg.name(),
            predecessor: entry.predecessor,
            twin: entry.twin.map(Case::name),
            timed_calls,
            performance_claim: PERFORMANCE_CLAIM,
            projection_sha256: None,
            verified_samples: None,
            observation: None,
            selected_target: None,
            legs_agree: None,
            real_file: None,
            ppt_text_batch: None,
        })
    }

    fn verified(
        mut self,
        projection: &impl Serialize,
        samples: usize,
    ) -> Result<Self, Box<dyn Error>> {
        self.projection_sha256 = Some(sha256_hex(&serde_json::to_vec(projection)?));
        self.verified_samples = Some(samples);
        Ok(self)
    }
}

fn finish(entry: &Entry, corpus: &Corpus, elapsed: Vec<u64>, summary: RouteSummary) -> CaseResult {
    result_with_source(
        entry.case,
        corpus,
        elapsed,
        SourceSummary {
            unified_ops: Some(Box::new(summary)),
            ..SourceSummary::default()
        },
    )
}

// ---------------------------------------------------------------------------
// Leaf owners shared with the filesystem pairs
// ---------------------------------------------------------------------------

/// The options of the unified eager route: the owning package from the eager
/// detection route.
pub(crate) fn eager_options() -> UnifiedOpenOptions {
    UnifiedOpenOptions::new().source_mode(SourceMode::Eager)
}

/// The leaf calls the unified eager DOCX owner makes: one eager OPC parse,
/// the owning package, and its one-pass text validation.
pub(crate) fn leaf_docx_eager(bytes: Vec<u8>) -> Result<litchi_docx::Package, Box<dyn Error>> {
    let opc = OpcPackage::from_bytes(&bytes, SourceLimits::default())?;
    drop(bytes);
    let package = litchi_docx::Package::from_opc_package(opc)?;
    let document = package.document()?;
    document.text()?;
    std::hint::black_box(document.tolerance_report());
    drop(document);
    Ok(package)
}

/// The leaf XLSX owner and the core properties the unified open reads with
/// it.
pub(crate) struct LeafXlsx {
    pub(crate) workbook: litchi_xlsx::SourceBackedWorkbook,
    pub(crate) metadata: Metadata,
}

/// The leaf calls the unified source-backed XLSX open makes: one OPC index,
/// the core-properties read, and the deferred workbook catalog.
pub(crate) fn leaf_xlsx_open(source: Arc<dyn ReadAt>) -> Result<LeafXlsx, Box<dyn Error>> {
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        litchi_opc::SourceReadPolicy::exact(),
        None,
    )?;
    let mut repairs = litchi_ooxml_common::leniency::ToleranceReport::default();
    let metadata =
        litchi_ooxml_common::properties::read_source_backed_with_report(&package, &mut repairs)?
            .unwrap_or_default();
    let workbook = litchi_xlsx::SourceBackedWorkbook::from_source_backed_package(package)?;
    Ok(LeafXlsx { workbook, metadata })
}

pub(crate) fn leaf_xlsx_names(workbook: &litchi_xlsx::SourceBackedWorkbook) -> Vec<String> {
    workbook.sheet_names().map(str::to_owned).collect()
}

/// Cell text for one leaf XLSX cell, as the unified workbook text renders it.
fn push_xlsx_cell_text(out: &mut String, cell: &litchi_xlsx::cell::Cell) {
    use litchi_xlsx::cell::{Cell, Value};

    fn number(out: &mut String, value: &litchi_xlsx::cell::Number) {
        match value.as_str().parse::<i64>() {
            Ok(integer) => out.push_str(&integer.to_string()),
            Err(_) => out.push_str(&value.as_f64().unwrap_or_default().to_string()),
        }
    }

    match cell {
        Cell::Empty => {},
        Cell::Value(Value::Bool(value)) => out.push_str(if *value { "TRUE" } else { "FALSE" }),
        Cell::Value(Value::Number(value)) => number(out, value),
        Cell::Value(Value::Text(value)) => out.push_str(value.as_str()),
        Cell::Value(Value::Date(value)) => out.push_str(value.as_str()),
        Cell::Value(Value::Error(value)) => out.push_str(value.as_str()),
        Cell::Value(_) => out.push_str("unknown XLSX value kind"),
        Cell::Formula(formula) => match formula.cached().map(|cache| cache.value()) {
            Some(Value::Text(value)) => out.push_str(value.as_str()),
            Some(Value::Date(value)) => out.push_str(value.as_str()),
            Some(Value::Number(value)) => number(out, value),
            Some(Value::Bool(value)) => out.push_str(if *value { "TRUE" } else { "FALSE" }),
            _ => {
                out.push('=');
                out.push_str(formula.text());
            },
        },
        Cell::Unknown(unknown) => {
            out.push_str("unknown XLSX cell kind: ");
            out.push_str(unknown.kind());
        },
        _ => out.push_str("unknown XLSX cell kind"),
    }
}

/// The workbook text the unified handle produces, written directly against
/// the leaf: every row of every grid sheet from column 0 to the last stored
/// column, cells separated by tabs, rows ended by newlines.
pub(crate) fn leaf_xlsx_text(
    workbook: &litchi_xlsx::SourceBackedWorkbook,
) -> Result<String, Box<dyn Error>> {
    let mut out = String::new();
    for index in 0..workbook.len() {
        let sheet = workbook
            .sheet(index)?
            .ok_or("leaf XLSX sheet position is outside the catalog")?;
        let extent = match sheet.stored_extent() {
            Ok(extent) => extent,
            Err(litchi_xlsx::Error::NotWorksheet { .. }) => None,
            Err(error) => return Err(error.into()),
        };
        let Some(extent) = extent else {
            continue;
        };
        let (end_row, end_column) = extent.end();
        let width = usize::try_from(end_column)?;
        let mut row_cells: Vec<Option<litchi_xlsx::cell::Cell>> = Vec::with_capacity(width);
        for row in 0..end_row {
            row_cells.clear();
            row_cells.resize(width, None);
            for entry in sheet.cells((row, 0, row + 1, end_column))? {
                let column = usize::try_from(entry.address.column().get())?;
                *row_cells
                    .get_mut(column)
                    .ok_or("leaf XLSX cell lies outside its row extent")? = Some(entry.cell);
            }
            for (column, cell) in row_cells.iter().enumerate() {
                if column > 0 {
                    out.push('\t');
                }
                if let Some(cell) = cell {
                    push_xlsx_cell_text(&mut out, cell);
                }
            }
            out.push('\n');
        }
    }
    Ok(out)
}

// ---------------------------------------------------------------------------
// Dispatch
// ---------------------------------------------------------------------------

/// The run options the in-process pairs read.
pub(crate) struct Selection<'a> {
    pub(crate) cases: &'a [Case],
    pub(crate) semantic_shapes: &'a [SemanticShape],
    pub(crate) rtf_variants: &'a [RtfSemanticVariant],
    pub(crate) ole2_files: &'a [PathBuf],
    pub(crate) warmup_iterations: usize,
    pub(crate) samples: usize,
}

/// Run every selected in-process pair route. Each family builds its corpus
/// once, from the same builder its predecessor uses.
pub(crate) fn run_selected(selection: &Selection<'_>) -> Result<Vec<CaseResult>, Box<dyn Error>> {
    let entries = selection
        .cases
        .iter()
        .filter_map(|case| entry(*case))
        .filter(|entry| !matches!(entry.plan, Plan::Filesystem))
        .copied()
        .collect::<Vec<_>>();
    let warmup = selection.warmup_iterations;
    let samples = selection.samples;
    let family = |predicate: fn(&Plan) -> bool| {
        entries
            .iter()
            .filter(move |entry| predicate(&entry.plan))
            .collect::<Vec<_>>()
    };
    let mut results = Vec::new();

    let ppt_text_batch = selection
        .cases
        .iter()
        .filter_map(|case| ppt_text_batch_entry(*case))
        .collect::<Vec<_>>();
    if !ppt_text_batch.is_empty() {
        results.extend(run_ppt_text_batch_selected(
            &ppt_text_batch,
            warmup,
            samples,
        )?);
    }

    let xlsx_bytes = family(|plan| matches!(plan, Plan::XlsxBytes { .. }));
    if !xlsx_bytes.is_empty() {
        let corpus = crate::build_xlsx_cell_crud_corpus(crate::XlsxCellCrudShape::Medium)?;
        for entry in xlsx_bytes {
            results.push(run_xlsx_bytes(entry, &corpus, warmup, samples)?);
        }
    }

    let docx_semantic = family(|plan| matches!(plan, Plan::DocxSemantic { .. }));
    if !docx_semantic.is_empty() {
        for shape in selection.semantic_shapes {
            let corpus = crate::build_semantic_docx_corpus(*shape)?;
            for entry in &docx_semantic {
                results.push(run_docx_semantic(entry, &corpus, warmup, samples)?);
            }
        }
    }

    let docx_producer = family(|plan| matches!(plan, Plan::DocxProducer));
    if !docx_producer.is_empty() {
        let producer = crate::producer_shape::build_docx_producer_corpus()?;
        for entry in docx_producer {
            results.push(run_docx_producer(entry, &producer, warmup, samples)?);
        }
    }

    let ole2 = family(|plan| matches!(plan, Plan::Ole2(_)));
    if !ole2.is_empty() {
        if selection.ole2_files.is_empty() {
            return Err(
                "the DOC/PPT unified and leaf selectors require --ole2-file PATH (repeatable)"
                    .into(),
            );
        }
        let inputs = crate::ole2_range_source::classify_inputs(selection.ole2_files)?;
        let mut doc = None;
        let mut ppt = None;
        for entry in ole2 {
            let Plan::Ole2(query) = entry.plan else {
                return Err("OLE2 family selected a non-OLE2 plan".into());
            };
            let corpus = if query.is_doc() {
                if doc.is_none() {
                    let path = inputs.doc.as_deref().ok_or(
                        "the DOC unified and leaf selectors require an --ole2-file DOC fixture",
                    )?;
                    doc = Some(build_ole2_corpus(path, true)?);
                }
                doc.as_ref().ok_or("DOC corpus is absent")?
            } else {
                if ppt.is_none() {
                    let path = inputs.ppt.as_deref().ok_or(
                        "the PPT unified and leaf selectors require an --ole2-file PPT fixture",
                    )?;
                    ppt = Some(build_ole2_corpus(path, false)?);
                }
                ppt.as_ref().ok_or("PPT corpus is absent")?
            };
            results.push(run_ole2(entry, query, corpus, warmup, samples)?);
        }
    }

    let rtf = family(|plan| matches!(plan, Plan::Rtf { .. }));
    if !rtf.is_empty() {
        let mut rows = 0_usize;
        for variant in selection
            .rtf_variants
            .iter()
            .filter(|variant| **variant != RtfSemanticVariant::Watermark)
        {
            for shape in selection
                .semantic_shapes
                .iter()
                .filter(|shape| variant.supports_shape(**shape))
            {
                let corpus = crate::build_semantic_rtf_corpus(*shape, *variant)?;
                for entry in &rtf {
                    results.push(run_rtf(entry, &corpus, warmup, samples)?);
                    rows += 1;
                }
            }
        }
        if rows == 0 {
            return Err("the selected RTF variants produce no unified/leaf RTF rows".into());
        }
    }

    let odt = family(|plan| matches!(plan, Plan::Odt { .. }));
    if !odt.is_empty() {
        let corpus = crate::build_odt_repeated_text_corpus()?;
        let source = TempSource::new("odt", &corpus.archive)?;
        for entry in odt {
            results.push(run_odt(entry, &corpus, source.path(), warmup, samples)?);
        }
        source.remove()?;
    }

    let odp = family(|plan| matches!(plan, Plan::Odp { .. }));
    if !odp.is_empty() {
        let corpus = crate::build_odp_media_corpus()?;
        let source = TempSource::new("odp", &corpus.archive)?;
        for entry in odp {
            results.push(run_odp(entry, &corpus, source.path(), warmup, samples)?);
        }
        source.remove()?;
    }

    let ods = family(|plan| matches!(plan, Plan::Ods { .. }));
    if !ods.is_empty() {
        let corpus = crate::build_ods_media_corpus()?;
        let source = TempSource::new("ods", &corpus.archive)?;
        for entry in ods {
            results.push(run_ods(entry, &corpus, source.path(), warmup, samples)?);
        }
        source.remove()?;
    }

    let detect = family(|plan| matches!(plan, Plan::Detect { .. }));
    if !detect.is_empty() {
        for shape in selection.semantic_shapes {
            for entry in &detect {
                // The detection predecessors time the facade's detection
                // functions, which M0a step S9 kept, so they stay registered.
                let predecessor = crate::parse_case(entry.predecessor)
                    .ok_or("a detection route's predecessor is not registered")?;
                let corpus = crate::build_detection_corpus(predecessor, *shape)?;
                results.push(run_detection(entry, &corpus, warmup, samples)?);
            }
        }
    }

    Ok(results)
}

const PPT_TEXT_BATCH_GENERATOR: &str = "litchi-ppt-text-batch-v1";
const PPT_TEXT_BATCH_NATIVE_CHUNK_SIZE: usize =
    litchi_ppt::slide_order::MAX_SHAPE_TEXT_REPLACEMENTS;
const PPT_TEXT_BATCH_NATIVE_OPERATION_LIMIT: usize = litchi_ppt::package::edit::MAX_OPERATIONS;
const PPT_TEXT_BATCH_LARGE_SNAPSHOT_LIMIT: usize = 1 << 30;

#[derive(Clone, Copy, Debug, Default)]
struct PptTextBatchNativeStats {
    changed: usize,
    inner_chunks: usize,
    aggregate_calls: usize,
    commits: usize,
}

#[derive(Clone, Copy, Debug, Default)]
struct PptTextBatchUnifiedStats {
    operations: usize,
    commits: usize,
}

struct PreparedPptTextBatch {
    native: litchi_ppt::Package,
    unified: litchi::File,
    replacements: Vec<litchi_ppt::slide_order::ShapeTextReplacement>,
    paths: Vec<String>,
    props: Vec<litchi_ops::Props<'static>>,
    source_sha256: String,
    expected_output: Vec<u8>,
    limits: SemanticLimits,
    native_stats: PptTextBatchNativeStats,
    unified_stats: PptTextBatchUnifiedStats,
}

enum PptTextBatchOutput {
    Native(litchi_ppt::Package),
    Unified(litchi::File),
}

fn ppt_text_batch_source_text(count: usize, index: usize) -> String {
    format!("litchi-ppt-text-batch-source-v1-{count}-{index:05}")
}

fn ppt_text_batch_updated_text(count: usize, index: usize) -> String {
    format!("litchi-ppt-text-batch-updated-v1-{count}-{index:05}")
}

const fn ppt_text_batch_shape_name(count: usize) -> &'static str {
    match count {
        1 => "1",
        10 => "10",
        100 => "100",
        256 => "256",
        1_000 => "1000",
        4_096 => "4096",
        10_000 => "10000",
        _ => "custom",
    }
}

fn build_ppt_text_batch_corpus(count: usize) -> Result<Corpus, Box<dyn Error>> {
    if count == 0 {
        return Err("PPT text-batch corpus count must be positive".into());
    }
    let mut builder = litchi_ppt::writer::Builder::new();
    let slide = builder.add_slide()?;
    for index in 0..count {
        let x = 24_i32
            .checked_add(i32::try_from(index % 16)?.saturating_mul(120))
            .ok_or("PPT text-batch x coordinate overflows i32")?;
        let y = 24_i32
            .checked_add(i32::try_from(index / 16)?.saturating_mul(72))
            .ok_or("PPT text-batch y coordinate overflows i32")?;
        builder.add_textbox(
            slide,
            x,
            y,
            96,
            48,
            &ppt_text_batch_source_text(count, index),
        )?;
    }
    let archive = builder.build()?.to_bytes()?;
    let package = litchi_ppt::Package::from_bytes(&archive)?;
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    let first = slides.first().ok_or("PPT text-batch corpus has no slide")?;
    if slides.len() != 1 || first.shape_count()? != count {
        return Err("PPT text-batch corpus shape count differs from its fixture".into());
    }
    let target_payload = ppt_text_batch_source_text(count, 0).into_bytes();
    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("ppt-text-batch-{count}"),
            generator: PPT_TEXT_BATCH_GENERATOR,
            package_format: "PPT/CFB",
            shape: ppt_text_batch_shape_name(count),
            payload_kind: "deterministic-text-boxes",
            compression: "none",
            entry_count: count,
            archive_member_count: litchi_cfb::OleFile::open(
                Cursor::new(archive.as_slice()),
                litchi_cfb::OleFileLimits::default(),
            )?
            .list_directory_entries(&[])?
            .len(),
            entry_bytes: target_payload.len(),
            uncompressed_payload_bytes: (0..count)
                .try_fold(0usize, |total, index| {
                    total.checked_add(ppt_text_batch_source_text(count, index).len())
                })
                .ok_or("PPT text-batch payload size overflows usize")?,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "slide:0/shape:0".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "slide:0/shape:0".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn open_ppt_text_batch_native(
    owner: PptTextBatchOwner,
    corpus: &Corpus,
    path: &Path,
) -> Result<litchi_ppt::Package, Box<dyn Error>> {
    Ok(match owner {
        PptTextBatchOwner::Eager => litchi_ppt::Package::from_bytes(&corpus.archive)?,
        PptTextBatchOwner::Positional => litchi_ppt::Package::open(path)?,
    })
}

fn open_ppt_text_batch_unified(
    owner: PptTextBatchOwner,
    corpus: &Corpus,
    path: &Path,
) -> Result<litchi::File, Box<dyn Error>> {
    Ok(match owner {
        PptTextBatchOwner::Eager => litchi::open_bytes(corpus.archive.clone())?,
        PptTextBatchOwner::Positional => litchi::open(path)?,
    })
}

fn apply_ppt_text_batch_native(
    package: &litchi_ppt::Package,
    replacements: &[litchi_ppt::slide_order::ShapeTextReplacement],
) -> Result<(litchi_ppt::Package, PptTextBatchNativeStats), Box<dyn Error>> {
    let mut current = package.clone();
    let mut stats = PptTextBatchNativeStats::default();
    for operation_chunk in replacements.chunks(PPT_TEXT_BATCH_NATIVE_OPERATION_LIMIT) {
        let mut edit = current.edit()?;
        let chunks = operation_chunk
            .chunks(PPT_TEXT_BATCH_NATIVE_CHUNK_SIZE)
            .collect::<Vec<_>>();
        stats.changed = stats
            .changed
            .checked_add(edit.set_shape_text_chunks(&chunks)?)
            .ok_or("PPT text-batch changed count overflows usize")?;
        stats.inner_chunks = stats
            .inner_chunks
            .checked_add(chunks.len())
            .ok_or("PPT text-batch native chunk count overflows usize")?;
        stats.aggregate_calls += 1;
        current = edit.commit()?.into_package();
        stats.commits = stats
            .commits
            .checked_add(1)
            .ok_or("PPT text-batch native commit count overflows usize")?;
    }
    Ok((current, stats))
}

fn apply_ppt_text_batch_unified(
    file: &litchi::File,
    paths: &[String],
    props: &[litchi_ops::Props<'static>],
    limits: &SemanticLimits,
) -> Result<(litchi::File, PptTextBatchUnifiedStats), Box<dyn Error>> {
    if paths.len() != props.len() {
        return Err("PPT text-batch path and property vectors differ".into());
    }
    let commit = file.edit_with_options(
        litchi::EditOptions::new().semantic_limits(*limits),
        |transaction| {
            for (path, props) in paths.iter().zip(props) {
                transaction.set(path.as_str(), props)?;
            }
            Ok(())
        },
    )?;
    let stats = PptTextBatchUnifiedStats {
        operations: commit.diagnostics().operations(),
        commits: commit.diagnostics().commits(),
    };
    Ok((commit.into_file(), stats))
}

fn prepare_ppt_text_batch(
    entry: PptTextBatchEntry,
    corpus: &Corpus,
    path: &Path,
) -> Result<PreparedPptTextBatch, Box<dyn Error>> {
    let limits = ppt_text_batch_limits(entry.count)?;
    prepare_ppt_text_batch_with_limits(entry, corpus, path, limits)
}

fn prepare_ppt_text_batch_with_limits(
    entry: PptTextBatchEntry,
    corpus: &Corpus,
    path: &Path,
    limits: SemanticLimits,
) -> Result<PreparedPptTextBatch, Box<dyn Error>> {
    let native = open_ppt_text_batch_native(entry.owner, corpus, path)?;
    let unified = open_ppt_text_batch_unified(entry.owner, corpus, path)?;
    if native.to_bytes()? != corpus.archive || unified.to_bytes()? != corpus.archive {
        return Err("PPT text-batch prepared owner changed the source before timing".into());
    }
    let mut replacements = Vec::with_capacity(entry.count);
    let mut paths = Vec::with_capacity(entry.count);
    let mut props = Vec::with_capacity(entry.count);
    for index in 0..entry.count {
        let target = litchi_ppt::text_edit::Target::new(Position::new(0), Position::new(index));
        let updated = ppt_text_batch_updated_text(entry.count, index);
        replacements.push(litchi_ppt::slide_order::ShapeTextReplacement::new(
            target, &updated,
        ));
        paths.push(format!("/slide[1]/shape[{}]", index + 1));
        props.push(litchi_ops::Props::new().with("text", updated));
    }
    let (native_output, native_stats) = apply_ppt_text_batch_native(&native, &replacements)
        .map_err(|error| format!("native PPT text-batch qualification failed: {error:?}"))?;
    let (unified_output, unified_stats) =
        apply_ppt_text_batch_unified(&unified, &paths, &props, &limits)
            .map_err(|error| format!("unified PPT text-batch qualification failed: {error:?}"))?;
    let native_bytes = native_output
        .to_bytes()
        .map_err(|error| format!("native PPT text-batch serialization failed: {error:?}"))?;
    let unified_bytes = unified_output
        .to_bytes()
        .map_err(|error| format!("unified PPT text-batch serialization failed: {error:?}"))?;
    if native_bytes != unified_bytes {
        return Err("PPT text-batch native and unified outputs differ".into());
    }
    if fs::read(path)? != corpus.archive {
        return Err("PPT text-batch qualification mutated the prepared source".into());
    }
    if native_stats.changed != entry.count || unified_stats.operations != entry.count {
        return Err("PPT text-batch qualification changed an unexpected number of targets".into());
    }
    let expected_commits = entry.count.div_ceil(PPT_TEXT_BATCH_NATIVE_OPERATION_LIMIT);
    if native_stats.commits != expected_commits
        || native_stats.aggregate_calls != expected_commits
        || native_stats.inner_chunks != entry.count.div_ceil(PPT_TEXT_BATCH_NATIVE_CHUNK_SIZE)
        || unified_stats.commits != expected_commits
    {
        return Err("PPT text-batch qualification did not coalesce at the native capacity".into());
    }
    let source_history = ppt_text_batch_history_length(&corpus.archive)?;
    let output_history = ppt_text_batch_history_length(&native_bytes)?;
    if output_history.checked_sub(source_history) != Some(expected_commits) {
        return Err("PPT aggregate publication history differs from native edit count".into());
    }
    verify_ppt_text_batch_texts(&native_bytes, entry.count)?;
    Ok(PreparedPptTextBatch {
        native,
        unified,
        replacements,
        paths,
        props,
        source_sha256: sha256_hex(&corpus.archive),
        expected_output: native_bytes,
        limits,
        native_stats,
        unified_stats,
    })
}

// Frozen v1 evidence retains the old repeated-publication contract. Aggregate
// routes have new names because the source corpus is identical but edit history
// and output bytes differ above one inner chunk.
fn ppt_text_batch_history_length(bytes: &[u8]) -> Result<usize, Box<dyn Error>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())?;
    let current = ole.open_stream(&["Current User"])?;
    let document = ole.open_stream(&["PowerPoint Document"])?;
    let mut offset = litchi_ppt::CurrentUser::parse(&current, &litchi_ppt::SourceLimits::default())?
        .current_edit_offset() as usize;
    let mut count = 0usize;
    let mut visited = std::collections::BTreeSet::new();
    while offset != 0 {
        if !visited.insert(offset) {
            return Err("PPT qualification found cyclic edit history".into());
        }
        let end = offset
            .checked_add(36)
            .ok_or("PPT history offset overflow")?;
        let record = document
            .get(offset..end)
            .ok_or("truncated PPT edit history")?;
        if u16::from_le_bytes(record[2..4].try_into()?) != 4085
            || u32::from_le_bytes(record[4..8].try_into()?) < 28
        {
            return Err("PPT qualification found invalid UserEditAtom".into());
        }
        offset = u32::from_le_bytes(record[16..20].try_into()?) as usize;
        count += 1;
    }
    Ok(count)
}

fn verify_ppt_text_batch_texts(bytes: &[u8], count: usize) -> Result<(), Box<dyn Error>> {
    let package = litchi_ppt::Package::from_bytes(bytes)?;
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    let slide = slides.first().ok_or("PPT aggregate output lacks slide")?;
    let shapes = slide.shapes()?;
    if slides.len() != 1 || shapes.len() != count {
        return Err("PPT aggregate output has unexpected shape topology".into());
    }
    for (index, shape) in shapes.iter().enumerate() {
        if shape.text()? != ppt_text_batch_updated_text(count, index) {
            return Err("PPT aggregate output text differs from requested value".into());
        }
    }
    Ok(())
}

fn run_ppt_text_batch(
    entry: PptTextBatchEntry,
    corpus: &Corpus,
    path: &Path,
    prepared: &PreparedPptTextBatch,
    warmup: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup, samples)? {
        let started = Instant::now();
        let (output, stats) = match entry.leg {
            PptTextBatchLeg::Native => {
                let (output, stats) =
                    apply_ppt_text_batch_native(&prepared.native, &prepared.replacements)?;
                (PptTextBatchOutput::Native(output), stats.changed)
            },
            PptTextBatchLeg::Unified => {
                let (output, stats) = apply_ppt_text_batch_unified(
                    &prepared.unified,
                    &prepared.paths,
                    &prepared.props,
                    &prepared.limits,
                )?;
                (PptTextBatchOutput::Unified(output), stats.operations)
            },
        };
        std::hint::black_box(&output);
        let duration = started.elapsed();
        let output_bytes = match output {
            PptTextBatchOutput::Native(output) => output.to_bytes()?,
            PptTextBatchOutput::Unified(output) => output.to_bytes()?,
        };
        if output_bytes != prepared.expected_output {
            return Err(format!(
                "{} produced output different from its qualified native output",
                entry.case.name()
            )
            .into());
        }
        if stats != entry.count {
            return Err(format!(
                "{} applied {stats} targets, expected {}",
                entry.case.name(),
                entry.count
            )
            .into());
        }
        if fs::read(path)? != corpus.archive {
            return Err(format!("{} mutated its prepared source", entry.case.name()).into());
        }
        record_elapsed(&mut elapsed, iteration, warmup, duration)?;
    }
    let summary = RouteSummary {
        leg: entry.leg.name(),
        predecessor: "ppt_text_batch_native_vs_unified_v1",
        twin: Some(entry.twin.name()),
        timed_calls: match entry.leg {
            PptTextBatchLeg::Native => {
                "prepared litchi_ppt::Package::edit with one set_shape_text_chunks publication per 16384-operation edit, inner chunks bounded at 256"
            },
            PptTextBatchLeg::Unified => {
                "prepared litchi::File::edit shape text sets from string paths (path parsing included) in one facade transaction with 16384-operation native rollover"
            },
        },
        performance_claim: PERFORMANCE_CLAIM,
        projection_sha256: Some(sha256_hex(&prepared.expected_output)),
        verified_samples: Some(samples),
        observation: None,
        selected_target: None,
        legs_agree: Some(true),
        real_file: None,
        ppt_text_batch: Some(PptTextBatchSummary {
            leg: entry.leg.name(),
            owner: entry.owner.name(),
            count: entry.count,
            native_chunk_size: PPT_TEXT_BATCH_NATIVE_CHUNK_SIZE,
            native_operation_limit: PPT_TEXT_BATCH_NATIVE_OPERATION_LIMIT,
            native_inner_chunks: prepared.native_stats.inner_chunks,
            native_aggregate_calls: prepared.native_stats.aggregate_calls,
            native_publications: prepared.native_stats.commits,
            native_commits: prepared.native_stats.commits,
            unified_operations: prepared.unified_stats.operations,
            unified_commits: prepared.unified_stats.commits,
            semantic_snapshot_limit_bytes: prepared.limits.max_snapshot_bytes(),
            source_sha256: prepared.source_sha256.clone(),
            expected_output_sha256: sha256_hex(&prepared.expected_output),
            exact_output_verified: true,
            unchanged_source_verified: true,
            timing_scope: "one prepared native or unified text-batch publication; opening, fixture construction, command-vector preparation, output serialization, exact equivalence and source checks are outside the timer; unified string-path parsing is timed",
        }),
    };
    Ok(result_with_source(
        entry.case,
        corpus,
        elapsed,
        SourceSummary {
            unified_ops: Some(Box::new(summary)),
            ..SourceSummary::default()
        },
    ))
}

fn ppt_text_batch_limits(count: usize) -> Result<SemanticLimits, Box<dyn Error>> {
    if count <= PPT_TEXT_BATCH_NATIVE_OPERATION_LIMIT {
        return Ok(SemanticLimits::default());
    }
    Ok(SemanticLimits::builder()
        .max_snapshot_bytes(PPT_TEXT_BATCH_LARGE_SNAPSHOT_LIMIT)?
        .build()?)
}

fn run_ppt_text_batch_selected(
    selected: &[PptTextBatchEntry],
    warmup: usize,
    samples: usize,
) -> Result<Vec<CaseResult>, Box<dyn Error>> {
    let mut results = Vec::new();
    let mut prepared_keys = Vec::new();
    for entry in selected {
        let key = (entry.owner, entry.count);
        if prepared_keys.contains(&key) {
            continue;
        }
        prepared_keys.push(key);
        let corpus = build_ppt_text_batch_corpus(entry.count)?;
        let source = TempSource::new("ppt", &corpus.archive)?;
        let prepared = prepare_ppt_text_batch(*entry, &corpus, source.path())?;
        for selected_entry in selected
            .iter()
            .copied()
            .filter(|candidate| (candidate.owner, candidate.count) == key)
        {
            results.push(run_ppt_text_batch(
                selected_entry,
                &corpus,
                source.path(),
                &prepared,
                warmup,
                samples,
            )?);
        }
        source.remove()?;
    }
    Ok(results)
}

/// A corpus written to a private scratch directory for the path routes, in
/// the temporary directory the predecessors also write to; removed when the
/// family finishes (and on drop, if a route failed first).
struct TempSource {
    root: PathBuf,
    path: PathBuf,
    removed: bool,
}

impl TempSource {
    fn new(extension: &str, bytes: &[u8]) -> Result<Self, Box<dyn Error>> {
        let root = crate::filesystem::scratch_root(None, "unified-ops")?;
        let source = Self {
            path: root.join(format!("source.{extension}")),
            root,
            removed: false,
        };
        let mut file = fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&source.path)?;
        file.write_all(bytes)?;
        file.flush()?;
        Ok(source)
    }

    fn path(&self) -> &Path {
        &self.path
    }

    fn remove(mut self) -> Result<(), Box<dyn Error>> {
        self.removed = true;
        fs::remove_dir_all(&self.root)?;
        Ok(())
    }
}

impl Drop for TempSource {
    fn drop(&mut self) {
        if !self.removed {
            let _ = fs::remove_dir_all(&self.root);
        }
    }
}

fn unexpected_plan(entry: &Entry) -> Box<dyn Error> {
    format!("{} reached the wrong in-process runner", entry.case.name()).into()
}

fn verify_projection<T: PartialEq>(
    entry: &Entry,
    observed: &T,
    expected: &T,
) -> Result<(), Box<dyn Error>> {
    if observed == expected {
        Ok(())
    } else {
        Err(format!(
            "{} sample projection differs from its oracle",
            entry.case.name()
        )
        .into())
    }
}

// ---------------------------------------------------------------------------
// XLSX bytes (xlsx_bytes_*)
// ---------------------------------------------------------------------------

#[derive(Debug, PartialEq, Serialize)]
struct WorkbookProjection {
    names: Vec<String>,
    count: usize,
    text_sha256: String,
    metadata_sha256: String,
}

fn workbook_projection(
    names: Vec<String>,
    count: usize,
    text: &str,
    metadata_sha256: String,
) -> WorkbookProjection {
    WorkbookProjection {
        names,
        count,
        text_sha256: sha256_hex(text.as_bytes()),
        metadata_sha256,
    }
}

fn unified_workbook_names_count_text(
    file: &litchi::File,
) -> Result<(Vec<String>, usize, String), Box<dyn Error>> {
    let workbook = file.workbook()?;
    Ok((
        workbook.worksheet_names()?,
        workbook.worksheet_count()?,
        file.text()?,
    ))
}

fn run_xlsx_bytes(
    entry: &Entry,
    corpus: &Corpus,
    warmup: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let Plan::XlsxBytes { lifecycle } = entry.plan else {
        return Err(unexpected_plan(entry));
    };
    // The predecessor's oracle: the typed eager owner and an independently
    // opened OPC/property package, both outside the timer.
    let spec = crate::xlsx_spec(corpus)?;
    let typed = litchi_xlsx::Package::from_bytes(&corpus.archive)?;
    crate::verify_xlsx_cells(&typed, spec, &[])?;
    let expected = workbook_projection(
        typed
            .workbook_ref()
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect(),
        typed.workbook_ref().len(),
        &crate::xlsx_typed_full_text(&typed, spec)?,
        crate::xlsx_root_metadata_digest(&crate::xlsx_root_metadata_oracle(
            &OpcPackage::from_bytes(&corpus.archive, litchi_opc::SourceLimits::default())?,
        )?)?,
    );
    drop(typed);

    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup, samples)? {
        // Vec ownership is prepared outside the timer, as in the predecessor.
        let owned = corpus.archive.clone();
        let (duration, observed) = match entry.leg {
            Leg::Unified => {
                let started = Instant::now();
                let file = litchi::open_bytes(owned)?;
                let timed = lifecycle
                    .then(|| unified_workbook_names_count_text(&file))
                    .transpose()?;
                let duration = started.elapsed();
                let (names, count, text) = match timed {
                    Some(projection) => projection,
                    None => unified_workbook_names_count_text(&file)?,
                };
                let metadata = file.metadata()?.unwrap_or_default();
                std::hint::black_box(&file);
                (
                    duration,
                    workbook_projection(
                        names,
                        count,
                        &text,
                        crate::xlsx_root_metadata_digest(&metadata)?,
                    ),
                )
            },
            Leg::Leaf => {
                let started = Instant::now();
                let leaf = leaf_xlsx_open(Arc::new(OwnedSource::new(owned)))?;
                let timed = lifecycle
                    .then(|| -> Result<_, Box<dyn Error>> {
                        Ok((
                            leaf_xlsx_names(&leaf.workbook),
                            leaf.workbook.len(),
                            leaf_xlsx_text(&leaf.workbook)?,
                        ))
                    })
                    .transpose()?;
                let duration = started.elapsed();
                let (names, count, text) = match timed {
                    Some(projection) => projection,
                    None => (
                        leaf_xlsx_names(&leaf.workbook),
                        leaf.workbook.len(),
                        leaf_xlsx_text(&leaf.workbook)?,
                    ),
                };
                let metadata_sha256 = crate::xlsx_root_metadata_digest(&leaf.metadata)?;
                std::hint::black_box(&leaf.workbook);
                (
                    duration,
                    workbook_projection(names, count, &text, metadata_sha256),
                )
            },
        };
        verify_projection(entry, &observed, &expected)?;
        record_elapsed(&mut elapsed, iteration, warmup, duration)?;
    }
    let timed_calls = match (entry.leg, lifecycle) {
        (Leg::Unified, false) => {
            "litchi::open_bytes(Vec<u8>) -> File::Xlsx (source-backed over the owned bytes)"
        },
        (Leg::Unified, true) => {
            "litchi::open_bytes(Vec<u8>) + File::workbook()?.worksheet_names() + worksheet_count() + File::text()"
        },
        (Leg::Leaf, false) => {
            "SourceBackedPackage::from_read_at(OwnedSource,
            litchi_opc::SourceLimits::default(),
            litchi_opc::SourceCacheLimits::default(),
            litchi_opc::SourceReadPolicy::exact(),
            None) + core-properties read + litchi_xlsx::SourceBackedWorkbook::from_source_backed_package"
        },
        (Leg::Leaf, true) => {
            "leaf open (as for open) + sheet names + len() + a row-by-row cell-text projection over SourceWorksheet::cells"
        },
    };
    let summary = RouteSummary::for_case(entry.case, timed_calls)?.verified(&expected, samples)?;
    Ok(finish(entry, corpus, elapsed, summary))
}

// ---------------------------------------------------------------------------
// DOCX read-one (docx_semantic_*, docx_producer_*)
// ---------------------------------------------------------------------------

fn run_docx_semantic(
    entry: &Entry,
    corpus: &Corpus,
    warmup: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let Plan::DocxSemantic { text } = entry.plan else {
        return Err(unexpected_plan(entry));
    };
    let shape = crate::semantic_shape(corpus)?;
    let index = shape.docx_paragraphs() / 2;
    let expected = crate::semantic_docx_text(index, false);
    // The corpus-level oracle the predecessor checks, once, on an independent
    // eager owner.
    crate::verify_semantic_docx(
        &litchi_docx::Package::from_reader(Cursor::new(corpus.archive.clone()))?,
        shape,
        &[],
    )?;

    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup, samples)? {
        // Each iteration prepares a fresh owner outside the timer: the
        // unified eager owner, or the leaf owner built by the same leaf calls.
        let (duration, observed) = match (entry.leg, text) {
            (Leg::Unified, false) => {
                let file = litchi::open_bytes_with(corpus.archive.clone(), &eager_options())?;
                let started = Instant::now();
                let paragraph = file.document()?.paragraph(index)?;
                let duration = started.elapsed();
                (
                    duration,
                    paragraph
                        .ok_or("unified DOCX selected paragraph is missing")?
                        .text()?,
                )
            },
            (Leg::Unified, true) => {
                let file = litchi::open_bytes_with(corpus.archive.clone(), &eager_options())?;
                let paragraph = file
                    .document()?
                    .paragraph(index)?
                    .ok_or("unified DOCX selected paragraph is missing")?;
                let started = Instant::now();
                let result = paragraph.text();
                let duration = started.elapsed();
                (duration, result?)
            },
            (Leg::Leaf, false) => {
                let package = leaf_docx_eager(corpus.archive.clone())?;
                let started = Instant::now();
                let paragraph = package.document()?.paragraph(index)?;
                let duration = started.elapsed();
                (
                    duration,
                    paragraph
                        .ok_or("leaf DOCX selected paragraph is missing")?
                        .text()?,
                )
            },
            (Leg::Leaf, true) => {
                let package = leaf_docx_eager(corpus.archive.clone())?;
                let document = package.document()?;
                let paragraph = document
                    .paragraph(index)?
                    .ok_or("leaf DOCX selected paragraph is missing")?;
                let started = Instant::now();
                let result = paragraph.text();
                let duration = started.elapsed();
                (duration, result?)
            },
        };
        verify_projection(entry, &observed, &expected)?;
        record_elapsed(&mut elapsed, iteration, warmup, duration)?;
    }
    let timed_calls = match (entry.leg, text) {
        (Leg::Unified, false) => {
            "File::document()?.paragraph(middle) on an eager File (open_bytes_with SourceMode::Eager) prepared outside the timer"
        },
        (Leg::Unified, true) => {
            "compat Paragraph::text() on the middle paragraph of an eager File, selected outside the timer"
        },
        (Leg::Leaf, false) => {
            "Package::document()?.paragraph(middle) on an eager litchi_docx::Package prepared outside the timer by the unified eager owner's leaf calls"
        },
        (Leg::Leaf, true) => {
            "litchi_docx::Paragraph::text() on the middle paragraph, selected outside the timer"
        },
    };
    let summary = RouteSummary::for_case(entry.case, timed_calls)?
        .verified(&(index, sha256_hex(expected.as_bytes())), samples)?;
    Ok(finish(entry, corpus, elapsed, summary))
}

fn run_docx_producer(
    entry: &Entry,
    producer: &crate::producer_shape::ProducerCorpus,
    warmup: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if !matches!(entry.plan, Plan::DocxProducer) {
        return Err(unexpected_plan(entry));
    }
    let index = producer.evidence.selected_sheet;
    let expected = producer.evidence.selected_target_text.clone();
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup, samples)? {
        let archive = producer.corpus.archive.clone();
        // A fresh source-backed owner per iteration, its main story derived
        // once outside the timer, as the predecessor prepares its document.
        let (duration, observed) = match entry.leg {
            Leg::Unified => {
                let file = litchi::open_bytes(archive)?;
                if !matches!(&file, litchi::File::Docx(docx) if docx.is_source_backed()) {
                    return Err("unified producer DOCX did not open source-backed".into());
                }
                file.document()?.paragraph_count()?;
                let started = Instant::now();
                let text = file.document()?.paragraph_text(index)?;
                let duration = started.elapsed();
                (duration, text)
            },
            Leg::Leaf => {
                let package = litchi_docx::source_backed::Package::from_read_at(Arc::new(
                    OwnedSource::new(archive),
                ))?;
                package.document()?.paragraph_count()?;
                let started = Instant::now();
                let text = package.document()?.paragraph_text(index)?;
                let duration = started.elapsed();
                (duration, text)
            },
        };
        verify_projection(entry, &observed.as_deref(), &Some(expected.as_str()))?;
        record_elapsed(&mut elapsed, iteration, warmup, duration)?;
    }
    let timed_calls = match entry.leg {
        Leg::Unified => {
            "File::document()?.paragraph_text(selected) on a source-backed File (open_bytes) whose main story was derived outside the timer"
        },
        Leg::Leaf => {
            "source_backed::Package::document()?.paragraph_text(selected) on a package whose main story was derived outside the timer"
        },
    };
    let summary = RouteSummary::for_case(entry.case, timed_calls)?
        .verified(&(index, sha256_hex(expected.as_bytes())), samples)?;
    Ok(finish(entry, &producer.corpus, elapsed, summary))
}

// ---------------------------------------------------------------------------
// Legacy OLE2 (doc_file_*, ppt_file_*)
// ---------------------------------------------------------------------------

/// The predecessor facade family's corpus identity, kept identical so the
/// unified and leaf rows share the catalog entry of the bytes they read.
const DOC_GENERATOR: &str = "litchi-doc-facade-real-file-v1";
const PPT_GENERATOR: &str = "litchi-ppt-facade-real-file-v1";

struct Ole2Corpus {
    corpus: Corpus,
    path: PathBuf,
    provenance: crate::producer_shape::RealFileProvenance,
    /// The selected paragraph (DOC) or slide (PPT).
    target: usize,
    /// Frozen per-leg outcomes, keyed by query.
    oracles: Vec<(Leg, Ole2Query, String)>,
}

fn text_outcome<E: std::fmt::Display>(text: Result<String, E>) -> String {
    match text {
        Ok(text) => format!("text:{}:{}", text.len(), sha256_hex(text.as_bytes())),
        Err(error) => format!("refused:{error}"),
    }
}

fn optional_text_outcome<E: std::fmt::Display>(text: Result<Option<String>, E>) -> String {
    match text {
        Ok(Some(text)) => format!("text:{}:{}", text.len(), sha256_hex(text.as_bytes())),
        Ok(None) => "absent".to_owned(),
        Err(error) => format!("refused:{error}"),
    }
}

/// The leaf calls the unified DOC open makes: the whole file read once (here
/// through the harness's bounded reader of a caller-named fixture), one
/// in-memory CFB parse, the DOC package, its reader, and the OLE
/// property-set metadata.
fn leaf_doc_open(path: &Path) -> Result<litchi_doc::Document, Box<dyn Error>> {
    let bytes = crate::ole2_range_source::read_bounded(path)?;
    let ole = litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())?;
    let mut package = litchi_doc::raw::Reader::from_ole_file(ole)?;
    let document = package.document()?;
    std::hint::black_box(package.ole_file().get_metadata().map(Metadata::from).ok());
    Ok(document)
}

/// The leaf calls the unified PPT path open makes: the positional package,
/// its OLE property-set metadata, and the reader.
fn leaf_ppt_open(path: &Path) -> Result<litchi_ppt::Presentation, Box<dyn Error>> {
    let package = litchi_ppt::SourceBackedPackage::open(path)?;
    std::hint::black_box(package.metadata().ok().map(Metadata::from));
    Ok(package.presentation()?)
}

fn leaf_doc_paragraph_text(
    document: &litchi_doc::Document,
    index: usize,
) -> Result<Option<String>, Box<dyn Error>> {
    document
        .paragraphs()?
        .into_iter()
        .nth(index)
        .map(|paragraph| Ok(paragraph.text()?.to_owned()))
        .transpose()
}

/// The leaf slide read the unified view makes: the slide, its text and its
/// shape count.
fn leaf_ppt_slide_text(
    presentation: &litchi_ppt::Presentation,
    index: usize,
) -> Result<Option<String>, Box<dyn Error>> {
    let Some(slide) = presentation.slide_at(index)? else {
        return Ok(None);
    };
    let text = slide.text()?.to_string();
    std::hint::black_box(slide.shape_count().unwrap_or(0));
    Ok(Some(text))
}

/// One fresh-open sample of `query` on `leg`: the elapsed time of the timed
/// region and the projection derived from it. A refusal is a projection, as
/// in the predecessor family: the route reads the file before refusing.
fn ole2_sample(
    leg: Leg,
    query: Ole2Query,
    path: &Path,
    target: usize,
) -> Result<(Duration, String), Box<dyn Error>> {
    Ok(match (leg, query) {
        (Leg::Unified, Ole2Query::DocOpen) => {
            let started = Instant::now();
            let file = litchi::open(path);
            let duration = started.elapsed();
            let projection = match &file {
                Ok(file) => format!("paragraphs:{}", file.document()?.paragraph_count()?),
                Err(error) => format!("refused:{error}"),
            };
            drop(file);
            (duration, projection)
        },
        (Leg::Leaf, Ole2Query::DocOpen) => {
            let started = Instant::now();
            let document = leaf_doc_open(path);
            let duration = started.elapsed();
            let projection = match &document {
                Ok(document) => format!("paragraphs:{}", document.paragraph_count()?),
                Err(error) => format!("refused:{error}"),
            };
            drop(document);
            (duration, projection)
        },
        (Leg::Unified, Ole2Query::DocFullText) | (Leg::Unified, Ole2Query::PptFullText) => {
            let started = Instant::now();
            let projection = match litchi::open(path) {
                Ok(file) => text_outcome(std::hint::black_box(file.text())),
                Err(error) => format!("refused:{error}"),
            };
            (started.elapsed(), projection)
        },
        (Leg::Leaf, Ole2Query::DocFullText) => {
            let started = Instant::now();
            let projection = match leaf_doc_open(path) {
                Ok(document) => text_outcome(std::hint::black_box(document.text())),
                Err(error) => format!("refused:{error}"),
            };
            (started.elapsed(), projection)
        },
        (Leg::Unified, Ole2Query::DocOneParagraph) => {
            let started = Instant::now();
            let projection = match litchi::open(path) {
                Ok(file) => optional_text_outcome(std::hint::black_box(
                    file.document()
                        .and_then(|document| document.paragraph_text(target)),
                )),
                Err(error) => format!("refused:{error}"),
            };
            (started.elapsed(), projection)
        },
        (Leg::Leaf, Ole2Query::DocOneParagraph) => {
            let started = Instant::now();
            let projection = match leaf_doc_open(path) {
                Ok(document) => optional_text_outcome(std::hint::black_box(
                    leaf_doc_paragraph_text(&document, target),
                )),
                Err(error) => format!("refused:{error}"),
            };
            (started.elapsed(), projection)
        },
        (Leg::Unified, Ole2Query::PptOpen) => {
            let started = Instant::now();
            let file = litchi::open(path);
            let duration = started.elapsed();
            let projection = match &file {
                Ok(file) => format!("slides:{}", file.presentation()?.slide_count()?),
                Err(error) => format!("refused:{error}"),
            };
            drop(file);
            (duration, projection)
        },
        (Leg::Leaf, Ole2Query::PptOpen) => {
            let started = Instant::now();
            let presentation = leaf_ppt_open(path);
            let duration = started.elapsed();
            let projection = match &presentation {
                Ok(presentation) => format!("slides:{}", presentation.slide_count()),
                Err(error) => format!("refused:{error}"),
            };
            drop(presentation);
            (duration, projection)
        },
        (Leg::Leaf, Ole2Query::PptFullText) => {
            let started = Instant::now();
            let projection = match leaf_ppt_open(path) {
                Ok(presentation) => text_outcome(std::hint::black_box(presentation.text())),
                Err(error) => format!("refused:{error}"),
            };
            (started.elapsed(), projection)
        },
        (Leg::Unified, Ole2Query::PptOneSlideText) => {
            let started = Instant::now();
            let projection = match litchi::open(path) {
                Ok(file) => optional_text_outcome(std::hint::black_box(
                    file.presentation()
                        .and_then(|presentation| presentation.slide(target))
                        .and_then(|slide| slide.map(|slide| slide.text()).transpose()),
                )),
                Err(error) => format!("refused:{error}"),
            };
            (started.elapsed(), projection)
        },
        (Leg::Leaf, Ole2Query::PptOneSlideText) => {
            let started = Instant::now();
            let projection = match leaf_ppt_open(path) {
                Ok(presentation) => optional_text_outcome(std::hint::black_box(
                    leaf_ppt_slide_text(&presentation, target),
                )),
                Err(error) => format!("refused:{error}"),
            };
            (started.elapsed(), projection)
        },
    })
}

/// Builds the corpus of a caller-named DOC or PPT fixture: the predecessor
/// family's corpus identity, the scenario target derived through the unified
/// API exactly as the predecessor derived it through the facade, and both
/// legs' frozen outcomes.
fn build_ole2_corpus(path: &Path, is_doc: bool) -> Result<Ole2Corpus, Box<dyn Error>> {
    use crate::ole2_range_source::{cfb_inventory, provenance_of, read_bounded};

    let archive = read_bounded(path)?;
    let provenance = provenance_of(path, &archive)?;
    let (streams, sector_size, target_path, target_payload) = if is_doc {
        cfb_inventory(&archive, &[&["WordDocument"]], "WordDocument")?
    } else {
        cfb_inventory(
            &archive,
            &[
                &["PowerPoint Document"],
                &["PP97_DUALSTORAGE", "PowerPoint Document"],
            ],
            "PowerPoint Document",
        )?
    };
    let target_entry = target_path.join("/");

    let (count, target) = match litchi::open(path) {
        Ok(file) if is_doc => {
            let document = file.document()?;
            let count = document.paragraph_count()?;
            let mut populated = Vec::new();
            for index in 0..count {
                if document
                    .paragraph_text(index)?
                    .is_some_and(|text| !text.trim().is_empty())
                {
                    populated.push(index);
                }
            }
            // The median populated paragraph, so the target is not
            // systematically the title.
            let target = populated.get(populated.len() / 2).copied().unwrap_or(0);
            (Some(count), target)
        },
        Ok(file) => {
            let presentation = file.presentation()?;
            let count = presentation.slide_count()?;
            let mut target = 0;
            for index in 0..count {
                if presentation
                    .slide(index)?
                    .map(|slide| slide.text())
                    .transpose()?
                    .is_some_and(|text| !text.trim().is_empty())
                {
                    target = index;
                    break;
                }
            }
            (Some(count), target)
        },
        Err(_) => (None, 0),
    };

    let queries = if is_doc {
        [
            Ole2Query::DocOpen,
            Ole2Query::DocFullText,
            Ole2Query::DocOneParagraph,
        ]
    } else {
        [
            Ole2Query::PptOpen,
            Ole2Query::PptFullText,
            Ole2Query::PptOneSlideText,
        ]
    };
    let mut oracles = Vec::with_capacity(6);
    for leg in [Leg::Unified, Leg::Leaf] {
        for query in queries {
            oracles.push((leg, query, ole2_sample(leg, query, path, target)?.1));
        }
    }

    let (generator, name, package_format) = if is_doc {
        (DOC_GENERATOR, "doc-facade-real-file", "DOC/CFB/OLE2")
    } else {
        (PPT_GENERATOR, "ppt-facade-real-file", "PPT/CFB/OLE2")
    };
    let manifest = crate::CorpusManifest {
        name: name.to_owned(),
        generator,
        package_format,
        shape: "real-file",
        payload_kind: "real-producer-ole2",
        compression: "none",
        entry_count: count.unwrap_or(0),
        archive_member_count: streams.len(),
        entry_bytes: sector_size,
        uncompressed_payload_bytes: target_payload.len(),
        archive_bytes: archive.len(),
        archive_sha256: sha256_hex(&archive),
        target_entry: target_entry.clone(),
        target_payload_bytes: target_payload.len(),
        target_payload_sha256: sha256_hex(&target_payload),
        rtf_variant: None,
        xlsx: None,
    };
    Ok(Ole2Corpus {
        corpus: Corpus {
            manifest,
            archive,
            target_name: target_entry,
            target_payload,
            xlsx: None,
        },
        path: path.to_path_buf(),
        provenance,
        target,
        oracles,
    })
}

fn ole2_oracle(corpus: &Ole2Corpus, leg: Leg, query: Ole2Query) -> Result<&str, Box<dyn Error>> {
    corpus
        .oracles
        .iter()
        .find(|(oracle_leg, oracle_query, _)| *oracle_leg == leg && *oracle_query == query)
        .map(|(_, _, observation)| observation.as_str())
        .ok_or_else(|| "OLE2 scenario has no frozen oracle".into())
}

fn run_ole2(
    entry: &Entry,
    query: Ole2Query,
    corpus: &Ole2Corpus,
    warmup: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let oracle = ole2_oracle(corpus, entry.leg, query)?.to_owned();
    let other = match entry.leg {
        Leg::Unified => Leg::Leaf,
        Leg::Leaf => Leg::Unified,
    };
    let twin_oracle = ole2_oracle(corpus, other, query)?;
    // Both legs read the same file: when neither refuses, they must agree.
    let both_succeeded = !oracle.starts_with("refused:") && !twin_oracle.starts_with("refused:");
    let legs_agree = oracle == twin_oracle;
    if both_succeeded && !legs_agree {
        return Err(format!(
            "{} and its twin reached different outcomes on the same file: {oracle} vs {twin_oracle}",
            entry.case.name()
        )
        .into());
    }

    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup, samples)? {
        let (duration, projection) = ole2_sample(entry.leg, query, &corpus.path, corpus.target)?;
        if projection != oracle {
            return Err(format!(
                "{} produced a sample that differs from its frozen outcome",
                entry.case.name()
            )
            .into());
        }
        record_elapsed(&mut elapsed, iteration, warmup, duration)?;
    }
    let timed_calls = match (entry.leg, query) {
        (Leg::Unified, Ole2Query::DocOpen) => {
            "litchi::open(Path) -> File::Doc; the projection is read after the timer"
        },
        (Leg::Unified, Ole2Query::DocFullText) => "fresh litchi::open(Path) + File::text()",
        (Leg::Unified, Ole2Query::DocOneParagraph) => {
            "fresh litchi::open(Path) + File::document()?.paragraph_text(target)"
        },
        (Leg::Unified, Ole2Query::PptOpen) => {
            "litchi::open(Path) -> File::Ppt; the projection is read after the timer"
        },
        (Leg::Unified, Ole2Query::PptFullText) => "fresh litchi::open(Path) + File::text()",
        (Leg::Unified, Ole2Query::PptOneSlideText) => {
            "fresh litchi::open(Path) + File::presentation()?.slide(target) + compat Slide::text()"
        },
        (Leg::Leaf, Ole2Query::DocOpen) => {
            "fs::read + litchi_cfb::OleFile::open + litchi_doc::raw::Reader::from_ole_file + Package::document() + OLE property-set metadata"
        },
        (Leg::Leaf, Ole2Query::DocFullText) => {
            "fresh leaf DOC open (as for open) + Document::text()"
        },
        (Leg::Leaf, Ole2Query::DocOneParagraph) => {
            "fresh leaf DOC open (as for open) + Document::paragraphs() + nth(target).text()"
        },
        (Leg::Leaf, Ole2Query::PptOpen) => {
            "litchi_ppt::SourceBackedPackage::open + OLE property-set metadata + presentation()"
        },
        (Leg::Leaf, Ole2Query::PptFullText) => {
            "fresh leaf PPT open (as for open) + Presentation::text()"
        },
        (Leg::Leaf, Ole2Query::PptOneSlideText) => {
            "fresh leaf PPT open (as for open) + Presentation::slide_at(target) + Slide::text() + shape_count()"
        },
    };
    let mut summary =
        RouteSummary::for_case(entry.case, timed_calls)?.verified(&oracle, samples)?;
    summary.observation = Some(oracle);
    summary.selected_target = Some(corpus.target);
    summary.legs_agree = Some(legs_agree);
    summary.real_file = Some(corpus.provenance.clone());
    Ok(finish(entry, &corpus.corpus, elapsed, summary))
}

// ---------------------------------------------------------------------------
// RTF (rtf_file_*)
// ---------------------------------------------------------------------------

#[derive(Debug, PartialEq, Serialize)]
struct TextProjection {
    text_sha256: String,
    paragraphs: Vec<String>,
}

fn run_rtf(
    entry: &Entry,
    corpus: &Corpus,
    warmup: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let Plan::Rtf { lifecycle } = entry.plan else {
        return Err(unexpected_plan(entry));
    };
    let shape = crate::semantic_shape(corpus)?;
    let variant = crate::semantic_rtf_variant(corpus)?;
    // The predecessor's native oracle, which does not use the code under
    // measurement.
    let native = litchi_rtf::Package::from_bytes(&corpus.archive)?;
    crate::verify_semantic_rtf(&native, shape, variant, &[])?;
    let expected = TextProjection {
        text_sha256: sha256_hex(native.text().as_bytes()),
        paragraphs: native
            .body()
            .paragraphs()
            .map(|paragraph| paragraph.to_text())
            .collect(),
    };
    let expected_text = native.text().to_owned();
    drop(native);
    // Match the raw leaf's required profile argument without charging its
    // construction to the timed parse.  OpenOptions is Copy, so each sample
    // receives the same checked profile value.
    let leaf_options = litchi_rtf::read::OpenOptions::default();

    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup, samples)? {
        let owned = corpus.archive.clone();
        let (duration, observed, lifecycle_text) = match entry.leg {
            Leg::Unified => {
                let started = Instant::now();
                let file = litchi::open_bytes(owned)?;
                let text = lifecycle.then(|| file.text()).transpose()?;
                let duration = started.elapsed();
                let observed = TextProjection {
                    text_sha256: sha256_hex(file.text()?.as_bytes()),
                    paragraphs: file
                        .document()?
                        .paragraphs()?
                        .iter()
                        .map(litchi::compat::word::Paragraph::text)
                        .collect::<Result<_, _>>()?,
                };
                std::hint::black_box(&file);
                (duration, observed, text)
            },
            Leg::Leaf => {
                let started = Instant::now();
                let document = litchi_rtf::raw::Tree::from_owned_bytes(owned, leaf_options)?;
                let text = lifecycle.then(|| document.text());
                let duration = started.elapsed();
                let observed = TextProjection {
                    text_sha256: sha256_hex(document.text().as_bytes()),
                    paragraphs: document
                        .paragraphs_with_content()
                        .iter()
                        .map(|paragraph| paragraph.text())
                        .collect(),
                };
                std::hint::black_box(&document);
                (duration, observed, text)
            },
        };
        verify_projection(entry, &observed, &expected)?;
        if let Some(text) = lifecycle_text
            && text != expected_text
        {
            return Err(format!("{} timed text differs from the oracle", entry.case.name()).into());
        }
        record_elapsed(&mut elapsed, iteration, warmup, duration)?;
    }
    let timed_calls = match (entry.leg, lifecycle) {
        (Leg::Unified, false) => "litchi::open_bytes(Vec<u8>) -> File::Rtf",
        (Leg::Unified, true) => "litchi::open_bytes(Vec<u8>) + File::text()",
        (Leg::Leaf, false) => "litchi_rtf::raw::Tree::from_owned_bytes(Vec<u8>, litchi_rtf::read::OpenOptions::default()) [options prepared outside timer]",
        (Leg::Leaf, true) => "litchi_rtf::raw::Tree::from_owned_bytes(Vec<u8>, litchi_rtf::read::OpenOptions::default()) [options prepared outside timer] + Tree::text()",
    };
    let summary = RouteSummary::for_case(entry.case, timed_calls)?.verified(&expected, samples)?;
    Ok(finish(entry, corpus, elapsed, summary))
}

// ---------------------------------------------------------------------------
// ODT (odt_file_*)
// ---------------------------------------------------------------------------

#[allow(
    clippy::large_enum_variant,
    reason = "a stack value of the harness; boxing a leaf owner would add an allocation to the timed leaf leg"
)]
enum OdtRoot {
    Unified(litchi::File),
    Eager(litchi_odt::Document),
    Source(litchi_odt::SourceBackedDocument),
}

impl OdtRoot {
    fn open(leg: Leg, ingress: Ingress, path: &Path) -> Result<Self, Box<dyn Error>> {
        Ok(match (leg, ingress) {
            (Leg::Unified, Ingress::Eager) => Self::Unified(litchi::open_bytes(fs::read(path)?)?),
            (Leg::Unified, Ingress::Source) => Self::Unified(litchi::open(path)?),
            (Leg::Leaf, Ingress::Eager) => {
                Self::Eager(litchi_odt::Package::from_bytes(&fs::read(path)?)?.document()?)
            },
            (Leg::Leaf, Ingress::Source) => {
                Self::Source(litchi_odt::SourceBackedDocument::open(path)?)
            },
        })
    }

    fn text(&self) -> Result<String, Box<dyn Error>> {
        Ok(match self {
            Self::Unified(file) => file.text()?,
            Self::Eager(document) => document.text()?,
            Self::Source(document) => document.text()?,
        })
    }

    fn paragraph_count(&self) -> Result<usize, Box<dyn Error>> {
        Ok(match self {
            Self::Unified(file) => file.document()?.paragraph_count()?,
            Self::Eager(document) => document.paragraph_count()?,
            Self::Source(document) => document.paragraph_count()?,
        })
    }
}

fn run_odt(
    entry: &Entry,
    corpus: &Corpus,
    path: &Path,
    warmup: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let Plan::Odt { ingress, lifecycle } = entry.plan else {
        return Err(unexpected_plan(entry));
    };
    // The oracle: the eager leaf owner over the corpus bytes, checked against
    // the generator.
    let oracle = litchi_odt::Package::from_bytes(&corpus.archive)?.document()?;
    let expected_paragraphs = SemanticShape::Large.docx_paragraphs();
    let expected_text = oracle.text()?;
    let generated = (0..expected_paragraphs)
        .map(|index| crate::semantic_odt_text(index, false))
        .collect::<Vec<_>>()
        .join("\n");
    if oracle.paragraph_count()? != expected_paragraphs || expected_text != generated {
        return Err("ODT repeated-text corpus differs from its generator".into());
    }
    drop(oracle);
    if fs::read(path)? != corpus.archive {
        return Err("ODT temporary source differs from the corpus".into());
    }
    let expected = (expected_paragraphs, sha256_hex(expected_text.as_bytes()));

    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup, samples)? {
        let started = Instant::now();
        let root = OdtRoot::open(entry.leg, ingress, path)?;
        let text = lifecycle.then(|| root.text()).transpose()?;
        let duration = started.elapsed();
        if let Some(text) = text.as_deref()
            && text != expected_text
        {
            return Err(format!("{} timed text differs from the oracle", entry.case.name()).into());
        }
        let observed = (root.paragraph_count()?, sha256_hex(root.text()?.as_bytes()));
        verify_projection(entry, &observed, &expected)?;
        std::hint::black_box(root);
        std::hint::black_box(text);
        record_elapsed(&mut elapsed, iteration, warmup, duration)?;
    }
    let timed_calls = match (entry.leg, ingress, lifecycle) {
        (Leg::Unified, Ingress::Eager, false) => {
            "fs::read + litchi::open_bytes(Vec<u8>) -> File::Odt"
        },
        (Leg::Unified, Ingress::Eager, true) => {
            "fs::read + litchi::open_bytes(Vec<u8>) + File::text()"
        },
        (Leg::Unified, Ingress::Source, false) => "litchi::open(Path) -> File::Odt (source-backed)",
        (Leg::Unified, Ingress::Source, true) => "litchi::open(Path) + File::text()",
        (Leg::Leaf, Ingress::Eager, false) => "fs::read + litchi_odt::Document::from_bytes",
        (Leg::Leaf, Ingress::Eager, true) => "fs::read + litchi_odt::Document::from_bytes + text()",
        (Leg::Leaf, Ingress::Source, false) => "litchi_odt::SourceBackedDocument::from_path",
        (Leg::Leaf, Ingress::Source, true) => {
            "litchi_odt::SourceBackedDocument::from_path + text()"
        },
    };
    let summary = RouteSummary::for_case(entry.case, timed_calls)?.verified(&expected, samples)?;
    Ok(finish(entry, corpus, elapsed, summary))
}

// ---------------------------------------------------------------------------
// ODP (odp_file_*)
// ---------------------------------------------------------------------------

#[allow(
    clippy::large_enum_variant,
    reason = "a stack value of the harness; boxing a leaf owner would add an allocation to the timed leaf leg"
)]
enum OdpRoot {
    Unified(litchi::File),
    Eager(litchi_odp::Presentation),
    Source(litchi_odp::SourceBackedPresentation),
}

impl OdpRoot {
    /// Open the root through the predecessor's ingress: owned bytes cloned
    /// outside the timer for `Eager`, the path for `Source`.
    fn open(
        leg: Leg,
        ingress: Ingress,
        owned: Option<Vec<u8>>,
        path: &Path,
    ) -> Result<Self, Box<dyn Error>> {
        Ok(match (leg, ingress) {
            (Leg::Unified, Ingress::Eager) => Self::Unified(litchi::open_bytes(
                owned.ok_or("eager ODP route has no owned bytes")?,
            )?),
            (Leg::Unified, Ingress::Source) => Self::Unified(litchi::open(path)?),
            (Leg::Leaf, Ingress::Eager) => Self::Eager(
                litchi_odp::Package::from_bytes(
                    &owned.ok_or("eager ODP route has no owned bytes")?,
                )?
                .presentation()?,
            ),
            (Leg::Leaf, Ingress::Source) => {
                Self::Source(litchi_odp::SourceBackedPresentation::open(path)?)
            },
        })
    }

    fn slide_count(&self) -> Result<usize, Box<dyn Error>> {
        Ok(match self {
            Self::Unified(file) => file.presentation()?.slide_count()?,
            Self::Eager(presentation) => presentation.slide_count()?,
            Self::Source(presentation) => presentation.slide_count()?,
        })
    }

    /// The timed selected-slide query; its text is read after the timer.
    fn slide(&self, index: usize) -> Result<OdpSlide, Box<dyn Error>> {
        Ok(match self {
            Self::Unified(file) => OdpSlide::Unified(
                file.presentation()?
                    .slide(index)?
                    .ok_or("unified ODP selected slide is missing")?,
            ),
            Self::Eager(presentation) => OdpSlide::Leaf(
                presentation
                    .slide(index)?
                    .ok_or("leaf ODP selected slide is missing")?,
            ),
            Self::Source(presentation) => OdpSlide::Leaf(
                presentation
                    .slide(index)?
                    .ok_or("leaf ODP selected slide is missing")?,
            ),
        })
    }
}

enum OdpSlide {
    Unified(litchi::compat::slide::Slide),
    Leaf(litchi_odp::Slide),
}

impl OdpSlide {
    fn text(&self) -> Result<String, Box<dyn Error>> {
        Ok(match self {
            Self::Unified(slide) => slide.text()?,
            Self::Leaf(slide) => slide.all_text(),
        })
    }
}

fn run_odp(
    entry: &Entry,
    corpus: &Corpus,
    path: &Path,
    warmup: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let Plan::Odp {
        ingress,
        selected_slide,
    } = entry.plan
    else {
        return Err(unexpected_plan(entry));
    };
    let slide_count = SemanticShape::Medium.pptx_slides();
    let target = slide_count / 2;
    let expected_text = format!(
        "{}\n{}",
        crate::semantic_odp_title(target, false),
        crate::semantic_odp_text(target, false)
    );
    if fs::read(path)? != corpus.archive {
        return Err("ODP temporary source differs from the corpus".into());
    }
    let expected = (slide_count, sha256_hex(expected_text.as_bytes()));

    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup, samples)? {
        let owned = (ingress == Ingress::Eager).then(|| corpus.archive.clone());
        let (duration, root, selected) = if selected_slide {
            // The root is prepared outside the timer, as in the predecessor.
            let root = OdpRoot::open(entry.leg, ingress, owned, path)?;
            let started = Instant::now();
            let slide = root.slide(target)?;
            (started.elapsed(), root, Some(slide))
        } else {
            let started = Instant::now();
            let root = OdpRoot::open(entry.leg, ingress, owned, path)?;
            (started.elapsed(), root, None)
        };
        let selected_text = match selected {
            Some(slide) => slide.text()?,
            None => root.slide(target)?.text()?,
        };
        let observed = (root.slide_count()?, sha256_hex(selected_text.as_bytes()));
        verify_projection(entry, &observed, &expected)?;
        std::hint::black_box(root);
        record_elapsed(&mut elapsed, iteration, warmup, duration)?;
    }
    let timed_calls = match (entry.leg, ingress, selected_slide) {
        (Leg::Unified, Ingress::Eager, false) => "litchi::open_bytes(Vec<u8>) -> File::Odp",
        (Leg::Unified, Ingress::Source, false) => "litchi::open(Path) -> File::Odp (source-backed)",
        (Leg::Unified, _, true) => {
            "File::presentation()?.slide(middle) on a File prepared outside the timer"
        },
        (Leg::Leaf, Ingress::Eager, false) => {
            "litchi_odp::Package::from_bytes(&Vec<u8>).map_err(litchi_core::Error::from).and_then(|package| package.presentation())"
        },
        (Leg::Leaf, Ingress::Source, false) => "litchi_odp::SourceBackedPresentation::from_path",
        (Leg::Leaf, Ingress::Eager, true) => {
            "litchi_odp::Presentation::slide(middle) on a presentation prepared outside the timer"
        },
        (Leg::Leaf, Ingress::Source, true) => {
            "litchi_odp::SourceBackedPresentation::slide(middle) on a presentation prepared outside the timer"
        },
    };
    let summary = RouteSummary::for_case(entry.case, timed_calls)?.verified(&expected, samples)?;
    Ok(finish(entry, corpus, elapsed, summary))
}

// ---------------------------------------------------------------------------
// ODS (ods_file_*_open)
// ---------------------------------------------------------------------------

fn run_ods(
    entry: &Entry,
    corpus: &Corpus,
    path: &Path,
    warmup: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let Plan::Ods { ingress } = entry.plan else {
        return Err(unexpected_plan(entry));
    };
    let typed = litchi_ods::Package::from_bytes(&corpus.archive)?.workbook()?;
    let expected = workbook_projection(
        typed.sheet_names(),
        typed.sheet_count(),
        &typed.text()?,
        crate::ods_root_metadata_digest(typed.metadata())?,
    );
    drop(typed);
    if fs::read(path)? != corpus.archive {
        return Err("ODS temporary source differs from the corpus".into());
    }

    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup, samples)? {
        let owned = (ingress == Ingress::Eager).then(|| corpus.archive.clone());
        let (duration, observed) = match (entry.leg, ingress) {
            (Leg::Unified, _) => {
                let started = Instant::now();
                let file = match owned {
                    Some(owned) => litchi::open_bytes(owned)?,
                    None => litchi::open(path)?,
                };
                let duration = started.elapsed();
                let (names, count, text) = unified_workbook_names_count_text(&file)?;
                let metadata = file.metadata()?.unwrap_or_default();
                let observed = workbook_projection(
                    names,
                    count,
                    &text,
                    crate::ods_root_metadata_digest(&metadata)?,
                );
                std::hint::black_box(file);
                (duration, observed)
            },
            (Leg::Leaf, Ingress::Eager) => {
                let owned = owned.ok_or("eager ODS route has no owned bytes")?;
                let started = Instant::now();
                let spreadsheet = litchi_ods::Package::from_bytes(&owned)?.workbook()?;
                let duration = started.elapsed();
                let observed = workbook_projection(
                    spreadsheet.sheet_names(),
                    spreadsheet.sheet_count(),
                    &spreadsheet.text()?,
                    crate::ods_root_metadata_digest(spreadsheet.metadata())?,
                );
                std::hint::black_box(spreadsheet);
                (duration, observed)
            },
            (Leg::Leaf, Ingress::Source) => {
                let started = Instant::now();
                let spreadsheet = litchi_ods::SourceBackedSpreadsheet::open(path)?;
                let duration = started.elapsed();
                let observed = workbook_projection(
                    spreadsheet.sheet_names()?,
                    spreadsheet.sheet_count()?,
                    &spreadsheet.text()?,
                    crate::ods_root_metadata_digest(spreadsheet.metadata()?)?,
                );
                std::hint::black_box(spreadsheet);
                (duration, observed)
            },
        };
        verify_projection(entry, &observed, &expected)?;
        record_elapsed(&mut elapsed, iteration, warmup, duration)?;
    }
    let timed_calls = match (entry.leg, ingress) {
        (Leg::Unified, Ingress::Eager) => "litchi::open_bytes(Vec<u8>) -> File::Ods",
        (Leg::Unified, Ingress::Source) => "litchi::open(Path) -> File::Ods (source-backed)",
        (Leg::Leaf, Ingress::Eager) => {
            "litchi_ods::Package::from_bytes(&Vec<u8>).map_err(litchi_core::Error::from).and_then(|package| package.workbook())"
        },
        (Leg::Leaf, Ingress::Source) => "litchi_ods::SourceBackedSpreadsheet::from_path",
    };
    let summary = RouteSummary::for_case(entry.case, timed_calls)?.verified(&expected, samples)?;
    Ok(finish(entry, corpus, elapsed, summary))
}

// ---------------------------------------------------------------------------
// Detection (detect_*_unified_*, detect_*_leaf_*)
// ---------------------------------------------------------------------------

fn run_detection(
    entry: &Entry,
    corpus: &Corpus,
    warmup: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let Plan::Detect { family, mode } = entry.plan else {
        return Err(unexpected_plan(entry));
    };
    let expected = mode.expected_format(family);
    if entry.leg == Leg::Leaf && !matches!(mode, DetectionMode::Bytes | DetectionMode::Reader) {
        return Err("only the ODF byte and reader inputs have a leaf detector".into());
    }
    let detection_limits = DetectSourceLimits::default();
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup, samples)? {
        let (observed, duration, cursor) = match mode {
            DetectionMode::Reader => {
                let original_position = 7_u64;
                let mut reader = Cursor::new(corpus.archive.as_slice());
                reader.set_position(original_position);
                let started = Instant::now();
                let observed = match entry.leg {
                    Leg::Unified => litchi_ops::detect::detect_format_from_reader(
                        &mut reader,
                        &detection_limits,
                    ),
                    Leg::Leaf => litchi_odf_common::detect::reader(
                        &mut reader,
                        litchi_odf_common::detect::Limits::default(),
                    )
                    .ok()
                    .flatten(),
                };
                let duration = started.elapsed();
                (
                    observed,
                    duration,
                    Some((reader.position(), original_position)),
                )
            },
            DetectionMode::Bytes | DetectionMode::Polyglot | DetectionMode::CatalogAlias => {
                let started = Instant::now();
                let observed = match entry.leg {
                    Leg::Unified => litchi_ops::detect::detect_file_format_from_bytes(
                        &corpus.archive,
                        &detection_limits,
                    ),
                    Leg::Leaf => litchi_odf_common::detect::bytes(
                        &corpus.archive,
                        litchi_odf_common::detect::Limits::default(),
                    )
                    .ok()
                    .flatten(),
                };
                (observed, started.elapsed(), None)
            },
        };
        if observed != Some(expected) {
            return Err(format!(
                "{} detected {observed:?}, the corpus expects {expected:?}",
                entry.case.name()
            )
            .into());
        }
        if let Some((observed_position, original_position)) = cursor
            && observed_position != original_position
        {
            return Err(format!("{} moved the reader cursor", entry.case.name()).into());
        }
        std::hint::black_box(observed);
        record_elapsed(&mut elapsed, iteration, warmup, duration)?;
    }
    let timed_calls = match (entry.leg, mode) {
        (Leg::Unified, DetectionMode::Reader) => {
            "litchi_ops::detect::detect_format_from_reader (the coordinator litchi::open runs, without the facade's iWork tail)"
        },
        (Leg::Unified, _) => {
            "litchi_ops::detect::detect_file_format_from_bytes (the coordinator litchi::open runs, without the facade's iWork tail)"
        },
        (Leg::Leaf, DetectionMode::Reader) => "litchi_odf_common::detect::reader",
        (Leg::Leaf, _) => "litchi_odf_common::detect::bytes",
    };
    let summary = RouteSummary::for_case(entry.case, timed_calls)?
        .verified(&format!("{expected:?}"), samples)?;
    Ok(finish(entry, corpus, elapsed, summary))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ppt_text_batch_routes_are_named_paired_and_opt_in() {
        // The `_aggregate_v3` identity is bound to this native operation
        // bound: it decides the expected publication count and the output
        // bytes (record 0272). A bound change must retire these names, as
        // record 0272 retired `_aggregate_v2` for the 4096 bound.
        assert_eq!(PPT_TEXT_BATCH_NATIVE_OPERATION_LIMIT, 16_384);
        assert_eq!(PPT_TEXT_BATCH_CATALOG.len(), 28);
        for entry in PPT_TEXT_BATCH_CATALOG {
            assert_eq!(crate::parse_case(entry.case.name()), Some(entry.case));
            assert_eq!(crate::parse_case(entry.twin.name()), Some(entry.twin));
            assert!(
                !Case::DEFAULT.contains(&entry.case),
                "{}",
                entry.case.name()
            );
            assert!(
                !Case::DEFAULT.contains(&entry.twin),
                "{}",
                entry.twin.name()
            );
            let twin = ppt_text_batch_entry(entry.twin).expect("text-batch twin is registered");
            assert_eq!(twin.twin, entry.case);
            assert_eq!(twin.owner, entry.owner);
            assert_eq!(twin.count, entry.count);
            assert_ne!(twin.leg, entry.leg);
            assert!(matches!(
                entry.count,
                1 | 10 | 100 | 256 | 1_000 | 4_096 | 10_000
            ));
            assert!(entry.case.name().ends_with("_aggregate_v3"));
            let retired = entry.case.name().replace("_aggregate_v3", "_v1");
            assert_eq!(crate::parse_case(&retired), None, "frozen v1 is historical");
            // Record 0272: the 0220 `_aggregate_v2` selectors qualified the
            // 4096-operation native bound; their frozen evidence keeps them.
            let retired = entry.case.name().replace("_aggregate_v3", "_aggregate_v2");
            assert_eq!(crate::parse_case(&retired), None, "frozen v2 is historical");
        }
        for invalid in [
            "ppt_text_batch_eager_native_0_aggregate_v3",
            "ppt_text_batch_eager_native_1",
            "ppt_text_batch_positional_unified_10001_aggregate_v3",
            "ppt_text_batch_eager_unified_10000_aggregate_v2",
            "ppt_text_batch_source_native_10_v1",
        ] {
            assert_eq!(crate::parse_case(invalid), None, "{invalid}");
        }
    }

    #[test]
    fn ppt_text_batch_qualification_covers_native_chunk_boundary_and_source_identity() {
        let corpus = build_ppt_text_batch_corpus(PPT_TEXT_BATCH_NATIVE_CHUNK_SIZE + 1)
            .expect("PPT text-batch fixture");
        let source = TempSource::new("ppt", &corpus.archive).expect("PPT text-batch source");
        let entry = PptTextBatchEntry {
            case: Case::PptTextAggregateEagerNative100,
            twin: Case::PptTextAggregateEagerUnified100,
            owner: PptTextBatchOwner::Eager,
            leg: PptTextBatchLeg::Native,
            count: PPT_TEXT_BATCH_NATIVE_CHUNK_SIZE + 1,
        };
        let native =
            open_ppt_text_batch_native(entry.owner, &corpus, source.path()).expect("native owner");
        let replacements = (0..entry.count)
            .map(|index| {
                let target =
                    litchi_ppt::text_edit::Target::new(Position::new(0), Position::new(index));
                let updated = ppt_text_batch_updated_text(entry.count, index);
                litchi_ppt::slide_order::ShapeTextReplacement::new(target, updated)
            })
            .collect::<Vec<_>>();
        let (native_output, native_stats) =
            apply_ppt_text_batch_native(&native, &replacements).expect("native qualification");
        assert_eq!(native_stats.inner_chunks, 2);
        assert_eq!(native_stats.aggregate_calls, 1);
        assert_eq!(native_stats.commits, 1);
        assert_eq!(native_stats.changed, PPT_TEXT_BATCH_NATIVE_CHUNK_SIZE + 1);
        let prepared = prepare_ppt_text_batch(entry, &corpus, source.path())
            .expect("coalesced native/unified qualification across chunk boundary");
        assert_eq!(prepared.unified_stats.commits, 1);
        assert_eq!(prepared.unified_stats.operations, entry.count);
        assert_ne!(
            native_output.to_bytes().expect("native output bytes"),
            corpus.archive
        );
        assert_eq!(
            fs::read(source.path()).expect("source bytes"),
            corpus.archive
        );
        let small_corpus = build_ppt_text_batch_corpus(1).expect("small PPT text-batch fixture");
        let small_source =
            TempSource::new("ppt", &small_corpus.archive).expect("small PPT text-batch source");
        let small_entry =
            ppt_text_batch_entry(Case::PptTextAggregateEagerNative1).expect("small native route");
        let prepared = prepare_ppt_text_batch(small_entry, &small_corpus, small_source.path())
            .expect("native/unified qualification");
        assert_eq!(prepared.unified_stats.operations, 1);
        assert_eq!(
            prepared.limits.max_snapshot_bytes(),
            SemanticLimits::default().max_snapshot_bytes()
        );
        let result = run_ppt_text_batch(
            small_entry,
            &small_corpus,
            small_source.path(),
            &prepared,
            0,
            1,
        )
        .expect("qualified native route");
        let summary = result
            .source
            .expect("source envelope")
            .unified_ops
            .expect("text-batch summary");
        let evidence = summary.ppt_text_batch.expect("text-batch evidence");
        assert!(evidence.exact_output_verified);
        assert!(evidence.unchanged_source_verified);
        assert_eq!(result.elapsed_ns.samples.len(), 1);
        small_source
            .remove()
            .expect("remove small PPT text-batch source");
        source.remove().expect("remove PPT text-batch source");
    }

    #[test]
    fn ppt_text_batch_snapshot_policy_expands_only_past_native_rollover() {
        assert_eq!(
            ppt_text_batch_limits(1_000)
                .expect("default profile")
                .max_snapshot_bytes(),
            SemanticLimits::default().max_snapshot_bytes()
        );
        assert_eq!(
            ppt_text_batch_limits(PPT_TEXT_BATCH_NATIVE_OPERATION_LIMIT)
                .expect("default profile at the native bound")
                .max_snapshot_bytes(),
            SemanticLimits::default().max_snapshot_bytes()
        );
        assert_eq!(
            ppt_text_batch_limits(PPT_TEXT_BATCH_NATIVE_OPERATION_LIMIT + 1)
                .expect("large profile")
                .max_snapshot_bytes(),
            PPT_TEXT_BATCH_LARGE_SNAPSHOT_LIMIT
        );
    }

    fn fixture(relative: &str) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(relative)
    }

    #[test]
    fn catalog_pairs_every_route_with_its_twin_and_predecessor() {
        let mut names = std::collections::BTreeSet::new();
        let mut kept = std::collections::BTreeSet::new();
        let mut deleted = std::collections::BTreeSet::new();
        for entry in CATALOG {
            let name = entry.case.name();
            assert!(names.insert(name), "{name} is registered twice");
            assert_eq!(crate::parse_case(name), Some(entry.case));
            assert!(!Case::DEFAULT.contains(&entry.case), "{name}");
            let infix = match entry.leg {
                Leg::Unified => "_unified_",
                Leg::Leaf => "_leaf_",
            };
            assert!(name.contains(infix), "{name}");
            // The name is the predecessor's with the leg infix inserted.
            let predecessor = entry.predecessor;
            let stripped = name.replacen(infix, "_", 1);
            assert!(
                stripped == predecessor
                    || stripped.replacen("_file_", "_facade_file_", 1) == predecessor,
                "{name} does not re-measure {predecessor}"
            );
            // M0a step S9 deleted the predecessors that timed the facade
            // types; their frozen records keep their names, which S10
            // compares against. The others stay registered and are not
            // themselves paired routes.
            match crate::parse_case(predecessor) {
                Some(case) => {
                    assert!(super::entry(case).is_none(), "{predecessor}");
                    kept.insert(predecessor);
                },
                None => {
                    deleted.insert(predecessor);
                },
            }
            match entry.twin {
                Some(twin) => {
                    let twin_entry = super::entry(twin).expect("twin is registered");
                    assert_eq!(twin_entry.twin, Some(entry.case));
                    assert_eq!(twin_entry.predecessor, entry.predecessor);
                    assert_ne!(twin_entry.leg, entry.leg);
                },
                None => {
                    assert_eq!(entry.leg, Leg::Unified);
                    assert!(matches!(
                        entry.plan,
                        Plan::Detect {
                            mode: DetectionMode::Polyglot | DetectionMode::CatalogAlias,
                            ..
                        }
                    ));
                },
            }
            assert_eq!(
                matches!(entry.plan, Plan::Filesystem),
                entry.case.is_unified_ops_filesystem(),
                "{name}"
            );
            assert_eq!(
                !matches!(entry.plan, Plan::Filesystem),
                entry.case.is_unified_ops_in_process(),
                "{name}"
            );
        }
        assert_eq!(CATALOG.len(), 118);
        // 47 predecessors timed the deleted facade types; the 15 kept ones
        // time the detection functions and the DOCX leaf owner.
        assert_eq!(deleted.len(), 47, "{deleted:?}");
        assert_eq!(kept.len(), 15, "{kept:?}");
        assert!(kept.iter().all(|name| {
            name.starts_with("detect_")
                || name.starts_with("docx_semantic_")
                || name.starts_with("docx_producer_")
        }));
    }

    #[test]
    fn xlsx_bytes_pair_reproduces_the_typed_oracle() {
        let corpus = crate::build_xlsx_cell_crud_corpus(crate::XlsxCellCrudShape::Medium).unwrap();
        for case in [
            Case::XlsxBytesUnifiedOpen,
            Case::XlsxBytesUnifiedOpenLifecycle,
            Case::XlsxBytesLeafOpen,
            Case::XlsxBytesLeafOpenLifecycle,
        ] {
            let entry = entry(case).unwrap();
            let result = run_xlsx_bytes(entry, &corpus, 0, 1).unwrap();
            assert_eq!(result.case, case.name());
            assert_eq!(result.elapsed_ns.samples.len(), 1);
            let summary = result.source.unwrap().unified_ops.unwrap();
            assert_eq!(summary.verified_samples, Some(1));
            assert_eq!(summary.predecessor, entry.predecessor);
        }
    }

    #[test]
    fn docx_read_one_pairs_return_the_generated_paragraph() {
        let corpus = crate::build_semantic_docx_corpus(SemanticShape::Tiny).unwrap();
        for case in [
            Case::DocxSemanticUnifiedOneParagraph,
            Case::DocxSemanticUnifiedOneParagraphText,
            Case::DocxSemanticLeafOneParagraph,
            Case::DocxSemanticLeafOneParagraphText,
        ] {
            let result = run_docx_semantic(entry(case).unwrap(), &corpus, 0, 1).unwrap();
            assert_eq!(result.elapsed_ns.samples.len(), 1);
        }
        let producer = crate::producer_shape::build_docx_producer_corpus().unwrap();
        for case in [
            Case::DocxProducerUnifiedSourceSelectedParagraph,
            Case::DocxProducerLeafSourceSelectedParagraph,
        ] {
            let result = run_docx_producer(entry(case).unwrap(), &producer, 0, 1).unwrap();
            assert_eq!(result.elapsed_ns.samples.len(), 1);
        }
    }

    #[test]
    fn rtf_and_detection_pairs_agree_with_their_oracles() {
        let corpus =
            crate::build_semantic_rtf_corpus(SemanticShape::Tiny, RtfSemanticVariant::Plain)
                .unwrap();
        for case in [
            Case::RtfFileUnifiedOpen,
            Case::RtfFileUnifiedOpenLifecycle,
            Case::RtfFileLeafOpen,
            Case::RtfFileLeafOpenLifecycle,
        ] {
            let result = run_rtf(entry(case).unwrap(), &corpus, 0, 1).unwrap();
            assert_eq!(result.elapsed_ns.samples.len(), 1);
        }
        for entry in CATALOG
            .iter()
            .filter(|entry| matches!(entry.plan, Plan::Detect { .. }))
        {
            let predecessor = crate::parse_case(entry.predecessor).unwrap();
            let corpus = crate::build_detection_corpus(predecessor, SemanticShape::Tiny).unwrap();
            let result = run_detection(entry, &corpus, 0, 1).unwrap();
            assert_eq!(result.case, entry.case.name());
        }
    }

    #[test]
    fn odf_pairs_open_and_query_the_fixed_corpora() {
        let selection_cases = CATALOG
            .iter()
            .filter(|entry| {
                matches!(
                    entry.plan,
                    Plan::Odt { .. } | Plan::Odp { .. } | Plan::Ods { .. }
                )
            })
            .map(|entry| entry.case)
            .collect::<Vec<_>>();
        let results = run_selected(&Selection {
            cases: &selection_cases,
            semantic_shapes: &[SemanticShape::Tiny],
            rtf_variants: &[RtfSemanticVariant::Plain],
            ole2_files: &[],
            warmup_iterations: 0,
            samples: 1,
        })
        .unwrap();
        assert_eq!(results.len(), selection_cases.len());
    }

    #[test]
    fn ole2_pairs_reach_the_same_outcome_on_real_fixtures() {
        for (path, cases) in [
            (
                "test-data/ole/doc/documentProperties.doc",
                [
                    Case::DocFileUnifiedOpen,
                    Case::DocFileUnifiedFullText,
                    Case::DocFileUnifiedOneParagraph,
                    Case::DocFileLeafOpen,
                    Case::DocFileLeafFullText,
                    Case::DocFileLeafOneParagraph,
                ],
            ),
            (
                "test-data/ole/ppt/SampleShow.ppt",
                [
                    Case::PptFileUnifiedOpen,
                    Case::PptFileUnifiedFullText,
                    Case::PptFileUnifiedOneSlideText,
                    Case::PptFileLeafOpen,
                    Case::PptFileLeafFullText,
                    Case::PptFileLeafOneSlideText,
                ],
            ),
        ] {
            let results = run_selected(&Selection {
                cases: &cases,
                semantic_shapes: &[SemanticShape::Tiny],
                rtf_variants: &[RtfSemanticVariant::Plain],
                ole2_files: &[fixture(path)],
                warmup_iterations: 0,
                samples: 1,
            })
            .unwrap();
            assert_eq!(results.len(), 6);
            for result in results {
                let summary = result.source.unwrap().unified_ops.unwrap();
                assert_eq!(summary.legs_agree, Some(true), "{}", result.case);
                assert!(
                    summary
                        .observation
                        .as_deref()
                        .is_some_and(|observation| !observation.starts_with("refused:")),
                    "{}",
                    result.case
                );
            }
        }
    }
}
