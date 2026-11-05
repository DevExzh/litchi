//! Paired unified-API and leaf-API filesystem routes (M0a step S8b).
//!
//! Plan 0003 §5.3 lists the open, list, read-one and full-text scenarios M0a
//! re-measures. Every DOCX, PPTX and XLSX scenario this module's parent runs
//! in fresh child processes gets two new selectors:
//!
//! - a **unified** leg that goes through `litchi::open`,
//!   `litchi::open_bytes(_with)`, `litchi::File` and its family views and
//!   format handles, and
//! - a **leaf** leg that makes the same leaf calls directly on the format's
//!   leaf crate. The format is known, so the leaf leg detects nothing.
//!
//! Each leg mirrors its predecessor selector: the same corpus, the same
//! ingress (`fs::read` plus a byte open, or a path open), the same untimed
//! preparation for query routes, and the same timer placement, including
//! where the opened root is dropped (inside the timed scope for DOCX and PPTX,
//! after every operation-only snapshot for XLSX). The predecessors stay
//! registered until M0a step S9 deletes the types they call; their frozen
//! records are never rewritten.
//!
//! Correctness is checked outside the timer exactly as for the predecessors:
//! the unified projection, the leaf projection and the corpus generator must
//! agree, and source-backed routes carry the same untimed leaf replay
//! evidence. Nothing here makes a performance claim.

use std::{error::Error, fs, path::Path, sync::Arc};

use litchi_core::{FileSource, Metadata, OwnedSource};
use sha2::{Digest, Sha256};

use super::{
    DOCX_FILE_CORPUS_GENERATOR, DOCX_PARAGRAPH_SEPARATOR, PPTX_FILE_CORPUS_GENERATOR,
    PPTX_FILE_SELECTED_POSITION, SourceReplay, XLSX_FILE_SELECTED_CELL_TARGET,
    XlsxRootSemanticProjection, XlsxSelectedCellEvidence, assert_pinned_xlsx_corpus,
    assert_source_sha256, docx_archive_signature, docx_package_signature, mixed_ascii_case,
    verify_docx_timed_full_text, xlsx_selected_cell_evidence_digest,
    xlsx_semantic_projection_sha256, xlsx_typed_semantic_oracle,
};
use crate::unified_ops::{
    LeafXlsx, RouteSummary, eager_options, leaf_docx_eager, leaf_xlsx_names, leaf_xlsx_open,
    leaf_xlsx_text,
};
use crate::{Case, PPTX_SOURCE_SLIDE_COUNT, SemanticShape};

/// Which side of a pair a route measures.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Leg {
    /// `litchi::open*`, `litchi::File`, its views and handles.
    Unified,
    /// The format's leaf crate, called directly.
    Leaf,
}

/// The ingress and owner state a DOCX or PPTX route mirrors.
///
/// The names follow the predecessor selectors. For DOCX, `Eager` is the
/// owning `litchi_docx::Package` built from `fs::read` bytes. For PPTX,
/// `Eager` is the byte route (`fs::read`, then a source-backed presentation
/// over the owned bytes), which is what `pptx_file_eager_*` measured.
/// `Source` is the path route on both.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Ingress {
    Eager,
    Source,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum DocxQuery {
    Open,
    ParagraphCount,
    ListParagraphs,
    FullText,
    OpenParagraphCountLifecycle,
    OpenFullTextLifecycle,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum PptxQuery {
    Open,
    ListSlides,
    SlideCount,
    SelectedSlide,
    OpenSlideCountLifecycle,
    OpenSelectedSlideLifecycle,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum XlsxQuery {
    Open,
    OpenLifecycle,
    SelectedCell,
}

/// One paired filesystem route.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum Route {
    Docx {
        leg: Leg,
        ingress: Ingress,
        query: DocxQuery,
    },
    Pptx {
        leg: Leg,
        ingress: Ingress,
        query: PptxQuery,
    },
    /// The XLSX path route (`xlsx_file_*`), source-backed on both legs.
    Xlsx { leg: Leg, query: XlsxQuery },
}

/// Every paired filesystem route and its selector, in `Case::name` order.
const ROUTES: [(Route, Case); 54] = [
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: DocxQuery::Open,
        },
        Case::DocxFileUnifiedEagerOpen,
    ),
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: DocxQuery::Open,
        },
        Case::DocxFileUnifiedSourceOpen,
    ),
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: DocxQuery::ParagraphCount,
        },
        Case::DocxFileUnifiedEagerParagraphCount,
    ),
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: DocxQuery::ParagraphCount,
        },
        Case::DocxFileUnifiedSourceParagraphCount,
    ),
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: DocxQuery::ListParagraphs,
        },
        Case::DocxFileUnifiedEagerListParagraphs,
    ),
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: DocxQuery::ListParagraphs,
        },
        Case::DocxFileUnifiedSourceListParagraphs,
    ),
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: DocxQuery::FullText,
        },
        Case::DocxFileUnifiedEagerFullText,
    ),
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: DocxQuery::FullText,
        },
        Case::DocxFileUnifiedSourceFullText,
    ),
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: DocxQuery::OpenParagraphCountLifecycle,
        },
        Case::DocxFileUnifiedEagerOpenParagraphCountLifecycle,
    ),
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: DocxQuery::OpenParagraphCountLifecycle,
        },
        Case::DocxFileUnifiedSourceOpenParagraphCountLifecycle,
    ),
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: DocxQuery::OpenFullTextLifecycle,
        },
        Case::DocxFileUnifiedEagerOpenFullTextLifecycle,
    ),
    (
        Route::Docx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: DocxQuery::OpenFullTextLifecycle,
        },
        Case::DocxFileUnifiedSourceOpenFullTextLifecycle,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: DocxQuery::Open,
        },
        Case::DocxFileLeafEagerOpen,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: DocxQuery::Open,
        },
        Case::DocxFileLeafSourceOpen,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: DocxQuery::ParagraphCount,
        },
        Case::DocxFileLeafEagerParagraphCount,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: DocxQuery::ParagraphCount,
        },
        Case::DocxFileLeafSourceParagraphCount,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: DocxQuery::ListParagraphs,
        },
        Case::DocxFileLeafEagerListParagraphs,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: DocxQuery::ListParagraphs,
        },
        Case::DocxFileLeafSourceListParagraphs,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: DocxQuery::FullText,
        },
        Case::DocxFileLeafEagerFullText,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: DocxQuery::FullText,
        },
        Case::DocxFileLeafSourceFullText,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: DocxQuery::OpenParagraphCountLifecycle,
        },
        Case::DocxFileLeafEagerOpenParagraphCountLifecycle,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: DocxQuery::OpenParagraphCountLifecycle,
        },
        Case::DocxFileLeafSourceOpenParagraphCountLifecycle,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: DocxQuery::OpenFullTextLifecycle,
        },
        Case::DocxFileLeafEagerOpenFullTextLifecycle,
    ),
    (
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: DocxQuery::OpenFullTextLifecycle,
        },
        Case::DocxFileLeafSourceOpenFullTextLifecycle,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: PptxQuery::Open,
        },
        Case::PptxFileUnifiedEagerOpen,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: PptxQuery::Open,
        },
        Case::PptxFileUnifiedSourceOpen,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: PptxQuery::ListSlides,
        },
        Case::PptxFileUnifiedEagerListSlides,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: PptxQuery::ListSlides,
        },
        Case::PptxFileUnifiedSourceListSlides,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: PptxQuery::SlideCount,
        },
        Case::PptxFileUnifiedEagerSlideCount,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: PptxQuery::SlideCount,
        },
        Case::PptxFileUnifiedSourceSlideCount,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: PptxQuery::SelectedSlide,
        },
        Case::PptxFileUnifiedEagerSelectedSlide,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: PptxQuery::SelectedSlide,
        },
        Case::PptxFileUnifiedSourceSelectedSlide,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: PptxQuery::OpenSlideCountLifecycle,
        },
        Case::PptxFileUnifiedEagerOpenSlideCountLifecycle,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: PptxQuery::OpenSlideCountLifecycle,
        },
        Case::PptxFileUnifiedSourceOpenSlideCountLifecycle,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Eager,
            query: PptxQuery::OpenSelectedSlideLifecycle,
        },
        Case::PptxFileUnifiedEagerOpenSelectedSlideLifecycle,
    ),
    (
        Route::Pptx {
            leg: Leg::Unified,
            ingress: Ingress::Source,
            query: PptxQuery::OpenSelectedSlideLifecycle,
        },
        Case::PptxFileUnifiedSourceOpenSelectedSlideLifecycle,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: PptxQuery::Open,
        },
        Case::PptxFileLeafEagerOpen,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: PptxQuery::Open,
        },
        Case::PptxFileLeafSourceOpen,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: PptxQuery::ListSlides,
        },
        Case::PptxFileLeafEagerListSlides,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: PptxQuery::ListSlides,
        },
        Case::PptxFileLeafSourceListSlides,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: PptxQuery::SlideCount,
        },
        Case::PptxFileLeafEagerSlideCount,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: PptxQuery::SlideCount,
        },
        Case::PptxFileLeafSourceSlideCount,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: PptxQuery::SelectedSlide,
        },
        Case::PptxFileLeafEagerSelectedSlide,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: PptxQuery::SelectedSlide,
        },
        Case::PptxFileLeafSourceSelectedSlide,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: PptxQuery::OpenSlideCountLifecycle,
        },
        Case::PptxFileLeafEagerOpenSlideCountLifecycle,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: PptxQuery::OpenSlideCountLifecycle,
        },
        Case::PptxFileLeafSourceOpenSlideCountLifecycle,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            query: PptxQuery::OpenSelectedSlideLifecycle,
        },
        Case::PptxFileLeafEagerOpenSelectedSlideLifecycle,
    ),
    (
        Route::Pptx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            query: PptxQuery::OpenSelectedSlideLifecycle,
        },
        Case::PptxFileLeafSourceOpenSelectedSlideLifecycle,
    ),
    (
        Route::Xlsx {
            leg: Leg::Unified,
            query: XlsxQuery::Open,
        },
        Case::XlsxFileUnifiedOpen,
    ),
    (
        Route::Xlsx {
            leg: Leg::Unified,
            query: XlsxQuery::OpenLifecycle,
        },
        Case::XlsxFileUnifiedOpenLifecycle,
    ),
    (
        Route::Xlsx {
            leg: Leg::Unified,
            query: XlsxQuery::SelectedCell,
        },
        Case::XlsxFileUnifiedSelectedCell,
    ),
    (
        Route::Xlsx {
            leg: Leg::Leaf,
            query: XlsxQuery::Open,
        },
        Case::XlsxFileLeafOpen,
    ),
    (
        Route::Xlsx {
            leg: Leg::Leaf,
            query: XlsxQuery::OpenLifecycle,
        },
        Case::XlsxFileLeafOpenLifecycle,
    ),
    (
        Route::Xlsx {
            leg: Leg::Leaf,
            query: XlsxQuery::SelectedCell,
        },
        Case::XlsxFileLeafSelectedCell,
    ),
];

impl Route {
    /// The route registered under `name`, if any.
    pub(super) fn parse(name: &str) -> Option<Self> {
        ROUTES
            .iter()
            .find(|(_, case)| case.name() == name)
            .map(|(route, _)| *route)
    }

    /// The selector of this route.
    pub(super) const fn case(self) -> Case {
        let mut index = 0;
        while index < ROUTES.len() {
            if ROUTES[index].0.same(self) {
                return ROUTES[index].1;
            }
            index += 1;
        }
        panic!("every filesystem route is registered")
    }

    /// `PartialEq` for the `const` lookup above.
    const fn same(self, other: Self) -> bool {
        match (self, other) {
            (
                Self::Docx {
                    leg,
                    ingress,
                    query,
                },
                Self::Docx {
                    leg: other_leg,
                    ingress: other_ingress,
                    query: other_query,
                },
            ) => {
                leg as u8 == other_leg as u8
                    && ingress as u8 == other_ingress as u8
                    && query as u8 == other_query as u8
            },
            (
                Self::Pptx {
                    leg,
                    ingress,
                    query,
                },
                Self::Pptx {
                    leg: other_leg,
                    ingress: other_ingress,
                    query: other_query,
                },
            ) => {
                leg as u8 == other_leg as u8
                    && ingress as u8 == other_ingress as u8
                    && query as u8 == other_query as u8
            },
            (
                Self::Xlsx { leg, query },
                Self::Xlsx {
                    leg: other_leg,
                    query: other_query,
                },
            ) => leg as u8 == other_leg as u8 && query as u8 == other_query as u8,
            _ => false,
        }
    }

    #[cfg(test)]
    const fn leg(self) -> Leg {
        match self {
            Self::Docx { leg, .. } | Self::Pptx { leg, .. } | Self::Xlsx { leg, .. } => leg,
        }
    }

    pub(super) const fn is_docx(self) -> bool {
        matches!(self, Self::Docx { .. })
    }

    pub(super) const fn is_pptx(self) -> bool {
        matches!(self, Self::Pptx { .. })
    }

    pub(super) const fn is_xlsx(self) -> bool {
        matches!(self, Self::Xlsx { .. })
    }

    /// Query routes prepare their root before the timer, as their
    /// predecessors do, so a page-cache proof would not describe the timed
    /// operation.
    pub(super) const fn is_prepared_query(self) -> bool {
        matches!(
            self,
            Self::Docx {
                query: DocxQuery::ParagraphCount | DocxQuery::ListParagraphs | DocxQuery::FullText,
                ..
            } | Self::Pptx {
                query: PptxQuery::ListSlides | PptxQuery::SlideCount | PptxQuery::SelectedSlide,
                ..
            } | Self::Xlsx {
                query: XlsxQuery::SelectedCell,
                ..
            }
        )
    }

    pub(super) const fn is_docx_full_text_lifecycle(self) -> bool {
        matches!(
            self,
            Self::Docx {
                query: DocxQuery::OpenFullTextLifecycle,
                ..
            }
        )
    }

    pub(super) const fn is_xlsx_selected_cell(self) -> bool {
        matches!(
            self,
            Self::Xlsx {
                query: XlsxQuery::SelectedCell,
                ..
            }
        )
    }

    /// The predecessor query whose untimed leaf source replay also describes
    /// this route: source-backed DOCX and PPTX routes on both legs.
    /// The replay opens the leaf owner over the same file and runs the same
    /// query, so its evidence has the predecessor's shape; it does not
    /// replay the unified route's format-detection reads.
    pub(super) const fn source_replay(self) -> Option<SourceReplay> {
        match self {
            Self::Docx {
                ingress: Ingress::Source,
                query,
                ..
            } => Some(match query {
                DocxQuery::Open => SourceReplay::DocxSourceOpen,
                DocxQuery::ParagraphCount => SourceReplay::DocxSourceParagraphCount,
                DocxQuery::ListParagraphs => SourceReplay::DocxSourceListParagraphs,
                DocxQuery::FullText => SourceReplay::DocxSourceFullText,
                DocxQuery::OpenParagraphCountLifecycle => {
                    SourceReplay::DocxSourceOpenParagraphCountLifecycle
                },
                DocxQuery::OpenFullTextLifecycle => SourceReplay::DocxSourceOpenFullTextLifecycle,
            }),
            Self::Pptx {
                ingress: Ingress::Source,
                query,
                ..
            } => Some(match query {
                PptxQuery::Open => SourceReplay::PptxSourceOpen,
                PptxQuery::ListSlides => SourceReplay::PptxSourceListSlides,
                PptxQuery::SlideCount => SourceReplay::PptxSourceSlideCount,
                PptxQuery::SelectedSlide => SourceReplay::PptxSourceSelectedSlide,
                PptxQuery::OpenSlideCountLifecycle => {
                    SourceReplay::PptxSourceOpenSlideCountLifecycle
                },
                PptxQuery::OpenSelectedSlideLifecycle => {
                    SourceReplay::PptxSourceOpenSelectedSlideLifecycle
                },
            }),
            _ => None,
        }
    }

    /// The generic filesystem counter scope, as the predecessor reports it.
    pub(super) const fn logical_read_counter_scope(self) -> &'static str {
        match self {
            Self::Docx {
                ingress: Ingress::Source,
                ..
            }
            | Self::Pptx {
                ingress: Ingress::Source,
                ..
            } => "untimed_source_replay_only",
            Self::Docx { .. } => "not_applicable_eager_docx",
            Self::Pptx { .. } => "not_applicable_eager_pptx",
            Self::Xlsx { .. } => "not_applicable_filesystem_xlsx",
        }
    }

    /// The public calls inside the timed scope.
    const fn timed_calls(self) -> &'static str {
        match self {
            Self::Docx {
                leg: Leg::Unified,
                ingress: Ingress::Eager,
                query,
            } => match query {
                DocxQuery::Open => {
                    "fs::read + litchi::open_bytes_with(SourceMode::Eager) -> File::Docx; the handle drops inside the timer"
                },
                DocxQuery::ParagraphCount => {
                    "File::document()?.paragraph_count() on a File prepared outside the timer"
                },
                DocxQuery::ListParagraphs => {
                    "File::document()?.paragraphs() on a File prepared outside the timer"
                },
                DocxQuery::FullText => "File::text() on a File prepared outside the timer",
                DocxQuery::OpenParagraphCountLifecycle => {
                    "fs::read + litchi::open_bytes_with(SourceMode::Eager) + File::document()?.paragraph_count()"
                },
                DocxQuery::OpenFullTextLifecycle => {
                    "fs::read + litchi::open_bytes_with(SourceMode::Eager) + File::text()"
                },
            },
            Self::Docx {
                leg: Leg::Unified,
                ingress: Ingress::Source,
                query,
            } => match query {
                DocxQuery::Open => {
                    "litchi::open(Path) -> File::Docx (source-backed); the handle drops inside the timer"
                },
                DocxQuery::ParagraphCount => {
                    "File::document()?.paragraph_count() on a File prepared outside the timer"
                },
                DocxQuery::ListParagraphs => {
                    "File::document()?.paragraphs() on a File prepared outside the timer"
                },
                DocxQuery::FullText => "File::text() on a File prepared outside the timer",
                DocxQuery::OpenParagraphCountLifecycle => {
                    "litchi::open(Path) + File::document()?.paragraph_count()"
                },
                DocxQuery::OpenFullTextLifecycle => "litchi::open(Path) + File::text()",
            },
            Self::Docx {
                leg: Leg::Leaf,
                ingress: Ingress::Eager,
                query,
            } => match query {
                DocxQuery::Open => {
                    "fs::read + OpcPackage::from_bytes + litchi_docx::Package::from_opc_package + Package::document()?.text() validation pass, as the unified eager owner runs it"
                },
                DocxQuery::ParagraphCount => {
                    "Package::document()?.paragraph_count() on a package prepared outside the timer"
                },
                DocxQuery::ListParagraphs => {
                    "Package::document()?.paragraphs() on a package prepared outside the timer"
                },
                DocxQuery::FullText => {
                    "Package::document()?.text_with_paragraph_separator(\"\\n\") on a package prepared outside the timer"
                },
                DocxQuery::OpenParagraphCountLifecycle => {
                    "leaf eager open (as for open) + Package::document()?.paragraph_count()"
                },
                DocxQuery::OpenFullTextLifecycle => {
                    "leaf eager open (as for open) + Package::document()?.text_with_paragraph_separator(\"\\n\")"
                },
            },
            Self::Docx {
                leg: Leg::Leaf,
                ingress: Ingress::Source,
                query,
            } => match query {
                DocxQuery::Open => "litchi_docx::source_backed::Package::open",
                DocxQuery::ParagraphCount => {
                    "source_backed::Package::document()?.paragraph_count() on a package prepared outside the timer"
                },
                DocxQuery::ListParagraphs => {
                    "source_backed::Package::document()?.paragraphs() on a package prepared outside the timer"
                },
                DocxQuery::FullText => {
                    "source_backed::Package::document()?.extract_text_with_paragraph_separator(\"\\n\") on a package prepared outside the timer"
                },
                DocxQuery::OpenParagraphCountLifecycle => {
                    "source_backed::Package::open + document()?.paragraph_count()"
                },
                DocxQuery::OpenFullTextLifecycle => {
                    "source_backed::Package::open + document()?.extract_text_with_paragraph_separator(\"\\n\")"
                },
            },
            Self::Pptx {
                leg: Leg::Unified,
                ingress: Ingress::Eager,
                query,
            } => match query {
                PptxQuery::Open => {
                    "fs::read + litchi::open_bytes(Vec<u8>) -> File::Pptx (source-backed over the owned bytes); the handle drops inside the timer"
                },
                PptxQuery::ListSlides => {
                    "File::presentation()?.slides() on a File prepared outside the timer"
                },
                PptxQuery::SlideCount => {
                    "File::presentation()?.slide_count() on a File prepared outside the timer"
                },
                PptxQuery::SelectedSlide => {
                    "File::presentation()?.slide(100) on a File prepared outside the timer"
                },
                PptxQuery::OpenSlideCountLifecycle => {
                    "fs::read + litchi::open_bytes + File::presentation()?.slide_count()"
                },
                PptxQuery::OpenSelectedSlideLifecycle => {
                    "fs::read + litchi::open_bytes + File::presentation()?.slide(100)"
                },
            },
            Self::Pptx {
                leg: Leg::Unified,
                ingress: Ingress::Source,
                query,
            } => match query {
                PptxQuery::Open => {
                    "litchi::open(Path) -> File::Pptx (source-backed); the handle drops inside the timer"
                },
                PptxQuery::ListSlides => {
                    "File::presentation()?.slides() on a File prepared outside the timer"
                },
                PptxQuery::SlideCount => {
                    "File::presentation()?.slide_count() on a File prepared outside the timer"
                },
                PptxQuery::SelectedSlide => {
                    "File::presentation()?.slide(100) on a File prepared outside the timer"
                },
                PptxQuery::OpenSlideCountLifecycle => {
                    "litchi::open(Path) + File::presentation()?.slide_count()"
                },
                PptxQuery::OpenSelectedSlideLifecycle => {
                    "litchi::open(Path) + File::presentation()?.slide(100)"
                },
            },
            Self::Pptx {
                leg: Leg::Leaf,
                ingress: Ingress::Eager,
                query,
            } => match query {
                PptxQuery::Open => {
                    "fs::read + litchi_pptx::SourceBackedPresentation::from_read_at(OwnedSource)"
                },
                PptxQuery::ListSlides => {
                    "SourceBackedPresentation::slides() + SourceSlide::text_and_name() per slide on a presentation prepared outside the timer"
                },
                PptxQuery::SlideCount => {
                    "SourceBackedPresentation::slide_count() on a presentation prepared outside the timer"
                },
                PptxQuery::SelectedSlide => {
                    "SourceBackedPresentation::slide(100) + text_and_name() on a presentation prepared outside the timer"
                },
                PptxQuery::OpenSlideCountLifecycle => {
                    "fs::read + SourceBackedPresentation::from_read_at(OwnedSource) + slide_count()"
                },
                PptxQuery::OpenSelectedSlideLifecycle => {
                    "fs::read + SourceBackedPresentation::from_read_at(OwnedSource) + slide(100) + text_and_name()"
                },
            },
            Self::Pptx {
                leg: Leg::Leaf,
                ingress: Ingress::Source,
                query,
            } => match query {
                PptxQuery::Open => "litchi_pptx::SourceBackedPresentation::from_path",
                PptxQuery::ListSlides => {
                    "SourceBackedPresentation::slides() + SourceSlide::text_and_name() per slide on a presentation prepared outside the timer"
                },
                PptxQuery::SlideCount => {
                    "SourceBackedPresentation::slide_count() on a presentation prepared outside the timer"
                },
                PptxQuery::SelectedSlide => {
                    "SourceBackedPresentation::slide(100) + text_and_name() on a presentation prepared outside the timer"
                },
                PptxQuery::OpenSlideCountLifecycle => {
                    "SourceBackedPresentation::from_path + slide_count()"
                },
                PptxQuery::OpenSelectedSlideLifecycle => {
                    "SourceBackedPresentation::from_path + slide(100) + text_and_name()"
                },
            },
            Self::Xlsx {
                leg: Leg::Unified,
                query,
            } => match query {
                XlsxQuery::Open => {
                    "litchi::open(Path) -> File::Xlsx (source-backed); the handle stays alive until every operation-only snapshot is taken"
                },
                XlsxQuery::OpenLifecycle => {
                    "litchi::open(Path) + File::workbook()?.worksheet_names() + worksheet_count() + File::text()"
                },
                XlsxQuery::SelectedCell => {
                    "Xlsx::sheet(mixed-case name) + SelectedWorksheet::cell(A1) on a File prepared outside the timer"
                },
            },
            Self::Xlsx {
                leg: Leg::Leaf,
                query,
            } => match query {
                XlsxQuery::Open => {
                    "SourceBackedPackage::from_read_at(FileSource,
            litchi_opc::SourceLimits::default(),
            litchi_opc::SourceCacheLimits::default(),
            litchi_opc::SourceReadPolicy::exact(),
            None) + core-properties read + litchi_xlsx::SourceBackedWorkbook::from_source_backed_package, as the unified open runs them"
                },
                XlsxQuery::OpenLifecycle => {
                    "leaf open (as for open) + sheet names + len() + a row-by-row cell-text projection over SourceWorksheet::cells"
                },
                XlsxQuery::SelectedCell => {
                    "SourceBackedWorkbook::sheet(mixed-case name) + SourceWorksheet::cell(A1) on a workbook prepared outside the timer"
                },
            },
        }
    }

    /// The pairing and timing metadata this route restates in its result.
    pub(super) fn summary(self) -> Result<RouteSummary, Box<dyn Error>> {
        RouteSummary::for_case(self.case(), self.timed_calls())
    }
}

// ---------------------------------------------------------------------------
// Child protocol
// ---------------------------------------------------------------------------

/// A root prepared outside the timer for a query route.
#[allow(
    clippy::large_enum_variant,
    reason = "a stack value of the harness; boxing a leaf owner would add an allocation to the timed leaf leg"
)]
pub(super) enum Prepared {
    Unified(litchi::File),
    DocxEager(litchi_docx::Package),
    DocxSource(litchi_docx::source_backed::Package),
    Pptx(litchi_pptx::SourceBackedPresentation),
    Xlsx(LeafXlsx),
}

/// Open the root a query route queries, outside the timer, through the same
/// ingress its timed lifecycle twin uses.
pub(super) fn prepare(route: Route, source: &Path) -> Result<Prepared, Box<dyn Error>> {
    Ok(match route {
        Route::Docx {
            leg: Leg::Unified,
            ingress,
            ..
        } => Prepared::Unified(open_unified_docx(ingress, source)?),
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Eager,
            ..
        } => Prepared::DocxEager(leaf_docx_eager(fs::read(source)?)?),
        Route::Docx {
            leg: Leg::Leaf,
            ingress: Ingress::Source,
            ..
        } => Prepared::DocxSource(litchi_docx::source_backed::Package::open(source)?),
        Route::Pptx {
            leg: Leg::Unified,
            ingress,
            ..
        } => Prepared::Unified(open_unified_pptx(ingress, source)?),
        Route::Pptx {
            leg: Leg::Leaf,
            ingress,
            ..
        } => Prepared::Pptx(open_leaf_pptx(ingress, source)?),
        Route::Xlsx {
            leg: Leg::Unified, ..
        } => Prepared::Unified(open_unified_xlsx(source)?),
        Route::Xlsx { leg: Leg::Leaf, .. } => {
            Prepared::Xlsx(leaf_xlsx_open(Arc::new(FileSource::open(source)?))?)
        },
    })
}

fn open_unified_docx(ingress: Ingress, source: &Path) -> Result<litchi::File, Box<dyn Error>> {
    let file = match ingress {
        Ingress::Eager => litchi::open_bytes_with(fs::read(source)?, &eager_options())?,
        Ingress::Source => litchi::open(source)?,
    };
    if !matches!(file, litchi::File::Docx(_)) {
        return Err("unified DOCX route opened a non-DOCX variant".into());
    }
    Ok(file)
}

fn open_unified_pptx(ingress: Ingress, source: &Path) -> Result<litchi::File, Box<dyn Error>> {
    let file = match ingress {
        Ingress::Eager => litchi::open_bytes(fs::read(source)?)?,
        Ingress::Source => litchi::open(source)?,
    };
    if !matches!(file, litchi::File::Pptx(_)) {
        return Err("unified PPTX route opened a non-PPTX variant".into());
    }
    Ok(file)
}

fn open_unified_xlsx(source: &Path) -> Result<litchi::File, Box<dyn Error>> {
    let file = litchi::open(source)?;
    if !matches!(file, litchi::File::Xlsx(_)) {
        return Err("unified XLSX route opened a non-XLSX variant".into());
    }
    Ok(file)
}

fn open_leaf_pptx(
    ingress: Ingress,
    source: &Path,
) -> Result<litchi_pptx::SourceBackedPresentation, Box<dyn Error>> {
    Ok(match ingress {
        Ingress::Eager => litchi_pptx::SourceBackedPresentation::from_read_at(Arc::new(
            OwnedSource::new(fs::read(source)?),
        ))?,
        Ingress::Source => litchi_pptx::SourceBackedPresentation::open(source)?,
    })
}

/// The XLSX root a timed route leaves behind.
#[allow(
    clippy::large_enum_variant,
    reason = "a stack value of the harness; boxing a leaf owner would add an allocation to the timed leaf leg"
)]
enum XlsxRoot {
    Unified(litchi::File),
    Leaf(LeafXlsx),
}

/// The selected worksheet handle a timed selected-cell route leaves behind.
enum SelectedSheet {
    Unified(litchi::format::xlsx::SelectedWorksheet),
    Leaf(litchi_xlsx::SourceWorksheet),
}

impl SelectedSheet {
    fn name(&self) -> &str {
        match self {
            Self::Unified(sheet) => sheet.name(),
            Self::Leaf(sheet) => sheet.name(),
        }
    }

    fn position(&self) -> usize {
        match self {
            Self::Unified(sheet) => sheet.position(),
            Self::Leaf(sheet) => sheet.position(),
        }
    }
}

/// What the timed XLSX scope produced; released only after the correctness
/// checks and every operation-only snapshot, as the predecessors do.
pub(super) struct XlsxTimed {
    root: XlsxRoot,
    projection: Option<(Vec<String>, usize, String)>,
    selected: Option<(SelectedSheet, litchi_xlsx::SourceCellView)>,
    selector: Option<(String, String)>,
}

/// The values a timed child operation leaves for the untimed checks.
#[derive(Default)]
pub(super) struct Timed {
    /// The full text a DOCX open-plus-full-text lifecycle returned.
    pub(super) docx_full_text: Option<String>,
    pub(super) xlsx: Option<XlsxTimed>,
}

fn expect_count(count: usize, expected: usize, context: &str) -> Result<(), Box<dyn Error>> {
    if count == expected {
        Ok(())
    } else {
        Err(format!("{context} returned {count}, the corpus pins {expected}").into())
    }
}

/// The timed body of one paired route. `prepared` is the root a query route
/// prepared before the timer; it stays owned by the caller (and so is dropped
/// after the child reports) except for the XLSX selected-cell route, whose
/// workbook moves into the retained values as its predecessor's does.
pub(super) fn run_timed(
    route: Route,
    source: &Path,
    prepared: &mut Option<Prepared>,
) -> Result<Timed, Box<dyn Error>> {
    let docx_paragraphs = SemanticShape::Medium.docx_paragraphs();
    match route {
        Route::Docx {
            leg,
            ingress,
            query,
        } => {
            let full_text = match query {
                DocxQuery::Open => {
                    match leg {
                        Leg::Unified => {
                            std::hint::black_box(open_unified_docx(ingress, source)?);
                        },
                        Leg::Leaf => match ingress {
                            Ingress::Eager => {
                                std::hint::black_box(leaf_docx_eager(fs::read(source)?)?);
                            },
                            Ingress::Source => {
                                std::hint::black_box(litchi_docx::source_backed::Package::open(
                                    source,
                                )?);
                            },
                        },
                    }
                    None
                },
                DocxQuery::ParagraphCount | DocxQuery::ListParagraphs | DocxQuery::FullText => {
                    let root = prepared
                        .as_ref()
                        .ok_or("DOCX query route has no prepared root")?;
                    docx_query(root, query)?;
                    None
                },
                DocxQuery::OpenParagraphCountLifecycle => {
                    let root = prepare(route, source)?;
                    let count = docx_paragraph_count(&root)?;
                    expect_count(count, docx_paragraphs, "DOCX lifecycle paragraph count")?;
                    std::hint::black_box((root, count));
                    None
                },
                DocxQuery::OpenFullTextLifecycle => {
                    // The String outlives the timer; the root is dropped inside
                    // it, as in the predecessor lifecycle.
                    let text = {
                        let root = prepare(route, source)?;
                        let text = docx_text(&root)?;
                        std::hint::black_box(&text);
                        text
                    };
                    Some(text)
                },
            };
            Ok(Timed {
                docx_full_text: full_text,
                xlsx: None,
            })
        },
        Route::Pptx { query, .. } => {
            match query {
                PptxQuery::Open => {
                    std::hint::black_box(prepare(route, source)?);
                },
                PptxQuery::ListSlides | PptxQuery::SlideCount | PptxQuery::SelectedSlide => {
                    let root = prepared
                        .as_ref()
                        .ok_or("PPTX query route has no prepared root")?;
                    pptx_query(root, query)?;
                },
                PptxQuery::OpenSlideCountLifecycle => {
                    let root = prepare(route, source)?;
                    let count = pptx_query(&root, PptxQuery::SlideCount)?;
                    std::hint::black_box((root, count));
                },
                PptxQuery::OpenSelectedSlideLifecycle => {
                    let root = prepare(route, source)?;
                    pptx_query(&root, PptxQuery::SelectedSlide)?;
                    std::hint::black_box(&root);
                },
            }
            Ok(Timed::default())
        },
        Route::Xlsx { leg, query } => {
            let xlsx = match query {
                XlsxQuery::Open | XlsxQuery::OpenLifecycle => {
                    let root = match leg {
                        Leg::Unified => XlsxRoot::Unified(open_unified_xlsx(source)?),
                        Leg::Leaf => {
                            XlsxRoot::Leaf(leaf_xlsx_open(Arc::new(FileSource::open(source)?))?)
                        },
                    };
                    let projection = (query == XlsxQuery::OpenLifecycle)
                        .then(|| xlsx_names_count_text(&root))
                        .transpose()?;
                    XlsxTimed {
                        root,
                        projection,
                        selected: None,
                        selector: None,
                    }
                },
                XlsxQuery::SelectedCell => {
                    let target = XLSX_FILE_SELECTED_CELL_TARGET;
                    let root = match prepared.take() {
                        Some(Prepared::Unified(file)) => XlsxRoot::Unified(file),
                        Some(Prepared::Xlsx(leaf)) => XlsxRoot::Leaf(leaf),
                        _ => return Err("XLSX selected-cell route has no prepared root".into()),
                    };
                    // The selector strings are derived the same way as the
                    // predecessor's, which prepares them outside the timer.
                    let selector = mixed_ascii_case(&crate::xlsx_sheet_name(target.sheet));
                    let address = crate::xlsx_address(target.row, target.column)?;
                    let selected = match &root {
                        XlsxRoot::Unified(litchi::File::Xlsx(xlsx)) => {
                            let sheet = xlsx
                                .sheet(selector.as_str())?
                                .ok_or("unified XLSX selected sheet is missing")?;
                            let cell = sheet.cell(address.as_str())?;
                            (SelectedSheet::Unified(sheet), cell)
                        },
                        XlsxRoot::Unified(_) => {
                            return Err("unified XLSX route holds a non-XLSX file".into());
                        },
                        XlsxRoot::Leaf(leaf) => {
                            let sheet = leaf
                                .workbook
                                .sheet(selector.as_str())?
                                .ok_or("leaf XLSX selected sheet is missing")?;
                            let cell = sheet.cell(address.as_str())?;
                            (SelectedSheet::Leaf(sheet), cell)
                        },
                    };
                    std::hint::black_box(&selected);
                    XlsxTimed {
                        root,
                        projection: None,
                        selected: Some(selected),
                        selector: Some((selector, address)),
                    }
                },
            };
            Ok(Timed {
                docx_full_text: None,
                xlsx: Some(xlsx),
            })
        },
    }
}

fn docx_paragraph_count(root: &Prepared) -> Result<usize, Box<dyn Error>> {
    Ok(match root {
        Prepared::Unified(file) => file.document()?.paragraph_count()?,
        Prepared::DocxEager(package) => package.document()?.paragraph_count()?,
        Prepared::DocxSource(package) => package.document()?.paragraph_count()?,
        _ => return Err("DOCX query reached a non-DOCX root".into()),
    })
}

/// The full text of a DOCX root. The leaf legs ask for the paragraph
/// separator `File::text()` writes since M0a step S8a, so both legs of a pair
/// build the same text with the same validating pass.
fn docx_text(root: &Prepared) -> Result<String, Box<dyn Error>> {
    Ok(match root {
        Prepared::Unified(file) => file.text()?,
        Prepared::DocxEager(package) => package
            .document()?
            .text_with_paragraph_separator(DOCX_PARAGRAPH_SEPARATOR)?,
        Prepared::DocxSource(package) => package
            .document()?
            .extract_text_with_paragraph_separator(DOCX_PARAGRAPH_SEPARATOR)?,
        _ => return Err("DOCX query reached a non-DOCX root".into()),
    })
}

fn docx_query(root: &Prepared, query: DocxQuery) -> Result<(), Box<dyn Error>> {
    match query {
        DocxQuery::ParagraphCount => {
            std::hint::black_box(docx_paragraph_count(root)?);
        },
        DocxQuery::ListParagraphs => match root {
            Prepared::Unified(file) => {
                std::hint::black_box(file.document()?.paragraphs()?);
            },
            Prepared::DocxEager(package) => {
                std::hint::black_box(package.document()?.paragraphs()?);
            },
            Prepared::DocxSource(package) => {
                std::hint::black_box(package.document()?.paragraphs()?);
            },
            _ => return Err("DOCX query reached a non-DOCX root".into()),
        },
        DocxQuery::FullText => {
            std::hint::black_box(docx_text(root)?);
        },
        _ => return Err("DOCX open route reached the query body".into()),
    }
    Ok(())
}

/// Run one PPTX query and return the slide count it observed (zero for the
/// selected-slide query, which returns no count).
fn pptx_query(root: &Prepared, query: PptxQuery) -> Result<usize, Box<dyn Error>> {
    match (root, query) {
        (Prepared::Unified(file), PptxQuery::ListSlides) => {
            let slides = file.presentation()?.slides()?;
            expect_count(slides.len(), PPTX_SOURCE_SLIDE_COUNT, "PPTX slide list")?;
            std::hint::black_box(slides);
            Ok(0)
        },
        (Prepared::Pptx(presentation), PptxQuery::ListSlides) => {
            let slides = presentation
                .slides()
                .map(|slide| slide.text_and_name())
                .collect::<Result<Vec<_>, _>>()?;
            expect_count(slides.len(), PPTX_SOURCE_SLIDE_COUNT, "PPTX slide list")?;
            std::hint::black_box(slides);
            Ok(0)
        },
        (Prepared::Unified(file), PptxQuery::SlideCount) => {
            let count = file.presentation()?.slide_count()?;
            expect_count(count, PPTX_SOURCE_SLIDE_COUNT, "PPTX slide count")?;
            Ok(std::hint::black_box(count))
        },
        (Prepared::Pptx(presentation), PptxQuery::SlideCount) => {
            let count = presentation.slide_count();
            expect_count(count, PPTX_SOURCE_SLIDE_COUNT, "PPTX slide count")?;
            Ok(std::hint::black_box(count))
        },
        (Prepared::Unified(file), PptxQuery::SelectedSlide) => {
            let slide = file
                .presentation()?
                .slide(PPTX_FILE_SELECTED_POSITION)?
                .ok_or("unified PPTX selected slide is missing")?;
            std::hint::black_box(slide);
            Ok(0)
        },
        (Prepared::Pptx(presentation), PptxQuery::SelectedSlide) => {
            let slide = presentation
                .slide(PPTX_FILE_SELECTED_POSITION)
                .ok_or("leaf PPTX selected slide is missing")?
                .text_and_name()?;
            std::hint::black_box(slide);
            Ok(0)
        },
        _ => Err("PPTX query reached a non-PPTX root or an open route".into()),
    }
}

fn xlsx_names_count_text(root: &XlsxRoot) -> Result<(Vec<String>, usize, String), Box<dyn Error>> {
    Ok(match root {
        XlsxRoot::Unified(file) => {
            let workbook = file.workbook()?;
            (
                workbook.worksheet_names()?,
                workbook.worksheet_count()?,
                file.text()?,
            )
        },
        XlsxRoot::Leaf(leaf) => (
            leaf_xlsx_names(&leaf.workbook),
            leaf.workbook.len(),
            leaf_xlsx_text(&leaf.workbook)?,
        ),
    })
}

fn xlsx_metadata(root: &XlsxRoot) -> Result<Metadata, Box<dyn Error>> {
    Ok(match root {
        XlsxRoot::Unified(file) => file.metadata()?.unwrap_or_default(),
        XlsxRoot::Leaf(leaf) => leaf.metadata.clone(),
    })
}

/// The untimed XLSX evidence: the final source digest, the semantic
/// projection digest checked against the typed oracle, and the selected-cell
/// evidence of the selected-cell routes.
pub(super) fn verify_xlsx(
    route: Route,
    source: &Path,
    corpus: &crate::Corpus,
    allow_page_aligned_source: bool,
    timed: &Timed,
) -> Result<(String, String, Option<XlsxSelectedCellEvidence>), Box<dyn Error>> {
    assert_pinned_xlsx_corpus(corpus)?;
    let xlsx = timed
        .xlsx
        .as_ref()
        .ok_or("XLSX route did not retain its timed root")?;
    let initial_bytes = fs::read(source)?;
    if !allow_page_aligned_source && initial_bytes.len() != corpus.archive.len() {
        return Err("XLSX filesystem source length differs from corpus manifest".into());
    }
    let expected = xlsx_typed_semantic_oracle(corpus)?;
    let (worksheet_names, worksheet_count, full_text) = match &xlsx.projection {
        Some((names, count, text)) => (names.clone(), *count, text.clone()),
        None => xlsx_names_count_text(&xlsx.root)?,
    };
    let observed = XlsxRootSemanticProjection {
        worksheet_names,
        worksheet_count,
        full_text,
        metadata_sha256: crate::xlsx_root_metadata_digest(&xlsx_metadata(&xlsx.root)?)?,
    };
    if observed != expected {
        return Err(format!(
            "{} semantic projection differs from the typed oracle",
            route.case().name()
        )
        .into());
    }
    let selected_cell = route
        .is_xlsx_selected_cell()
        .then(|| verify_xlsx_selected_cell(xlsx, corpus))
        .transpose()?;
    let semantic_sha256 = xlsx_semantic_projection_sha256(&observed)?;
    let final_bytes = fs::read(source)?;
    let source_sha256 = crate::sha256_hex(&final_bytes);
    if !allow_page_aligned_source && source_sha256 != corpus.manifest.archive_sha256 {
        return Err("XLSX filesystem source hash differs from its fixed pin".into());
    }
    Ok((source_sha256, semantic_sha256, selected_cell))
}

fn verify_xlsx_selected_cell(
    xlsx: &XlsxTimed,
    corpus: &crate::Corpus,
) -> Result<XlsxSelectedCellEvidence, Box<dyn Error>> {
    let target = XLSX_FILE_SELECTED_CELL_TARGET;
    let (sheet, observed) = xlsx
        .selected
        .as_ref()
        .ok_or("XLSX selected-cell route omitted its timed cell view")?;
    let (selector, address) = xlsx
        .selector
        .as_ref()
        .ok_or("XLSX selected-cell route omitted its selector")?;
    let expected_sheet_name = crate::xlsx_sheet_name(target.sheet);
    let expected_address = crate::xlsx_address(target.row, target.column)?;
    if *selector != mixed_ascii_case(&expected_sheet_name) || *address != expected_address {
        return Err("XLSX selected-cell route changed its selector".into());
    }
    if sheet.name() != expected_sheet_name || sheet.position() != target.sheet {
        return Err("XLSX selected-cell worksheet identity differs from the corpus".into());
    }
    let typed = litchi_xlsx::Package::from_bytes(&corpus.archive)?;
    let typed_sheet = typed
        .workbook_ref()
        .sheet(expected_sheet_name.as_str())?
        .ok_or("typed XLSX selected-cell oracle sheet is missing")?;
    let expected = match typed_sheet.cell(expected_address.as_str())? {
        litchi_xlsx::cell::View::Missing => litchi_xlsx::SourceCellView::Missing,
        litchi_xlsx::cell::View::Covered(range) => litchi_xlsx::SourceCellView::Covered(range),
        litchi_xlsx::cell::View::Stored(cell) => litchi_xlsx::SourceCellView::Stored(cell.clone()),
        _ => return Err("typed XLSX selected-cell oracle returned an unknown view".into()),
    };
    if *observed != expected {
        return Err("XLSX selected-cell view differs from the typed eager oracle".into());
    }
    let expected_value = crate::xlsx_value(target).to_string();
    match observed {
        litchi_xlsx::SourceCellView::Stored(litchi_xlsx::Cell::Value(
            litchi_xlsx::Value::Number(value),
        )) if value.as_str() == expected_value => {
            let evidence = XlsxSelectedCellEvidence {
                canonical_sheet_name: expected_sheet_name,
                sheet_position: target.sheet,
                prepared_selector: selector.clone(),
                cell_address: expected_address,
                view: "stored".to_owned(),
                value_kind: "number".to_owned(),
                lexical_value: value.as_str().to_owned(),
                digest: String::new(),
            };
            let digest = xlsx_selected_cell_evidence_digest(&evidence)?;
            Ok(XlsxSelectedCellEvidence { digest, ..evidence })
        },
        _ => Err("XLSX selected-cell value differs from the deterministic generator".into()),
    }
}

// ---------------------------------------------------------------------------
// Untimed output verification
// ---------------------------------------------------------------------------

/// The semantic signature of a DOCX through the unified document view, in
/// the shape `docx_package_signature` gives the leaf package.
fn unified_docx_signature(file: &litchi::File) -> Result<String, Box<dyn Error>> {
    let document = file.document()?;
    let paragraphs = document
        .paragraphs()?
        .into_iter()
        .map(|paragraph| paragraph.text())
        .collect::<Result<Vec<_>, _>>()?;
    let tables = document
        .tables()?
        .iter()
        .map(unified_table_projection)
        .collect::<Result<Vec<_>, _>>()?;
    let elements = document
        .elements()?
        .into_iter()
        .map(|element| -> Result<(&str, String), Box<dyn Error>> {
            if let Some(paragraph) = element.as_paragraph() {
                return Ok(("paragraph", paragraph.text()?));
            }
            if let Some(table) = element.as_table() {
                return Ok((
                    "table",
                    serde_json::to_string(&unified_table_projection(table)?)?,
                ));
            }
            Err("unified DOCX element has no supported projection".into())
        })
        .collect::<Result<Vec<_>, _>>()?;
    let metadata = file.metadata()?.unwrap_or_default();
    Ok(crate::sha256_hex(&serde_json::to_vec(&(
        document.paragraph_count()?,
        file.text()?,
        paragraphs,
        tables,
        elements,
        metadata,
    ))?))
}

fn unified_table_projection(
    table: &litchi::compat::word::Table,
) -> Result<(usize, Vec<(usize, Vec<String>)>), Box<dyn Error>> {
    let row_count = table.row_count()?;
    let rows = table.rows()?;
    if rows.len() != row_count {
        return Err("unified DOCX table row count disagrees with its rows".into());
    }
    let rows = rows
        .into_iter()
        .map(|row| -> Result<(usize, Vec<String>), Box<dyn Error>> {
            let cell_count = row.cell_count()?;
            let cells = row
                .cells()?
                .into_iter()
                .map(|cell| cell.text())
                .collect::<Result<Vec<_>, _>>()?;
            if cells.len() != cell_count {
                return Err("unified DOCX row cell count disagrees with its cells".into());
            }
            Ok((cell_count, cells))
        })
        .collect::<Result<Vec<_>, _>>()?;
    Ok((row_count, rows))
}

type PptxSignature = (usize, i64, i64, String, String, String);

fn pptx_signature_digests(
    slides: impl Iterator<Item = Result<(String, String), Box<dyn Error>>>,
    selected: (String, String),
) -> Result<(String, String), Box<dyn Error>> {
    let mut slides_hasher = Sha256::new();
    for (position, slide) in slides.enumerate() {
        let (text, name) = slide?;
        slides_hasher.update(position.to_le_bytes());
        slides_hasher.update((text.len() as u64).to_le_bytes());
        slides_hasher.update(text.as_bytes());
        slides_hasher.update((name.len() as u64).to_le_bytes());
        slides_hasher.update(name.as_bytes());
    }
    let mut selected_hasher = Sha256::new();
    selected_hasher.update(selected.0.as_bytes());
    selected_hasher.update([0]);
    selected_hasher.update(selected.1.as_bytes());
    Ok((
        crate::sha256_hex(slides_hasher.finalize().as_slice()),
        crate::sha256_hex(selected_hasher.finalize().as_slice()),
    ))
}

fn unified_pptx_signature(file: &litchi::File) -> Result<PptxSignature, Box<dyn Error>> {
    let litchi::File::Pptx(handle) = file else {
        return Err("unified PPTX signature needs a PPTX file".into());
    };
    let (width, height) = handle.slide_size()?;
    let metadata = serde_json::to_string(&file.metadata()?)?;
    let presentation = file.presentation()?;
    let slides = presentation.slides()?;
    let selected = presentation
        .slide(PPTX_FILE_SELECTED_POSITION)?
        .ok_or("unified PPTX signature selected slide is missing")?;
    let (slides_sha256, selected_sha256) = pptx_signature_digests(
        slides
            .iter()
            .map(|slide| Ok((slide.text()?, slide.name()?.unwrap_or_default()))),
        (selected.text()?, selected.name()?.unwrap_or_default()),
    )?;
    Ok((
        presentation.slide_count()?,
        width,
        height,
        metadata,
        slides_sha256,
        selected_sha256,
    ))
}

fn leaf_pptx_signature(
    presentation: &litchi_pptx::SourceBackedPresentation,
) -> Result<PptxSignature, Box<dyn Error>> {
    let (width, height) = presentation.slide_size()?;
    let metadata = presentation
        .properties_with_tolerance(&mut litchi_pptx::ToleranceReport::default())?
        .filter(Metadata::has_data);
    let metadata = serde_json::to_string(&metadata)?;
    let selected = presentation
        .slide(PPTX_FILE_SELECTED_POSITION)
        .ok_or("leaf PPTX signature selected slide is missing")?
        .text_and_name()?;
    let (slides_sha256, selected_sha256) = pptx_signature_digests(
        presentation
            .slides()
            .map(|slide| slide.text_and_name().map_err(Into::into)),
        selected,
    )?;
    Ok((
        presentation.slide_count(),
        width,
        height,
        metadata,
        slides_sha256,
        selected_sha256,
    ))
}

/// Untimed correctness for a paired route, run by the child after every
/// operation-only snapshot and by the parent after every sample: the unified
/// projections over both owners and the leaf projection must agree, and the
/// corpus generator pins the counts and the timed full text.
pub(super) fn verify_output(
    route: Route,
    source: &Path,
    corpus: &crate::Corpus,
    allow_page_aligned_source: bool,
    timed_docx_full_text: Option<&str>,
) -> Result<(), Box<dyn Error>> {
    match route {
        Route::Docx { .. } => verify_docx(
            route,
            source,
            corpus,
            allow_page_aligned_source,
            timed_docx_full_text,
        ),
        Route::Pptx { .. } => verify_pptx(source, corpus, allow_page_aligned_source),
        // XLSX correctness is validated against the exact timed root before
        // this output-only dispatch, as for the predecessors.
        Route::Xlsx { .. } => Ok(()),
    }
}

fn verify_source_identity(
    bytes: &[u8],
    corpus: &crate::Corpus,
    allow_page_aligned_source: bool,
    format: &str,
) -> Result<(), Box<dyn Error>> {
    if !allow_page_aligned_source
        && (bytes.len() != corpus.manifest.archive_bytes
            || crate::sha256_hex(bytes) != corpus.manifest.archive_sha256)
    {
        return Err(format!("{format} filesystem source differs from the corpus manifest").into());
    }
    Ok(())
}

fn verify_docx(
    route: Route,
    source: &Path,
    corpus: &crate::Corpus,
    allow_page_aligned_source: bool,
    timed_docx_full_text: Option<&str>,
) -> Result<(), Box<dyn Error>> {
    if corpus.manifest.generator != DOCX_FILE_CORPUS_GENERATOR {
        return Err("DOCX filesystem source has the wrong corpus generator".into());
    }
    let bytes = fs::read(source)?;
    let source_sha256 = crate::sha256_hex(&bytes);
    verify_source_identity(&bytes, corpus, allow_page_aligned_source, "DOCX")?;
    let leaf = leaf_docx_eager(bytes.clone())?;
    let leaf_metadata = leaf.metadata().cloned().unwrap_or_default();
    let leaf_signature = docx_package_signature(&leaf, &leaf_metadata)?;
    let unified_eager = open_unified_docx(Ingress::Eager, source)?;
    let unified_source = open_unified_docx(Ingress::Source, source)?;
    if unified_docx_signature(&unified_eager)? != leaf_signature
        || unified_docx_signature(&unified_source)? != leaf_signature
    {
        return Err("DOCX unified and leaf semantic signatures differ".into());
    }
    expect_count(
        leaf.document()?.paragraph_count()?,
        SemanticShape::Medium.docx_paragraphs(),
        "DOCX corpus paragraph count",
    )?;
    if let Some(observed) = timed_docx_full_text {
        if !route.is_docx_full_text_lifecycle() {
            return Err("a non-lifecycle DOCX route retained timed full text".into());
        }
        verify_docx_timed_full_text(observed)?;
    }
    if docx_archive_signature(&bytes)?.semantic_sha256
        != docx_archive_signature(&corpus.archive)?.semantic_sha256
    {
        return Err("DOCX filesystem archive topology or payload hashes differ".into());
    }
    assert_source_sha256(source, &source_sha256)?;
    Ok(())
}

fn verify_pptx(
    source: &Path,
    corpus: &crate::Corpus,
    allow_page_aligned_source: bool,
) -> Result<(), Box<dyn Error>> {
    if corpus.manifest.generator != PPTX_FILE_CORPUS_GENERATOR {
        return Err("PPTX filesystem source has the wrong corpus generator".into());
    }
    let bytes = fs::read(source)?;
    verify_source_identity(&bytes, corpus, allow_page_aligned_source, "PPTX")?;
    let leaf = open_leaf_pptx(Ingress::Eager, source)?;
    let leaf_signature = leaf_pptx_signature(&leaf)?;
    let unified_bytes = open_unified_pptx(Ingress::Eager, source)?;
    let unified_path = open_unified_pptx(Ingress::Source, source)?;
    if unified_pptx_signature(&unified_bytes)? != leaf_signature
        || unified_pptx_signature(&unified_path)? != leaf_signature
    {
        return Err("PPTX unified and leaf semantic signatures differ".into());
    }
    expect_count(
        leaf.slide_count(),
        PPTX_SOURCE_SLIDE_COUNT,
        "PPTX corpus slide count",
    )?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_route_round_trips_through_its_selector() {
        for (route, case) in ROUTES {
            assert_eq!(route.case(), case);
            assert_eq!(Route::parse(case.name()), Some(route));
            assert!(case.is_unified_ops_filesystem(), "{}", case.name());
            assert!(case.is_filesystem(), "{}", case.name());
            assert!(!Case::DEFAULT.contains(&case), "{}", case.name());
            let name = case.name();
            match route.leg() {
                Leg::Unified => assert!(name.contains("_file_unified_"), "{name}"),
                Leg::Leaf => assert!(name.contains("_file_leaf_"), "{name}"),
            }
            assert_eq!(route.is_docx(), name.starts_with("docx_file_"));
            assert_eq!(route.is_pptx(), name.starts_with("pptx_file_"));
            assert_eq!(route.is_xlsx(), name.starts_with("xlsx_file_"));
        }
        assert!(Route::parse("docx_file_source_open").is_none());
    }

    #[test]
    fn query_routes_are_prepared_and_never_cold_verified() {
        for (route, case) in ROUTES {
            let operation = super::super::Operation::Ops(route);
            let name = case.name();
            let prepared = !name.contains("_open");
            assert_eq!(route.is_prepared_query(), prepared, "{name}");
            assert_eq!(operation.supports_cold_verified(), !prepared, "{name}");
        }
    }

    #[test]
    fn source_routes_replay_the_matching_predecessor_query() {
        for (route, case) in ROUTES {
            let name = case.name();
            match route.source_replay() {
                Some(predecessor) => {
                    assert!(name.contains("_source_"), "{name}");
                    let suffix = name
                        .split_once("_source_")
                        .map(|(_, suffix)| suffix)
                        .expect("source route name");
                    assert!(predecessor.predecessor_name().ends_with(suffix), "{name}");
                    assert_eq!(
                        route.logical_read_counter_scope(),
                        "untimed_source_replay_only"
                    );
                },
                None => assert!(!name.contains("_source_") || route.is_xlsx(), "{name}"),
            }
        }
    }

    #[test]
    fn docx_full_text_legs_time_the_separated_text_of_the_oracle() {
        let corpus = crate::build_docx_source_edit_corpus().unwrap();
        let scratch = crate::filesystem::scratch_root(None, "docx-full-text-legs").unwrap();
        let path = scratch.join("source.docx");
        fs::write(&path, &corpus.archive).unwrap();
        let expected = super::super::expected_docx_full_text();
        assert_eq!(
            expected.matches(DOCX_PARAGRAPH_SEPARATOR).count() + 1,
            SemanticShape::Medium.docx_paragraphs()
        );
        let mut legs = 0;
        for (route, case) in ROUTES {
            if !matches!(
                route,
                Route::Docx {
                    query: DocxQuery::FullText,
                    ..
                }
            ) {
                continue;
            }
            let root = prepare(route, &path).unwrap();
            assert_eq!(docx_text(&root).unwrap(), expected, "{}", case.name());
            legs += 1;
        }
        fs::remove_dir_all(&scratch).unwrap();
        assert_eq!(legs, 4);
    }

    #[test]
    fn leaf_xlsx_text_matches_the_unified_handle_and_the_typed_oracle() {
        let corpus = crate::build_xlsx_cell_crud_corpus(crate::XlsxCellCrudShape::Medium).unwrap();
        let workbook = litchi_xlsx::SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(
            corpus.archive.clone(),
        )))
        .unwrap();
        let unified = litchi::open_bytes(corpus.archive.clone()).unwrap();
        let leaf_text = leaf_xlsx_text(&workbook).unwrap();
        assert_eq!(leaf_text, unified.text().unwrap());
        assert_eq!(
            leaf_text,
            xlsx_typed_semantic_oracle(&corpus).unwrap().full_text
        );
    }
}
