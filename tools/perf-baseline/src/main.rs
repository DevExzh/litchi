//! Reproducible, content-free baseline measurements for Litchi's OPC and CFB substrates.
//!
//! This is deliberately a standalone tool rather than a public crate dependency.
//! It creates all inputs in memory from fixed specifications and writes JSON that
//! identifies the exact generated corpus by SHA-256.

#![forbid(unsafe_code)]

mod filesystem;
mod process_metrics;

use std::{
    collections::{BTreeMap, BTreeSet},
    error::Error,
    fs::{self, File},
    io::{self, Cursor, Seek, SeekFrom, Write},
    num::{NonZeroU64, NonZeroUsize},
    ops::Range,
    path::PathBuf,
    process::Command,
    sync::{
        Arc, Barrier, Condvar, Mutex,
        atomic::{AtomicU64, Ordering},
        mpsc,
    },
    thread::JoinHandle,
    time::{Duration, Instant},
};

use litchi_cfb::{OleFile, OleWriter, SharedOleFile, SharedOleFileLimits};
use litchi_core::{
    Budget, CancellationSource, CheckStatus, ExecutionContext, ExecutionError, ExecutionLimits,
    Limits, ReadAt, SourceVersion, ValidateReport,
};
use litchi_core::{OwnedSource, Position, Resource};
use litchi_ole_common::object::{
    Editor as OleObjectEditor, Limits as OleObjectLimits, Targets as OleObjectTargets,
};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, OpenSession, PackURI, PackageWriter, PartData, ReadLimits,
    Relationships, SourceBackedPackage, SourceCacheDiagnostics, SourceCacheLimits, TargetMode,
    constants::{content_type as opc_content_type, relationship_type},
};
use litchi_xlsx::{
    Cell as XlsxCell, Rect, SourceBackedWorkbook, StreamingCell, StreamingCellValue,
    StreamingWorkbookLimits, StreamingWorkbookWriter, Value as XlsxValue, Workbook,
};
use serde::Serialize;
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;
use soapberry_zip::{PreservationIndex, ZipArchive};

const SCHEMA_VERSION: u32 = 1;
const DEFAULT_SAMPLES: usize = 15;
const DEFAULT_WARMUP_ITERATIONS: usize = 3;
const DEFAULT_RANGE_FIXED_LATENCY_US: u64 = 100;
const DEFAULT_RANGE_REQUEST_OVERHEAD_US: u64 = 25;
const DEFAULT_RANGE_BANDWIDTH_BYTES_PER_SECOND: u64 = 50 * 1024 * 1024;
const DEFAULT_RANGE_MAX_PHYSICAL_BYTES: usize = 4 * 1024;
const OPC_CACHE_SLOW_SOURCE_DELAY_US: u64 = 10_000;
const OPC_CACHE_COHORT_TIMEOUT: Duration = Duration::from_secs(10);
const CONTENT_TYPE: &str = "application/octet-stream";
const OPC_CORPUS_GENERATOR: &str = "litchi-opc-synthetic-v2";
const CFB_CORPUS_GENERATOR: &str = "litchi-cfb-synthetic-v1";
const CFB_SELECTIVE_CORPUS_GENERATOR: &str = "litchi-cfb-selective-read-v1";
const LEGACY_WRITER_CORPUS_GENERATOR: &str = "litchi-legacy-writer-v1";
const XLSX_CORPUS_GENERATOR: &str = "litchi-xlsx-synthetic-v1";
const SEMANTIC_DOCX_CORPUS_GENERATOR: &str = "litchi-docx-semantic-v1";
const DOCX_SOURCE_EDIT_CORPUS_GENERATOR: &str = "litchi-docx-source-edit-media-v1";
const SEMANTIC_PPTX_CORPUS_GENERATOR: &str = "litchi-pptx-semantic-v1";
const PPTX_SOURCE_EDIT_CORPUS_GENERATOR: &str = "litchi-pptx-source-edit-media-v1";
const XLSX_CALC_SOURCE_EDIT_CORPUS_GENERATOR: &str =
    "litchi-xlsx-calculation-metadata-source-edit-media-v1";
const XLSX_DEFINED_NAMES_SOURCE_EDIT_CORPUS_GENERATOR: &str =
    "litchi-xlsx-defined-names-source-edit-media-v1";
const XLSX_PAGE_BREAK_SOURCE_EDIT_CORPUS_GENERATOR: &str =
    "litchi-xlsx-page-break-source-edit-media-v1";
const XLSX_PAGE_MARGIN_SOURCE_EDIT_CORPUS_GENERATOR: &str =
    "litchi-xlsx-page-margin-source-edit-media-v1";
const XLSX_PAGE_SETUP_SOURCE_EDIT_CORPUS_GENERATOR: &str =
    "litchi-xlsx-page-setup-source-edit-media-v1";
const XLSX_PRINT_OPTIONS_SOURCE_EDIT_CORPUS_GENERATOR: &str =
    "litchi-xlsx-print-options-source-edit-media-v1";
const XLSX_SHEET_PROTECTION_SOURCE_EDIT_CORPUS_GENERATOR: &str =
    "litchi-xlsx-sheet-protection-source-edit-media-v1";
const XLSX_DATA_VALIDATION_SOURCE_EDIT_CORPUS_GENERATOR: &str =
    "litchi-xlsx-data-validation-source-edit-media-v1";
const XLSX_AUTO_FILTER_SOURCE_EDIT_CORPUS_GENERATOR: &str =
    "litchi-xlsx-auto-filter-source-edit-media-v1";
const XLSX_CONDITIONAL_FORMATTING_SOURCE_EDIT_CORPUS_GENERATOR: &str =
    "litchi-xlsx-conditional-formatting-source-edit-media-v1";
const XLSX_CELL_VALUES_SOURCE_EDIT_CORPUS_GENERATOR: &str =
    "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1";
const XLSX_MERGE_EDIT_CORPUS_GENERATOR: &str = "litchi-xlsx-merge-edit-sparse-a1-b2-v1";
const SEMANTIC_ODT_CORPUS_GENERATOR: &str = "litchi-odt-semantic-v1";
const ODF_REPAIR_CORPUS_GENERATOR: &str = "litchi-odf-mimetype-repair-v1";
const ODF_REPAIR_LOCAL_EXTRA: &[u8] = &[0x55, 0x54, 0x05, 0x00, 0x01, 0, 0, 0, 0];
const ODF_REPAIR_PUBLICATION_SCRATCH_BYTES: u64 = 64 * 1024;
const ODT_MEDIA_CORPUS_GENERATOR: &str = "litchi-odt-media-paragraph-publication-v1";
const ODT_RESOURCE_BATCH_CORPUS_GENERATOR: &str =
    "litchi-odt-embedded-resource-batch-publication-v1";
const ODT_MEDIA_APPEND_RUN_TEXT: &str = " appended run";
const ODT_MEDIA_APPEND_HYPERLINK_HREF: &str = "https://example.invalid/performance";
const ODT_MEDIA_APPEND_HYPERLINK_TEXT: &str = " performance link";
const ODT_MEDIA_INSERT_PARAGRAPH_TEXT: &str = "Inserted performance paragraph";
const SEMANTIC_ODS_CORPUS_GENERATOR: &str = "litchi-ods-semantic-v1";
const ODS_MEDIA_CORPUS_GENERATOR: &str = "litchi-ods-media-publication-v1";
const SEMANTIC_ODP_CORPUS_GENERATOR: &str = "litchi-odp-semantic-v1";
const ODP_MEDIA_CORPUS_GENERATOR: &str = "litchi-odp-media-textbox-publication-v1";
const ODP_TEXT_BOX_BATCH_CORPUS_GENERATOR: &str = "litchi-odp-cross-slide-textbox-publication-v1";
const SEMANTIC_RTF_CORPUS_GENERATOR: &str = "litchi-rtf-semantic-v2";
const RTF_LIFECYCLE_CORPUS_GENERATOR: &str = "litchi-rtf-paragraph-lifecycle-v1";
const XLSX_STREAMING_CORPUS_GENERATOR: &str = "litchi-xlsx-streaming-create-v1";
const RTF_STREAMING_CORPUS_GENERATOR: &str = "litchi-rtf-streaming-create-v1";
const RTF_LOGICAL_TAIL_SINK_WINDOW_BYTES: usize = 16 * 1024;
const XLSX_STREAMING_ROW_BYTES: u64 = 4 * 1024;
const RTF_STREAMING_SCRATCH_BYTES: u64 = 37;
const XLS_COMMENTS_EDIT_CORPUS_GENERATOR: &str = "litchi-xls-comments-opaque-heavy-v1";
const XLS_COMMENTS_SOURCE_COUNT: usize = 256;
const XLS_COMMENTS_BATCH_COUNT: usize = 256;
const XLS_COMMENTS_OPAQUE_STREAM_COUNT: usize = 8;
const XLS_COMMENTS_OPAQUE_STREAM_BYTES: usize = 2 * 1024 * 1024;
const XLS_VISIBILITY_CORPUS_GENERATOR: &str = "litchi-xls-visibility-opaque-v1";
const XLS_VISIBILITY_SHEET_COUNT: usize = litchi_xls::sheet_visibility::MAX_VISIBILITY_CHANGES + 2;
const XLS_VISIBILITY_BATCH_COUNT: usize = litchi_xls::sheet_visibility::MAX_VISIBILITY_CHANGES;
const XLS_VISIBILITY_OPAQUE_STREAM_COUNT: usize = 8;
const XLS_VISIBILITY_OPAQUE_STREAM_BYTES: usize = 256 * 1024;
const ODS_MEDIA_ENTRY_COUNT: usize = 8;
const ODS_MEDIA_ENTRY_BYTES: usize = 2 * 1024 * 1024;
const DOCX_SOURCE_MEDIA_ENTRY_COUNT: usize = 8;
const DOCX_SOURCE_MEDIA_ENTRY_BYTES: usize = 2 * 1024 * 1024;
const PPTX_SOURCE_MEDIA_ENTRY_COUNT: usize = 8;
const PPTX_SOURCE_MEDIA_ENTRY_BYTES: usize = 2 * 1024 * 1024;
const PPTX_SOURCE_SLIDE_COUNT: usize = 200;
const PPTX_SOURCE_TEXT_BOXES_PER_SLIDE: usize = 8;
const PPTX_MULTI_SLIDE_BATCH_COUNT: usize = 8;
const ODP_TEXT_BOX_BATCH_COUNT: usize = 8;
const ODT_RESOURCE_BATCH_COUNT: usize = 64;
const ODT_RESOURCE_PAYLOAD_BYTES: usize = 4 * 1024;
const XLSX_CALC_MEDIA_ENTRY_COUNT: usize = 8;
const XLSX_CALC_MEDIA_ENTRY_BYTES: usize = 2 * 1024 * 1024;
const XLSX_CELL_VALUES_MEDIA_ENTRY_COUNT: usize = 8;
const XLSX_CELL_VALUES_MEDIA_ENTRY_BYTES: usize = 512 * 1024;
const ODP_MEDIA_TEXT_BOX_NAME: &str = "litchi-perf-media-text-box";
const OLE_COMMON_CORPUS_GENERATOR: &str = "litchi-ole-common-copy-elision-v1";
const OLE_COMMON_TARGET: &str = "ole_common_edit_target.bin";
const OLE_COMMON_ORIGINAL: &[u8] = b"litchi-ole-common-original-stream-v1";
const OLE_COMMON_REPLACEMENT: &[u8] = b"litchi-ole-common-edited-stream-v1";
static NEXT_INSTRUMENTED_SOURCE_ID: AtomicU64 = AtomicU64::new(1);

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CorpusShape {
    Tiny,
    ManySmall,
    FewLarge,
    WideRoot,
}

impl CorpusShape {
    const ALL: [Self; 4] = [Self::Tiny, Self::ManySmall, Self::FewLarge, Self::WideRoot];

    const fn name(self) -> &'static str {
        match self {
            Self::Tiny => "tiny",
            Self::ManySmall => "many-small",
            Self::FewLarge => "few-large",
            Self::WideRoot => "wide-root",
        }
    }

    const fn entry_count(self) -> usize {
        match self {
            Self::Tiny => 3,
            Self::ManySmall => 256,
            Self::FewLarge => 4,
            Self::WideRoot => 2048,
        }
    }

    const fn entry_bytes(self) -> usize {
        match self {
            Self::Tiny => 512,
            Self::ManySmall => 1024,
            Self::FewLarge => 4 * 1024 * 1024,
            Self::WideRoot => 64,
        }
    }
}

/// Bounded document shapes for fresh legacy writer runs.
///
/// These are independent from `CorpusShape`: the container corpus matrix has
/// payload/compression concerns which are not applicable to writer API runs.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum WriterShape {
    Tiny,
    Large,
    PayloadHeavy,
}

/// Bounded, deterministic worksheet shapes for end-to-end XLSX measurements.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum XlsxShape {
    Tiny,
    Medium,
    DenseWide,
}

/// Deterministic media-rich multi-sheet corpora for scalar-cell CRUD.
///
/// This matrix is opt-in because it intentionally exercises the new bounded
/// multi-worksheet source closure rather than the historical 36-case matrix.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum XlsxCellCrudShape {
    Medium,
    DenseSparse,
}

impl XlsxCellCrudShape {
    const ALL: [Self; 2] = [Self::Medium, Self::DenseSparse];

    const fn name(self) -> &'static str {
        match self {
            Self::Medium => "medium",
            Self::DenseSparse => "dense-sparse",
        }
    }
}

/// Small, complete public-API DOCX/PPTX corpora.  These cases are opt-in so
/// their intentionally semantic workload does not alter the substrate matrix.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SemanticShape {
    Tiny,
    Medium,
    Large,
}

/// Transport and producer variants for the opt-in semantic RTF matrix.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RtfSemanticVariant {
    Plain,
    Byte1252,
    Lzfu,
    Watermark,
}

impl RtfSemanticVariant {
    #[cfg(test)]
    const ALL: [Self; 4] = [Self::Plain, Self::Byte1252, Self::Lzfu, Self::Watermark];

    const fn name(self) -> &'static str {
        match self {
            Self::Plain => "plain",
            Self::Byte1252 => "byte1252",
            Self::Lzfu => "lzfu",
            Self::Watermark => "watermark",
        }
    }

    const fn supports_shape(self, shape: SemanticShape) -> bool {
        !matches!(self, Self::Watermark) || matches!(shape, SemanticShape::Tiny)
    }

    const fn supports_case(self, case: Case) -> bool {
        case.uses_semantic_rtf()
            && (!matches!(
                case,
                Case::RtfSemanticOneEditSave
                    | Case::RtfSemanticOnePercentEditSave
                    | Case::RtfSemanticRemoveParagraphSave
                    | Case::RtfSemanticMoveParagraphSave
                    | Case::RtfLogicalTailAppend
                    | Case::RtfLogicalTailNoopSave
            ) || matches!(self, Self::Plain))
            && (!matches!(case, Case::RtfSemanticTextToSink) || !matches!(self, Self::Watermark))
    }

    const fn supports_validation(self) -> bool {
        !matches!(self, Self::Watermark)
    }
}

impl SemanticShape {
    const ALL: [Self; 3] = [Self::Tiny, Self::Medium, Self::Large];

    const fn name(self) -> &'static str {
        match self {
            Self::Tiny => "tiny",
            Self::Medium => "medium",
            Self::Large => "large",
        }
    }

    const fn docx_paragraphs(self) -> usize {
        match self {
            Self::Tiny => 24,
            Self::Medium => 200,
            Self::Large => 10_000,
        }
    }

    const fn rtf_paragraphs(self) -> usize {
        self.docx_paragraphs()
    }

    const fn streaming_units(self) -> usize {
        match self {
            Self::Tiny => 64,
            Self::Medium => 8_192,
            Self::Large => 131_072,
        }
    }

    const fn pptx_slides(self) -> usize {
        match self {
            Self::Tiny => 3,
            Self::Medium => 12,
            Self::Large => 100,
        }
    }

    const fn pptx_text_boxes_per_slide(self) -> usize {
        match self {
            Self::Tiny => 4,
            Self::Medium => 8,
            Self::Large => 100,
        }
    }

    const fn ods_sheet_count(self) -> usize {
        match self {
            Self::Tiny => 1,
            Self::Medium | Self::Large => 2,
        }
    }

    const fn ods_rows_per_sheet(self) -> usize {
        match self {
            Self::Tiny => 8,
            Self::Medium => 32,
            Self::Large => 128,
        }
    }

    const fn ods_columns_per_sheet(self) -> usize {
        self.ods_rows_per_sheet()
    }

    const fn ods_cell_count(self) -> usize {
        self.ods_sheet_count() * self.ods_rows_per_sheet() * self.ods_columns_per_sheet()
    }
}

impl XlsxShape {
    const ALL: [Self; 3] = [Self::Tiny, Self::Medium, Self::DenseWide];

    const fn name(self) -> &'static str {
        match self {
            Self::Tiny => "tiny",
            Self::Medium => "medium",
            Self::DenseWide => "dense-wide",
        }
    }

    const fn sheet_count(self) -> usize {
        match self {
            Self::Tiny => 3,
            Self::Medium => 4,
            Self::DenseWide => 2,
        }
    }

    const fn row_count(self) -> usize {
        match self {
            Self::Tiny => 8,
            Self::Medium => 32,
            Self::DenseWide => 256,
        }
    }

    const fn column_count(self) -> usize {
        match self {
            Self::Tiny => 8,
            Self::Medium => 32,
            Self::DenseWide => 256,
        }
    }
}

impl WriterShape {
    const ALL: [Self; 3] = [Self::Tiny, Self::Large, Self::PayloadHeavy];

    const fn name(self) -> &'static str {
        match self {
            Self::Tiny => "tiny",
            Self::Large => "large",
            Self::PayloadHeavy => "payload-heavy",
        }
    }

    const fn doc_paragraph_count(self) -> usize {
        match self {
            Self::Tiny => 3,
            Self::Large => 512,
            Self::PayloadHeavy => 128,
        }
    }

    const fn xls_dimensions(self) -> Option<(usize, usize, usize)> {
        match self {
            Self::Tiny => Some((1, 4, 4)),
            Self::Large => Some((4, 128, 16)),
            Self::PayloadHeavy => None,
        }
    }

    const fn ppt_dimensions(self) -> (usize, usize) {
        match self {
            Self::Tiny => (1, 2),
            Self::Large => (12, 12),
            Self::PayloadHeavy => (16, 8),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum PayloadKind {
    Compressible,
    Incompressible,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CfbSelectiveTarget {
    Mini,
    Fat,
}

impl CfbSelectiveTarget {
    const fn name(self) -> &'static str {
        match self {
            Self::Mini => "minifat-36-byte",
            Self::Fat => "fat-4mib",
        }
    }

    const fn target_bytes(self) -> usize {
        match self {
            Self::Mini => 36,
            Self::Fat => 4 * 1024 * 1024,
        }
    }
}

impl PayloadKind {
    const ALL: [Self; 2] = [Self::Compressible, Self::Incompressible];

    const fn name(self) -> &'static str {
        match self {
            Self::Compressible => "compressible",
            Self::Incompressible => "incompressible",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Case {
    ZipIndex,
    ZipReadOne,
    OpcOpen,
    OpcOpenOwned,
    OpcNoopSave,
    OpcMutatedSave,
    OpcSourceOpen,
    OpcSourceOpenMainRead,
    OpcSourceCachedMainRead,
    OpcSourceConcurrentSamePart,
    OpcSourceCacheBudgetBoundary,
    OpcSourceCacheControlContention,
    OpcSourceCacheManagedContention,
    OpcSourceOverlayOnePartSave,
    OpcFileEagerOpen,
    OpcFileSourceOpen,
    OpcFileEagerOnePartAtomicSave,
    OpcFileSourceOnePartAtomicSave,
    CfbFileSameLengthOverlayAtomicSave,
    DocxSourceBackedOneEditSave,
    PptxSourceBackedOneEditSave,
    PptxEagerBatchEditSave,
    PptxSourceBackedBatchEditSave,
    PptxEagerMultiSlideBatchEditSave,
    PptxSourceBackedMultiSlideBatchEditSave,
    XlsxEagerCalculationMetadataEditSave,
    XlsxSourceBackedCalculationMetadataEditSave,
    XlsxEagerDefinedNamesEditSave,
    XlsxSourceBackedDefinedNamesEditSave,
    XlsxEagerPageBreakEditSave,
    XlsxSourceBackedPageBreakEditSave,
    XlsxEagerPageMarginEditSave,
    XlsxSourceBackedPageMarginEditSave,
    XlsxEagerPageSetupEditSave,
    XlsxSourceBackedPageSetupEditSave,
    XlsxEagerPrintOptionsEditSave,
    XlsxSourceBackedPrintOptionsEditSave,
    XlsxEagerSheetProtectionEditSave,
    XlsxSourceBackedSheetProtectionEditSave,
    XlsxEagerDataValidationEditSave,
    XlsxSourceBackedDataValidationEditSave,
    XlsxEagerAutoFilterEditSave,
    XlsxSourceBackedAutoFilterEditSave,
    XlsxEagerConditionalFormattingEditSave,
    XlsxSourceBackedConditionalFormattingEditSave,
    XlsxEagerMergeCommitSave,
    XlsxEagerUnmergeCommitSave,
    XlsxEagerCellValuesOneEditSave,
    XlsxSourceBackedCellValuesOneEditSave,
    XlsxEagerCellValuesOnePercentEditSave,
    XlsxSourceBackedCellValuesOnePercentEditSave,
    XlsxEagerCellValuesBatchEditSave,
    XlsxSourceBackedCellValuesBatchEditSave,
    CfbOpen,
    CfbListStreams,
    CfbReadOne,
    CfbCreateStreamBorrowed,
    CfbCreateStreamOwned,
    OleCommonOpen,
    OleCommonPutStreamPublish,
    OleCommonFinishRender,
    OleCommonOneEditSave,
    CfbSharedOpen,
    CfbSharedReadOne,
    CfbSharedConcurrentReads,
    CfbSelectiveMiniLegacyRead,
    CfbSelectiveMiniSharedRead,
    CfbSelectiveFatLegacyRead,
    CfbSelectiveFatSharedRead,
    DocFreshWriteTo,
    XlsFreshWriteTo,
    PptFreshWriteTo,
    DocSemanticOpen,
    DocSemanticListParagraphs,
    DocSemanticOneParagraph,
    DocSemanticFullText,
    DocSemanticNoopEditSave,
    DocSemanticOneEditSave,
    DocBodySnapshotListParagraphs,
    XlsSemanticOpen,
    XlsSemanticListWorksheets,
    XlsSemanticOneCell,
    XlsSemanticFullCellScan,
    XlsSemanticNoopEditSave,
    XlsSemanticOneEditSave,
    XlsValidationReport,
    XlsCommentsEagerEditSave,
    XlsCommentsSourceBackedEditSave,
    XlsCommentsEagerBatchEditSave,
    XlsCommentsSourceBackedBatchEditSave,
    XlsVisibilityEagerEditSave,
    XlsVisibilitySourceBackedEditSave,
    XlsVisibilityEagerBatchEditSave,
    XlsVisibilitySourceBackedBatchEditSave,
    PptSemanticOpen,
    PptSemanticListSlides,
    PptSemanticOneShapeText,
    PptSemanticFullText,
    PptSlideOrderSnapshotOpen,
    PptTextEditOneEditSave,
    PptSemanticNoopEditSave,
    PptSemanticOneEditSave,
    XlsxOpenOwned,
    XlsxListSheets,
    XlsxFirstCell,
    XlsxFullCellScan,
    XlsxNarrowColumnRangeScan,
    XlsxNoopCommit,
    XlsxNoopCommitSave,
    XlsxOneCellCommit,
    XlsxOneCellCommitFirstRead,
    XlsxOneCellCommitSave,
    XlsxOnePercentCommit,
    XlsxOnePercentCommitSave,
    XlsxSourceOpen,
    XlsxSourceListSheets,
    XlsxSourceFirstCell,
    XlsxSourceNarrowColumnRangeScan,
    XlsxStreamingCreate,
    OpcRangeSourceOpen,
    OpcRangeSourceOpenMainRead,
    XlsxRangeSourceOpen,
    XlsxRangeSourceListSheets,
    XlsxRangeSourceFirstCell,
    XlsxRangeSourceNarrowColumnRangeScan,
    OpcOpenSessionScaling,
    CfbBulkReadScaling,
    RtfSemanticOpen,
    RtfSemanticParagraphCount,
    RtfSemanticListParagraphs,
    RtfSemanticCollectParagraphs,
    RtfSemanticOneParagraph,
    RtfSemanticFullText,
    RtfSemanticTextToSink,
    RtfSemanticStreamSave,
    RtfSemanticNoopEditSave,
    RtfSemanticOneEditSave,
    RtfSemanticOnePercentEditSave,
    RtfSemanticRemoveParagraphSave,
    RtfSemanticMoveParagraphSave,
    RtfLogicalTailAppend,
    RtfLogicalTailNoopSave,
    RtfValidationReport,
    RtfStreamingCreate,
    DocxSemanticOpen,
    DocxSemanticListParagraphs,
    DocxSemanticOneParagraph,
    DocxSemanticFullText,
    DocxSemanticCreateSmall,
    DocxSemanticNoopEditSave,
    DocxSemanticOneEditSave,
    DocxSemanticOnePercentEditSave,
    DocxValidationReport,
    DocxSectionInventory,
    PptxSemanticOpen,
    PptxSemanticListSlides,
    PptxSemanticOneSlide,
    PptxSemanticFullText,
    PptxSemanticCreateSmall,
    PptxSemanticNoopEditSave,
    PptxSemanticOneEditSave,
    PptxSemanticOnePercentEditSave,
    PptxValidationReport,
    OdtSemanticOpen,
    OdtSemanticListParagraphs,
    OdtSemanticOneParagraph,
    OdtSemanticFullText,
    OdtSemanticCreateSmall,
    OdtSemanticNoopEditSave,
    OdtSemanticOneEditSave,
    OdtSemanticOnePercentEditSave,
    OdfValidationReport,
    OdfMimetypeRepairPlan,
    OdtMediaParagraphEditSave,
    OdtMediaLineBreakEditSave,
    OdtMediaAppendRunEditSave,
    OdtMediaAppendHyperlinkEditSave,
    OdtMediaInsertParagraphEditSave,
    OdtMediaRemoveParagraphEditSave,
    OdtEmbeddedResourceScalarReplaceSave,
    OdtEmbeddedResourceBatchReplaceSave,
    OdsSemanticOpen,
    OdsSemanticListSheets,
    OdsSemanticOneCell,
    OdsSemanticCellSweep,
    OdsSemanticFullCellText,
    OdsSemanticCreateSmall,
    OdsSemanticNoopEditSave,
    OdsSemanticOneEditSave,
    OdsSemanticOnePercentEditSave,
    OdsMediaOneEditSave,
    OdpSemanticOpen,
    OdpSemanticListSlides,
    OdpSemanticOneSlide,
    OdpSemanticFullText,
    OdpSemanticCreateSmall,
    OdpSemanticNoopEditSave,
    OdpSemanticOneEditSave,
    OdpMediaTextBoxEditSave,
    OdpMediaTextBoxScalarReplaceSave,
    OdpMediaTextBoxBatchReplaceSave,
}

impl Case {
    const DEFAULT: [Self; 36] = [
        Self::ZipIndex,
        Self::ZipReadOne,
        Self::OpcOpen,
        Self::OpcOpenOwned,
        Self::OpcNoopSave,
        Self::OpcMutatedSave,
        Self::OpcSourceOpen,
        Self::OpcSourceOpenMainRead,
        Self::OpcSourceCachedMainRead,
        Self::OpcSourceConcurrentSamePart,
        Self::CfbOpen,
        Self::CfbListStreams,
        Self::CfbReadOne,
        Self::CfbCreateStreamBorrowed,
        Self::CfbCreateStreamOwned,
        Self::CfbSharedOpen,
        Self::CfbSharedReadOne,
        Self::CfbSharedConcurrentReads,
        Self::DocFreshWriteTo,
        Self::XlsFreshWriteTo,
        Self::PptFreshWriteTo,
        Self::XlsxOpenOwned,
        Self::XlsxListSheets,
        Self::XlsxFirstCell,
        Self::XlsxFullCellScan,
        Self::XlsxNarrowColumnRangeScan,
        Self::XlsxNoopCommit,
        Self::XlsxNoopCommitSave,
        Self::XlsxOneCellCommit,
        Self::XlsxOneCellCommitSave,
        Self::XlsxOnePercentCommit,
        Self::XlsxOnePercentCommitSave,
        Self::XlsxSourceOpen,
        Self::XlsxSourceListSheets,
        Self::XlsxSourceFirstCell,
        Self::XlsxSourceNarrowColumnRangeScan,
    ];

    const fn name(self) -> &'static str {
        match self {
            Self::ZipIndex => "zip_index",
            Self::ZipReadOne => "zip_read_one",
            Self::OpcOpen => "opc_open",
            Self::OpcOpenOwned => "opc_open_owned",
            Self::OpcNoopSave => "opc_noop_save",
            Self::OpcMutatedSave => "opc_mutated_save",
            Self::OpcSourceOpen => "opc_source_open",
            Self::OpcSourceOpenMainRead => "opc_source_open_main_read",
            Self::OpcSourceCachedMainRead => "opc_source_cached_main_read",
            Self::OpcSourceConcurrentSamePart => "opc_source_concurrent_same_part",
            Self::OpcSourceCacheBudgetBoundary => "opc_source_cache_budget_boundary",
            Self::OpcSourceCacheControlContention => "opc_source_cache_control_contention",
            Self::OpcSourceCacheManagedContention => "opc_source_cache_managed_contention",
            Self::OpcSourceOverlayOnePartSave => "opc_source_overlay_one_part_save",
            Self::OpcFileEagerOpen => "opc_file_eager_open",
            Self::OpcFileSourceOpen => "opc_file_source_open",
            Self::OpcFileEagerOnePartAtomicSave => "opc_file_eager_one_part_atomic_save",
            Self::OpcFileSourceOnePartAtomicSave => "opc_file_source_one_part_atomic_save",
            Self::CfbFileSameLengthOverlayAtomicSave => "cfb_file_same_length_overlay_atomic_save",
            Self::DocxSourceBackedOneEditSave => "docx_source_backed_one_edit_save",
            Self::PptxSourceBackedOneEditSave => "pptx_source_backed_one_edit_save",
            Self::PptxEagerBatchEditSave => "pptx_eager_batch_edit_save",
            Self::PptxSourceBackedBatchEditSave => "pptx_source_backed_batch_edit_save",
            Self::PptxEagerMultiSlideBatchEditSave => "pptx_eager_multi_slide_batch_edit_save",
            Self::PptxSourceBackedMultiSlideBatchEditSave => {
                "pptx_source_backed_multi_slide_batch_edit_save"
            },
            Self::XlsxEagerCalculationMetadataEditSave => {
                "xlsx_eager_calculation_metadata_edit_save"
            },
            Self::XlsxSourceBackedCalculationMetadataEditSave => {
                "xlsx_source_backed_calculation_metadata_edit_save"
            },
            Self::XlsxEagerDefinedNamesEditSave => "xlsx_eager_defined_names_edit_save",
            Self::XlsxSourceBackedDefinedNamesEditSave => {
                "xlsx_source_backed_defined_names_edit_save"
            },
            Self::XlsxEagerPageBreakEditSave => "xlsx_eager_page_break_edit_save",
            Self::XlsxSourceBackedPageBreakEditSave => "xlsx_source_backed_page_break_edit_save",
            Self::XlsxEagerPageMarginEditSave => "xlsx_eager_page_margin_edit_save",
            Self::XlsxSourceBackedPageMarginEditSave => "xlsx_source_backed_page_margin_edit_save",
            Self::XlsxEagerPageSetupEditSave => "xlsx_eager_page_setup_edit_save",
            Self::XlsxSourceBackedPageSetupEditSave => "xlsx_source_backed_page_setup_edit_save",
            Self::XlsxEagerPrintOptionsEditSave => "xlsx_eager_print_options_edit_save",
            Self::XlsxSourceBackedPrintOptionsEditSave => {
                "xlsx_source_backed_print_options_edit_save"
            },
            Self::XlsxEagerSheetProtectionEditSave => "xlsx_eager_sheet_protection_edit_save",
            Self::XlsxSourceBackedSheetProtectionEditSave => {
                "xlsx_source_backed_sheet_protection_edit_save"
            },
            Self::XlsxEagerDataValidationEditSave => "xlsx_eager_data_validation_edit_save",
            Self::XlsxSourceBackedDataValidationEditSave => {
                "xlsx_source_backed_data_validation_edit_save"
            },
            Self::XlsxEagerAutoFilterEditSave => "xlsx_eager_auto_filter_edit_save",
            Self::XlsxSourceBackedAutoFilterEditSave => "xlsx_source_backed_auto_filter_edit_save",
            Self::XlsxEagerConditionalFormattingEditSave => {
                "xlsx_eager_conditional_formatting_edit_save"
            },
            Self::XlsxSourceBackedConditionalFormattingEditSave => {
                "xlsx_source_backed_conditional_formatting_edit_save"
            },
            Self::XlsxEagerMergeCommitSave => "xlsx_eager_merge_commit_save",
            Self::XlsxEagerUnmergeCommitSave => "xlsx_eager_unmerge_commit_save",
            Self::XlsxEagerCellValuesOneEditSave => "xlsx_eager_cell_values_one_edit_save",
            Self::XlsxSourceBackedCellValuesOneEditSave => {
                "xlsx_source_backed_cell_values_one_edit_save"
            },
            Self::XlsxEagerCellValuesOnePercentEditSave => {
                "xlsx_eager_cell_values_one_percent_edit_save"
            },
            Self::XlsxSourceBackedCellValuesOnePercentEditSave => {
                "xlsx_source_backed_cell_values_one_percent_edit_save"
            },
            Self::XlsxEagerCellValuesBatchEditSave => "xlsx_eager_cell_values_batch_edit_save",
            Self::XlsxSourceBackedCellValuesBatchEditSave => {
                "xlsx_source_backed_cell_values_batch_edit_save"
            },
            Self::CfbOpen => "cfb_open",
            Self::CfbListStreams => "cfb_list_streams",
            Self::CfbReadOne => "cfb_read_one",
            Self::CfbCreateStreamBorrowed => "cfb_create_stream_borrowed",
            Self::CfbCreateStreamOwned => "cfb_create_stream_owned",
            Self::OleCommonOpen => "ole_common_open",
            Self::OleCommonPutStreamPublish => "ole_common_put_stream_publish",
            Self::OleCommonFinishRender => "ole_common_finish_render",
            Self::OleCommonOneEditSave => "ole_common_one_edit_save",
            Self::CfbSharedOpen => "cfb_shared_open",
            Self::CfbSharedReadOne => "cfb_shared_read_one",
            Self::CfbSharedConcurrentReads => "cfb_shared_concurrent_reads",
            Self::CfbSelectiveMiniLegacyRead => "cfb_selective_mini_legacy_read",
            Self::CfbSelectiveMiniSharedRead => "cfb_selective_mini_shared_read",
            Self::CfbSelectiveFatLegacyRead => "cfb_selective_fat_legacy_read",
            Self::CfbSelectiveFatSharedRead => "cfb_selective_fat_shared_read",
            Self::DocFreshWriteTo => "doc_fresh_write_to",
            Self::XlsFreshWriteTo => "xls_fresh_write_to",
            Self::PptFreshWriteTo => "ppt_fresh_write_to",
            Self::DocSemanticOpen => "doc_semantic_open",
            Self::DocSemanticListParagraphs => "doc_semantic_list_paragraphs",
            Self::DocSemanticOneParagraph => "doc_semantic_one_paragraph",
            Self::DocSemanticFullText => "doc_semantic_full_text",
            Self::DocSemanticNoopEditSave => "doc_semantic_noop_edit_save",
            Self::DocSemanticOneEditSave => "doc_semantic_one_edit_save",
            Self::DocBodySnapshotListParagraphs => "doc_body_snapshot_list_paragraphs",
            Self::XlsSemanticOpen => "xls_semantic_open",
            Self::XlsSemanticListWorksheets => "xls_semantic_list_worksheets",
            Self::XlsSemanticOneCell => "xls_semantic_one_cell",
            Self::XlsSemanticFullCellScan => "xls_semantic_full_cell_scan",
            Self::XlsSemanticNoopEditSave => "xls_semantic_noop_edit_save",
            Self::XlsSemanticOneEditSave => "xls_semantic_one_edit_save",
            Self::XlsValidationReport => "xls_validation_report",
            Self::XlsCommentsEagerEditSave => "xls_comments_eager_edit_save",
            Self::XlsCommentsSourceBackedEditSave => "xls_comments_source_backed_edit_save",
            Self::XlsCommentsEagerBatchEditSave => "xls_comments_eager_batch_edit_save",
            Self::XlsCommentsSourceBackedBatchEditSave => {
                "xls_comments_source_backed_batch_edit_save"
            },
            Self::XlsVisibilityEagerEditSave => "xls_visibility_eager_edit_save",
            Self::XlsVisibilitySourceBackedEditSave => "xls_visibility_source_backed_edit_save",
            Self::XlsVisibilityEagerBatchEditSave => "xls_visibility_eager_batch_edit_save",
            Self::XlsVisibilitySourceBackedBatchEditSave => {
                "xls_visibility_source_backed_batch_edit_save"
            },
            Self::PptSemanticOpen => "ppt_semantic_open",
            Self::PptSemanticListSlides => "ppt_semantic_list_slides",
            Self::PptSemanticOneShapeText => "ppt_semantic_one_shape_text",
            Self::PptSemanticFullText => "ppt_semantic_full_text",
            Self::PptSlideOrderSnapshotOpen => "ppt_slide_order_snapshot_open",
            Self::PptTextEditOneEditSave => "ppt_text_edit_one_edit_save",
            Self::PptSemanticNoopEditSave => "ppt_semantic_noop_edit_save",
            Self::PptSemanticOneEditSave => "ppt_semantic_one_edit_save",
            Self::XlsxOpenOwned => "xlsx_open_owned",
            Self::XlsxListSheets => "xlsx_list_sheets",
            Self::XlsxFirstCell => "xlsx_first_cell",
            Self::XlsxFullCellScan => "xlsx_full_cell_scan",
            Self::XlsxNarrowColumnRangeScan => "xlsx_narrow_column_range_scan",
            Self::XlsxNoopCommit => "xlsx_noop_commit",
            Self::XlsxNoopCommitSave => "xlsx_noop_commit_save",
            Self::XlsxOneCellCommit => "xlsx_one_cell_commit",
            Self::XlsxOneCellCommitFirstRead => "xlsx_one_cell_commit_first_read",
            Self::XlsxOneCellCommitSave => "xlsx_one_cell_commit_save",
            Self::XlsxOnePercentCommit => "xlsx_one_percent_commit",
            Self::XlsxOnePercentCommitSave => "xlsx_one_percent_commit_save",
            Self::XlsxSourceOpen => "xlsx_source_open",
            Self::XlsxSourceListSheets => "xlsx_source_list_sheets",
            Self::XlsxSourceFirstCell => "xlsx_source_first_cell",
            Self::XlsxSourceNarrowColumnRangeScan => "xlsx_source_narrow_column_range_scan",
            Self::XlsxStreamingCreate => "xlsx_streaming_create",
            Self::OpcRangeSourceOpen => "opc_range_source_open",
            Self::OpcRangeSourceOpenMainRead => "opc_range_source_open_main_read",
            Self::XlsxRangeSourceOpen => "xlsx_range_source_open",
            Self::XlsxRangeSourceListSheets => "xlsx_range_source_list_sheets",
            Self::XlsxRangeSourceFirstCell => "xlsx_range_source_first_cell",
            Self::XlsxRangeSourceNarrowColumnRangeScan => {
                "xlsx_range_source_narrow_column_range_scan"
            },
            Self::OpcOpenSessionScaling => "opc_open_session_scaling",
            Self::CfbBulkReadScaling => "cfb_bulk_read_scaling",
            Self::RtfSemanticOpen => "rtf_semantic_open",
            Self::RtfSemanticParagraphCount => "rtf_semantic_paragraph_count",
            Self::RtfSemanticListParagraphs => "rtf_semantic_list_paragraphs",
            Self::RtfSemanticCollectParagraphs => "rtf_semantic_collect_paragraphs",
            Self::RtfSemanticOneParagraph => "rtf_semantic_one_paragraph",
            Self::RtfSemanticFullText => "rtf_semantic_full_text",
            Self::RtfSemanticTextToSink => "rtf_semantic_text_to_sink",
            Self::RtfSemanticStreamSave => "rtf_semantic_stream_save",
            Self::RtfSemanticNoopEditSave => "rtf_semantic_noop_edit_save",
            Self::RtfSemanticOneEditSave => "rtf_semantic_one_edit_save",
            Self::RtfSemanticOnePercentEditSave => "rtf_semantic_one_percent_edit_save",
            Self::RtfSemanticRemoveParagraphSave => "rtf_semantic_remove_paragraph_save",
            Self::RtfSemanticMoveParagraphSave => "rtf_semantic_move_paragraph_save",
            Self::RtfLogicalTailAppend => "rtf_logical_tail_append",
            Self::RtfLogicalTailNoopSave => "rtf_logical_tail_noop_save",
            Self::RtfValidationReport => "rtf_validation_report",
            Self::RtfStreamingCreate => "rtf_streaming_create",
            Self::DocxSemanticOpen => "docx_semantic_open",
            Self::DocxSemanticListParagraphs => "docx_semantic_list_paragraphs",
            Self::DocxSemanticOneParagraph => "docx_semantic_one_paragraph",
            Self::DocxSemanticFullText => "docx_semantic_full_text",
            Self::DocxSemanticCreateSmall => "docx_semantic_create_small",
            Self::DocxSemanticNoopEditSave => "docx_semantic_noop_edit_save",
            Self::DocxSemanticOneEditSave => "docx_semantic_one_edit_save",
            Self::DocxSemanticOnePercentEditSave => "docx_semantic_one_percent_edit_save",
            Self::DocxValidationReport => "docx_validation_report",
            Self::DocxSectionInventory => "docx_section_inventory",
            Self::PptxSemanticOpen => "pptx_semantic_open",
            Self::PptxSemanticListSlides => "pptx_semantic_list_slides",
            Self::PptxSemanticOneSlide => "pptx_semantic_one_slide",
            Self::PptxSemanticFullText => "pptx_semantic_full_text",
            Self::PptxSemanticCreateSmall => "pptx_semantic_create_small",
            Self::PptxSemanticNoopEditSave => "pptx_semantic_noop_edit_save",
            Self::PptxSemanticOneEditSave => "pptx_semantic_one_edit_save",
            Self::PptxSemanticOnePercentEditSave => "pptx_semantic_one_percent_edit_save",
            Self::PptxValidationReport => "pptx_validation_report",
            Self::OdtSemanticOpen => "odt_semantic_open",
            Self::OdtSemanticListParagraphs => "odt_semantic_list_paragraphs",
            Self::OdtSemanticOneParagraph => "odt_semantic_one_paragraph",
            Self::OdtSemanticFullText => "odt_semantic_full_text",
            Self::OdtSemanticCreateSmall => "odt_semantic_create_small",
            Self::OdtSemanticNoopEditSave => "odt_semantic_noop_edit_save",
            Self::OdtSemanticOneEditSave => "odt_semantic_one_edit_save",
            Self::OdtSemanticOnePercentEditSave => "odt_semantic_one_percent_edit_save",
            Self::OdfValidationReport => "odf_validation_report",
            Self::OdfMimetypeRepairPlan => "odf_mimetype_repair_plan",
            Self::OdtMediaParagraphEditSave => "odt_media_paragraph_edit_save",
            Self::OdtMediaLineBreakEditSave => "odt_media_line_break_edit_save",
            Self::OdtMediaAppendRunEditSave => "odt_media_append_run_edit_save",
            Self::OdtMediaAppendHyperlinkEditSave => "odt_media_append_hyperlink_edit_save",
            Self::OdtMediaInsertParagraphEditSave => "odt_media_insert_paragraph_edit_save",
            Self::OdtMediaRemoveParagraphEditSave => "odt_media_remove_paragraph_edit_save",
            Self::OdtEmbeddedResourceScalarReplaceSave => {
                "odt_embedded_resource_scalar_replace_save"
            },
            Self::OdtEmbeddedResourceBatchReplaceSave => "odt_embedded_resource_batch_replace_save",
            Self::OdsSemanticOpen => "ods_semantic_open",
            Self::OdsSemanticListSheets => "ods_semantic_list_sheets",
            Self::OdsSemanticOneCell => "ods_semantic_one_cell",
            Self::OdsSemanticCellSweep => "ods_semantic_cell_sweep",
            Self::OdsSemanticFullCellText => "ods_semantic_full_cell_text",
            Self::OdsSemanticCreateSmall => "ods_semantic_create_small",
            Self::OdsSemanticNoopEditSave => "ods_semantic_noop_edit_save",
            Self::OdsSemanticOneEditSave => "ods_semantic_one_edit_save",
            Self::OdsSemanticOnePercentEditSave => "ods_semantic_one_percent_edit_save",
            Self::OdsMediaOneEditSave => "ods_media_one_edit_save",
            Self::OdpSemanticOpen => "odp_semantic_open",
            Self::OdpSemanticListSlides => "odp_semantic_list_slides",
            Self::OdpSemanticOneSlide => "odp_semantic_one_slide",
            Self::OdpSemanticFullText => "odp_semantic_full_text",
            Self::OdpSemanticCreateSmall => "odp_semantic_create_small",
            Self::OdpSemanticNoopEditSave => "odp_semantic_noop_edit_save",
            Self::OdpSemanticOneEditSave => "odp_semantic_one_edit_save",
            Self::OdpMediaTextBoxEditSave => "odp_media_textbox_edit_save",
            Self::OdpMediaTextBoxScalarReplaceSave => "odp_media_textbox_scalar_replace_save",
            Self::OdpMediaTextBoxBatchReplaceSave => "odp_media_textbox_batch_replace_save",
        }
    }

    const fn uses_synthetic_cfb(self) -> bool {
        matches!(
            self,
            Self::CfbOpen
                | Self::CfbListStreams
                | Self::CfbReadOne
                | Self::CfbCreateStreamBorrowed
                | Self::CfbCreateStreamOwned
                | Self::OleCommonOpen
                | Self::OleCommonPutStreamPublish
                | Self::OleCommonFinishRender
                | Self::OleCommonOneEditSave
                | Self::CfbSharedOpen
                | Self::CfbSharedReadOne
                | Self::CfbSharedConcurrentReads
                | Self::CfbSelectiveMiniLegacyRead
                | Self::CfbSelectiveMiniSharedRead
                | Self::CfbSelectiveFatLegacyRead
                | Self::CfbSelectiveFatSharedRead
                | Self::CfbBulkReadScaling
        )
    }

    const fn uses_synthetic_opc(self) -> bool {
        matches!(
            self,
            Self::ZipIndex
                | Self::ZipReadOne
                | Self::OpcOpen
                | Self::OpcOpenOwned
                | Self::OpcNoopSave
                | Self::OpcMutatedSave
                | Self::OpcSourceOpen
                | Self::OpcSourceOpenMainRead
                | Self::OpcSourceCachedMainRead
                | Self::OpcSourceConcurrentSamePart
                | Self::OpcRangeSourceOpen
                | Self::OpcRangeSourceOpenMainRead
                | Self::OpcOpenSessionScaling
        )
    }

    const fn is_fresh_writer(self) -> bool {
        matches!(
            self,
            Self::DocFreshWriteTo | Self::XlsFreshWriteTo | Self::PptFreshWriteTo
        )
    }

    const fn uses_semantic_doc(self) -> bool {
        matches!(
            self,
            Self::DocSemanticOpen
                | Self::DocSemanticListParagraphs
                | Self::DocSemanticOneParagraph
                | Self::DocSemanticFullText
                | Self::DocSemanticNoopEditSave
                | Self::DocSemanticOneEditSave
                | Self::DocBodySnapshotListParagraphs
        )
    }

    const fn uses_semantic_xls(self) -> bool {
        matches!(
            self,
            Self::XlsSemanticOpen
                | Self::XlsSemanticListWorksheets
                | Self::XlsSemanticOneCell
                | Self::XlsSemanticFullCellScan
                | Self::XlsSemanticNoopEditSave
                | Self::XlsSemanticOneEditSave
        )
    }

    const fn is_xls_comments_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsCommentsEagerEditSave
                | Self::XlsCommentsSourceBackedEditSave
                | Self::XlsCommentsEagerBatchEditSave
                | Self::XlsCommentsSourceBackedBatchEditSave
        )
    }

    const fn is_xls_visibility_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsVisibilityEagerEditSave
                | Self::XlsVisibilitySourceBackedEditSave
                | Self::XlsVisibilityEagerBatchEditSave
                | Self::XlsVisibilitySourceBackedBatchEditSave
        )
    }

    const fn uses_semantic_ppt(self) -> bool {
        matches!(
            self,
            Self::PptSemanticOpen
                | Self::PptSemanticListSlides
                | Self::PptSemanticOneShapeText
                | Self::PptSemanticFullText
                | Self::PptSlideOrderSnapshotOpen
                | Self::PptTextEditOneEditSave
                | Self::PptSemanticNoopEditSave
                | Self::PptSemanticOneEditSave
        )
    }

    const fn uses_xlsx(self) -> bool {
        matches!(
            self,
            Self::XlsxOpenOwned
                | Self::XlsxListSheets
                | Self::XlsxFirstCell
                | Self::XlsxFullCellScan
                | Self::XlsxNarrowColumnRangeScan
                | Self::XlsxNoopCommit
                | Self::XlsxNoopCommitSave
                | Self::XlsxOneCellCommit
                | Self::XlsxOneCellCommitFirstRead
                | Self::XlsxOneCellCommitSave
                | Self::XlsxOnePercentCommit
                | Self::XlsxOnePercentCommitSave
                | Self::XlsxSourceOpen
                | Self::XlsxSourceListSheets
                | Self::XlsxSourceFirstCell
                | Self::XlsxSourceNarrowColumnRangeScan
                | Self::XlsxRangeSourceOpen
                | Self::XlsxRangeSourceListSheets
                | Self::XlsxRangeSourceFirstCell
                | Self::XlsxRangeSourceNarrowColumnRangeScan
        )
    }

    const fn uses_xlsx_cell_values(self) -> bool {
        self.is_xlsx_cell_values_edit_save()
    }

    const fn uses_streaming_creation(self) -> bool {
        matches!(self, Self::XlsxStreamingCreate | Self::RtfStreamingCreate)
    }

    const fn uses_semantic_docx(self) -> bool {
        matches!(
            self,
            Self::DocxSemanticOpen
                | Self::DocxSemanticListParagraphs
                | Self::DocxSemanticOneParagraph
                | Self::DocxSemanticFullText
                | Self::DocxSemanticCreateSmall
                | Self::DocxSemanticNoopEditSave
                | Self::DocxSemanticOneEditSave
                | Self::DocxSemanticOnePercentEditSave
        )
    }

    const fn uses_semantic_rtf(self) -> bool {
        matches!(
            self,
            Self::RtfSemanticOpen
                | Self::RtfSemanticParagraphCount
                | Self::RtfSemanticListParagraphs
                | Self::RtfSemanticCollectParagraphs
                | Self::RtfSemanticOneParagraph
                | Self::RtfSemanticFullText
                | Self::RtfSemanticTextToSink
                | Self::RtfSemanticStreamSave
                | Self::RtfSemanticNoopEditSave
                | Self::RtfSemanticOneEditSave
                | Self::RtfSemanticOnePercentEditSave
                | Self::RtfSemanticRemoveParagraphSave
                | Self::RtfSemanticMoveParagraphSave
                | Self::RtfLogicalTailAppend
                | Self::RtfLogicalTailNoopSave
        )
    }

    const fn is_rtf_lifecycle(self) -> bool {
        matches!(
            self,
            Self::RtfSemanticRemoveParagraphSave | Self::RtfSemanticMoveParagraphSave
        )
    }

    const fn is_rtf_logical_tail(self) -> bool {
        matches!(
            self,
            Self::RtfLogicalTailAppend | Self::RtfLogicalTailNoopSave
        )
    }

    const fn uses_semantic_pptx(self) -> bool {
        matches!(
            self,
            Self::PptxSemanticOpen
                | Self::PptxSemanticListSlides
                | Self::PptxSemanticOneSlide
                | Self::PptxSemanticFullText
                | Self::PptxSemanticCreateSmall
                | Self::PptxSemanticNoopEditSave
                | Self::PptxSemanticOneEditSave
                | Self::PptxSemanticOnePercentEditSave
        )
    }

    const fn uses_semantic_odt(self) -> bool {
        matches!(
            self,
            Self::OdtSemanticOpen
                | Self::OdtSemanticListParagraphs
                | Self::OdtSemanticOneParagraph
                | Self::OdtSemanticFullText
                | Self::OdtSemanticCreateSmall
                | Self::OdtSemanticNoopEditSave
                | Self::OdtSemanticOneEditSave
                | Self::OdtSemanticOnePercentEditSave
        )
    }

    const fn uses_odt_media(self) -> bool {
        matches!(
            self,
            Self::OdtMediaParagraphEditSave
                | Self::OdtMediaLineBreakEditSave
                | Self::OdtMediaAppendRunEditSave
                | Self::OdtMediaAppendHyperlinkEditSave
                | Self::OdtMediaInsertParagraphEditSave
                | Self::OdtMediaRemoveParagraphEditSave
        )
    }

    const fn uses_odt_resource_batch(self) -> bool {
        matches!(
            self,
            Self::OdtEmbeddedResourceScalarReplaceSave | Self::OdtEmbeddedResourceBatchReplaceSave
        )
    }

    const fn uses_semantic_ods(self) -> bool {
        matches!(
            self,
            Self::OdsSemanticOpen
                | Self::OdsSemanticListSheets
                | Self::OdsSemanticOneCell
                | Self::OdsSemanticCellSweep
                | Self::OdsSemanticFullCellText
                | Self::OdsSemanticCreateSmall
                | Self::OdsSemanticNoopEditSave
                | Self::OdsSemanticOneEditSave
                | Self::OdsSemanticOnePercentEditSave
        )
    }

    const fn uses_semantic_odp(self) -> bool {
        matches!(
            self,
            Self::OdpSemanticOpen
                | Self::OdpSemanticListSlides
                | Self::OdpSemanticOneSlide
                | Self::OdpSemanticFullText
                | Self::OdpSemanticCreateSmall
                | Self::OdpSemanticNoopEditSave
                | Self::OdpSemanticOneEditSave
        )
    }

    const fn uses_odp_media(self) -> bool {
        matches!(self, Self::OdpMediaTextBoxEditSave)
    }

    const fn uses_odp_text_box_batch(self) -> bool {
        matches!(
            self,
            Self::OdpMediaTextBoxScalarReplaceSave | Self::OdpMediaTextBoxBatchReplaceSave
        )
    }

    const fn uses_validation_xls(self) -> bool {
        matches!(self, Self::XlsValidationReport)
    }

    const fn uses_validation_rtf(self) -> bool {
        matches!(self, Self::RtfValidationReport)
    }

    const fn uses_validation_docx(self) -> bool {
        matches!(
            self,
            Self::DocxValidationReport | Self::DocxSectionInventory
        )
    }

    const fn uses_validation_pptx(self) -> bool {
        matches!(self, Self::PptxValidationReport)
    }

    const fn uses_validation_odf(self) -> bool {
        matches!(self, Self::OdfValidationReport)
    }

    const fn uses_odf_repair(self) -> bool {
        matches!(self, Self::OdfMimetypeRepairPlan)
    }

    const fn uses_ods_media(self) -> bool {
        matches!(self, Self::OdsMediaOneEditSave)
    }

    const fn is_semantic_create_small(self) -> bool {
        matches!(
            self,
            Self::DocxSemanticCreateSmall
                | Self::PptxSemanticCreateSmall
                | Self::OdtSemanticCreateSmall
                | Self::OdsSemanticCreateSmall
                | Self::OdpSemanticCreateSmall
        )
    }

    const fn is_scaling(self) -> bool {
        matches!(self, Self::OpcOpenSessionScaling | Self::CfbBulkReadScaling)
    }

    const fn is_cfb_selective(self) -> bool {
        matches!(
            self,
            Self::CfbSelectiveMiniLegacyRead
                | Self::CfbSelectiveMiniSharedRead
                | Self::CfbSelectiveFatLegacyRead
                | Self::CfbSelectiveFatSharedRead
        )
    }

    const fn is_opc_source_cache_evidence(self) -> bool {
        matches!(
            self,
            Self::OpcSourceCacheBudgetBoundary
                | Self::OpcSourceCacheControlContention
                | Self::OpcSourceCacheManagedContention
        )
    }

    const fn is_opc_source_overlay_save(self) -> bool {
        matches!(self, Self::OpcSourceOverlayOnePartSave)
    }

    const fn is_filesystem(self) -> bool {
        matches!(
            self,
            Self::OpcFileEagerOpen
                | Self::OpcFileSourceOpen
                | Self::OpcFileEagerOnePartAtomicSave
                | Self::OpcFileSourceOnePartAtomicSave
                | Self::CfbFileSameLengthOverlayAtomicSave
        )
    }

    const fn is_docx_source_edit_save(self) -> bool {
        matches!(self, Self::DocxSourceBackedOneEditSave)
    }

    const fn is_pptx_source_edit_save(self) -> bool {
        matches!(
            self,
            Self::PptxSourceBackedOneEditSave
                | Self::PptxEagerBatchEditSave
                | Self::PptxSourceBackedBatchEditSave
                | Self::PptxEagerMultiSlideBatchEditSave
                | Self::PptxSourceBackedMultiSlideBatchEditSave
        )
    }

    const fn is_xlsx_calculation_metadata_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerCalculationMetadataEditSave
                | Self::XlsxSourceBackedCalculationMetadataEditSave
        )
    }

    const fn is_xlsx_defined_names_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerDefinedNamesEditSave | Self::XlsxSourceBackedDefinedNamesEditSave
        )
    }

    const fn is_xlsx_page_break_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerPageBreakEditSave | Self::XlsxSourceBackedPageBreakEditSave
        )
    }

    const fn is_xlsx_page_margin_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerPageMarginEditSave | Self::XlsxSourceBackedPageMarginEditSave
        )
    }

    const fn is_xlsx_page_setup_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerPageSetupEditSave | Self::XlsxSourceBackedPageSetupEditSave
        )
    }

    const fn is_xlsx_print_options_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerPrintOptionsEditSave | Self::XlsxSourceBackedPrintOptionsEditSave
        )
    }

    const fn is_xlsx_sheet_protection_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerSheetProtectionEditSave | Self::XlsxSourceBackedSheetProtectionEditSave
        )
    }

    const fn is_xlsx_data_validation_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerDataValidationEditSave | Self::XlsxSourceBackedDataValidationEditSave
        )
    }

    const fn is_xlsx_auto_filter_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerAutoFilterEditSave | Self::XlsxSourceBackedAutoFilterEditSave
        )
    }

    const fn is_xlsx_conditional_formatting_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerConditionalFormattingEditSave
                | Self::XlsxSourceBackedConditionalFormattingEditSave
        )
    }

    const fn is_xlsx_merge_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerMergeCommitSave | Self::XlsxEagerUnmergeCommitSave
        )
    }

    const fn is_xlsx_cell_values_edit_save(self) -> bool {
        matches!(
            self,
            Self::XlsxEagerCellValuesOneEditSave
                | Self::XlsxSourceBackedCellValuesOneEditSave
                | Self::XlsxEagerCellValuesOnePercentEditSave
                | Self::XlsxSourceBackedCellValuesOnePercentEditSave
                | Self::XlsxEagerCellValuesBatchEditSave
                | Self::XlsxSourceBackedCellValuesBatchEditSave
        )
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
struct RangeSimulationConfig {
    fixed_latency_us: u64,
    request_overhead_us: u64,
    bandwidth_bytes_per_second: u64,
    max_physical_range_bytes: usize,
}

impl Default for RangeSimulationConfig {
    fn default() -> Self {
        Self {
            fixed_latency_us: DEFAULT_RANGE_FIXED_LATENCY_US,
            request_overhead_us: DEFAULT_RANGE_REQUEST_OVERHEAD_US,
            bandwidth_bytes_per_second: DEFAULT_RANGE_BANDWIDTH_BYTES_PER_SECOND,
            max_physical_range_bytes: DEFAULT_RANGE_MAX_PHYSICAL_BYTES,
        }
    }
}

#[derive(Debug)]
struct Options {
    samples: usize,
    warmup_iterations: usize,
    filesystem_cache: filesystem::CacheSelection,
    filesystem_root: Option<PathBuf>,
    cases: Vec<Case>,
    shapes: Vec<CorpusShape>,
    payloads: Vec<PayloadKind>,
    writer_shapes: Vec<WriterShape>,
    xlsx_shapes: Vec<XlsxShape>,
    xlsx_cell_crud_shapes: Vec<XlsxCellCrudShape>,
    semantic_shapes: Vec<SemanticShape>,
    rtf_variants: Vec<RtfSemanticVariant>,
    range_simulation: RangeSimulationConfig,
    execution_workers: Vec<usize>,
    output: Option<PathBuf>,
}

#[derive(Debug)]
struct Corpus {
    manifest: CorpusManifest,
    archive: Vec<u8>,
    target_name: String,
    target_payload: Vec<u8>,
    xlsx: Option<XlsxCorpus>,
}

#[derive(Clone, Debug)]
struct StreamingCorpus {
    manifest: CorpusManifest,
    shape: SemanticShape,
    metrics: StreamingMetrics,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct StreamingMetrics {
    rows: u64,
    cells: u64,
    paragraphs: u64,
    runs: u64,
    input_bytes: u64,
    authored_part_bytes: u64,
    retained_authoring_window_bytes: u64,
}

#[derive(Clone, Debug)]
struct XlsxCorpus {
    sheet_count: usize,
    row_count: usize,
    column_count: usize,
    one_percent_updates: Vec<XlsxCoordinate>,
    /// Optional sparse coordinate inventory used by the media-rich CRUD
    /// matrix. `None` preserves the historical rectangular corpus behavior.
    cell_inventory: Option<Vec<Vec<XlsxCoordinate>>>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct XlsxCoordinate {
    sheet: usize,
    row: usize,
    column: usize,
}

#[derive(Serialize)]
struct Report {
    schema_version: u32,
    tool: Tool,
    environment: Environment,
    configuration: Configuration,
    results: Vec<CaseResult>,
    #[serde(skip_serializing_if = "Option::is_none")]
    filesystem_evidence: Option<Vec<filesystem::Evidence>>,
}

#[derive(Serialize)]
struct Tool {
    name: &'static str,
    version: &'static str,
    profile: &'static str,
    target_os: &'static str,
    target_arch: &'static str,
}

#[derive(Serialize)]
struct Environment {
    rustc_version: Option<String>,
    git_revision: Option<String>,
    git_worktree_dirty: Option<bool>,
    logical_cpus_available: usize,
    allocator: &'static str,
    rustflags: Option<String>,
    cargo_build_target: Option<String>,
    perf_event_paranoid: Option<String>,
    os: Option<String>,
    kernel: Option<String>,
    cpu_model: Option<String>,
    total_memory_bytes: Option<u64>,
    page_size_bytes: Option<u64>,
    filesystem_type: Option<String>,
    source_destination_same_device: Option<bool>,
    cpu_affinity: Option<String>,
    storage_identifier: Option<String>,
}

#[derive(Serialize)]
struct Configuration {
    samples_per_case: usize,
    warmup_iterations_per_case: usize,
    filesystem_cache_states: Vec<&'static str>,
    filesystem_fresh_child_per_sample: bool,
    filesystem_process_isolated: bool,
    filesystem_root_selected: bool,
    cases: Vec<&'static str>,
    corpus_shapes: Vec<&'static str>,
    payload_kinds: Vec<&'static str>,
    writer_shapes: Vec<&'static str>,
    xlsx_shapes: Vec<&'static str>,
    xlsx_cell_crud_shapes: Vec<&'static str>,
    semantic_shapes: Vec<&'static str>,
    rtf_variants: Vec<&'static str>,
    range_simulation: RangeSimulationConfig,
    execution_workers: Vec<usize>,
}

#[derive(Clone, Debug, Serialize)]
struct CorpusManifest {
    name: String,
    generator: &'static str,
    package_format: &'static str,
    shape: &'static str,
    payload_kind: &'static str,
    compression: &'static str,
    entry_count: usize,
    archive_member_count: usize,
    entry_bytes: usize,
    uncompressed_payload_bytes: usize,
    archive_bytes: usize,
    archive_sha256: String,
    target_entry: String,
    target_payload_bytes: usize,
    target_payload_sha256: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    rtf_variant: Option<&'static str>,
    xlsx: Option<XlsxManifest>,
}

#[derive(Clone, Debug, Serialize)]
struct XlsxManifest {
    sheet_count: usize,
    rows_per_sheet: usize,
    columns_per_sheet: usize,
    one_percent_update_count: usize,
    source_members: XlsxSourceMembersManifest,
}

#[derive(Clone, Debug, Serialize)]
struct XlsxSourceMembersManifest {
    workbook: String,
    worksheets: Vec<String>,
    shared_strings: Option<String>,
    styles: Option<String>,
}

#[derive(Serialize)]
struct CaseResult {
    case: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    cache_state: Option<&'static str>,
    corpus: CorpusManifest,
    elapsed_ns: Statistics,
    sink: Option<SinkSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    source: Option<SourceSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    execution: Option<ExecutionSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    output_sha256: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
struct CfbSelectiveImplementationEvidence {
    implementation: &'static str,
    open_ns: Vec<u64>,
    read_ns: Vec<u64>,
    total_ns: Vec<u64>,
    open_read_calls: Vec<u64>,
    open_read_bytes: Vec<u64>,
    open_range_sizes: Vec<Vec<u64>>,
    read_calls: Vec<u64>,
    read_bytes: Vec<u64>,
    read_range_sizes: Vec<Vec<u64>>,
    returned_payload_bytes: Vec<u64>,
    selected_payload_sha256: String,
}

#[derive(Clone, Debug, Serialize)]
struct CfbSelectiveEvidence {
    timing_scope: &'static str,
    sink: &'static str,
    selected_target_kind: &'static str,
    legacy_or_positional: CfbSelectiveImplementationEvidence,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct ValidationSummary {
    report_sha256: String,
    check_ids: Vec<String>,
    check_statuses: Vec<String>,
    issue_codes: Vec<String>,
    issue_count: usize,
    complete: bool,
    has_errors: bool,
    counts: BTreeMap<String, u64>,
    source_sha256_before: String,
    source_sha256_after: String,
    source_bytes: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    source_read_calls: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    source_read_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    section_inventory: Option<SectionInventorySummary>,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct SectionInventorySummary {
    section_count: usize,
    paragraph_count: usize,
    descriptors: Vec<SectionDescriptorSummary>,
}

#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
struct SectionDescriptorSummary {
    position: usize,
    ownership: String,
    paragraph_start: usize,
    paragraph_end: usize,
    page_width_emu: Option<i64>,
    page_height_emu: Option<i64>,
    page_orientation: String,
    margin_left_emu: Option<i64>,
    margin_right_emu: Option<i64>,
    margin_top_emu: Option<i64>,
    margin_bottom_emu: Option<i64>,
    margin_header_emu: Option<i64>,
    margin_footer_emu: Option<i64>,
    margin_gutter_emu: Option<i64>,
    start: Option<String>,
    headers: Vec<String>,
    footers: Vec<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
struct ExecutionSummary {
    worker_count: usize,
    logical_tasks: usize,
    logical_bytes: u64,
}

#[derive(Clone, Debug, Default, Serialize)]
struct SourceSummary {
    read_calls: Vec<u64>,
    read_bytes: Vec<u64>,
    ordinary_payload_read_calls: Vec<u64>,
    ordinary_payload_read_bytes: Vec<u64>,
    max_in_flight_reads: Vec<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    ordinary_payload_materializations: Option<Vec<u64>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    xlsx: Option<XlsxSourceSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    xls_comments: Option<XlsCommentsSourceSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    xls_visibility: Option<XlsVisibilitySourceSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    simulation: Option<RangeSimulationSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    opc_cache: Option<OpcCacheEvidenceSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    validation: Option<ValidationSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    odf_repair: Option<OdfRepairSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    cfb_selective: Option<CfbSelectiveEvidence>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize)]
struct OdfRepairSummary {
    schema: &'static str,
    repair_id: &'static str,
    intent: &'static str,
    validation_issue_id: String,
    plan_json_sha256: String,
    source_bytes: u64,
    output_bytes: u64,
    source_sha256: String,
    output_sha256: String,
    member_count: usize,
    extra_field_id: u16,
    extra_field_bytes: u16,
    changed_members: Vec<&'static str>,
    changed_regions: Vec<&'static str>,
    member_payloads_preserved: bool,
    reversible: bool,
    exact_canonical_recovery_verified: bool,
    patch_verified: bool,
    inverse_verified: bool,
    stale_source_refusal_verified: bool,
    canonical_no_plan_verified: bool,
    partial_sink_progress_verified: bool,
}

#[derive(Clone, Debug, Serialize)]
struct OpcCacheEvidenceSummary {
    cache_mode: &'static str,
    scenario: &'static str,
    capacity_ratio: &'static str,
    capacity_entries: usize,
    capacity_bytes: usize,
    working_set_parts: usize,
    working_set_bytes: u64,
    worker_count: usize,
    persistent_worker_teams_created: usize,
    fixed_source_delay_us: u64,
    timing_scope: &'static str,
    diagnostics: OpcCacheDiagnosticsSummary,
    #[serde(skip_serializing_if = "Option::is_none")]
    gate: Option<OpcCacheGateSummary>,
    budget_used_after_handles_drop: Vec<u64>,
    budget_used_after_package_drop: Vec<u64>,
    scaling: OpcCacheScalingSummary,
}

#[derive(Clone, Debug, Default, Serialize)]
struct OpcCacheDiagnosticsSummary {
    hits: Vec<u64>,
    cold_loads: Vec<u64>,
    waiter_joins: Vec<u64>,
    successful_loads: Vec<u64>,
    failed_loads: Vec<u64>,
    evictions: Vec<u64>,
    bypasses: Vec<u64>,
    oversized_bypasses: Vec<u64>,
    allocation_bypasses: Vec<u64>,
    budget_reservation_failures: Vec<u64>,
    retained_entries: Vec<usize>,
    retained_bytes: Vec<usize>,
    in_flight_loads: Vec<usize>,
    budget_memory_used: Vec<u64>,
    budget_cache_reserved_bytes: Vec<u64>,
    budget_memory_limit: Vec<Option<u64>>,
}

#[derive(Clone, Debug, Default, Serialize)]
struct OpcCacheGateSummary {
    initial_arrivals: Vec<u64>,
    delayed_payload_arrivals: Vec<u64>,
    max_concurrent_delays: Vec<u64>,
    pre_release_flights: Vec<usize>,
    pre_release_waiters: Vec<u64>,
}

#[derive(Clone, Debug, Serialize)]
struct OpcCacheScalingSummary {
    model: &'static str,
    classification: &'static str,
    baseline_worker_count: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    p50_speedup: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    p50_efficiency: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    amdahl_serial_fraction: Option<f64>,
    p50_requests_per_second: f64,
    relative_request_throughput: f64,
}

#[derive(Clone, Debug, Default, Serialize)]
struct XlsCommentsSourceSummary {
    source_counter_scope: &'static str,
    source_backed: bool,
    update_count: usize,
    semantic_staging_plan_ns: Vec<u64>,
    publication_ns: Vec<u64>,
    changed_comments: Vec<usize>,
    touched_streams: Vec<usize>,
    source_bytes: Vec<u64>,
    source_workbook_bytes: Vec<u64>,
    target_workbook_bytes: Vec<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    splice_count: Option<Vec<usize>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    replacement_bytes: Option<Vec<u64>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    changed_spans: Option<Vec<usize>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    source_fingerprints: Option<Vec<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    target_fingerprints: Option<Vec<String>>,
}

#[derive(Clone, Debug)]
struct XlsCommentsIterationEvidence {
    source_backed: bool,
    update_count: usize,
    semantic_staging_plan_ns: u64,
    publication_ns: u64,
    changed_comments: usize,
    touched_streams: usize,
    source_bytes: u64,
    source_workbook_bytes: u64,
    target_workbook_bytes: u64,
    splice_count: Option<usize>,
    replacement_bytes: Option<u64>,
    changed_spans: Option<usize>,
    source_fingerprint: Option<String>,
    target_fingerprint: Option<String>,
}

#[derive(Clone, Debug, Default, Serialize)]
struct XlsVisibilitySourceSummary {
    source_counter_scope: &'static str,
    source_backed: bool,
    update_count: usize,
    semantic_staging_plan_ns: Vec<u64>,
    publication_ns: Vec<u64>,
    changed_worksheets: Vec<usize>,
    touched_streams: Vec<usize>,
    source_bytes: Vec<u64>,
    source_workbook_bytes: Vec<u64>,
    target_workbook_bytes: Vec<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    splice_count: Option<Vec<usize>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    replacement_bytes: Option<Vec<u64>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    changed_spans: Option<Vec<usize>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    source_fingerprints: Option<Vec<String>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    target_fingerprints: Option<Vec<String>>,
}

#[derive(Clone, Debug)]
struct XlsVisibilityIterationEvidence {
    source_backed: bool,
    update_count: usize,
    semantic_staging_plan_ns: u64,
    publication_ns: u64,
    changed_worksheets: usize,
    touched_streams: usize,
    source_bytes: u64,
    source_workbook_bytes: u64,
    target_workbook_bytes: u64,
    splice_count: Option<usize>,
    replacement_bytes: Option<u64>,
    changed_spans: Option<usize>,
    source_fingerprint: Option<String>,
    target_fingerprint: Option<String>,
}

#[derive(Clone, Debug, Default, Serialize)]
struct RangeSimulationSummary {
    logical_read_calls: Vec<u64>,
    logical_read_bytes: Vec<u64>,
    physical_request_count: Vec<u64>,
    physical_request_bytes: Vec<u64>,
    physical_request_sizes: Vec<Vec<u64>>,
    physical_request_size_buckets: Vec<RequestSizeBuckets>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize)]
struct RequestSizeBuckets {
    bytes_1_to_512: u64,
    bytes_513_to_4096: u64,
    bytes_4097_to_16384: u64,
    bytes_16385_to_65536: u64,
    bytes_over_65536: u64,
}

#[derive(Clone, Debug, Default, Serialize)]
struct XlsxSourceSummary {
    workbook_read_calls: Vec<u64>,
    workbook_read_bytes: Vec<u64>,
    selected_worksheet_read_calls: Vec<u64>,
    selected_worksheet_read_bytes: Vec<u64>,
    unselected_worksheet_read_calls: Vec<u64>,
    unselected_worksheet_read_bytes: Vec<u64>,
    shared_strings_read_calls: Vec<u64>,
    shared_strings_read_bytes: Vec<u64>,
    styles_read_calls: Vec<u64>,
    styles_read_bytes: Vec<u64>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct RangeSnapshot {
    read_calls: u64,
    read_bytes: u64,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct XlsxSourceSnapshot {
    workbook: RangeSnapshot,
    selected_worksheet: RangeSnapshot,
    unselected_worksheets: RangeSnapshot,
    shared_strings: RangeSnapshot,
    styles: RangeSnapshot,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct SourceSnapshot {
    read_calls: u64,
    read_bytes: u64,
    ordinary_payload_read_calls: u64,
    ordinary_payload_read_bytes: u64,
    max_in_flight_reads: u64,
    xlsx: XlsxSourceSnapshot,
}

#[derive(Clone, Debug, Default)]
struct XlsxTrackedRanges {
    workbook: Vec<Range<u64>>,
    selected_worksheet: Vec<Range<u64>>,
    unselected_worksheets: Vec<Range<u64>>,
    shared_strings: Vec<Range<u64>>,
    styles: Vec<Range<u64>>,
}

#[derive(Debug, Default)]
struct AtomicRangeCounter {
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
}

#[derive(Debug)]
struct InstrumentedSource {
    bytes: Arc<Vec<u8>>,
    version: SourceVersion,
    ordinary_payload_ranges: Vec<Range<u64>>,
    xlsx_ranges: XlsxTrackedRanges,
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
    ordinary_payload_read_calls: AtomicU64,
    ordinary_payload_read_bytes: AtomicU64,
    in_flight_reads: AtomicU64,
    max_in_flight_reads: AtomicU64,
    xlsx_workbook: AtomicRangeCounter,
    xlsx_selected_worksheet: AtomicRangeCounter,
    xlsx_unselected_worksheets: AtomicRangeCounter,
    xlsx_shared_strings: AtomicRangeCounter,
    xlsx_styles: AtomicRangeCounter,
}

#[derive(Clone, Debug, Default)]
struct SelectiveReadSnapshot {
    read_calls: u64,
    read_bytes: u64,
    range_sizes: Vec<u64>,
}

#[derive(Debug, Default)]
struct SelectiveReadMetrics {
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
    range_sizes: Mutex<Vec<u64>>,
}

impl SelectiveReadMetrics {
    fn reset(&self) -> io::Result<()> {
        self.read_calls.store(0, Ordering::SeqCst);
        self.read_bytes.store(0, Ordering::SeqCst);
        self.range_sizes
            .lock()
            .map_err(|_| io::Error::other("selective CFB range metrics are poisoned"))?
            .clear();
        Ok(())
    }

    fn snapshot(&self) -> io::Result<SelectiveReadSnapshot> {
        let mut range_sizes = self
            .range_sizes
            .lock()
            .map_err(|_| io::Error::other("selective CFB range metrics are poisoned"))?
            .clone();
        range_sizes.sort_unstable();
        Ok(SelectiveReadSnapshot {
            read_calls: self.read_calls.load(Ordering::SeqCst),
            read_bytes: self.read_bytes.load(Ordering::SeqCst),
            range_sizes,
        })
    }

    fn record(&self, count: usize) -> io::Result<()> {
        let count = u64::try_from(count)
            .map_err(|_| io::Error::other("selective CFB range does not fit u64"))?;
        self.read_calls.fetch_add(1, Ordering::SeqCst);
        self.read_bytes.fetch_add(count, Ordering::SeqCst);
        self.range_sizes
            .lock()
            .map_err(|_| io::Error::other("selective CFB range metrics are poisoned"))?
            .push(count);
        Ok(())
    }
}

#[derive(Debug)]
struct SelectiveReadAt {
    bytes: Arc<Vec<u8>>,
    metrics: Arc<SelectiveReadMetrics>,
    version: SourceVersion,
}

impl SelectiveReadAt {
    fn new(bytes: Arc<Vec<u8>>, metrics: Arc<SelectiveReadMetrics>) -> Self {
        Self {
            bytes,
            metrics,
            version: SourceVersion::new(
                NEXT_INSTRUMENTED_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
                0,
            ),
        }
    }
}

impl ReadAt for SelectiveReadAt {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("selective CFB source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let count = self
            .bytes
            .get(start..)
            .map_or(0, |remaining| remaining.len().min(output.len()));
        if count != 0 {
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
        }
        self.metrics.record(count)?;
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

#[derive(Debug)]
struct SelectiveCursor {
    bytes: Arc<Vec<u8>>,
    position: u64,
    metrics: Arc<SelectiveReadMetrics>,
}

impl SelectiveCursor {
    fn new(bytes: Arc<Vec<u8>>, metrics: Arc<SelectiveReadMetrics>) -> Self {
        Self {
            bytes,
            position: 0,
            metrics,
        }
    }
}

impl io::Read for SelectiveCursor {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(self.position).unwrap_or(usize::MAX);
        let count = self
            .bytes
            .get(start..)
            .map_or(0, |remaining| remaining.len().min(output.len()));
        if count != 0 {
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
            self.position =
                self.position
                    .checked_add(u64::try_from(count).map_err(|_| {
                        io::Error::other("selective CFB cursor count does not fit u64")
                    })?)
                    .ok_or_else(|| io::Error::other("selective CFB cursor position overflows"))?;
        }
        self.metrics.record(count)?;
        Ok(count)
    }
}

impl Seek for SelectiveCursor {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let length = u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("selective CFB cursor length does not fit u64"))?;
        let next = match position {
            SeekFrom::Start(offset) => i128::from(offset),
            SeekFrom::Current(offset) => i128::from(self.position) + i128::from(offset),
            SeekFrom::End(offset) => i128::from(length) + i128::from(offset),
        };
        if next < 0 || next > i128::from(u64::MAX) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "selective CFB cursor seek is outside the source",
            ));
        }
        self.position = u64::try_from(next)
            .map_err(|_| io::Error::other("selective CFB cursor seek does not fit u64"))?;
        Ok(self.position)
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
struct RangeSimulationSnapshot {
    logical_read_calls: u64,
    logical_read_bytes: u64,
    physical_request_count: u64,
    physical_request_bytes: u64,
    physical_request_sizes: Vec<u64>,
    physical_request_size_buckets: RequestSizeBuckets,
}

#[derive(Debug)]
struct SimulatedRangeSource {
    backing: Arc<InstrumentedSource>,
    config: RangeSimulationConfig,
    logical_read_calls: AtomicU64,
    logical_read_bytes: AtomicU64,
    physical_request_count: AtomicU64,
    physical_request_bytes: AtomicU64,
    physical_request_sizes: Mutex<Vec<u64>>,
}

#[derive(Serialize)]
struct Statistics {
    unit: &'static str,
    samples: Vec<u64>,
    min: u64,
    p50: u64,
    p95: u64,
    p99: u64,
    max: u64,
    mean: f64,
    standard_deviation: f64,
    confidence_interval_95: ConfidenceInterval,
}

#[derive(Serialize)]
struct ConfidenceInterval {
    method: &'static str,
    lower: f64,
    upper: f64,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize)]
struct SinkSummary {
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    retained_output_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    retained_authoring_window_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    rows: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    cells: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    paragraphs: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    runs: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    input_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    authored_part_bytes: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    rtf_tail_append: Option<RtfTailAppendSummary>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
struct RtfTailAppendSummary {
    operation: &'static str,
    source_bytes: u64,
    input_bytes: u64,
    inserted_bytes: u64,
    output_bytes: u64,
    paragraphs: u64,
    runs: u64,
    sink_window_bytes: u64,
    exact_noop_verified: bool,
    in_memory_patch_verified: bool,
    durable_patch_verified: bool,
    reopen_verified: bool,
    source_conflict_verified: bool,
}

/// A non-seekable, bounded memory sink that consumes every output byte.
///
/// The complete byte budget is reserved before timing. Writes therefore
/// measure the memory copy performed by a real sequential consumer without
/// introducing allocator growth into the timed interval.
#[derive(Debug)]
struct CountingSink {
    summary: SinkSummary,
    max_bytes: u64,
    max_write: u64,
    bytes: Vec<u8>,
}

impl CountingSink {
    const fn bounded(max_bytes: u64, max_write: u64) -> Self {
        Self {
            summary: SinkSummary {
                accepted_bytes: 0,
                write_calls: 0,
                largest_write: 0,
                retained_output_bytes: None,
                retained_authoring_window_bytes: None,
                rows: None,
                cells: None,
                paragraphs: None,
                runs: None,
                input_bytes: None,
                authored_part_bytes: None,
                rtf_tail_append: None,
            },
            max_bytes,
            max_write,
            bytes: Vec::new(),
        }
    }

    fn reserve_budget(&mut self) -> io::Result<()> {
        let maximum = usize::try_from(self.max_bytes)
            .map_err(|_error| io::Error::other("sink byte budget does not fit usize"))?;
        self.bytes
            .try_reserve_exact(maximum)
            .map_err(|error| io::Error::other(format!("cannot reserve sink byte budget: {error}")))
    }

    const fn summary(&self) -> SinkSummary {
        self.summary
    }
}

impl Write for CountingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let length = u64::try_from(bytes.len())
            .map_err(|_error| io::Error::other("write length does not fit u64"))?;
        if length > self.max_write {
            return Err(io::Error::other(
                "sequential sink write exceeds configured maximum",
            ));
        }
        let accepted = self
            .summary
            .accepted_bytes
            .checked_add(length)
            .ok_or_else(|| io::Error::other("sequential sink byte count overflows u64"))?;
        if accepted > self.max_bytes {
            return Err(io::Error::other("sequential sink byte budget exceeded"));
        }

        self.summary.accepted_bytes = accepted;
        self.summary.write_calls = self
            .summary
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("sequential sink write count overflows u64"))?;
        self.summary.largest_write = self.summary.largest_write.max(length);
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct PrefixFailSink {
    accepted: u64,
    fail_after: u64,
}

impl Write for PrefixFailSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.fail_after {
            return Err(io::Error::other("intentional repair sink failure"));
        }
        let remaining = self.fail_after - self.accepted;
        let accepted =
            usize::try_from(remaining.min(u64::try_from(bytes.len()).map_err(|_error| {
                io::Error::other("repair sink write length does not fit u64")
            })?))
            .map_err(|_error| io::Error::other("repair sink accepted length does not fit usize"))?;
        self.accepted = self
            .accepted
            .checked_add(u64::try_from(accepted).map_err(|_error| {
                io::Error::other("repair sink accepted length does not fit u64")
            })?)
            .ok_or_else(|| io::Error::other("repair sink progress overflows u64"))?;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// A fixed-memory, non-seek sink for streaming-creation timings.
///
/// Only SHA-256 state and scalar counters are retained. The complete artifact
/// used for reopen verification is generated separately outside timing.
#[derive(Debug)]
struct HashingDiscardSink {
    summary: SinkSummary,
    maximum: u64,
    digest: Sha256,
}

impl HashingDiscardSink {
    fn new(maximum: u64, retained_authoring_window_bytes: u64) -> Self {
        Self {
            summary: SinkSummary {
                retained_output_bytes: Some(0),
                retained_authoring_window_bytes: Some(retained_authoring_window_bytes),
                ..SinkSummary::default()
            },
            maximum,
            digest: Sha256::new(),
        }
    }

    fn without_authoring_window(maximum: u64) -> Self {
        Self {
            summary: SinkSummary {
                retained_output_bytes: Some(0),
                ..SinkSummary::default()
            },
            maximum,
            digest: Sha256::new(),
        }
    }

    fn finish(self) -> (SinkSummary, String) {
        let digest = self.digest.finalize();
        let mut output = String::with_capacity(digest.len() * 2);
        for byte in digest {
            use std::fmt::Write as _;
            let _ = write!(output, "{byte:02x}");
        }
        (self.summary, output)
    }
}

impl Write for HashingDiscardSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let length = u64::try_from(bytes.len())
            .map_err(|_error| io::Error::other("streaming write length does not fit u64"))?;
        let accepted = self
            .summary
            .accepted_bytes
            .checked_add(length)
            .ok_or_else(|| io::Error::other("streaming output byte count overflows u64"))?;
        if accepted > self.maximum {
            return Err(io::Error::other("streaming output byte ceiling exceeded"));
        }
        self.digest.update(bytes);
        self.summary.accepted_bytes = accepted;
        self.summary.write_calls = self
            .summary
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("streaming write count overflows u64"))?;
        self.summary.largest_write = self.summary.largest_write.max(length);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// A bounded forward-only sink used by logical-tail publication measurements.
///
/// The tail-append API owns a validated candidate snapshot, so this sink keeps
/// publication accounting separate from candidate retention. It accepts at
/// most one fixed window per call, retains no output, and hashes the complete
/// stream for an untimed digest comparison.
#[derive(Debug)]
struct WindowedHashingSink {
    summary: SinkSummary,
    maximum: u64,
    window: usize,
    digest: Sha256,
}

impl WindowedHashingSink {
    fn new(maximum: u64, window: usize) -> io::Result<Self> {
        if window == 0 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "logical-tail sink window must be non-zero",
            ));
        }
        Ok(Self {
            summary: SinkSummary {
                retained_output_bytes: Some(0),
                retained_authoring_window_bytes: Some(u64::try_from(window).map_err(|_error| {
                    io::Error::new(
                        io::ErrorKind::InvalidInput,
                        "logical-tail sink window does not fit u64",
                    )
                })?),
                ..SinkSummary::default()
            },
            maximum,
            window,
            digest: Sha256::new(),
        })
    }

    fn finish(self) -> (SinkSummary, String) {
        let digest = self.digest.finalize();
        let mut output = String::with_capacity(digest.len() * 2);
        for byte in digest {
            use std::fmt::Write as _;
            let _ = write!(output, "{byte:02x}");
        }
        (self.summary, output)
    }
}

impl Write for WindowedHashingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = bytes.len().min(self.window);
        let length = u64::try_from(count)
            .map_err(|_error| io::Error::other("logical-tail write length does not fit u64"))?;
        let accepted = self
            .summary
            .accepted_bytes
            .checked_add(length)
            .ok_or_else(|| io::Error::other("logical-tail output byte count overflows u64"))?;
        if accepted > self.maximum {
            return Err(io::Error::other("logical-tail output ceiling exceeded"));
        }
        self.digest.update(bytes.get(..count).ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidData,
                "logical-tail sink window is outside the write buffer",
            )
        })?);
        self.summary.accepted_bytes = accepted;
        self.summary.write_calls = self
            .summary
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("logical-tail write count overflows u64"))?;
        self.summary.largest_write = self.summary.largest_write.max(length);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// A seekable counterpart used only where the public DOCX API requires it.
/// The summary counts actual writes even when the ZIP serializer rewrites a
/// header after seeking; `bytes` remains the final accepted package.
#[derive(Debug, Default)]
struct CountingSeekSink {
    cursor: Cursor<Vec<u8>>,
    summary: SinkSummary,
}

impl CountingSeekSink {
    fn into_parts(self) -> (Vec<u8>, SinkSummary) {
        (self.cursor.into_inner(), self.summary)
    }
}

impl Write for CountingSeekSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let accepted = self.cursor.write(bytes)?;
        let accepted_u64 = u64::try_from(accepted)
            .map_err(|_error| io::Error::other("seekable sink write length does not fit u64"))?;
        self.summary.accepted_bytes = self
            .summary
            .accepted_bytes
            .checked_add(accepted_u64)
            .ok_or_else(|| io::Error::other("seekable sink byte count overflows u64"))?;
        self.summary.write_calls = self
            .summary
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("seekable sink write count overflows u64"))?;
        self.summary.largest_write = self.summary.largest_write.max(accepted_u64);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.cursor.flush()
    }
}

impl Seek for CountingSeekSink {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.cursor.seek(position)
    }
}

impl InstrumentedSource {
    fn new(bytes: Vec<u8>, ordinary_payload_ranges: Vec<Range<u64>>) -> Self {
        Self::new_xlsx(bytes, ordinary_payload_ranges, XlsxTrackedRanges::default())
    }

    fn new_xlsx(
        bytes: Vec<u8>,
        ordinary_payload_ranges: Vec<Range<u64>>,
        xlsx_ranges: XlsxTrackedRanges,
    ) -> Self {
        Self {
            bytes: Arc::new(bytes),
            version: SourceVersion::new(
                NEXT_INSTRUMENTED_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
                0,
            ),
            ordinary_payload_ranges,
            xlsx_ranges,
            read_calls: AtomicU64::new(0),
            read_bytes: AtomicU64::new(0),
            ordinary_payload_read_calls: AtomicU64::new(0),
            ordinary_payload_read_bytes: AtomicU64::new(0),
            in_flight_reads: AtomicU64::new(0),
            max_in_flight_reads: AtomicU64::new(0),
            xlsx_workbook: AtomicRangeCounter::default(),
            xlsx_selected_worksheet: AtomicRangeCounter::default(),
            xlsx_unselected_worksheets: AtomicRangeCounter::default(),
            xlsx_shared_strings: AtomicRangeCounter::default(),
            xlsx_styles: AtomicRangeCounter::default(),
        }
    }

    fn snapshot(&self) -> SourceSnapshot {
        SourceSnapshot {
            read_calls: self.read_calls.load(Ordering::SeqCst),
            read_bytes: self.read_bytes.load(Ordering::SeqCst),
            ordinary_payload_read_calls: self.ordinary_payload_read_calls.load(Ordering::SeqCst),
            ordinary_payload_read_bytes: self.ordinary_payload_read_bytes.load(Ordering::SeqCst),
            max_in_flight_reads: self.max_in_flight_reads.load(Ordering::SeqCst),
            xlsx: XlsxSourceSnapshot {
                workbook: self.xlsx_workbook.snapshot(),
                selected_worksheet: self.xlsx_selected_worksheet.snapshot(),
                unselected_worksheets: self.xlsx_unselected_worksheets.snapshot(),
                shared_strings: self.xlsx_shared_strings.snapshot(),
                styles: self.xlsx_styles.snapshot(),
            },
        }
    }

    fn reset(&self) {
        debug_assert_eq!(self.in_flight_reads.load(Ordering::SeqCst), 0);
        self.read_calls.store(0, Ordering::SeqCst);
        self.read_bytes.store(0, Ordering::SeqCst);
        self.ordinary_payload_read_calls.store(0, Ordering::SeqCst);
        self.ordinary_payload_read_bytes.store(0, Ordering::SeqCst);
        self.max_in_flight_reads.store(0, Ordering::SeqCst);
        self.xlsx_workbook.reset();
        self.xlsx_selected_worksheet.reset();
        self.xlsx_unselected_worksheets.reset();
        self.xlsx_shared_strings.reset();
        self.xlsx_styles.reset();
    }
}

impl AtomicRangeCounter {
    fn observe(&self, bytes: u64) {
        if bytes != 0 {
            self.read_calls.fetch_add(1, Ordering::SeqCst);
            self.read_bytes.fetch_add(bytes, Ordering::SeqCst);
        }
    }

    fn snapshot(&self) -> RangeSnapshot {
        RangeSnapshot {
            read_calls: self.read_calls.load(Ordering::SeqCst),
            read_bytes: self.read_bytes.load(Ordering::SeqCst),
        }
    }

    fn reset(&self) {
        self.read_calls.store(0, Ordering::SeqCst);
        self.read_bytes.store(0, Ordering::SeqCst);
    }
}

fn range_overlap_bytes(ranges: &[Range<u64>], start: u64, end: u64) -> u64 {
    ranges
        .iter()
        .map(|range| end.min(range.end).saturating_sub(start.max(range.start)))
        .sum()
}

impl ReadAt for InstrumentedSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_error| io::Error::other("instrumented source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.read_calls.fetch_add(1, Ordering::SeqCst);
        let in_flight = self.in_flight_reads.fetch_add(1, Ordering::SeqCst) + 1;
        self.max_in_flight_reads
            .fetch_max(in_flight, Ordering::SeqCst);
        let _guard = InFlightReadGuard {
            counter: &self.in_flight_reads,
        };

        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let count = self
            .bytes
            .get(start..)
            .map_or(0, |remaining| remaining.len().min(output.len()));
        if count != 0 {
            output[..count].copy_from_slice(&self.bytes[start..start + count]);
        }
        let count_u64 = u64::try_from(count)
            .map_err(|_error| io::Error::other("instrumented read length does not fit u64"))?;
        self.read_bytes.fetch_add(count_u64, Ordering::SeqCst);

        let end = offset.saturating_add(count_u64);
        let ordinary_bytes = range_overlap_bytes(&self.ordinary_payload_ranges, offset, end);
        if ordinary_bytes != 0 {
            self.ordinary_payload_read_calls
                .fetch_add(1, Ordering::SeqCst);
            self.ordinary_payload_read_bytes
                .fetch_add(ordinary_bytes, Ordering::SeqCst);
        }
        self.xlsx_workbook
            .observe(range_overlap_bytes(&self.xlsx_ranges.workbook, offset, end));
        self.xlsx_selected_worksheet.observe(range_overlap_bytes(
            &self.xlsx_ranges.selected_worksheet,
            offset,
            end,
        ));
        self.xlsx_unselected_worksheets.observe(range_overlap_bytes(
            &self.xlsx_ranges.unselected_worksheets,
            offset,
            end,
        ));
        self.xlsx_shared_strings.observe(range_overlap_bytes(
            &self.xlsx_ranges.shared_strings,
            offset,
            end,
        ));
        self.xlsx_styles
            .observe(range_overlap_bytes(&self.xlsx_ranges.styles, offset, end));
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

impl RequestSizeBuckets {
    fn observe(&mut self, bytes: u64) {
        match bytes {
            0 => {},
            1..=512 => self.bytes_1_to_512 += 1,
            513..=4096 => self.bytes_513_to_4096 += 1,
            4097..=16384 => self.bytes_4097_to_16384 += 1,
            16385..=65536 => self.bytes_16385_to_65536 += 1,
            _ => self.bytes_over_65536 += 1,
        }
    }
}

fn simulated_request_delay(config: RangeSimulationConfig, bytes: usize) -> Duration {
    let base_nanos = u128::from(
        config
            .fixed_latency_us
            .saturating_add(config.request_overhead_us),
    ) * 1_000;
    let transfer_nanos = (bytes as u128)
        .saturating_mul(1_000_000_000)
        .div_ceil(u128::from(config.bandwidth_bytes_per_second));
    let nanos = base_nanos
        .saturating_add(transfer_nanos)
        .min(u128::from(u64::MAX));
    Duration::from_nanos(nanos as u64)
}

impl SimulatedRangeSource {
    fn new(backing: Arc<InstrumentedSource>, config: RangeSimulationConfig) -> Self {
        Self {
            backing,
            config,
            logical_read_calls: AtomicU64::new(0),
            logical_read_bytes: AtomicU64::new(0),
            physical_request_count: AtomicU64::new(0),
            physical_request_bytes: AtomicU64::new(0),
            physical_request_sizes: Mutex::new(Vec::new()),
        }
    }

    fn snapshot(&self) -> io::Result<RangeSimulationSnapshot> {
        let mut sizes = self
            .physical_request_sizes
            .lock()
            .map_err(|_error| io::Error::other("range simulator request sizes are poisoned"))?
            .clone();
        sizes.sort_unstable();
        let mut buckets = RequestSizeBuckets::default();
        for &size in &sizes {
            buckets.observe(size);
        }
        Ok(RangeSimulationSnapshot {
            logical_read_calls: self.logical_read_calls.load(Ordering::SeqCst),
            logical_read_bytes: self.logical_read_bytes.load(Ordering::SeqCst),
            physical_request_count: self.physical_request_count.load(Ordering::SeqCst),
            physical_request_bytes: self.physical_request_bytes.load(Ordering::SeqCst),
            physical_request_sizes: sizes,
            physical_request_size_buckets: buckets,
        })
    }

    fn reset(&self) -> io::Result<()> {
        self.logical_read_calls.store(0, Ordering::SeqCst);
        self.logical_read_bytes.store(0, Ordering::SeqCst);
        self.physical_request_count.store(0, Ordering::SeqCst);
        self.physical_request_bytes.store(0, Ordering::SeqCst);
        self.physical_request_sizes
            .lock()
            .map_err(|_error| io::Error::other("range simulator request sizes are poisoned"))?
            .clear();
        self.backing.reset();
        Ok(())
    }
}

impl ReadAt for SimulatedRangeSource {
    fn len(&self) -> io::Result<u64> {
        self.backing.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.logical_read_calls.fetch_add(1, Ordering::SeqCst);
        let mut total = 0usize;
        while total < output.len() {
            let requested = (output.len() - total).min(self.config.max_physical_range_bytes);
            std::thread::sleep(simulated_request_delay(self.config, requested));
            let physical_offset = offset.saturating_add(u64::try_from(total).unwrap_or(u64::MAX));
            let read = self
                .backing
                .read_at(physical_offset, &mut output[total..total + requested])?;
            let read_u64 = u64::try_from(read)
                .map_err(|_error| io::Error::other("physical request size does not fit u64"))?;
            self.physical_request_count.fetch_add(1, Ordering::SeqCst);
            self.physical_request_bytes
                .fetch_add(read_u64, Ordering::SeqCst);
            self.physical_request_sizes
                .lock()
                .map_err(|_error| io::Error::other("range simulator request sizes are poisoned"))?
                .push(read_u64);
            total += read;
            if read < requested {
                break;
            }
        }
        self.logical_read_bytes
            .fetch_add(u64::try_from(total).unwrap_or(u64::MAX), Ordering::SeqCst);
        Ok(total)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.backing.version()
    }
}

struct InFlightReadGuard<'counter> {
    counter: &'counter AtomicU64,
}

impl Drop for InFlightReadGuard<'_> {
    fn drop(&mut self) {
        self.counter.fetch_sub(1, Ordering::SeqCst);
    }
}

impl SourceSummary {
    fn record(&mut self, snapshot: SourceSnapshot) {
        self.read_calls.push(snapshot.read_calls);
        self.read_bytes.push(snapshot.read_bytes);
        self.ordinary_payload_read_calls
            .push(snapshot.ordinary_payload_read_calls);
        self.ordinary_payload_read_bytes
            .push(snapshot.ordinary_payload_read_bytes);
        self.max_in_flight_reads.push(snapshot.max_in_flight_reads);
    }

    fn record_opc(&mut self, snapshot: SourceSnapshot, payload_materializations: u64) {
        self.record(snapshot);
        self.ordinary_payload_materializations
            .get_or_insert_with(Vec::new)
            .push(payload_materializations);
    }

    fn record_xlsx(&mut self, snapshot: SourceSnapshot) {
        self.record(snapshot);
        let summary = self.xlsx.get_or_insert_with(XlsxSourceSummary::default);
        summary
            .workbook_read_calls
            .push(snapshot.xlsx.workbook.read_calls);
        summary
            .workbook_read_bytes
            .push(snapshot.xlsx.workbook.read_bytes);
        summary
            .selected_worksheet_read_calls
            .push(snapshot.xlsx.selected_worksheet.read_calls);
        summary
            .selected_worksheet_read_bytes
            .push(snapshot.xlsx.selected_worksheet.read_bytes);
        summary
            .unselected_worksheet_read_calls
            .push(snapshot.xlsx.unselected_worksheets.read_calls);
        summary
            .unselected_worksheet_read_bytes
            .push(snapshot.xlsx.unselected_worksheets.read_bytes);
        summary
            .shared_strings_read_calls
            .push(snapshot.xlsx.shared_strings.read_calls);
        summary
            .shared_strings_read_bytes
            .push(snapshot.xlsx.shared_strings.read_bytes);
        summary
            .styles_read_calls
            .push(snapshot.xlsx.styles.read_calls);
        summary
            .styles_read_bytes
            .push(snapshot.xlsx.styles.read_bytes);
    }

    fn record_xls_comments(
        &mut self,
        snapshot: SourceSnapshot,
        evidence: XlsCommentsIterationEvidence,
    ) -> Result<(), Box<dyn Error>> {
        self.record(snapshot);
        let summary = self
            .xls_comments
            .get_or_insert_with(|| XlsCommentsSourceSummary {
                source_counter_scope: "owned-source-ingress-only",
                source_backed: evidence.source_backed,
                update_count: evidence.update_count,
                ..XlsCommentsSourceSummary::default()
            });
        if summary.source_backed != evidence.source_backed
            || summary.update_count != evidence.update_count
        {
            return Err("XLS comment source evidence mixed incompatible cases".into());
        }
        summary
            .semantic_staging_plan_ns
            .push(evidence.semantic_staging_plan_ns);
        summary.publication_ns.push(evidence.publication_ns);
        summary.changed_comments.push(evidence.changed_comments);
        summary.touched_streams.push(evidence.touched_streams);
        summary.source_bytes.push(evidence.source_bytes);
        summary
            .source_workbook_bytes
            .push(evidence.source_workbook_bytes);
        summary
            .target_workbook_bytes
            .push(evidence.target_workbook_bytes);
        if let Some(splice_count) = evidence.splice_count {
            summary
                .splice_count
                .get_or_insert_with(Vec::new)
                .push(splice_count);
        }
        if let Some(replacement_bytes) = evidence.replacement_bytes {
            summary
                .replacement_bytes
                .get_or_insert_with(Vec::new)
                .push(replacement_bytes);
        }
        if let Some(changed_spans) = evidence.changed_spans {
            summary
                .changed_spans
                .get_or_insert_with(Vec::new)
                .push(changed_spans);
        }
        if let Some(fingerprint) = evidence.source_fingerprint {
            summary
                .source_fingerprints
                .get_or_insert_with(Vec::new)
                .push(fingerprint);
        }
        if let Some(fingerprint) = evidence.target_fingerprint {
            summary
                .target_fingerprints
                .get_or_insert_with(Vec::new)
                .push(fingerprint);
        }
        Ok(())
    }

    fn record_xls_visibility(
        &mut self,
        snapshot: SourceSnapshot,
        evidence: XlsVisibilityIterationEvidence,
    ) -> Result<(), Box<dyn Error>> {
        self.record(snapshot);
        let summary = self
            .xls_visibility
            .get_or_insert_with(|| XlsVisibilitySourceSummary {
                source_counter_scope: "owned-source-ingress-only",
                source_backed: evidence.source_backed,
                update_count: evidence.update_count,
                ..XlsVisibilitySourceSummary::default()
            });
        if summary.source_backed != evidence.source_backed
            || summary.update_count != evidence.update_count
        {
            return Err("XLS visibility source evidence mixed incompatible cases".into());
        }
        summary
            .semantic_staging_plan_ns
            .push(evidence.semantic_staging_plan_ns);
        summary.publication_ns.push(evidence.publication_ns);
        summary.changed_worksheets.push(evidence.changed_worksheets);
        summary.touched_streams.push(evidence.touched_streams);
        summary.source_bytes.push(evidence.source_bytes);
        summary
            .source_workbook_bytes
            .push(evidence.source_workbook_bytes);
        summary
            .target_workbook_bytes
            .push(evidence.target_workbook_bytes);
        if let Some(splice_count) = evidence.splice_count {
            summary
                .splice_count
                .get_or_insert_with(Vec::new)
                .push(splice_count);
        }
        if let Some(replacement_bytes) = evidence.replacement_bytes {
            summary
                .replacement_bytes
                .get_or_insert_with(Vec::new)
                .push(replacement_bytes);
        }
        if let Some(changed_spans) = evidence.changed_spans {
            summary
                .changed_spans
                .get_or_insert_with(Vec::new)
                .push(changed_spans);
        }
        if let Some(fingerprint) = evidence.source_fingerprint {
            summary
                .source_fingerprints
                .get_or_insert_with(Vec::new)
                .push(fingerprint);
        }
        if let Some(fingerprint) = evidence.target_fingerprint {
            summary
                .target_fingerprints
                .get_or_insert_with(Vec::new)
                .push(fingerprint);
        }
        Ok(())
    }

    fn record_simulation(&mut self, snapshot: RangeSimulationSnapshot) {
        let summary = self
            .simulation
            .get_or_insert_with(RangeSimulationSummary::default);
        summary.logical_read_calls.push(snapshot.logical_read_calls);
        summary.logical_read_bytes.push(snapshot.logical_read_bytes);
        summary
            .physical_request_count
            .push(snapshot.physical_request_count);
        summary
            .physical_request_bytes
            .push(snapshot.physical_request_bytes);
        summary
            .physical_request_sizes
            .push(snapshot.physical_request_sizes);
        summary
            .physical_request_size_buckets
            .push(snapshot.physical_request_size_buckets);
    }
}

fn check_status_name(status: &CheckStatus) -> &'static str {
    match status {
        CheckStatus::Complete => "complete",
        CheckStatus::NotApplicable => "not_applicable",
        CheckStatus::Blocked { .. } => "blocked",
        CheckStatus::StoppedBy { .. } => "stopped_by",
        _ => "unknown",
    }
}

fn require_complete_validation(
    case: Case,
    summary: &ValidationSummary,
) -> Result<(), Box<dyn Error>> {
    if summary.check_ids.is_empty() || !summary.complete || summary.has_errors {
        return Err(format!(
            "{} produced an incomplete or error validation report",
            case.name()
        )
        .into());
    }
    Ok(())
}

fn rtf_status_name(status: litchi_rtf::ValidationStatus) -> &'static str {
    match status {
        litchi_rtf::ValidationStatus::Valid => "valid",
        litchi_rtf::ValidationStatus::Present => "present",
        litchi_rtf::ValidationStatus::Absent => "absent",
        litchi_rtf::ValidationStatus::NotApplicable => "not_applicable",
        litchi_rtf::ValidationStatus::Unsupported => "unsupported",
        litchi_rtf::ValidationStatus::Unknown => "unknown",
        _ => "unknown",
    }
}

fn generic_validation_summary(
    report: &ValidateReport,
    source: &[u8],
    source_read_calls: Option<u64>,
    source_read_bytes: Option<u64>,
) -> Result<ValidationSummary, Box<dyn Error>> {
    let encoded = serde_json::to_vec(report)?;
    let check_ids = report
        .checks()
        .iter()
        .map(|check| check.id().as_str().to_owned())
        .collect::<Vec<_>>();
    let check_statuses = report
        .checks()
        .iter()
        .map(|check| check_status_name(check.status()).to_owned())
        .collect::<Vec<_>>();
    let issue_codes = report
        .issues()
        .iter()
        .map(|issue| issue.code().to_owned())
        .collect::<Vec<_>>();
    let mut counts = BTreeMap::new();
    counts.insert(
        "checks".to_owned(),
        u64::try_from(report.checks().len()).map_err(|_| "validation check count overflows u64")?,
    );
    counts.insert(
        "issues".to_owned(),
        u64::try_from(report.issues().len()).map_err(|_| "validation issue count overflows u64")?,
    );
    counts.insert("complete".to_owned(), u64::from(report.is_complete()));
    counts.insert("has_errors".to_owned(), u64::from(report.has_errors()));
    Ok(ValidationSummary {
        report_sha256: sha256_hex(&encoded),
        check_ids,
        check_statuses,
        issue_codes,
        issue_count: report.issues().len(),
        complete: report.is_complete(),
        has_errors: report.has_errors(),
        counts,
        source_sha256_before: sha256_hex(source),
        source_sha256_after: sha256_hex(source),
        source_bytes: u64::try_from(source.len())
            .map_err(|_| "validation source length overflows u64")?,
        source_read_calls,
        source_read_bytes,
        section_inventory: None,
    })
}

fn rtf_validation_summary(
    report: &litchi_rtf::ValidationReport,
    source: &[u8],
) -> Result<ValidationSummary, Box<dyn Error>> {
    let checks = [
        ("syntax", report.syntax()),
        ("root", report.root()),
        ("document", report.document()),
        ("compressed_transport", report.compressed_transport()),
        ("fields", report.fields()),
        ("external_links", report.external_links()),
        ("objects", report.objects()),
        ("pictures", report.pictures()),
        ("active_content", report.active_content()),
        ("unsupported_syntax", report.unsupported_syntax()),
        ("external_resolution", report.external_resolution()),
        ("execution", report.execution()),
        ("repair", report.repair()),
        ("security", report.security()),
    ];
    let check_ids = checks
        .iter()
        .map(|(id, _)| (*id).to_owned())
        .collect::<Vec<_>>();
    let check_statuses = checks
        .iter()
        .map(|(_, check)| rtf_status_name(check.status()).to_owned())
        .collect::<Vec<_>>();
    let complete = checks.iter().all(|(_, check)| {
        matches!(
            check.status(),
            litchi_rtf::ValidationStatus::Valid
                | litchi_rtf::ValidationStatus::Present
                | litchi_rtf::ValidationStatus::Absent
                | litchi_rtf::ValidationStatus::NotApplicable
        )
    });
    let counts_value = report.counts();
    let mut counts = BTreeMap::new();
    counts.insert(
        "source_bytes".to_owned(),
        u64::try_from(counts_value.source_bytes())
            .map_err(|_| "RTF source byte count overflows u64")?,
    );
    counts.insert(
        "fields".to_owned(),
        u64::try_from(counts_value.fields()).map_err(|_| "RTF field count overflows u64")?,
    );
    counts.insert(
        "objects".to_owned(),
        u64::try_from(counts_value.objects()).map_err(|_| "RTF object count overflows u64")?,
    );
    counts.insert(
        "pictures".to_owned(),
        u64::try_from(counts_value.pictures()).map_err(|_| "RTF picture count overflows u64")?,
    );
    counts.insert(
        "form_fields".to_owned(),
        u64::try_from(counts_value.form_fields())
            .map_err(|_| "RTF form-field count overflows u64")?,
    );
    counts.insert(
        "opaque_nodes".to_owned(),
        u64::try_from(counts_value.opaque_nodes())
            .map_err(|_| "RTF opaque-node count overflows u64")?,
    );
    counts.insert(
        "opaque_bytes".to_owned(),
        u64::try_from(counts_value.opaque_bytes())
            .map_err(|_| "RTF opaque-byte count overflows u64")?,
    );
    counts.insert(
        "unknown_syntax_markers".to_owned(),
        u64::try_from(counts_value.unknown_syntax_markers())
            .map_err(|_| "RTF unknown-syntax count overflows u64")?,
    );
    let canonical = serde_json::to_vec(&(&check_ids, &check_statuses, &counts))?;
    Ok(ValidationSummary {
        report_sha256: sha256_hex(&canonical),
        check_ids,
        check_statuses,
        issue_codes: Vec::new(),
        issue_count: 0,
        complete,
        has_errors: !complete,
        counts,
        source_sha256_before: sha256_hex(source),
        source_sha256_after: sha256_hex(source),
        source_bytes: u64::try_from(source.len()).map_err(|_| "RTF source length overflows u64")?,
        source_read_calls: None,
        source_read_bytes: None,
        section_inventory: None,
    })
}

fn section_inventory_summary(snapshot: &litchi_docx::section::Snapshot) -> SectionInventorySummary {
    let inventory = snapshot.inventory();
    let descriptors = inventory
        .sections()
        .iter()
        .map(|section| SectionDescriptorSummary {
            position: section.position().get(),
            ownership: match section.ownership() {
                litchi_docx::section::Ownership::Paragraph(position) => {
                    format!("paragraph:{}", position.get())
                },
                litchi_docx::section::Ownership::BodyFinal => "body_final".to_owned(),
                litchi_docx::section::Ownership::Implicit => "implicit".to_owned(),
                _ => "unknown".to_owned(),
            },
            paragraph_start: section.paragraphs().start().get(),
            paragraph_end: section.paragraphs().end().get(),
            page_width_emu: section
                .page_size()
                .and_then(|page| page.width.map(|value| value.0)),
            page_height_emu: section
                .page_size()
                .and_then(|page| page.height.map(|value| value.0)),
            page_orientation: section.page_size().map_or_else(
                || "none".to_owned(),
                |page| format!("{:?}", page.orientation),
            ),
            margin_left_emu: section
                .margins()
                .and_then(|margins| margins.left.map(|value| value.0)),
            margin_right_emu: section
                .margins()
                .and_then(|margins| margins.right.map(|value| value.0)),
            margin_top_emu: section
                .margins()
                .and_then(|margins| margins.top.map(|value| value.0)),
            margin_bottom_emu: section
                .margins()
                .and_then(|margins| margins.bottom.map(|value| value.0)),
            margin_header_emu: section
                .margins()
                .and_then(|margins| margins.header.map(|value| value.0)),
            margin_footer_emu: section
                .margins()
                .and_then(|margins| margins.footer.map(|value| value.0)),
            margin_gutter_emu: section
                .margins()
                .and_then(|margins| margins.gutter.map(|value| value.0)),
            start: section.start().map(|value| format!("{:?}", value)),
            headers: section
                .headers()
                .iter()
                .map(|reference| format!("{:?}:{}", reference.kind, reference.relationship_id))
                .collect(),
            footers: section
                .footers()
                .iter()
                .map(|reference| format!("{:?}:{}", reference.kind, reference.relationship_id))
                .collect(),
        })
        .collect::<Vec<_>>();
    SectionInventorySummary {
        section_count: descriptors.len(),
        paragraph_count: inventory.paragraph_count(),
        descriptors,
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    if filesystem::run_child_if_requested()? {
        return Ok(());
    }
    let options = parse_options()?;
    let mut results = Vec::new();
    let filesystem_runs = filesystem::run_selected(
        &options.cases,
        options.warmup_iterations,
        options.samples,
        options.filesystem_cache,
        options.filesystem_root.as_deref(),
    )?;
    let mut filesystem_evidence = Vec::with_capacity(filesystem_runs.len());
    for run in filesystem_runs {
        if let Some(result) = run.warm_result {
            results.push(result);
        }
        if let Some(result) = run.cold_result {
            results.push(result);
        }
        filesystem_evidence.push(run.evidence);
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_opc_source_cache_evidence())
    {
        let corpus = build_opc_corpus(CorpusShape::ManySmall, PayloadKind::Incompressible)?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_opc_source_cache_evidence())
        {
            match case {
                Case::OpcSourceCacheBudgetBoundary => {
                    results.extend(run_opc_source_cache_budget_boundary(
                        &corpus,
                        options.warmup_iterations,
                        options.samples,
                    )?)
                },
                Case::OpcSourceCacheControlContention => {
                    results.extend(run_opc_source_cache_contention(
                        *case,
                        &corpus,
                        options.warmup_iterations,
                        options.samples,
                        &options.execution_workers,
                        OpcCacheMode::Control,
                    )?)
                },
                Case::OpcSourceCacheManagedContention => {
                    results.extend(run_opc_source_cache_contention(
                        *case,
                        &corpus,
                        options.warmup_iterations,
                        options.samples,
                        &options.execution_workers,
                        OpcCacheMode::Managed,
                    )?)
                },
                _ => unreachable!("filtered OPC source-cache evidence case"),
            }
        }
    }

    for shape in &options.shapes {
        for payload in &options.payloads {
            let opc_corpus = options
                .cases
                .iter()
                .any(|case| case.uses_synthetic_opc())
                .then(|| build_opc_corpus(*shape, *payload))
                .transpose()?;
            let cfb_corpus = options
                .cases
                .iter()
                .any(|case| case.uses_synthetic_cfb())
                .then(|| build_cfb_corpus(*shape, *payload))
                .transpose()?;
            for case in options.cases.iter().filter(|case| {
                !case.is_fresh_writer()
                    && !case.uses_semantic_doc()
                    && !case.uses_semantic_xls()
                    && !case.uses_semantic_ppt()
                    && !case.uses_xlsx()
                    && !case.uses_xlsx_cell_values()
                    && !case.uses_streaming_creation()
                    && !case.uses_semantic_rtf()
                    && !case.uses_semantic_docx()
                    && !case.uses_semantic_pptx()
                    && !case.uses_semantic_odt()
                    && !case.uses_odt_media()
                    && !case.uses_odt_resource_batch()
                    && !case.uses_semantic_ods()
                    && !case.uses_ods_media()
                    && !case.uses_semantic_odp()
                    && !case.uses_odp_media()
                    && !case.uses_odp_text_box_batch()
                    && !case.is_opc_source_overlay_save()
                    && !case.is_opc_source_cache_evidence()
                    && !case.is_filesystem()
                    && !case.is_docx_source_edit_save()
                    && !case.is_pptx_source_edit_save()
                    && !case.is_xlsx_calculation_metadata_edit_save()
                    && !case.is_xlsx_defined_names_edit_save()
                    && !case.is_xlsx_page_break_edit_save()
                    && !case.is_xlsx_page_margin_edit_save()
                    && !case.is_xlsx_page_setup_edit_save()
                    && !case.is_xlsx_print_options_edit_save()
                    && !case.is_xlsx_sheet_protection_edit_save()
                    && !case.is_xlsx_data_validation_edit_save()
                    && !case.is_xlsx_auto_filter_edit_save()
                    && !case.is_xlsx_conditional_formatting_edit_save()
                    && !case.is_xlsx_merge_edit_save()
                    && !case.is_xls_comments_edit_save()
                    && !case.is_xls_visibility_edit_save()
                    && !case.uses_validation_xls()
                    && !case.uses_validation_rtf()
                    && !case.uses_validation_docx()
                    && !case.uses_validation_pptx()
                    && !case.uses_validation_odf()
                    && !case.uses_odf_repair()
                    && !case.is_cfb_selective()
            }) {
                let corpus = if case.uses_synthetic_cfb() {
                    cfb_corpus
                        .as_ref()
                        .ok_or("CFB case has no generated CFB corpus")?
                } else {
                    opc_corpus
                        .as_ref()
                        .ok_or("OPC case has no generated OPC corpus")?
                };
                if case.is_scaling() {
                    for &workers in &options.execution_workers {
                        results.push(run_scaling_case(
                            *case,
                            corpus,
                            options.warmup_iterations,
                            options.samples,
                            workers,
                        )?);
                    }
                } else {
                    results.push(run_case_with_config(
                        *case,
                        corpus,
                        options.warmup_iterations,
                        options.samples,
                        options.range_simulation,
                    )?);
                }
            }
        }
    }

    if options.cases.iter().any(|case| case.is_cfb_selective()) {
        for shape in options
            .shapes
            .iter()
            .copied()
            .filter(|shape| matches!(shape, CorpusShape::ManySmall | CorpusShape::WideRoot))
        {
            for target in [CfbSelectiveTarget::Mini, CfbSelectiveTarget::Fat] {
                let corpus = build_cfb_selective_corpus(shape, target)?;
                for case in options.cases.iter().copied().filter(|case| {
                    case.is_cfb_selective()
                        && match target {
                            CfbSelectiveTarget::Mini => {
                                matches!(
                                    case,
                                    Case::CfbSelectiveMiniLegacyRead
                                        | Case::CfbSelectiveMiniSharedRead
                                )
                            },
                            CfbSelectiveTarget::Fat => {
                                matches!(
                                    case,
                                    Case::CfbSelectiveFatLegacyRead
                                        | Case::CfbSelectiveFatSharedRead
                                )
                            },
                        }
                }) {
                    results.push(run_cfb_selective_read(
                        case,
                        &corpus,
                        options.warmup_iterations,
                        options.samples,
                    )?);
                }
            }
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_opc_source_overlay_save())
    {
        let corpus = build_opc_corpus(CorpusShape::FewLarge, PayloadKind::Incompressible)?;
        results.push(run_opc_source_overlay_one_part_save(
            &corpus,
            options.warmup_iterations,
            options.samples,
        )?);
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xls_comments_edit_save())
    {
        let corpus = build_xls_comments_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xls_comments_edit_save())
        {
            results.push(run_xls_comments_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xls_visibility_edit_save())
    {
        let corpus = build_xls_visibility_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xls_visibility_edit_save())
        {
            results.push(run_xls_visibility_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xlsx_page_break_edit_save())
    {
        let corpus = build_xlsx_page_break_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xlsx_page_break_edit_save())
        {
            results.push(run_xlsx_page_break_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xlsx_page_margin_edit_save())
    {
        let corpus = build_xlsx_page_margin_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xlsx_page_margin_edit_save())
        {
            results.push(run_xlsx_page_margin_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xlsx_print_options_edit_save())
    {
        let corpus = build_xlsx_print_options_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xlsx_print_options_edit_save())
        {
            results.push(run_xlsx_print_options_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xlsx_sheet_protection_edit_save())
    {
        let corpus = build_xlsx_sheet_protection_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xlsx_sheet_protection_edit_save())
        {
            results.push(run_xlsx_sheet_protection_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xlsx_data_validation_edit_save())
    {
        let corpus = build_xlsx_data_validation_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xlsx_data_validation_edit_save())
        {
            results.push(run_xlsx_data_validation_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xlsx_auto_filter_edit_save())
    {
        let corpus = build_xlsx_auto_filter_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xlsx_auto_filter_edit_save())
        {
            results.push(run_xlsx_auto_filter_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xlsx_conditional_formatting_edit_save())
    {
        let corpus = build_xlsx_conditional_formatting_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xlsx_conditional_formatting_edit_save())
        {
            results.push(run_xlsx_conditional_formatting_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xlsx_page_setup_edit_save())
    {
        let corpus = build_xlsx_page_setup_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xlsx_page_setup_edit_save())
        {
            results.push(run_xlsx_page_setup_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_docx_source_edit_save())
    {
        let corpus = build_docx_source_edit_corpus()?;
        results.push(run_docx_source_backed_one_edit_save(
            &corpus,
            options.warmup_iterations,
            options.samples,
        )?);
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_pptx_source_edit_save())
    {
        let corpus = build_pptx_source_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_pptx_source_edit_save())
        {
            results.push(match case {
                Case::PptxSourceBackedOneEditSave => run_pptx_source_backed_one_edit_save(
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                )?,
                Case::PptxEagerBatchEditSave | Case::PptxSourceBackedBatchEditSave => {
                    run_pptx_batch_edit_save(
                        *case,
                        &corpus,
                        options.warmup_iterations,
                        options.samples,
                    )?
                },
                Case::PptxEagerMultiSlideBatchEditSave
                | Case::PptxSourceBackedMultiSlideBatchEditSave => {
                    run_pptx_multi_slide_batch_edit_save(
                        *case,
                        &corpus,
                        options.warmup_iterations,
                        options.samples,
                    )?
                },
                _ => unreachable!("filtered PPTX source-edit case"),
            });
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xlsx_calculation_metadata_edit_save())
    {
        let corpus = build_xlsx_calculation_metadata_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xlsx_calculation_metadata_edit_save())
        {
            results.push(run_xlsx_calculation_metadata_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xlsx_defined_names_edit_save())
    {
        let corpus = build_xlsx_defined_names_edit_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xlsx_defined_names_edit_save())
        {
            results.push(run_xlsx_defined_names_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.is_xlsx_merge_edit_save())
    {
        for case in options
            .cases
            .iter()
            .filter(|case| case.is_xlsx_merge_edit_save())
        {
            let corpus = build_xlsx_merge_edit_corpus(*case)?;
            results.push(run_xlsx_merge_edit_save(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
            )?);
        }
    }

    for shape in &options.writer_shapes {
        for case in options.cases.iter().filter(|case| case.is_fresh_writer()) {
            let corpus = build_writer_corpus(*case, *shape)?;
            results.push(run_case_with_config(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
                options.range_simulation,
            )?);
        }
    }

    for shape in options
        .writer_shapes
        .iter()
        .filter(|shape| **shape != WriterShape::PayloadHeavy)
    {
        if options.cases.iter().any(|case| case.uses_semantic_doc()) {
            let corpus = build_writer_corpus(Case::DocFreshWriteTo, *shape)?;
            for case in options.cases.iter().filter(|case| case.uses_semantic_doc()) {
                results.push(run_case_with_config(
                    *case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
        if options.cases.iter().any(|case| case.uses_semantic_xls()) {
            let corpus = build_writer_corpus(Case::XlsFreshWriteTo, *shape)?;
            for case in options.cases.iter().filter(|case| case.uses_semantic_xls()) {
                results.push(run_case_with_config(
                    *case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
        if options.cases.iter().any(|case| case.uses_semantic_ppt()) {
            let corpus = build_writer_corpus(Case::PptFreshWriteTo, *shape)?;
            for case in options.cases.iter().filter(|case| case.uses_semantic_ppt()) {
                results.push(run_case_with_config(
                    *case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    if options.cases.iter().any(|case| case.uses_xlsx()) {
        for shape in &options.xlsx_shapes {
            let corpus = build_xlsx_corpus(*shape)?;
            for case in options.cases.iter().filter(|case| case.uses_xlsx()) {
                results.push(run_case_with_config(
                    *case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.uses_xlsx_cell_values())
    {
        for shape in &options.xlsx_cell_crud_shapes {
            let corpus = build_xlsx_cell_crud_corpus(*shape)?;
            for case in options
                .cases
                .iter()
                .filter(|case| case.uses_xlsx_cell_values())
            {
                results.push(run_xlsx_cell_values_edit_save(
                    *case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                )?);
            }
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.uses_streaming_creation())
    {
        for shape in &options.semantic_shapes {
            for case in options
                .cases
                .iter()
                .filter(|case| case.uses_streaming_creation())
            {
                let corpus = build_streaming_corpus(*case, *shape)?;
                results.push(run_streaming_creation(
                    *case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                )?);
            }
        }
    }

    if options.cases.iter().any(|case| case.uses_semantic_docx()) {
        for shape in &options.semantic_shapes {
            let corpus = build_semantic_docx_corpus(*shape)?;
            for case in options.cases.iter().filter(|case| {
                case.uses_semantic_docx()
                    && (*shape == SemanticShape::Tiny || !case.is_semantic_create_small())
            }) {
                results.push(run_case_with_config(
                    *case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    if options.cases.iter().any(|case| case.uses_semantic_rtf()) {
        let mut rtf_rows = 0usize;
        for variant in &options.rtf_variants {
            for shape in options
                .semantic_shapes
                .iter()
                .filter(|shape| variant.supports_shape(**shape))
            {
                let semantic_corpus = build_semantic_rtf_corpus(*shape, *variant)?;
                let lifecycle_corpus = (*variant == RtfSemanticVariant::Plain
                    && options
                        .cases
                        .iter()
                        .any(|case| case.is_rtf_lifecycle() || case.is_rtf_logical_tail()))
                .then(|| build_rtf_lifecycle_corpus(*shape))
                .transpose()?;
                for case in options
                    .cases
                    .iter()
                    .filter(|case| variant.supports_case(**case))
                {
                    let corpus = if case.is_rtf_lifecycle() || case.is_rtf_logical_tail() {
                        lifecycle_corpus
                            .as_ref()
                            .ok_or("RTF lifecycle case has no plain lifecycle corpus")?
                    } else {
                        &semantic_corpus
                    };
                    results.push(run_case_with_config(
                        *case,
                        corpus,
                        options.warmup_iterations,
                        options.samples,
                        options.range_simulation,
                    )?);
                    rtf_rows = rtf_rows
                        .checked_add(1)
                        .ok_or("semantic RTF result count overflows usize")?;
                }
            }
        }
        if rtf_rows == 0 {
            return Err("selected RTF variants and shapes produce no supported cases".into());
        }
    }

    if options.cases.iter().any(|case| case.uses_semantic_pptx()) {
        for shape in &options.semantic_shapes {
            let corpus = build_semantic_pptx_corpus(*shape)?;
            for case in options.cases.iter().filter(|case| {
                case.uses_semantic_pptx()
                    && (*shape == SemanticShape::Tiny || !case.is_semantic_create_small())
            }) {
                results.push(run_case_with_config(
                    *case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    if options.cases.iter().any(|case| case.uses_semantic_odt()) {
        for shape in &options.semantic_shapes {
            let corpus = build_semantic_odt_corpus(*shape)?;
            for case in options.cases.iter().filter(|case| {
                case.uses_semantic_odt()
                    && (*shape == SemanticShape::Tiny || !case.is_semantic_create_small())
            }) {
                results.push(run_case_with_config(
                    *case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    if options.cases.iter().any(|case| case.uses_semantic_ods()) {
        for shape in &options.semantic_shapes {
            let corpus = build_semantic_ods_corpus(*shape)?;
            for case in options.cases.iter().filter(|case| {
                case.uses_semantic_ods()
                    && (*shape == SemanticShape::Tiny || !case.is_semantic_create_small())
            }) {
                results.push(run_case_with_config(
                    *case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    if options.cases.iter().any(|case| case.uses_odt_media()) {
        let corpus = build_odt_media_corpus()?;
        for case in options.cases.iter().filter(|case| case.uses_odt_media()) {
            results.push(run_case_with_config(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
                options.range_simulation,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.uses_odt_resource_batch())
    {
        let corpus = build_odt_resource_batch_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.uses_odt_resource_batch())
        {
            results.push(run_case_with_config(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
                options.range_simulation,
            )?);
        }
    }

    if options.cases.iter().any(|case| case.uses_ods_media()) {
        let corpus = build_ods_media_corpus()?;
        for case in options.cases.iter().filter(|case| case.uses_ods_media()) {
            results.push(run_case_with_config(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
                options.range_simulation,
            )?);
        }
    }

    if options.cases.iter().any(|case| case.uses_semantic_odp()) {
        for shape in &options.semantic_shapes {
            let corpus = build_semantic_odp_corpus(*shape)?;
            for case in options.cases.iter().filter(|case| {
                case.uses_semantic_odp()
                    && (*shape == SemanticShape::Tiny || !case.is_semantic_create_small())
            }) {
                results.push(run_case_with_config(
                    *case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    if options.cases.iter().any(|case| case.uses_odp_media()) {
        let corpus = build_odp_media_corpus()?;
        for case in options.cases.iter().filter(|case| case.uses_odp_media()) {
            results.push(run_case_with_config(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
                options.range_simulation,
            )?);
        }
    }

    if options
        .cases
        .iter()
        .any(|case| case.uses_odp_text_box_batch())
    {
        let corpus = build_odp_text_box_batch_corpus()?;
        for case in options
            .cases
            .iter()
            .filter(|case| case.uses_odp_text_box_batch())
        {
            results.push(run_case_with_config(
                *case,
                &corpus,
                options.warmup_iterations,
                options.samples,
                options.range_simulation,
            )?);
        }
    }

    if options.cases.iter().any(|case| case.uses_validation_xls()) {
        for shape in options
            .writer_shapes
            .iter()
            .copied()
            .filter(|shape| *shape != WriterShape::PayloadHeavy)
        {
            let corpus = build_writer_corpus(Case::XlsFreshWriteTo, shape)?;
            for case in options
                .cases
                .iter()
                .copied()
                .filter(|case| case.uses_validation_xls())
            {
                results.push(run_case_with_config(
                    case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    if options.cases.iter().any(|case| case.uses_validation_docx()) {
        for shape in &options.semantic_shapes {
            let corpus = build_semantic_docx_corpus(*shape)?;
            for case in options
                .cases
                .iter()
                .copied()
                .filter(|case| case.uses_validation_docx())
            {
                results.push(run_case_with_config(
                    case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    if options.cases.iter().any(|case| case.uses_validation_pptx()) {
        for shape in &options.semantic_shapes {
            let corpus = build_semantic_pptx_corpus(*shape)?;
            for case in options
                .cases
                .iter()
                .copied()
                .filter(|case| case.uses_validation_pptx())
            {
                results.push(run_case_with_config(
                    case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    if options.cases.iter().any(|case| case.uses_validation_rtf()) {
        let mut rtf_validation_rows = 0usize;
        for variant in &options.rtf_variants {
            for shape in options
                .semantic_shapes
                .iter()
                .filter(|shape| variant.supports_validation() && variant.supports_shape(**shape))
            {
                let corpus = build_semantic_rtf_corpus(*shape, *variant)?;
                for case in options
                    .cases
                    .iter()
                    .copied()
                    .filter(|case| case.uses_validation_rtf())
                {
                    results.push(run_case_with_config(
                        case,
                        &corpus,
                        options.warmup_iterations,
                        options.samples,
                        options.range_simulation,
                    )?);
                    rtf_validation_rows = rtf_validation_rows
                        .checked_add(1)
                        .ok_or("RTF validation result count overflows usize")?;
                }
            }
        }
        if rtf_validation_rows == 0 {
            return Err(
                "selected RTF variants and shapes produce no validation-supported cases".into(),
            );
        }
    }

    if options.cases.iter().any(|case| case.uses_validation_odf()) {
        for shape in &options.semantic_shapes {
            let corpus = build_semantic_odt_corpus(*shape)?;
            for case in options
                .cases
                .iter()
                .copied()
                .filter(|case| case.uses_validation_odf())
            {
                results.push(run_case_with_config(
                    case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    if options.cases.iter().any(|case| case.uses_odf_repair()) {
        for shape in &options.semantic_shapes {
            let corpus = build_odf_repair_corpus(*shape)?;
            for case in options
                .cases
                .iter()
                .copied()
                .filter(|case| case.uses_odf_repair())
            {
                results.push(run_case_with_config(
                    case,
                    &corpus,
                    options.warmup_iterations,
                    options.samples,
                    options.range_simulation,
                )?);
            }
        }
    }

    let report = Report {
        schema_version: SCHEMA_VERSION,
        tool: Tool {
            name: env!("CARGO_PKG_NAME"),
            version: env!("CARGO_PKG_VERSION"),
            profile: if cfg!(debug_assertions) {
                "debug"
            } else {
                "release"
            },
            target_os: std::env::consts::OS,
            target_arch: std::env::consts::ARCH,
        },
        environment: environment(filesystem::host_evidence(
            options.filesystem_root.as_deref(),
            !filesystem_evidence.is_empty(),
        )),
        configuration: Configuration {
            samples_per_case: options.samples,
            warmup_iterations_per_case: options.warmup_iterations,
            filesystem_cache_states: options.filesystem_cache.names(),
            filesystem_fresh_child_per_sample: true,
            filesystem_process_isolated: true,
            filesystem_root_selected: options.filesystem_root.is_some(),
            cases: options.cases.iter().map(|case| case.name()).collect(),
            corpus_shapes: options.shapes.iter().map(|shape| shape.name()).collect(),
            payload_kinds: options.payloads.iter().map(|kind| kind.name()).collect(),
            writer_shapes: options
                .writer_shapes
                .iter()
                .map(|shape| shape.name())
                .collect(),
            xlsx_shapes: options
                .xlsx_shapes
                .iter()
                .map(|shape| shape.name())
                .collect(),
            xlsx_cell_crud_shapes: options
                .xlsx_cell_crud_shapes
                .iter()
                .map(|shape| shape.name())
                .collect(),
            semantic_shapes: options
                .semantic_shapes
                .iter()
                .map(|shape| shape.name())
                .collect(),
            rtf_variants: options
                .rtf_variants
                .iter()
                .map(|variant| variant.name())
                .collect(),
            range_simulation: options.range_simulation,
            execution_workers: options.execution_workers,
        },
        results,
        filesystem_evidence: (!filesystem_evidence.is_empty()).then_some(filesystem_evidence),
    };

    write_report(&report, options.output.as_ref())
}

fn parse_options() -> Result<Options, Box<dyn Error>> {
    let mut samples = DEFAULT_SAMPLES;
    let mut warmup_iterations = DEFAULT_WARMUP_ITERATIONS;
    let mut filesystem_cache = filesystem::CacheSelection::default();
    let mut filesystem_root = None;
    let mut cases = Case::DEFAULT.to_vec();
    let mut shapes = CorpusShape::ALL.to_vec();
    let mut payloads = PayloadKind::ALL.to_vec();
    let mut writer_shapes = WriterShape::ALL.to_vec();
    let mut xlsx_shapes = XlsxShape::ALL.to_vec();
    let mut xlsx_cell_crud_shapes = XlsxCellCrudShape::ALL.to_vec();
    let mut semantic_shapes = SemanticShape::ALL.to_vec();
    let mut rtf_variants = vec![RtfSemanticVariant::Plain];
    let mut range_simulation = RangeSimulationConfig::default();
    let mut execution_workers = default_execution_workers()?;
    let mut output = None;
    let mut arguments = std::env::args().skip(1);

    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--samples" => {
                let value = arguments
                    .next()
                    .ok_or("--samples requires a positive integer")?;
                samples = value.parse()?;
                if samples == 0 {
                    return Err("--samples must be greater than zero".into());
                }
            },
            "--warmup" => {
                let value = arguments
                    .next()
                    .ok_or("--warmup requires a non-negative integer")?;
                warmup_iterations = value.parse()?;
            },
            "--filesystem-cache" => {
                filesystem_cache = filesystem::CacheSelection::parse(
                    &arguments
                        .next()
                        .ok_or("--filesystem-cache requires warm,cold-requested")?,
                )?;
            },
            "--filesystem-root" => {
                filesystem_root = Some(PathBuf::from(
                    arguments.next().ok_or("--filesystem-root requires PATH")?,
                ));
            },
            "--case" => cases = parse_selection(arguments.next(), "--case", parse_case)?,
            "--shape" => shapes = parse_selection(arguments.next(), "--shape", parse_shape)?,
            "--payload" => {
                payloads = parse_selection(arguments.next(), "--payload", parse_payload)?;
            },
            "--writer-shape" => {
                writer_shapes =
                    parse_selection(arguments.next(), "--writer-shape", parse_writer_shape)?;
            },
            "--xlsx-shape" => {
                xlsx_shapes = parse_selection(arguments.next(), "--xlsx-shape", parse_xlsx_shape)?;
            },
            "--xlsx-cell-crud-shape" => {
                xlsx_cell_crud_shapes = parse_selection(
                    arguments.next(),
                    "--xlsx-cell-crud-shape",
                    parse_xlsx_cell_crud_shape,
                )?;
            },
            "--semantic-shape" => {
                semantic_shapes =
                    parse_selection(arguments.next(), "--semantic-shape", parse_semantic_shape)?;
            },
            "--rtf-variant" => {
                rtf_variants =
                    parse_selection(arguments.next(), "--rtf-variant", parse_rtf_variant)?;
            },
            "--range-fixed-latency-us" => {
                range_simulation.fixed_latency_us =
                    parse_u64_option(arguments.next(), "--range-fixed-latency-us", true)?;
            },
            "--range-request-overhead-us" => {
                range_simulation.request_overhead_us =
                    parse_u64_option(arguments.next(), "--range-request-overhead-us", true)?;
            },
            "--range-bandwidth-bytes-per-sec" => {
                range_simulation.bandwidth_bytes_per_second =
                    parse_u64_option(arguments.next(), "--range-bandwidth-bytes-per-sec", false)?;
            },
            "--range-max-physical-bytes" => {
                range_simulation.max_physical_range_bytes = usize::try_from(parse_u64_option(
                    arguments.next(),
                    "--range-max-physical-bytes",
                    false,
                )?)?;
            },
            "--workers" => {
                execution_workers = parse_execution_workers(arguments.next())?;
            },
            "--json" => {
                let value = arguments.next().ok_or("--json requires PATH or -")?;
                if value != "-" {
                    output = Some(PathBuf::from(value));
                }
            },
            "--help" | "-h" => {
                print_usage();
                std::process::exit(0);
            },
            _ => return Err(format!("unrecognized argument {argument:?}; use --help").into()),
        }
    }

    Ok(Options {
        samples,
        warmup_iterations,
        filesystem_cache,
        filesystem_root,
        cases,
        shapes,
        payloads,
        writer_shapes,
        xlsx_shapes,
        xlsx_cell_crud_shapes,
        semantic_shapes,
        rtf_variants,
        range_simulation,
        execution_workers,
        output,
    })
}

fn parse_u64_option(
    value: Option<String>,
    option: &str,
    allow_zero: bool,
) -> Result<u64, Box<dyn Error>> {
    let value = value.ok_or_else(|| format!("{option} requires an integer"))?;
    let parsed = value.parse::<u64>()?;
    if !allow_zero && parsed == 0 {
        return Err(format!("{option} must be greater than zero").into());
    }
    Ok(parsed)
}

fn default_execution_workers() -> Result<Vec<usize>, Box<dyn Error>> {
    resolve_execution_workers(["1", "2", "4", "8", "available"])
}

fn parse_execution_workers(value: Option<String>) -> Result<Vec<usize>, Box<dyn Error>> {
    let value = value.ok_or("--workers requires a comma-separated list")?;
    if value.is_empty() {
        return Err("--workers selection must not be empty".into());
    }
    resolve_execution_workers(value.split(','))
}

fn resolve_execution_workers<'a>(
    values: impl IntoIterator<Item = &'a str>,
) -> Result<Vec<usize>, Box<dyn Error>> {
    let available = std::thread::available_parallelism().map_or(1, NonZeroUsize::get);
    let mut workers = BTreeSet::new();
    for value in values {
        let requested = if value == "available" {
            available
        } else {
            value.parse::<usize>()?
        };
        if requested == 0 {
            return Err("worker counts must be greater than zero".into());
        }
        workers.insert(requested.min(available));
    }
    if workers.is_empty() {
        return Err("worker selection must not be empty".into());
    }
    Ok(workers.into_iter().collect())
}

fn parse_selection<T>(
    value: Option<String>,
    option: &str,
    parse: impl Fn(&str) -> Option<T>,
) -> Result<Vec<T>, Box<dyn Error>> {
    let value = value.ok_or_else(|| format!("{option} requires a comma-separated value"))?;
    if value.is_empty() {
        return Err(format!("{option} selection must not be empty").into());
    }
    if value == "all" {
        return Err(format!("{option}=all is implicit; omit {option} instead").into());
    }
    let selected: Option<Vec<T>> = value.split(',').map(parse).collect();
    selected.ok_or_else(|| format!("invalid {option} selection {value:?}").into())
}

fn parse_case(value: &str) -> Option<Case> {
    match value {
        "zip_index" => Some(Case::ZipIndex),
        "zip_read_one" => Some(Case::ZipReadOne),
        "opc_open" => Some(Case::OpcOpen),
        "opc_open_owned" => Some(Case::OpcOpenOwned),
        "opc_noop_save" => Some(Case::OpcNoopSave),
        "opc_mutated_save" => Some(Case::OpcMutatedSave),
        "opc_source_open" => Some(Case::OpcSourceOpen),
        "opc_source_open_main_read" => Some(Case::OpcSourceOpenMainRead),
        "opc_source_cached_main_read" => Some(Case::OpcSourceCachedMainRead),
        "opc_source_concurrent_same_part" => Some(Case::OpcSourceConcurrentSamePart),
        "opc_source_cache_budget_boundary" => Some(Case::OpcSourceCacheBudgetBoundary),
        "opc_source_cache_control_contention" => Some(Case::OpcSourceCacheControlContention),
        "opc_source_cache_managed_contention" => Some(Case::OpcSourceCacheManagedContention),
        "opc_source_overlay_one_part_save" => Some(Case::OpcSourceOverlayOnePartSave),
        "opc_file_eager_open" => Some(Case::OpcFileEagerOpen),
        "opc_file_source_open" => Some(Case::OpcFileSourceOpen),
        "opc_file_eager_one_part_atomic_save" => Some(Case::OpcFileEagerOnePartAtomicSave),
        "opc_file_source_one_part_atomic_save" => Some(Case::OpcFileSourceOnePartAtomicSave),
        "cfb_file_same_length_overlay_atomic_save" => {
            Some(Case::CfbFileSameLengthOverlayAtomicSave)
        },
        "cfb_selective_mini_legacy_read" => Some(Case::CfbSelectiveMiniLegacyRead),
        "cfb_selective_mini_shared_read" => Some(Case::CfbSelectiveMiniSharedRead),
        "cfb_selective_fat_legacy_read" => Some(Case::CfbSelectiveFatLegacyRead),
        "cfb_selective_fat_shared_read" => Some(Case::CfbSelectiveFatSharedRead),
        "docx_source_backed_one_edit_save" => Some(Case::DocxSourceBackedOneEditSave),
        "pptx_source_backed_one_edit_save" => Some(Case::PptxSourceBackedOneEditSave),
        "pptx_eager_batch_edit_save" => Some(Case::PptxEagerBatchEditSave),
        "pptx_source_backed_batch_edit_save" => Some(Case::PptxSourceBackedBatchEditSave),
        "pptx_eager_multi_slide_batch_edit_save" => Some(Case::PptxEagerMultiSlideBatchEditSave),
        "pptx_source_backed_multi_slide_batch_edit_save" => {
            Some(Case::PptxSourceBackedMultiSlideBatchEditSave)
        },
        "xlsx_eager_calculation_metadata_edit_save" => {
            Some(Case::XlsxEagerCalculationMetadataEditSave)
        },
        "xlsx_source_backed_calculation_metadata_edit_save" => {
            Some(Case::XlsxSourceBackedCalculationMetadataEditSave)
        },
        "xlsx_eager_defined_names_edit_save" => Some(Case::XlsxEagerDefinedNamesEditSave),
        "xlsx_source_backed_defined_names_edit_save" => {
            Some(Case::XlsxSourceBackedDefinedNamesEditSave)
        },
        "xlsx_eager_page_break_edit_save" => Some(Case::XlsxEagerPageBreakEditSave),
        "xlsx_source_backed_page_break_edit_save" => Some(Case::XlsxSourceBackedPageBreakEditSave),
        "xlsx_eager_page_margin_edit_save" => Some(Case::XlsxEagerPageMarginEditSave),
        "xlsx_source_backed_page_margin_edit_save" => {
            Some(Case::XlsxSourceBackedPageMarginEditSave)
        },
        "xlsx_eager_page_setup_edit_save" => Some(Case::XlsxEagerPageSetupEditSave),
        "xlsx_source_backed_page_setup_edit_save" => Some(Case::XlsxSourceBackedPageSetupEditSave),
        "xlsx_eager_print_options_edit_save" => Some(Case::XlsxEagerPrintOptionsEditSave),
        "xlsx_source_backed_print_options_edit_save" => {
            Some(Case::XlsxSourceBackedPrintOptionsEditSave)
        },
        "xlsx_eager_sheet_protection_edit_save" => Some(Case::XlsxEagerSheetProtectionEditSave),
        "xlsx_source_backed_sheet_protection_edit_save" => {
            Some(Case::XlsxSourceBackedSheetProtectionEditSave)
        },
        "xlsx_eager_data_validation_edit_save" => Some(Case::XlsxEagerDataValidationEditSave),
        "xlsx_source_backed_data_validation_edit_save" => {
            Some(Case::XlsxSourceBackedDataValidationEditSave)
        },
        "xlsx_eager_auto_filter_edit_save" => Some(Case::XlsxEagerAutoFilterEditSave),
        "xlsx_source_backed_auto_filter_edit_save" => {
            Some(Case::XlsxSourceBackedAutoFilterEditSave)
        },
        "xlsx_eager_conditional_formatting_edit_save" => {
            Some(Case::XlsxEagerConditionalFormattingEditSave)
        },
        "xlsx_source_backed_conditional_formatting_edit_save" => {
            Some(Case::XlsxSourceBackedConditionalFormattingEditSave)
        },
        "xlsx_eager_merge_commit_save" => Some(Case::XlsxEagerMergeCommitSave),
        "xlsx_eager_unmerge_commit_save" => Some(Case::XlsxEagerUnmergeCommitSave),
        "xlsx_eager_cell_values_one_edit_save" => Some(Case::XlsxEagerCellValuesOneEditSave),
        "xlsx_source_backed_cell_values_one_edit_save" => {
            Some(Case::XlsxSourceBackedCellValuesOneEditSave)
        },
        "xlsx_eager_cell_values_one_percent_edit_save" => {
            Some(Case::XlsxEagerCellValuesOnePercentEditSave)
        },
        "xlsx_source_backed_cell_values_one_percent_edit_save" => {
            Some(Case::XlsxSourceBackedCellValuesOnePercentEditSave)
        },
        "xlsx_eager_cell_values_batch_edit_save" => Some(Case::XlsxEagerCellValuesBatchEditSave),
        "xlsx_source_backed_cell_values_batch_edit_save" => {
            Some(Case::XlsxSourceBackedCellValuesBatchEditSave)
        },
        "cfb_open" => Some(Case::CfbOpen),
        "cfb_list_streams" => Some(Case::CfbListStreams),
        "cfb_read_one" => Some(Case::CfbReadOne),
        "cfb_create_stream_borrowed" => Some(Case::CfbCreateStreamBorrowed),
        "cfb_create_stream_owned" => Some(Case::CfbCreateStreamOwned),
        "ole_common_open" => Some(Case::OleCommonOpen),
        "ole_common_put_stream_publish" => Some(Case::OleCommonPutStreamPublish),
        "ole_common_finish_render" => Some(Case::OleCommonFinishRender),
        "ole_common_one_edit_save" => Some(Case::OleCommonOneEditSave),
        "cfb_shared_open" => Some(Case::CfbSharedOpen),
        "cfb_shared_read_one" => Some(Case::CfbSharedReadOne),
        "cfb_shared_concurrent_reads" => Some(Case::CfbSharedConcurrentReads),
        "doc_fresh_write_to" => Some(Case::DocFreshWriteTo),
        "xls_fresh_write_to" => Some(Case::XlsFreshWriteTo),
        "ppt_fresh_write_to" => Some(Case::PptFreshWriteTo),
        "doc_semantic_open" => Some(Case::DocSemanticOpen),
        "doc_semantic_list_paragraphs" => Some(Case::DocSemanticListParagraphs),
        "doc_semantic_one_paragraph" => Some(Case::DocSemanticOneParagraph),
        "doc_semantic_full_text" => Some(Case::DocSemanticFullText),
        "doc_semantic_noop_edit_save" => Some(Case::DocSemanticNoopEditSave),
        "doc_semantic_one_edit_save" => Some(Case::DocSemanticOneEditSave),
        "doc_body_snapshot_list_paragraphs" => Some(Case::DocBodySnapshotListParagraphs),
        "xls_semantic_open" => Some(Case::XlsSemanticOpen),
        "xls_semantic_list_worksheets" => Some(Case::XlsSemanticListWorksheets),
        "xls_semantic_one_cell" => Some(Case::XlsSemanticOneCell),
        "xls_semantic_full_cell_scan" => Some(Case::XlsSemanticFullCellScan),
        "xls_semantic_noop_edit_save" => Some(Case::XlsSemanticNoopEditSave),
        "xls_semantic_one_edit_save" => Some(Case::XlsSemanticOneEditSave),
        "xls_validation_report" => Some(Case::XlsValidationReport),
        "xls_comments_eager_edit_save" => Some(Case::XlsCommentsEagerEditSave),
        "xls_comments_source_backed_edit_save" => Some(Case::XlsCommentsSourceBackedEditSave),
        "xls_comments_eager_batch_edit_save" => Some(Case::XlsCommentsEagerBatchEditSave),
        "xls_comments_source_backed_batch_edit_save" => {
            Some(Case::XlsCommentsSourceBackedBatchEditSave)
        },
        "xls_visibility_eager_edit_save" => Some(Case::XlsVisibilityEagerEditSave),
        "xls_visibility_source_backed_edit_save" => Some(Case::XlsVisibilitySourceBackedEditSave),
        "xls_visibility_eager_batch_edit_save" => Some(Case::XlsVisibilityEagerBatchEditSave),
        "xls_visibility_source_backed_batch_edit_save" => {
            Some(Case::XlsVisibilitySourceBackedBatchEditSave)
        },
        "ppt_semantic_open" => Some(Case::PptSemanticOpen),
        "ppt_semantic_list_slides" => Some(Case::PptSemanticListSlides),
        "ppt_semantic_one_shape_text" => Some(Case::PptSemanticOneShapeText),
        "ppt_semantic_full_text" => Some(Case::PptSemanticFullText),
        "ppt_slide_order_snapshot_open" => Some(Case::PptSlideOrderSnapshotOpen),
        "ppt_text_edit_one_edit_save" => Some(Case::PptTextEditOneEditSave),
        "ppt_semantic_noop_edit_save" => Some(Case::PptSemanticNoopEditSave),
        "ppt_semantic_one_edit_save" => Some(Case::PptSemanticOneEditSave),
        "xlsx_open_owned" => Some(Case::XlsxOpenOwned),
        "xlsx_list_sheets" => Some(Case::XlsxListSheets),
        "xlsx_first_cell" => Some(Case::XlsxFirstCell),
        "xlsx_full_cell_scan" => Some(Case::XlsxFullCellScan),
        "xlsx_narrow_column_range_scan" => Some(Case::XlsxNarrowColumnRangeScan),
        "xlsx_noop_commit" => Some(Case::XlsxNoopCommit),
        "xlsx_noop_commit_save" => Some(Case::XlsxNoopCommitSave),
        "xlsx_one_cell_commit" => Some(Case::XlsxOneCellCommit),
        "xlsx_one_cell_commit_first_read" => Some(Case::XlsxOneCellCommitFirstRead),
        "xlsx_one_cell_commit_save" => Some(Case::XlsxOneCellCommitSave),
        "xlsx_one_percent_commit" => Some(Case::XlsxOnePercentCommit),
        "xlsx_one_percent_commit_save" => Some(Case::XlsxOnePercentCommitSave),
        "xlsx_source_open" => Some(Case::XlsxSourceOpen),
        "xlsx_source_list_sheets" => Some(Case::XlsxSourceListSheets),
        "xlsx_source_first_cell" => Some(Case::XlsxSourceFirstCell),
        "xlsx_source_narrow_column_range_scan" => Some(Case::XlsxSourceNarrowColumnRangeScan),
        "xlsx_streaming_create" => Some(Case::XlsxStreamingCreate),
        "opc_range_source_open" => Some(Case::OpcRangeSourceOpen),
        "opc_range_source_open_main_read" => Some(Case::OpcRangeSourceOpenMainRead),
        "xlsx_range_source_open" => Some(Case::XlsxRangeSourceOpen),
        "xlsx_range_source_list_sheets" => Some(Case::XlsxRangeSourceListSheets),
        "xlsx_range_source_first_cell" => Some(Case::XlsxRangeSourceFirstCell),
        "xlsx_range_source_narrow_column_range_scan" => {
            Some(Case::XlsxRangeSourceNarrowColumnRangeScan)
        },
        "opc_open_session_scaling" => Some(Case::OpcOpenSessionScaling),
        "cfb_bulk_read_scaling" => Some(Case::CfbBulkReadScaling),
        "rtf_semantic_open" => Some(Case::RtfSemanticOpen),
        "rtf_semantic_paragraph_count" => Some(Case::RtfSemanticParagraphCount),
        "rtf_semantic_list_paragraphs" => Some(Case::RtfSemanticListParagraphs),
        "rtf_semantic_collect_paragraphs" => Some(Case::RtfSemanticCollectParagraphs),
        "rtf_semantic_one_paragraph" => Some(Case::RtfSemanticOneParagraph),
        "rtf_semantic_full_text" => Some(Case::RtfSemanticFullText),
        "rtf_semantic_text_to_sink" => Some(Case::RtfSemanticTextToSink),
        "rtf_semantic_stream_save" => Some(Case::RtfSemanticStreamSave),
        "rtf_semantic_noop_edit_save" => Some(Case::RtfSemanticNoopEditSave),
        "rtf_semantic_one_edit_save" => Some(Case::RtfSemanticOneEditSave),
        "rtf_semantic_one_percent_edit_save" => Some(Case::RtfSemanticOnePercentEditSave),
        "rtf_semantic_remove_paragraph_save" => Some(Case::RtfSemanticRemoveParagraphSave),
        "rtf_semantic_move_paragraph_save" => Some(Case::RtfSemanticMoveParagraphSave),
        "rtf_logical_tail_append" => Some(Case::RtfLogicalTailAppend),
        "rtf_logical_tail_noop_save" => Some(Case::RtfLogicalTailNoopSave),
        "rtf_validation_report" => Some(Case::RtfValidationReport),
        "rtf_streaming_create" => Some(Case::RtfStreamingCreate),
        "docx_semantic_open" => Some(Case::DocxSemanticOpen),
        "docx_semantic_list_paragraphs" => Some(Case::DocxSemanticListParagraphs),
        "docx_semantic_one_paragraph" => Some(Case::DocxSemanticOneParagraph),
        "docx_semantic_full_text" => Some(Case::DocxSemanticFullText),
        "docx_semantic_create_small" => Some(Case::DocxSemanticCreateSmall),
        "docx_semantic_noop_edit_save" => Some(Case::DocxSemanticNoopEditSave),
        "docx_semantic_one_edit_save" => Some(Case::DocxSemanticOneEditSave),
        "docx_semantic_one_percent_edit_save" => Some(Case::DocxSemanticOnePercentEditSave),
        "docx_validation_report" => Some(Case::DocxValidationReport),
        "docx_section_inventory" => Some(Case::DocxSectionInventory),
        "pptx_semantic_open" => Some(Case::PptxSemanticOpen),
        "pptx_semantic_list_slides" => Some(Case::PptxSemanticListSlides),
        "pptx_semantic_one_slide" => Some(Case::PptxSemanticOneSlide),
        "pptx_semantic_full_text" => Some(Case::PptxSemanticFullText),
        "pptx_semantic_create_small" => Some(Case::PptxSemanticCreateSmall),
        "pptx_semantic_noop_edit_save" => Some(Case::PptxSemanticNoopEditSave),
        "pptx_semantic_one_edit_save" => Some(Case::PptxSemanticOneEditSave),
        "pptx_semantic_one_percent_edit_save" => Some(Case::PptxSemanticOnePercentEditSave),
        "pptx_validation_report" => Some(Case::PptxValidationReport),
        "odt_semantic_open" => Some(Case::OdtSemanticOpen),
        "odt_semantic_list_paragraphs" => Some(Case::OdtSemanticListParagraphs),
        "odt_semantic_one_paragraph" => Some(Case::OdtSemanticOneParagraph),
        "odt_semantic_full_text" => Some(Case::OdtSemanticFullText),
        "odt_semantic_create_small" => Some(Case::OdtSemanticCreateSmall),
        "odt_semantic_noop_edit_save" => Some(Case::OdtSemanticNoopEditSave),
        "odt_semantic_one_edit_save" => Some(Case::OdtSemanticOneEditSave),
        "odt_semantic_one_percent_edit_save" => Some(Case::OdtSemanticOnePercentEditSave),
        "odf_validation_report" => Some(Case::OdfValidationReport),
        "odf_mimetype_repair_plan" => Some(Case::OdfMimetypeRepairPlan),
        "odt_media_paragraph_edit_save" => Some(Case::OdtMediaParagraphEditSave),
        "odt_media_line_break_edit_save" => Some(Case::OdtMediaLineBreakEditSave),
        "odt_media_append_run_edit_save" => Some(Case::OdtMediaAppendRunEditSave),
        "odt_media_append_hyperlink_edit_save" => Some(Case::OdtMediaAppendHyperlinkEditSave),
        "odt_media_insert_paragraph_edit_save" => Some(Case::OdtMediaInsertParagraphEditSave),
        "odt_media_remove_paragraph_edit_save" => Some(Case::OdtMediaRemoveParagraphEditSave),
        "odt_embedded_resource_scalar_replace_save" => {
            Some(Case::OdtEmbeddedResourceScalarReplaceSave)
        },
        "odt_embedded_resource_batch_replace_save" => {
            Some(Case::OdtEmbeddedResourceBatchReplaceSave)
        },
        "ods_semantic_open" => Some(Case::OdsSemanticOpen),
        "ods_semantic_list_sheets" => Some(Case::OdsSemanticListSheets),
        "ods_semantic_one_cell" => Some(Case::OdsSemanticOneCell),
        "ods_semantic_cell_sweep" => Some(Case::OdsSemanticCellSweep),
        "ods_semantic_full_cell_text" => Some(Case::OdsSemanticFullCellText),
        "ods_semantic_create_small" => Some(Case::OdsSemanticCreateSmall),
        "ods_semantic_noop_edit_save" => Some(Case::OdsSemanticNoopEditSave),
        "ods_semantic_one_edit_save" => Some(Case::OdsSemanticOneEditSave),
        "ods_semantic_one_percent_edit_save" => Some(Case::OdsSemanticOnePercentEditSave),
        "ods_media_one_edit_save" => Some(Case::OdsMediaOneEditSave),
        "odp_semantic_open" => Some(Case::OdpSemanticOpen),
        "odp_semantic_list_slides" => Some(Case::OdpSemanticListSlides),
        "odp_semantic_one_slide" => Some(Case::OdpSemanticOneSlide),
        "odp_semantic_full_text" => Some(Case::OdpSemanticFullText),
        "odp_semantic_create_small" => Some(Case::OdpSemanticCreateSmall),
        "odp_semantic_noop_edit_save" => Some(Case::OdpSemanticNoopEditSave),
        "odp_semantic_one_edit_save" => Some(Case::OdpSemanticOneEditSave),
        "odp_media_textbox_edit_save" => Some(Case::OdpMediaTextBoxEditSave),
        "odp_media_textbox_scalar_replace_save" => Some(Case::OdpMediaTextBoxScalarReplaceSave),
        "odp_media_textbox_batch_replace_save" => Some(Case::OdpMediaTextBoxBatchReplaceSave),
        _ => None,
    }
}

fn parse_shape(value: &str) -> Option<CorpusShape> {
    match value {
        "tiny" => Some(CorpusShape::Tiny),
        "many-small" => Some(CorpusShape::ManySmall),
        "few-large" => Some(CorpusShape::FewLarge),
        "wide-root" => Some(CorpusShape::WideRoot),
        _ => None,
    }
}

fn parse_payload(value: &str) -> Option<PayloadKind> {
    match value {
        "compressible" => Some(PayloadKind::Compressible),
        "incompressible" => Some(PayloadKind::Incompressible),
        _ => None,
    }
}

fn parse_writer_shape(value: &str) -> Option<WriterShape> {
    match value {
        "tiny" => Some(WriterShape::Tiny),
        "large" => Some(WriterShape::Large),
        "payload-heavy" => Some(WriterShape::PayloadHeavy),
        _ => None,
    }
}

fn parse_xlsx_shape(value: &str) -> Option<XlsxShape> {
    match value {
        "tiny" => Some(XlsxShape::Tiny),
        "medium" => Some(XlsxShape::Medium),
        "dense-wide" => Some(XlsxShape::DenseWide),
        _ => None,
    }
}

fn parse_xlsx_cell_crud_shape(value: &str) -> Option<XlsxCellCrudShape> {
    match value {
        "medium" => Some(XlsxCellCrudShape::Medium),
        "dense-sparse" => Some(XlsxCellCrudShape::DenseSparse),
        _ => None,
    }
}

fn parse_semantic_shape(value: &str) -> Option<SemanticShape> {
    match value {
        "tiny" => Some(SemanticShape::Tiny),
        "medium" => Some(SemanticShape::Medium),
        "large" => Some(SemanticShape::Large),
        _ => None,
    }
}

fn parse_rtf_variant(value: &str) -> Option<RtfSemanticVariant> {
    match value {
        "plain" => Some(RtfSemanticVariant::Plain),
        "byte1252" => Some(RtfSemanticVariant::Byte1252),
        "lzfu" => Some(RtfSemanticVariant::Lzfu),
        "watermark" => Some(RtfSemanticVariant::Watermark),
        _ => None,
    }
}

fn print_usage() {
    println!(
        "Usage: cargo run --release --manifest-path tools/perf-baseline/Cargo.toml -- [OPTIONS]\n\n\
         Options:\n\
           --samples N                 Samples per case (default: {DEFAULT_SAMPLES})\n\
           --warmup N                  Untimed iterations per case (default: {DEFAULT_WARMUP_ITERATIONS})\n\
           --filesystem-cache LIST     Filesystem states: warm,cold-requested\n\
           --filesystem-root PATH      Parent directory for filesystem samples\n\
           --case LIST                 zip_index,zip_read_one,opc_open,opc_open_owned,\n\
                                       opc_noop_save,opc_mutated_save,opc_source_open,\n\
                                       opc_source_open_main_read,opc_source_cached_main_read,\n\
                                       opc_source_concurrent_same_part,\n\
                                       opc_source_cache_budget_boundary,\n\
                                       opc_source_cache_control_contention,\n\
                                       opc_source_cache_managed_contention,\n\
                                       opc_source_overlay_one_part_save,\n\
                                       opc_file_eager_open,opc_file_source_open,\n\
                                       opc_file_eager_one_part_atomic_save,\n\
                                       opc_file_source_one_part_atomic_save,\n\
                                       cfb_file_same_length_overlay_atomic_save,\n\
                                       docx_source_backed_one_edit_save,\n\
                                       pptx_source_backed_one_edit_save,\n\
                                       pptx_eager_batch_edit_save,\n\
                                       pptx_source_backed_batch_edit_save,\n\
                                       pptx_eager_multi_slide_batch_edit_save,\n\
                                       pptx_source_backed_multi_slide_batch_edit_save,\n\
                                       xlsx_eager_calculation_metadata_edit_save,\n\
                                       xlsx_source_backed_calculation_metadata_edit_save,\n\
                                       xlsx_eager_defined_names_edit_save,\n\
                                       xlsx_source_backed_defined_names_edit_save,\n\
                                       xlsx_eager_page_break_edit_save,\n\
                                       xlsx_source_backed_page_break_edit_save,\n\
                                       xlsx_eager_page_margin_edit_save,\n\
                                       xlsx_source_backed_page_margin_edit_save,\n\
                                       xlsx_eager_page_setup_edit_save,\n\
                                       xlsx_source_backed_page_setup_edit_save,\n\
                                       xlsx_eager_print_options_edit_save,\n\
                                       xlsx_source_backed_print_options_edit_save,\n\
                                       xlsx_eager_sheet_protection_edit_save,\n\
                                       xlsx_source_backed_sheet_protection_edit_save,\n\
                                       xlsx_eager_data_validation_edit_save,\n\
                                       xlsx_source_backed_data_validation_edit_save,\n\
                                       xlsx_eager_auto_filter_edit_save,\n\
                                       xlsx_source_backed_auto_filter_edit_save,\n\
                                       xlsx_eager_conditional_formatting_edit_save,\n\
                                       xlsx_source_backed_conditional_formatting_edit_save,\n\
                                       xlsx_eager_merge_commit_save,\n\
                                       xlsx_eager_unmerge_commit_save,\n\
                                       cfb_open,cfb_list_streams,cfb_read_one,\n\
                                       cfb_create_stream_borrowed,cfb_create_stream_owned,\n\
                                       ole_common_open,ole_common_put_stream_publish,\n\
                                       ole_common_finish_render,\n\
                                       ole_common_one_edit_save,\n\
                                       cfb_shared_open,cfb_shared_read_one,\n\
                                       cfb_shared_concurrent_reads,\n\
                                       cfb_selective_mini_legacy_read,\n\
                                       cfb_selective_mini_shared_read,\n\
                                       cfb_selective_fat_legacy_read,\n\
                                       cfb_selective_fat_shared_read,\n\
                                       doc_fresh_write_to,xls_fresh_write_to,ppt_fresh_write_to,\n\
                                       doc_semantic_open,doc_semantic_list_paragraphs,\n\
                                       doc_semantic_one_paragraph,doc_semantic_full_text,\n\
                                       doc_semantic_noop_edit_save,doc_semantic_one_edit_save,\n\
                                       doc_body_snapshot_list_paragraphs,\n\
                                       xls_semantic_open,xls_semantic_list_worksheets,\n\
                                       xls_semantic_one_cell,xls_semantic_full_cell_scan,\n\
                                       xls_semantic_noop_edit_save,xls_semantic_one_edit_save,\n\
                                       xls_comments_eager_edit_save,\n\
                                       xls_comments_source_backed_edit_save,\n\
                                       xls_comments_eager_batch_edit_save,\n\
                                       xls_comments_source_backed_batch_edit_save,\n\
                                       xls_visibility_eager_edit_save,\n\
                                       xls_visibility_source_backed_edit_save,\n\
                                       xls_visibility_eager_batch_edit_save,\n\
                                       xls_visibility_source_backed_batch_edit_save,\n\
                                       ppt_semantic_open,ppt_semantic_list_slides,\n\
                                       ppt_semantic_one_shape_text,ppt_semantic_full_text,\n\
                                       ppt_slide_order_snapshot_open,\n\
                                       ppt_text_edit_one_edit_save,\n\
                                       ppt_semantic_noop_edit_save,ppt_semantic_one_edit_save,\n\
                                       xlsx_open_owned,xlsx_list_sheets,xlsx_first_cell,\n\
                                       xlsx_full_cell_scan,xlsx_narrow_column_range_scan,\n\
                                       xlsx_noop_commit,xlsx_noop_commit_save,\n\
                                       xlsx_one_cell_commit,xlsx_one_cell_commit_first_read,\n\
                                       xlsx_one_cell_commit_save,\n\
                                       xlsx_one_percent_commit,xlsx_one_percent_commit_save,\n\
                                       xlsx_eager_cell_values_one_edit_save,\n\
                                       xlsx_source_backed_cell_values_one_edit_save,\n\
                                       xlsx_eager_cell_values_one_percent_edit_save,\n\
                                       xlsx_source_backed_cell_values_one_percent_edit_save,\n\
                                       xlsx_eager_cell_values_batch_edit_save,\n\
                                       xlsx_source_backed_cell_values_batch_edit_save,\n\
                                       xlsx_source_open,xlsx_source_list_sheets,\n\
                                       xlsx_source_first_cell,\n\
                                       xlsx_source_narrow_column_range_scan,\n\
                                       xlsx_streaming_create,\n\
                                       opc_range_source_open,opc_range_source_open_main_read,\n\
                                       xlsx_range_source_open,xlsx_range_source_list_sheets,\n\
                                       xlsx_range_source_first_cell,\n\
                                       xlsx_range_source_narrow_column_range_scan,\n\
                                       opc_open_session_scaling,cfb_bulk_read_scaling,\n\
                                       rtf_semantic_open,rtf_semantic_paragraph_count,\n\
                                       rtf_semantic_list_paragraphs,rtf_semantic_collect_paragraphs,\n\
                                       rtf_semantic_one_paragraph,rtf_semantic_full_text,\n\
                                       rtf_semantic_text_to_sink,\n\
                                       rtf_semantic_stream_save,rtf_semantic_noop_edit_save,\n\
                                       rtf_semantic_one_edit_save,rtf_semantic_one_percent_edit_save,\n\
                                       rtf_semantic_remove_paragraph_save,rtf_semantic_move_paragraph_save,\n\
                                       rtf_logical_tail_append,rtf_logical_tail_noop_save,\n\
                                       rtf_streaming_create,\n\
                                       docx_semantic_open,docx_semantic_list_paragraphs,\n\
                                       docx_semantic_one_paragraph,docx_semantic_full_text,\n\
                                       docx_semantic_create_small,docx_semantic_noop_edit_save,\n\
                                       docx_semantic_one_edit_save,docx_semantic_one_percent_edit_save,\n\
                                       pptx_semantic_open,pptx_semantic_list_slides,\n\
                                       pptx_semantic_one_slide,pptx_semantic_full_text,\n\
                                       pptx_semantic_create_small,pptx_semantic_noop_edit_save,\n\
                                       pptx_semantic_one_edit_save,pptx_semantic_one_percent_edit_save,\n\
                                       odt_semantic_open,odt_semantic_list_paragraphs,\n\
                                       odt_semantic_one_paragraph,odt_semantic_full_text,\n\
                                       odt_semantic_create_small,odt_semantic_noop_edit_save,\n\
                                       odt_semantic_one_edit_save,odt_semantic_one_percent_edit_save,\n\
                                       odt_media_paragraph_edit_save,odt_media_line_break_edit_save,\n\
                                       odt_media_append_run_edit_save,\n\
                                       odt_media_append_hyperlink_edit_save,\n\
                                       odt_media_insert_paragraph_edit_save,\n\
                                       odt_media_remove_paragraph_edit_save,\n\
                                       odt_embedded_resource_scalar_replace_save,\n\
                                       odt_embedded_resource_batch_replace_save,\n\
                                       ods_semantic_open,\n\
                                       ods_semantic_list_sheets,ods_semantic_one_cell,\n\
                                       ods_semantic_cell_sweep,\n\
                                       ods_semantic_full_cell_text,ods_semantic_create_small,\n\
                                       ods_semantic_noop_edit_save,ods_semantic_one_edit_save,\n\
                                       ods_semantic_one_percent_edit_save,\n\
                                       ods_media_one_edit_save,\n\
                                       odp_semantic_open,odp_semantic_list_slides,\n\
                                       odp_semantic_one_slide,odp_semantic_full_text,\n\
                                       odp_semantic_create_small,odp_semantic_noop_edit_save,\n\
                                       odp_semantic_one_edit_save,odp_media_textbox_edit_save,\n\
                                       odp_media_textbox_scalar_replace_save,\n\
                                       odp_media_textbox_batch_replace_save\n\
           --shape LIST                tiny,many-small,few-large,wide-root\n\
           --payload LIST              compressible,incompressible\n\
           --writer-shape LIST         tiny,large,payload-heavy\n\
           --xlsx-shape LIST           tiny,medium,dense-wide\n\
           --xlsx-cell-crud-shape LIST medium,dense-sparse (only used by matched scalar-cell cases)\n\
           --semantic-shape LIST       tiny,medium,large (only used by opt-in Office semantic cases)\n\
           --rtf-variant LIST          plain,byte1252,lzfu,watermark (default: plain)\n\
           --range-fixed-latency-us N  Fixed latency per request (default: {DEFAULT_RANGE_FIXED_LATENCY_US})\n\
           --range-request-overhead-us N\n\
                                       Request overhead (default: {DEFAULT_RANGE_REQUEST_OVERHEAD_US})\n\
           --range-bandwidth-bytes-per-sec N\n\
                                       Bandwidth (default: {DEFAULT_RANGE_BANDWIDTH_BYTES_PER_SECOND})\n\
           --range-max-physical-bytes N\n\
                                       Maximum physical range (default: {DEFAULT_RANGE_MAX_PHYSICAL_BYTES})\n\
           --workers LIST              Scaling workers: 1,2,4,8,available (capped/deduped)\n\
           --json PATH                 Write JSON to PATH; use - or omit for stdout\n\
           --help                      Show this help"
    );
}

fn build_opc_corpus(
    shape: CorpusShape,
    payload_kind: PayloadKind,
) -> Result<Corpus, Box<dyn Error>> {
    let entry_count = shape.entry_count();
    let entry_bytes = shape.entry_bytes();
    let target_index = entry_count / 2;
    let target_name = entry_name(target_index);
    let mut package = OpcPackage::new();
    let mut target_payload = None;

    for index in 0..entry_count {
        let name = entry_name(index);
        let payload = payload_bytes(payload_kind, index, entry_bytes);
        if index == target_index {
            target_payload = Some(payload.clone());
        }
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{name}"))?,
            CONTENT_TYPE.to_owned(),
            payload,
        )))?;
    }
    package.rels_mut().try_add_relationship(
        relationship_type::OFFICE_DOCUMENT.to_owned(),
        target_name.clone(),
        "rIdBenchmarkMain".to_owned(),
        TargetMode::Internal,
    )?;

    let archive = PackageWriter::to_bytes(&package)?;
    let archive_member_count = ArchiveReader::new(&archive)?.file_names().count();
    let target_payload = target_payload.ok_or("generated corpus has no target entry")?;
    let entry_count_u64 = u64::try_from(entry_count)?;
    let entry_bytes_u64 = u64::try_from(entry_bytes)?;
    let uncompressed_payload_bytes = entry_count_u64
        .checked_mul(entry_bytes_u64)
        .ok_or("generated corpus payload byte count overflow")?;
    let name = format!("{}-{}", shape.name(), payload_kind.name());

    Ok(Corpus {
        manifest: CorpusManifest {
            name,
            generator: OPC_CORPUS_GENERATOR,
            package_format: "OPC/ZIP",
            shape: shape.name(),
            payload_kind: payload_kind.name(),
            compression: "deflate",
            entry_count,
            archive_member_count,
            entry_bytes,
            uncompressed_payload_bytes: usize::try_from(uncompressed_payload_bytes)?,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: target_name.clone(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name,
        target_payload,
        xlsx: None,
    })
}

fn build_cfb_corpus(
    shape: CorpusShape,
    payload_kind: PayloadKind,
) -> Result<Corpus, Box<dyn Error>> {
    let entry_count = shape.entry_count();
    let entry_bytes = shape.entry_bytes();
    let target_index = entry_count
        .checked_sub(1)
        .ok_or("generated CFB corpus has no target stream")?;
    let target_name = cfb_entry_name(target_index);
    let target_payload = payload_bytes(payload_kind, target_index, entry_bytes);
    let mut writer = OleWriter::new();

    for index in 0..entry_count {
        let name = cfb_entry_name(index);
        let payload = payload_bytes(payload_kind, index, entry_bytes);
        writer.create_stream_owned(&[name.as_str()], payload)?;
    }

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let archive = output.into_inner();
    let mut parsed = OleFile::open(Cursor::new(archive.as_slice()))?;
    let archive_member_count = parsed.list_streams().len();
    if archive_member_count != entry_count {
        return Err("CFB stream count differs from generated corpus specification".into());
    }
    if parsed.open_stream(&[target_name.as_str()])? != target_payload {
        return Err("CFB target differs from deterministic corpus payload".into());
    }

    let uncompressed_payload_bytes = entry_count
        .checked_mul(entry_bytes)
        .ok_or("generated CFB corpus payload byte count overflow")?;

    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("cfb-{}-{}", shape.name(), payload_kind.name()),
            generator: CFB_CORPUS_GENERATOR,
            package_format: "CFB/OLE2",
            shape: shape.name(),
            payload_kind: payload_kind.name(),
            compression: "none",
            entry_count,
            archive_member_count,
            entry_bytes,
            uncompressed_payload_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: target_name.clone(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name,
        target_payload,
        xlsx: None,
    })
}

fn build_cfb_selective_corpus(
    shape: CorpusShape,
    target: CfbSelectiveTarget,
) -> Result<Corpus, Box<dyn Error>> {
    let entry_count: usize = match shape {
        CorpusShape::ManySmall => 256,
        CorpusShape::WideRoot => 2048,
        _ => return Err("CFB selective corpus requires 256 or 2048 siblings".into()),
    };
    let base_entry_bytes = 1024usize;
    let target_index = entry_count
        .checked_sub(1)
        .ok_or("CFB selective corpus has no target stream")?;
    let target_payload = payload_bytes(
        PayloadKind::Incompressible,
        900_000 + target_index,
        target.target_bytes(),
    );
    let mut writer = OleWriter::new();
    for index in 0..entry_count {
        let payload = if index == target_index {
            target_payload.clone()
        } else {
            payload_bytes(PayloadKind::Incompressible, index, base_entry_bytes)
        };
        let name = cfb_entry_name(index);
        writer.create_stream_owned(&[name.as_str()], payload)?;
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let archive = output.into_inner();
    let target_name = cfb_entry_name(target_index);
    let mut parsed = OleFile::open(Cursor::new(archive.as_slice()))?;
    if parsed.list_streams().len() != entry_count
        || parsed.open_stream(&[target_name.as_str()])? != target_payload
    {
        return Err("CFB selective corpus failed deterministic stream validation".into());
    }
    let uncompressed_payload_bytes = (entry_count - 1)
        .checked_mul(base_entry_bytes)
        .and_then(|bytes| bytes.checked_add(target_payload.len()))
        .ok_or("CFB selective corpus payload byte count overflows usize")?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("cfb-selective-{}-{}", target.name(), shape.name()),
            generator: CFB_SELECTIVE_CORPUS_GENERATOR,
            package_format: "CFB/OLE2",
            shape: shape.name(),
            payload_kind: "incompressible",
            compression: "none",
            entry_count,
            archive_member_count: entry_count,
            entry_bytes: base_entry_bytes,
            uncompressed_payload_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: target_name,
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: cfb_entry_name(target_index),
        target_payload,
        xlsx: None,
    })
}

fn build_ole_common_corpus(base: &Corpus) -> Result<Corpus, Box<dyn Error>> {
    let kind = corpus_payload_kind(base)?;
    let unchanged_stream_count = base.manifest.entry_count;
    let entry_count = unchanged_stream_count
        .checked_add(1)
        .ok_or("OLE common corpus stream count overflow")?;
    let uncompressed_payload_bytes = base
        .manifest
        .uncompressed_payload_bytes
        .checked_add(OLE_COMMON_ORIGINAL.len())
        .ok_or("OLE common corpus payload byte count overflow")?;
    let mut writer = OleWriter::new();
    for index in 0..unchanged_stream_count {
        let name = cfb_entry_name(index);
        writer.create_stream_owned(
            &[name.as_str()],
            payload_bytes(kind, index, base.manifest.entry_bytes),
        )?;
    }
    writer.create_stream(&[OLE_COMMON_TARGET], OLE_COMMON_ORIGINAL)?;

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let archive = output.into_inner();
    let mut parsed = OleFile::open(Cursor::new(archive.as_slice()))?;
    if parsed.list_streams().len() != entry_count
        || parsed.open_stream(&[OLE_COMMON_TARGET])? != OLE_COMMON_ORIGINAL
    {
        return Err("OLE common corpus differs from its specification".into());
    }

    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("ole-common-{}-{}", base.manifest.shape, kind.name()),
            generator: OLE_COMMON_CORPUS_GENERATOR,
            package_format: "CFB/OLE2",
            shape: base.manifest.shape,
            payload_kind: kind.name(),
            compression: "none",
            entry_count,
            archive_member_count: entry_count,
            entry_bytes: base.manifest.entry_bytes,
            uncompressed_payload_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: OLE_COMMON_TARGET.to_owned(),
            target_payload_bytes: OLE_COMMON_ORIGINAL.len(),
            target_payload_sha256: sha256_hex(OLE_COMMON_ORIGINAL),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: OLE_COMMON_TARGET.to_owned(),
        target_payload: OLE_COMMON_ORIGINAL.to_vec(),
        xlsx: None,
    })
}

fn build_writer_corpus(case: Case, shape: WriterShape) -> Result<Corpus, Box<dyn Error>> {
    let (archive, entry_count, uncompressed_payload_bytes, package_format, target_name) = match case
    {
        Case::DocFreshWriteTo => {
            let (archive, entry_count, content_bytes) = write_fresh_doc(shape)?;
            (
                archive,
                entry_count,
                content_bytes,
                "DOC/CFB",
                "WordDocument",
            )
        },
        Case::XlsFreshWriteTo => {
            let (archive, entry_count, content_bytes) = write_fresh_xls(shape)?;
            (archive, entry_count, content_bytes, "XLS/CFB", "Workbook")
        },
        Case::PptFreshWriteTo => {
            let (archive, entry_count, content_bytes) = write_fresh_ppt(shape)?;
            (
                archive,
                entry_count,
                content_bytes,
                "PPT/CFB",
                "PowerPoint Document",
            )
        },
        _ => return Err("synthetic container case does not have a writer corpus".into()),
    };
    let mut parsed = OleFile::open(Cursor::new(archive.as_slice()))?;
    let archive_member_count = parsed.list_streams().len();
    let target_payload = parsed.open_stream(&[target_name])?;
    if archive_member_count == 0 || target_payload.is_empty() {
        return Err("fresh writer corpus has no packaged target stream".into());
    }

    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!(
                "{}-{}",
                case.name().trim_end_matches("_fresh_write_to"),
                shape.name()
            ),
            generator: LEGACY_WRITER_CORPUS_GENERATOR,
            package_format,
            shape: shape.name(),
            payload_kind: "not-applicable",
            compression: "none",
            entry_count,
            archive_member_count,
            entry_bytes: 0,
            uncompressed_payload_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: target_name.to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: target_name.to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn write_fresh_doc(shape: WriterShape) -> Result<(Vec<u8>, usize, usize), Box<dyn Error>> {
    let paragraph_count = shape.doc_paragraph_count();
    let payload_bytes = (shape == WriterShape::PayloadHeavy).then_some(20_000);
    let mut writer = litchi_doc::writer::Writer::new();
    let mut content_bytes = 0;
    for paragraph in 0..paragraph_count {
        let text = payload_bytes.map_or_else(
            || writer_text("doc", 0, paragraph, 0),
            |length| writer_payload_text("doc", 0, paragraph, 0, length),
        );
        content_bytes += text.len();
        writer.add_paragraph(&text)?;
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok((output.into_inner(), paragraph_count, content_bytes))
}

fn write_fresh_xls(shape: WriterShape) -> Result<(Vec<u8>, usize, usize), Box<dyn Error>> {
    let mut writer = litchi_xls::writer::Writer::new();
    let mut content_bytes = 0;
    let mut cell_count = 0;

    match shape {
        WriterShape::Tiny | WriterShape::Large => {
            let (sheet_count, row_count, column_count) = shape
                .xls_dimensions()
                .ok_or("non-numeric writer shape reached numeric XLS corpus")?;
            for sheet in 0..sheet_count {
                let worksheet = writer.add_worksheet(&format!("Bench{sheet:02}"))?;
                for row in 0..row_count {
                    for column in 0..column_count {
                        let value =
                            (sheet * row_count * column_count + row * column_count + column) as f64;
                        content_bytes += std::mem::size_of_val(&value);
                        cell_count += 1;
                        writer.write_number(worksheet, row as u32, column as u16, value)?;
                    }
                }
            }
        },
        WriterShape::PayloadHeavy => {
            // `WritableWorksheet` stores cells in a hash map, so use one
            // string cell per worksheet. Worksheets are serialized in
            // insertion order and one-cell maps traverse deterministically.
            // That keeps exact output checks while placing ~4 MiB of legal,
            // distinct 32,700-byte strings in the shared-string table.
            for sheet in 0..128 {
                let worksheet = writer.add_worksheet(&format!("Payload{sheet:03}"))?;
                let text = writer_payload_text("xls", sheet, 0, 0, 32_700);
                content_bytes += text.len();
                cell_count += 1;
                writer.write_string(worksheet, 0, 0, &text)?;
            }
        },
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok((output.into_inner(), cell_count, content_bytes))
}

fn write_fresh_ppt(shape: WriterShape) -> Result<(Vec<u8>, usize, usize), Box<dyn Error>> {
    let (slide_count, boxes_per_slide) = shape.ppt_dimensions();
    let payload_bytes = (shape == WriterShape::PayloadHeavy).then_some(40_000);
    let mut writer = litchi_ppt::writer::Writer::new();
    let mut content_bytes = 0;
    let mut text_box_count = 0;
    for slide_number in 0..slide_count {
        let slide = writer.add_slide()?;
        for box_number in 0..boxes_per_slide {
            let text = payload_bytes.map_or_else(
                || writer_text("ppt", slide_number, box_number, 0),
                |length| writer_payload_text("ppt", slide_number, box_number, 0, length),
            );
            content_bytes += text.len();
            text_box_count += 1;
            let x = 36 + (box_number % 3) as i32 * 180;
            let y = 36 + (box_number / 3) as i32 * 90;
            writer.add_textbox(slide, x, y, 144, 54, &text)?;
        }
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok((output.into_inner(), text_box_count, content_bytes))
}

fn semantic_shape(corpus: &Corpus) -> Result<SemanticShape, Box<dyn Error>> {
    match corpus.manifest.shape {
        "tiny" => Ok(SemanticShape::Tiny),
        "medium" => Ok(SemanticShape::Medium),
        "large" => Ok(SemanticShape::Large),
        _ => Err("semantic corpus has an unknown shape".into()),
    }
}

fn semantic_rtf_variant(corpus: &Corpus) -> Result<RtfSemanticVariant, Box<dyn Error>> {
    match corpus.manifest.rtf_variant {
        Some("plain") => Ok(RtfSemanticVariant::Plain),
        Some("byte1252") => Ok(RtfSemanticVariant::Byte1252),
        Some("lzfu") => Ok(RtfSemanticVariant::Lzfu),
        Some("watermark") => Ok(RtfSemanticVariant::Watermark),
        Some(_) => Err("semantic RTF corpus has an unknown transport variant".into()),
        None => Err("semantic RTF corpus has no transport variant".into()),
    }
}

fn writer_shape(corpus: &Corpus) -> Result<WriterShape, Box<dyn Error>> {
    match corpus.manifest.shape {
        "tiny" => Ok(WriterShape::Tiny),
        "large" => Ok(WriterShape::Large),
        "payload-heavy" => Ok(WriterShape::PayloadHeavy),
        _ => Err("legacy writer corpus has an unknown writer shape".into()),
    }
}

fn semantic_docx_text(index: usize, updated: bool) -> String {
    let state = if updated { "updated" } else { "source" };
    format!("litchi-perf-baseline-docx-semantic-v1-{state}-{index:05}")
}

fn semantic_pptx_text(slide: usize, shape: usize, updated: bool) -> String {
    let state = if updated { "updated" } else { "source" };
    format!("litchi-perf-baseline-pptx-semantic-v1-{state}-{slide:03}-{shape:03}")
}

fn semantic_update_indices(count: usize) -> Result<Vec<usize>, Box<dyn Error>> {
    if count == 0 {
        return Err("semantic corpus has no editable objects".into());
    }
    let updates = count
        .checked_add(99)
        .ok_or("semantic update count overflows usize")?
        / 100;
    Ok((0..updates).map(|index| index * count / updates).collect())
}

fn semantic_rtf_text(index: usize, updated: bool) -> String {
    let state = if updated { "updated" } else { "source" };
    format!("litchi-perf-baseline-rtf-semantic-v1-{state}-{index:05}")
}

fn semantic_rtf_variant_text(variant: RtfSemanticVariant, index: usize, updated: bool) -> String {
    match variant {
        RtfSemanticVariant::Plain | RtfSemanticVariant::Lzfu => semantic_rtf_text(index, updated),
        RtfSemanticVariant::Byte1252 => {
            let state = if updated { "updated" } else { "source" };
            format!("litchi-perf-baseline-rtf-byte1252-{state}-{index:05}-caf\u{e9}")
        },
        RtfSemanticVariant::Watermark => String::new(),
    }
}

fn semantic_rtf_paragraph_count(shape: SemanticShape, variant: RtfSemanticVariant) -> usize {
    match variant {
        RtfSemanticVariant::Watermark => 1,
        RtfSemanticVariant::Plain | RtfSemanticVariant::Byte1252 | RtfSemanticVariant::Lzfu => {
            shape.rtf_paragraphs()
        },
    }
}

fn semantic_rtf_expected_text(
    shape: SemanticShape,
    variant: RtfSemanticVariant,
    updated: &[usize],
) -> String {
    if variant == RtfSemanticVariant::Watermark {
        return "\n".to_owned();
    }
    (0..semantic_rtf_paragraph_count(shape, variant))
        .map(|index| {
            semantic_rtf_variant_text(variant, index, updated.binary_search(&index).is_ok())
        })
        .collect::<Vec<_>>()
        .join("\n")
}

fn semantic_rtf_plain_bytes(shape: SemanticShape) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut source = String::from(r"{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss Arial;}}\f0\fs20 ");
    for index in 0..shape.rtf_paragraphs() {
        if index != 0 {
            source.push_str(r"\par ");
        }
        source.push_str(&semantic_rtf_text(index, false));
    }
    source.push('}');

    let document = litchi_rtf::Document::parse(&source)?;
    let bytes = document.to_bytes()?;
    if bytes != source.as_bytes() {
        return Err("semantic RTF generator lost exact source identity".into());
    }
    Ok(bytes)
}

fn semantic_rtf_bytes(
    shape: SemanticShape,
    variant: RtfSemanticVariant,
) -> Result<Vec<u8>, Box<dyn Error>> {
    let archive = match variant {
        RtfSemanticVariant::Plain => semantic_rtf_plain_bytes(shape)?,
        RtfSemanticVariant::Byte1252 => {
            let mut source =
                br"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss Arial;}}\f0\fs20 ".to_vec();
            for index in 0..shape.rtf_paragraphs() {
                if index != 0 {
                    source.extend_from_slice(br"\par ");
                }
                let prefix = format!("litchi-perf-baseline-rtf-byte1252-source-{index:05}-caf");
                source.extend_from_slice(prefix.as_bytes());
                source.push(0xe9);
            }
            source.push(b'}');
            source
        },
        RtfSemanticVariant::Lzfu => {
            let plain = semantic_rtf_plain_bytes(shape)?;
            let compressed = litchi_rtf::transport::compress(&plain, true)?;
            if !litchi_rtf::transport::is_compressed_rtf(&compressed)
                || litchi_rtf::transport::decompress(&compressed)? != plain
            {
                return Err("semantic RTF LZFu transport verification failed".into());
            }
            compressed
        },
        RtfSemanticVariant::Watermark => {
            if shape != SemanticShape::Tiny {
                return Err("semantic RTF watermark corpus is only available for tiny".into());
            }
            let source = include_bytes!("../../../test-data/rtf/watermark.rtf").to_vec();
            if sha256_hex(&source)
                != "48d62dcd959e737b06ebb8255780bcaaf1e88056ff9c3d5a21d3ff5cd3ddf9cb"
            {
                return Err("semantic RTF watermark fixture hash differs from inventory".into());
            }
            source
        },
    };

    let document = litchi_rtf::Document::from_bytes(&archive)?;
    if document.to_bytes()? != archive {
        return Err("semantic RTF variant lost exact source identity".into());
    }
    Ok(archive)
}

fn build_semantic_rtf_corpus(
    shape: SemanticShape,
    variant: RtfSemanticVariant,
) -> Result<Corpus, Box<dyn Error>> {
    let archive = semantic_rtf_bytes(shape, variant)?;
    let document = litchi_rtf::Document::from_bytes(&archive)?;
    verify_semantic_rtf(&document, shape, variant, &[])?;
    let target_payload = semantic_rtf_variant_text(variant, 0, false).into_bytes();
    let content_bytes = semantic_rtf_expected_text(shape, variant, &[]).len();
    let name = if variant == RtfSemanticVariant::Watermark {
        "rtf-semantic-watermark".to_owned()
    } else {
        format!("rtf-semantic-{}-{}", variant.name(), shape.name())
    };
    Ok(Corpus {
        manifest: CorpusManifest {
            name,
            generator: SEMANTIC_RTF_CORPUS_GENERATOR,
            package_format: "RTF",
            shape: shape.name(),
            payload_kind: match variant {
                RtfSemanticVariant::Plain => "deterministic-semantic-text",
                RtfSemanticVariant::Byte1252 => "deterministic-byte1252-text",
                RtfSemanticVariant::Lzfu => "deterministic-semantic-text",
                RtfSemanticVariant::Watermark => "producer-watermark-drawing",
            },
            compression: if variant == RtfSemanticVariant::Lzfu {
                "lzfu"
            } else {
                "none"
            },
            entry_count: semantic_rtf_paragraph_count(shape, variant),
            archive_member_count: 1,
            entry_bytes: target_payload.len(),
            uncompressed_payload_bytes: content_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "paragraph:0".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: Some(variant.name()),
            xlsx: None,
        },
        archive,
        target_name: "paragraph:0".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn build_rtf_lifecycle_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let mut source = String::from(r"{\rtf1\ansi ");
    for index in 0..shape.rtf_paragraphs() {
        if index != 0 {
            source.push_str(r"\par ");
        }
        source.push_str(&semantic_rtf_text(index, false));
    }
    source.push('}');
    let archive = source.into_bytes();
    let document = litchi_rtf::Document::from_bytes(&archive)?;
    if document.to_bytes()? != archive {
        return Err("RTF lifecycle corpus lost exact source identity".into());
    }
    verify_semantic_rtf(&document, shape, RtfSemanticVariant::Plain, &[])?;
    let target_payload = semantic_rtf_text(0, false).into_bytes();
    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("rtf-paragraph-lifecycle-plain-{}", shape.name()),
            generator: RTF_LIFECYCLE_CORPUS_GENERATOR,
            package_format: "RTF",
            shape: shape.name(),
            payload_kind: "deterministic-default-formatted-text",
            compression: "none",
            entry_count: shape.rtf_paragraphs(),
            archive_member_count: 1,
            entry_bytes: target_payload.len(),
            uncompressed_payload_bytes: semantic_rtf_expected_text(
                shape,
                RtfSemanticVariant::Plain,
                &[],
            )
            .len(),
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "paragraph:0".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: Some(RtfSemanticVariant::Plain.name()),
            xlsx: None,
        },
        archive,
        target_name: "paragraph:0".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn rtf_logical_tail_paragraph_count(shape: SemanticShape) -> usize {
    match shape {
        SemanticShape::Tiny => 4,
        SemanticShape::Medium => 64,
        SemanticShape::Large => 256,
    }
}

fn rtf_logical_tail_text(shape: SemanticShape, index: usize) -> String {
    format!(
        "litchi-perf-baseline-rtf-tail-v1-{}-{index:04}",
        shape.name()
    )
}

fn rtf_logical_tail_limits(
    source_bytes: usize,
    input_bytes: usize,
    paragraph_count: usize,
) -> Result<litchi_rtf::TailAppendLimits, Box<dyn Error>> {
    let paragraph_overhead = paragraph_count
        .checked_mul(32)
        .ok_or("RTF logical-tail paragraph bound overflows usize")?;
    let inserted_bound = input_bytes
        .checked_add(paragraph_overhead)
        .and_then(|value| value.checked_add(1024))
        .ok_or("RTF logical-tail inserted-byte bound overflows usize")?;
    let output_bound = source_bytes
        .checked_add(inserted_bound)
        .ok_or("RTF logical-tail output bound overflows usize")?;
    let patch_bound = inserted_bound
        .checked_mul(2)
        .and_then(|value| value.checked_add(4096))
        .ok_or("RTF logical-tail patch bound overflows usize")?;
    Ok(litchi_rtf::TailAppendLimits::new(
        paragraph_count,
        paragraph_count,
        input_bytes,
        inserted_bound,
        output_bound,
        patch_bound,
    ))
}

fn stage_rtf_logical_tail(
    source: &litchi_rtf::Document,
    paragraphs: &[&str],
    limits: litchi_rtf::TailAppendLimits,
) -> Result<litchi_rtf::TailAppendCommit, Box<dyn Error>> {
    let mut edit = source.tail_append_with_limits(litchi_rtf::TailSelector::Body, limits);
    edit.append_text_paragraphs(paragraphs)?;
    Ok(edit.commit()?)
}

fn verify_rtf_logical_tail_projection(
    document: &litchi_rtf::Document,
    source: &litchi_rtf::Document,
    appended: &[String],
) -> Result<(), Box<dyn Error>> {
    let mut expected = source
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.to_text())
        .collect::<Vec<_>>();
    expected.extend(appended.iter().cloned());
    let actual = document
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.to_text())
        .collect::<Vec<_>>();
    if actual != expected || document.paragraph_count() != expected.len() {
        return Err(
            "RTF logical-tail reopen paragraph projection differs from specification".into(),
        );
    }
    let mut expected_text = expected.join("\n");
    if !expected.is_empty() {
        expected_text.push('\n');
    }
    if document.text() != expected_text {
        return Err("RTF logical-tail reopen text differs from specification".into());
    }
    Ok(())
}

fn verify_rtf_logical_tail_gates(
    source: &litchi_rtf::Document,
    changed: &litchi_rtf::TailAppendCommit,
    noop: &litchi_rtf::TailAppendCommit,
    appended: &[String],
    limits: litchi_rtf::TailAppendLimits,
    expected: &[u8],
) -> Result<(), Box<dyn Error>> {
    if !changed.diagnostics().changed()
        || changed.diagnostics().operation_count() != 1
        || changed.diagnostics().paragraphs() != appended.len()
        || changed.diagnostics().runs() != appended.len()
    {
        return Err("RTF logical-tail append diagnostics differ from specification".into());
    }
    if noop.diagnostics().changed()
        || noop.diagnostics().operation_count() != 0
        || !noop.snapshot().same_snapshot(source)
    {
        return Err("RTF logical-tail empty append was not an exact no-op".into());
    }

    let mut published = Vec::new();
    changed.write_to(&mut published, limits)?;
    if published != expected {
        return Err("RTF logical-tail sequential publication differs from candidate".into());
    }
    verify_rtf_logical_tail_projection(
        &litchi_rtf::Document::from_bytes(expected)?,
        source,
        appended,
    )?;

    let applied = changed.patch().apply(source)?;
    if applied.to_bytes()? != expected {
        return Err("RTF logical-tail in-memory patch replay differs from publication".into());
    }
    let restored = changed.patch().inverse().apply(&applied)?;
    if restored.to_bytes()? != source.to_bytes()? {
        return Err("RTF logical-tail in-memory inverse did not restore exact source".into());
    }

    let durable = changed.patch().to_durable(limits)?;
    let encoded = durable.to_deterministic_json()?;
    let decoded = litchi_rtf::DurableTailAppendPatch::from_deterministic_json(&encoded, limits)?;
    let durable_applied = decoded.apply(source)?;
    if durable_applied.to_bytes()? != expected {
        return Err("RTF logical-tail durable replay differs from publication".into());
    }
    let inverse_encoded = decoded.inverse().to_deterministic_json()?;
    let inverse =
        litchi_rtf::DurableTailAppendPatch::from_deterministic_json(&inverse_encoded, limits)?;
    let durable_restored = inverse.apply(&durable_applied)?;
    if durable_restored.to_bytes()? != source.to_bytes()? {
        return Err("RTF logical-tail durable inverse did not restore exact source".into());
    }

    let mut noop_output = Vec::new();
    noop.write_to(&mut noop_output, limits)?;
    if noop_output != source.to_bytes()? {
        return Err("RTF logical-tail no-op publication changed source bytes".into());
    }
    let noop_durable = noop.patch().to_durable(limits)?;
    let noop_json = noop_durable.to_deterministic_json()?;
    let noop_decoded =
        litchi_rtf::DurableTailAppendPatch::from_deterministic_json(&noop_json, limits)?;
    let noop_applied = noop_decoded.apply(source)?;
    if !noop_applied.same_snapshot(source) || noop_applied.to_bytes()? != source.to_bytes()? {
        return Err("RTF logical-tail durable no-op lost exact source identity".into());
    }

    let foreign = litchi_rtf::Document::parse(r"{\rtf1\ansi foreign source}")?;
    if !matches!(
        changed.patch().apply(&foreign),
        Err(litchi_rtf::TailAppendError::PatchConflict)
    ) || !matches!(
        decoded.apply(&foreign),
        Err(litchi_rtf::TailAppendError::PatchConflict)
    ) {
        return Err("RTF logical-tail source-conflict gate accepted a foreign source".into());
    }
    Ok(())
}

fn semantic_docx_bytes(shape: SemanticShape) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut package = litchi_docx::Package::new()?;
    let document = package.document_mut()?;
    for index in 0..shape.docx_paragraphs() {
        document.add_paragraph_with_text(&semantic_docx_text(index, false));
    }
    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output)?;
    Ok(output.into_inner())
}

fn build_semantic_docx_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let archive = semantic_docx_bytes(shape)?;
    let package = litchi_docx::Package::from_reader(Cursor::new(archive.clone()))?;
    verify_semantic_docx(&package, shape, &[])?;
    let archive_member_count = ArchiveReader::new(&archive)?.file_names().count();
    let target_payload = semantic_docx_text(0, false).into_bytes();
    let content_bytes = (0..shape.docx_paragraphs())
        .try_fold(0usize, |total, index| {
            total.checked_add(semantic_docx_text(index, false).len())
        })
        .ok_or("semantic DOCX text byte count overflows usize")?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("docx-semantic-{}", shape.name()),
            generator: SEMANTIC_DOCX_CORPUS_GENERATOR,
            package_format: "DOCX/OPC/ZIP",
            shape: shape.name(),
            payload_kind: "deterministic-semantic-text",
            compression: "deflate",
            entry_count: shape.docx_paragraphs(),
            archive_member_count,
            entry_bytes: semantic_docx_text(0, false).len(),
            uncompressed_payload_bytes: content_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "paragraph:0".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "paragraph:0".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn docx_source_media_payload(index: usize) -> Vec<u8> {
    let mut bytes = payload_bytes(
        PayloadKind::Incompressible,
        40_000 + index,
        DOCX_SOURCE_MEDIA_ENTRY_BYTES,
    );
    bytes[..8].copy_from_slice(b"\x89PNG\r\n\x1a\n");
    bytes
}

fn docx_source_edit_bytes() -> Result<Vec<u8>, Box<dyn Error>> {
    let mut package = litchi_docx::Package::new()?;
    let document = package.document_mut()?;
    for index in 0..SemanticShape::Medium.docx_paragraphs() {
        let paragraph = document.add_paragraph_with_text(&semantic_docx_text(index, false));
        if index < DOCX_SOURCE_MEDIA_ENTRY_COUNT {
            paragraph.add_picture_from_bytes(
                docx_source_media_payload(index),
                Some(914_400),
                Some(914_400),
            )?;
        }
    }
    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output)?;
    Ok(output.into_inner())
}

fn pptx_source_media_payload(index: usize) -> Vec<u8> {
    let mut bytes = payload_bytes(
        PayloadKind::Incompressible,
        50_000 + index,
        PPTX_SOURCE_MEDIA_ENTRY_BYTES,
    );
    bytes[..8].copy_from_slice(b"\x89PNG\r\n\x1a\n");
    bytes
}

fn pptx_source_edit_bytes() -> Result<Vec<u8>, Box<dyn Error>> {
    let mut authored = litchi_pptx::Package::new()?;
    let presentation = authored.presentation_mut()?;
    for slide_index in 0..PPTX_SOURCE_SLIDE_COUNT {
        let slide = presentation.add_slide()?;
        for shape_index in 0..PPTX_SOURCE_TEXT_BOXES_PER_SLIDE {
            slide.add_text_box(
                &semantic_pptx_text(slide_index, shape_index, false),
                36 + i64::try_from(shape_index % 4)? * 180,
                36 + i64::try_from(shape_index / 4)? * 90,
                144,
                54,
            );
        }
    }
    let mut package = litchi_pptx::Package::from_bytes(&authored.to_bytes()?)?;
    let mut edit = package.opened_presentation_transaction()?;
    for index in 0..PPTX_SOURCE_MEDIA_ENTRY_COUNT {
        let resource = litchi_pptx::media_parts::Resource::new(
            format!("/ppt/media/litchi-perf-source-media-{index:02}.png"),
            "image/png",
            pptx_source_media_payload(index),
        );
        edit.add_picture(
            index,
            format!("litchi-perf-source-media-{index:02}"),
            &resource,
            (800, 800, 72, 72),
        )?;
    }
    let commit = edit.commit()?;
    if !commit.is_changed() {
        return Err("PPTX source-edit corpus media transaction did not change the package".into());
    }
    package.apply_opened_presentation_commit(commit)?;
    Ok(package.to_bytes()?)
}

fn build_docx_source_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let archive = docx_source_edit_bytes()?;
    let package = litchi_docx::Package::from_reader(Cursor::new(archive.clone()))?;
    verify_semantic_docx(&package, SemanticShape::Medium, &[])?;
    let opc = OpcPackage::from_bytes(&archive)?;
    let entry_count = opc.part_count();
    let uncompressed_payload_bytes = opc.iter_parts().try_fold(0usize, |total, part| {
        total
            .checked_add(part.blob().len())
            .ok_or("DOCX source-edit logical byte count overflows usize")
    })?;
    for index in 0..DOCX_SOURCE_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/word/media/image{}.png", index + 1))?;
        if opc.get_part(&uri)?.blob() != docx_source_media_payload(index) {
            return Err("DOCX source-edit media payload differs from specification".into());
        }
    }
    let archive_member_count = ArchiveReader::new(&archive)?.file_names().count();
    let target = SemanticShape::Medium.docx_paragraphs() / 2;
    let target_payload = semantic_docx_text(target, false).into_bytes();
    Ok(Corpus {
        manifest: CorpusManifest {
            name: "docx-source-backed-media".to_owned(),
            generator: DOCX_SOURCE_EDIT_CORPUS_GENERATOR,
            package_format: "DOCX/OPC/ZIP",
            shape: "media-rich",
            payload_kind: "deterministic-incompressible-media",
            compression: "deflate",
            entry_count,
            archive_member_count,
            entry_bytes: DOCX_SOURCE_MEDIA_ENTRY_BYTES,
            uncompressed_payload_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: format!("paragraph:{target}"),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "word/document.xml".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn build_pptx_source_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let archive = pptx_source_edit_bytes()?;
    let package = litchi_pptx::Package::from_bytes(&archive)?;
    verify_pptx_source_edit_semantics(&package, 0)?;
    let opc = OpcPackage::from_bytes(&archive)?;
    let entry_count = opc.part_count();
    let uncompressed_payload_bytes = opc.iter_parts().try_fold(0usize, |total, part| {
        total
            .checked_add(part.blob().len())
            .ok_or("PPTX source-edit logical byte count overflows usize")
    })?;
    for index in 0..PPTX_SOURCE_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!(
            "/ppt/media/litchi-perf-source-media-{index:02}.png"
        ))?;
        if opc.get_part(&uri)?.blob() != pptx_source_media_payload(index) {
            return Err("PPTX source-edit media payload differs from specification".into());
        }
    }
    let target_slide = PPTX_SOURCE_SLIDE_COUNT / 2;
    let target_payload = semantic_pptx_text(target_slide, 0, false).into_bytes();
    Ok(Corpus {
        manifest: CorpusManifest {
            name: "pptx-source-backed-media".to_owned(),
            generator: PPTX_SOURCE_EDIT_CORPUS_GENERATOR,
            package_format: "PPTX/OPC/ZIP",
            shape: "media-rich",
            payload_kind: "deterministic-incompressible-media",
            compression: "deflate",
            entry_count,
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: PPTX_SOURCE_MEDIA_ENTRY_BYTES,
            uncompressed_payload_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: format!("slide:{target_slide}/shape:0"),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: format!("ppt/slides/slide{}.xml", target_slide + 1),
        target_payload,
        xlsx: None,
    })
}

fn xlsx_calculation_media_payload(index: usize) -> Vec<u8> {
    let mut bytes = payload_bytes(
        PayloadKind::Incompressible,
        60_000 + index,
        XLSX_CALC_MEDIA_ENTRY_BYTES,
    );
    bytes[..8].copy_from_slice(b"\x89PNG\r\n\x1a\n");
    bytes
}

fn xlsx_calculation_metadata_edit_bytes() -> Result<Vec<u8>, Box<dyn Error>> {
    let mut package = litchi_xlsx::Package::create()?;
    let mut edit = package.edit_calculation_metadata()?;
    edit.set_properties(
        litchi_xlsx::calculation_properties::Properties::new().with_calculation_id(Some(7)),
    );
    edit.commit()?;
    let mut opc = package.into_plain_opc();
    let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
    opc.get_part_mut(&worksheet_uri)?.set_blob(
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rIdDrawing"/></worksheet>"#.to_vec(),
    );

    let drawing_uri = PackURI::new("/xl/drawings/drawing1.xml")?;
    let mut drawing_xml = String::from(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        use std::fmt::Write as _;
        write!(
            drawing_xml,
            r#"<xdr:twoCellAnchor><xdr:from><xdr:col>{index}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{}" name="Picture {}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImage{index}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"#,
            index + 1,
            index + 1,
            index + 1,
        )?;
    }
    drawing_xml.push_str("</xdr:wsDr>");
    opc.try_add_part(Box::new(BlobPart::new(
        drawing_uri.clone(),
        opc_content_type::OFC_DRAWING.to_owned(),
        drawing_xml.into_bytes(),
    )))?;
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let media_uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        opc.try_add_part(Box::new(BlobPart::new(
            media_uri,
            opc_content_type::PNG.to_owned(),
            xlsx_calculation_media_payload(index),
        )))?;
        opc.get_part_mut(&drawing_uri)?
            .rels_mut()
            .try_add_relationship(
                relationship_type::IMAGE.to_owned(),
                format!("../media/image{}.png", index + 1),
                format!("rIdImage{index}"),
                TargetMode::Internal,
            )?;
    }
    opc.get_part_mut(&worksheet_uri)?
        .rels_mut()
        .try_add_relationship(
            relationship_type::DRAWING.to_owned(),
            "../drawings/drawing1.xml".to_owned(),
            "rIdDrawing".to_owned(),
            TargetMode::Internal,
        )?;
    let package = litchi_xlsx::Package::from_opc(opc)?;
    package.to_bytes().map_err(Into::into)
}

fn build_xlsx_calculation_metadata_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let archive = xlsx_calculation_metadata_edit_bytes()?;
    let package = litchi_xlsx::Package::from_slice(&archive)?;
    if package
        .calculation_metadata()?
        .properties()
        .ok_or("XLSX calculation corpus has no calcPr")?
        .calculation_id()
        != 7
    {
        return Err("XLSX calculation corpus has unexpected calculation ID".into());
    }
    let opc = OpcPackage::from_bytes(&archive)?;
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        if opc.get_part(&uri)?.blob() != xlsx_calculation_media_payload(index) {
            return Err("XLSX calculation corpus media differs from specification".into());
        }
    }
    let target_uri = PackURI::new("/xl/workbook.xml")?;
    let target_payload = opc.get_part(&target_uri)?.blob().to_vec();
    let entry_count = opc.part_count();
    let uncompressed_payload_bytes = opc.iter_parts().try_fold(0usize, |total, part| {
        total
            .checked_add(part.blob().len())
            .ok_or("XLSX calculation corpus logical byte count overflows usize")
    })?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: "xlsx-calculation-metadata-media".to_owned(),
            generator: XLSX_CALC_SOURCE_EDIT_CORPUS_GENERATOR,
            package_format: "XLSX/OPC/ZIP",
            shape: "media-rich",
            payload_kind: "deterministic-incompressible-media",
            compression: "deflate",
            entry_count,
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: XLSX_CALC_MEDIA_ENTRY_BYTES,
            uncompressed_payload_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "workbook:calculation-metadata".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "xl/workbook.xml".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn build_xlsx_defined_names_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let mut corpus = build_xlsx_calculation_metadata_edit_corpus()?;
    corpus.manifest.name = "xlsx-defined-names-media".to_owned();
    corpus.manifest.generator = XLSX_DEFINED_NAMES_SOURCE_EDIT_CORPUS_GENERATOR;
    corpus.manifest.target_entry = "workbook:definedNames".to_owned();
    Ok(corpus)
}

fn build_xlsx_page_break_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let mut corpus = build_xlsx_calculation_metadata_edit_corpus()?;
    let target_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
    let opc = OpcPackage::from_bytes(&corpus.archive)?;
    let target_payload = opc.get_part(&target_uri)?.blob().to_vec();
    corpus.manifest.name = "xlsx-page-break-media".to_owned();
    corpus.manifest.generator = XLSX_PAGE_BREAK_SOURCE_EDIT_CORPUS_GENERATOR;
    corpus.manifest.target_entry = "worksheet:Sheet1:rowBreaks".to_owned();
    corpus.manifest.target_payload_bytes = target_payload.len();
    corpus.manifest.target_payload_sha256 = sha256_hex(&target_payload);
    corpus.target_name = "xl/worksheets/sheet1.xml".to_owned();
    corpus.target_payload = target_payload;
    Ok(corpus)
}

fn build_xlsx_page_margin_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let mut corpus = build_xlsx_calculation_metadata_edit_corpus()?;
    let target_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
    let opc = OpcPackage::from_bytes(&corpus.archive)?;
    let target_payload = opc.get_part(&target_uri)?.blob().to_vec();
    corpus.manifest.name = "xlsx-page-margin-media".to_owned();
    corpus.manifest.generator = XLSX_PAGE_MARGIN_SOURCE_EDIT_CORPUS_GENERATOR;
    corpus.manifest.target_entry = "worksheet:Sheet1:pageMargins".to_owned();
    corpus.manifest.target_payload_bytes = target_payload.len();
    corpus.manifest.target_payload_sha256 = sha256_hex(&target_payload);
    corpus.target_name = "xl/worksheets/sheet1.xml".to_owned();
    corpus.target_payload = target_payload;
    Ok(corpus)
}

fn build_xlsx_page_setup_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let mut corpus = build_xlsx_calculation_metadata_edit_corpus()?;
    let target_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
    let opc = OpcPackage::from_bytes(&corpus.archive)?;
    let target_payload = opc.get_part(&target_uri)?.blob().to_vec();
    corpus.manifest.name = "xlsx-page-setup-media".to_owned();
    corpus.manifest.generator = XLSX_PAGE_SETUP_SOURCE_EDIT_CORPUS_GENERATOR;
    corpus.manifest.target_entry = "worksheet:Sheet1:pageSetup".to_owned();
    corpus.manifest.target_payload_bytes = target_payload.len();
    corpus.manifest.target_payload_sha256 = sha256_hex(&target_payload);
    corpus.target_name = "xl/worksheets/sheet1.xml".to_owned();
    corpus.target_payload = target_payload;
    Ok(corpus)
}

fn build_xlsx_print_options_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let mut corpus = build_xlsx_calculation_metadata_edit_corpus()?;
    let target_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
    let opc = OpcPackage::from_bytes(&corpus.archive)?;
    let target_payload = opc.get_part(&target_uri)?.blob().to_vec();
    corpus.manifest.name = "xlsx-print-options-media".to_owned();
    corpus.manifest.generator = XLSX_PRINT_OPTIONS_SOURCE_EDIT_CORPUS_GENERATOR;
    corpus.manifest.target_entry = "worksheet:Sheet1:printOptions".to_owned();
    corpus.manifest.target_payload_bytes = target_payload.len();
    corpus.manifest.target_payload_sha256 = sha256_hex(&target_payload);
    corpus.target_name = "xl/worksheets/sheet1.xml".to_owned();
    corpus.target_payload = target_payload;
    Ok(corpus)
}

fn build_xlsx_sheet_protection_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let mut corpus = build_xlsx_calculation_metadata_edit_corpus()?;
    let target_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
    let opc = OpcPackage::from_bytes(&corpus.archive)?;
    let target_payload = opc.get_part(&target_uri)?.blob().to_vec();
    corpus.manifest.name = "xlsx-sheet-protection-media".to_owned();
    corpus.manifest.generator = XLSX_SHEET_PROTECTION_SOURCE_EDIT_CORPUS_GENERATOR;
    corpus.manifest.target_entry = "worksheet:Sheet1:protection".to_owned();
    corpus.manifest.target_payload_bytes = target_payload.len();
    corpus.manifest.target_payload_sha256 = sha256_hex(&target_payload);
    corpus.target_name = "xl/worksheets/sheet1.xml".to_owned();
    corpus.target_payload = target_payload;
    Ok(corpus)
}

fn build_xlsx_data_validation_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let mut corpus = build_xlsx_calculation_metadata_edit_corpus()?;
    let target_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
    let mut opc = OpcPackage::from_bytes(&corpus.archive)?;
    let seeded = litchi_xlsx::data_validation::replace_data_validation_collections(
        opc.get_part(&target_uri)?.blob(),
        &xlsx_data_validation_values(false)?,
    )?;
    opc.get_part_mut(&target_uri)?.set_blob(seeded);
    corpus.archive = PackageWriter::to_bytes(&opc)?;
    corpus.manifest.archive_member_count =
        ArchiveReader::new(&corpus.archive)?.file_names().count();
    corpus.manifest.uncompressed_payload_bytes =
        opc.iter_parts().try_fold(0usize, |total, part| {
            total
                .checked_add(part.blob().len())
                .ok_or("XLSX data-validation corpus logical byte count overflows usize")
        })?;
    corpus.manifest.archive_bytes = corpus.archive.len();
    corpus.manifest.archive_sha256 = sha256_hex(&corpus.archive);
    let target_payload = opc.get_part(&target_uri)?.blob().to_vec();
    corpus.manifest.name = "xlsx-data-validation-media".to_owned();
    corpus.manifest.generator = XLSX_DATA_VALIDATION_SOURCE_EDIT_CORPUS_GENERATOR;
    corpus.manifest.target_entry = "worksheet:Sheet1:data-validations".to_owned();
    corpus.manifest.target_payload_bytes = target_payload.len();
    corpus.manifest.target_payload_sha256 = sha256_hex(&target_payload);
    corpus.target_name = "xl/worksheets/sheet1.xml".to_owned();
    corpus.target_payload = target_payload;
    Ok(corpus)
}

fn build_xlsx_auto_filter_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let mut corpus = build_xlsx_calculation_metadata_edit_corpus()?;
    let target_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
    let mut opc = OpcPackage::from_bytes(&corpus.archive)?;
    let seeded = litchi_xlsx::auto_filter::replace_auto_filter(
        opc.get_part(&target_uri)?.blob(),
        Some(&xlsx_auto_filter_value(false)?),
    )?;
    opc.get_part_mut(&target_uri)?.set_blob(seeded);
    corpus.archive = PackageWriter::to_bytes(&opc)?;
    corpus.manifest.archive_member_count =
        ArchiveReader::new(&corpus.archive)?.file_names().count();
    corpus.manifest.uncompressed_payload_bytes =
        opc.iter_parts().try_fold(0usize, |total, part| {
            total
                .checked_add(part.blob().len())
                .ok_or("XLSX auto-filter corpus logical byte count overflows usize")
        })?;
    corpus.manifest.archive_bytes = corpus.archive.len();
    corpus.manifest.archive_sha256 = sha256_hex(&corpus.archive);
    let target_payload = opc.get_part(&target_uri)?.blob().to_vec();
    corpus.manifest.name = "xlsx-auto-filter-media".to_owned();
    corpus.manifest.generator = XLSX_AUTO_FILTER_SOURCE_EDIT_CORPUS_GENERATOR;
    corpus.manifest.target_entry = "worksheet:Sheet1:autoFilter".to_owned();
    corpus.manifest.target_payload_bytes = target_payload.len();
    corpus.manifest.target_payload_sha256 = sha256_hex(&target_payload);
    corpus.target_name = "xl/worksheets/sheet1.xml".to_owned();
    corpus.target_payload = target_payload;
    Ok(corpus)
}

fn build_xlsx_conditional_formatting_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let mut corpus = build_xlsx_calculation_metadata_edit_corpus()?;
    let target_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
    let mut opc = OpcPackage::from_bytes(&corpus.archive)?;
    let seeded = litchi_xlsx::conditional_formatting::replace_conditional_formattings(
        opc.get_part(&target_uri)?.blob(),
        &xlsx_conditional_formatting_values(false)?,
        0,
    )?;
    opc.get_part_mut(&target_uri)?.set_blob(seeded);
    corpus.archive = PackageWriter::to_bytes(&opc)?;
    corpus.manifest.archive_member_count =
        ArchiveReader::new(&corpus.archive)?.file_names().count();
    corpus.manifest.uncompressed_payload_bytes =
        opc.iter_parts().try_fold(0usize, |total, part| {
            total
                .checked_add(part.blob().len())
                .ok_or("XLSX conditional-formatting corpus logical byte count overflows usize")
        })?;
    corpus.manifest.archive_bytes = corpus.archive.len();
    corpus.manifest.archive_sha256 = sha256_hex(&corpus.archive);
    let target_payload = opc.get_part(&target_uri)?.blob().to_vec();
    corpus.manifest.name = "xlsx-conditional-formatting-media".to_owned();
    corpus.manifest.generator = XLSX_CONDITIONAL_FORMATTING_SOURCE_EDIT_CORPUS_GENERATOR;
    corpus.manifest.target_entry = "worksheet:Sheet1:conditional-formatting".to_owned();
    corpus.manifest.target_payload_bytes = target_payload.len();
    corpus.manifest.target_payload_sha256 = sha256_hex(&target_payload);
    corpus.target_name = "xl/worksheets/sheet1.xml".to_owned();
    corpus.target_payload = target_payload;
    Ok(corpus)
}

fn xlsx_merge_fixture() -> Result<(Workbook, Workbook), Box<dyn Error>> {
    let workbook = Workbook::new()?;
    let mut edit = workbook.edit()?;
    let mut sheet = edit
        .sheet("Sheet1")?
        .ok_or("XLSX merge fixture is missing Sheet1")?;
    sheet
        .set("A1", "litchi-xlsx-merge-anchor-v1")?
        .set("C1", "litchi-xlsx-merge-unrelated-v1")?;
    let unmerged = edit.commit()?.into_workbook();

    let mut merge_edit = unmerged.edit()?;
    merge_edit
        .sheet("Sheet1")?
        .ok_or("XLSX merge fixture is missing Sheet1")?
        .merge("A1:B2")?;
    let merged = merge_edit.commit()?.into_workbook();
    Ok((unmerged, merged))
}

fn build_xlsx_merge_edit_corpus(case: Case) -> Result<Corpus, Box<dyn Error>> {
    if !case.is_xlsx_merge_edit_save() {
        return Err("XLSX merge corpus requires a merge or unmerge case".into());
    }
    let (unmerged, merged) = xlsx_merge_fixture()?;
    let source = if case == Case::XlsxEagerMergeCommitSave {
        &unmerged
    } else {
        &merged
    };
    let archive = source.to_bytes()?;
    let opc = OpcPackage::from_bytes(&archive)?;
    let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
    let target_payload = opc.get_part(&worksheet_uri)?.blob().to_vec();
    let entry_count = opc.part_count();
    let uncompressed_payload_bytes = opc.iter_parts().try_fold(0usize, |total, part| {
        total
            .checked_add(part.blob().len())
            .ok_or("XLSX merge fixture logical byte count overflows usize")
    })?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!(
                "xlsx-{}-edit-sparse-a1-b2",
                if case == Case::XlsxEagerMergeCommitSave {
                    "merge"
                } else {
                    "unmerge"
                }
            ),
            generator: XLSX_MERGE_EDIT_CORPUS_GENERATOR,
            package_format: "XLSX/OPC/ZIP",
            shape: "sparse-a1-b2",
            payload_kind: "deterministic-semantic-cells",
            compression: "deflate",
            entry_count,
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: target_payload.len(),
            uncompressed_payload_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: format!(
                "worksheet:Sheet1:{}:A1:B2",
                if case == Case::XlsxEagerMergeCommitSave {
                    "merge"
                } else {
                    "unmerge"
                }
            ),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "xl/worksheets/sheet1.xml".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn semantic_pptx_bytes(shape: SemanticShape) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut package = litchi_pptx::Package::new()?;
    let presentation = package.presentation_mut()?;
    for slide_index in 0..shape.pptx_slides() {
        let slide = presentation.add_slide()?;
        for shape_index in 0..shape.pptx_text_boxes_per_slide() {
            slide.add_text_box(
                &semantic_pptx_text(slide_index, shape_index, false),
                36 + i64::try_from(shape_index % 4)? * 180,
                36 + i64::try_from(shape_index / 4)? * 90,
                144,
                54,
            );
        }
    }
    Ok(package.to_bytes()?)
}

fn semantic_odt_text(index: usize, updated: bool) -> String {
    let state = if updated { "updated" } else { "source" };
    format!("litchi-perf-baseline-odt-semantic-v1-{state}-{index:05}")
}

fn semantic_ods_sheet_name(index: usize) -> String {
    format!("Sheet {index}")
}

fn semantic_ods_text(sheet: usize, row: usize, column: usize, updated: bool) -> String {
    let state = if updated { "updated" } else { "source" };
    format!("litchi-perf-baseline-ods-semantic-v1-{state}-{sheet:02}-{row:03}-{column:03}")
}

fn semantic_odp_title(index: usize, updated: bool) -> String {
    let state = if updated { "updated" } else { "source" };
    format!("litchi-perf-baseline-odp-title-v1-{state}-{index:03}")
}

fn semantic_odp_text(index: usize, updated: bool) -> String {
    let state = if updated { "updated" } else { "source" };
    format!("litchi-perf-baseline-odp-body-v1-{state}-{index:03}")
}

fn semantic_odt_bytes(shape: SemanticShape) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut builder = litchi_odt::Builder::new();
    for index in 0..shape.docx_paragraphs() {
        builder.add_paragraph(&semantic_odt_text(index, false))?;
    }
    Ok(builder.build()?)
}

fn odt_media_path(index: usize) -> String {
    format!("Pictures/litchi-perf-odt-media-{index:02}.bin")
}

fn odt_media_payload(index: usize) -> Vec<u8> {
    payload_bytes(
        PayloadKind::Incompressible,
        30_000 + index,
        ODS_MEDIA_ENTRY_BYTES,
    )
}

fn odt_resource_batch_name(index: usize) -> String {
    format!("litchi-perf-odt-resource-owner-{index:02}")
}

fn odt_resource_batch_path(index: usize, updated: bool) -> String {
    let state = if updated { "target" } else { "source" };
    format!("Pictures/litchi-perf-odt-resource-{state}-{index:02}.png")
}

fn odt_resource_batch_payload(index: usize, updated: bool) -> Vec<u8> {
    let seed = if updated { 50_000 } else { 40_000 };
    payload_bytes(
        PayloadKind::Incompressible,
        seed + index,
        ODT_RESOURCE_PAYLOAD_BYTES,
    )
}

fn odt_resource_batch_image(
    index: usize,
    updated: bool,
) -> litchi_odt::package::embedded::EmbeddedResource {
    litchi_odt::package::embedded::EmbeddedResource {
        kind: litchi_odt::package::embedded::EmbeddedResourceKind::Image,
        source: litchi_odt::package::embedded::EmbeddedResourceSource::PackageFile {
            bytes: odt_resource_batch_payload(index, updated),
            media_type: "image/png".to_owned(),
            preferred_path: Some(odt_resource_batch_path(index, updated)),
        },
        frame_name: Some(odt_resource_batch_name(index)),
        xml_id: None,
        class_id: None,
    }
}

fn odt_media_archive() -> Result<Vec<u8>, Box<dyn Error>> {
    let mut document = litchi_odt::mutable::MutableDocument::new();
    for index in 0..SemanticShape::Medium.docx_paragraphs() {
        document.add_paragraph(&semantic_odt_text(index, false))?;
    }
    let base = document.to_bytes()?;
    let source = ArchiveReader::new(&base)?;
    let mut writer = litchi_odt::core::PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.text")?;
    for path in source.file_names() {
        if matches!(path, "mimetype" | "META-INF/manifest.xml") || path.ends_with('/') {
            continue;
        }
        writer.add_file(path, &source.read(path)?)?;
    }
    writer.add_manifest_directory("Pictures/", "")?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        writer.add_file_with_media_type(
            &odt_media_path(index),
            &odt_media_payload(index),
            "application/octet-stream",
        )?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn odt_resource_batch_archive() -> Result<Vec<u8>, Box<dyn Error>> {
    let base = odt_media_archive()?;
    let source = ArchiveReader::new(&base)?;
    let mut content = String::from_utf8(source.read("content.xml")?)?;
    let insertion = content
        .rfind("</office:text>")
        .ok_or("ODT embedded-resource corpus office:text close tag is missing")?;
    let frames = (0..ODT_RESOURCE_BATCH_COUNT)
        .map(|index| {
            format!(
                r#"<draw:frame xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" draw:name="{}"><draw:image draw:mime-type="image/png" xlink:href="{}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"></draw:image></draw:frame>"#,
                odt_resource_batch_name(index),
                odt_resource_batch_path(index, false),
            )
        })
        .collect::<String>();
    content.insert_str(insertion, &frames);

    let mut writer = litchi_odt::core::PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.text")?;
    for path in source.file_names() {
        if matches!(path, "mimetype" | "META-INF/manifest.xml" | "content.xml")
            || path.ends_with('/')
        {
            continue;
        }
        writer.add_file(path, &source.read(path)?)?;
    }
    writer.add_file("content.xml", content.as_bytes())?;
    for index in 0..ODT_RESOURCE_BATCH_COUNT {
        writer.add_file_with_media_type(
            &odt_resource_batch_path(index, false),
            &odt_resource_batch_payload(index, false),
            "image/png",
        )?;
    }
    let archive = writer.finish_to_bytes()?;
    verify_odt_resource_batch_archive(&archive, false)?;
    Ok(archive)
}

fn semantic_ods_bytes(shape: SemanticShape) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut builder = litchi_ods::Builder::new();
    for sheet_index in 0..shape.ods_sheet_count() {
        let mut sheet = litchi_ods::Sheet::new(semantic_ods_sheet_name(sheet_index))?;
        for row in 0..shape.ods_rows_per_sheet() {
            let mut cells = Vec::with_capacity(shape.ods_columns_per_sheet());
            for column in 0..shape.ods_columns_per_sheet() {
                let text = semantic_ods_text(sheet_index, row, column, false);
                cells.push(litchi_ods::Cell::new(
                    litchi_ods::CellValue::Text(text.clone()),
                    text,
                ));
            }
            sheet.rows.push(litchi_ods::Row {
                cells,
                style_name: None,
                default_cell_style_name: None,
                repeat: std::num::NonZeroUsize::MIN,
            });
        }
        builder.add_sheet(sheet)?;
    }
    Ok(builder.build()?)
}

fn ods_media_path(index: usize) -> String {
    format!("Pictures/litchi-perf-media-{index:02}.bin")
}

fn ods_media_payload(index: usize) -> Vec<u8> {
    payload_bytes(
        PayloadKind::Incompressible,
        10_000 + index,
        ODS_MEDIA_ENTRY_BYTES,
    )
}

fn odp_media_path(index: usize) -> String {
    format!("Pictures/litchi-perf-odp-media-{index:02}.bin")
}

fn odp_media_payload(index: usize) -> Vec<u8> {
    payload_bytes(
        PayloadKind::Incompressible,
        20_000 + index,
        ODS_MEDIA_ENTRY_BYTES,
    )
}

fn odp_media_text() -> String {
    "litchi-perf-baseline-odp-media-text-box-v1".to_owned()
}

fn odp_text_box_batch_name(index: usize) -> String {
    format!("litchi-perf-odp-batch-text-box-{index:02}")
}

fn odp_text_box_batch_text(index: usize, updated: bool) -> String {
    let state = if updated { "updated" } else { "source" };
    format!("litchi-perf-baseline-odp-batch-v1-{state}-{index:02}")
}

fn odp_text_box_batch_page(index: usize) -> usize {
    index * (SemanticShape::Medium.pptx_slides() - 1) / (ODP_TEXT_BOX_BATCH_COUNT - 1)
}

fn ods_media_archive() -> Result<Vec<u8>, Box<dyn Error>> {
    let base = semantic_ods_bytes(SemanticShape::Medium)?;
    let source = ArchiveReader::new(&base)?;
    let mut writer = litchi_odf_common::core::PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.spreadsheet")?;
    for path in source.file_names() {
        if matches!(path, "mimetype" | "META-INF/manifest.xml") || path.ends_with('/') {
            continue;
        }
        writer.add_file(path, &source.read(path)?)?;
    }
    writer.add_manifest_directory("Pictures/", "")?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        writer.add_file_with_media_type(
            &ods_media_path(index),
            &ods_media_payload(index),
            "application/octet-stream",
        )?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn semantic_odp_bytes(shape: SemanticShape) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut builder = litchi_odp::Builder::new();
    for index in 0..shape.pptx_slides() {
        builder.add_slide_with_title(
            &semantic_odp_title(index, false),
            &semantic_odp_text(index, false),
        )?;
    }
    // The public ODP builder currently records wall-clock timestamps in
    // `meta.xml`. Keep its public authored content/style path, but rebuild the
    // generated package with fixed metadata so a benchmark corpus has one
    // stable identity across runs and machines.
    let generated = builder.build()?;
    let reader = ArchiveReader::new(&generated)?;
    let content = reader.read("content.xml")?;
    let styles = reader.read("styles.xml")?;
    let mut writer = litchi_odp::core::PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.presentation")?;
    writer.add_file("content.xml", &content)?;
    writer.add_file("styles.xml", &styles)?;
    writer.add_file(
        "meta.xml",
        br#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>litchi-perf-baseline</meta:generator></office:meta></office:document-meta>"#,
    )?;
    Ok(writer.finish_to_bytes()?)
}

fn odp_media_archive() -> Result<Vec<u8>, Box<dyn Error>> {
    let base = semantic_odp_bytes(SemanticShape::Medium)?;
    let source = ArchiveReader::new(&base)?;
    let mut writer = litchi_odp::core::PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.presentation")?;
    for path in source.file_names() {
        if matches!(path, "mimetype" | "META-INF/manifest.xml") || path.ends_with('/') {
            continue;
        }
        writer.add_file(path, &source.read(path)?)?;
    }
    writer.add_manifest_directory("Pictures/", "")?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        writer.add_file_with_media_type(
            &odp_media_path(index),
            &odp_media_payload(index),
            "application/octet-stream",
        )?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn odp_text_box_batch_archive() -> Result<Vec<u8>, Box<dyn Error>> {
    let source = litchi_odp::authoring::edit::Snapshot::from_bytes(semantic_odp_bytes(
        SemanticShape::Medium,
    )?)?;
    let mut transaction = source.transaction()?;
    for index in 0..ODP_TEXT_BOX_BATCH_COUNT {
        let text_box = litchi_odp::content::TextBox::new(
            odp_text_box_batch_name(index),
            litchi_odp::content::RichText::plain(odp_text_box_batch_text(index, false))?,
        )?;
        transaction.add_text_box(odp_text_box_batch_page(index), &text_box)?;
    }
    let commit = transaction.commit()?;
    if !commit.changed() || commit.patch().is_noop() {
        return Err("ODP text-box batch corpus construction was an exact no-op".into());
    }
    let staged = ArchiveReader::new(commit.snapshot().bytes())?;
    let mut writer = litchi_odp::core::PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.presentation")?;
    for path in staged.file_names() {
        if matches!(path, "mimetype" | "META-INF/manifest.xml") || path.ends_with('/') {
            continue;
        }
        writer.add_file(path, &staged.read(path)?)?;
    }
    writer.add_manifest_directory("Pictures/", "")?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        writer.add_file_with_media_type(
            &odp_media_path(index),
            &odp_media_payload(index),
            "application/octet-stream",
        )?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn build_semantic_odt_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let archive = semantic_odt_bytes(shape)?;
    let document = litchi_odt::Document::from_bytes(archive.clone())?;
    verify_semantic_odt(&document, shape, &[])?;
    let target_payload = semantic_odt_text(0, false).into_bytes();
    let content_bytes = (0..shape.docx_paragraphs()).try_fold(0usize, |total, index| {
        total
            .checked_add(semantic_odt_text(index, false).len())
            .ok_or("semantic ODT text byte count overflows usize")
    })?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("odt-semantic-{}", shape.name()),
            generator: SEMANTIC_ODT_CORPUS_GENERATOR,
            package_format: "ODT/ODF/ZIP",
            shape: shape.name(),
            payload_kind: "deterministic-semantic-text",
            compression: "deflate",
            entry_count: shape.docx_paragraphs(),
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: target_payload.len(),
            uncompressed_payload_bytes: content_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "paragraph:0".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "paragraph:0".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn build_odf_repair_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let canonical = build_semantic_odt_corpus(shape)?;
    let archive = inject_mimetype_local_timestamp_extra(canonical.archive.clone())?;
    let report = litchi_odf_common::validate_package(&archive)?;
    if !report.is_complete()
        || report.issues().len() != 1
        || report.issues()[0].code() != litchi_odf_common::MIMETYPE_LOCAL_EXTRA_ISSUE
        || report.issues()[0].repair().repair_id()
            != Some(litchi_odf_common::MIMETYPE_LOCAL_EXTRA_REPAIR)
    {
        return Err("ODF repair corpus did not produce the sole supported repair issue".into());
    }

    let mut manifest = canonical.manifest;
    manifest.name = format!("odf-mimetype-repair-{}", shape.name());
    manifest.generator = ODF_REPAIR_CORPUS_GENERATOR;
    manifest.archive_bytes = archive.len();
    manifest.archive_sha256 = sha256_hex(&archive);
    manifest.target_entry = "mimetype:local-extra:0x5455".to_owned();
    manifest.target_payload_bytes = ODF_REPAIR_LOCAL_EXTRA.len();
    manifest.target_payload_sha256 = sha256_hex(ODF_REPAIR_LOCAL_EXTRA);
    Ok(Corpus {
        manifest,
        archive,
        target_name: "mimetype:local-extra:0x5455".to_owned(),
        target_payload: ODF_REPAIR_LOCAL_EXTRA.to_vec(),
        xlsx: None,
    })
}

fn inject_mimetype_local_timestamp_extra(mut source: Vec<u8>) -> Result<Vec<u8>, Box<dyn Error>> {
    const LOCAL_FIXED: usize = 30;
    const CENTRAL_LOCAL_OFFSET: std::ops::Range<usize> = 42..46;
    const EOCD_DIRECTORY_OFFSET: std::ops::Range<usize> = 16..20;

    let archive = ZipArchive::from_slice(&source)?;
    let central_start = usize::try_from(archive.directory_offset())?;
    let eocd = usize::try_from(archive.eocd_offset())?;
    let records = archive
        .entries()
        .map(|entry| {
            let entry = entry?;
            Ok((
                usize::try_from(entry.central_directory_offset())?,
                usize::try_from(entry.local_header_offset())?,
            ))
        })
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
    let first_local = records.first().ok_or("ODF repair corpus is empty")?.1;
    if first_local != 0 || source.get(..4) != Some(&0x0403_4b50_u32.to_le_bytes()) {
        return Err("ODF repair corpus does not begin with a local ZIP header".into());
    }
    let name_len = usize::from(u16::from_le_bytes(
        source
            .get(first_local + 26..first_local + 28)
            .ok_or("ODF repair local name length is truncated")?
            .try_into()?,
    ));
    let extra_len = usize::from(u16::from_le_bytes(
        source
            .get(first_local + 28..first_local + 30)
            .ok_or("ODF repair local extra length is truncated")?
            .try_into()?,
    ));
    let name_start = first_local
        .checked_add(LOCAL_FIXED)
        .ok_or("ODF repair local name offset overflows usize")?;
    let insert_at = name_start
        .checked_add(name_len)
        .ok_or("ODF repair local extra offset overflows usize")?;
    if source.get(name_start..insert_at) != Some(b"mimetype") || extra_len != 0 {
        return Err("ODF repair corpus has a non-canonical first mimetype header".into());
    }
    let shift = ODF_REPAIR_LOCAL_EXTRA.len();
    let shift_u32 = u32::try_from(shift)?;
    source.splice(insert_at..insert_at, ODF_REPAIR_LOCAL_EXTRA.iter().copied());
    source[first_local + 28..first_local + 30]
        .copy_from_slice(&u16::try_from(shift)?.to_le_bytes());

    for (central, local) in records {
        let shifted_central = central
            .checked_add(shift)
            .ok_or("ODF repair central offset overflows usize")?;
        let shifted_local = if local > first_local {
            u32::try_from(local)?
                .checked_add(shift_u32)
                .ok_or("ODF repair shifted local-header offset overflows classic ZIP")?
        } else {
            u32::try_from(local)?
        };
        let range = shifted_central + CENTRAL_LOCAL_OFFSET.start
            ..shifted_central + CENTRAL_LOCAL_OFFSET.end;
        source
            .get_mut(range)
            .ok_or("ODF repair central record is truncated")?
            .copy_from_slice(&shifted_local.to_le_bytes());
    }

    let shifted_eocd = eocd
        .checked_add(shift)
        .ok_or("ODF repair EOCD offset overflows usize")?;
    let shifted_directory = u32::try_from(central_start)?
        .checked_add(shift_u32)
        .ok_or("ODF repair central directory offset overflows classic ZIP")?;
    let range =
        shifted_eocd + EOCD_DIRECTORY_OFFSET.start..shifted_eocd + EOCD_DIRECTORY_OFFSET.end;
    source
        .get_mut(range)
        .ok_or("ODF repair EOCD is truncated")?
        .copy_from_slice(&shifted_directory.to_le_bytes());

    ZipArchive::from_slice(&source)?;
    Ok(source)
}

fn build_odt_media_corpus() -> Result<Corpus, Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let archive = odt_media_archive()?;
    verify_odt_media_archive(&archive, false)?;
    let target = shape.docx_paragraphs() / 2;
    let target_payload = semantic_odt_text(target, false).into_bytes();
    let paragraph_bytes = (0..shape.docx_paragraphs()).try_fold(0usize, |total, index| {
        total
            .checked_add(semantic_odt_text(index, false).len())
            .ok_or("media-rich ODT text byte count overflows usize")
    })?;
    let media_bytes = ODS_MEDIA_ENTRY_COUNT
        .checked_mul(ODS_MEDIA_ENTRY_BYTES)
        .ok_or("media-rich ODT payload byte count overflows usize")?;
    let entry_count = shape
        .docx_paragraphs()
        .checked_add(ODS_MEDIA_ENTRY_COUNT)
        .ok_or("media-rich ODT logical entry count overflows usize")?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: "odt-media-paragraph-publication".to_owned(),
            generator: ODT_MEDIA_CORPUS_GENERATOR,
            package_format: "ODT/ODF/ZIP",
            shape: "media-rich",
            payload_kind: "deterministic-incompressible-media",
            compression: "deflate",
            entry_count,
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: ODS_MEDIA_ENTRY_BYTES,
            uncompressed_payload_bytes: paragraph_bytes
                .checked_add(media_bytes)
                .ok_or("media-rich ODT aggregate byte count overflows usize")?,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: format!("paragraph:{target}"),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: format!("paragraph:{target}"),
        target_payload,
        xlsx: None,
    })
}

fn build_odt_resource_batch_corpus() -> Result<Corpus, Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let archive = odt_resource_batch_archive()?;
    verify_odt_resource_batch_archive(&archive, false)?;
    let target_payload = odt_resource_batch_payload(0, false);
    let paragraph_bytes = (0..shape.docx_paragraphs()).try_fold(0usize, |total, index| {
        total
            .checked_add(semantic_odt_text(index, false).len())
            .ok_or("ODT embedded-resource corpus text byte count overflows usize")
    })?;
    let retained_media_bytes = ODS_MEDIA_ENTRY_COUNT
        .checked_mul(ODS_MEDIA_ENTRY_BYTES)
        .ok_or("ODT embedded-resource retained-media byte count overflows usize")?;
    let resource_bytes = ODT_RESOURCE_BATCH_COUNT
        .checked_mul(ODT_RESOURCE_PAYLOAD_BYTES)
        .ok_or("ODT embedded-resource payload byte count overflows usize")?;
    let entry_count = shape
        .docx_paragraphs()
        .checked_add(ODS_MEDIA_ENTRY_COUNT)
        .and_then(|count| count.checked_add(ODT_RESOURCE_BATCH_COUNT))
        .ok_or("ODT embedded-resource logical entry count overflows usize")?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: "odt-embedded-resource-batch-publication".to_owned(),
            generator: ODT_RESOURCE_BATCH_CORPUS_GENERATOR,
            package_format: "ODT/ODF/ZIP",
            shape: "media-rich-64-image-owners",
            payload_kind: "deterministic-incompressible-media-and-image-resources",
            compression: "deflate",
            entry_count,
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: ODT_RESOURCE_PAYLOAD_BYTES,
            uncompressed_payload_bytes: paragraph_bytes
                .checked_add(retained_media_bytes)
                .and_then(|total| total.checked_add(resource_bytes))
                .ok_or("ODT embedded-resource aggregate byte count overflows usize")?,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: format!("image:0/{}", odt_resource_batch_path(0, false)),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: odt_resource_batch_name(0),
        target_payload,
        xlsx: None,
    })
}

fn build_semantic_ods_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let archive = semantic_ods_bytes(shape)?;
    let spreadsheet = litchi_ods::Spreadsheet::from_bytes(archive.clone())?;
    verify_semantic_ods(&spreadsheet, shape, false)?;
    let target_payload = semantic_ods_text(0, 0, 0, false).into_bytes();
    let content_bytes = (0..shape.ods_sheet_count()).try_fold(0usize, |total, sheet| {
        (0..shape.ods_rows_per_sheet()).try_fold(total, |total, row| {
            (0..shape.ods_columns_per_sheet()).try_fold(total, |total, column| {
                total
                    .checked_add(semantic_ods_text(sheet, row, column, false).len())
                    .ok_or("semantic ODS text byte count overflows usize")
            })
        })
    })?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("ods-semantic-{}", shape.name()),
            generator: SEMANTIC_ODS_CORPUS_GENERATOR,
            package_format: "ODS/ODF/ZIP",
            shape: shape.name(),
            payload_kind: "deterministic-semantic-text",
            compression: "deflate",
            entry_count: shape.ods_cell_count(),
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: target_payload.len(),
            uncompressed_payload_bytes: content_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "Sheet 0!R0C0".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "Sheet 0!R0C0".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn build_ods_media_corpus() -> Result<Corpus, Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let archive = ods_media_archive()?;
    verify_ods_media_archive(&archive, false)?;
    let target_payload = semantic_ods_text(0, 0, 0, false).into_bytes();
    let cell_bytes = (0..shape.ods_sheet_count()).try_fold(0usize, |total, sheet| {
        (0..shape.ods_rows_per_sheet()).try_fold(total, |total, row| {
            (0..shape.ods_columns_per_sheet()).try_fold(total, |total, column| {
                total
                    .checked_add(semantic_ods_text(sheet, row, column, false).len())
                    .ok_or("media-rich ODS text byte count overflows usize")
            })
        })
    })?;
    let media_bytes = ODS_MEDIA_ENTRY_COUNT
        .checked_mul(ODS_MEDIA_ENTRY_BYTES)
        .ok_or("media-rich ODS payload byte count overflows usize")?;
    let entry_count = shape
        .ods_cell_count()
        .checked_add(ODS_MEDIA_ENTRY_COUNT)
        .ok_or("media-rich ODS logical entry count overflows usize")?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: "ods-media-publication".to_owned(),
            generator: ODS_MEDIA_CORPUS_GENERATOR,
            package_format: "ODS/ODF/ZIP",
            shape: "media-rich",
            payload_kind: "deterministic-incompressible-media",
            compression: "deflate",
            entry_count,
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: ODS_MEDIA_ENTRY_BYTES,
            uncompressed_payload_bytes: cell_bytes
                .checked_add(media_bytes)
                .ok_or("media-rich ODS aggregate byte count overflows usize")?,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "Sheet 1!R16C16".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "Sheet 1!R16C16".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn build_semantic_odp_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let archive = semantic_odp_bytes(shape)?;
    let presentation = litchi_odp::Presentation::from_bytes(archive.clone())?;
    verify_semantic_odp(&presentation, shape, false)?;
    let target_payload = semantic_odp_text(0, false).into_bytes();
    let content_bytes = (0..shape.pptx_slides()).try_fold(0usize, |total, index| {
        total
            .checked_add(semantic_odp_title(index, false).len())
            .and_then(|total| total.checked_add(semantic_odp_text(index, false).len()))
            .ok_or("semantic ODP text byte count overflows usize")
    })?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("odp-semantic-{}", shape.name()),
            generator: SEMANTIC_ODP_CORPUS_GENERATOR,
            package_format: "ODP/ODF/ZIP",
            shape: shape.name(),
            payload_kind: "deterministic-semantic-text",
            compression: "deflate",
            entry_count: shape.pptx_slides(),
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: target_payload.len(),
            uncompressed_payload_bytes: content_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "slide:0".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "slide:0".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn build_odp_media_corpus() -> Result<Corpus, Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let archive = odp_media_archive()?;
    verify_odp_media_archive(&archive, false)?;
    let target_payload = odp_media_text().into_bytes();
    let content_bytes = (0..shape.pptx_slides()).try_fold(0usize, |total, index| {
        total
            .checked_add(semantic_odp_title(index, false).len())
            .and_then(|total| total.checked_add(semantic_odp_text(index, false).len()))
            .ok_or("media-rich ODP text byte count overflows usize")
    })?;
    let media_bytes = ODS_MEDIA_ENTRY_COUNT
        .checked_mul(ODS_MEDIA_ENTRY_BYTES)
        .ok_or("media-rich ODP payload byte count overflows usize")?;
    let entry_count = shape
        .pptx_slides()
        .checked_add(ODS_MEDIA_ENTRY_COUNT)
        .ok_or("media-rich ODP logical entry count overflows usize")?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: "odp-media-textbox-publication".to_owned(),
            generator: ODP_MEDIA_CORPUS_GENERATOR,
            package_format: "ODP/ODF/ZIP",
            shape: "media-rich",
            payload_kind: "deterministic-incompressible-media",
            compression: "deflate",
            entry_count,
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: ODS_MEDIA_ENTRY_BYTES,
            uncompressed_payload_bytes: content_bytes
                .checked_add(media_bytes)
                .ok_or("media-rich ODP aggregate byte count overflows usize")?,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: format!("slide:0/{ODP_MEDIA_TEXT_BOX_NAME}"),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: ODP_MEDIA_TEXT_BOX_NAME.to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn build_odp_text_box_batch_corpus() -> Result<Corpus, Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let archive = odp_text_box_batch_archive()?;
    verify_odp_text_box_batch_archive(&archive, false)?;
    let target_payload = odp_text_box_batch_text(0, false).into_bytes();
    let slide_bytes = (0..shape.pptx_slides()).try_fold(0usize, |total, index| {
        total
            .checked_add(semantic_odp_title(index, false).len())
            .and_then(|total| total.checked_add(semantic_odp_text(index, false).len()))
            .ok_or("ODP text-box batch slide byte count overflows usize")
    })?;
    let text_box_bytes = (0..ODP_TEXT_BOX_BATCH_COUNT).try_fold(0usize, |total, index| {
        total
            .checked_add(odp_text_box_batch_text(index, false).len())
            .ok_or("ODP text-box batch text byte count overflows usize")
    })?;
    let media_bytes = ODS_MEDIA_ENTRY_COUNT
        .checked_mul(ODS_MEDIA_ENTRY_BYTES)
        .ok_or("ODP text-box batch media byte count overflows usize")?;
    let entry_count = shape
        .pptx_slides()
        .checked_add(ODP_TEXT_BOX_BATCH_COUNT)
        .and_then(|count| count.checked_add(ODS_MEDIA_ENTRY_COUNT))
        .ok_or("ODP text-box batch logical entry count overflows usize")?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: "odp-cross-slide-textbox-publication".to_owned(),
            generator: ODP_TEXT_BOX_BATCH_CORPUS_GENERATOR,
            package_format: "ODP/ODF/ZIP",
            shape: "media-rich-cross-slide",
            payload_kind: "deterministic-incompressible-media-and-rich-text",
            compression: "deflate",
            entry_count,
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: ODS_MEDIA_ENTRY_BYTES,
            uncompressed_payload_bytes: slide_bytes
                .checked_add(text_box_bytes)
                .and_then(|total| total.checked_add(media_bytes))
                .ok_or("ODP text-box batch aggregate byte count overflows usize")?,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: format!(
                "slide:{}/{}",
                odp_text_box_batch_page(0),
                odp_text_box_batch_name(0)
            ),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: odp_text_box_batch_name(0),
        target_payload,
        xlsx: None,
    })
}

fn build_semantic_pptx_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let archive = semantic_pptx_bytes(shape)?;
    let package = litchi_pptx::Package::from_bytes(&archive)?;
    verify_semantic_pptx(&package, shape, &[])?;
    let count = shape
        .pptx_slides()
        .checked_mul(shape.pptx_text_boxes_per_slide())
        .ok_or("semantic PPTX shape count overflows usize")?;
    let content_bytes = (0..shape.pptx_slides())
        .try_fold(0usize, |total, slide| {
            (0..shape.pptx_text_boxes_per_slide()).try_fold(total, |total, object| {
                total.checked_add(semantic_pptx_text(slide, object, false).len())
            })
        })
        .ok_or("semantic PPTX text byte count overflows usize")?;
    let target_payload = semantic_pptx_text(0, 0, false).into_bytes();
    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("pptx-semantic-{}", shape.name()),
            generator: SEMANTIC_PPTX_CORPUS_GENERATOR,
            package_format: "PPTX/OPC/ZIP",
            shape: shape.name(),
            payload_kind: "deterministic-semantic-text",
            compression: "deflate",
            entry_count: count,
            archive_member_count: ArchiveReader::new(&archive)?.file_names().count(),
            entry_bytes: semantic_pptx_text(0, 0, false).len(),
            uncompressed_payload_bytes: content_bytes,
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

fn build_xlsx_corpus(shape: XlsxShape) -> Result<Corpus, Box<dyn Error>> {
    let spec = XlsxCorpus {
        sheet_count: shape.sheet_count(),
        row_count: shape.row_count(),
        column_count: shape.column_count(),
        one_percent_updates: xlsx_one_percent_updates(shape)?,
        cell_inventory: None,
    };
    let workbook = build_xlsx_workbook(&spec)?;
    let archive = workbook.to_bytes()?;
    let reopened = Workbook::from_bytes(archive.clone())?;
    verify_xlsx_cells(&reopened, &spec, &[])?;

    let cell_count = xlsx_cell_count(&spec)?;
    let target = XlsxCoordinate {
        sheet: 0,
        row: 0,
        column: 0,
    };
    let target_name = xlsx_cell_name(target);
    let target_payload = xlsx_value(target).to_string().into_bytes();
    let archive_member_count = ArchiveReader::new(&archive)?.file_names().count();
    let (_source_ranges, source_members) = xlsx_source_layout(&archive, spec.sheet_count)?;

    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("xlsx-{}", shape.name()),
            generator: XLSX_CORPUS_GENERATOR,
            package_format: "XLSX/OPC/ZIP",
            shape: shape.name(),
            payload_kind: "deterministic-integer-grid",
            compression: "deflate",
            entry_count: cell_count,
            archive_member_count,
            entry_bytes: std::mem::size_of::<i32>(),
            uncompressed_payload_bytes: cell_count
                .checked_mul(std::mem::size_of::<i32>())
                .ok_or("XLSX logical payload size overflows usize")?,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: target_name.clone(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: Some(XlsxManifest {
                sheet_count: spec.sheet_count,
                rows_per_sheet: spec.row_count,
                columns_per_sheet: spec.column_count,
                one_percent_update_count: spec.one_percent_updates.len(),
                source_members,
            }),
        },
        archive,
        target_name,
        target_payload,
        xlsx: Some(spec),
    })
}

fn xlsx_cell_crud_inventory(shape: XlsxCellCrudShape) -> Vec<Vec<XlsxCoordinate>> {
    let sheet_count = 4;
    let mut inventory = Vec::with_capacity(sheet_count);
    for sheet in 0..sheet_count {
        let mut cells = Vec::new();
        match shape {
            XlsxCellCrudShape::Medium => {
                for row in 0..48 {
                    for column in 0..48 {
                        cells.push(XlsxCoordinate { sheet, row, column });
                    }
                }
            },
            XlsxCellCrudShape::DenseSparse => match sheet {
                0 => {
                    for row in 0..128 {
                        for column in 0..128 {
                            cells.push(XlsxCoordinate { sheet, row, column });
                        }
                    }
                },
                1 => {
                    for row in (0..128).step_by(4) {
                        for column in (0..128).step_by(4) {
                            cells.push(XlsxCoordinate { sheet, row, column });
                        }
                    }
                },
                2 => {
                    for row in (0..128).step_by(8) {
                        for column in (0..128).step_by(8) {
                            cells.push(XlsxCoordinate { sheet, row, column });
                        }
                    }
                },
                _ => {
                    for index in 0..128 {
                        cells.push(XlsxCoordinate {
                            sheet,
                            row: index,
                            column: index,
                        });
                    }
                },
            },
        }
        inventory.push(cells);
    }
    inventory
}

fn xlsx_cell_crud_updates(
    inventory: &[Vec<XlsxCoordinate>],
) -> Result<Vec<XlsxCoordinate>, Box<dyn Error>> {
    let all = inventory.iter().flatten().copied().collect::<Vec<_>>();
    let update_count = all
        .len()
        .checked_add(99)
        .ok_or("XLSX cell CRUD update count overflows usize")?
        / 100;
    let mut updates = Vec::with_capacity(update_count);
    for index in 0..update_count {
        let position = index
            .checked_mul(all.len())
            .ok_or("XLSX cell CRUD update position overflows usize")?
            / update_count;
        updates.push(
            *all.get(position)
                .ok_or("XLSX cell CRUD update position is outside inventory")?,
        );
    }
    Ok(updates)
}

fn xlsx_cell_crud_media_payload(index: usize) -> Vec<u8> {
    let mut payload = payload_bytes(
        PayloadKind::Incompressible,
        70_000 + index,
        XLSX_CELL_VALUES_MEDIA_ENTRY_BYTES,
    );
    payload[..8].copy_from_slice(b"\x89PNG\r\n\x1a\n");
    payload
}

fn strip_xlsx_cell_crud_calc_properties(package: &mut OpcPackage) -> Result<(), Box<dyn Error>> {
    let workbook_uri = PackURI::new("/xl/workbook.xml")?;
    let workbook_xml = package.get_part(&workbook_uri)?.blob();
    let mut workbook_xml = String::from_utf8(workbook_xml.to_vec())?;
    if let Some(start) = workbook_xml.find("<calcPr") {
        let end = workbook_xml[start..]
            .find("/>")
            .map(|offset| start + offset + 2)
            .ok_or("XLSX CRUD workbook calcPr is not self-closing")?;
        workbook_xml.replace_range(start..end, "");
    }
    package
        .get_part_mut(&workbook_uri)?
        .set_blob(workbook_xml.into_bytes());
    Ok(())
}

fn build_xlsx_cell_crud_corpus(shape: XlsxCellCrudShape) -> Result<Corpus, Box<dyn Error>> {
    let inventory = xlsx_cell_crud_inventory(shape);
    let updates = xlsx_cell_crud_updates(&inventory)?;
    let (row_count, column_count) = match shape {
        XlsxCellCrudShape::Medium => (48, 48),
        XlsxCellCrudShape::DenseSparse => (128, 128),
    };
    let spec = XlsxCorpus {
        sheet_count: inventory.len(),
        row_count,
        column_count,
        one_percent_updates: updates,
        cell_inventory: Some(inventory),
    };
    let workbook = build_xlsx_workbook(&spec)?;
    let mut archive = workbook.to_bytes()?;
    let mut package = OpcPackage::from_bytes(&archive)?;
    strip_xlsx_cell_crud_calc_properties(&mut package)?;
    for index in 0..XLSX_CELL_VALUES_MEDIA_ENTRY_COUNT {
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/xl/media/litchi-cell-crud-{index:02}.png"))?,
            opc_content_type::PNG.to_owned(),
            xlsx_cell_crud_media_payload(index),
        )))?;
    }
    archive = PackageWriter::to_bytes(&package)?;
    let reopened = Workbook::from_bytes(archive.clone())?;
    verify_xlsx_cells(&reopened, &spec, &[])?;
    let cell_count = xlsx_cell_count(&spec)?;
    let target = *spec
        .one_percent_updates
        .first()
        .ok_or("XLSX cell CRUD corpus has no target cell")?;
    let target_name = xlsx_cell_name(target);
    let target_payload = xlsx_value(target).to_string().into_bytes();
    let archive_member_count = ArchiveReader::new(&archive)?.file_names().count();
    let (_source_ranges, source_members) = xlsx_source_layout(&archive, spec.sheet_count)?;
    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("xlsx-cell-values-{}", shape.name()),
            generator: XLSX_CELL_VALUES_SOURCE_EDIT_CORPUS_GENERATOR,
            package_format: "XLSX/OPC/ZIP",
            shape: shape.name(),
            payload_kind: "deterministic-multi-sheet-scalar-grid-with-media",
            compression: "deflate",
            entry_count: cell_count,
            archive_member_count,
            entry_bytes: std::mem::size_of::<i32>(),
            uncompressed_payload_bytes: cell_count
                .checked_mul(std::mem::size_of::<i32>())
                .and_then(|bytes| {
                    bytes.checked_add(
                        XLSX_CELL_VALUES_MEDIA_ENTRY_COUNT
                            .checked_mul(XLSX_CELL_VALUES_MEDIA_ENTRY_BYTES)?,
                    )
                })
                .ok_or("XLSX cell CRUD logical byte count overflows usize")?,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: target_name.clone(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: Some(XlsxManifest {
                sheet_count: spec.sheet_count,
                rows_per_sheet: spec.row_count,
                columns_per_sheet: spec.column_count,
                one_percent_update_count: spec.one_percent_updates.len(),
                source_members,
            }),
        },
        archive,
        target_name,
        target_payload,
        xlsx: Some(spec),
    })
}

fn build_xlsx_workbook(spec: &XlsxCorpus) -> Result<Workbook, Box<dyn Error>> {
    let workbook = Workbook::new()?;
    let mut edit = workbook.edit()?;
    if let Some(inventory) = spec.cell_inventory.as_ref() {
        for sheet_index in 0..spec.sheet_count {
            let name = xlsx_sheet_name(sheet_index);
            let coordinates = inventory
                .get(sheet_index)
                .ok_or("XLSX CRUD sheet inventory is missing")?;
            if sheet_index == 0 {
                let mut sheet = edit
                    .sheet(name.as_str())?
                    .ok_or("XLSX CRUD sheet is missing")?;
                for coordinate in coordinates {
                    sheet.set(
                        xlsx_address(coordinate.row, coordinate.column)?,
                        xlsx_value(*coordinate),
                    )?;
                }
            } else {
                let mut sheet = edit.add(name)?;
                for coordinate in coordinates {
                    sheet.set(
                        xlsx_address(coordinate.row, coordinate.column)?,
                        xlsx_value(*coordinate),
                    )?;
                }
            }
        }
    } else {
        {
            let mut sheet = edit
                .sheet("Sheet1")?
                .ok_or("XLSX baseline sheet is missing")?;
            for row in 0..spec.row_count {
                for column in 0..spec.column_count {
                    let coordinate = XlsxCoordinate {
                        sheet: 0,
                        row,
                        column,
                    };
                    sheet.set(xlsx_address(row, column)?, xlsx_value(coordinate))?;
                }
            }
        }
        for sheet_index in 1..spec.sheet_count {
            let name = xlsx_sheet_name(sheet_index);
            let mut sheet = edit.add(name)?;
            for row in 0..spec.row_count {
                for column in 0..spec.column_count {
                    let coordinate = XlsxCoordinate {
                        sheet: sheet_index,
                        row,
                        column,
                    };
                    sheet.set(xlsx_address(row, column)?, xlsx_value(coordinate))?;
                }
            }
        }
    }
    Ok(edit.commit()?.workbook().clone())
}

fn xlsx_one_percent_updates(shape: XlsxShape) -> Result<Vec<XlsxCoordinate>, Box<dyn Error>> {
    let sheet_count = shape.sheet_count();
    let row_count = shape.row_count();
    let column_count = shape.column_count();
    let total = sheet_count
        .checked_mul(row_count)
        .and_then(|value| value.checked_mul(column_count))
        .ok_or("XLSX cell count overflows usize")?;
    let update_count = total
        .checked_add(99)
        .ok_or("XLSX update count overflows usize")?
        / 100;
    let mut updates = Vec::with_capacity(update_count);
    for index in 0..update_count {
        let linear = index
            .checked_mul(total)
            .ok_or("XLSX update position overflows usize")?
            / update_count;
        let sheet = linear / (row_count * column_count);
        let within_sheet = linear % (row_count * column_count);
        updates.push(XlsxCoordinate {
            sheet,
            row: within_sheet / column_count,
            column: within_sheet % column_count,
        });
    }
    Ok(updates)
}

fn xlsx_cell_count(spec: &XlsxCorpus) -> Result<usize, Box<dyn Error>> {
    if let Some(inventory) = spec.cell_inventory.as_ref() {
        return inventory.iter().try_fold(0usize, |total, sheet| {
            total
                .checked_add(sheet.len())
                .ok_or_else(|| "XLSX cell inventory count overflows usize".into())
        });
    }
    spec.sheet_count
        .checked_mul(spec.row_count)
        .and_then(|value| value.checked_mul(spec.column_count))
        .ok_or_else(|| "XLSX cell count overflows usize".into())
}

fn xlsx_sheet_name(index: usize) -> String {
    if index == 0 {
        "Sheet1".to_owned()
    } else {
        format!("Bench{index:02}")
    }
}

fn xlsx_address(row: usize, column: usize) -> Result<String, Box<dyn Error>> {
    let row = row
        .checked_add(1)
        .ok_or("XLSX row number overflows usize")?;
    let mut value = column
        .checked_add(1)
        .ok_or("XLSX column number overflows usize")?;
    let mut label = String::new();
    while value != 0 {
        let remainder = (value - 1) % 26;
        label.insert(0, char::from(b'A' + u8::try_from(remainder)?));
        value = (value - 1) / 26;
    }
    Ok(format!("{label}{row}"))
}

fn xlsx_cell_name(coordinate: XlsxCoordinate) -> String {
    format!(
        "{}!{}",
        xlsx_sheet_name(coordinate.sheet),
        xlsx_address(coordinate.row, coordinate.column)
            .expect("bounded XLSX benchmark coordinate must be valid")
    )
}

fn xlsx_cell_crud_updates_for_case(
    case: Case,
    spec: &XlsxCorpus,
) -> Result<Vec<XlsxCoordinate>, Box<dyn Error>> {
    let inventory = spec
        .cell_inventory
        .as_ref()
        .ok_or("XLSX cell CRUD case has no cell inventory")?;
    let total = inventory.iter().map(Vec::len).sum::<usize>();
    let count = match case {
        Case::XlsxEagerCellValuesOneEditSave | Case::XlsxSourceBackedCellValuesOneEditSave => 1,
        Case::XlsxEagerCellValuesOnePercentEditSave
        | Case::XlsxSourceBackedCellValuesOnePercentEditSave => total.div_ceil(100),
        Case::XlsxEagerCellValuesBatchEditSave | Case::XlsxSourceBackedCellValuesBatchEditSave => {
            litchi_xlsx::cell_values::MAX_BATCH_EDITS
        },
        _ => return Err("invalid XLSX cell CRUD case".into()),
    };
    if count == 0 || count > total {
        return Err("XLSX cell CRUD update count is outside corpus inventory".into());
    }
    if matches!(
        case,
        Case::XlsxEagerCellValuesOnePercentEditSave
            | Case::XlsxSourceBackedCellValuesOnePercentEditSave
    ) {
        if spec.one_percent_updates.len() != count {
            return Err("XLSX CRUD 1% manifest count differs from selected updates".into());
        }
        return Ok(spec.one_percent_updates.clone());
    }
    if count == 1 {
        return Ok(vec![
            *inventory
                .first()
                .and_then(|cells| cells.first())
                .ok_or("XLSX cell CRUD corpus has no first cell")?,
        ]);
    }
    let mut updates = Vec::with_capacity(count);
    for index in 0..count {
        let sheet = index % inventory.len();
        let local = index / inventory.len();
        updates.push(
            *inventory
                .get(sheet)
                .and_then(|cells| cells.get(local))
                .ok_or("XLSX cross-sheet update position is outside inventory")?,
        );
    }
    Ok(updates)
}

fn xlsx_update_sheet_selectors(updates: &[XlsxCoordinate]) -> Vec<litchi_xlsx::Selector<'static>> {
    let mut positions = updates
        .iter()
        .map(|coordinate| coordinate.sheet)
        .collect::<Vec<_>>();
    positions.sort_unstable();
    positions.dedup();
    positions
        .into_iter()
        .map(litchi_xlsx::Selector::from)
        .collect()
}

fn verify_xlsx_cell_crud_output(
    corpus: &Corpus,
    output: &[u8],
    updated: &[XlsxCoordinate],
) -> Result<(), Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let workbook = Workbook::from_bytes(output.to_vec())?;
    verify_xlsx_cells(&workbook, spec, updated)?;
    verify_xlsx_cell_crud_package_identity(corpus, output)
}

fn verify_xlsx_cell_crud_package_identity(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != candidate.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("XLSX cell CRUD output changed package topology".into());
    }
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("XLSX cell CRUD output changed Part metadata".into());
        }
        if source_part.partname().membername().starts_with("xl/media/")
            && candidate_part.blob() != source_part.blob()
        {
            return Err("XLSX cell CRUD output changed untouched media".into());
        }
    }
    Ok(())
}

fn verify_xlsx_cell_crud_raw_source_output(
    corpus: &Corpus,
    output: &[u8],
    updated: &[XlsxCoordinate],
) -> Result<(), Box<dyn Error>> {
    let source = raw_zip_members(&corpus.archive)?;
    let candidate = raw_zip_members(output)?;
    if source.keys().ne(candidate.keys()) {
        return Err("XLSX source CRUD raw ZIP member set differs from source".into());
    }
    let touched = updated
        .iter()
        .map(|coordinate| format!("xl/worksheets/sheet{}.xml", coordinate.sheet + 1))
        .collect::<BTreeSet<_>>();
    for (name, source_member) in source {
        if !touched.contains(&name) && candidate.get(&name) != Some(&source_member) {
            return Err(
                format!("XLSX source CRUD changed raw unselected ZIP member {name}").into(),
            );
        }
    }
    Ok(())
}

fn run_xlsx_cell_value_lifecycle_gates(
    corpus: &Corpus,
    spec: &XlsxCorpus,
) -> Result<(), Box<dyn Error>> {
    let selectors = (0..spec.sheet_count)
        .map(litchi_xlsx::Selector::from)
        .collect::<Vec<_>>();
    let source = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let editor = litchi_xlsx::cell_values::SourceBackedEditor::from_read_at(source)?;
    let noop = editor.edit_sheets(selectors.clone())?;
    let noop_commit = noop.commit()?;
    if noop_commit.changed() || !noop_commit.patch().is_empty() {
        return Err("XLSX cell CRUD exact no-op produced a change".into());
    }
    let mut noop_output = Vec::new();
    editor.publish_multi_commit_to_stream(&mut noop_output, &noop_commit)?;
    if noop_output != corpus.archive {
        return Err("XLSX cell CRUD exact no-op changed source bytes".into());
    }

    let target = *spec
        .cell_inventory
        .as_ref()
        .and_then(|inventory| inventory.first())
        .and_then(|cells| cells.first())
        .ok_or("XLSX cell CRUD corpus has no lifecycle target")?;
    for remove in [false, true] {
        let source = Arc::new(OwnedSource::new(corpus.archive.clone()));
        let editor = litchi_xlsx::cell_values::SourceBackedEditor::from_read_at(source)?;
        let mut edit = editor.edit_sheets([litchi_xlsx::Selector::from(target.sheet)])?;
        let address =
            litchi_xlsx::Address::at(u32::try_from(target.row)?, u32::try_from(target.column)?)?;
        if remove {
            edit.remove(target.sheet, address)?;
        } else {
            edit.clear(target.sheet, address)?;
        }
        let commit = edit.commit()?;
        if !commit.changed() || commit.diagnostics().changed_cells() != 1 {
            return Err(
                "XLSX cell CRUD clear/remove lifecycle gate did not change one cell".into(),
            );
        }
        let mut published = Vec::new();
        editor.publish_multi_commit_to_stream(&mut published, &commit)?;
        verify_xlsx_cell_crud_package_identity(corpus, &published)?;
        let workbook = Workbook::from_bytes(published.clone())?;
        verify_xlsx_cell_crud_lifecycle_state(&workbook, spec, target, remove)?;
        let mut applied = OpcPackage::from_bytes(&corpus.archive)?;
        commit.patch().apply(&mut applied)?;
        let applied_bytes = PackageWriter::to_bytes(&applied)?;
        verify_xlsx_cell_crud_package_identity(corpus, &applied_bytes)?;
        let applied_workbook = Workbook::from_bytes(applied_bytes)?;
        verify_xlsx_cell_crud_lifecycle_state(&applied_workbook, spec, target, remove)?;
        if commit.patch().apply(&mut applied).is_ok() {
            return Err("XLSX lifecycle stale source was accepted".into());
        }
        commit.patch().inverse().apply(&mut applied)?;
        let restored_bytes = PackageWriter::to_bytes(&applied)?;
        verify_xlsx_cell_crud_package_identity(corpus, &restored_bytes)?;
        let restored_workbook = Workbook::from_bytes(restored_bytes)?;
        verify_xlsx_cells(&restored_workbook, spec, &[])?;
    }
    Ok(())
}

fn verify_xlsx_cell_crud_lifecycle_state(
    workbook: &Workbook,
    spec: &XlsxCorpus,
    target: XlsxCoordinate,
    removed: bool,
) -> Result<(), Box<dyn Error>> {
    let sheet = workbook
        .sheet(xlsx_sheet_name(target.sheet).as_str())?
        .ok_or("XLSX lifecycle state is missing its target sheet")?;
    let address = xlsx_address(target.row, target.column)?;
    let stored = sheet.cell(address.as_str())?.stored();
    if removed {
        if stored.is_some() {
            return Err("XLSX remove lifecycle retained the cell owner".into());
        }
    } else if !matches!(stored, Some(XlsxCell::Empty)) {
        return Err("XLSX clear lifecycle did not retain an empty cell owner".into());
    }
    let expected = spec
        .cell_inventory
        .as_ref()
        .and_then(|inventory| inventory.get(target.sheet))
        .map(|cells| cells.len() - usize::from(removed))
        .ok_or("XLSX lifecycle target sheet inventory is missing")?;
    if sheet.cells("A1:XFD1048576")?.count() != expected {
        return Err("XLSX lifecycle state changed the wrong cell owners".into());
    }
    Ok(())
}

fn xlsx_value(coordinate: XlsxCoordinate) -> i32 {
    let sheet = i32::try_from(coordinate.sheet).expect("bounded XLSX sheet count fits i32");
    let row = i32::try_from(coordinate.row).expect("bounded XLSX row count fits i32");
    let column = i32::try_from(coordinate.column).expect("bounded XLSX column count fits i32");
    sheet * 1_000_000 + row * 1_000 + column
}

fn prepare_xlsx_updates(
    workbook: &Workbook,
    updates: &[XlsxCoordinate],
) -> Result<litchi_xlsx::Edit, Box<dyn Error>> {
    let mut edit = workbook.edit()?;
    for coordinate in updates {
        let name = xlsx_sheet_name(coordinate.sheet);
        let mut sheet = edit
            .sheet(name.as_str())?
            .ok_or("XLSX update target sheet is missing")?;
        sheet.set(
            xlsx_address(coordinate.row, coordinate.column)?,
            xlsx_value(*coordinate) + 1,
        )?;
    }
    Ok(edit)
}

fn verify_xlsx_cells(
    workbook: &Workbook,
    spec: &XlsxCorpus,
    updated: &[XlsxCoordinate],
) -> Result<(), Box<dyn Error>> {
    if workbook.len() != spec.sheet_count {
        return Err("XLSX workbook sheet count differs from corpus specification".into());
    }
    for sheet_index in 0..spec.sheet_count {
        let name = xlsx_sheet_name(sheet_index);
        let sheet = workbook
            .sheet(name.as_str())?
            .ok_or("XLSX workbook sheet is missing")?;
        let expected = spec.cell_inventory.as_ref().map_or_else(
            || {
                spec.row_count
                    .checked_mul(spec.column_count)
                    .ok_or("XLSX per-sheet cell count overflows usize")
            },
            |inventory| {
                inventory
                    .get(sheet_index)
                    .map(|cells| cells.len())
                    .ok_or("XLSX cell inventory is missing a sheet")
            },
        )?;
        let observed = sheet.cells("A1:XFD1048576")?.count();
        if observed != expected {
            return Err("XLSX stored cell count differs from corpus specification".into());
        }
        let coordinates = spec.cell_inventory.as_ref().map_or_else(
            || {
                (0..spec.row_count)
                    .flat_map(|row| {
                        (0..spec.column_count).map(move |column| XlsxCoordinate {
                            sheet: sheet_index,
                            row,
                            column,
                        })
                    })
                    .collect::<Vec<_>>()
            },
            |inventory| inventory[sheet_index].clone(),
        );
        for coordinate in coordinates {
            let expected = xlsx_value(coordinate) + i32::from(updated.contains(&coordinate));
            let address = xlsx_address(coordinate.row, coordinate.column)?;
            let actual = sheet
                .cell(address.as_str())?
                .stored()
                .ok_or("XLSX expected stored cell is missing")?;
            let XlsxCell::Value(XlsxValue::Number(actual)) = actual else {
                return Err("XLSX expected numeric cell has another value type".into());
            };
            if actual.as_str() != expected.to_string() {
                return Err("XLSX numeric cell differs from deterministic expectation".into());
            }
        }
    }
    Ok(())
}

fn writer_text(kind: &str, first: usize, second: usize, third: usize) -> String {
    format!(
        "litchi-perf-baseline-{kind}-v1-{first:03}-{second:05}-{third:03} deterministic payload"
    )
}

fn updated_writer_text(kind: &str, first: usize, second: usize, third: usize) -> String {
    format!(
        "litchi-perf-baseline-{kind}-v2-{first:03}-{second:05}-{third:03} deterministic payload"
    )
}

fn writer_payload_text(
    kind: &str,
    first: usize,
    second: usize,
    third: usize,
    length: usize,
) -> String {
    const REPEATED_TEXT: &str = "litchi-perf-baseline-payload-heavy-v1 ";
    let mut text = writer_text(kind, first, second, third);
    while text.len() < length {
        text.push_str(REPEATED_TEXT);
    }
    text.truncate(length);
    text
}

fn entry_name(index: usize) -> String {
    format!("benchmark/parts/{index:05}.bin")
}

fn cfb_entry_name(index: usize) -> String {
    format!("benchmark_stream_{index:05}.bin")
}

fn payload_bytes(kind: PayloadKind, index: usize, length: usize) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(length);
    match kind {
        PayloadKind::Compressible => {
            const BLOCK: &[u8] = b"litchi-perf-baseline-compressible-payload-v1\n";
            for offset in 0..length {
                bytes.push(BLOCK[(offset + index) % BLOCK.len()]);
            }
        },
        PayloadKind::Incompressible => {
            let mut state = (index as u64)
                .wrapping_mul(0x9e37_79b9_7f4a_7c15)
                .wrapping_add(0xd1b5_4a32_d192_ed03);
            for _ in 0..length {
                state ^= state << 13;
                state ^= state >> 7;
                state ^= state << 17;
                bytes.push((state >> 24) as u8);
            }
        },
    }
    bytes
}

#[cfg(test)]
fn run_case(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    run_case_with_config(
        case,
        corpus,
        warmup_iterations,
        samples,
        RangeSimulationConfig::default(),
    )
}

fn run_case_with_config(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    range_simulation: RangeSimulationConfig,
) -> Result<CaseResult, Box<dyn Error>> {
    match case {
        Case::ZipIndex => run_zip_index(corpus, warmup_iterations, samples),
        Case::ZipReadOne => run_zip_read_one(corpus, warmup_iterations, samples),
        Case::OpcOpen => run_opc_open(corpus, warmup_iterations, samples),
        Case::OpcOpenOwned => run_opc_open_owned(corpus, warmup_iterations, samples),
        Case::OpcNoopSave => run_opc_noop_save(corpus, warmup_iterations, samples),
        Case::OpcMutatedSave => run_opc_mutated_save(corpus, warmup_iterations, samples),
        Case::OpcSourceOpen => run_opc_source_open(corpus, warmup_iterations, samples, false),
        Case::OpcSourceOpenMainRead => {
            run_opc_source_open(corpus, warmup_iterations, samples, true)
        },
        Case::OpcSourceCachedMainRead => {
            run_opc_source_cached_main_read(corpus, warmup_iterations, samples)
        },
        Case::OpcSourceConcurrentSamePart => {
            run_opc_source_concurrent_same_part(corpus, warmup_iterations, samples)
        },
        Case::OpcSourceCacheBudgetBoundary
        | Case::OpcSourceCacheControlContention
        | Case::OpcSourceCacheManagedContention => {
            Err("OPC source-cache evidence uses its fixed matrix runner".into())
        },
        Case::OpcSourceOverlayOnePartSave => {
            run_opc_source_overlay_one_part_save(corpus, warmup_iterations, samples)
        },
        Case::OpcFileEagerOpen
        | Case::OpcFileSourceOpen
        | Case::OpcFileEagerOnePartAtomicSave
        | Case::OpcFileSourceOnePartAtomicSave
        | Case::CfbFileSameLengthOverlayAtomicSave => {
            Err("filesystem cases are dispatched by the child-process evidence runner".into())
        },
        Case::DocxSourceBackedOneEditSave => {
            run_docx_source_backed_one_edit_save(corpus, warmup_iterations, samples)
        },
        Case::PptxSourceBackedOneEditSave => {
            run_pptx_source_backed_one_edit_save(corpus, warmup_iterations, samples)
        },
        Case::PptxEagerBatchEditSave | Case::PptxSourceBackedBatchEditSave => {
            run_pptx_batch_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::PptxEagerMultiSlideBatchEditSave | Case::PptxSourceBackedMultiSlideBatchEditSave => {
            run_pptx_multi_slide_batch_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxEagerCalculationMetadataEditSave
        | Case::XlsxSourceBackedCalculationMetadataEditSave => {
            run_xlsx_calculation_metadata_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxEagerDefinedNamesEditSave | Case::XlsxSourceBackedDefinedNamesEditSave => {
            run_xlsx_defined_names_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxEagerPageBreakEditSave | Case::XlsxSourceBackedPageBreakEditSave => {
            run_xlsx_page_break_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxEagerPageMarginEditSave | Case::XlsxSourceBackedPageMarginEditSave => {
            run_xlsx_page_margin_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxEagerPageSetupEditSave | Case::XlsxSourceBackedPageSetupEditSave => {
            run_xlsx_page_setup_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxEagerPrintOptionsEditSave | Case::XlsxSourceBackedPrintOptionsEditSave => {
            run_xlsx_print_options_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxEagerSheetProtectionEditSave | Case::XlsxSourceBackedSheetProtectionEditSave => {
            run_xlsx_sheet_protection_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxEagerDataValidationEditSave | Case::XlsxSourceBackedDataValidationEditSave => {
            run_xlsx_data_validation_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxEagerAutoFilterEditSave | Case::XlsxSourceBackedAutoFilterEditSave => {
            run_xlsx_auto_filter_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxEagerConditionalFormattingEditSave
        | Case::XlsxSourceBackedConditionalFormattingEditSave => {
            run_xlsx_conditional_formatting_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxEagerMergeCommitSave | Case::XlsxEagerUnmergeCommitSave => {
            run_xlsx_merge_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::CfbOpen => run_cfb_open(corpus, warmup_iterations, samples),
        Case::CfbListStreams => run_cfb_list_streams(corpus, warmup_iterations, samples),
        Case::CfbReadOne => run_cfb_read_one(corpus, warmup_iterations, samples),
        Case::CfbCreateStreamBorrowed => {
            run_cfb_create_stream(corpus, warmup_iterations, samples, false)
        },
        Case::CfbCreateStreamOwned => {
            run_cfb_create_stream(corpus, warmup_iterations, samples, true)
        },
        Case::OleCommonOpen => run_ole_common_open(corpus, warmup_iterations, samples),
        Case::OleCommonPutStreamPublish => {
            run_ole_common_put_stream_publish(corpus, warmup_iterations, samples)
        },
        Case::OleCommonFinishRender => {
            run_ole_common_finish_render(corpus, warmup_iterations, samples)
        },
        Case::OleCommonOneEditSave => {
            run_ole_common_one_edit_save(corpus, warmup_iterations, samples)
        },
        Case::CfbSharedOpen => run_cfb_shared_open(corpus, warmup_iterations, samples),
        Case::CfbSharedReadOne => run_cfb_shared_read_one(corpus, warmup_iterations, samples),
        Case::CfbSharedConcurrentReads => {
            run_cfb_shared_concurrent_reads(corpus, warmup_iterations, samples)
        },
        Case::CfbSelectiveMiniLegacyRead
        | Case::CfbSelectiveMiniSharedRead
        | Case::CfbSelectiveFatLegacyRead
        | Case::CfbSelectiveFatSharedRead => {
            Err("selective CFB case requires its dedicated corpus dispatcher".into())
        },
        Case::DocFreshWriteTo | Case::XlsFreshWriteTo | Case::PptFreshWriteTo => {
            run_fresh_writer(case, corpus, warmup_iterations, samples)
        },
        Case::DocSemanticOpen
        | Case::DocSemanticListParagraphs
        | Case::DocSemanticOneParagraph
        | Case::DocSemanticFullText
        | Case::DocSemanticNoopEditSave
        | Case::DocSemanticOneEditSave => {
            run_semantic_doc(case, corpus, warmup_iterations, samples)
        },
        Case::DocBodySnapshotListParagraphs => {
            run_doc_body_snapshot_list_paragraphs(corpus, warmup_iterations, samples)
        },
        Case::XlsSemanticOpen
        | Case::XlsSemanticListWorksheets
        | Case::XlsSemanticOneCell
        | Case::XlsSemanticFullCellScan
        | Case::XlsSemanticNoopEditSave
        | Case::XlsSemanticOneEditSave => {
            run_semantic_xls(case, corpus, warmup_iterations, samples)
        },
        Case::XlsValidationReport => run_xls_validation_report(corpus, warmup_iterations, samples),
        Case::XlsCommentsEagerEditSave
        | Case::XlsCommentsSourceBackedEditSave
        | Case::XlsCommentsEagerBatchEditSave
        | Case::XlsCommentsSourceBackedBatchEditSave => {
            run_xls_comments_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsVisibilityEagerEditSave
        | Case::XlsVisibilitySourceBackedEditSave
        | Case::XlsVisibilityEagerBatchEditSave
        | Case::XlsVisibilitySourceBackedBatchEditSave => {
            run_xls_visibility_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::PptSemanticOpen
        | Case::PptSemanticListSlides
        | Case::PptSemanticOneShapeText
        | Case::PptSemanticFullText
        | Case::PptSlideOrderSnapshotOpen
        | Case::PptSemanticNoopEditSave
        | Case::PptSemanticOneEditSave => {
            run_semantic_ppt(case, corpus, warmup_iterations, samples)
        },
        Case::PptTextEditOneEditSave => {
            run_ppt_text_edit_one_edit_save(corpus, warmup_iterations, samples)
        },
        Case::XlsxOpenOwned => run_xlsx_open_owned(corpus, warmup_iterations, samples),
        Case::XlsxListSheets => run_xlsx_list_sheets(corpus, warmup_iterations, samples),
        Case::XlsxFirstCell => run_xlsx_first_cell(corpus, warmup_iterations, samples),
        Case::XlsxFullCellScan => run_xlsx_full_cell_scan(corpus, warmup_iterations, samples),
        Case::XlsxNarrowColumnRangeScan => {
            run_xlsx_narrow_column_range_scan(corpus, warmup_iterations, samples)
        },
        Case::XlsxNoopCommit => run_xlsx_noop_commit(corpus, warmup_iterations, samples),
        Case::XlsxNoopCommitSave => run_xlsx_noop_commit_save(corpus, warmup_iterations, samples),
        Case::XlsxOneCellCommit => run_xlsx_update_commit(corpus, warmup_iterations, samples, 1),
        Case::XlsxOneCellCommitFirstRead => {
            run_xlsx_one_cell_commit_first_read(corpus, warmup_iterations, samples)
        },
        Case::XlsxOneCellCommitSave => {
            run_xlsx_update_commit_save(corpus, warmup_iterations, samples, 1)
        },
        Case::XlsxOnePercentCommit => run_xlsx_update_commit(corpus, warmup_iterations, samples, 0),
        Case::XlsxOnePercentCommitSave => {
            run_xlsx_update_commit_save(corpus, warmup_iterations, samples, 0)
        },
        Case::XlsxEagerCellValuesOneEditSave
        | Case::XlsxSourceBackedCellValuesOneEditSave
        | Case::XlsxEagerCellValuesOnePercentEditSave
        | Case::XlsxSourceBackedCellValuesOnePercentEditSave
        | Case::XlsxEagerCellValuesBatchEditSave
        | Case::XlsxSourceBackedCellValuesBatchEditSave => {
            run_xlsx_cell_values_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::XlsxSourceOpen => run_xlsx_source_open(corpus, warmup_iterations, samples),
        Case::XlsxSourceListSheets => {
            run_xlsx_source_list_sheets(corpus, warmup_iterations, samples)
        },
        Case::XlsxSourceFirstCell => run_xlsx_source_first_cell(corpus, warmup_iterations, samples),
        Case::XlsxSourceNarrowColumnRangeScan => {
            run_xlsx_source_narrow_column_range_scan(corpus, warmup_iterations, samples)
        },
        Case::XlsxStreamingCreate | Case::RtfStreamingCreate => {
            Err("streaming creation cases use their bounded corpus runner".into())
        },
        Case::OpcRangeSourceOpen => {
            run_opc_range_source_open(corpus, warmup_iterations, samples, range_simulation, false)
        },
        Case::OpcRangeSourceOpenMainRead => {
            run_opc_range_source_open(corpus, warmup_iterations, samples, range_simulation, true)
        },
        Case::XlsxRangeSourceOpen => {
            run_xlsx_range_source_open(corpus, warmup_iterations, samples, range_simulation)
        },
        Case::XlsxRangeSourceListSheets => {
            run_xlsx_range_source_list_sheets(corpus, warmup_iterations, samples, range_simulation)
        },
        Case::XlsxRangeSourceFirstCell => {
            run_xlsx_range_source_first_cell(corpus, warmup_iterations, samples, range_simulation)
        },
        Case::XlsxRangeSourceNarrowColumnRangeScan => {
            run_xlsx_range_source_narrow_column_range_scan(
                corpus,
                warmup_iterations,
                samples,
                range_simulation,
            )
        },
        Case::RtfSemanticOpen
        | Case::RtfSemanticParagraphCount
        | Case::RtfSemanticListParagraphs
        | Case::RtfSemanticCollectParagraphs
        | Case::RtfSemanticOneParagraph
        | Case::RtfSemanticFullText
        | Case::RtfSemanticTextToSink
        | Case::RtfSemanticStreamSave
        | Case::RtfSemanticNoopEditSave
        | Case::RtfSemanticOneEditSave
        | Case::RtfSemanticOnePercentEditSave
        | Case::RtfSemanticRemoveParagraphSave
        | Case::RtfSemanticMoveParagraphSave => {
            run_semantic_rtf(case, corpus, warmup_iterations, samples)
        },
        Case::RtfValidationReport => run_rtf_validation_report(corpus, warmup_iterations, samples),
        Case::RtfLogicalTailAppend | Case::RtfLogicalTailNoopSave => {
            run_rtf_logical_tail_append(case, corpus, warmup_iterations, samples)
        },
        Case::DocxSemanticOpen
        | Case::DocxSemanticListParagraphs
        | Case::DocxSemanticOneParagraph
        | Case::DocxSemanticFullText
        | Case::DocxSemanticCreateSmall
        | Case::DocxSemanticNoopEditSave
        | Case::DocxSemanticOneEditSave
        | Case::DocxSemanticOnePercentEditSave => {
            run_semantic_docx(case, corpus, warmup_iterations, samples)
        },
        Case::DocxValidationReport => {
            run_docx_validation_report(corpus, warmup_iterations, samples)
        },
        Case::DocxSectionInventory => {
            run_docx_section_inventory(corpus, warmup_iterations, samples)
        },
        Case::PptxSemanticOpen
        | Case::PptxSemanticListSlides
        | Case::PptxSemanticOneSlide
        | Case::PptxSemanticFullText
        | Case::PptxSemanticCreateSmall
        | Case::PptxSemanticNoopEditSave
        | Case::PptxSemanticOneEditSave
        | Case::PptxSemanticOnePercentEditSave => {
            run_semantic_pptx(case, corpus, warmup_iterations, samples)
        },
        Case::PptxValidationReport => {
            run_pptx_validation_report(corpus, warmup_iterations, samples)
        },
        Case::OdtSemanticOpen
        | Case::OdtSemanticListParagraphs
        | Case::OdtSemanticOneParagraph
        | Case::OdtSemanticFullText
        | Case::OdtSemanticCreateSmall
        | Case::OdtSemanticNoopEditSave
        | Case::OdtSemanticOneEditSave
        | Case::OdtSemanticOnePercentEditSave => {
            run_semantic_odt(case, corpus, warmup_iterations, samples)
        },
        Case::OdfValidationReport => run_odf_validation_report(corpus, warmup_iterations, samples),
        Case::OdfMimetypeRepairPlan => {
            run_odf_mimetype_repair_plan(corpus, warmup_iterations, samples)
        },
        Case::OdtMediaParagraphEditSave => {
            run_odt_media_paragraph_edit_save(corpus, warmup_iterations, samples)
        },
        Case::OdtMediaLineBreakEditSave => {
            run_odt_media_line_break_edit_save(corpus, warmup_iterations, samples)
        },
        Case::OdtMediaAppendRunEditSave => {
            run_odt_media_append_run_edit_save(corpus, warmup_iterations, samples)
        },
        Case::OdtMediaAppendHyperlinkEditSave => {
            run_odt_media_append_hyperlink_edit_save(corpus, warmup_iterations, samples)
        },
        Case::OdtMediaInsertParagraphEditSave | Case::OdtMediaRemoveParagraphEditSave => {
            run_odt_media_structural_paragraph_edit_save(case, corpus, warmup_iterations, samples)
        },
        Case::OdtEmbeddedResourceScalarReplaceSave | Case::OdtEmbeddedResourceBatchReplaceSave => {
            run_odt_embedded_resource_publication(case, corpus, warmup_iterations, samples)
        },
        Case::OdsSemanticOpen
        | Case::OdsSemanticListSheets
        | Case::OdsSemanticOneCell
        | Case::OdsSemanticCellSweep
        | Case::OdsSemanticFullCellText
        | Case::OdsSemanticCreateSmall
        | Case::OdsSemanticNoopEditSave
        | Case::OdsSemanticOneEditSave
        | Case::OdsSemanticOnePercentEditSave => {
            run_semantic_ods(case, corpus, warmup_iterations, samples)
        },
        Case::OdsMediaOneEditSave => {
            run_ods_media_one_edit_save(corpus, warmup_iterations, samples)
        },
        Case::OdpSemanticOpen
        | Case::OdpSemanticListSlides
        | Case::OdpSemanticOneSlide
        | Case::OdpSemanticFullText
        | Case::OdpSemanticCreateSmall
        | Case::OdpSemanticNoopEditSave
        | Case::OdpSemanticOneEditSave => {
            run_semantic_odp(case, corpus, warmup_iterations, samples)
        },
        Case::OdpMediaTextBoxEditSave => {
            run_odp_media_textbox_edit_save(corpus, warmup_iterations, samples)
        },
        Case::OdpMediaTextBoxScalarReplaceSave | Case::OdpMediaTextBoxBatchReplaceSave => {
            run_odp_text_box_model_publication(case, corpus, warmup_iterations, samples)
        },
        Case::OpcOpenSessionScaling | Case::CfbBulkReadScaling => {
            Err("scaling case requires an explicit worker count".into())
        },
    }
}

fn execution_context(
    workers: usize,
    tasks: usize,
    in_flight_bytes: u64,
    input_bytes: u64,
    work_bytes: u64,
) -> Result<ExecutionContext, Box<dyn Error>> {
    let workers = NonZeroUsize::new(workers).ok_or("worker count must be nonzero")?;
    let max_tasks = NonZeroUsize::new(tasks.max(workers.get()))
        .ok_or("execution task count must be nonzero")?;
    let max_bytes =
        NonZeroU64::new(in_flight_bytes.max(1)).ok_or("execution byte bound must be nonzero")?;
    let limits = ExecutionLimits::new(workers, max_tasks, max_bytes, 0)?;
    let memory = in_flight_bytes
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("execution memory budget overflows u64")?;
    let input = input_bytes
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("execution input budget overflows u64")?;
    let work = work_bytes
        .checked_mul(3)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("execution work budget overflows u64")?;
    let objects = u64::try_from(tasks)?
        .checked_mul(4)
        .and_then(|value| value.checked_add(1_024))
        .ok_or("execution object budget overflows u64")?;
    let (_cancellation, token) = CancellationSource::pair();
    Ok(ExecutionContext::new(
        Budget::root(
            "litchi-perf-scaling",
            Limits::new(memory, input, 64 * 1024, objects, 1_024, work),
        ),
        token,
        limits,
    ))
}

fn zip_logical_work(bytes: &[u8]) -> Result<(usize, u64), Box<dyn Error>> {
    let archive = ArchiveReader::new(bytes)?;
    let names = archive.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut logical_bytes = 0u64;
    for name in &names {
        logical_bytes = logical_bytes
            .checked_add(u64::try_from(archive.read(name)?.len())?)
            .ok_or("ZIP logical work bytes overflow u64")?;
    }
    Ok((names.len(), logical_bytes))
}

fn run_scaling_case(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    workers: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    match case {
        Case::OpcOpenSessionScaling => {
            run_opc_open_session_scaling(corpus, warmup_iterations, samples, workers)
        },
        Case::CfbBulkReadScaling => {
            run_cfb_bulk_read_scaling(corpus, warmup_iterations, samples, workers)
        },
        _ => Err("non-scaling case passed to scaling runner".into()),
    }
}

fn run_opc_open_session_scaling(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    workers: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let (logical_tasks, logical_bytes) = zip_logical_work(&corpus.archive)?;
    let input_bytes = u64::try_from(corpus.archive.len())?;
    let kind = corpus_payload_kind(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let context = execution_context(
            workers,
            logical_tasks,
            logical_bytes,
            input_bytes,
            logical_bytes,
        )?;
        let session = OpenSession::new(context)?;
        let started = Instant::now();
        let package = session.from_bytes(&corpus.archive, ReadLimits::default())?;
        let duration = started.elapsed();
        verify_opc_scaling_package(&package, corpus, kind)?;
        std::hint::black_box(&package);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_execution(
        Case::OpcOpenSessionScaling,
        corpus,
        elapsed,
        ExecutionSummary {
            worker_count: workers,
            logical_tasks,
            logical_bytes,
        },
    ))
}

fn verify_opc_scaling_package(
    package: &OpcPackage,
    corpus: &Corpus,
    kind: PayloadKind,
) -> Result<(), Box<dyn Error>> {
    if package.iter_parts().count() != corpus.manifest.entry_count {
        return Err("explicit OPC open session part count differs from manifest".into());
    }
    for index in 0..corpus.manifest.entry_count {
        let uri = PackURI::new(format!("/{}", entry_name(index)))?;
        if package.get_part(&uri)?.blob() != payload_bytes(kind, index, corpus.manifest.entry_bytes)
        {
            return Err(
                format!("explicit OPC open session part {index} differs from corpus").into(),
            );
        }
    }
    Ok(())
}

fn run_cfb_bulk_read_scaling(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    workers: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let logical_tasks = corpus.manifest.entry_count;
    let logical_bytes = u64::try_from(corpus.manifest.uncompressed_payload_bytes)?;
    let input_bytes = u64::try_from(corpus.archive.len())?;
    let names = (0..logical_tasks).map(cfb_entry_name).collect::<Vec<_>>();
    let path_storage = names
        .iter()
        .map(|name| vec![name.as_str()])
        .collect::<Vec<_>>();
    let paths = path_storage.iter().map(Vec::as_slice).collect::<Vec<_>>();
    let kind = corpus_payload_kind(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = cfb_instrumented_source(corpus);
        let ole = SharedOleFile::open_with_limits(source.clone(), cfb_shared_limits(corpus)?)?;
        let context = execution_context(
            workers,
            logical_tasks,
            logical_bytes,
            input_bytes,
            logical_bytes,
        )?;
        let session = ole.bulk_read(context);
        let preload = session.read_streams(&paths)?;
        verify_cfb_bulk_outputs(&preload, corpus, kind)?;
        source.reset();
        let started = Instant::now();
        let outputs = session.read_streams(&paths)?;
        let duration = started.elapsed();
        verify_cfb_bulk_outputs(&outputs, corpus, kind)?;
        std::hint::black_box(&outputs);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_execution(
        Case::CfbBulkReadScaling,
        corpus,
        elapsed,
        ExecutionSummary {
            worker_count: workers,
            logical_tasks,
            logical_bytes,
        },
    ))
}

fn verify_cfb_bulk_outputs(
    outputs: &[Vec<u8>],
    corpus: &Corpus,
    kind: PayloadKind,
) -> Result<(), Box<dyn Error>> {
    if outputs.len() != corpus.manifest.entry_count {
        return Err("CFB bulk read output count differs from manifest".into());
    }
    for (index, output) in outputs.iter().enumerate() {
        if *output != payload_bytes(kind, index, corpus.manifest.entry_bytes) {
            return Err(format!("CFB bulk read output {index} differs from corpus").into());
        }
    }
    Ok(())
}

fn xlsx_spec(corpus: &Corpus) -> Result<&XlsxCorpus, Box<dyn Error>> {
    corpus
        .xlsx
        .as_ref()
        .ok_or_else(|| "XLSX case has no generated XLSX corpus".into())
}

fn xlsx_output_ceiling(bytes: usize) -> Result<u64, Box<dyn Error>> {
    u64::try_from(bytes)?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or_else(|| "XLSX sequential output ceiling overflows u64".into())
}

fn verify_semantic_rtf(
    document: &litchi_rtf::Document,
    shape: SemanticShape,
    variant: RtfSemanticVariant,
    updated: &[usize],
) -> Result<(), Box<dyn Error>> {
    let paragraph_count = semantic_rtf_paragraph_count(shape, variant);
    if document.paragraph_count() != paragraph_count {
        return Err("semantic RTF paragraph count differs from specification".into());
    }
    let mut count = 0usize;
    for (index, paragraph) in document.body().paragraphs().enumerate() {
        if paragraph.to_text()
            != semantic_rtf_variant_text(variant, index, updated.binary_search(&index).is_ok())
        {
            return Err("semantic RTF paragraph text differs from specification".into());
        }
        count = count
            .checked_add(1)
            .ok_or("semantic RTF paragraph count overflows usize")?;
    }
    if count != paragraph_count {
        return Err("semantic RTF paragraph traversal differs from specification".into());
    }
    if document.text() != semantic_rtf_expected_text(shape, variant, updated) {
        return Err("semantic RTF full text differs from specification".into());
    }
    if variant == RtfSemanticVariant::Watermark {
        let header_shapes = document
            .sections()
            .iter()
            .flat_map(|section| &section.headers_footers)
            .flat_map(|header_footer| &header_footer.shapes)
            .collect::<Vec<_>>();
        if document.sections().len() != 1
            || header_shapes.len() != 3
            || header_shapes
                .first()
                .and_then(|shape| shape.property("gtextUNICODE"))
                != Some("ASAP")
        {
            return Err("semantic RTF watermark drawing projection differs from fixture".into());
        }
    }
    Ok(())
}

fn semantic_rtf_lifecycle_projection(
    case: Case,
    shape: SemanticShape,
) -> Result<Vec<String>, Box<dyn Error>> {
    let paragraph_count = shape.rtf_paragraphs();
    if paragraph_count < 2 {
        return Err("semantic RTF lifecycle corpus needs at least two paragraphs".into());
    }
    let selected = paragraph_count / 2;
    let mut paragraphs = (0..paragraph_count)
        .map(|index| semantic_rtf_variant_text(RtfSemanticVariant::Plain, index, false))
        .collect::<Vec<_>>();
    match case {
        Case::RtfSemanticRemoveParagraphSave => {
            paragraphs.remove(selected);
        },
        Case::RtfSemanticMoveParagraphSave => {
            let paragraph = paragraphs.remove(0);
            paragraphs.push(paragraph);
        },
        _ => return Err("non-lifecycle RTF case requested a lifecycle projection".into()),
    }
    Ok(paragraphs)
}

fn stage_semantic_rtf_lifecycle(
    case: Case,
    document: &litchi_rtf::Document,
) -> Result<litchi_rtf::edit::Commit, Box<dyn Error>> {
    let paragraph_count = document.paragraph_count();
    if paragraph_count < 2 {
        return Err("semantic RTF lifecycle corpus needs at least two paragraphs".into());
    }
    let selected = paragraph_count / 2;
    let mut edit = document.edit();
    match case {
        Case::RtfSemanticRemoveParagraphSave => {
            edit.remove_paragraph(selected)?;
        },
        Case::RtfSemanticMoveParagraphSave => {
            edit.move_paragraph(0, paragraph_count - 1)?;
        },
        _ => return Err("non-lifecycle RTF case reached lifecycle staging".into()),
    }
    Ok(edit.commit()?)
}

fn semantic_rtf_durable_limits() -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        1024 * 1024,
        1,
        8,
        256 * 1024,
        512 * 1024,
    )
}

fn verify_semantic_rtf_lifecycle_projection(
    document: &litchi_rtf::Document,
    expected: &[String],
) -> Result<(), Box<dyn Error>> {
    if document.paragraph_count() != expected.len() {
        return Err("semantic RTF lifecycle paragraph count differs from specification".into());
    }
    let paragraphs = document
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.to_text())
        .collect::<Vec<_>>();
    if paragraphs != expected || document.text() != expected.join("\n") {
        return Err("semantic RTF lifecycle full projection differs from specification".into());
    }
    Ok(())
}

fn verify_semantic_rtf_lifecycle_commit(
    case: Case,
    source: &litchi_rtf::Document,
    commit: &litchi_rtf::edit::Commit,
    expected_projection: &[String],
    expected_bytes: &[u8],
) -> Result<(), Box<dyn Error>> {
    if !commit.diagnostics().changed() || commit.diagnostics().operation_count() != 1 {
        return Err("semantic RTF lifecycle commit has unexpected diagnostics".into());
    }
    if commit.snapshot().to_bytes()? != expected_bytes {
        return Err("semantic RTF lifecycle commit differs from expected bytes".into());
    }
    let reopened = litchi_rtf::Document::from_bytes(expected_bytes)?;
    verify_semantic_rtf_lifecycle_projection(&reopened, expected_projection)?;

    let applied = commit.patch().apply(source)?;
    if applied.to_bytes()? != expected_bytes {
        return Err("semantic RTF lifecycle patch replay differs from publication".into());
    }
    let restored = commit.patch().inverse().apply(&applied)?;
    if restored.to_bytes()? != source.to_bytes()? {
        return Err("semantic RTF lifecycle inverse did not restore exact source bytes".into());
    }

    let limits = semantic_rtf_durable_limits();
    let durable = commit.patch().to_durable(limits)?;
    let encoded = durable.to_deterministic_json()?;
    let decoded =
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
            &encoded, limits,
        )?;
    let durable_applied = source.apply_durable(&decoded)?;
    if durable_applied.to_bytes()? != expected_bytes {
        return Err("semantic RTF durable lifecycle replay differs from publication".into());
    }
    let durable_restored = durable_applied.apply_durable(&decoded.inverse())?;
    if durable_restored.to_bytes()? != source.to_bytes()? {
        return Err("semantic RTF durable inverse did not restore exact source bytes".into());
    }

    let mut stale_edit = source.edit();
    stale_edit.replace_paragraph_text(0, "litchi-perf-baseline-rtf-stale-source")?;
    let stale = stale_edit.commit()?.into_snapshot();
    if !matches!(
        stale.apply_durable(&decoded),
        Err(litchi_rtf::edit::Error::PatchConflict)
    ) {
        return Err("semantic RTF durable lifecycle patch accepted a stale source".into());
    }

    if case == Case::RtfSemanticMoveParagraphSave {
        let mut noop = source.edit();
        noop.move_paragraph(0, 0)?;
        let noop = noop.commit()?;
        if noop.diagnostics().changed()
            || noop.diagnostics().operation_count() != 1
            || !noop.snapshot().same_snapshot(source)
            || noop.snapshot().to_bytes()? != source.to_bytes()?
            || !noop.patch().to_durable(limits)?.operations().is_empty()
        {
            return Err("semantic RTF equal-position move was not an exact no-op".into());
        }
    }
    Ok(())
}

fn verify_semantic_docx(
    package: &litchi_docx::Package,
    shape: SemanticShape,
    updated: &[usize],
) -> Result<(), Box<dyn Error>> {
    let document = package.document()?;
    if document.paragraph_count()? != shape.docx_paragraphs() {
        return Err("semantic DOCX paragraph count differs from specification".into());
    }
    let paragraphs = document.paragraphs()?;
    if paragraphs.len() != shape.docx_paragraphs() {
        return Err("semantic DOCX paragraph list differs from specification".into());
    }
    let mut full_text = Vec::with_capacity(paragraphs.len());
    for (index, paragraph) in paragraphs.into_iter().enumerate() {
        let expected = semantic_docx_text(index, updated.contains(&index));
        let actual = paragraph.text()?;
        if actual != expected {
            return Err("semantic DOCX paragraph text differs from specification".into());
        }
        full_text.push(actual);
    }
    if document.text()? != full_text.concat() {
        return Err("semantic DOCX full text differs from paragraph scan".into());
    }
    Ok(())
}

fn verify_semantic_pptx(
    package: &litchi_pptx::Package,
    shape: SemanticShape,
    updated: &[usize],
) -> Result<(), Box<dyn Error>> {
    let presentation = package.presentation()?;
    if presentation.slide_count()? != shape.pptx_slides() {
        return Err("semantic PPTX slide count differs from specification".into());
    }
    let mut presentation_text = Vec::with_capacity(shape.pptx_slides());
    for slide_index in 0..shape.pptx_slides() {
        let slide = presentation
            .slide(slide_index)?
            .ok_or("semantic PPTX slide is missing")?;
        let scene = slide.shapes()?;
        if scene.len() != shape.pptx_text_boxes_per_slide() {
            return Err("semantic PPTX shape count differs from specification".into());
        }
        let mut slide_text = Vec::with_capacity(scene.len());
        for (shape_index, object) in scene.iter().enumerate() {
            let linear = slide_index * shape.pptx_text_boxes_per_slide() + shape_index;
            let expected = semantic_pptx_text(slide_index, shape_index, updated.contains(&linear));
            if object.text() != Some(expected.as_str()) {
                return Err("semantic PPTX shape text differs from specification".into());
            }
            slide_text.push(expected);
        }
        let expected_slide_text = slide_text.join("\n");
        if slide.text()? != expected_slide_text {
            return Err("semantic PPTX slide text differs from shape scan".into());
        }
        presentation_text.push(expected_slide_text);
    }
    if presentation.text()? != presentation_text.join("\n") {
        return Err("semantic PPTX full text differs from slide scan".into());
    }
    Ok(())
}

fn verify_pptx_source_edit_semantics(
    package: &litchi_pptx::Package,
    updated_shapes: usize,
) -> Result<(), Box<dyn Error>> {
    let target_slide = PPTX_SOURCE_SLIDE_COUNT / 2;
    verify_pptx_source_edit_semantics_for(package, &[(target_slide, updated_shapes)])
}

fn verify_pptx_source_edit_semantics_for(
    package: &litchi_pptx::Package,
    updated_slides: &[(usize, usize)],
) -> Result<(), Box<dyn Error>> {
    let presentation = package.presentation()?;
    if presentation.slide_count()? != PPTX_SOURCE_SLIDE_COUNT {
        return Err("PPTX source-edit slide count differs from specification".into());
    }
    let mut presentation_text = Vec::with_capacity(PPTX_SOURCE_SLIDE_COUNT);
    for slide_index in 0..PPTX_SOURCE_SLIDE_COUNT {
        let slide = presentation
            .slide(slide_index)?
            .ok_or("PPTX source-edit slide is missing")?;
        let scene = slide.shapes()?;
        let expected_shape_count = PPTX_SOURCE_TEXT_BOXES_PER_SLIDE
            + usize::from(slide_index < PPTX_SOURCE_MEDIA_ENTRY_COUNT);
        if scene.len() != expected_shape_count {
            return Err("PPTX source-edit shape count differs from specification".into());
        }
        let text_shapes = scene
            .iter()
            .filter_map(|shape| shape.text())
            .collect::<Vec<_>>();
        if text_shapes.len() != PPTX_SOURCE_TEXT_BOXES_PER_SLIDE {
            return Err("PPTX source-edit text shape count differs from specification".into());
        }
        let mut slide_text = Vec::with_capacity(PPTX_SOURCE_TEXT_BOXES_PER_SLIDE);
        for (shape_index, actual) in text_shapes.into_iter().enumerate() {
            let updated_shapes = updated_slides
                .iter()
                .find_map(|&(position, count)| (position == slide_index).then_some(count))
                .unwrap_or_default();
            let is_updated = shape_index < updated_shapes;
            let expected = semantic_pptx_text(slide_index, shape_index, is_updated);
            if actual != expected {
                return Err("PPTX source-edit shape text differs from specification".into());
            }
            slide_text.push(expected);
        }
        let expected_slide_text = slide_text.join("\n");
        if slide.text()? != expected_slide_text {
            return Err("PPTX source-edit slide text differs from shape scan".into());
        }
        presentation_text.push(expected_slide_text);
    }
    if presentation.text()? != presentation_text.join("\n") {
        return Err("PPTX source-edit full text differs from slide scan".into());
    }
    Ok(())
}

fn verify_semantic_odt(
    document: &litchi_odt::Document,
    shape: SemanticShape,
    updated: &[usize],
) -> Result<(), Box<dyn Error>> {
    let paragraphs = document.paragraphs()?;
    if paragraphs.len() != shape.docx_paragraphs() {
        return Err("semantic ODT paragraph count differs from specification".into());
    }
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let is_updated = updated.binary_search(&index).is_ok();
        if paragraph.text()? != semantic_odt_text(index, is_updated) {
            return Err("semantic ODT paragraph text differs from specification".into());
        }
    }
    let expected = (0..shape.docx_paragraphs())
        .map(|index| semantic_odt_text(index, updated.binary_search(&index).is_ok()))
        .collect::<Vec<_>>()
        .join("\n");
    if document.text()? != expected {
        return Err("semantic ODT full text differs from paragraph scan".into());
    }
    Ok(())
}

fn verify_odt_media_archive(bytes: &[u8], updated: bool) -> Result<(), Box<dyn Error>> {
    let document = litchi_odt::Document::from_bytes(bytes.to_vec())?;
    let updated = updated.then_some(SemanticShape::Medium.docx_paragraphs() / 2);
    verify_semantic_odt(&document, SemanticShape::Medium, updated.as_slice())?;

    let package = litchi_odf_common::core::OwnedPackage::from_bytes(bytes.to_vec())?;
    let package = package.package()?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        let path = odt_media_path(index);
        if package.manifest().get_media_type(&path) != Some("application/octet-stream") {
            return Err(format!("media-rich ODT manifest entry differs for '{path}'").into());
        }
        if package.get_file(&path)? != odt_media_payload(index) {
            return Err(format!("media-rich ODT payload differs for '{path}'").into());
        }
    }
    Ok(())
}

fn verify_odt_media_line_break_archive(bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let target = shape.docx_paragraphs() / 2;
    let document = litchi_odt::Document::from_bytes(bytes.to_vec())?;
    let paragraphs = document.paragraphs()?;
    if paragraphs.len() != shape.docx_paragraphs() {
        return Err("media-rich ODT line-break paragraph count differs from specification".into());
    }
    let mut expected_text = Vec::with_capacity(shape.docx_paragraphs());
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let mut expected = semantic_odt_text(index, false);
        if index == target {
            expected.push('\n');
        }
        if paragraph.text()? != expected {
            return Err("media-rich ODT line-break text differs from specification".into());
        }
        expected_text.push(expected);
    }
    if document.text()? != expected_text.join("\n") {
        return Err("media-rich ODT line-break full text differs from paragraph scan".into());
    }

    let package = litchi_odf_common::core::OwnedPackage::from_bytes(bytes.to_vec())?;
    let package = package.package()?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        let path = odt_media_path(index);
        if package.manifest().get_media_type(&path) != Some("application/octet-stream") {
            return Err(format!("media-rich ODT manifest entry differs for '{path}'").into());
        }
        if package.get_file(&path)? != odt_media_payload(index) {
            return Err(format!("media-rich ODT payload differs for '{path}'").into());
        }
    }
    Ok(())
}

fn verify_odt_media_append_run_archive(bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let target = shape.docx_paragraphs() / 2;
    let document = litchi_odt::Document::from_bytes(bytes.to_vec())?;
    let paragraphs = document.paragraphs()?;
    if paragraphs.len() != shape.docx_paragraphs() {
        return Err("media-rich ODT append-run paragraph count differs from specification".into());
    }
    let mut expected_text = Vec::with_capacity(shape.docx_paragraphs());
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let mut expected = semantic_odt_text(index, false);
        if index == target {
            expected.push_str(ODT_MEDIA_APPEND_RUN_TEXT);
        }
        if paragraph.text()? != expected {
            return Err("media-rich ODT append-run text differs from specification".into());
        }
        expected_text.push(expected);
    }
    if document.text()? != expected_text.join("\n") {
        return Err("media-rich ODT append-run full text differs from paragraph scan".into());
    }

    let package = litchi_odf_common::core::OwnedPackage::from_bytes(bytes.to_vec())?;
    let package = package.package()?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        let path = odt_media_path(index);
        if package.manifest().get_media_type(&path) != Some("application/octet-stream") {
            return Err(format!("media-rich ODT manifest entry differs for '{path}'").into());
        }
        if package.get_file(&path)? != odt_media_payload(index) {
            return Err(format!("media-rich ODT payload differs for '{path}'").into());
        }
    }
    Ok(())
}

fn verify_odt_media_append_hyperlink_archive(bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let target = shape.docx_paragraphs() / 2;
    let document = litchi_odt::Document::from_bytes(bytes.to_vec())?;
    let paragraphs = document.paragraphs()?;
    if paragraphs.len() != shape.docx_paragraphs() {
        return Err(
            "media-rich ODT append-hyperlink paragraph count differs from specification".into(),
        );
    }
    let mut expected_text = Vec::with_capacity(shape.docx_paragraphs());
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let mut expected = semantic_odt_text(index, false);
        if index == target {
            expected.push_str(ODT_MEDIA_APPEND_HYPERLINK_TEXT);
        }
        if paragraph.text()? != expected {
            return Err("media-rich ODT append-hyperlink text differs from specification".into());
        }
        expected_text.push(expected);
    }
    if document.text()? != expected_text.join("\n") {
        return Err("media-rich ODT append-hyperlink full text differs from paragraph scan".into());
    }
    if document.hyperlinks()?
        != [(
            ODT_MEDIA_APPEND_HYPERLINK_TEXT.to_string(),
            ODT_MEDIA_APPEND_HYPERLINK_HREF.to_string(),
        )]
    {
        return Err("media-rich ODT hyperlink semantics differ from specification".into());
    }

    let package = litchi_odf_common::core::OwnedPackage::from_bytes(bytes.to_vec())?;
    let package = package.package()?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        let path = odt_media_path(index);
        if package.manifest().get_media_type(&path) != Some("application/octet-stream") {
            return Err(format!("media-rich ODT manifest entry differs for '{path}'").into());
        }
        if package.get_file(&path)? != odt_media_payload(index) {
            return Err(format!("media-rich ODT payload differs for '{path}'").into());
        }
    }
    Ok(())
}

fn verify_odt_media_structural_paragraph_archive(
    bytes: &[u8],
    inserted: bool,
) -> Result<(), Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let original_count = shape.docx_paragraphs();
    let target = original_count / 2;
    let document = litchi_odt::Document::from_bytes(bytes.to_vec())?;
    let paragraphs = document.paragraphs()?;
    let expected_count = if inserted {
        original_count
            .checked_add(1)
            .ok_or("media-rich ODT inserted paragraph count overflows usize")?
    } else {
        original_count
            .checked_sub(1)
            .ok_or("media-rich ODT removed paragraph count underflows usize")?
    };
    if paragraphs.len() != expected_count {
        return Err("media-rich ODT structural paragraph count differs from specification".into());
    }
    let mut expected_text = Vec::with_capacity(expected_count);
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let expected = if inserted && index == target {
            ODT_MEDIA_INSERT_PARAGRAPH_TEXT.to_owned()
        } else {
            let original = if inserted {
                index - usize::from(index > target)
            } else {
                index + usize::from(index >= target)
            };
            semantic_odt_text(original, false)
        };
        if paragraph.text()? != expected {
            return Err(
                "media-rich ODT structural paragraph text differs from specification".into(),
            );
        }
        expected_text.push(expected);
    }
    if document.text()? != expected_text.join("\n") {
        return Err("media-rich ODT structural full text differs from paragraph scan".into());
    }

    let package = litchi_odf_common::core::OwnedPackage::from_bytes(bytes.to_vec())?;
    let package = package.package()?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        let path = odt_media_path(index);
        if package.manifest().get_media_type(&path) != Some("application/octet-stream") {
            return Err(format!("media-rich ODT manifest entry differs for '{path}'").into());
        }
        if package.get_file(&path)? != odt_media_payload(index) {
            return Err(format!("media-rich ODT payload differs for '{path}'").into());
        }
    }
    Ok(())
}

type OdtResourceProjection = (String, Vec<(String, String, String, String)>);

fn verify_odt_resource_batch_archive(
    bytes: &[u8],
    updated: bool,
) -> Result<OdtResourceProjection, Box<dyn Error>> {
    let document = litchi_odt::Document::from_bytes(bytes.to_vec())?;
    verify_semantic_odt(&document, SemanticShape::Medium, &[])?;
    if !document.embedded_objects()?.is_empty() {
        return Err("ODT embedded-resource corpus unexpectedly contains objects".into());
    }
    let images = document.images()?;
    if images.len() != ODT_RESOURCE_BATCH_COUNT {
        return Err("ODT embedded-resource image count differs from specification".into());
    }

    let package = litchi_odt::core::OwnedPackage::from_bytes(bytes.to_vec())?;
    let package = package.package()?;
    let mut projection = Vec::with_capacity(ODT_RESOURCE_BATCH_COUNT);
    for (index, image) in images.iter().enumerate() {
        let name = odt_resource_batch_name(index);
        let path = odt_resource_batch_path(index, updated);
        let litchi_odf_common::media::Source::PackagePart {
            href,
            path: actual_path,
            manifest_media_type,
        } = &image.source
        else {
            return Err(format!("ODT embedded-resource image {index} is not packaged").into());
        };
        if image.part != litchi_odf_common::drawing::Part::Content
            || image.frame.as_ref().and_then(|frame| frame.name.as_deref()) != Some(name.as_str())
            || image.xml_id.is_some()
            || image.declared_media_type.as_deref() != Some("image/png")
            || image.alternative_index != 0
            || href != &path
            || actual_path != &path
            || manifest_media_type.as_deref() != Some("image/png")
            || package.manifest().get_media_type(&path) != Some("image/png")
        {
            return Err(format!("ODT embedded-resource owner {index} differs").into());
        }
        let actual_payload = package.get_file(&path)?;
        let expected_payload = odt_resource_batch_payload(index, updated);
        let actual_digest = sha256_hex(&actual_payload);
        let expected_digest = sha256_hex(&expected_payload);
        if actual_payload.len() != ODT_RESOURCE_PAYLOAD_BYTES || actual_digest != expected_digest {
            return Err(
                format!("ODT embedded-resource payload digest differs for '{path}'").into(),
            );
        }
        projection.push((name, path, "image/png".to_owned(), actual_digest));
    }
    if updated {
        for index in 0..ODT_RESOURCE_BATCH_COUNT {
            let path = odt_resource_batch_path(index, false);
            if package.manifest().get_media_type(&path) != Some("image/png")
                || sha256_hex(&package.get_file(&path)?)
                    != sha256_hex(&odt_resource_batch_payload(index, false))
            {
                return Err(format!(
                    "ODT embedded-resource displaced source payload differs for '{path}'"
                )
                .into());
            }
        }
    }
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        let path = odt_media_path(index);
        if package.manifest().get_media_type(&path) != Some("application/octet-stream")
            || sha256_hex(&package.get_file(&path)?) != sha256_hex(&odt_media_payload(index))
        {
            return Err(
                format!("ODT embedded-resource retained media differs for '{path}'").into(),
            );
        }
    }
    Ok((document.text()?, projection))
}

fn verify_odt_resource_batch_raw_members(
    source: &[u8],
    published: &[u8],
) -> Result<(), Box<dyn Error>> {
    let identical = litchi_odf_common::package::raw_identical_members(source, published)
        .ok_or("ODT embedded-resource raw-member comparison failed")?;
    if identical.contains("content.xml") {
        return Err("ODT embedded-resource content.xml remained raw-identical".into());
    }
    for path in ArchiveReader::new(source)?.file_names() {
        if !matches!(path, "content.xml" | "META-INF/manifest.xml") && !identical.contains(path) {
            return Err(format!(
                "ODT embedded-resource publication changed untouched raw member '{path}'; identical={identical:?}"
            )
            .into());
        }
    }
    Ok(())
}

fn semantic_ods_full_cell_text(
    spreadsheet: &litchi_ods::Spreadsheet,
    shape: SemanticShape,
) -> Result<String, Box<dyn Error>> {
    let mut values = Vec::with_capacity(shape.ods_cell_count());
    for sheet in 0..shape.ods_sheet_count() {
        let name = semantic_ods_sheet_name(sheet);
        for row in 0..shape.ods_rows_per_sheet() {
            for column in 0..shape.ods_columns_per_sheet() {
                let cell = spreadsheet
                    .cell(&name, row, column)
                    .ok_or("semantic ODS sheet is missing")?;
                let litchi_ods::CellView::Stored(cell) = cell else {
                    return Err("semantic ODS cell is missing".into());
                };
                values.push(cell.text.clone());
            }
        }
    }
    Ok(values.join("\n"))
}

fn semantic_ods_cell_sweep(
    spreadsheet: &litchi_ods::Spreadsheet,
    shape: SemanticShape,
) -> Result<usize, Box<dyn Error>> {
    let mut stored_cells = 0usize;
    for sheet in 0..shape.ods_sheet_count() {
        let name = semantic_ods_sheet_name(sheet);
        for row in 0..shape.ods_rows_per_sheet() {
            for column in 0..shape.ods_columns_per_sheet() {
                let cell = spreadsheet
                    .cell(&name, row, column)
                    .ok_or("semantic ODS sheet is missing")?;
                let litchi_ods::CellView::Stored(cell) = cell else {
                    return Err("semantic ODS cell is missing".into());
                };
                std::hint::black_box(cell);
                stored_cells = stored_cells
                    .checked_add(1)
                    .ok_or("semantic ODS stored-cell count overflowed")?;
            }
        }
    }
    Ok(stored_cells)
}

fn semantic_ods_flat_index(shape: SemanticShape, sheet: usize, row: usize, column: usize) -> usize {
    sheet * shape.ods_rows_per_sheet() * shape.ods_columns_per_sheet()
        + row * shape.ods_columns_per_sheet()
        + column
}

fn expected_semantic_ods_full_cell_text(shape: SemanticShape, updated_indices: &[usize]) -> String {
    let mut values = Vec::with_capacity(shape.ods_cell_count());
    for sheet in 0..shape.ods_sheet_count() {
        for row in 0..shape.ods_rows_per_sheet() {
            for column in 0..shape.ods_columns_per_sheet() {
                let index = semantic_ods_flat_index(shape, sheet, row, column);
                values.push(semantic_ods_text(
                    sheet,
                    row,
                    column,
                    updated_indices.binary_search(&index).is_ok(),
                ));
            }
        }
    }
    values.join("\n")
}

fn verify_semantic_ods(
    spreadsheet: &litchi_ods::Spreadsheet,
    shape: SemanticShape,
    updated: bool,
) -> Result<(), Box<dyn Error>> {
    let updated_indices = if updated {
        vec![semantic_ods_flat_index(
            shape,
            shape.ods_sheet_count() / 2,
            shape.ods_rows_per_sheet() / 2,
            shape.ods_columns_per_sheet() / 2,
        )]
    } else {
        Vec::new()
    };
    verify_semantic_ods_updates(spreadsheet, shape, &updated_indices)
}

fn verify_semantic_ods_updates(
    spreadsheet: &litchi_ods::Spreadsheet,
    shape: SemanticShape,
    updated_indices: &[usize],
) -> Result<(), Box<dyn Error>> {
    if spreadsheet.sheets().len() != shape.ods_sheet_count() {
        return Err("semantic ODS sheet count differs from specification".into());
    }
    for sheet in 0..shape.ods_sheet_count() {
        let name = semantic_ods_sheet_name(sheet);
        let sheet_value = spreadsheet
            .sheet(&name)
            .ok_or("semantic ODS named sheet is missing")?;
        if sheet_value.logical_row_count() != shape.ods_rows_per_sheet()
            || sheet_value.logical_column_count() != shape.ods_columns_per_sheet()
        {
            return Err("semantic ODS sheet dimensions differ from specification".into());
        }
        for row in 0..shape.ods_rows_per_sheet() {
            for column in 0..shape.ods_columns_per_sheet() {
                let index = semantic_ods_flat_index(shape, sheet, row, column);
                let is_updated = updated_indices.binary_search(&index).is_ok();
                let cell = spreadsheet
                    .cell(&name, row, column)
                    .ok_or("semantic ODS sheet is missing")?;
                let litchi_ods::CellView::Stored(cell) = cell else {
                    return Err("semantic ODS cell is missing".into());
                };
                if cell.text != semantic_ods_text(sheet, row, column, is_updated) {
                    return Err("semantic ODS cell text differs from specification".into());
                }
            }
        }
    }
    if semantic_ods_full_cell_text(spreadsheet, shape)?
        != expected_semantic_ods_full_cell_text(shape, updated_indices)
    {
        return Err("semantic ODS full cell text differs from specification".into());
    }
    Ok(())
}

fn verify_ods_media_archive(bytes: &[u8], updated: bool) -> Result<(), Box<dyn Error>> {
    let spreadsheet = litchi_ods::Spreadsheet::from_bytes(bytes.to_vec())?;
    verify_semantic_ods(&spreadsheet, SemanticShape::Medium, updated)?;

    let package = litchi_odf_common::core::OwnedPackage::from_bytes(bytes.to_vec())?;
    let package = package.package()?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        let path = ods_media_path(index);
        if package.manifest().get_media_type(&path) != Some("application/octet-stream") {
            return Err(format!("media-rich ODS manifest entry differs for '{path}'").into());
        }
        if package.get_file(&path)? != ods_media_payload(index) {
            return Err(format!("media-rich ODS payload differs for '{path}'").into());
        }
    }
    Ok(())
}

fn verify_semantic_odp(
    presentation: &litchi_odp::Presentation,
    shape: SemanticShape,
    updated: bool,
) -> Result<(), Box<dyn Error>> {
    let slides = presentation.slides()?;
    let expected_slide_count = shape.pptx_slides() + usize::from(updated);
    if slides.len() != expected_slide_count {
        return Err("semantic ODP slide count differs from specification".into());
    }
    let expected = (0..shape.pptx_slides())
        .map(|index| {
            format!(
                "{}\n{}",
                semantic_odp_title(index, false),
                semantic_odp_text(index, false)
            )
        })
        .chain(updated.then(|| {
            let index = shape.pptx_slides();
            format!(
                "{}\n{}",
                semantic_odp_title(index, true),
                semantic_odp_text(index, true)
            )
        }))
        .collect::<Vec<_>>();
    for (index, slide) in slides.iter().enumerate() {
        let is_added = updated && index == shape.pptx_slides();
        if slide.title.as_deref() != Some(semantic_odp_title(index, is_added).as_str())
            || slide.all_text() != expected[index]
        {
            return Err("semantic ODP slide differs from specification".into());
        }
    }
    if presentation.text()? != expected.join("\n\n") {
        return Err("semantic ODP full text differs from slide scan".into());
    }
    Ok(())
}

fn verify_odp_media_archive(bytes: &[u8], text_box_added: bool) -> Result<(), Box<dyn Error>> {
    let presentation = litchi_odp::Presentation::from_bytes(bytes.to_vec())?;
    if text_box_added {
        let slides = presentation.slides()?;
        if slides.len() != SemanticShape::Medium.pptx_slides() {
            return Err("media-rich ODP slide count differs from specification".into());
        }
        for (index, slide) in slides.iter().enumerate() {
            let original = format!(
                "{}\n{}",
                semantic_odp_title(index, false),
                semantic_odp_text(index, false)
            );
            if slide.title.as_deref() != Some(semantic_odp_title(index, false).as_str())
                || (index == 0 && (slide.all_text() != format!("{original}\n{}", odp_media_text())))
                || (index != 0 && slide.all_text() != original)
            {
                return Err("media-rich ODP slide differs from specification".into());
            }
        }
    } else {
        verify_semantic_odp(&presentation, SemanticShape::Medium, false)?;
    }
    let snapshot = presentation.snapshot()?;
    let inventory = snapshot.rich_content()?;
    let matching = inventory
        .text_boxes()
        .iter()
        .filter(|text_box| text_box.name() == ODP_MEDIA_TEXT_BOX_NAME)
        .collect::<Vec<_>>();
    if matching.len() != usize::from(text_box_added) {
        return Err("media-rich ODP text-box inventory differs from specification".into());
    }
    if let Some(text_box) = matching.first()
        && (text_box.page() != 0
            || text_box.paragraph_count() != 1
            || !text_box.xml().contains(&odp_media_text()))
    {
        return Err("media-rich ODP inserted text box differs from specification".into());
    }

    let package = litchi_odf_common::core::OwnedPackage::from_bytes(bytes.to_vec())?;
    let package = package.package()?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        let path = odp_media_path(index);
        if package.manifest().get_media_type(&path) != Some("application/octet-stream") {
            return Err(format!("media-rich ODP manifest entry differs for '{path}'").into());
        }
        if package.get_file(&path)? != odp_media_payload(index) {
            return Err(format!("media-rich ODP payload differs for '{path}'").into());
        }
    }
    Ok(())
}

fn verify_odp_text_box_batch_archive(bytes: &[u8], updated: bool) -> Result<(), Box<dyn Error>> {
    let presentation = litchi_odp::Presentation::from_bytes(bytes.to_vec())?;
    let slides = presentation.slides()?;
    if slides.len() != SemanticShape::Medium.pptx_slides() {
        return Err("ODP text-box batch slide count differs from specification".into());
    }
    for (page, slide) in slides.iter().enumerate() {
        let mut expected = format!(
            "{}\n{}",
            semantic_odp_title(page, false),
            semantic_odp_text(page, false)
        );
        for index in 0..ODP_TEXT_BOX_BATCH_COUNT {
            if odp_text_box_batch_page(index) == page {
                expected.push('\n');
                expected.push_str(&odp_text_box_batch_text(index, updated));
            }
        }
        if slide.title.as_deref() != Some(semantic_odp_title(page, false).as_str())
            || slide.all_text() != expected
        {
            return Err("ODP text-box batch slide projection differs from specification".into());
        }
    }
    let expected_full_text = slides
        .iter()
        .map(litchi_odp::Slide::all_text)
        .collect::<Vec<_>>()
        .join("\n\n");
    if presentation.text()? != expected_full_text {
        return Err("ODP text-box batch full text differs from slide scan".into());
    }

    let snapshot = presentation.snapshot()?;
    let inventory = snapshot.rich_content()?;
    for index in 0..ODP_TEXT_BOX_BATCH_COUNT {
        let name = odp_text_box_batch_name(index);
        let matching = inventory
            .text_boxes()
            .iter()
            .filter(|model| model.name() == name)
            .collect::<Vec<_>>();
        if matching.len() != 1 {
            return Err(format!("ODP text-box batch owner '{name}' is not unique").into());
        }
        let model = matching[0];
        if model.page() != odp_text_box_batch_page(index)
            || model.name() != name
            || model.paragraph_count() != 1
            || model.list_count() != 0
            || !model
                .xml()
                .contains(&odp_text_box_batch_text(index, updated))
        {
            return Err(format!("ODP text-box batch owner '{name}' differs").into());
        }
    }

    let package = litchi_odf_common::core::OwnedPackage::from_bytes(bytes.to_vec())?;
    let package = package.package()?;
    for index in 0..ODS_MEDIA_ENTRY_COUNT {
        let path = odp_media_path(index);
        if package.manifest().get_media_type(&path) != Some("application/octet-stream") {
            return Err(format!("ODP text-box batch manifest entry differs for '{path}'").into());
        }
        if package.get_file(&path)? != odp_media_payload(index) {
            return Err(format!("ODP text-box batch payload differs for '{path}'").into());
        }
    }
    Ok(())
}

fn verify_odp_text_box_batch_raw_members(
    source: &[u8],
    published: &[u8],
    require_manifest: bool,
) -> Result<(), Box<dyn Error>> {
    let identical = litchi_odf_common::package::raw_identical_members(source, published)
        .ok_or("ODP text-box batch raw-member comparison failed")?;
    if identical.contains("content.xml") {
        return Err("ODP text-box batch content.xml remained raw-identical".into());
    }
    if !require_manifest && identical.contains("META-INF/manifest.xml") {
        return Err("ODP scalar text-box staging unexpectedly raw-preserved the manifest".into());
    }
    for path in ArchiveReader::new(source)?.file_names() {
        if path != "content.xml"
            && (require_manifest || path != "META-INF/manifest.xml")
            && !identical.contains(path)
        {
            return Err(format!(
                "ODP text-box batch changed raw member '{path}'; identical={identical:?}"
            )
            .into());
        }
    }
    Ok(())
}

fn verify_semantic_doc(
    bytes: &[u8],
    shape: WriterShape,
    updated: Option<usize>,
) -> Result<(), Box<dyn Error>> {
    use litchi_doc::body_text::{Projection, Snapshot};

    let count = shape.doc_paragraph_count();
    let snapshot = Snapshot::from_bytes(bytes.to_vec())?;
    let projected = snapshot.paragraphs(Projection::All)?;
    if projected.len() != count {
        return Err("semantic DOC paragraph count differs from writer specification".into());
    }
    for (index, paragraph) in projected.iter().enumerate() {
        let expected = if updated == Some(index) {
            updated_writer_text("doc", 0, index, 0)
        } else {
            writer_text("doc", 0, index, 0)
        };
        if paragraph.position() != Position::new(index) || paragraph.text() != expected {
            return Err(
                "semantic DOC projected paragraph differs from writer specification".into(),
            );
        }
    }

    let mut package = litchi_doc::Package::from_reader(Cursor::new(bytes))?;
    let document = package.document()?;
    let paragraphs = document.paragraphs()?;
    if paragraphs.len() != count {
        return Err(
            "semantic DOC document paragraph count differs from writer specification".into(),
        );
    }
    let mut expected_full = String::new();
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let expected = if updated == Some(index) {
            updated_writer_text("doc", 0, index, 0)
        } else {
            writer_text("doc", 0, index, 0)
        };
        if paragraph.text()? != expected {
            return Err("semantic DOC paragraph text differs from writer specification".into());
        }
        expected_full.push_str(&expected);
        expected_full.push('\r');
    }
    if document.text()? != expected_full {
        return Err("semantic DOC full text differs from writer specification".into());
    }
    Ok(())
}

fn run_doc_body_snapshot_list_paragraphs(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    use litchi_doc::body_text::{Projection, Snapshot};

    let shape = writer_shape(corpus)?;
    if shape == WriterShape::PayloadHeavy {
        return Err("payload-heavy DOC corpus is excluded from semantic cases".into());
    }
    let snapshot = Snapshot::from_bytes(corpus.archive.clone())?;
    let count = shape.doc_paragraph_count();
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let paragraphs = snapshot.paragraphs(Projection::All)?;
        let duration = started.elapsed();
        if paragraphs.len() != count {
            return Err("DOC body snapshot paragraph count differs from specification".into());
        }
        for (index, paragraph) in paragraphs.iter().enumerate() {
            if paragraph.position() != Position::new(index)
                || paragraph.text() != writer_text("doc", 0, index, 0)
            {
                return Err("DOC body snapshot paragraph differs from specification".into());
            }
        }
        std::hint::black_box(paragraphs);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(
        Case::DocBodySnapshotListParagraphs,
        corpus,
        elapsed,
        None,
    ))
}

fn run_semantic_doc(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    use litchi_doc::body_text::Snapshot;

    let shape = writer_shape(corpus)?;
    if shape == WriterShape::PayloadHeavy {
        return Err("payload-heavy DOC corpus is excluded from semantic cases".into());
    }
    let selected = shape.doc_paragraph_count() / 2;
    let expected_changed = if case == Case::DocSemanticOneEditSave {
        let source = Snapshot::from_bytes(corpus.archive.clone())?;
        let mut edit = source.edit()?;
        edit.replace_paragraph(
            Position::new(selected),
            &updated_writer_text("doc", 0, selected, 0),
        )?;
        edit.commit()?.snapshot().finish()
    } else {
        corpus.archive.clone()
    };
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        match case {
            Case::DocSemanticOpen => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let mut package = litchi_doc::Package::from_reader(Cursor::new(owned))?;
                let document = package.document()?;
                let duration = started.elapsed();
                verify_semantic_doc(&corpus.archive, shape, None)?;
                std::hint::black_box(document);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::DocSemanticListParagraphs => {
                let mut package =
                    litchi_doc::Package::from_reader(Cursor::new(corpus.archive.as_slice()))?;
                let document = package.document()?;
                let started = Instant::now();
                let paragraphs = document.paragraphs()?;
                let duration = started.elapsed();
                if paragraphs.len() != shape.doc_paragraph_count() {
                    return Err(
                        "semantic DOC paragraph list differs from writer specification".into(),
                    );
                }
                verify_semantic_doc(&corpus.archive, shape, None)?;
                std::hint::black_box(paragraphs);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::DocSemanticOneParagraph => {
                let mut package =
                    litchi_doc::Package::from_reader(Cursor::new(corpus.archive.as_slice()))?;
                let document = package.document()?;
                let started = Instant::now();
                let paragraph = document
                    .paragraphs()?
                    .into_iter()
                    .nth(selected)
                    .ok_or("semantic DOC selected paragraph is missing")?;
                let duration = started.elapsed();
                if paragraph.text()? != writer_text("doc", 0, selected, 0) {
                    return Err(
                        "semantic DOC selected paragraph differs from writer specification".into(),
                    );
                }
                verify_semantic_doc(&corpus.archive, shape, None)?;
                std::hint::black_box(paragraph);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::DocSemanticFullText => {
                let mut package =
                    litchi_doc::Package::from_reader(Cursor::new(corpus.archive.as_slice()))?;
                let document = package.document()?;
                let started = Instant::now();
                let text = document.text()?;
                let duration = started.elapsed();
                verify_semantic_doc(&corpus.archive, shape, None)?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::DocSemanticNoopEditSave | Case::DocSemanticOneEditSave => {
                let source = Snapshot::from_bytes(corpus.archive.clone())?;
                let updated = case == Case::DocSemanticOneEditSave;
                let started = Instant::now();
                let mut edit = source.edit()?;
                if updated {
                    edit.replace_paragraph(
                        Position::new(selected),
                        &updated_writer_text("doc", 0, selected, 0),
                    )?;
                }
                let commit = edit.commit()?;
                let bytes = commit.snapshot().finish();
                let duration = started.elapsed();
                if commit.changed() != updated
                    || commit.patch().is_noop() == updated
                    || commit.patch().changes().len() != usize::from(updated)
                    || bytes != expected_changed
                {
                    return Err("semantic DOC edit/save has unexpected publication state".into());
                }
                let applied = commit.patch().apply(&source)?;
                if applied.bytes() != commit.snapshot().bytes() {
                    return Err("semantic DOC exact patch replay differs from commit".into());
                }
                let restored = commit.patch().inverse().apply(&applied)?;
                if restored.bytes() != source.bytes() {
                    return Err("semantic DOC inverse did not restore exact source bytes".into());
                }
                verify_semantic_doc(&bytes, shape, updated.then_some(selected))?;
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-DOC semantic case passed to DOC runner".into()),
        }
    }
    Ok(result(case, corpus, elapsed, None))
}

fn xls_expected_value(
    shape: WriterShape,
    sheet: usize,
    row: usize,
    column: usize,
) -> Result<f64, Box<dyn Error>> {
    let (_, rows, columns) = shape
        .xls_dimensions()
        .ok_or("payload-heavy XLS corpus has no numeric semantic grid")?;
    Ok((sheet * rows * columns + row * columns + column) as f64)
}

fn xls_comment_value(
    index: usize,
    updated: bool,
) -> Result<litchi_xls::comments::Value, Box<dyn Error>> {
    let state = if updated { "target" } else { "source" };
    let author_state = if updated { "Target" } else { "Source" };
    Ok(litchi_xls::comments::Value::new(
        format!("{author_state} {index:03}"),
        format!("comment {state} {index:03}"),
    )?)
}

fn xls_comment_opaque_stream_name(index: usize) -> String {
    format!("Payload{index:03}")
}

fn build_xls_comments_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let mut workbook_writer = litchi_xls::writer::Writer::new();
    let comments = workbook_writer.add_worksheet("Comments")?;
    for index in 0..XLS_COMMENTS_SOURCE_COUNT {
        let value = xls_comment_value(index, false)?;
        workbook_writer.add_comment(
            comments,
            u32::try_from(index)?,
            1,
            value.author(),
            value.text(),
        )?;
    }
    let untouched = workbook_writer.add_worksheet("Untouched")?;
    workbook_writer.write_number(untouched, 20, 4, 42.0)?;
    workbook_writer.add_comment(untouched, 4, 3, "Sentinel", "untouched sentinel")?;
    let mut workbook_package = Cursor::new(Vec::new());
    workbook_writer.write_to(&mut workbook_package)?;
    let mut parsed = OleFile::open(Cursor::new(workbook_package.into_inner()))?;
    let workbook_stream = parsed.open_stream(&["Workbook"])?;

    let mut writer = OleWriter::new();
    writer.create_stream_owned(&["Workbook"], workbook_stream)?;
    writer.create_storage(&["OpaquePayloads"])?;
    for index in 0..XLS_COMMENTS_OPAQUE_STREAM_COUNT {
        let name = xls_comment_opaque_stream_name(index);
        writer.create_stream_owned(
            &["OpaquePayloads", name.as_str()],
            payload_bytes(
                PayloadKind::Incompressible,
                10_000 + index,
                XLS_COMMENTS_OPAQUE_STREAM_BYTES,
            ),
        )?;
    }
    writer.create_stream(
        &["OpaqueMetadata"],
        b"litchi-xls-comments-opaque-metadata-v1",
    )?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let archive = output.into_inner();

    verify_xls_comments_static_guards(&archive)?;
    let mut parsed = OleFile::open(Cursor::new(archive.as_slice()))?;
    let archive_member_count = parsed.list_streams().len();
    let target_payload = parsed.open_stream(&["Workbook"])?;
    let uncompressed_payload_bytes = target_payload
        .len()
        .checked_add(XLS_COMMENTS_OPAQUE_STREAM_COUNT * XLS_COMMENTS_OPAQUE_STREAM_BYTES)
        .and_then(|value| value.checked_add(b"litchi-xls-comments-opaque-metadata-v1".len()))
        .ok_or("XLS comments corpus payload size overflow")?;
    if archive_member_count != XLS_COMMENTS_OPAQUE_STREAM_COUNT + 2 {
        return Err("XLS comments corpus stream inventory differs from specification".into());
    }

    Ok(Corpus {
        manifest: CorpusManifest {
            name: "xls-comments-opaque-heavy".to_string(),
            generator: XLS_COMMENTS_EDIT_CORPUS_GENERATOR,
            package_format: "XLS/CFB",
            shape: "256-comments-opaque-heavy",
            payload_kind: PayloadKind::Incompressible.name(),
            compression: "none",
            entry_count: XLS_COMMENTS_SOURCE_COUNT + 1,
            archive_member_count,
            entry_bytes: XLS_COMMENTS_OPAQUE_STREAM_BYTES,
            uncompressed_payload_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "Workbook".to_string(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "Workbook".to_string(),
        target_payload,
        xlsx: None,
    })
}

fn xls_visibility_opaque_stream_name(index: usize) -> String {
    format!("Payload{index:03}")
}

fn build_xls_visibility_edit_corpus() -> Result<Corpus, Box<dyn Error>> {
    let mut workbook_writer = litchi_xls::writer::Writer::new();
    for index in 0..XLS_VISIBILITY_SHEET_COUNT {
        let worksheet = workbook_writer.add_worksheet(&format!("Visibility{index:02}"))?;
        workbook_writer.write_number(worksheet, 0, 0, index as f64)?;
    }
    let mut workbook_package = Cursor::new(Vec::new());
    workbook_writer.write_to(&mut workbook_package)?;
    let mut parsed = OleFile::open(Cursor::new(workbook_package.into_inner()))?;
    let workbook_stream = parsed.open_stream(&["Workbook"])?;

    let mut writer = OleWriter::new();
    writer.create_stream_owned(&["Workbook"], workbook_stream)?;
    writer.create_storage(&["OpaquePayloads"])?;
    for index in 0..XLS_VISIBILITY_OPAQUE_STREAM_COUNT {
        let name = xls_visibility_opaque_stream_name(index);
        writer.create_stream_owned(
            &["OpaquePayloads", name.as_str()],
            payload_bytes(
                PayloadKind::Incompressible,
                20_000 + index,
                XLS_VISIBILITY_OPAQUE_STREAM_BYTES,
            ),
        )?;
    }
    writer.create_stream(
        &["OpaqueMetadata"],
        b"litchi-xls-visibility-opaque-metadata-v1",
    )?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    let archive = output.into_inner();

    verify_xls_visibility_static_guards(&archive)?;
    let mut parsed = OleFile::open(Cursor::new(archive.as_slice()))?;
    let archive_member_count = parsed.list_streams().len();
    let target_payload = parsed.open_stream(&["Workbook"])?;
    let opaque_bytes = XLS_VISIBILITY_OPAQUE_STREAM_COUNT
        .checked_mul(XLS_VISIBILITY_OPAQUE_STREAM_BYTES)
        .and_then(|bytes| bytes.checked_add(b"litchi-xls-visibility-opaque-metadata-v1".len()))
        .ok_or("XLS visibility opaque payload size overflow")?;
    let uncompressed_payload_bytes = target_payload
        .len()
        .checked_add(opaque_bytes)
        .ok_or("XLS visibility corpus payload size overflow")?;
    if archive_member_count != XLS_VISIBILITY_OPAQUE_STREAM_COUNT + 2 {
        return Err("XLS visibility corpus stream inventory differs from specification".into());
    }

    Ok(Corpus {
        manifest: CorpusManifest {
            name: "xls-visibility-opaque".to_owned(),
            generator: XLS_VISIBILITY_CORPUS_GENERATOR,
            package_format: "XLS/CFB",
            shape: "66-worksheets-opaque-heavy",
            payload_kind: PayloadKind::Incompressible.name(),
            compression: "none",
            entry_count: XLS_VISIBILITY_SHEET_COUNT,
            archive_member_count,
            entry_bytes: XLS_VISIBILITY_OPAQUE_STREAM_BYTES,
            uncompressed_payload_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "Workbook".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "Workbook".to_owned(),
        target_payload,
        xlsx: None,
    })
}

fn xls_visibility_case_parameters(case: Case) -> Result<(bool, usize), Box<dyn Error>> {
    match case {
        Case::XlsVisibilityEagerEditSave => Ok((false, 1)),
        Case::XlsVisibilitySourceBackedEditSave => Ok((true, 1)),
        Case::XlsVisibilityEagerBatchEditSave => Ok((false, XLS_VISIBILITY_BATCH_COUNT)),
        Case::XlsVisibilitySourceBackedBatchEditSave => Ok((true, XLS_VISIBILITY_BATCH_COUNT)),
        _ => Err("non-XLS-visibility case passed to XLS visibility runner".into()),
    }
}

fn stage_xls_visibility_updates(
    transaction: &mut litchi_xls::sheet_visibility::Transaction,
    update_count: usize,
) -> Result<(), Box<dyn Error>> {
    if update_count == 1 {
        transaction.hide(1usize.into())?;
        return Ok(());
    }
    if update_count != XLS_VISIBILITY_BATCH_COUNT {
        return Err("XLS visibility benchmark update count is outside its fixed closure".into());
    }
    for position in 1..=XLS_VISIBILITY_BATCH_COUNT {
        transaction.hide(position.into())?;
    }
    Ok(())
}

enum XlsVisibilityPublication {
    Eager {
        commit: litchi_xls::sheet_visibility::Commit,
        source_bytes: u64,
        source_workbook_bytes: u64,
    },
    SourceBacked(litchi_xls::sheet_visibility::SourceBackedCommit),
}

impl XlsVisibilityPublication {
    fn write_to<W: Write>(
        &self,
        writer: &mut W,
    ) -> Result<Option<litchi_cfb::PublishReport>, Box<dyn Error>> {
        match self {
            Self::Eager { commit, .. } => {
                for chunk in commit.snapshot().bytes().chunks(64 * 1024) {
                    writer.write_all(chunk)?;
                }
                writer.flush()?;
                Ok(None)
            },
            Self::SourceBacked(commit) => Ok(Some(commit.write_to(writer)?)),
        }
    }

    fn evidence(
        &self,
        source_backed: bool,
        update_count: usize,
        semantic_staging_plan: Duration,
        publication: Duration,
    ) -> Result<XlsVisibilityIterationEvidence, Box<dyn Error>> {
        match self {
            Self::Eager {
                commit,
                source_bytes,
                source_workbook_bytes,
            } => {
                let diagnostics = commit.diagnostics();
                Ok(XlsVisibilityIterationEvidence {
                    source_backed,
                    update_count,
                    semantic_staging_plan_ns: elapsed_ns(semantic_staging_plan)?,
                    publication_ns: elapsed_ns(publication)?,
                    changed_worksheets: diagnostics.changed_worksheets(),
                    touched_streams: diagnostics.touched_streams(),
                    source_bytes: *source_bytes,
                    source_workbook_bytes: *source_workbook_bytes,
                    target_workbook_bytes: u64::try_from(
                        commit.snapshot().workbook_stream().len(),
                    )?,
                    splice_count: None,
                    replacement_bytes: None,
                    changed_spans: None,
                    source_fingerprint: None,
                    target_fingerprint: None,
                })
            },
            Self::SourceBacked(commit) => {
                let diagnostics = commit.diagnostics();
                Ok(XlsVisibilityIterationEvidence {
                    source_backed,
                    update_count,
                    semantic_staging_plan_ns: elapsed_ns(semantic_staging_plan)?,
                    publication_ns: elapsed_ns(publication)?,
                    changed_worksheets: diagnostics.changed_worksheets(),
                    touched_streams: diagnostics.touched_streams(),
                    source_bytes: diagnostics.source_bytes(),
                    source_workbook_bytes: diagnostics.source_workbook_bytes(),
                    target_workbook_bytes: diagnostics.target_workbook_bytes(),
                    splice_count: Some(diagnostics.splice_count()),
                    replacement_bytes: Some(diagnostics.replacement_bytes()),
                    changed_spans: Some(diagnostics.changed_spans()),
                    source_fingerprint: Some(fingerprint_hex(
                        diagnostics.source_fingerprint().as_bytes(),
                    )),
                    target_fingerprint: Some(fingerprint_hex(
                        diagnostics.target_fingerprint().as_bytes(),
                    )),
                })
            },
        }
    }
}

fn prepare_xls_visibility_publication(
    source: litchi_xls::sheet_visibility::Snapshot,
    source_backed: bool,
    update_count: usize,
) -> Result<XlsVisibilityPublication, Box<dyn Error>> {
    let source_bytes = u64::try_from(source.bytes().len())?;
    let source_workbook_bytes = u64::try_from(source.workbook_stream().len())?;
    let mut transaction = source.transaction();
    stage_xls_visibility_updates(&mut transaction, update_count)?;
    if source_backed {
        Ok(XlsVisibilityPublication::SourceBacked(
            transaction.commit_source_backed()?,
        ))
    } else {
        Ok(XlsVisibilityPublication::Eager {
            commit: transaction.commit()?,
            source_bytes,
            source_workbook_bytes,
        })
    }
}

fn read_xls_visibility_source(
    source: &InstrumentedSource,
    output: &mut [u8],
) -> Result<(), Box<dyn Error>> {
    let mut offset = 0_u64;
    for chunk in output.chunks_mut(64 * 1024) {
        source.read_exact_at(offset, chunk)?;
        offset = offset
            .checked_add(u64::try_from(chunk.len())?)
            .ok_or("XLS visibility source offset overflow")?;
    }
    Ok(())
}

fn xls_visibility_expected_updates(update_count: usize, position: usize) -> bool {
    (update_count == XLS_VISIBILITY_BATCH_COUNT
        && (1..=XLS_VISIBILITY_BATCH_COUNT).contains(&position))
        || (update_count == 1 && position == 1)
}

fn xls_visibility_bound_sheet_offsets(workbook: &[u8]) -> Result<Vec<usize>, Box<dyn Error>> {
    let mut offsets = Vec::new();
    let mut offset = 0_usize;
    while offset < workbook.len() {
        let header_end = offset
            .checked_add(4)
            .ok_or("XLS visibility BIFF header offset overflow")?;
        let header = workbook
            .get(offset..header_end)
            .ok_or("XLS visibility BIFF record has a truncated header")?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let payload_len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let end = header_end
            .checked_add(payload_len)
            .ok_or("XLS visibility BIFF record length overflow")?;
        if workbook.get(header_end..end).is_none() {
            return Err("XLS visibility BIFF record has a truncated payload".into());
        }
        if kind == 0x0085 {
            let state_offset = offset
                .checked_add(8)
                .ok_or("XLS visibility hsState offset overflow")?;
            if state_offset >= end {
                return Err("XLS visibility BoundSheet8 has no hsState field".into());
            }
            offsets.push(state_offset);
        }
        offset = end;
    }
    Ok(offsets)
}

fn verify_xls_visibility_output(
    corpus: &Corpus,
    output: &[u8],
    update_count: usize,
) -> Result<(), Box<dyn Error>> {
    use litchi_xls::SheetVisibility;

    if sha256_hex(&corpus.archive) != corpus.manifest.archive_sha256 {
        return Err("XLS visibility source digest differs from its manifest".into());
    }
    verify_xls_untouched_streams(&corpus.archive, output, true)?;
    let snapshot = litchi_xls::sheet_visibility::Snapshot::from_bytes(output.to_vec())?;
    if snapshot.worksheet_count() != XLS_VISIBILITY_SHEET_COUNT {
        return Err("XLS visibility worksheet inventory differs from source".into());
    }
    for position in 0..XLS_VISIBILITY_SHEET_COUNT {
        let worksheet = snapshot
            .worksheet(position.into())?
            .ok_or("XLS visibility worksheet owner disappeared")?;
        let expected = if xls_visibility_expected_updates(update_count, position) {
            SheetVisibility::Hidden
        } else {
            SheetVisibility::Visible
        };
        if worksheet.visibility() != expected {
            return Err(format!(
                "XLS visibility semantic readback differs at position {position}: observed {:?}, expected {:?}",
                worksheet.visibility(),
                expected,
            )
            .into());
        }
    }
    let source_workbook =
        litchi_xls::sheet_visibility::Snapshot::from_bytes(corpus.archive.clone())?
            .workbook_stream()
            .to_vec();
    let source_offsets = xls_visibility_bound_sheet_offsets(&source_workbook)?;
    let target_offsets = xls_visibility_bound_sheet_offsets(snapshot.workbook_stream())?;
    if source_offsets != target_offsets || source_offsets.len() != XLS_VISIBILITY_SHEET_COUNT {
        return Err("XLS visibility BoundSheet8 offset inventory differs from source".into());
    }
    let changed_offsets = source_workbook
        .iter()
        .zip(snapshot.workbook_stream())
        .enumerate()
        .filter_map(|(offset, (before, after))| (before != after).then_some(offset))
        .collect::<Vec<_>>();
    let expected_offsets = source_offsets
        .iter()
        .enumerate()
        .filter_map(|(position, offset)| {
            xls_visibility_expected_updates(update_count, position).then_some(*offset)
        })
        .collect::<Vec<_>>();
    if changed_offsets != expected_offsets {
        return Err(format!(
            "XLS visibility changed offsets differ: observed {changed_offsets:?}, expected {expected_offsets:?}"
        )
        .into());
    }
    Ok(())
}

fn verify_xls_visibility_static_guards(source: &[u8]) -> Result<(), Box<dyn Error>> {
    use litchi_xls::SheetVisibility;

    let snapshot = litchi_xls::sheet_visibility::Snapshot::from_bytes(source.to_vec())?;
    if snapshot.worksheet_count() != XLS_VISIBILITY_SHEET_COUNT {
        return Err("XLS visibility source worksheet count differs from specification".into());
    }
    let noop = snapshot.transaction().commit_source_backed()?;
    if !noop.is_noop() || noop.diagnostics().changed_worksheets() != 0 {
        return Err("XLS visibility source-backed no-op was not exact".into());
    }
    let mut noop_bytes = Vec::new();
    noop.write_to(&mut noop_bytes)?;
    if noop_bytes != source {
        return Err("XLS visibility source-backed no-op changed source bytes".into());
    }

    let mut capped = snapshot.transaction();
    stage_xls_visibility_updates(&mut capped, XLS_VISIBILITY_BATCH_COUNT)?;
    let capped = capped.commit_source_backed()?;
    if capped.diagnostics().changed_worksheets() != XLS_VISIBILITY_BATCH_COUNT
        || capped.diagnostics().touched_streams() != 1
        || capped.diagnostics().splice_count() != XLS_VISIBILITY_BATCH_COUNT
        || capped.diagnostics().replacement_bytes() != u64::try_from(XLS_VISIBILITY_BATCH_COUNT)?
    {
        return Err("XLS visibility exact-cap plan has unexpected diagnostics".into());
    }
    let mut capped_bytes = Vec::new();
    capped.write_to(&mut capped_bytes)?;
    verify_xls_visibility_output(
        &Corpus {
            manifest: CorpusManifest {
                name: "guard".to_owned(),
                generator: XLS_VISIBILITY_CORPUS_GENERATOR,
                package_format: "XLS/CFB",
                shape: "guard",
                payload_kind: "guard",
                compression: "none",
                entry_count: XLS_VISIBILITY_SHEET_COUNT,
                archive_member_count: 0,
                entry_bytes: 0,
                uncompressed_payload_bytes: 0,
                archive_bytes: source.len(),
                archive_sha256: sha256_hex(source),
                target_entry: "Workbook".to_owned(),
                target_payload_bytes: 0,
                target_payload_sha256: String::new(),
                rtf_variant: None,
                xlsx: None,
            },
            archive: source.to_vec(),
            target_name: "Workbook".to_owned(),
            target_payload: Vec::new(),
            xlsx: None,
        },
        &capped_bytes,
        XLS_VISIBILITY_BATCH_COUNT,
    )?;

    let mut over_cap = snapshot.transaction();
    if over_cap
        .set_visibility_batch(
            (1..=XLS_VISIBILITY_BATCH_COUNT + 1)
                .map(|position| (position.into(), SheetVisibility::Hidden)),
        )
        .is_ok()
    {
        return Err("XLS visibility cap-plus-one staging unexpectedly succeeded".into());
    }
    if !over_cap.commit_source_backed()?.is_noop() {
        return Err("XLS visibility cap-plus-one refusal mutated its source".into());
    }

    let mut protected_writer = litchi_xls::writer::Writer::new();
    let sheet = protected_writer.add_worksheet("Protected")?;
    protected_writer.write_number(sheet, 0, 0, 1.0)?;
    protected_writer.protect_sheet(sheet, Some("password"), true, false)?;
    let mut protected_bytes = Cursor::new(Vec::new());
    protected_writer.write_to(&mut protected_bytes)?;
    if litchi_xls::sheet_visibility::Snapshot::from_bytes(protected_bytes.into_inner()).is_ok() {
        return Err("XLS visibility protected source was accepted".into());
    }
    Ok(())
}

fn verify_xls_visibility_prepared(
    source: &litchi_xls::sheet_visibility::Snapshot,
    prepared: &XlsVisibilityPublication,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    match prepared {
        XlsVisibilityPublication::Eager { commit, .. } => {
            let applied = commit.patch().apply(source)?;
            if applied.bytes() != commit.snapshot().bytes()
                || commit.patch().apply(commit.snapshot()).is_ok()
            {
                return Err("XLS visibility eager patch did not enforce exact source".into());
            }
            let restored = commit.patch().inverse().apply(&applied)?;
            if restored.bytes() != source.bytes() {
                return Err("XLS visibility eager inverse did not restore exact source".into());
            }
        },
        XlsVisibilityPublication::SourceBacked(commit) => {
            let diagnostics = commit.diagnostics();
            let source_digest: [u8; 32] = Sha256::digest(source.bytes()).into();
            let target_digest: [u8; 32] = Sha256::digest(output).into();
            if diagnostics.source_fingerprint().as_bytes() != &source_digest
                || diagnostics.target_fingerprint().as_bytes() != &target_digest
            {
                return Err(
                    "XLS visibility overlay fingerprints differ from exact artifacts".into(),
                );
            }
        },
    }
    Ok(())
}

fn run_xls_visibility_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let (source_backed, update_count) = xls_visibility_case_parameters(case)?;
    if corpus.manifest.generator != XLS_VISIBILITY_CORPUS_GENERATOR {
        return Err("XLS visibility edit case requires its fixed opaque corpus".into());
    }

    let expected_source =
        litchi_xls::sheet_visibility::Snapshot::from_bytes(corpus.archive.clone())?;
    let expected_prepared =
        prepare_xls_visibility_publication(expected_source.clone(), source_backed, update_count)?;
    let mut expected = Vec::new();
    let expected_report = expected_prepared.write_to(&mut expected)?;
    if expected == corpus.archive {
        return Err("XLS visibility expected publication did not change source bytes".into());
    }
    verify_xls_visibility_output(corpus, &expected, update_count)?;
    verify_xls_visibility_prepared(&expected_source, &expected_prepared, &expected)?;
    if let Some(report) = expected_report
        && (report.bytes() != u64::try_from(expected.len())? || report.changed_spans() == 0)
    {
        return Err("XLS visibility expected overlay report is incomplete".into());
    }
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = InstrumentedSource::new(corpus.archive.clone(), Vec::new());
        let mut source_bytes = vec![0_u8; corpus.archive.len()];
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;

        read_xls_visibility_source(&source, &mut source_bytes)?;
        let snapshot = litchi_xls::sheet_visibility::Snapshot::from_bytes(source_bytes)?;
        let plan_started = Instant::now();
        let prepared = prepare_xls_visibility_publication(snapshot, source_backed, update_count)?;
        let semantic_staging_plan = plan_started.elapsed();

        let publication_started = Instant::now();
        let report = prepared.write_to(&mut sink)?;
        let publication = publication_started.elapsed();
        let duration = semantic_staging_plan
            .checked_add(publication)
            .ok_or("XLS visibility total duration overflow")?;
        let metrics = source.snapshot();
        let evidence = prepared.evidence(
            source_backed,
            update_count,
            semantic_staging_plan,
            publication,
        )?;

        if metrics.read_calls == 0
            || metrics.read_bytes != u64::try_from(corpus.archive.len())?
            || evidence.changed_worksheets != update_count
            || evidence.touched_streams != 1
            || evidence.source_workbook_bytes != evidence.target_workbook_bytes
            || sink.bytes != expected
        {
            return Err(
                "XLS visibility iteration has unexpected source/publication evidence".into(),
            );
        }
        if source_backed {
            let expected_splices = update_count;
            let expected_replacement_bytes = u64::try_from(update_count)?;
            if evidence.splice_count != Some(expected_splices)
                || evidence.replacement_bytes != Some(expected_replacement_bytes)
            {
                return Err(
                    "XLS visibility source-backed splice diagnostics disagree with changed BoundSheet8 ranges".into(),
                );
            }
            let report = report.ok_or("XLS visibility source-backed publication has no report")?;
            if evidence.changed_spans != Some(report.changed_spans())
                || report.bytes() != sink.summary().accepted_bytes
                || evidence.source_fingerprint
                    != Some(fingerprint_hex(report.source_fingerprint().as_bytes()))
                || evidence.target_fingerprint
                    != Some(fingerprint_hex(report.target_fingerprint().as_bytes()))
            {
                return Err("XLS visibility overlay diagnostics disagree with publication".into());
            }
        } else if report.is_some() || evidence.changed_spans.is_some() {
            return Err("XLS eager visibility publication reported overlay-only evidence".into());
        }
        verify_xls_visibility_output(corpus, &sink.bytes, update_count)?;
        verify_xls_visibility_prepared(
            &litchi_xls::sheet_visibility::Snapshot::from_bytes(corpus.archive.clone())?,
            &prepared,
            &sink.bytes,
        )?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("XLS visibility output digest differs from expected output".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_xls_visibility(metrics, evidence)?;
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLS visibility measured output hashes are unstable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(deterministic_sink_summary(
            &sink_summaries,
            "XLS visibility publication",
        )?),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn xls_comments_case_parameters(case: Case) -> Result<(bool, usize), Box<dyn Error>> {
    match case {
        Case::XlsCommentsEagerEditSave => Ok((false, 1)),
        Case::XlsCommentsSourceBackedEditSave => Ok((true, 1)),
        Case::XlsCommentsEagerBatchEditSave => Ok((false, XLS_COMMENTS_BATCH_COUNT)),
        Case::XlsCommentsSourceBackedBatchEditSave => Ok((true, XLS_COMMENTS_BATCH_COUNT)),
        _ => Err("non-XLS-comment case passed to XLS comment runner".into()),
    }
}

fn stage_xls_comment_updates(
    edit: &mut litchi_xls::comments::Edit,
    update_count: usize,
) -> Result<(), Box<dyn Error>> {
    use litchi_xls::cell_values::{Reference, Selector};
    use litchi_xls::comments::Update;

    if update_count == 1 {
        let index = XLS_COMMENTS_SOURCE_COUNT / 2;
        edit.replace(
            Selector::Position(0),
            Reference::new(u32::try_from(index)?, 1)?,
            xls_comment_value(index, true)?,
        )?;
        return Ok(());
    }
    if update_count != XLS_COMMENTS_BATCH_COUNT {
        return Err("XLS comment benchmark update count is outside its fixed closure".into());
    }
    let updates = (0..update_count)
        .map(|index| {
            Ok(Update::new(
                Reference::new(u32::try_from(index)?, 1)?,
                xls_comment_value(index, true)?,
            ))
        })
        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
    edit.replace_many(Selector::Position(0), updates)?;
    Ok(())
}

enum XlsCommentsPublication {
    Eager {
        commit: litchi_xls::comments::Commit,
        source_bytes: u64,
        source_workbook_bytes: u64,
    },
    SourceBacked(litchi_xls::comments::SourceBackedCommit),
}

impl XlsCommentsPublication {
    fn write_to<W: Write>(
        &self,
        writer: &mut W,
    ) -> Result<Option<litchi_cfb::PublishReport>, Box<dyn Error>> {
        match self {
            Self::Eager { commit, .. } => {
                for chunk in commit.snapshot().bytes().chunks(64 * 1024) {
                    writer.write_all(chunk)?;
                }
                writer.flush()?;
                Ok(None)
            },
            Self::SourceBacked(commit) => Ok(Some(commit.write_to(writer)?)),
        }
    }

    fn evidence(
        &self,
        source_backed: bool,
        update_count: usize,
        semantic_staging_plan: Duration,
        publication: Duration,
    ) -> Result<XlsCommentsIterationEvidence, Box<dyn Error>> {
        match self {
            Self::Eager {
                commit,
                source_bytes,
                source_workbook_bytes,
            } => {
                let diagnostics = commit.diagnostics();
                Ok(XlsCommentsIterationEvidence {
                    source_backed,
                    update_count,
                    semantic_staging_plan_ns: elapsed_ns(semantic_staging_plan)?,
                    publication_ns: elapsed_ns(publication)?,
                    changed_comments: diagnostics.changed_comments(),
                    touched_streams: diagnostics.touched_streams(),
                    source_bytes: *source_bytes,
                    source_workbook_bytes: *source_workbook_bytes,
                    target_workbook_bytes: u64::try_from(
                        commit.snapshot().workbook_stream().len(),
                    )?,
                    splice_count: None,
                    replacement_bytes: None,
                    changed_spans: None,
                    source_fingerprint: None,
                    target_fingerprint: None,
                })
            },
            Self::SourceBacked(commit) => {
                let diagnostics = commit.diagnostics();
                Ok(XlsCommentsIterationEvidence {
                    source_backed,
                    update_count,
                    semantic_staging_plan_ns: elapsed_ns(semantic_staging_plan)?,
                    publication_ns: elapsed_ns(publication)?,
                    changed_comments: diagnostics.changed_comments(),
                    touched_streams: diagnostics.touched_streams(),
                    source_bytes: diagnostics.source_bytes(),
                    source_workbook_bytes: diagnostics.source_workbook_bytes(),
                    target_workbook_bytes: diagnostics.target_workbook_bytes(),
                    splice_count: Some(diagnostics.splice_count()),
                    replacement_bytes: Some(diagnostics.replacement_bytes()),
                    changed_spans: Some(diagnostics.changed_spans()),
                    source_fingerprint: Some(fingerprint_hex(
                        diagnostics.source_fingerprint().as_bytes(),
                    )),
                    target_fingerprint: Some(fingerprint_hex(
                        diagnostics.target_fingerprint().as_bytes(),
                    )),
                })
            },
        }
    }
}

fn prepare_xls_comments_publication(
    source: litchi_xls::comments::Snapshot,
    source_backed: bool,
    update_count: usize,
) -> Result<XlsCommentsPublication, Box<dyn Error>> {
    let source_bytes = u64::try_from(source.bytes().len())?;
    let source_workbook_bytes = u64::try_from(source.workbook_stream().len())?;
    let mut edit = source.edit();
    stage_xls_comment_updates(&mut edit, update_count)?;
    if source_backed {
        Ok(XlsCommentsPublication::SourceBacked(
            edit.commit_source_backed()?,
        ))
    } else {
        Ok(XlsCommentsPublication::Eager {
            commit: edit.commit()?,
            source_bytes,
            source_workbook_bytes,
        })
    }
}

fn read_xls_comments_source(
    source: &InstrumentedSource,
    output: &mut [u8],
) -> Result<(), Box<dyn Error>> {
    let mut offset = 0_u64;
    for chunk in output.chunks_mut(64 * 1024) {
        source.read_exact_at(offset, chunk)?;
        offset = offset
            .checked_add(u64::try_from(chunk.len())?)
            .ok_or("XLS comments source offset overflow")?;
    }
    Ok(())
}

fn verify_xls_untouched_streams(
    source: &[u8],
    candidate: &[u8],
    require_equal_workbook_length: bool,
) -> Result<(), Box<dyn Error>> {
    let mut source_ole = OleFile::open(Cursor::new(source))?;
    let mut candidate_ole = OleFile::open(Cursor::new(candidate))?;
    let mut source_paths = source_ole.list_streams();
    let mut candidate_paths = candidate_ole.list_streams();
    source_paths.sort();
    candidate_paths.sort();
    if source_paths != candidate_paths
        || !source_ole.directory_exists(&["OpaquePayloads"])
        || !candidate_ole.directory_exists(&["OpaquePayloads"])
    {
        return Err("XLS comment publication changed the CFB stream/storage inventory".into());
    }
    for path in source_paths {
        let borrowed = path.iter().map(String::as_str).collect::<Vec<_>>();
        let source_stream = source_ole.open_stream(&borrowed)?;
        let candidate_stream = candidate_ole.open_stream(&borrowed)?;
        let workbook = path.len() == 1
            && path
                .first()
                .is_some_and(|name| name == "Workbook" || name == "Book");
        if workbook {
            if source_stream == candidate_stream {
                return Err("XLS comment publication left the Workbook stream unchanged".into());
            }
            if require_equal_workbook_length && source_stream.len() != candidate_stream.len() {
                return Err(format!(
                    "XLS comment publication changed Workbook length from {} to {}",
                    source_stream.len(),
                    candidate_stream.len()
                )
                .into());
            }
        } else if source_stream != candidate_stream {
            return Err("XLS comment publication changed an untouched opaque stream".into());
        }
    }
    Ok(())
}

fn verify_xls_comments_output(
    corpus: &Corpus,
    output: &[u8],
    update_count: usize,
) -> Result<(), Box<dyn Error>> {
    use litchi_core::sheet::Cell as _;
    use litchi_xls::cell_values::{Reference, Selector};

    if sha256_hex(&corpus.archive) != corpus.manifest.archive_sha256 {
        return Err("XLS comment source digest differs from its manifest".into());
    }
    verify_xls_untouched_streams(&corpus.archive, output, false)?;
    let snapshot = litchi_xls::comments::Snapshot::from_bytes(output.to_vec())?;
    if snapshot.worksheet_count() != 2 {
        return Err("XLS comment output worksheet inventory differs from source".into());
    }
    let comments = snapshot
        .worksheet(Selector::Name("Comments"))?
        .ok_or("XLS comment output lost its selected worksheet")?;
    if comments.comments().len() != XLS_COMMENTS_SOURCE_COUNT {
        return Err("XLS comment output comment inventory differs from source".into());
    }
    for index in 0..XLS_COMMENTS_SOURCE_COUNT {
        let should_update = update_count == XLS_COMMENTS_BATCH_COUNT
            || (update_count == 1 && index == XLS_COMMENTS_SOURCE_COUNT / 2);
        let comment = comments
            .comment(Reference::new(u32::try_from(index)?, 1)?)?
            .ok_or("XLS comment output lost a selected NOTE owner")?;
        if litchi_xls::comments::Value::from_comment(comment)
            != xls_comment_value(index, should_update)?
        {
            return Err("XLS comment output semantic value differs from expectation".into());
        }
    }
    let untouched = snapshot
        .worksheet(Selector::Name("Untouched"))?
        .ok_or("XLS comment output lost its untouched worksheet")?;
    let sentinel = untouched
        .comment(Reference::new(4, 3)?)?
        .ok_or("XLS comment output lost its untouched sentinel")?;
    if sentinel.author() != "Sentinel" || sentinel.text() != "untouched sentinel" {
        return Err("XLS comment output changed the untouched sentinel".into());
    }
    let workbook = litchi_xls::Workbook::new(Cursor::new(output))?;
    let untouched_metadata = workbook
        .sheets()
        .iter()
        .find(|sheet| sheet.name() == "Untouched")
        .ok_or("XLS comment output lost its untouched worksheet metadata")?;
    let untouched_cells = workbook.xls_worksheet(
        untouched_metadata
            .parsed_worksheet_index()
            .ok_or("XLS comment untouched tab is not a worksheet")?,
    )?;
    let numeric = untouched_cells
        .get_cell(20, 4)
        .ok_or("XLS comment output lost its untouched numeric cell")?
        .value()
        .as_float()
        .ok_or("XLS comment untouched numeric cell is no longer numeric")?;
    if numeric.to_bits() != 42.0_f64.to_bits() {
        return Err("XLS comment output changed its untouched numeric cell".into());
    }
    Ok(())
}

fn verify_xls_comments_static_guards(source: &[u8]) -> Result<(), Box<dyn Error>> {
    use litchi_xls::cell_values::{Reference, Selector};
    use litchi_xls::comments::Value;

    let snapshot = litchi_xls::comments::Snapshot::from_bytes(source.to_vec())?;
    let mut rejected = snapshot.edit();
    rejected.replace(
        Selector::Position(0),
        Reference::new(0, 1)?,
        Value::new("Target 000", "comment target 000!")?,
    )?;
    if rejected.commit_source_backed().is_ok() {
        return Err("XLS source-backed comment edit accepted a length change".into());
    }

    let mut rejected_width = snapshot.edit();
    rejected_width.replace(
        Selector::Position(0),
        Reference::new(0, 1)?,
        Value::new("Target 000", "作者作者作者作者作")?,
    )?;
    if rejected_width.commit_source_backed().is_ok() {
        return Err("XLS source-backed comment edit accepted an encoding-width change".into());
    }

    let mut fallback = snapshot.edit();
    fallback.replace(
        Selector::Position(0),
        Reference::new(0, 1)?,
        Value::new("Target 000", "comment target 000!")?,
    )?;
    let fallback = fallback.commit()?;
    verify_xls_untouched_streams(source, fallback.snapshot().bytes(), false)?;
    let changed = fallback
        .snapshot()
        .worksheet(Selector::Position(0))?
        .ok_or("XLS eager fallback lost its worksheet")?
        .comment(Reference::new(0, 1)?)?
        .ok_or("XLS eager fallback lost its comment")?;
    if changed.text() != "comment target 000!" {
        return Err("XLS eager fallback did not publish the explicit length change".into());
    }

    let mut protected_writer = litchi_xls::writer::Writer::new();
    let sheet = protected_writer.add_worksheet("Protected")?;
    protected_writer.add_comment(sheet, 0, 0, "Source", "source")?;
    protected_writer.protect_sheet(sheet, Some("password"), true, false)?;
    let mut protected_bytes = Cursor::new(Vec::new());
    protected_writer.write_to(&mut protected_bytes)?;
    let protected = litchi_xls::comments::Snapshot::from_bytes(protected_bytes.into_inner())?;
    let mut edit = protected.edit();
    if edit
        .replace(
            Selector::Position(0),
            Reference::new(0, 0)?,
            Value::new("Target", "target")?,
        )
        .is_ok()
    {
        return Err("XLS comment edit accepted a protected worksheet".into());
    }
    Ok(())
}

fn verify_xls_comments_prepared(
    source: &litchi_xls::comments::Snapshot,
    prepared: &XlsCommentsPublication,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    match prepared {
        XlsCommentsPublication::Eager { commit, .. } => {
            let applied = commit.patch().apply(source)?;
            if applied.bytes() != commit.snapshot().bytes()
                || commit.patch().apply(commit.snapshot()).is_ok()
            {
                return Err("XLS comment eager patch did not enforce its exact source".into());
            }
            let restored = commit.patch().inverse().apply(&applied)?;
            if restored.bytes() != source.bytes() {
                return Err("XLS comment eager inverse did not restore exact source bytes".into());
            }
        },
        XlsCommentsPublication::SourceBacked(commit) => {
            let diagnostics = commit.diagnostics();
            let source_digest: [u8; 32] = Sha256::digest(source.bytes()).into();
            let target_digest: [u8; 32] = Sha256::digest(output).into();
            if diagnostics.source_fingerprint().as_bytes() != &source_digest
                || diagnostics.target_fingerprint().as_bytes() != &target_digest
            {
                return Err("XLS comment overlay fingerprints differ from exact artifacts".into());
            }
        },
    }
    Ok(())
}

fn run_xls_comments_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let (source_backed, update_count) = xls_comments_case_parameters(case)?;
    if corpus.manifest.generator != XLS_COMMENTS_EDIT_CORPUS_GENERATOR {
        return Err("XLS comment edit case requires its fixed opaque-heavy corpus".into());
    }

    let expected_source = litchi_xls::comments::Snapshot::from_bytes(corpus.archive.clone())?;
    let expected_prepared =
        prepare_xls_comments_publication(expected_source.clone(), source_backed, update_count)?;
    let mut expected = Vec::new();
    let expected_report = expected_prepared.write_to(&mut expected)?;
    if expected == corpus.archive {
        return Err("XLS comment expected publication did not change source bytes".into());
    }
    verify_xls_comments_output(corpus, &expected, update_count)?;
    verify_xls_comments_prepared(&expected_source, &expected_prepared, &expected)?;
    if let Some(report) = expected_report
        && (report.bytes() != u64::try_from(expected.len())? || report.changed_spans() == 0)
    {
        return Err("XLS comment expected overlay report is incomplete".into());
    }
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = InstrumentedSource::new(corpus.archive.clone(), Vec::new());
        let mut source_bytes = vec![0_u8; corpus.archive.len()];
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;

        read_xls_comments_source(&source, &mut source_bytes)?;
        let snapshot = litchi_xls::comments::Snapshot::from_bytes(source_bytes)?;
        let plan_started = Instant::now();
        let prepared = prepare_xls_comments_publication(snapshot, source_backed, update_count)?;
        let semantic_staging_plan = plan_started.elapsed();

        let publication_started = Instant::now();
        let report = prepared.write_to(&mut sink)?;
        let publication = publication_started.elapsed();
        let duration = semantic_staging_plan
            .checked_add(publication)
            .ok_or("XLS comment total duration overflow")?;
        let metrics = source.snapshot();
        let evidence = prepared.evidence(
            source_backed,
            update_count,
            semantic_staging_plan,
            publication,
        )?;

        if metrics.read_calls == 0
            || metrics.read_bytes != u64::try_from(corpus.archive.len())?
            || evidence.changed_comments != update_count
            || evidence.touched_streams != 1
            || (source_backed && evidence.source_workbook_bytes != evidence.target_workbook_bytes)
            || sink.bytes != expected
        {
            return Err("XLS comment iteration has unexpected source/publication evidence".into());
        }
        if source_backed {
            let expected_splices = evidence
                .changed_comments
                .checked_mul(2)
                .ok_or("XLS comment splice-count expectation overflow")?;
            let replacement_bytes = evidence.replacement_bytes.ok_or(
                "XLS comment source-backed publication omitted replacement-byte diagnostics",
            )?;
            if evidence.splice_count != Some(expected_splices)
                || replacement_bytes == 0
                || replacement_bytes >= evidence.source_workbook_bytes
            {
                return Err(
                    "XLS comment source-backed splice diagnostics disagree with NOTE/TXO ranges"
                        .into(),
                );
            }
            let report = report.ok_or("XLS comment source-backed publication has no report")?;
            if evidence.changed_spans != Some(report.changed_spans())
                || report.bytes() != sink.summary().accepted_bytes
                || evidence.source_fingerprint
                    != Some(fingerprint_hex(report.source_fingerprint().as_bytes()))
                || evidence.target_fingerprint
                    != Some(fingerprint_hex(report.target_fingerprint().as_bytes()))
            {
                return Err("XLS comment overlay diagnostics disagree with publication".into());
            }
        } else if report.is_some() || evidence.changed_spans.is_some() {
            return Err("XLS eager comment publication reported overlay-only evidence".into());
        }
        verify_xls_comments_output(corpus, &sink.bytes, update_count)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("XLS comment output digest differs from expected output".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_xls_comments(metrics, evidence)?;
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLS comment measured output hashes are unstable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(deterministic_sink_summary(
            &sink_summaries,
            "XLS comment publication",
        )?),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn fingerprint_hex(bytes: &[u8; 32]) -> String {
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn verify_semantic_xls(
    bytes: &[u8],
    shape: WriterShape,
    updated: Option<(usize, usize, usize)>,
) -> Result<(), Box<dyn Error>> {
    use litchi_xls::cell_values::{Reference, Snapshot, Value};

    let (sheet_count, rows, columns) = shape
        .xls_dimensions()
        .ok_or("payload-heavy XLS corpus is excluded from semantic verification")?;
    let snapshot = Snapshot::from_bytes(bytes.to_vec())?;
    if snapshot.worksheet_count() != sheet_count {
        return Err("semantic XLS worksheet count differs from writer specification".into());
    }
    for (sheet, worksheet) in snapshot.worksheets().enumerate() {
        if worksheet.position() != sheet || worksheet.name() != format!("Bench{sheet:02}") {
            return Err("semantic XLS worksheet identity differs from writer specification".into());
        }
        if worksheet.cells().len() != rows * columns {
            return Err("semantic XLS cell count differs from writer specification".into());
        }
        for cell in worksheet.cells() {
            let reference = cell.reference();
            let row = usize::from(reference.row());
            let column = usize::from(reference.column());
            let mut expected = xls_expected_value(shape, sheet, row, column)?;
            if updated == Some((sheet, row, column)) {
                expected += 0.5;
            }
            if !matches!(cell.value(), Value::Number(actual) if actual.to_bits() == expected.to_bits())
                || worksheet.cell(Reference::new(row as u32, column as u32)?)? != Some(cell)
            {
                return Err("semantic XLS cell differs from writer specification".into());
            }
        }
    }
    Ok(())
}

fn run_semantic_xls(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    use litchi_core::sheet::{Cell as _, Worksheet as _};
    use litchi_xls::cell_values::{Reference, Snapshot};

    let shape = writer_shape(corpus)?;
    let (sheet_count, rows, columns) = shape
        .xls_dimensions()
        .ok_or("payload-heavy XLS corpus is excluded from semantic cases")?;
    let selected = (sheet_count / 2, rows / 2, columns / 2);
    let reference = Reference::new(selected.1 as u32, selected.2 as u32)?;
    let replacement = xls_expected_value(shape, selected.0, selected.1, selected.2)? + 0.5;
    let expected_changed = if case == Case::XlsSemanticOneEditSave {
        let source = Snapshot::from_bytes(corpus.archive.clone())?;
        let mut edit = source.edit();
        edit.set_number(selected.0.into(), reference, replacement)?;
        edit.commit()?.snapshot().bytes().to_vec()
    } else {
        corpus.archive.clone()
    };
    let expected_count = sheet_count * rows * columns;
    let expected_sum = (expected_count * (expected_count - 1) / 2) as f64;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        match case {
            Case::XlsSemanticOpen => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let workbook = litchi_xls::Workbook::new(Cursor::new(owned))?;
                let duration = started.elapsed();
                verify_semantic_xls(&corpus.archive, shape, None)?;
                std::hint::black_box(workbook);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::XlsSemanticListWorksheets => {
                let workbook = litchi_xls::Workbook::new(Cursor::new(corpus.archive.as_slice()))?;
                let started = Instant::now();
                let names = workbook
                    .sheets()
                    .iter()
                    .map(|sheet| sheet.name().to_owned())
                    .collect::<Vec<_>>();
                let duration = started.elapsed();
                let expected = (0..sheet_count)
                    .map(|sheet| format!("Bench{sheet:02}"))
                    .collect::<Vec<_>>();
                if names != expected {
                    return Err(
                        "semantic XLS worksheet list differs from writer specification".into(),
                    );
                }
                verify_semantic_xls(&corpus.archive, shape, None)?;
                std::hint::black_box(names);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::XlsSemanticOneCell => {
                let workbook = litchi_xls::Workbook::new(Cursor::new(corpus.archive.as_slice()))?;
                let started = Instant::now();
                let metadata = workbook
                    .sheet(selected.0)
                    .ok_or("semantic XLS selected tab is missing")?;
                let worksheet = workbook.xls_worksheet(
                    metadata
                        .parsed_worksheet_index()
                        .ok_or("semantic XLS selected tab is not a worksheet")?,
                )?;
                let cell = worksheet
                    .get_cell(selected.1 as u32, selected.2 as u32)
                    .ok_or("semantic XLS selected cell is missing")?;
                let value = cell
                    .value()
                    .as_float()
                    .ok_or("semantic XLS selected cell is not numeric")?;
                let duration = started.elapsed();
                if value.to_bits()
                    != xls_expected_value(shape, selected.0, selected.1, selected.2)?.to_bits()
                {
                    return Err(
                        "semantic XLS selected cell differs from writer specification".into(),
                    );
                }
                verify_semantic_xls(&corpus.archive, shape, None)?;
                std::hint::black_box(value);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::XlsSemanticFullCellScan => {
                let workbook = litchi_xls::Workbook::new(Cursor::new(corpus.archive.as_slice()))?;
                let started = Instant::now();
                let mut count = 0_usize;
                let mut sum = 0.0_f64;
                for metadata in workbook.sheets() {
                    let worksheet = workbook.xls_worksheet(
                        metadata
                            .parsed_worksheet_index()
                            .ok_or("semantic XLS tab is not a worksheet")?,
                    )?;
                    let mut cells = worksheet.cells();
                    while let Some(cell) = cells.next() {
                        let cell = cell.map_err(|error| io::Error::other(error.to_string()))?;
                        sum += cell
                            .value()
                            .as_float()
                            .ok_or("semantic XLS scan found a non-numeric cell")?;
                        count += 1;
                    }
                }
                let duration = started.elapsed();
                if count != expected_count || sum.to_bits() != expected_sum.to_bits() {
                    return Err("semantic XLS full scan differs from writer specification".into());
                }
                verify_semantic_xls(&corpus.archive, shape, None)?;
                std::hint::black_box((count, sum));
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::XlsSemanticNoopEditSave | Case::XlsSemanticOneEditSave => {
                let source = Snapshot::from_bytes(corpus.archive.clone())?;
                let updated = case == Case::XlsSemanticOneEditSave;
                let started = Instant::now();
                let mut edit = source.edit();
                if updated {
                    edit.set_number(selected.0.into(), reference, replacement)?;
                }
                let commit = edit.commit()?;
                let bytes = commit.snapshot().bytes().to_vec();
                let duration = started.elapsed();
                let diagnostics = commit.diagnostics();
                if commit.patch().is_empty() == updated
                    || diagnostics.changed_cells() != usize::from(updated)
                    || diagnostics.changed_number_fields() != usize::from(updated)
                    || diagnostics.touched_streams() != usize::from(updated)
                    || bytes != expected_changed
                {
                    return Err("semantic XLS edit/save has unexpected publication state".into());
                }
                let applied = commit.patch().apply(&source)?;
                if applied.bytes() != commit.snapshot().bytes() {
                    return Err("semantic XLS exact patch replay differs from commit".into());
                }
                let restored = commit.patch().inverse().apply(&applied)?;
                if restored.bytes() != source.bytes() {
                    return Err("semantic XLS inverse did not restore exact source bytes".into());
                }
                verify_semantic_xls(&bytes, shape, updated.then_some(selected))?;
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-XLS semantic case passed to XLS runner".into()),
        }
    }
    Ok(result(case, corpus, elapsed, None))
}

fn verify_semantic_ppt(
    bytes: &[u8],
    shape: WriterShape,
    updated: Option<(usize, usize)>,
) -> Result<(), Box<dyn Error>> {
    let (slide_count, boxes_per_slide) = shape.ppt_dimensions();
    let mut package = litchi_ppt::Package::from_reader(Cursor::new(bytes))?;
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    if slides.len() != slide_count {
        return Err("semantic PPT slide count differs from writer specification".into());
    }
    let mut expected_slides = Vec::with_capacity(slide_count);
    for (slide, item) in slides.iter().enumerate() {
        let shapes = item.shapes()?;
        if shapes.len() != boxes_per_slide {
            return Err("semantic PPT shape count differs from writer specification".into());
        }
        let mut expected_shapes = Vec::with_capacity(boxes_per_slide);
        for (shape_index, object) in shapes.iter().enumerate() {
            let expected = if updated == Some((slide, shape_index)) {
                updated_writer_text("ppt", slide, shape_index, 0)
            } else {
                writer_text("ppt", slide, shape_index, 0)
            };
            if object.text()? != expected {
                return Err("semantic PPT shape text differs from writer specification".into());
            }
            expected_shapes.push(expected);
        }
        let expected_slide = expected_shapes.join("\n");
        if item.text()? != expected_slide {
            return Err("semantic PPT slide text differs from writer specification".into());
        }
        expected_slides.push(expected_slide);
    }
    if presentation.text()? != expected_slides.join("\n\n") {
        return Err("semantic PPT full text differs from writer specification".into());
    }
    Ok(())
}

fn run_semantic_ppt(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    use litchi_ppt::{slide_order::Snapshot, text_edit::Target};

    let shape = writer_shape(corpus)?;
    if shape == WriterShape::PayloadHeavy {
        return Err("payload-heavy PPT corpus is excluded from semantic cases".into());
    }
    let (slide_count, boxes_per_slide) = shape.ppt_dimensions();
    let linear = slide_count * boxes_per_slide / 2;
    let selected = (linear / boxes_per_slide, linear % boxes_per_slide);
    let target = Target::new(Position::new(selected.0), Position::new(selected.1));
    let expected_changed = if case == Case::PptSemanticOneEditSave {
        let source = Snapshot::from_bytes(corpus.archive.clone())?;
        let mut edit = source.edit()?;
        edit.set_shape_text(
            target,
            updated_writer_text("ppt", selected.0, selected.1, 0),
        )?;
        edit.commit()?.snapshot().bytes().to_vec()
    } else {
        corpus.archive.clone()
    };
    let mut elapsed = Vec::with_capacity(samples);
    let mut final_slide_order_snapshot = None;
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        match case {
            Case::PptSemanticOpen => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let mut package = litchi_ppt::Package::from_reader(Cursor::new(owned))?;
                let presentation = package.presentation()?;
                let duration = started.elapsed();
                verify_semantic_ppt(&corpus.archive, shape, None)?;
                std::hint::black_box(presentation);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::PptSemanticListSlides => {
                let mut package =
                    litchi_ppt::Package::from_reader(Cursor::new(corpus.archive.as_slice()))?;
                let presentation = package.presentation()?;
                let started = Instant::now();
                let slides = presentation.slides()?;
                let duration = started.elapsed();
                if slides.len() != slide_count {
                    return Err("semantic PPT slide list differs from writer specification".into());
                }
                verify_semantic_ppt(&corpus.archive, shape, None)?;
                std::hint::black_box(slides);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::PptSemanticOneShapeText => {
                let mut package =
                    litchi_ppt::Package::from_reader(Cursor::new(corpus.archive.as_slice()))?;
                let presentation = package.presentation()?;
                let started = Instant::now();
                let text = presentation
                    .slides()?
                    .into_iter()
                    .nth(selected.0)
                    .ok_or("semantic PPT selected slide is missing")?
                    .shapes()?
                    .get(selected.1)
                    .ok_or("semantic PPT selected shape is missing")?
                    .text()?;
                let duration = started.elapsed();
                if text != writer_text("ppt", selected.0, selected.1, 0) {
                    return Err(
                        "semantic PPT selected shape differs from writer specification".into(),
                    );
                }
                verify_semantic_ppt(&corpus.archive, shape, None)?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::PptSemanticFullText => {
                let mut package =
                    litchi_ppt::Package::from_reader(Cursor::new(corpus.archive.as_slice()))?;
                let presentation = package.presentation()?;
                let started = Instant::now();
                let text = presentation.text()?;
                let duration = started.elapsed();
                verify_semantic_ppt(&corpus.archive, shape, None)?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::PptSlideOrderSnapshotOpen => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let snapshot = Snapshot::from_bytes(owned)?;
                let duration = started.elapsed();
                if snapshot.slide_count() != slide_count {
                    return Err(
                        "PPT slide-order snapshot count differs from writer specification".into(),
                    );
                }
                std::hint::black_box(&snapshot);
                final_slide_order_snapshot = Some(snapshot);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::PptSemanticNoopEditSave | Case::PptSemanticOneEditSave => {
                let source = Snapshot::from_bytes(corpus.archive.clone())?;
                let updated = case == Case::PptSemanticOneEditSave;
                let started = Instant::now();
                let mut edit = source.edit()?;
                if updated {
                    edit.set_shape_text(
                        target,
                        updated_writer_text("ppt", selected.0, selected.1, 0),
                    )?;
                }
                let commit = edit.commit()?;
                let bytes = commit.snapshot().bytes().to_vec();
                let duration = started.elapsed();
                if commit.patch().is_empty() == updated
                    || commit.patch().shape_text_changes().len() != usize::from(updated)
                    || bytes != expected_changed
                {
                    return Err("semantic PPT edit/save has unexpected publication state".into());
                }
                if updated {
                    let [change] = commit.patch().shape_text_changes() else {
                        return Err("semantic PPT edit patch lacks its one shape change".into());
                    };
                    if change.target() != target
                        || change.before() != writer_text("ppt", selected.0, selected.1, 0)
                        || change.after() != updated_writer_text("ppt", selected.0, selected.1, 0)
                    {
                        return Err("semantic PPT shape change differs from specification".into());
                    }
                }
                let applied = commit.patch().apply(&source)?;
                if applied.bytes() != commit.snapshot().bytes() {
                    return Err("semantic PPT exact patch replay differs from commit".into());
                }
                let restored = commit.patch().inverse().apply(&applied)?;
                if restored.bytes() != source.bytes() {
                    return Err("semantic PPT inverse did not restore exact source bytes".into());
                }
                verify_semantic_ppt(&bytes, shape, updated.then_some(selected))?;
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-PPT semantic case passed to PPT runner".into()),
        }
    }
    if let Some(snapshot) = final_slide_order_snapshot {
        verify_semantic_ppt(snapshot.bytes(), shape, None)?;
    }
    Ok(result(case, corpus, elapsed, None))
}

fn run_ppt_text_edit_one_edit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    use litchi_ppt::text_edit::{Snapshot, Target};

    let shape = writer_shape(corpus)?;
    if shape == WriterShape::PayloadHeavy {
        return Err("payload-heavy PPT corpus is excluded from semantic cases".into());
    }
    let (slide_count, boxes_per_slide) = shape.ppt_dimensions();
    let linear = slide_count * boxes_per_slide / 2;
    let selected = (linear / boxes_per_slide, linear % boxes_per_slide);
    let target = Target::new(Position::new(selected.0), Position::new(selected.1));
    let source_text = writer_text("ppt", selected.0, selected.1, 0);
    let replacement = updated_writer_text("ppt", selected.0, selected.1, 0);

    let expected_changed = {
        let source = Snapshot::from_bytes(corpus.archive.clone())?;
        let mut edit = source.edit_text(target)?;
        if edit.text() != source_text {
            return Err("PPT text-edit source differs from writer specification".into());
        }
        edit.set_text(replacement.clone())?;
        edit.commit()?.snapshot().bytes().to_vec()
    };

    let mut elapsed = Vec::with_capacity(samples);
    let mut final_publication = None;
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        // Snapshot construction is deliberately outside the measured interval:
        // this case attributes the public text-edit transaction itself.
        let source = Snapshot::from_bytes(corpus.archive.clone())?;
        let started = Instant::now();
        let mut edit = source.edit_text(target)?;
        edit.set_text(replacement.clone())?;
        let commit = edit.commit()?;
        let duration = started.elapsed();

        if commit.patch().is_empty()
            || commit.patch().before() != corpus.archive
            || commit.patch().after() != expected_changed
            || commit.snapshot().bytes() != expected_changed
        {
            return Err("PPT text-edit publication differs from specification".into());
        }
        std::hint::black_box(commit.snapshot());
        final_publication = Some((source, commit));
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let (source, commit) = final_publication.ok_or("PPT text-edit case produced no publication")?;
    let applied = commit.patch().apply(&source)?;
    if applied.bytes() != commit.snapshot().bytes() {
        return Err("PPT text-edit exact patch replay differs from commit".into());
    }
    let restored = commit.patch().inverse().apply(&applied)?;
    if restored.bytes() != source.bytes() {
        return Err("PPT text-edit inverse did not restore exact source bytes".into());
    }
    let readback = commit.snapshot().edit_text(target)?;
    if readback.target() != target || readback.text() != replacement {
        return Err("PPT text-edit public readback differs from replacement".into());
    }
    verify_semantic_ppt(commit.snapshot().bytes(), shape, Some(selected))?;

    Ok(result(Case::PptTextEditOneEditSave, corpus, elapsed, None))
}

fn run_semantic_rtf(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let shape = semantic_shape(corpus)?;
    let variant = semantic_rtf_variant(corpus)?;
    if !variant.supports_shape(shape) || !variant.supports_case(case) {
        return Err("semantic RTF case is unsupported for the selected variant and shape".into());
    }
    let paragraph_count = semantic_rtf_paragraph_count(shape, variant);
    let selected = paragraph_count / 2;
    let updated = match case {
        Case::RtfSemanticOneEditSave => vec![selected],
        Case::RtfSemanticOnePercentEditSave => semantic_update_indices(paragraph_count)?,
        _ => Vec::new(),
    };
    let replacements = updated
        .iter()
        .map(|&position| {
            litchi_rtf::edit::ParagraphTextReplacement::new(
                position,
                semantic_rtf_variant_text(variant, position, true),
            )
        })
        .collect::<Vec<_>>();
    let lifecycle_projection = matches!(
        case,
        Case::RtfSemanticRemoveParagraphSave | Case::RtfSemanticMoveParagraphSave
    )
    .then(|| semantic_rtf_lifecycle_projection(case, shape))
    .transpose()?;
    let expected_changed = if lifecycle_projection.is_some() {
        let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
        stage_semantic_rtf_lifecycle(case, &document)?
            .snapshot()
            .to_bytes()?
    } else if !updated.is_empty() {
        let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
        let mut edit = document.edit();
        if case == Case::RtfSemanticOnePercentEditSave {
            edit.replace_body_paragraph_texts(&replacements)?;
        } else {
            edit.replace_paragraph_text(
                selected,
                semantic_rtf_variant_text(variant, selected, true),
            )?;
        }
        edit.commit()?.snapshot().to_bytes()?
    } else {
        corpus.archive.clone()
    };
    if let Some(expected_projection) = lifecycle_projection.as_deref() {
        let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
        let commit = stage_semantic_rtf_lifecycle(case, &document)?;
        verify_semantic_rtf_lifecycle_commit(
            case,
            &document,
            &commit,
            expected_projection,
            &expected_changed,
        )?;
    }
    let expected_text = semantic_rtf_expected_text(shape, variant, &[]);
    let sink_ceiling = if case == Case::RtfSemanticTextToSink {
        u64::try_from(expected_text.len())?
    } else {
        u64::try_from(expected_changed.len())?
    };
    let mut elapsed = Vec::with_capacity(samples);
    let mut sinks = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        match case {
            Case::RtfSemanticOpen => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let document = litchi_rtf::Document::from_bytes(&owned)?;
                let duration = started.elapsed();
                verify_semantic_rtf(&document, shape, variant, &[])?;
                std::hint::black_box(document);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::RtfSemanticParagraphCount => {
                let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
                let started = Instant::now();
                let count = document.paragraph_count();
                let duration = started.elapsed();
                if count != paragraph_count {
                    return Err("semantic RTF paragraph count differs from specification".into());
                }
                verify_semantic_rtf(&document, shape, variant, &[])?;
                std::hint::black_box(count);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::RtfSemanticListParagraphs => {
                let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
                let started = Instant::now();
                let count = document.body().paragraphs().count();
                let duration = started.elapsed();
                if count != paragraph_count {
                    return Err("semantic RTF paragraph list differs from specification".into());
                }
                verify_semantic_rtf(&document, shape, variant, &[])?;
                std::hint::black_box(count);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::RtfSemanticCollectParagraphs => {
                let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
                let started = Instant::now();
                let paragraphs = document.body().paragraphs().collect::<Vec<_>>();
                let duration = started.elapsed();
                if paragraphs.len() != paragraph_count {
                    return Err(
                        "semantic RTF paragraph collection differs from specification".into(),
                    );
                }
                verify_semantic_rtf(&document, shape, variant, &[])?;
                std::hint::black_box(paragraphs);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::RtfSemanticOneParagraph => {
                let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
                let started = Instant::now();
                let paragraph = document
                    .body()
                    .paragraphs()
                    .nth(selected)
                    .ok_or("semantic RTF selected paragraph is missing")?
                    .to_text();
                let duration = started.elapsed();
                if paragraph != semantic_rtf_variant_text(variant, selected, false) {
                    return Err("semantic RTF selected paragraph differs from specification".into());
                }
                verify_semantic_rtf(&document, shape, variant, &[])?;
                std::hint::black_box(paragraph);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::RtfSemanticFullText => {
                let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
                let started = Instant::now();
                let text = document.text();
                let duration = started.elapsed();
                if text != expected_text {
                    return Err("semantic RTF full text differs from specification".into());
                }
                verify_semantic_rtf(&document, shape, variant, &[])?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::RtfSemanticTextToSink => {
                let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
                let mut sink = CountingSink::bounded(sink_ceiling, sink_ceiling);
                sink.reserve_budget()?;
                let options = litchi_core::TextOutputOptions::new(
                    "\n",
                    "",
                    sink_ceiling,
                    u64::try_from(paragraph_count)?,
                );
                let started = Instant::now();
                let report = document.write_text_to(&mut sink, options)?;
                let duration = started.elapsed();
                let summary = sink.summary();
                if sink.bytes != expected_text.as_bytes() {
                    return Err(
                        "semantic RTF text sink differs from expected UTF-8 body text".into(),
                    );
                }
                if report.bytes_written() != sink_ceiling
                    || report.objects_written() != u64::try_from(paragraph_count)?
                {
                    return Err("semantic RTF text sink progress differs from specification".into());
                }
                verify_semantic_rtf(&document, shape, variant, &[])?;
                std::hint::black_box(report);
                if iteration >= warmup_iterations {
                    sinks.push(summary);
                }
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::RtfSemanticStreamSave
            | Case::RtfSemanticNoopEditSave
            | Case::RtfSemanticOneEditSave
            | Case::RtfSemanticOnePercentEditSave
            | Case::RtfSemanticRemoveParagraphSave
            | Case::RtfSemanticMoveParagraphSave => {
                let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
                let mut sink = CountingSink::bounded(sink_ceiling, sink_ceiling);
                sink.reserve_budget()?;
                let started = Instant::now();
                let (published, expected_updates, commit) = match case {
                    Case::RtfSemanticStreamSave => (document.clone(), &[][..], None),
                    Case::RtfSemanticNoopEditSave => {
                        let commit = document.edit().commit()?;
                        if commit.diagnostics().changed()
                            || commit.diagnostics().operation_count() != 0
                            || !commit.snapshot().same_snapshot(&document)
                        {
                            return Err("semantic RTF no-op commit changed its source".into());
                        }
                        (commit.into_snapshot(), &[][..], None)
                    },
                    Case::RtfSemanticOneEditSave => {
                        let mut edit = document.edit();
                        edit.replace_paragraph_text(
                            selected,
                            semantic_rtf_variant_text(variant, selected, true),
                        )?;
                        let commit = edit.commit()?;
                        if !commit.diagnostics().changed()
                            || commit.diagnostics().operation_count() != 1
                        {
                            return Err(
                                "semantic RTF changed commit has unexpected diagnostics".into()
                            );
                        }
                        let published = commit.snapshot().clone();
                        (published, updated.as_slice(), Some(commit))
                    },
                    Case::RtfSemanticOnePercentEditSave => {
                        let mut edit = document.edit();
                        edit.replace_body_paragraph_texts(&replacements)?;
                        let commit = edit.commit()?;
                        if !commit.diagnostics().changed()
                            || commit.diagnostics().operation_count() != updated.len()
                        {
                            return Err(
                                "semantic RTF one-percent commit has unexpected diagnostics".into(),
                            );
                        }
                        let published = commit.snapshot().clone();
                        (published, updated.as_slice(), Some(commit))
                    },
                    Case::RtfSemanticRemoveParagraphSave | Case::RtfSemanticMoveParagraphSave => {
                        let commit = stage_semantic_rtf_lifecycle(case, &document)?;
                        if !commit.diagnostics().changed()
                            || commit.diagnostics().operation_count() != 1
                        {
                            return Err(
                                "semantic RTF lifecycle commit has unexpected diagnostics".into()
                            );
                        }
                        let published = commit.snapshot().clone();
                        (published, &[][..], Some(commit))
                    },
                    _ => return Err("non-save RTF case reached save branch".into()),
                };
                published.write_to(&mut sink)?;
                let duration = started.elapsed();
                let summary = sink.summary();
                if sink.bytes != expected_changed {
                    return Err("semantic RTF save differs from deterministic output".into());
                }
                let reopened = litchi_rtf::Document::from_bytes(&sink.bytes)?;
                if let Some(expected_projection) = lifecycle_projection.as_deref() {
                    verify_semantic_rtf_lifecycle_projection(&reopened, expected_projection)?;
                } else {
                    verify_semantic_rtf(&reopened, shape, variant, expected_updates)?;
                }
                if let Some(commit) = commit {
                    let applied = commit.patch().apply(&document)?;
                    if applied.to_bytes()? != sink.bytes {
                        return Err("semantic RTF patch replay differs from publication".into());
                    }
                    let restored = commit.patch().inverse().apply(&applied)?;
                    if restored.to_bytes()? != corpus.archive {
                        return Err(
                            "semantic RTF inverse did not restore exact source bytes".into()
                        );
                    }
                }
                if iteration >= warmup_iterations {
                    sinks.push(summary);
                }
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-RTF semantic case passed to RTF runner".into()),
        }
    }
    let sink = (!sinks.is_empty())
        .then(|| deterministic_sink_summary(&sinks, "semantic RTF sequential output"))
        .transpose()?;
    let mut result = result(case, corpus, elapsed, sink);
    if lifecycle_projection.is_some() {
        result.output_sha256 = Some(sha256_hex(&expected_changed));
    }
    Ok(result)
}

fn run_rtf_logical_tail_append(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if !matches!(semantic_rtf_variant(corpus)?, RtfSemanticVariant::Plain) {
        return Err("RTF logical-tail cases require the plain uncompressed corpus".into());
    }
    let shape = semantic_shape(corpus)?;
    let source = litchi_rtf::Document::from_bytes(&corpus.archive)?;
    let appended = (0..rtf_logical_tail_paragraph_count(shape))
        .map(|index| rtf_logical_tail_text(shape, index))
        .collect::<Vec<_>>();
    let paragraph_inputs = appended.iter().map(String::as_str).collect::<Vec<_>>();
    let input_bytes = appended.iter().try_fold(0usize, |total, text| {
        total
            .checked_add(text.len())
            .ok_or("RTF logical-tail input byte count overflows usize")
    })?;
    let limits = rtf_logical_tail_limits(corpus.archive.len(), input_bytes, appended.len())?;
    let changed = stage_rtf_logical_tail(&source, &paragraph_inputs, limits)?;
    let noop = {
        let edit = source.tail_append_with_limits(litchi_rtf::TailSelector::Body, limits);
        edit.commit()?
    };
    let expected = changed.snapshot().to_bytes()?;
    verify_rtf_logical_tail_gates(&source, &changed, &noop, &appended, limits, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let source_bytes = u64::try_from(corpus.archive.len())?;
    let output_bytes = u64::try_from(expected.len())?;
    let inserted_bytes = output_bytes
        .checked_sub(source_bytes)
        .ok_or("RTF logical-tail output is smaller than its source")?;
    let maximum = output_bytes;
    let mut elapsed = Vec::with_capacity(samples);
    let mut summaries = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let mut sink = WindowedHashingSink::new(maximum, RTF_LOGICAL_TAIL_SINK_WINDOW_BYTES)?;
        let started = Instant::now();
        let written = match case {
            Case::RtfLogicalTailAppend => {
                let commit = stage_rtf_logical_tail(&source, &paragraph_inputs, limits)?;
                commit.write_to(&mut sink, limits)?
            },
            Case::RtfLogicalTailNoopSave => {
                let edit = source.tail_append_with_limits(litchi_rtf::TailSelector::Body, limits);
                let commit = edit.commit()?;
                commit.write_to(&mut sink, limits)?
            },
            _ => return Err("non-tail case passed to the logical-tail runner".into()),
        };
        let duration = started.elapsed();
        let (mut summary, digest) = sink.finish();
        let expected_written = if case == Case::RtfLogicalTailAppend {
            output_bytes
        } else {
            source_bytes
        };
        let expected_digest_for_case = if case == Case::RtfLogicalTailAppend {
            &expected_digest
        } else {
            &corpus.manifest.archive_sha256
        };
        if u64::try_from(written)? != expected_written || summary.accepted_bytes != expected_written
        {
            return Err("RTF logical-tail sink byte count differs from expected output".into());
        }
        if &digest != expected_digest_for_case {
            return Err("RTF logical-tail sink digest differs from expected output".into());
        }
        summary.rtf_tail_append = Some(RtfTailAppendSummary {
            operation: if case == Case::RtfLogicalTailAppend {
                "append"
            } else {
                "exact_noop"
            },
            source_bytes,
            input_bytes: if case == Case::RtfLogicalTailAppend {
                u64::try_from(input_bytes)?
            } else {
                0
            },
            inserted_bytes: if case == Case::RtfLogicalTailAppend {
                inserted_bytes
            } else {
                0
            },
            output_bytes: expected_written,
            paragraphs: if case == Case::RtfLogicalTailAppend {
                u64::try_from(appended.len())?
            } else {
                0
            },
            runs: if case == Case::RtfLogicalTailAppend {
                u64::try_from(appended.len())?
            } else {
                0
            },
            sink_window_bytes: u64::try_from(RTF_LOGICAL_TAIL_SINK_WINDOW_BYTES)?,
            exact_noop_verified: true,
            in_memory_patch_verified: true,
            durable_patch_verified: true,
            reopen_verified: true,
            source_conflict_verified: true,
        });
        if iteration >= warmup_iterations {
            summaries.push(summary);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let sink = deterministic_sink_summary(&summaries, case.name())?;
    if sink.retained_output_bytes != Some(0)
        || sink.retained_authoring_window_bytes
            != Some(u64::try_from(RTF_LOGICAL_TAIL_SINK_WINDOW_BYTES)?)
        || sink.largest_write > u64::try_from(RTF_LOGICAL_TAIL_SINK_WINDOW_BYTES)?
    {
        return Err("RTF logical-tail sink exceeded its fixed publication window".into());
    }
    let mut result = result(case, corpus, elapsed, Some(sink));
    result.output_sha256 = Some(if case == Case::RtfLogicalTailAppend {
        expected_digest
    } else {
        corpus.manifest.archive_sha256.clone()
    });
    Ok(result)
}

fn run_read_at_validation_report<F>(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    validate: F,
) -> Result<CaseResult, Box<dyn Error>>
where
    F: Fn(Arc<dyn ReadAt>) -> Result<ValidateReport, Box<dyn Error>>,
{
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut expected_validation = None;
    let mut measured_validation = None;
    let source_before = sha256_hex(&corpus.archive);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(corpus.archive.clone(), Vec::new()));
        let source_for_validation: Arc<dyn ReadAt> = source.clone();
        let started = Instant::now();
        let report = validate(source_for_validation)?;
        let duration = started.elapsed();
        let snapshot = source.snapshot();
        let mut summary = generic_validation_summary(
            &report,
            &corpus.archive,
            Some(snapshot.read_calls),
            Some(snapshot.read_bytes),
        )?;
        summary.source_sha256_before.clone_from(&source_before);
        summary.source_sha256_after = sha256_hex(&corpus.archive);
        if summary.source_sha256_before != summary.source_sha256_after {
            return Err(format!("{} mutated its source bytes", case.name()).into());
        }
        require_complete_validation(case, &summary)?;
        if let Some(expected) = &expected_validation {
            if expected != &summary {
                return Err(
                    format!("{} validation topology changed across samples", case.name()).into(),
                );
            }
        } else {
            expected_validation = Some(summary.clone());
        }
        if iteration >= warmup_iterations {
            if measured_validation.is_none() {
                measured_validation = Some(summary);
            }
            source_summary.record(snapshot);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let validation = measured_validation
        .or(expected_validation)
        .ok_or("validation case produced no samples")?;
    if source_summary.read_calls.len() != samples || source_summary.read_bytes.contains(&0) {
        return Err(format!("{} did not record bounded source reads", case.name()).into());
    }
    source_summary.validation = Some(validation);
    Ok(result_with_source(case, corpus, elapsed, source_summary))
}

fn run_xls_validation_report(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    run_read_at_validation_report(
        Case::XlsValidationReport,
        corpus,
        warmup_iterations,
        samples,
        |source| Ok(litchi_xls::validation::validate_source(source)?),
    )
}

fn run_docx_validation_report(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    run_read_at_validation_report(
        Case::DocxValidationReport,
        corpus,
        warmup_iterations,
        samples,
        |source| Ok(litchi_docx::validate_read_at(source)?),
    )
}

fn run_pptx_validation_report(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    run_read_at_validation_report(
        Case::PptxValidationReport,
        corpus,
        warmup_iterations,
        samples,
        |source| Ok(litchi_pptx::validate_source(source)?),
    )
}

fn run_borrowed_validation_report<R, F, S>(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    validate: F,
    summarize: S,
) -> Result<CaseResult, Box<dyn Error>>
where
    F: Fn(&[u8]) -> Result<R, Box<dyn Error>>,
    S: Fn(&R, &[u8]) -> Result<ValidationSummary, Box<dyn Error>>,
{
    let mut elapsed = Vec::with_capacity(samples);
    let mut expected_validation = None;
    let mut measured_validation = None;
    let source_before = sha256_hex(&corpus.archive);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let (report, duration) = {
            let started = Instant::now();
            let report = validate(&corpus.archive)?;
            (report, started.elapsed())
        };
        let mut summary = summarize(&report, &corpus.archive)?;
        summary.source_sha256_before.clone_from(&source_before);
        summary.source_sha256_after = sha256_hex(&corpus.archive);
        if summary.source_sha256_before != summary.source_sha256_after {
            return Err(format!("{} mutated its source bytes", case.name()).into());
        }
        require_complete_validation(case, &summary)?;
        if let Some(expected) = &expected_validation {
            if expected != &summary {
                return Err(
                    format!("{} validation topology changed across samples", case.name()).into(),
                );
            }
        } else {
            expected_validation = Some(summary.clone());
        }
        if iteration >= warmup_iterations && measured_validation.is_none() {
            measured_validation = Some(summary);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let source = SourceSummary {
        validation: Some(
            measured_validation
                .or(expected_validation)
                .ok_or("validation case produced no samples")?,
        ),
        ..SourceSummary::default()
    };
    Ok(result_with_source(case, corpus, elapsed, source))
}

fn run_rtf_validation_report(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    run_borrowed_validation_report(
        Case::RtfValidationReport,
        corpus,
        warmup_iterations,
        samples,
        |bytes| Ok(litchi_rtf::ValidationReport::from_bytes(bytes)?),
        rtf_validation_summary,
    )
}

fn run_odf_validation_report(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    run_borrowed_validation_report(
        Case::OdfValidationReport,
        corpus,
        warmup_iterations,
        samples,
        |bytes| Ok(litchi_odf_common::validate_package(bytes)?),
        |report, bytes| generic_validation_summary(report, bytes, None, None),
    )
}

fn verify_odf_mimetype_repair(
    corpus: &Corpus,
) -> Result<(OdfRepairSummary, Vec<u8>), Box<dyn Error>> {
    let shape = semantic_shape(corpus)?;
    let canonical = semantic_odt_bytes(shape)?;
    let report = litchi_odf_common::validate_package(&corpus.archive)?;
    let plan = litchi_odf_common::plan_odf_repair(
        &corpus.archive,
        &report,
        litchi_odf_common::OdfRepairLimits::default(),
    )?;
    let preview = plan.preview();
    let plan_json = plan.to_json()?;
    if preview.repair_id() != litchi_odf_common::MIMETYPE_LOCAL_EXTRA_REPAIR
        || preview.schema() != litchi_odf_common::MIMETYPE_REPAIR_PLAN_SCHEMA
        || preview.intent() != litchi_odf_common::RepairIntentKind::NonDestructive
        || preview.is_noop()
        || preview.source_len() != u64::try_from(corpus.archive.len())?
        || preview.output_len() != u64::try_from(canonical.len())?
        || preview.source_fingerprint().as_hex() != corpus.manifest.archive_sha256
        || preview.output_fingerprint().as_hex() != sha256_hex(&canonical)
    {
        return Err("ODF repair preview differs from the deterministic corpus".into());
    }

    let patch = plan.apply()?;
    if patch.target_bytes() != canonical {
        return Err("ODF repair patch did not recover the canonical ODT bytes".into());
    }
    let mut applied = Vec::new();
    patch.write_to(&mut applied)?;
    if applied != canonical {
        return Err("ODF repair patch publication differs from the canonical ODT".into());
    }
    let mut restored = Vec::new();
    patch.inverse().write_to(&mut restored)?;
    if restored != corpus.archive {
        return Err("ODF repair inverse did not restore the exact source".into());
    }

    let mut stale = corpus.archive.clone();
    let last = stale
        .last_mut()
        .ok_or("ODF repair corpus is unexpectedly empty")?;
    *last ^= 1;
    let mut stale_output = Vec::new();
    if !matches!(
        patch.apply_to(&stale, &mut stale_output),
        Err(litchi_odf_common::RepairError::SourceChanged { .. })
    ) || !stale_output.is_empty()
    {
        return Err("ODF repair patch did not refuse a stale source before output".into());
    }

    let canonical_report = litchi_odf_common::validate_package(&canonical)?;
    if !canonical_report.is_complete()
        || canonical_report.has_errors()
        || !matches!(
            litchi_odf_common::plan_odf_repair(
                &canonical,
                &canonical_report,
                litchi_odf_common::OdfRepairLimits::default(),
            ),
            Err(litchi_odf_common::RepairError::ReportMismatch)
        )
    {
        return Err("canonical ODT did not produce an exact repair no-plan refusal".into());
    }

    let mut partial = PrefixFailSink {
        accepted: 0,
        fail_after: 1,
    };
    let partial_verified = matches!(
        plan.write_to(&mut partial),
        Err(litchi_odf_common::RepairError::IncompleteOutput {
            progress: litchi_odf_common::RepairOutputProgress::Prefix {
                accepted: 1,
                expected,
            },
            ..
        }) if expected == preview.output_len()
    );
    if !partial_verified {
        return Err("ODF repair did not report exact partial-sink progress".into());
    }

    let effects = preview.effects();
    Ok((
        OdfRepairSummary {
            schema: preview.schema(),
            repair_id: preview.repair_id(),
            intent: preview.intent().as_str(),
            validation_issue_id: preview.validation_issue_id().to_string(),
            plan_json_sha256: sha256_hex(plan_json.as_bytes()),
            source_bytes: preview.source_len(),
            output_bytes: preview.output_len(),
            source_sha256: preview.source_fingerprint().as_hex(),
            output_sha256: preview.output_fingerprint().as_hex(),
            member_count: preview.member_count(),
            extra_field_id: plan.action().field_id(),
            extra_field_bytes: plan.action().field_bytes(),
            changed_members: effects.changed_members().to_vec(),
            changed_regions: effects
                .changed_regions()
                .iter()
                .map(|region| region.as_str())
                .collect(),
            member_payloads_preserved: effects.member_payloads_preserved(),
            reversible: effects.reversible(),
            exact_canonical_recovery_verified: true,
            patch_verified: true,
            inverse_verified: true,
            stale_source_refusal_verified: true,
            canonical_no_plan_verified: true,
            partial_sink_progress_verified: true,
        },
        canonical,
    ))
}

fn run_odf_mimetype_repair_plan(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let case = Case::OdfMimetypeRepairPlan;
    let (repair_summary, expected) = verify_odf_mimetype_repair(corpus)?;
    let expected_digest = sha256_hex(&expected);
    let mut elapsed = Vec::with_capacity(samples);
    let mut sinks = Vec::with_capacity(samples);
    let mut digests = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        // The sink retains only hash state, but repair planning deliberately
        // performs a bounded full-candidate preflight. Do not present the
        // publisher's 64 KiB write request ceiling as a total memory bound.
        let mut sink = HashingDiscardSink::without_authoring_window(u64::try_from(expected.len())?);
        let started = Instant::now();
        let report = litchi_odf_common::validate_package(&corpus.archive)?;
        let plan = litchi_odf_common::plan_odf_repair(
            &corpus.archive,
            &report,
            litchi_odf_common::OdfRepairLimits::default(),
        )?;
        let publication = plan.write_to(&mut sink)?;
        let duration = started.elapsed();
        let (summary, digest) = sink.finish();
        if publication.bytes() != u64::try_from(expected.len())?
            || publication.source_fingerprint().as_hex() != corpus.manifest.archive_sha256
            || publication.target_fingerprint().as_hex() != expected_digest
            || summary.accepted_bytes != u64::try_from(expected.len())?
            || summary.largest_write > ODF_REPAIR_PUBLICATION_SCRATCH_BYTES
            || digest != expected_digest
        {
            return Err("ODF repair publication evidence changed across iterations".into());
        }
        if iteration >= warmup_iterations {
            sinks.push(summary);
            digests.push(digest);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let sink = deterministic_sink_summary(&sinks, case.name())?;
    if digests.iter().any(|digest| digest != &expected_digest) {
        return Err("ODF repair output digest changed across samples".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(SourceSummary {
            odf_repair: Some(repair_summary),
            ..SourceSummary::default()
        }),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn run_docx_section_inventory(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut expected_validation = None;
    let mut measured_validation = None;
    let source_before = sha256_hex(&corpus.archive);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(corpus.archive.clone(), Vec::new()));
        let source_for_package: Arc<dyn ReadAt> = source.clone();
        let started = Instant::now();
        let package = litchi_docx::source_backed::Package::from_read_at(source_for_package)?;
        let snapshot = package.section_inventory_snapshot()?;
        let duration = started.elapsed();
        let inventory = section_inventory_summary(&snapshot);
        let canonical = serde_json::to_vec(&inventory)?;
        let source_after = sha256_hex(&corpus.archive);
        let summary = ValidationSummary {
            report_sha256: sha256_hex(&canonical),
            check_ids: vec!["docx.section_inventory".to_owned()],
            check_statuses: vec!["complete".to_owned()],
            issue_codes: Vec::new(),
            issue_count: 0,
            complete: true,
            has_errors: false,
            counts: BTreeMap::from([
                (
                    "sections".to_owned(),
                    u64::try_from(inventory.section_count)
                        .map_err(|_| "section count overflows u64")?,
                ),
                (
                    "paragraphs".to_owned(),
                    u64::try_from(inventory.paragraph_count)
                        .map_err(|_| "paragraph count overflows u64")?,
                ),
            ]),
            source_sha256_before: source_before.clone(),
            source_sha256_after: source_after,
            source_bytes: u64::try_from(corpus.archive.len())
                .map_err(|_| "DOCX source length overflows u64")?,
            source_read_calls: Some(source.snapshot().read_calls),
            source_read_bytes: Some(source.snapshot().read_bytes),
            section_inventory: Some(inventory),
        };
        if summary.source_sha256_before != summary.source_sha256_after {
            return Err("DOCX section inventory mutated its source bytes".into());
        }
        require_complete_validation(Case::DocxSectionInventory, &summary)?;
        if let Some(expected) = &expected_validation {
            if expected != &summary {
                return Err("DOCX section inventory topology changed across samples".into());
            }
        } else {
            expected_validation = Some(summary.clone());
        }
        if iteration >= warmup_iterations {
            if measured_validation.is_none() {
                measured_validation = Some(summary);
            }
            source_summary.record(source.snapshot());
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    source_summary.validation = Some(
        measured_validation
            .or(expected_validation)
            .ok_or("section case produced no samples")?,
    );
    Ok(result_with_source(
        Case::DocxSectionInventory,
        corpus,
        elapsed,
        source_summary,
    ))
}

fn run_semantic_docx(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let shape = semantic_shape(corpus)?;
    let updates = semantic_update_indices(shape.docx_paragraphs())?;
    let selected = match case {
        Case::DocxSemanticOneEditSave => vec![updates[0]],
        Case::DocxSemanticOnePercentEditSave => updates,
        _ => Vec::new(),
    };
    let mut elapsed = Vec::with_capacity(samples);
    let mut sinks = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        match case {
            Case::DocxSemanticCreateSmall => {
                let started = Instant::now();
                let bytes = semantic_docx_bytes(SemanticShape::Tiny)?;
                let duration = started.elapsed();
                let reopened = litchi_docx::Package::from_reader(Cursor::new(bytes.clone()))?;
                verify_semantic_docx(&reopened, SemanticShape::Tiny, &[])?;
                if bytes != corpus.archive && shape == SemanticShape::Tiny {
                    return Err(
                        "semantic DOCX creation digest differs from its deterministic corpus"
                            .into(),
                    );
                }
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::DocxSemanticOpen => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let package = litchi_docx::Package::from_reader(Cursor::new(owned))?;
                let duration = started.elapsed();
                verify_semantic_docx(&package, shape, &[])?;
                std::hint::black_box(package.document()?.paragraph_count()?);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::DocxSemanticListParagraphs => {
                let package =
                    litchi_docx::Package::from_reader(Cursor::new(corpus.archive.clone()))?;
                let document = package.document()?;
                let started = Instant::now();
                let paragraphs = document.paragraphs()?;
                let duration = started.elapsed();
                if paragraphs.len() != shape.docx_paragraphs() {
                    return Err("semantic DOCX paragraph list differs from specification".into());
                }
                verify_semantic_docx(&package, shape, &[])?;
                std::hint::black_box(paragraphs);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::DocxSemanticOneParagraph => {
                let package =
                    litchi_docx::Package::from_reader(Cursor::new(corpus.archive.clone()))?;
                let document = package.document()?;
                let index = shape.docx_paragraphs() / 2;
                let started = Instant::now();
                let paragraph = document.paragraph(index)?;
                let duration = started.elapsed();
                if paragraph
                    .ok_or("semantic DOCX selected paragraph is missing")?
                    .text()?
                    != semantic_docx_text(index, false)
                {
                    return Err(
                        "semantic DOCX selected paragraph differs from specification".into(),
                    );
                }
                verify_semantic_docx(&package, shape, &[])?;
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::DocxSemanticFullText => {
                let package =
                    litchi_docx::Package::from_reader(Cursor::new(corpus.archive.clone()))?;
                let document = package.document()?;
                let started = Instant::now();
                let text = document.text()?;
                let duration = started.elapsed();
                verify_semantic_docx(&package, shape, &[])?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::DocxSemanticNoopEditSave
            | Case::DocxSemanticOneEditSave
            | Case::DocxSemanticOnePercentEditSave => {
                let mut package =
                    litchi_docx::Package::from_reader(Cursor::new(corpus.archive.clone()))?;
                let started = Instant::now();
                let mut edit = package.edit_document()?;
                if selected.len() > 1 {
                    let replacements = selected
                        .iter()
                        .map(|index| {
                            litchi_docx::document::ParagraphTextReplacement::new(
                                Position::new(*index),
                                semantic_docx_text(*index, true),
                            )
                        })
                        .collect::<Vec<_>>();
                    edit.replace_body_paragraph_texts(&replacements)?;
                } else {
                    for index in &selected {
                        edit.replace_paragraph_text(
                            Position::new(*index),
                            semantic_docx_text(*index, true),
                        )?;
                    }
                }
                let mut sink = CountingSeekSink::default();
                let commit = package.publish_document_edit(edit)?;
                package.to_stream(&mut sink)?;
                let duration = started.elapsed();
                if commit.patch().changed() == selected.is_empty()
                    || commit.diagnostics().operations() != selected.len()
                {
                    return Err("semantic DOCX commit has an unexpected operation count".into());
                }
                let (bytes, summary) = sink.into_parts();
                let reopened = litchi_docx::Package::from_reader(Cursor::new(bytes))?;
                verify_semantic_docx(&reopened, shape, &selected)?;
                if iteration >= warmup_iterations {
                    sinks.push(summary);
                }
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-DOCX semantic case passed to DOCX runner".into()),
        }
    }
    let sink = (!sinks.is_empty())
        .then(|| deterministic_sink_summary(&sinks, "semantic DOCX edit/save"))
        .transpose()?;
    Ok(result(case, corpus, elapsed, sink))
}

fn run_semantic_pptx(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let shape = semantic_shape(corpus)?;
    let total = shape.pptx_slides() * shape.pptx_text_boxes_per_slide();
    let updates = semantic_update_indices(total)?;
    let selected = match case {
        Case::PptxSemanticOneEditSave => vec![updates[0]],
        Case::PptxSemanticOnePercentEditSave => updates,
        _ => Vec::new(),
    };
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        match case {
            Case::PptxSemanticCreateSmall => {
                let started = Instant::now();
                let bytes = semantic_pptx_bytes(SemanticShape::Tiny)?;
                let duration = started.elapsed();
                let reopened = litchi_pptx::Package::from_bytes(&bytes)?;
                verify_semantic_pptx(&reopened, SemanticShape::Tiny, &[])?;
                if bytes != corpus.archive && shape == SemanticShape::Tiny {
                    return Err(
                        "semantic PPTX creation digest differs from its deterministic corpus"
                            .into(),
                    );
                }
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::PptxSemanticOpen => {
                let started = Instant::now();
                let package = litchi_pptx::Package::from_bytes(&corpus.archive)?;
                let duration = started.elapsed();
                verify_semantic_pptx(&package, shape, &[])?;
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::PptxSemanticListSlides => {
                let package = litchi_pptx::Package::from_bytes(&corpus.archive)?;
                let presentation = package.presentation()?;
                let started = Instant::now();
                let slides = presentation.slides()?;
                let duration = started.elapsed();
                if slides.len() != shape.pptx_slides() {
                    return Err("semantic PPTX slide list differs from specification".into());
                }
                verify_semantic_pptx(&package, shape, &[])?;
                std::hint::black_box(slides);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::PptxSemanticOneSlide => {
                let package = litchi_pptx::Package::from_bytes(&corpus.archive)?;
                let presentation = package.presentation()?;
                let index = shape.pptx_slides() / 2;
                let started = Instant::now();
                let slide = presentation.slide(index)?;
                let duration = started.elapsed();
                let slide = slide.ok_or("semantic PPTX selected slide is missing")?;
                if slide.shapes()?.len() != shape.pptx_text_boxes_per_slide() {
                    return Err(
                        "semantic PPTX selected slide shape count differs from specification"
                            .into(),
                    );
                }
                verify_semantic_pptx(&package, shape, &[])?;
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::PptxSemanticFullText => {
                let package = litchi_pptx::Package::from_bytes(&corpus.archive)?;
                let presentation = package.presentation()?;
                let started = Instant::now();
                let text = presentation.text()?;
                let duration = started.elapsed();
                verify_semantic_pptx(&package, shape, &[])?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::PptxSemanticNoopEditSave
            | Case::PptxSemanticOneEditSave
            | Case::PptxSemanticOnePercentEditSave => {
                let mut package = litchi_pptx::Package::from_vec(corpus.archive.clone())?;
                let started = Instant::now();
                let mut edit = package.opened_presentation_transaction()?;
                for linear in &selected {
                    let slide = *linear / shape.pptx_text_boxes_per_slide();
                    let object = *linear % shape.pptx_text_boxes_per_slide();
                    if !edit.set_shape_text(
                        slide,
                        object,
                        semantic_pptx_text(slide, object, true),
                    )? {
                        return Err("semantic PPTX edit unexpectedly reported no change".into());
                    }
                }
                let commit = edit.commit()?;
                if commit.is_changed() == selected.is_empty() {
                    return Err("semantic PPTX commit has unexpected changed state".into());
                }
                package.apply_opened_presentation_commit(commit)?;
                let bytes = package.to_bytes()?;
                let duration = started.elapsed();
                let reopened = litchi_pptx::Package::from_bytes(&bytes)?;
                verify_semantic_pptx(&reopened, shape, &selected)?;
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-PPTX semantic case passed to PPTX runner".into()),
        }
    }
    Ok(result(case, corpus, elapsed, None))
}

fn run_semantic_odt(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let shape = semantic_shape(corpus)?;
    let index = shape.docx_paragraphs() / 2;
    let updates = semantic_update_indices(shape.docx_paragraphs())?;
    let selected = match case {
        Case::OdtSemanticOneEditSave => vec![index],
        Case::OdtSemanticOnePercentEditSave => updates,
        _ => Vec::new(),
    };
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        match case {
            Case::OdtSemanticCreateSmall => {
                let started = Instant::now();
                let bytes = semantic_odt_bytes(SemanticShape::Tiny)?;
                let duration = started.elapsed();
                let reopened = litchi_odt::Document::from_bytes(bytes.clone())?;
                verify_semantic_odt(&reopened, SemanticShape::Tiny, &[])?;
                if bytes != corpus.archive && shape == SemanticShape::Tiny {
                    return Err(
                        "semantic ODT creation differs from its deterministic corpus".into(),
                    );
                }
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdtSemanticOpen => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let document = litchi_odt::Document::from_bytes(owned)?;
                let duration = started.elapsed();
                verify_semantic_odt(&document, shape, &[])?;
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdtSemanticListParagraphs => {
                let document = litchi_odt::Document::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let paragraphs = document.paragraphs()?;
                let duration = started.elapsed();
                if paragraphs.len() != shape.docx_paragraphs() {
                    return Err("semantic ODT paragraph list differs from specification".into());
                }
                verify_semantic_odt(&document, shape, &[])?;
                std::hint::black_box(paragraphs);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdtSemanticOneParagraph => {
                let document = litchi_odt::Document::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let text = document
                    .paragraph(index)?
                    .ok_or("semantic ODT selected paragraph is missing")?
                    .text()?;
                let duration = started.elapsed();
                if text != semantic_odt_text(index, false) {
                    return Err("semantic ODT selected paragraph differs from specification".into());
                }
                verify_semantic_odt(&document, shape, &[])?;
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdtSemanticFullText => {
                let document = litchi_odt::Document::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let text = document.text()?;
                let duration = started.elapsed();
                verify_semantic_odt(&document, shape, &[])?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdtSemanticNoopEditSave
            | Case::OdtSemanticOneEditSave
            | Case::OdtSemanticOnePercentEditSave => {
                let document = litchi_odt::Document::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let mut edit = document.edit()?;
                for index in &selected {
                    edit.replace_paragraph(Position::new(*index), semantic_odt_text(*index, true))?;
                }
                let commit = edit.commit()?;
                let bytes = commit.snapshot().as_bytes().to_vec();
                let duration = started.elapsed();
                if (bytes != corpus.archive) == selected.is_empty()
                    || commit.results().len() != selected.len()
                {
                    return Err(
                        "semantic ODT edit/save changed-state or result count differs from specification"
                            .into(),
                    );
                }
                let reopened = litchi_odt::Document::from_bytes(bytes.clone())?;
                verify_semantic_odt(&reopened, shape, &selected)?;
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-ODT semantic case passed to ODT runner".into()),
        }
    }
    Ok(result(case, corpus, elapsed, None))
}

fn run_odt_media_paragraph_edit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let target = SemanticShape::Medium.docx_paragraphs() / 2;
    let mut expected_output_digest = None;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let source = litchi_odt::transaction::Snapshot::from_bytes(corpus.archive.clone())?;
        let mut edit = source.edit();
        edit.replace_paragraph(Position::new(target), semantic_odt_text(target, true))?;
        let commit = edit.commit()?;
        let bytes = commit.snapshot().as_bytes().to_vec();
        let duration = started.elapsed();
        if bytes == corpus.archive {
            return Err("media-rich ODT paragraph edit reported an exact no-op".into());
        }

        verify_odt_media_archive(&bytes, true)?;
        let replayed = commit.patch().apply(&source)?;
        if replayed.as_bytes() != bytes {
            return Err("media-rich ODT paragraph patch replay differs from commit".into());
        }
        let restored = commit.patch().inverse().apply(&replayed)?;
        if restored.as_bytes() != corpus.archive {
            return Err("media-rich ODT paragraph inverse did not restore the source".into());
        }
        if commit.patch().apply(&replayed).is_ok() {
            return Err("media-rich ODT paragraph patch accepted a stale source".into());
        }
        let digest = sha256_hex(&bytes);
        if let Some(expected) = &expected_output_digest {
            if expected != &digest {
                return Err("media-rich ODT paragraph publication is not deterministic".into());
            }
        } else {
            expected_output_digest = Some(digest);
        }
        std::hint::black_box(bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(
        Case::OdtMediaParagraphEditSave,
        corpus,
        elapsed,
        None,
    ))
}

fn run_odt_media_line_break_edit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let target = SemanticShape::Medium.docx_paragraphs() / 2;
    let mut expected_output_digest = None;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let source = litchi_odt::transaction::Snapshot::from_bytes(corpus.archive.clone())?;
        let mut edit = source.edit();
        edit.append_line_break(litchi_odt::transaction::ParagraphSelector::position(
            Position::new(target),
        ))?;
        let commit = edit.commit()?;
        let bytes = commit.snapshot().as_bytes().to_vec();
        let duration = started.elapsed();
        if bytes == corpus.archive {
            return Err("media-rich ODT line-break edit reported an exact no-op".into());
        }

        verify_odt_media_line_break_archive(&bytes)?;
        let replayed = commit.patch().apply(&source)?;
        if replayed.as_bytes() != bytes {
            return Err("media-rich ODT line-break patch replay differs from commit".into());
        }
        let restored = commit.patch().inverse().apply(&replayed)?;
        if restored.as_bytes() != corpus.archive {
            return Err("media-rich ODT line-break inverse did not restore the source".into());
        }
        if commit.patch().apply(&replayed).is_ok() {
            return Err("media-rich ODT line-break patch accepted a stale source".into());
        }
        let digest = sha256_hex(&bytes);
        if let Some(expected) = &expected_output_digest {
            if expected != &digest {
                return Err("media-rich ODT line-break publication is not deterministic".into());
            }
        } else {
            expected_output_digest = Some(digest);
        }
        std::hint::black_box(bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let mut measured = result(Case::OdtMediaLineBreakEditSave, corpus, elapsed, None);
    measured.output_sha256 = expected_output_digest;
    Ok(measured)
}

fn run_odt_media_append_run_edit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let target = SemanticShape::Medium.docx_paragraphs() / 2;
    let mut expected_output_digest = None;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let source = litchi_odt::transaction::Snapshot::from_bytes(corpus.archive.clone())?;
        let mut edit = source.edit();
        edit.append_run(Position::new(target), ODT_MEDIA_APPEND_RUN_TEXT, None)?;
        let commit = edit.commit()?;
        let bytes = commit.snapshot().as_bytes().to_vec();
        let duration = started.elapsed();
        if bytes == corpus.archive {
            return Err("media-rich ODT append-run edit reported an exact no-op".into());
        }

        verify_odt_media_append_run_archive(&bytes)?;
        let replayed = commit.patch().apply(&source)?;
        if replayed.as_bytes() != bytes {
            return Err("media-rich ODT append-run patch replay differs from commit".into());
        }
        let restored = commit.patch().inverse().apply(&replayed)?;
        if restored.as_bytes() != corpus.archive {
            return Err("media-rich ODT append-run inverse did not restore the source".into());
        }
        if commit.patch().apply(&replayed).is_ok() {
            return Err("media-rich ODT append-run patch accepted a stale source".into());
        }
        let digest = sha256_hex(&bytes);
        if let Some(expected) = &expected_output_digest {
            if expected != &digest {
                return Err("media-rich ODT append-run publication is not deterministic".into());
            }
        } else {
            expected_output_digest = Some(digest);
        }
        std::hint::black_box(bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let mut measured = result(Case::OdtMediaAppendRunEditSave, corpus, elapsed, None);
    measured.output_sha256 = expected_output_digest;
    Ok(measured)
}

fn run_odt_media_append_hyperlink_edit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let target = SemanticShape::Medium.docx_paragraphs() / 2;
    let mut expected_output_digest = None;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let source = litchi_odt::transaction::Snapshot::from_bytes(corpus.archive.clone())?;
        let mut edit = source.edit();
        edit.append_hyperlink(
            Position::new(target),
            ODT_MEDIA_APPEND_HYPERLINK_HREF,
            ODT_MEDIA_APPEND_HYPERLINK_TEXT,
        )?;
        let commit = edit.commit()?;
        let bytes = commit.snapshot().as_bytes().to_vec();
        let duration = started.elapsed();
        if bytes == corpus.archive {
            return Err("media-rich ODT append-hyperlink edit reported an exact no-op".into());
        }

        verify_odt_media_append_hyperlink_archive(&bytes)?;
        let replayed = commit.patch().apply(&source)?;
        if replayed.as_bytes() != bytes {
            return Err("media-rich ODT append-hyperlink patch replay differs from commit".into());
        }
        let restored = commit.patch().inverse().apply(&replayed)?;
        if restored.as_bytes() != corpus.archive {
            return Err(
                "media-rich ODT append-hyperlink inverse did not restore the source".into(),
            );
        }
        if commit.patch().apply(&replayed).is_ok() {
            return Err("media-rich ODT append-hyperlink patch accepted a stale source".into());
        }
        let digest = sha256_hex(&bytes);
        if let Some(expected) = &expected_output_digest {
            if expected != &digest {
                return Err(
                    "media-rich ODT append-hyperlink publication is not deterministic".into(),
                );
            }
        } else {
            expected_output_digest = Some(digest);
        }
        std::hint::black_box(bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let mut measured = result(Case::OdtMediaAppendHyperlinkEditSave, corpus, elapsed, None);
    measured.output_sha256 = expected_output_digest;
    Ok(measured)
}

fn run_odt_media_structural_paragraph_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let target = SemanticShape::Medium.docx_paragraphs() / 2;
    let inserted = match case {
        Case::OdtMediaInsertParagraphEditSave => true,
        Case::OdtMediaRemoveParagraphEditSave => false,
        _ => return Err("non-structural ODT case passed to structural runner".into()),
    };
    let operation = if inserted { "insert" } else { "remove" };
    let mut expected_output_digest = None;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let source = litchi_odt::transaction::Snapshot::from_bytes(corpus.archive.clone())?;
        let mut edit = source.edit();
        if inserted {
            edit.insert_paragraph(Position::new(target), ODT_MEDIA_INSERT_PARAGRAPH_TEXT)?;
        } else {
            edit.remove_paragraph(Position::new(target))?;
        }
        let commit = edit.commit()?;
        let bytes = commit.snapshot().as_bytes().to_vec();
        let duration = started.elapsed();
        if bytes == corpus.archive {
            return Err(
                format!("media-rich ODT paragraph {operation} reported an exact no-op").into(),
            );
        }

        verify_odt_media_structural_paragraph_archive(&bytes, inserted)?;
        let replayed = commit.patch().apply(&source)?;
        if replayed.as_bytes() != bytes {
            return Err(format!(
                "media-rich ODT paragraph {operation} patch replay differs from commit"
            )
            .into());
        }
        let restored = commit.patch().inverse().apply(&replayed)?;
        if restored.as_bytes() != corpus.archive {
            return Err(format!(
                "media-rich ODT paragraph {operation} inverse did not restore the source"
            )
            .into());
        }
        if commit.patch().apply(&replayed).is_ok() {
            return Err(format!(
                "media-rich ODT paragraph {operation} patch accepted a stale source"
            )
            .into());
        }
        let digest = sha256_hex(&bytes);
        if let Some(expected) = &expected_output_digest {
            if expected != &digest {
                return Err(format!(
                    "media-rich ODT paragraph {operation} publication is not deterministic"
                )
                .into());
            }
        } else {
            expected_output_digest = Some(digest);
        }
        std::hint::black_box(bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let mut measured = result(case, corpus, elapsed, None);
    measured.output_sha256 = expected_output_digest;
    Ok(measured)
}

fn publish_odt_embedded_resources(
    case: Case,
    source: &litchi_odt::transaction::Snapshot,
    replacements: &[litchi_odt::package::embedded::EmbeddedResource],
    changes: &[litchi_odt::package::embedded::EmbeddedResourceChange],
) -> Result<litchi_odt::transaction::Commit, Box<dyn Error>> {
    let mut edit = source.edit();
    match case {
        Case::OdtEmbeddedResourceScalarReplaceSave => {
            for (index, replacement) in replacements.iter().enumerate() {
                edit.replace_embedded_image(index, replacement)?;
            }
        },
        Case::OdtEmbeddedResourceBatchReplaceSave => {
            edit.edit_embedded_resources(changes)?;
        },
        _ => return Err("non-resource ODT case passed to resource publisher".into()),
    }
    Ok(edit.commit()?)
}

fn verify_odt_embedded_resource_commit(
    case: Case,
    source: &litchi_odt::transaction::Snapshot,
    commit: &litchi_odt::transaction::Commit,
    expected: &[u8],
) -> Result<OdtResourceProjection, Box<dyn Error>> {
    let results_match = match case {
        Case::OdtEmbeddedResourceScalarReplaceSave => {
            commit.results().len() == ODT_RESOURCE_BATCH_COUNT
                && commit
                    .results()
                    .iter()
                    .all(|result| result == &litchi_odt::transaction::OperationResult::Unit)
        },
        Case::OdtEmbeddedResourceBatchReplaceSave => {
            commit.results() == [litchi_odt::transaction::OperationResult::Indices(Vec::new())]
        },
        _ => false,
    };
    if commit.snapshot().as_bytes() != expected || expected == source.as_bytes() || !results_match {
        return Err(format!(
            "ODT embedded-resource commit differs: exact_output={}, changed={}, results={:?}",
            commit.snapshot().as_bytes() == expected,
            expected != source.as_bytes(),
            commit.results()
        )
        .into());
    }
    let projection = verify_odt_resource_batch_archive(expected, true)?;
    verify_odt_resource_batch_raw_members(source.as_bytes(), expected)?;

    let replayed = commit.patch().apply(source)?;
    if replayed.as_bytes() != expected {
        return Err("ODT embedded-resource volatile replay differs from publication".into());
    }
    if commit.patch().inverse().apply(&replayed)?.as_bytes() != source.as_bytes() {
        return Err("ODT embedded-resource volatile inverse did not restore source".into());
    }
    if commit.patch().apply(&replayed).is_ok() {
        return Err("ODT embedded-resource volatile patch accepted stale source".into());
    }

    let durable_json = commit.patch().durable()?.to_deterministic_json()?;
    let durable = litchi_odt::transaction::DurablePatch::from_deterministic_json(&durable_json)?;
    let durable_replayed = durable.apply(source)?;
    if durable_replayed.as_bytes() != expected {
        return Err("ODT embedded-resource durable replay differs from publication".into());
    }
    if durable.inverse().apply(&durable_replayed)?.as_bytes() != source.as_bytes() {
        return Err("ODT embedded-resource durable inverse did not restore source".into());
    }
    if durable.apply(&durable_replayed).is_ok() {
        return Err("ODT embedded-resource durable patch accepted stale source".into());
    }
    Ok(projection)
}

fn run_odt_embedded_resource_publication(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if !case.uses_odt_resource_batch() {
        return Err("non-resource ODT case passed to resource publication runner".into());
    }
    let source = litchi_odt::transaction::Snapshot::from_bytes(corpus.archive.clone())?;
    let replacements = (0..ODT_RESOURCE_BATCH_COUNT)
        .map(|index| odt_resource_batch_image(index, true))
        .collect::<Vec<_>>();
    let changes = replacements
        .iter()
        .enumerate()
        .map(|(index, replacement)| {
            litchi_odt::package::embedded::EmbeddedResourceChange::replace_image(
                Position::new(index),
                replacement,
            )
        })
        .collect::<Vec<_>>();

    let batch = publish_odt_embedded_resources(
        Case::OdtEmbeddedResourceBatchReplaceSave,
        &source,
        &replacements,
        &changes,
    )?;
    let expected_batch = batch.snapshot().as_bytes().to_vec();
    let batch_projection = verify_odt_embedded_resource_commit(
        Case::OdtEmbeddedResourceBatchReplaceSave,
        &source,
        &batch,
        &expected_batch,
    )?;
    let scalar = publish_odt_embedded_resources(
        Case::OdtEmbeddedResourceScalarReplaceSave,
        &source,
        &replacements,
        &changes,
    )?;
    let expected_scalar = scalar.snapshot().as_bytes().to_vec();
    let scalar_projection = verify_odt_embedded_resource_commit(
        Case::OdtEmbeddedResourceScalarReplaceSave,
        &source,
        &scalar,
        &expected_scalar,
    )?;
    if scalar_projection != batch_projection {
        return Err("matched ODT scalar and batch resource projections differ".into());
    }
    let expected = match case {
        Case::OdtEmbeddedResourceScalarReplaceSave => expected_scalar,
        Case::OdtEmbeddedResourceBatchReplaceSave => expected_batch,
        _ => return Err("non-resource ODT case reached output selection".into()),
    };

    let sink_ceiling = u64::try_from(expected.len())?;
    let expected_digest = sha256_hex(&expected);
    let mut elapsed = Vec::with_capacity(samples);
    let mut sinks = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let mut sink = CountingSink::bounded(sink_ceiling, sink_ceiling);
        sink.reserve_budget()?;
        let started = Instant::now();
        let commit = publish_odt_embedded_resources(case, &source, &replacements, &changes)?;
        sink.write_all(commit.snapshot().as_bytes())?;
        let duration = started.elapsed();

        if sink.bytes != expected {
            return Err("measured ODT embedded-resource publication differs from oracle".into());
        }
        let replayed = commit.patch().apply(&source)?;
        if replayed.as_bytes() != expected
            || commit.patch().inverse().apply(&replayed)?.as_bytes() != source.as_bytes()
            || commit.patch().apply(&replayed).is_ok()
        {
            return Err("measured ODT embedded-resource patch contract differs".into());
        }
        sinks.push(sink.summary());
        std::hint::black_box(commit);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let sink = deterministic_sink_summary(&sinks, "ODT embedded-resource publication")?;
    let mut measured = result(case, corpus, elapsed, Some(sink));
    measured.output_sha256 = Some(expected_digest);
    Ok(measured)
}

fn run_semantic_ods(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let shape = semantic_shape(corpus)?;
    let sheet = shape.ods_sheet_count() / 2;
    let row = shape.ods_rows_per_sheet() / 2;
    let column = shape.ods_columns_per_sheet() / 2;
    let sheet_name = semantic_ods_sheet_name(sheet);
    let single_update = [semantic_ods_flat_index(shape, sheet, row, column)];
    let one_percent_updates = semantic_update_indices(shape.ods_cell_count())?;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        match case {
            Case::OdsSemanticCreateSmall => {
                let started = Instant::now();
                let bytes = semantic_ods_bytes(SemanticShape::Tiny)?;
                let duration = started.elapsed();
                let reopened = litchi_ods::Spreadsheet::from_bytes(bytes.clone())?;
                verify_semantic_ods(&reopened, SemanticShape::Tiny, false)?;
                if bytes != corpus.archive && shape == SemanticShape::Tiny {
                    return Err(
                        "semantic ODS creation differs from its deterministic corpus".into(),
                    );
                }
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdsSemanticOpen => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let spreadsheet = litchi_ods::Spreadsheet::from_bytes(owned)?;
                let duration = started.elapsed();
                verify_semantic_ods(&spreadsheet, shape, false)?;
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdsSemanticListSheets => {
                let spreadsheet = litchi_ods::Spreadsheet::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let sheets = spreadsheet.sheets();
                let duration = started.elapsed();
                if sheets.len() != shape.ods_sheet_count() {
                    return Err("semantic ODS sheet list differs from specification".into());
                }
                verify_semantic_ods(&spreadsheet, shape, false)?;
                std::hint::black_box(sheets);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdsSemanticOneCell => {
                let spreadsheet = litchi_ods::Spreadsheet::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let cell = spreadsheet
                    .cell(&sheet_name, row, column)
                    .ok_or("semantic ODS selected sheet is missing")?;
                let duration = started.elapsed();
                let litchi_ods::CellView::Stored(cell) = cell else {
                    return Err("semantic ODS selected cell is missing".into());
                };
                if cell.text != semantic_ods_text(sheet, row, column, false) {
                    return Err("semantic ODS selected cell differs from specification".into());
                }
                verify_semantic_ods(&spreadsheet, shape, false)?;
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdsSemanticCellSweep => {
                let spreadsheet = litchi_ods::Spreadsheet::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let stored_cells = semantic_ods_cell_sweep(&spreadsheet, shape)?;
                let duration = started.elapsed();
                if stored_cells != shape.ods_cell_count() {
                    return Err("semantic ODS stored-cell count differs from specification".into());
                }
                verify_semantic_ods(&spreadsheet, shape, false)?;
                std::hint::black_box(stored_cells);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdsSemanticFullCellText => {
                let spreadsheet = litchi_ods::Spreadsheet::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let text = semantic_ods_full_cell_text(&spreadsheet, shape)?;
                let duration = started.elapsed();
                verify_semantic_ods(&spreadsheet, shape, false)?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdsSemanticNoopEditSave
            | Case::OdsSemanticOneEditSave
            | Case::OdsSemanticOnePercentEditSave => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let snapshot = litchi_ods::document::Snapshot::from_bytes(owned)?;
                let mut edit = snapshot.edit();
                let updated_indices = match case {
                    Case::OdsSemanticNoopEditSave => &[][..],
                    Case::OdsSemanticOneEditSave => &single_update[..],
                    Case::OdsSemanticOnePercentEditSave => &one_percent_updates,
                    _ => unreachable!("ODS edit/save case was matched above"),
                };
                if case == Case::OdsSemanticOneEditSave {
                    let text = semantic_ods_text(sheet, row, column, true);
                    edit.worksheets(|worksheets| {
                        worksheets
                            .set_cell(
                                sheet_name.as_str(),
                                row,
                                column,
                                litchi_ods::Cell::new(
                                    litchi_ods::CellValue::Text(text.clone()),
                                    text,
                                ),
                            )?
                            .ok_or_else(|| {
                                litchi_core::Error::InvalidFormat(
                                    "semantic ODS selected sheet is missing".to_owned(),
                                )
                            })?;
                        Ok(())
                    })?;
                } else if case == Case::OdsSemanticOnePercentEditSave {
                    let rows_per_sheet = shape.ods_rows_per_sheet();
                    let columns_per_sheet = shape.ods_columns_per_sheet();
                    let cells_per_sheet = rows_per_sheet * columns_per_sheet;
                    edit.worksheets(|worksheets| {
                        let mut changed = 0usize;
                        for selected_sheet in 0..shape.ods_sheet_count() {
                            let start = selected_sheet * cells_per_sheet;
                            let end = start + cells_per_sheet;
                            let changes = updated_indices
                                .iter()
                                .copied()
                                .filter(|index| (start..end).contains(index))
                                .map(|index| {
                                    let local = index - start;
                                    let selected_row = local / columns_per_sheet;
                                    let selected_column = local % columns_per_sheet;
                                    let text = semantic_ods_text(
                                        selected_sheet,
                                        selected_row,
                                        selected_column,
                                        true,
                                    );
                                    litchi_ods::worksheet::CellChange::new(
                                        selected_row,
                                        selected_column,
                                        litchi_ods::Cell::new(
                                            litchi_ods::CellValue::Text(text.clone()),
                                            text,
                                        ),
                                    )
                                })
                                .collect();
                            let selected_sheet_name = semantic_ods_sheet_name(selected_sheet);
                            changed += worksheets
                                .set_cells(selected_sheet_name.as_str(), changes)?
                                .ok_or_else(|| {
                                    litchi_core::Error::InvalidFormat(
                                        "semantic ODS selected sheet is missing".to_owned(),
                                    )
                                })?;
                        }
                        if changed != updated_indices.len() {
                            return Err(litchi_core::Error::InvalidFormat(
                                "semantic ODS 1% batch changed an unexpected cell count".to_owned(),
                            ));
                        }
                        Ok(())
                    })?;
                }
                let commit = edit.commit()?;
                let bytes = commit.snapshot().as_bytes().to_vec();
                let duration = started.elapsed();
                let updated = !updated_indices.is_empty();
                if (bytes != corpus.archive) != updated || commit.changed() != updated {
                    return Err(
                        "semantic ODS edit/save changed-state differs from specification".into(),
                    );
                }
                let reopened = litchi_ods::Spreadsheet::from_bytes(bytes.clone())?;
                verify_semantic_ods_updates(&reopened, shape, updated_indices)?;
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-ODS semantic case passed to ODS runner".into()),
        }
    }
    Ok(result(case, corpus, elapsed, None))
}

fn run_ods_media_one_edit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let sheet = shape.ods_sheet_count() / 2;
    let row = shape.ods_rows_per_sheet() / 2;
    let column = shape.ods_columns_per_sheet() / 2;
    let sheet_name = semantic_ods_sheet_name(sheet);
    let mut expected_output_digest = None;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let snapshot = litchi_ods::document::Snapshot::from_bytes(corpus.archive.clone())?;
        let mut edit = snapshot.edit();
        let text = semantic_ods_text(sheet, row, column, true);
        edit.worksheets(|worksheets| {
            worksheets
                .set_cell(
                    sheet_name.as_str(),
                    row,
                    column,
                    litchi_ods::Cell::new(litchi_ods::CellValue::Text(text.clone()), text),
                )?
                .ok_or_else(|| {
                    litchi_core::Error::InvalidFormat(
                        "media-rich ODS selected sheet is missing".to_owned(),
                    )
                })?;
            Ok(())
        })?;
        let commit = edit.commit()?;
        let bytes = commit.snapshot().as_bytes().to_vec();
        let duration = started.elapsed();
        if !commit.changed() || bytes == corpus.archive {
            return Err("media-rich ODS one-cell edit reported an exact no-op".into());
        }

        verify_ods_media_archive(&bytes, true)?;
        let digest = sha256_hex(&bytes);
        if let Some(expected) = &expected_output_digest {
            if expected != &digest {
                return Err("media-rich ODS publication is not deterministic".into());
            }
        } else {
            expected_output_digest = Some(digest);
        }
        std::hint::black_box(bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(Case::OdsMediaOneEditSave, corpus, elapsed, None))
}

fn run_semantic_odp(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let shape = semantic_shape(corpus)?;
    let index = shape.pptx_slides() / 2;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        match case {
            Case::OdpSemanticCreateSmall => {
                let started = Instant::now();
                let bytes = semantic_odp_bytes(SemanticShape::Tiny)?;
                let duration = started.elapsed();
                let reopened = litchi_odp::Presentation::from_bytes(bytes.clone())?;
                verify_semantic_odp(&reopened, SemanticShape::Tiny, false)?;
                if bytes != corpus.archive && shape == SemanticShape::Tiny {
                    return Err(
                        "semantic ODP creation differs from its deterministic corpus".into(),
                    );
                }
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdpSemanticOpen => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let presentation = litchi_odp::Presentation::from_bytes(owned)?;
                let duration = started.elapsed();
                verify_semantic_odp(&presentation, shape, false)?;
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdpSemanticListSlides => {
                let presentation = litchi_odp::Presentation::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let slides = presentation.slides()?;
                let duration = started.elapsed();
                if slides.len() != shape.pptx_slides() {
                    return Err("semantic ODP slide list differs from specification".into());
                }
                verify_semantic_odp(&presentation, shape, false)?;
                std::hint::black_box(slides);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdpSemanticOneSlide => {
                let presentation = litchi_odp::Presentation::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let slide = presentation
                    .slide(index)?
                    .ok_or("semantic ODP selected slide is missing")?;
                let text = slide.all_text();
                let duration = started.elapsed();
                let expected = format!(
                    "{}\n{}",
                    semantic_odp_title(index, false),
                    semantic_odp_text(index, false)
                );
                if text != expected {
                    return Err("semantic ODP selected slide differs from specification".into());
                }
                verify_semantic_odp(&presentation, shape, false)?;
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdpSemanticFullText => {
                let presentation = litchi_odp::Presentation::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let text = presentation.text()?;
                let duration = started.elapsed();
                verify_semantic_odp(&presentation, shape, false)?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdpSemanticNoopEditSave | Case::OdpSemanticOneEditSave => {
                let presentation = litchi_odp::Presentation::from_bytes(corpus.archive.clone())?;
                let snapshot = presentation.snapshot()?;
                let started = Instant::now();
                let mut transaction = snapshot.transaction()?;
                let updated = matches!(case, Case::OdpSemanticOneEditSave);
                if updated {
                    let index = shape.pptx_slides();
                    transaction.add(
                        &semantic_odp_title(index, true),
                        &semantic_odp_text(index, true),
                    )?;
                }
                let commit = transaction.commit()?;
                let bytes = commit.snapshot().bytes().to_vec();
                let duration = started.elapsed();
                if (bytes != corpus.archive) != updated || commit.changed() != updated {
                    return Err(
                        "semantic ODP edit/save changed-state differs from specification".into(),
                    );
                }
                let reopened = litchi_odp::Presentation::from_bytes(bytes.clone())?;
                verify_semantic_odp(&reopened, shape, updated)?;
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-ODP semantic case passed to ODP runner".into()),
        }
    }
    Ok(result(case, corpus, elapsed, None))
}

fn run_odp_media_textbox_edit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let text_box = litchi_odp::content::TextBox::new(
        ODP_MEDIA_TEXT_BOX_NAME,
        litchi_odp::content::RichText::plain(odp_media_text())?,
    )?;
    let mut expected_output_digest = None;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let source = litchi_odp::authoring::edit::Snapshot::from_bytes(corpus.archive.clone())?;
        let mut transaction = source.transaction()?;
        transaction.add_text_box(0usize, &text_box)?;
        let commit = transaction.commit()?;
        let bytes = commit.snapshot().bytes().to_vec();
        let duration = started.elapsed();
        if !commit.changed() || bytes == corpus.archive || commit.patch().is_noop() {
            return Err("media-rich ODP text-box edit reported an exact no-op".into());
        }

        verify_odp_media_archive(&bytes, true)?;
        let replayed = commit.patch().apply(&source)?;
        if replayed.bytes() != bytes {
            return Err("media-rich ODP text-box patch replay differs from commit".into());
        }
        let restored = commit.patch().inverse().apply(&replayed)?;
        if restored.bytes() != corpus.archive {
            return Err("media-rich ODP text-box inverse did not restore the source".into());
        }
        if commit.patch().apply(&replayed).is_ok() {
            return Err("media-rich ODP text-box patch accepted a stale source".into());
        }
        let digest = sha256_hex(&bytes);
        if let Some(expected) = &expected_output_digest {
            if expected != &digest {
                return Err("media-rich ODP text-box publication is not deterministic".into());
            }
        } else {
            expected_output_digest = Some(digest);
        }
        std::hint::black_box(bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(Case::OdpMediaTextBoxEditSave, corpus, elapsed, None))
}

fn odp_text_box_batch_models(
    source: &litchi_odp::authoring::edit::Snapshot,
) -> Result<Vec<litchi_odp::content::TextBoxModel>, Box<dyn Error>> {
    let inventory = source.rich_content()?;
    let mut models = Vec::with_capacity(ODP_TEXT_BOX_BATCH_COUNT);
    for index in 0..ODP_TEXT_BOX_BATCH_COUNT {
        let name = odp_text_box_batch_name(index);
        let mut matching = inventory
            .text_boxes()
            .iter()
            .filter(|model| model.page() == odp_text_box_batch_page(index) && model.name() == name);
        let mut model = matching
            .next()
            .ok_or_else(|| format!("ODP text-box batch source owner '{name}' is missing"))?
            .clone();
        if matching.next().is_some() {
            return Err(format!("ODP text-box batch source owner '{name}' is ambiguous").into());
        }
        model.replace_paragraph(
            0,
            &litchi_odp::content::Paragraph::plain(odp_text_box_batch_text(index, true))?,
        )?;
        if model.name() != name {
            return Err("ODP text-box batch model unexpectedly renamed its owner".into());
        }
        models.push(model);
    }
    Ok(models)
}

fn publish_odp_text_box_models(
    case: Case,
    source: &litchi_odp::authoring::edit::Snapshot,
    replacements: &[litchi_odp::content::TextBoxModelReplacement<'_>],
) -> Result<litchi_odp::authoring::edit::Commit, Box<dyn Error>> {
    let mut transaction = source.transaction()?;
    match case {
        Case::OdpMediaTextBoxScalarReplaceSave => {
            for replacement in replacements {
                transaction.replace_text_box_model(replacement.name(), replacement.model())?;
            }
        },
        Case::OdpMediaTextBoxBatchReplaceSave => {
            let changed = transaction.replace_text_box_models(replacements)?;
            if changed != ODP_TEXT_BOX_BATCH_COUNT {
                return Err("ODP text-box batch changed-owner count differs".into());
            }
        },
        _ => return Err("non-model ODP case passed to text-box publisher".into()),
    }
    Ok(transaction.commit()?)
}

fn verify_odp_text_box_model_commit(
    case: Case,
    source: &litchi_odp::authoring::edit::Snapshot,
    commit: &litchi_odp::authoring::edit::Commit,
    expected: &[u8],
) -> Result<(), Box<dyn Error>> {
    if !commit.changed()
        || commit.patch().is_noop()
        || commit.patch().domains() != [litchi_odp::authoring::edit::Domain::Content]
        || commit.snapshot().bytes() != expected
    {
        return Err(format!(
            "ODP text-box model commit differs: changed={}, noop={}, domains={:?}, exact_output={}",
            commit.changed(),
            commit.patch().is_noop(),
            commit.patch().domains(),
            commit.snapshot().bytes() == expected
        )
        .into());
    }
    verify_odp_text_box_batch_archive(expected, true)?;
    verify_odp_text_box_batch_raw_members(
        source.bytes(),
        expected,
        case == Case::OdpMediaTextBoxBatchReplaceSave,
    )?;

    let replayed = commit.patch().apply(source)?;
    if replayed.bytes() != expected {
        return Err("ODP text-box model patch replay differs from publication".into());
    }
    let restored = commit.patch().inverse().apply(&replayed)?;
    if restored.bytes() != source.bytes() {
        return Err("ODP text-box model inverse did not restore exact source bytes".into());
    }
    if commit.patch().apply(&replayed).is_ok() {
        return Err("ODP text-box model patch accepted a stale source".into());
    }

    let durable = litchi_odp::authoring::edit::Patch::from_durable_bytes(
        &commit.patch().to_durable_bytes()?,
    )?;
    let durable_replayed = durable.apply(source)?;
    if durable_replayed.bytes() != expected {
        return Err("durable ODP text-box model replay differs from publication".into());
    }
    let durable_restored = durable.inverse().apply(&durable_replayed)?;
    if durable_restored.bytes() != source.bytes() {
        return Err("durable ODP text-box model inverse did not restore source".into());
    }
    if durable.apply(&durable_replayed).is_ok() {
        return Err("durable ODP text-box model patch accepted a stale source".into());
    }
    Ok(())
}

fn run_odp_text_box_model_publication(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if !case.uses_odp_text_box_batch() {
        return Err("non-model ODP case passed to text-box publication runner".into());
    }
    let source = litchi_odp::authoring::edit::Snapshot::from_bytes(corpus.archive.clone())?;
    let models = odp_text_box_batch_models(&source)?;
    let replacements = models
        .iter()
        .enumerate()
        .map(|(index, model)| {
            litchi_odp::content::TextBoxModelReplacement::at(
                odp_text_box_batch_page(index),
                model.name(),
                model,
            )
        })
        .collect::<Vec<_>>();

    let batch = publish_odp_text_box_models(
        Case::OdpMediaTextBoxBatchReplaceSave,
        &source,
        &replacements,
    )?;
    let expected_batch = batch.snapshot().bytes().to_vec();
    verify_odp_text_box_model_commit(
        Case::OdpMediaTextBoxBatchReplaceSave,
        &source,
        &batch,
        &expected_batch,
    )?;
    let scalar = publish_odp_text_box_models(
        Case::OdpMediaTextBoxScalarReplaceSave,
        &source,
        &replacements,
    )?;
    let expected_scalar = scalar.snapshot().bytes().to_vec();
    verify_odp_text_box_model_commit(
        Case::OdpMediaTextBoxScalarReplaceSave,
        &source,
        &scalar,
        &expected_scalar,
    )?;
    let scalar_projection =
        litchi_odp::Presentation::from_bytes(expected_scalar.clone())?.text()?;
    let batch_projection = litchi_odp::Presentation::from_bytes(expected_batch.clone())?.text()?;
    if scalar_projection != batch_projection {
        return Err("matched ODP scalar and batch semantic projections differ".into());
    }
    let expected = match case {
        Case::OdpMediaTextBoxScalarReplaceSave => expected_scalar,
        Case::OdpMediaTextBoxBatchReplaceSave => expected_batch,
        _ => return Err("non-model ODP case reached output selection".into()),
    };

    let sink_ceiling = u64::try_from(expected.len())?;
    let expected_digest = sha256_hex(&expected);
    let mut elapsed = Vec::with_capacity(samples);
    let mut sinks = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let mut sink = CountingSink::bounded(sink_ceiling, sink_ceiling);
        sink.reserve_budget()?;
        let started = Instant::now();
        let commit = publish_odp_text_box_models(case, &source, &replacements)?;
        sink.write_all(commit.snapshot().bytes())?;
        let duration = started.elapsed();

        if !commit.changed()
            || commit.patch().is_noop()
            || commit.patch().domains() != [litchi_odp::authoring::edit::Domain::Content]
            || sink.bytes != expected
        {
            return Err("measured ODP text-box model publication differs from oracle".into());
        }
        let replayed = commit.patch().apply(&source)?;
        if replayed.bytes() != expected
            || commit.patch().inverse().apply(&replayed)?.bytes() != source.bytes()
            || commit.patch().apply(&replayed).is_ok()
        {
            return Err("measured ODP text-box model patch contract differs".into());
        }
        sinks.push(sink.summary());
        std::hint::black_box(commit);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let sink = deterministic_sink_summary(&sinks, "ODP text-box model publication")?;
    let mut measured = result(case, corpus, elapsed, Some(sink));
    measured.output_sha256 = Some(expected_digest);
    Ok(measured)
}

fn run_xlsx_open_owned(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let owned = corpus.archive.clone();
        let started = Instant::now();
        let workbook = Workbook::from_bytes(owned)?;
        let duration = started.elapsed();
        if workbook.len() != spec.sheet_count {
            return Err("owned XLSX open sheet count differs from corpus specification".into());
        }
        std::hint::black_box(&workbook);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(Case::XlsxOpenOwned, corpus, elapsed, None))
}

fn run_xlsx_list_sheets(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let workbook = Workbook::from_bytes(corpus.archive.clone())?;
        let started = Instant::now();
        let names = workbook
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>();
        let duration = started.elapsed();
        let expected = (0..spec.sheet_count)
            .map(xlsx_sheet_name)
            .collect::<Vec<_>>();
        if names != expected {
            return Err("XLSX sheet listing differs from corpus specification".into());
        }
        std::hint::black_box(&names);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(Case::XlsxListSheets, corpus, elapsed, None))
}

fn run_xlsx_first_cell(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let workbook = Workbook::from_bytes(corpus.archive.clone())?;
        let sheet = workbook
            .sheet("Sheet1")?
            .ok_or("XLSX first sheet is missing")?;
        let started = Instant::now();
        let cell = sheet.cell("A1")?.stored().cloned();
        let duration = started.elapsed();
        if !matches!(cell, Some(XlsxCell::Value(XlsxValue::Number(ref value))) if value.as_str() == "0")
        {
            return Err("XLSX first cell differs from deterministic expectation".into());
        }
        std::hint::black_box(&cell);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(Case::XlsxFirstCell, corpus, elapsed, None))
}

fn run_xlsx_full_cell_scan(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let expected = xlsx_cell_count(spec)?;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let workbook = Workbook::from_bytes(corpus.archive.clone())?;
        let sheet = workbook
            .sheet("Sheet1")?
            .ok_or("XLSX first sheet is missing")?;
        let started = Instant::now();
        let count = sheet.cells("A1:XFD1048576")?.count();
        let duration = started.elapsed();
        if count != expected / spec.sheet_count {
            return Err("XLSX full cell scan count differs from corpus specification".into());
        }
        std::hint::black_box(count);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(Case::XlsxFullCellScan, corpus, elapsed, None))
}

fn run_xlsx_narrow_column_range_scan(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let range = format!("B1:B{}", spec.row_count);
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let workbook = Workbook::from_bytes(corpus.archive.clone())?;
        let sheet = workbook
            .sheet("Sheet1")?
            .ok_or("XLSX first sheet is missing")?;
        // Preload the full sparse store outside the timed traversal. The
        // following narrow range is deliberately one column wide in a dense,
        // row-major worksheet, exposing iterator work that is proportional to
        // all stored cells in the selected rows.
        let _ = sheet.cell("A1")?;
        let started = Instant::now();
        let count = sheet.cells(range.as_str())?.count();
        let duration = started.elapsed();
        if count != spec.row_count {
            return Err("XLSX narrow-column range scan count differs from specification".into());
        }
        std::hint::black_box(count);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(
        Case::XlsxNarrowColumnRangeScan,
        corpus,
        elapsed,
        None,
    ))
}

fn xlsx_source_cell_has_value(cell: Option<&XlsxCell>, expected: i32) -> bool {
    matches!(cell, Some(XlsxCell::Value(XlsxValue::Number(value))) if value.as_str().parse() == Ok(expected))
}

fn verify_xlsx_source_cell(
    sheet: &litchi_xlsx::SourceWorksheet,
    address: &str,
    expected: i32,
) -> Result<(), Box<dyn Error>> {
    let view = sheet.cell(address)?;
    if !xlsx_source_cell_has_value(view.stored(), expected) {
        return Err(format!("source-backed XLSX cell {address} differs from expectation").into());
    }
    Ok(())
}

fn prove_xlsx_unselected_sheet_deferred(
    workbook: &SourceBackedWorkbook,
    source: &InstrumentedSource,
    before: SourceSnapshot,
    spec: &XlsxCorpus,
) -> Result<(), Box<dyn Error>> {
    if spec.sheet_count < 2 {
        return Err("XLSX source deferral proof requires at least two worksheets".into());
    }
    let sheet = workbook
        .sheet(xlsx_sheet_name(1).as_str())?
        .ok_or("source-backed XLSX second sheet is missing")?;
    verify_xlsx_source_cell(
        &sheet,
        "A1",
        xlsx_value(XlsxCoordinate {
            sheet: 1,
            row: 0,
            column: 0,
        }),
    )?;
    let after = source.snapshot();
    if after.xlsx.unselected_worksheets.read_calls <= before.xlsx.unselected_worksheets.read_calls
        || after.xlsx.unselected_worksheets.read_bytes
            <= before.xlsx.unselected_worksheets.read_bytes
    {
        return Err(
            "fresh unselected XLSX worksheet access performed no additional member read".into(),
        );
    }
    Ok(())
}

fn run_xlsx_source_open(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = xlsx_instrumented_source(corpus)?;
        let started = Instant::now();
        let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
        let duration = started.elapsed();
        if workbook.len() != spec.sheet_count {
            return Err("source-backed XLSX open sheet count differs from specification".into());
        }
        let metrics = source.snapshot();
        if metrics.xlsx.workbook.read_calls == 0 || metrics.xlsx.workbook.read_bytes == 0 {
            return Err("source-backed XLSX open did not read the workbook member".into());
        }

        let first = workbook
            .sheet("Sheet1")?
            .ok_or("source-backed XLSX first sheet is missing")?;
        verify_xlsx_source_cell(&first, "A1", 0)?;
        let after_proof = source.snapshot();
        if after_proof.xlsx.selected_worksheet.read_calls
            <= metrics.xlsx.selected_worksheet.read_calls
            || after_proof.xlsx.selected_worksheet.read_bytes
                <= metrics.xlsx.selected_worksheet.read_bytes
        {
            return Err("source-backed XLSX open had already materialized the first sheet".into());
        }

        std::hint::black_box(&workbook);
        if iteration >= warmup_iterations {
            source_summary.record_xlsx(metrics);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::XlsxSourceOpen,
        corpus,
        elapsed,
        source_summary,
    ))
}

fn run_xlsx_source_list_sheets(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let expected = (0..spec.sheet_count)
        .map(xlsx_sheet_name)
        .collect::<Vec<_>>();
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = xlsx_instrumented_source(corpus)?;
        let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
        source.reset();
        let started = Instant::now();
        let names = workbook
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>();
        let duration = started.elapsed();
        if names != expected {
            return Err("source-backed XLSX sheet listing differs from specification".into());
        }
        let metrics = source.snapshot();
        if metrics != SourceSnapshot::default() {
            return Err("source-backed XLSX sheet listing performed a positional read".into());
        }

        let first = workbook
            .sheet("Sheet1")?
            .ok_or("source-backed XLSX first sheet is missing")?;
        verify_xlsx_source_cell(&first, "A1", 0)?;
        let after_proof = source.snapshot();
        if after_proof.xlsx.selected_worksheet.read_calls == 0
            || after_proof.xlsx.selected_worksheet.read_bytes == 0
        {
            return Err(
                "source-backed XLSX listing had already materialized the first sheet".into(),
            );
        }

        std::hint::black_box(&names);
        if iteration >= warmup_iterations {
            source_summary.record_xlsx(metrics);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::XlsxSourceListSheets,
        corpus,
        elapsed,
        source_summary,
    ))
}

fn run_xlsx_source_first_cell(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = xlsx_instrumented_source(corpus)?;
        let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
        let first = workbook
            .sheet("Sheet1")?
            .ok_or("source-backed XLSX first sheet is missing")?;
        source.reset();
        let started = Instant::now();
        let cell = first.cell("A1")?;
        let duration = started.elapsed();
        if !xlsx_source_cell_has_value(cell.stored(), 0) {
            return Err("source-backed XLSX first cell differs from expectation".into());
        }
        let metrics = source.snapshot();
        if metrics.xlsx.selected_worksheet.read_calls == 0
            || metrics.xlsx.selected_worksheet.read_bytes == 0
        {
            return Err("source-backed XLSX first cell read no selected worksheet bytes".into());
        }
        prove_xlsx_unselected_sheet_deferred(&workbook, &source, metrics, spec)?;

        std::hint::black_box(&cell);
        if iteration >= warmup_iterations {
            source_summary.record_xlsx(metrics);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::XlsxSourceFirstCell,
        corpus,
        elapsed,
        source_summary,
    ))
}

fn run_xlsx_source_narrow_column_range_scan(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let range = format!("B1:B{}", spec.row_count);
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = xlsx_instrumented_source(corpus)?;
        let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
        let first = workbook
            .sheet("Sheet1")?
            .ok_or("source-backed XLSX first sheet is missing")?;
        source.reset();
        let started = Instant::now();
        let cells = first.cells(range.as_str())?;
        let duration = started.elapsed();
        if cells.len() != spec.row_count {
            return Err("source-backed XLSX narrow range count differs from specification".into());
        }
        for (row, cell) in cells.iter().enumerate() {
            let expected_address = xlsx_address(row, 1)?;
            let expected_value = xlsx_value(XlsxCoordinate {
                sheet: 0,
                row,
                column: 1,
            });
            if cell.address.to_string() != expected_address
                || !xlsx_source_cell_has_value(Some(&cell.cell), expected_value)
            {
                return Err(
                    "source-backed XLSX narrow range contents differ from specification".into(),
                );
            }
        }
        let metrics = source.snapshot();
        if metrics.xlsx.selected_worksheet.read_calls == 0
            || metrics.xlsx.selected_worksheet.read_bytes == 0
        {
            return Err("source-backed XLSX narrow range read no selected worksheet bytes".into());
        }
        prove_xlsx_unselected_sheet_deferred(&workbook, &source, metrics, spec)?;

        std::hint::black_box(&cells);
        if iteration >= warmup_iterations {
            source_summary.record_xlsx(metrics);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::XlsxSourceNarrowColumnRangeScan,
        corpus,
        elapsed,
        source_summary,
    ))
}

fn verify_unselected_xlsx_ranges_untouched(
    snapshot: SourceSnapshot,
    context: &str,
) -> Result<(), Box<dyn Error>> {
    if snapshot.xlsx.unselected_worksheets != RangeSnapshot::default() {
        return Err(format!("{context} touched an unselected worksheet range").into());
    }
    Ok(())
}

fn run_xlsx_range_source_open(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    config: RangeSimulationConfig,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let backing = xlsx_instrumented_source(corpus)?;
        let source = simulated_source(backing.clone(), config);
        let started = Instant::now();
        let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
        let duration = started.elapsed();
        if workbook.len() != spec.sheet_count {
            return Err(
                "simulated source-backed XLSX sheet count differs from specification".into(),
            );
        }
        let metrics = backing.snapshot();
        let simulation = source.snapshot()?;
        verify_simulation_snapshot(&simulation, config, "simulated XLSX source open")?;
        verify_unselected_xlsx_ranges_untouched(metrics, "simulated XLSX source open")?;

        let first = workbook
            .sheet("Sheet1")?
            .ok_or("simulated source-backed XLSX first sheet is missing")?;
        verify_xlsx_source_cell(&first, "A1", 0)?;
        if backing.snapshot().xlsx.selected_worksheet.read_calls
            <= metrics.xlsx.selected_worksheet.read_calls
        {
            return Err("simulated XLSX open had already materialized the first sheet".into());
        }

        if iteration >= warmup_iterations {
            summary.record_xlsx(metrics);
            summary.record_simulation(simulation);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::XlsxRangeSourceOpen,
        corpus,
        elapsed,
        summary,
    ))
}

fn run_xlsx_range_source_list_sheets(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    config: RangeSimulationConfig,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let expected = (0..spec.sheet_count)
        .map(xlsx_sheet_name)
        .collect::<Vec<_>>();
    let mut elapsed = Vec::with_capacity(samples);
    let mut summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let backing = xlsx_instrumented_source(corpus)?;
        let source = simulated_source(backing.clone(), config);
        let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
        source.reset()?;
        let started = Instant::now();
        let names = workbook
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>();
        let duration = started.elapsed();
        if names != expected {
            return Err("simulated source-backed XLSX listing differs from specification".into());
        }
        let metrics = backing.snapshot();
        let simulation = source.snapshot()?;
        if metrics != SourceSnapshot::default() || simulation != RangeSimulationSnapshot::default()
        {
            return Err("simulated XLSX listing performed a timed source request".into());
        }

        let first = workbook
            .sheet("Sheet1")?
            .ok_or("simulated source-backed XLSX first sheet is missing")?;
        verify_xlsx_source_cell(&first, "A1", 0)?;
        if source.snapshot()?.physical_request_count == 0
            || backing.snapshot().xlsx.selected_worksheet.read_calls == 0
        {
            return Err("simulated XLSX listing had already materialized the first sheet".into());
        }

        if iteration >= warmup_iterations {
            summary.record_xlsx(metrics);
            summary.record_simulation(simulation);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::XlsxRangeSourceListSheets,
        corpus,
        elapsed,
        summary,
    ))
}

fn run_xlsx_range_source_first_cell(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    config: RangeSimulationConfig,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let backing = xlsx_instrumented_source(corpus)?;
        let source = simulated_source(backing.clone(), config);
        let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
        let first = workbook
            .sheet("Sheet1")?
            .ok_or("simulated source-backed XLSX first sheet is missing")?;
        source.reset()?;
        let started = Instant::now();
        let cell = first.cell("A1")?;
        let duration = started.elapsed();
        if !xlsx_source_cell_has_value(cell.stored(), 0) {
            return Err("simulated source-backed XLSX first cell differs from expectation".into());
        }
        let metrics = backing.snapshot();
        let simulation = source.snapshot()?;
        verify_simulation_snapshot(&simulation, config, "simulated XLSX first-cell read")?;
        if metrics.xlsx.selected_worksheet.read_calls == 0 {
            return Err(
                "simulated XLSX first-cell read touched no selected worksheet range".into(),
            );
        }
        verify_unselected_xlsx_ranges_untouched(metrics, "simulated XLSX first-cell read")?;
        prove_xlsx_unselected_sheet_deferred(&workbook, &backing, metrics, spec)?;

        if iteration >= warmup_iterations {
            summary.record_xlsx(metrics);
            summary.record_simulation(simulation);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::XlsxRangeSourceFirstCell,
        corpus,
        elapsed,
        summary,
    ))
}

fn run_xlsx_range_source_narrow_column_range_scan(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    config: RangeSimulationConfig,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let range = format!("B1:B{}", spec.row_count);
    let mut elapsed = Vec::with_capacity(samples);
    let mut summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let backing = xlsx_instrumented_source(corpus)?;
        let source = simulated_source(backing.clone(), config);
        let workbook = SourceBackedWorkbook::from_read_at(source.clone())?;
        let first = workbook
            .sheet("Sheet1")?
            .ok_or("simulated source-backed XLSX first sheet is missing")?;
        source.reset()?;
        let started = Instant::now();
        let cells = first.cells(range.as_str())?;
        let duration = started.elapsed();
        if cells.len() != spec.row_count {
            return Err("simulated XLSX narrow range count differs from specification".into());
        }
        for (row, cell) in cells.iter().enumerate() {
            let expected_address = xlsx_address(row, 1)?;
            let expected_value = xlsx_value(XlsxCoordinate {
                sheet: 0,
                row,
                column: 1,
            });
            if cell.address.to_string() != expected_address
                || !xlsx_source_cell_has_value(Some(&cell.cell), expected_value)
            {
                return Err(
                    "simulated XLSX narrow range contents differ from specification".into(),
                );
            }
        }
        let metrics = backing.snapshot();
        let simulation = source.snapshot()?;
        verify_simulation_snapshot(&simulation, config, "simulated XLSX narrow range")?;
        if metrics.xlsx.selected_worksheet.read_calls == 0 {
            return Err("simulated XLSX narrow range touched no selected worksheet range".into());
        }
        verify_unselected_xlsx_ranges_untouched(metrics, "simulated XLSX narrow range")?;
        prove_xlsx_unselected_sheet_deferred(&workbook, &backing, metrics, spec)?;

        if iteration >= warmup_iterations {
            summary.record_xlsx(metrics);
            summary.record_simulation(simulation);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::XlsxRangeSourceNarrowColumnRangeScan,
        corpus,
        elapsed,
        summary,
    ))
}

fn run_xlsx_noop_commit(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let workbook = Workbook::from_bytes(corpus.archive.clone())?;
        let edit = workbook.edit()?;
        let started = Instant::now();
        let commit = edit.commit()?;
        let duration = started.elapsed();
        if !commit.patch().is_empty() {
            return Err("XLSX no-op commit produced semantic changes".into());
        }
        verify_xlsx_cells(commit.workbook(), spec, &[])?;
        std::hint::black_box(&commit);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(Case::XlsxNoopCommit, corpus, elapsed, None))
}

fn run_xlsx_noop_commit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let maximum = xlsx_output_ceiling(corpus.archive.len())?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let workbook = Workbook::from_bytes(corpus.archive.clone())?;
        let edit = workbook.edit()?;
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let commit = edit.commit()?;
        commit.workbook().write_to(&mut sink)?;
        let duration = started.elapsed();
        if !commit.patch().is_empty() || sink.bytes != corpus.archive {
            return Err("XLSX no-op commit/save is not byte-exact for generated corpus".into());
        }
        let reopened = Workbook::from_bytes(sink.bytes.clone())?;
        verify_xlsx_cells(&reopened, spec, &[])?;
        if iteration >= warmup_iterations {
            sink_summaries.push(sink.summary());
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let sink = deterministic_sink_summary(&sink_summaries, "XLSX no-op commit/save")?;
    Ok(result(
        Case::XlsxNoopCommitSave,
        corpus,
        elapsed,
        Some(sink),
    ))
}

fn run_xlsx_update_commit(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    one_cell: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let updates = if one_cell == 1 {
        &spec.one_percent_updates[..1]
    } else {
        spec.one_percent_updates.as_slice()
    };
    let case = if one_cell == 1 {
        Case::XlsxOneCellCommit
    } else {
        Case::XlsxOnePercentCommit
    };
    let mut elapsed = Vec::with_capacity(samples);
    let mut final_commit = None;
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let workbook = Workbook::from_bytes(corpus.archive.clone())?;
        let edit = prepare_xlsx_updates(&workbook, updates)?;
        let started = Instant::now();
        let commit = edit.commit()?;
        let duration = started.elapsed();
        if commit.patch().len() != updates.len() {
            return Err("XLSX update commit has an unexpected semantic change count".into());
        }
        std::hint::black_box(&commit);
        final_commit = Some(commit);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    verify_xlsx_cells(
        final_commit
            .as_ref()
            .ok_or("XLSX update commit produced no final snapshot")?
            .workbook(),
        spec,
        updates,
    )?;
    Ok(result(case, corpus, elapsed, None))
}

fn run_xlsx_one_cell_commit_first_read(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let coordinate = *spec
        .one_percent_updates
        .first()
        .ok_or("XLSX corpus has no update coordinate")?;
    let updates = std::slice::from_ref(&coordinate);
    let sheet_name = xlsx_sheet_name(coordinate.sheet);
    let address = xlsx_address(coordinate.row, coordinate.column)?;
    let expected = (xlsx_value(coordinate) + 1).to_string();
    let mut elapsed = Vec::with_capacity(samples);
    let mut final_commit = None;
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let workbook = Workbook::from_bytes(corpus.archive.clone())?;
        let edit = prepare_xlsx_updates(&workbook, updates)?;
        let started = Instant::now();
        let commit = edit.commit()?;
        let sheet = commit
            .workbook()
            .sheet(sheet_name.as_str())?
            .ok_or("XLSX updated sheet is missing")?;
        let cell = sheet
            .cell(address.as_str())?
            .stored()
            .ok_or("XLSX updated cell is missing")?;
        let XlsxCell::Value(XlsxValue::Number(actual)) = cell else {
            return Err("XLSX updated cell is not numeric".into());
        };
        std::hint::black_box(actual.as_str());
        let duration = started.elapsed();
        if actual.as_str() != expected {
            return Err("XLSX updated cell differs from deterministic expectation".into());
        }
        if commit.patch().len() != 1 {
            return Err("XLSX update commit has an unexpected semantic change count".into());
        }
        final_commit = Some(commit);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    verify_xlsx_cells(
        final_commit
            .as_ref()
            .ok_or("XLSX commit/read produced no final snapshot")?
            .workbook(),
        spec,
        updates,
    )?;
    Ok(result(
        Case::XlsxOneCellCommitFirstRead,
        corpus,
        elapsed,
        None,
    ))
}

fn run_xlsx_update_commit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    one_cell: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let updates = if one_cell == 1 {
        &spec.one_percent_updates[..1]
    } else {
        spec.one_percent_updates.as_slice()
    };
    let case = if one_cell == 1 {
        Case::XlsxOneCellCommitSave
    } else {
        Case::XlsxOnePercentCommitSave
    };
    let expected = xlsx_expected_output(corpus, updates)?;
    let maximum = xlsx_output_ceiling(expected.len())?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let workbook = Workbook::from_bytes(corpus.archive.clone())?;
        let edit = prepare_xlsx_updates(&workbook, updates)?;
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let commit = edit.commit()?;
        commit.workbook().write_to(&mut sink)?;
        let duration = started.elapsed();
        if sink.bytes != expected {
            return Err(
                "XLSX changed commit/save differs from deterministic expected output".into(),
            );
        }
        let reopened = Workbook::from_bytes(sink.bytes.clone())?;
        verify_xlsx_cells(&reopened, spec, updates)?;
        if iteration >= warmup_iterations {
            sink_summaries.push(sink.summary());
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let sink = deterministic_sink_summary(&sink_summaries, "XLSX changed commit/save")?;
    Ok(result(case, corpus, elapsed, Some(sink)))
}

fn xlsx_cell_crud_eager_output(
    corpus: &Corpus,
    updates: &[XlsxCoordinate],
) -> Result<Vec<u8>, Box<dyn Error>> {
    let workbook = Workbook::from_bytes(corpus.archive.clone())?;
    let mut edit = workbook.edit()?;
    for coordinate in updates {
        let mut sheet = edit
            .sheet(xlsx_sheet_name(coordinate.sheet))?
            .ok_or("XLSX cell CRUD eager target sheet is missing")?;
        sheet.set(
            xlsx_address(coordinate.row, coordinate.column)?,
            xlsx_value(*coordinate) + 1,
        )?;
    }
    let commit = edit.commit()?;
    let output = commit.workbook().to_bytes()?;
    verify_xlsx_cell_crud_output(corpus, &output, updates)?;
    Ok(output)
}

fn run_xlsx_cell_values_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_CELL_VALUES_SOURCE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_cell_values_edit_save()
    {
        return Err("XLSX cell CRUD case requires its fixed multi-sheet media corpus".into());
    }
    let spec = xlsx_spec(corpus)?;
    let updates = xlsx_cell_crud_updates_for_case(case, spec)?;
    run_xlsx_cell_value_lifecycle_gates(corpus, spec)?;
    let expected = xlsx_cell_crud_eager_output(corpus, &updates)?;
    let expected_digest = sha256_hex(&expected);
    let source_backed = matches!(
        case,
        Case::XlsxSourceBackedCellValuesOneEditSave
            | Case::XlsxSourceBackedCellValuesOnePercentEditSave
            | Case::XlsxSourceBackedCellValuesBatchEditSave
    );
    let expected_touched = xlsx_update_sheet_selectors(&updates).len();
    let maximum = xlsx_output_ceiling(expected.len())?;
    let payload_ranges = source_backed
        .then(|| xlsx_cell_crud_payload_ranges(corpus))
        .transpose()?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut output_digests = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let backing = source_backed.then(|| {
            Arc::new(InstrumentedSource::new(
                corpus.archive.clone(),
                payload_ranges
                    .clone()
                    .expect("source CRUD ranges are present"),
            ))
        });
        let eager_source = (!source_backed).then(|| corpus.archive.clone());
        let mut duration = Duration::ZERO;
        let mut source_metrics = None;
        let mut materializations = 0u64;
        if source_backed {
            let read_at: Arc<dyn ReadAt> = backing.clone().expect("source CRUD backing exists");
            let started = Instant::now();
            let editor = litchi_xlsx::cell_values::SourceBackedEditor::from_read_at(read_at)?;
            let selectors = xlsx_update_sheet_selectors(&updates);
            let mut edit = editor.edit_sheets(selectors)?;
            for coordinate in &updates {
                edit.set(
                    coordinate.sheet,
                    litchi_xlsx::Address::at(
                        u32::try_from(coordinate.row)?,
                        u32::try_from(coordinate.column)?,
                    )?,
                    xlsx_value(*coordinate) + 1,
                )?;
            }
            let commit = edit.commit()?;
            duration += started.elapsed();
            if commit.diagnostics().touched_worksheets() != expected_touched {
                return Err("XLSX source CRUD touched an unexpected worksheet count".into());
            }
            materializations = editor.cache_diagnostics().successful_loads;
            let publish_started = Instant::now();
            editor.publish_multi_commit_to_stream(&mut sink, &commit)?;
            duration += publish_started.elapsed();
            source_metrics = Some(
                backing
                    .clone()
                    .expect("source CRUD backing exists")
                    .snapshot(),
            );
        } else {
            let started = Instant::now();
            let workbook = Workbook::from_bytes(eager_source.expect("eager source exists"))?;
            let mut edit = workbook.edit()?;
            for coordinate in &updates {
                edit.sheet(xlsx_sheet_name(coordinate.sheet))?
                    .ok_or("XLSX cell CRUD eager target sheet is missing")?
                    .set(
                        xlsx_address(coordinate.row, coordinate.column)?,
                        xlsx_value(*coordinate) + 1,
                    )?;
            }
            let commit = edit.commit()?;
            duration += started.elapsed();
            let started = Instant::now();
            commit.workbook().write_to(&mut sink)?;
            duration += started.elapsed();
        }
        if source_backed && materializations < expected_touched as u64 {
            return Err("XLSX source CRUD materialized fewer worksheets than touched".into());
        }
        if sink.bytes.is_empty() {
            return Err("XLSX cell CRUD publication emitted no bytes".into());
        }
        verify_xlsx_cell_crud_output(corpus, &sink.bytes, &updates)?;
        if source_backed {
            verify_xlsx_cell_crud_raw_source_output(corpus, &sink.bytes, &updates)?;
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("XLSX cell CRUD publication exceeded sequential sink bound".into());
        }
        let digest = sha256_hex(&sink.bytes);
        if iteration >= warmup_iterations {
            if let Some(metrics) = source_metrics {
                source_summary.record_opc(metrics, materializations);
            }
            sink_summaries.push(sink.summary());
            output_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if output_digests.windows(2).any(|pair| pair[0] != pair[1]) {
        return Err("XLSX cell CRUD output hashes are not deterministic".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: source_backed.then_some(source_summary),
        execution: None,
        output_sha256: output_digests.first().cloned().or(Some(expected_digest)),
    })
}

fn prepare_xlsx_merge_edit(
    workbook: &Workbook,
    merge: bool,
) -> Result<litchi_xlsx::Edit, Box<dyn Error>> {
    let mut edit = workbook.edit()?;
    let mut sheet = edit
        .sheet("Sheet1")?
        .ok_or("XLSX merge fixture is missing Sheet1")?;
    if merge {
        sheet.merge("A1:B2")?;
    } else {
        sheet.unmerge("B2")?;
    }
    Ok(edit)
}

fn verify_xlsx_merge_output(output: &[u8], merged: bool) -> Result<(), Box<dyn Error>> {
    let workbook = Workbook::from_bytes(output.to_vec())?;
    let sheet = workbook
        .sheet("Sheet1")?
        .ok_or("XLSX merge output is missing Sheet1")?;
    let ranges = sheet.merges()?.map(Rect::a1).collect::<Vec<_>>();
    let expected_ranges = if merged { vec!["A1:B2"] } else { Vec::new() };
    if ranges != expected_ranges {
        return Err("XLSX merge output has unexpected merge membership".into());
    }

    let anchor = sheet.cell("A1")?;
    if !matches!(
        anchor,
        litchi_xlsx::cell::View::Stored(
            XlsxCell::Value(XlsxValue::Text(text))
        ) if text.as_str() == "litchi-xlsx-merge-anchor-v1"
    ) {
        return Err("XLSX merge output did not retain the anchor cell".into());
    }
    let unrelated = sheet.cell("C1")?;
    if !matches!(
        unrelated,
        litchi_xlsx::cell::View::Stored(
            XlsxCell::Value(XlsxValue::Text(text))
        ) if text.as_str() == "litchi-xlsx-merge-unrelated-v1"
    ) {
        return Err("XLSX merge output changed an unrelated cell".into());
    }

    for address in ["A2", "B1", "B2"] {
        let view = sheet.cell(address)?;
        if merged {
            if view.merge().map(Rect::a1).as_deref() != Some("A1:B2") {
                return Err(format!("XLSX merge output did not cover {address}").into());
            }
        } else if !view.is_missing() {
            return Err(format!("XLSX unmerge output retained {address}").into());
        }
    }
    if !sheet.cell("C2")?.is_missing() {
        return Err("XLSX merge output changed an uncovered cell".into());
    }
    Ok(())
}

fn verify_xlsx_merge_durable_patch(
    source: &Workbook,
    commit: &litchi_xlsx::Commit,
    expected: &[u8],
) -> Result<(), Box<dyn Error>> {
    if commit.patch().is_empty() || commit.patch().len() != 1 {
        return Err("XLSX merge edit produced an unexpected semantic patch".into());
    }
    let durable = commit.patch().durable()?;
    let first_json = durable.to_deterministic_json()?;
    if first_json != durable.to_deterministic_json()? {
        return Err("XLSX merge durable patch encoding is not deterministic".into());
    }
    let parsed = litchi_xlsx::DurablePatch::from_deterministic_json(&first_json)?;
    let applied = parsed.apply(source)?;
    if applied.to_plain_bytes()? != expected {
        return Err("XLSX merge durable patch did not reproduce the expected output".into());
    }
    let restored = parsed.inverse().apply(&applied)?;
    if restored.to_plain_bytes()? != source.to_plain_bytes()? {
        return Err("XLSX merge durable inverse did not restore the exact source".into());
    }

    let mut stale_edit = source.edit()?;
    stale_edit
        .sheet("Sheet1")?
        .ok_or("XLSX merge fixture is missing Sheet1")?
        .set("C1", "litchi-xlsx-merge-stale-v1")?;
    let stale = stale_edit.commit()?.into_workbook();
    if parsed.apply(&stale).is_ok() {
        return Err("XLSX merge durable patch accepted a stale source".into());
    }
    Ok(())
}

fn run_xlsx_merge_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_MERGE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_merge_edit_save()
    {
        return Err("XLSX merge case requires its sparse A1:B2 corpus".into());
    }
    let merge = case == Case::XlsxEagerMergeCommitSave;
    let (unmerged, merged) = xlsx_merge_fixture()?;
    let source = if merge { unmerged } else { merged };
    let source_bytes = source.to_bytes()?;
    if source_bytes != corpus.archive {
        return Err("XLSX merge corpus source bytes differ from its fixture".into());
    }
    let expected_commit = prepare_xlsx_merge_edit(&source, merge)?.commit()?;
    let expected = expected_commit.workbook().to_bytes()?;
    verify_xlsx_merge_output(&expected, merge)?;
    verify_xlsx_merge_durable_patch(&source, &expected_commit, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = xlsx_output_ceiling(expected.len())?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let workbook = Workbook::from_bytes(source_bytes.clone())?;
        let edit = prepare_xlsx_merge_edit(&workbook, merge)?;
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let commit = edit.commit()?;
        commit.workbook().write_to(&mut sink)?;
        let duration = started.elapsed();
        if sink.bytes != expected {
            return Err("XLSX merge commit/save differs from deterministic output".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("XLSX merge save exceeded the sequential sink write bound".into());
        }
        verify_xlsx_merge_output(&sink.bytes, merge)?;
        if sha256_hex(&sink.bytes) != expected_digest {
            return Err("XLSX merge output digest differs from expected output".into());
        }
        if iteration >= warmup_iterations {
            sink_summaries.push(sink.summary());
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: None,
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn xlsx_expected_output(
    corpus: &Corpus,
    updates: &[XlsxCoordinate],
) -> Result<Vec<u8>, Box<dyn Error>> {
    let workbook = Workbook::from_bytes(corpus.archive.clone())?;
    let commit = prepare_xlsx_updates(&workbook, updates)?.commit()?;
    Ok(commit.workbook().to_bytes()?)
}

fn deterministic_sink_summary(
    summaries: &[SinkSummary],
    context: &str,
) -> Result<SinkSummary, Box<dyn Error>> {
    let first = *summaries
        .first()
        .ok_or("XLSX save produced no measured sink summary")?;
    if summaries.iter().any(|summary| *summary != first) {
        return Err(format!("{context} produced differing sink summaries").into());
    }
    Ok(first)
}

fn run_zip_index(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let archive = ArchiveReader::new(&corpus.archive)?;
        let observed = archive.file_names().count();
        std::hint::black_box(observed);
        record_elapsed(
            &mut elapsed,
            iteration,
            warmup_iterations,
            started.elapsed(),
        )?;
    }
    Ok(result(Case::ZipIndex, corpus, elapsed, None))
}

fn run_zip_read_one(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let archive = ArchiveReader::new(&corpus.archive)?;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let bytes = archive.read(&corpus.target_name)?;
        if bytes != corpus.target_payload {
            return Err("ZIP read result differs from deterministic corpus payload".into());
        }
        std::hint::black_box(&bytes);
        record_elapsed(
            &mut elapsed,
            iteration,
            warmup_iterations,
            started.elapsed(),
        )?;
    }
    Ok(result(Case::ZipReadOne, corpus, elapsed, None))
}

fn opc_payload_ranges(corpus: &Corpus) -> Result<(Vec<Range<u64>>, Range<u64>), Box<dyn Error>> {
    let archive = soapberry_zip::ZipArchive::from_slice(&corpus.archive)?;
    let mut ordinary = Vec::with_capacity(corpus.manifest.entry_count);
    let mut target = None;
    for header in archive.entries() {
        let header = header?;
        let path = header.file_path().try_normalize()?;
        let name = path.as_ref();
        if !name.starts_with("benchmark/parts/") {
            continue;
        }
        let entry = archive.get_entry(header.wayfinder())?;
        let (start, end) = entry.compressed_data_range();
        let range = start..end;
        if name == corpus.target_name {
            target = Some(range.clone());
        }
        ordinary.push(range);
    }
    if ordinary.len() != corpus.manifest.entry_count {
        return Err("OPC payload range count differs from corpus manifest".into());
    }
    Ok((
        ordinary,
        target.ok_or("OPC main target has no compressed source range")?,
    ))
}

fn zip_member_ranges(bytes: &[u8]) -> Result<Vec<(String, Range<u64>)>, Box<dyn Error>> {
    let archive = soapberry_zip::ZipArchive::from_slice(bytes)?;
    let mut members = Vec::new();
    for header in archive.entries() {
        let header = header?;
        let name = header.file_path().try_normalize()?.as_ref().to_owned();
        let entry = archive.get_entry(header.wayfinder())?;
        let (start, end) = entry.compressed_data_range();
        members.push((name, start..end));
    }
    Ok(members)
}

fn docx_source_payload_ranges(
    corpus: &Corpus,
) -> Result<(Vec<Range<u64>>, Range<u64>), Box<dyn Error>> {
    let mut ordinary = Vec::with_capacity(corpus.manifest.entry_count);
    let mut target = None;
    for (name, range) in zip_member_ranges(&corpus.archive)? {
        if name == "[Content_Types].xml" || name.ends_with(".rels") {
            continue;
        }
        if name == corpus.target_name {
            target = Some(range.clone());
        }
        ordinary.push(range);
    }
    if ordinary.len() != corpus.manifest.entry_count {
        return Err("DOCX source-edit payload count differs from corpus manifest".into());
    }
    Ok((
        ordinary,
        target.ok_or("DOCX source-edit main document has no compressed source range")?,
    ))
}

fn pptx_source_payload_ranges(corpus: &Corpus) -> Result<Vec<Range<u64>>, Box<dyn Error>> {
    let ordinary = zip_member_ranges(&corpus.archive)?
        .into_iter()
        .filter_map(|(name, range)| {
            (name != "[Content_Types].xml" && !name.ends_with(".rels")).then_some(range)
        })
        .collect::<Vec<_>>();
    if ordinary.len() != corpus.manifest.entry_count {
        return Err("PPTX source-edit payload count differs from corpus manifest".into());
    }
    Ok(ordinary)
}

fn xlsx_calculation_payload_ranges(corpus: &Corpus) -> Result<Vec<Range<u64>>, Box<dyn Error>> {
    let ordinary = zip_member_ranges(&corpus.archive)?
        .into_iter()
        .filter_map(|(name, range)| {
            (name != "[Content_Types].xml" && !name.ends_with(".rels")).then_some(range)
        })
        .collect::<Vec<_>>();
    if ordinary.len() != corpus.manifest.entry_count {
        return Err("XLSX calculation-edit payload count differs from corpus manifest".into());
    }
    Ok(ordinary)
}

fn xlsx_cell_crud_payload_ranges(corpus: &Corpus) -> Result<Vec<Range<u64>>, Box<dyn Error>> {
    Ok(zip_member_ranges(&corpus.archive)?
        .into_iter()
        .filter_map(|(name, range)| {
            (name != "[Content_Types].xml" && !name.ends_with(".rels")).then_some(range)
        })
        .collect())
}

fn xlsx_source_layout(
    bytes: &[u8],
    expected_sheet_count: usize,
) -> Result<(XlsxTrackedRanges, XlsxSourceMembersManifest), Box<dyn Error>> {
    let members = zip_member_ranges(bytes)?;
    let mut ranges = XlsxTrackedRanges::default();
    let mut workbook = None;
    let mut worksheets = Vec::new();
    let mut shared_strings = None;
    let mut styles = None;

    for (name, range) in members {
        match name.as_str() {
            "xl/workbook.xml" => {
                workbook = Some(name);
                ranges.workbook.push(range);
            },
            "xl/sharedStrings.xml" => {
                shared_strings = Some(name);
                ranges.shared_strings.push(range);
            },
            "xl/styles.xml" => {
                styles = Some(name);
                ranges.styles.push(range);
            },
            "xl/worksheets/sheet1.xml" => {
                worksheets.push(name);
                ranges.selected_worksheet.push(range);
            },
            _ if name.starts_with("xl/worksheets/") && name.ends_with(".xml") => {
                worksheets.push(name);
                ranges.unselected_worksheets.push(range);
            },
            _ => {},
        }
    }
    worksheets.sort();
    if ranges.workbook.len() != 1 {
        return Err("XLSX archive does not contain exactly one workbook member".into());
    }
    if ranges.selected_worksheet.len() != 1 || worksheets.len() != expected_sheet_count {
        return Err("XLSX worksheet member count differs from corpus specification".into());
    }

    Ok((
        ranges,
        XlsxSourceMembersManifest {
            workbook: workbook.ok_or("XLSX workbook member is missing")?,
            worksheets,
            shared_strings,
            styles,
        },
    ))
}

fn xlsx_instrumented_source(corpus: &Corpus) -> Result<Arc<InstrumentedSource>, Box<dyn Error>> {
    let spec = xlsx_spec(corpus)?;
    let (ranges, _manifest) = xlsx_source_layout(&corpus.archive, spec.sheet_count)?;
    let mut ordinary = Vec::new();
    ordinary.extend(ranges.workbook.iter().cloned());
    ordinary.extend(ranges.selected_worksheet.iter().cloned());
    ordinary.extend(ranges.unselected_worksheets.iter().cloned());
    ordinary.extend(ranges.shared_strings.iter().cloned());
    ordinary.extend(ranges.styles.iter().cloned());
    Ok(Arc::new(InstrumentedSource::new_xlsx(
        corpus.archive.clone(),
        ordinary,
        ranges,
    )))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum OpcCacheMode {
    Control,
    Managed,
}

impl OpcCacheMode {
    const fn name(self) -> &'static str {
        match self {
            Self::Control => "finite-control",
            Self::Managed => "budget-managed",
        }
    }

    const fn is_managed(self) -> bool {
        matches!(self, Self::Managed)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum OpcCacheScenario {
    SamePart,
    DisjointParts,
}

impl OpcCacheScenario {
    const fn name(self) -> &'static str {
        match self {
            Self::SamePart => "same-part",
            Self::DisjointParts => "disjoint-parts",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum OpcCacheCapacity {
    Half,
    One,
    Two,
}

impl OpcCacheCapacity {
    const ALL: [Self; 3] = [Self::Half, Self::One, Self::Two];

    const fn name(self) -> &'static str {
        match self {
            Self::Half => "1/2x",
            Self::One => "1x",
            Self::Two => "2x",
        }
    }

    const fn parts(self, working_set_parts: usize) -> usize {
        match self {
            Self::Half => working_set_parts / 2,
            Self::One => working_set_parts,
            Self::Two => working_set_parts * 2,
        }
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct OpcCacheCounterDelta {
    hits: u64,
    cold_loads: u64,
    waiter_joins: u64,
    successful_loads: u64,
    failed_loads: u64,
    evictions: u64,
    bypasses: u64,
    oversized_bypasses: u64,
    allocation_bypasses: u64,
    budget_reservation_failures: u64,
}

fn opc_cache_counter_delta(
    before: SourceCacheDiagnostics,
    after: SourceCacheDiagnostics,
) -> Result<OpcCacheCounterDelta, Box<dyn Error>> {
    let subtract = |name: &str, after: u64, before: u64| -> Result<u64, Box<dyn Error>> {
        after
            .checked_sub(before)
            .ok_or_else(|| format!("OPC cache counter {name} moved backwards").into())
    };
    Ok(OpcCacheCounterDelta {
        hits: subtract("hits", after.hits, before.hits)?,
        cold_loads: subtract("cold_loads", after.cold_loads, before.cold_loads)?,
        waiter_joins: subtract("waiter_joins", after.waiter_joins, before.waiter_joins)?,
        successful_loads: subtract(
            "successful_loads",
            after.successful_loads,
            before.successful_loads,
        )?,
        failed_loads: subtract("failed_loads", after.failed_loads, before.failed_loads)?,
        evictions: subtract("evictions", after.evictions, before.evictions)?,
        bypasses: subtract("bypasses", after.bypasses, before.bypasses)?,
        oversized_bypasses: subtract(
            "oversized_bypasses",
            after.oversized_bypasses,
            before.oversized_bypasses,
        )?,
        allocation_bypasses: subtract(
            "allocation_bypasses",
            after.allocation_bypasses,
            before.allocation_bypasses,
        )?,
        budget_reservation_failures: subtract(
            "budget_reservation_failures",
            after.budget_reservation_failures,
            before.budget_reservation_failures,
        )?,
    })
}

impl OpcCacheDiagnosticsSummary {
    fn record(&mut self, delta: OpcCacheCounterDelta, current: SourceCacheDiagnostics) {
        self.hits.push(delta.hits);
        self.cold_loads.push(delta.cold_loads);
        self.waiter_joins.push(delta.waiter_joins);
        self.successful_loads.push(delta.successful_loads);
        self.failed_loads.push(delta.failed_loads);
        self.evictions.push(delta.evictions);
        self.bypasses.push(delta.bypasses);
        self.oversized_bypasses.push(delta.oversized_bypasses);
        self.allocation_bypasses.push(delta.allocation_bypasses);
        self.budget_reservation_failures
            .push(delta.budget_reservation_failures);
        self.retained_entries.push(current.retained_entries);
        self.retained_bytes.push(current.retained_bytes);
        self.in_flight_loads.push(current.in_flight_loads);
        self.budget_memory_used.push(current.budget_memory_used);
        self.budget_cache_reserved_bytes
            .push(current.budget_cache_reserved_bytes);
        self.budget_memory_limit.push(current.budget_memory_limit);
    }
}

#[derive(Debug, Default)]
struct OpcCacheGateState {
    armed: bool,
    released: bool,
    expected_initial: usize,
    initial_arrivals: u64,
    initial_started: usize,
    delayed_payload_arrivals: u64,
    seen_ranges: BTreeSet<usize>,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
struct OpcCacheGateSnapshot {
    initial_arrivals: u64,
    delayed_payload_arrivals: u64,
    max_concurrent_delays: u64,
}

#[derive(Debug)]
struct CoordinatedSlowSource {
    backing: Arc<InstrumentedSource>,
    payload_ranges: Vec<Range<u64>>,
    delay: Duration,
    state: Mutex<OpcCacheGateState>,
    changed: Condvar,
    current_delays: AtomicU64,
    max_concurrent_delays: AtomicU64,
}

impl CoordinatedSlowSource {
    fn new(
        backing: Arc<InstrumentedSource>,
        payload_ranges: Vec<Range<u64>>,
        delay: Duration,
    ) -> Self {
        Self {
            backing,
            payload_ranges,
            delay,
            state: Mutex::new(OpcCacheGateState::default()),
            changed: Condvar::new(),
            current_delays: AtomicU64::new(0),
            max_concurrent_delays: AtomicU64::new(0),
        }
    }

    fn arm(&self, expected_initial: usize) -> Result<(), Box<dyn Error>> {
        if expected_initial == 0 {
            return Err("OPC cache slow-source gate requires an initial arrival".into());
        }
        let mut state = self
            .state
            .lock()
            .map_err(|_error| "OPC cache slow-source gate is poisoned")?;
        if state.armed {
            return Err("OPC cache slow-source gate was armed twice".into());
        }
        state.armed = true;
        state.expected_initial = expected_initial;
        Ok(())
    }

    fn wait_for_initial_arrivals(&self, timeout: Duration) -> Result<(), Box<dyn Error>> {
        let deadline = Instant::now() + timeout;
        let mut state = self
            .state
            .lock()
            .map_err(|_error| "OPC cache slow-source gate is poisoned")?;
        while state.initial_arrivals < state.expected_initial as u64 {
            let now = Instant::now();
            if now >= deadline {
                return Err(format!(
                    "OPC cache slow-source gate observed {} of {} initial arrivals",
                    state.initial_arrivals, state.expected_initial
                )
                .into());
            }
            let remaining = deadline.saturating_duration_since(now);
            let (next, _) = self
                .changed
                .wait_timeout(state, remaining)
                .map_err(|_error| "OPC cache slow-source gate is poisoned")?;
            state = next;
        }
        Ok(())
    }

    fn release(&self) {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        state.released = true;
        self.changed.notify_all();
    }

    fn snapshot(&self) -> Result<OpcCacheGateSnapshot, Box<dyn Error>> {
        let state = self
            .state
            .lock()
            .map_err(|_error| "OPC cache slow-source gate is poisoned")?;
        Ok(OpcCacheGateSnapshot {
            initial_arrivals: state.initial_arrivals,
            delayed_payload_arrivals: state.delayed_payload_arrivals,
            max_concurrent_delays: self.max_concurrent_delays.load(Ordering::SeqCst),
        })
    }

    fn payload_range_index(&self, offset: u64, requested: usize) -> Option<usize> {
        let end = offset.saturating_add(u64::try_from(requested).unwrap_or(u64::MAX));
        self.payload_ranges
            .iter()
            .position(|range| offset < range.end && range.start < end)
    }
}

impl ReadAt for CoordinatedSlowSource {
    fn len(&self) -> io::Result<u64> {
        self.backing.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let range_index = self.payload_range_index(offset, output.len());
        let mut initial = false;
        let delayed = if let Some(range_index) = range_index {
            let mut state = self
                .state
                .lock()
                .map_err(|_error| io::Error::other("OPC cache slow-source gate is poisoned"))?;
            if state.armed && state.seen_ranges.insert(range_index) {
                state.delayed_payload_arrivals = state.delayed_payload_arrivals.saturating_add(1);
                if !state.released {
                    initial = true;
                    state.initial_arrivals = state.initial_arrivals.saturating_add(1);
                    self.changed.notify_all();
                    while !state.released {
                        state = self.changed.wait(state).map_err(|_error| {
                            io::Error::other("OPC cache slow-source gate is poisoned")
                        })?;
                    }
                }
                true
            } else {
                false
            }
        } else {
            false
        };

        if delayed {
            let current = self.current_delays.fetch_add(1, Ordering::SeqCst) + 1;
            self.max_concurrent_delays
                .fetch_max(current, Ordering::SeqCst);
            let _guard = InFlightReadGuard {
                counter: &self.current_delays,
            };
            if initial {
                let mut state = self
                    .state
                    .lock()
                    .map_err(|_error| io::Error::other("OPC cache slow-source gate is poisoned"))?;
                state.initial_started = state.initial_started.saturating_add(1);
                self.changed.notify_all();
                while state.initial_started < state.expected_initial {
                    state = self.changed.wait(state).map_err(|_error| {
                        io::Error::other("OPC cache slow-source gate is poisoned")
                    })?;
                }
                self.changed.notify_all();
            }
            std::thread::sleep(self.delay);
        }
        self.backing.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.backing.version()
    }
}

#[derive(Clone)]
struct OpcCacheRequest {
    logical_index: usize,
    partname: PackURI,
}

struct OpcCacheWorkerJob {
    package: Arc<SourceBackedPackage>,
    start: Arc<Barrier>,
    requests: Vec<OpcCacheRequest>,
}

enum OpcCacheWorkerCommand {
    Run(OpcCacheWorkerJob),
    Stop,
}

type OpcCacheWorkerResult = std::result::Result<Vec<(usize, PartData)>, String>;

struct OpcCacheWorkerTeam {
    senders: Vec<mpsc::Sender<OpcCacheWorkerCommand>>,
    results: mpsc::Receiver<OpcCacheWorkerResult>,
    workers: Vec<JoinHandle<()>>,
}

#[derive(Clone, Copy, Debug)]
struct OpcCacheCellConfig {
    case: Case,
    mode: OpcCacheMode,
    scenario: OpcCacheScenario,
    capacity: OpcCacheCapacity,
    working_set_parts: usize,
    worker_count: usize,
    warmup_iterations: usize,
    samples: usize,
}

impl OpcCacheWorkerTeam {
    fn new(worker_count: usize) -> Result<Self, Box<dyn Error>> {
        if worker_count == 0 {
            return Err("OPC cache worker team cannot be empty".into());
        }
        let (result_sender, results) = mpsc::channel();
        let mut senders = Vec::with_capacity(worker_count);
        let mut workers = Vec::with_capacity(worker_count);
        for index in 0..worker_count {
            let (sender, receiver) = mpsc::channel();
            let result_sender = result_sender.clone();
            let worker = std::thread::Builder::new()
                .name(format!("litchi-opc-cache-{index}"))
                .spawn(move || {
                    while let Ok(command) = receiver.recv() {
                        let OpcCacheWorkerCommand::Run(job) = command else {
                            break;
                        };
                        let OpcCacheWorkerJob {
                            package,
                            start,
                            requests,
                        } = job;
                        start.wait();
                        let result: OpcCacheWorkerResult = requests
                            .into_iter()
                            .map(|request| {
                                package
                                    .part(&request.partname)
                                    .and_then(|part| part.data())
                                    .map(|data| (request.logical_index, data))
                                    .map_err(|error| error.to_string())
                            })
                            .collect();
                        drop(package);
                        if result_sender.send(result).is_err() {
                            break;
                        }
                    }
                })?;
            senders.push(sender);
            workers.push(worker);
        }
        Ok(Self {
            senders,
            results,
            workers,
        })
    }

    fn dispatch(
        &self,
        package: Arc<SourceBackedPackage>,
        assignments: Vec<Vec<OpcCacheRequest>>,
    ) -> Result<(), Box<dyn Error>> {
        if assignments.len() != self.senders.len()
            || assignments.iter().any(|requests| requests.is_empty())
        {
            return Err("OPC cache worker assignments must be nonempty and match the team".into());
        }
        let start = Arc::new(Barrier::new(self.senders.len() + 1));
        for (sender, requests) in self.senders.iter().zip(assignments) {
            sender
                .send(OpcCacheWorkerCommand::Run(OpcCacheWorkerJob {
                    package: Arc::clone(&package),
                    start: Arc::clone(&start),
                    requests,
                }))
                .map_err(|_error| "OPC cache worker stopped before dispatch")?;
        }
        start.wait();
        Ok(())
    }

    fn collect(&self) -> Result<Vec<(usize, PartData)>, Box<dyn Error>> {
        let mut collected = Vec::new();
        for _ in 0..self.senders.len() {
            let worker = self.results.recv()?;
            collected.extend(worker.map_err(|error| format!("OPC cache worker failed: {error}"))?);
        }
        Ok(collected)
    }
}

impl Drop for OpcCacheWorkerTeam {
    fn drop(&mut self) {
        for sender in &self.senders {
            let _ = sender.send(OpcCacheWorkerCommand::Stop);
        }
        for worker in std::mem::take(&mut self.workers) {
            let _ = worker.join();
        }
    }
}

fn validate_opc_cache_corpus(corpus: &Corpus) -> Result<(), Box<dyn Error>> {
    if corpus.manifest.generator != OPC_CORPUS_GENERATOR
        || corpus.manifest.shape != CorpusShape::ManySmall.name()
        || corpus.manifest.payload_kind != PayloadKind::Incompressible.name()
        || corpus.manifest.entry_bytes != CorpusShape::ManySmall.entry_bytes()
        || corpus.manifest.entry_count != CorpusShape::ManySmall.entry_count()
        || sha256_hex(&corpus.archive) != corpus.manifest.archive_sha256
    {
        return Err(
            "OPC cache evidence requires its fixed many-small incompressible corpus".into(),
        );
    }
    Ok(())
}

fn opc_cache_parts_and_ranges(
    corpus: &Corpus,
) -> Result<(Vec<PackURI>, Vec<Range<u64>>), Box<dyn Error>> {
    let mut by_name = zip_member_ranges(&corpus.archive)?
        .into_iter()
        .collect::<BTreeMap<_, _>>();
    let mut parts = Vec::with_capacity(corpus.manifest.entry_count);
    let mut ranges = Vec::with_capacity(corpus.manifest.entry_count);
    for index in 0..corpus.manifest.entry_count {
        let name = entry_name(index);
        parts.push(PackURI::new(format!("/{name}"))?);
        ranges.push(
            by_name
                .remove(&name)
                .ok_or("OPC cache evidence Part has no compressed source range")?,
        );
    }
    if parts.len() != corpus.manifest.entry_count {
        return Err("OPC cache evidence Part count differs from corpus manifest".into());
    }
    Ok((parts, ranges))
}

fn opc_cache_context(
    memory_limit: u64,
    worker_count: usize,
) -> Result<(Budget, ExecutionContext), Box<dyn Error>> {
    let workers =
        NonZeroUsize::new(worker_count).ok_or("OPC cache worker count must be nonzero")?;
    let in_flight = NonZeroU64::new(memory_limit.max(1))
        .ok_or("OPC cache in-flight byte limit must be nonzero")?;
    let execution_limits = ExecutionLimits::new(workers, workers, in_flight, 0)?;
    let budget = Budget::root(
        "litchi-perf-opc-source-cache",
        Limits::new(
            memory_limit,
            64 * 1024 * 1024,
            64 * 1024 * 1024,
            1_000_000,
            256,
            1_000_000_000,
        ),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    Ok((budget, context))
}

fn opc_cache_package(
    mode: OpcCacheMode,
    source: Arc<dyn ReadAt>,
    cache_limits: SourceCacheLimits,
    memory_limit: u64,
    worker_count: usize,
) -> Result<(Arc<SourceBackedPackage>, Option<Budget>), Box<dyn Error>> {
    match mode {
        OpcCacheMode::Control => Ok((
            Arc::new(
                SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
                    source,
                    ReadLimits::default(),
                    cache_limits,
                )?,
            ),
            None,
        )),
        OpcCacheMode::Managed => {
            let (budget, context) = opc_cache_context(memory_limit, worker_count)?;
            let package =
                SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                    source,
                    ReadLimits::default(),
                    cache_limits,
                    context,
                )?;
            Ok((Arc::new(package), Some(budget)))
        },
    }
}

fn opc_cache_budget_used(budget: Option<&Budget>) -> u64 {
    budget.map_or(0, |budget| budget.used(Resource::Memory))
}

fn opc_cache_assert_diagnostics_mode(
    mode: OpcCacheMode,
    diagnostics: SourceCacheDiagnostics,
    memory_limit: u64,
) -> Result<(), Box<dyn Error>> {
    if diagnostics.budget_managed != mode.is_managed()
        || diagnostics.budget_memory_limit != mode.is_managed().then_some(memory_limit)
        || (!mode.is_managed()
            && (diagnostics.budget_memory_used != 0
                || diagnostics.budget_cache_reserved_bytes != 0
                || diagnostics.budget_reservation_failures != 0))
    {
        return Err("OPC cache diagnostics disagree with the selected cache mode".into());
    }
    Ok(())
}

fn opc_cache_expected_delta(
    scenario: OpcCacheScenario,
    capacity: OpcCacheCapacity,
    worker_count: usize,
    working_set_parts: usize,
    capacity_parts: usize,
) -> Result<OpcCacheCounterDelta, Box<dyn Error>> {
    Ok(match scenario {
        OpcCacheScenario::SamePart => OpcCacheCounterDelta {
            cold_loads: 1,
            waiter_joins: u64::try_from(worker_count.saturating_sub(1))?,
            successful_loads: 1,
            evictions: u64::from(matches!(capacity, OpcCacheCapacity::One)),
            bypasses: u64::from(matches!(capacity, OpcCacheCapacity::Half)),
            oversized_bypasses: u64::from(matches!(capacity, OpcCacheCapacity::Half)),
            ..OpcCacheCounterDelta::default()
        },
        OpcCacheScenario::DisjointParts => {
            let bypasses = working_set_parts.saturating_sub(capacity_parts);
            let evictions = match capacity {
                OpcCacheCapacity::Half => capacity_parts,
                OpcCacheCapacity::One => working_set_parts,
                OpcCacheCapacity::Two => 0,
            };
            OpcCacheCounterDelta {
                cold_loads: u64::try_from(working_set_parts)?,
                successful_loads: u64::try_from(working_set_parts)?,
                evictions: u64::try_from(evictions)?,
                bypasses: u64::try_from(bypasses)?,
                ..OpcCacheCounterDelta::default()
            }
        },
    })
}

fn opc_cache_wait_for_cohort(
    package: &SourceBackedPackage,
    expected_flights: usize,
    expected_waiters: u64,
    timeout: Duration,
) -> Result<SourceCacheDiagnostics, Box<dyn Error>> {
    let deadline = Instant::now() + timeout;
    loop {
        let diagnostics = package.cache_diagnostics();
        if diagnostics.in_flight_loads == expected_flights
            && diagnostics.waiter_joins >= expected_waiters
        {
            return Ok(diagnostics);
        }
        if Instant::now() >= deadline {
            return Err(format!(
                "OPC cache cohort reached {} flights and {} waiters, expected {expected_flights} flights and at least {expected_waiters} waiters",
                diagnostics.in_flight_loads, diagnostics.waiter_joins
            )
            .into());
        }
        std::thread::yield_now();
    }
}

fn opc_cache_result(
    case: Case,
    corpus: &Corpus,
    elapsed: Vec<u64>,
    mut source: SourceSummary,
    evidence: OpcCacheEvidenceSummary,
) -> CaseResult {
    source.opc_cache = Some(evidence);
    result_with_source(case, corpus, elapsed, source)
}

fn opc_cache_empty_scaling(model: &'static str) -> OpcCacheScalingSummary {
    OpcCacheScalingSummary {
        model,
        classification: "not-applicable",
        baseline_worker_count: 1,
        p50_speedup: None,
        p50_efficiency: None,
        amdahl_serial_fraction: None,
        p50_requests_per_second: 0.0,
        relative_request_throughput: 1.0,
    }
}

fn run_opc_source_cache_budget_boundary(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<Vec<CaseResult>, Box<dyn Error>> {
    validate_opc_cache_corpus(corpus)?;
    let (parts, ranges) = opc_cache_parts_and_ranges(corpus)?;
    let target_index = corpus.manifest.entry_count / 2;
    let target_part = parts
        .get(target_index)
        .ok_or("OPC cache boundary target index is outside the corpus")?;
    let payload_size = u64::try_from(corpus.manifest.entry_bytes)?;
    let cache_limits = SourceCacheLimits::new(corpus.manifest.entry_bytes, 1)?;
    let mut results = Vec::with_capacity(2);

    for (scenario, memory_limit, succeeds) in [
        ("exact-budget", payload_size, true),
        (
            "one-under-budget",
            payload_size
                .checked_sub(1)
                .ok_or("OPC cache payload must be nonempty")?,
            false,
        ),
    ] {
        let mut elapsed = Vec::with_capacity(samples);
        let mut source_summary = SourceSummary::default();
        let mut diagnostics_summary = OpcCacheDiagnosticsSummary::default();
        let mut after_handles = Vec::with_capacity(samples);
        let mut after_packages = Vec::with_capacity(samples);
        for iteration in 0..iteration_count(warmup_iterations, samples)? {
            let source = Arc::new(InstrumentedSource::new(
                corpus.archive.clone(),
                ranges.clone(),
            ));
            let read_at: Arc<dyn ReadAt> = source.clone();
            let (package, budget) = opc_cache_package(
                OpcCacheMode::Managed,
                read_at,
                cache_limits,
                memory_limit,
                1,
            )?;
            source.reset();
            let before = package.cache_diagnostics();
            opc_cache_assert_diagnostics_mode(OpcCacheMode::Managed, before, memory_limit)?;

            let started = Instant::now();
            let loaded = package.part(target_part)?.data();
            let duration = started.elapsed();
            let after = package.cache_diagnostics();
            opc_cache_assert_diagnostics_mode(OpcCacheMode::Managed, after, memory_limit)?;
            let delta = opc_cache_counter_delta(before, after)?;
            let source_snapshot = source.snapshot();

            if succeeds {
                let data = loaded?;
                let expected = payload_bytes(
                    PayloadKind::Incompressible,
                    target_index,
                    corpus.manifest.entry_bytes,
                );
                if data.as_bytes() != expected
                    || delta
                        != (OpcCacheCounterDelta {
                            cold_loads: 1,
                            successful_loads: 1,
                            ..OpcCacheCounterDelta::default()
                        })
                    || after.retained_entries != 1
                    || after.retained_bytes != corpus.manifest.entry_bytes
                    || after.in_flight_loads != 0
                    || after.budget_memory_used != payload_size
                    || after.budget_cache_reserved_bytes != payload_size
                    || opc_cache_budget_used(budget.as_ref()) != payload_size
                    || source_snapshot.ordinary_payload_read_calls == 0
                    || source_snapshot.ordinary_payload_read_bytes == 0
                {
                    return Err("OPC exact-budget boundary invariant failed".into());
                }
                drop(data);
                let used_after_handle = opc_cache_budget_used(budget.as_ref());
                if used_after_handle != payload_size {
                    return Err(
                        "OPC exact-budget cache reservation did not outlive its handle".into(),
                    );
                }
                if iteration >= warmup_iterations {
                    after_handles.push(used_after_handle);
                }
            } else {
                if !matches!(
                    loaded,
                    Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
                        if limit.resource == Resource::Memory
                ) || delta
                    != (OpcCacheCounterDelta {
                        budget_reservation_failures: 2,
                        ..OpcCacheCounterDelta::default()
                    })
                    || after.retained_entries != 0
                    || after.retained_bytes != 0
                    || after.in_flight_loads != 0
                    || after.budget_memory_used != 0
                    || after.budget_cache_reserved_bytes != 0
                    || opc_cache_budget_used(budget.as_ref()) != 0
                    || source_snapshot != SourceSnapshot::default()
                {
                    return Err(
                        "OPC one-under-budget refusal was not exact and zero-payload-I/O".into(),
                    );
                }
                if iteration >= warmup_iterations {
                    after_handles.push(0);
                }
            }

            drop(package);
            let used_after_package = opc_cache_budget_used(budget.as_ref());
            if used_after_package != 0 {
                return Err("OPC boundary package drop leaked a memory reservation".into());
            }
            if iteration >= warmup_iterations {
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
                source_summary.record_opc(source_snapshot, delta.successful_loads);
                diagnostics_summary.record(delta, after);
                after_packages.push(used_after_package);
            }
        }

        results.push(opc_cache_result(
            Case::OpcSourceCacheBudgetBoundary,
            corpus,
            elapsed,
            source_summary,
            OpcCacheEvidenceSummary {
                cache_mode: OpcCacheMode::Managed.name(),
                scenario,
                capacity_ratio: "1x",
                capacity_entries: 1,
                capacity_bytes: corpus.manifest.entry_bytes,
                working_set_parts: 1,
                working_set_bytes: payload_size,
                worker_count: 1,
                persistent_worker_teams_created: 0,
                fixed_source_delay_us: 0,
                timing_scope: "single PartData request; package open and verification excluded",
                diagnostics: diagnostics_summary,
                gate: None,
                budget_used_after_handles_drop: after_handles,
                budget_used_after_package_drop: after_packages,
                scaling: opc_cache_empty_scaling("budget-boundary-no-scaling"),
            },
        ));
    }
    Ok(results)
}

fn opc_cache_assignments(
    scenario: OpcCacheScenario,
    worker_count: usize,
    working_set_parts: usize,
    timed_parts: &[PackURI],
) -> Result<Vec<Vec<OpcCacheRequest>>, Box<dyn Error>> {
    let mut assignments = vec![Vec::new(); worker_count];
    match scenario {
        OpcCacheScenario::SamePart => {
            let target = timed_parts
                .first()
                .ok_or("OPC same-Part scenario has no target")?;
            for requests in &mut assignments {
                requests.push(OpcCacheRequest {
                    logical_index: 0,
                    partname: target.clone(),
                });
            }
        },
        OpcCacheScenario::DisjointParts => {
            if timed_parts.len() != working_set_parts || working_set_parts < worker_count {
                return Err("OPC disjoint-Part worker assignments cannot make progress".into());
            }
            for (logical_index, partname) in timed_parts.iter().enumerate() {
                assignments[logical_index % worker_count].push(OpcCacheRequest {
                    logical_index,
                    partname: partname.clone(),
                });
            }
        },
    }
    Ok(assignments)
}

fn opc_cache_verify_outputs(
    scenario: OpcCacheScenario,
    corpus: &Corpus,
    timed_start_index: usize,
    outputs: &mut [(usize, PartData)],
    expected_outputs: usize,
) -> Result<(), Box<dyn Error>> {
    if outputs.len() != expected_outputs {
        return Err("OPC cache worker team returned an unexpected result count".into());
    }
    outputs.sort_by_key(|(logical_index, _data)| *logical_index);
    for (position, (logical_index, data)) in outputs.iter().enumerate() {
        let payload_index = match scenario {
            OpcCacheScenario::SamePart => timed_start_index,
            OpcCacheScenario::DisjointParts => timed_start_index
                .checked_add(position)
                .ok_or("OPC cache verification index overflows")?,
        };
        if *logical_index
            != if matches!(scenario, OpcCacheScenario::SamePart) {
                0
            } else {
                position
            }
            || data.as_bytes()
                != payload_bytes(
                    PayloadKind::Incompressible,
                    payload_index,
                    corpus.manifest.entry_bytes,
                )
        {
            return Err("OPC cache worker result differs from deterministic payload".into());
        }
    }
    if matches!(scenario, OpcCacheScenario::SamePart)
        && outputs.first().is_some_and(|(_, first)| {
            outputs
                .iter()
                .skip(1)
                .any(|(_, other)| !first.shares_allocation_with(other))
        })
    {
        return Err("OPC same-Part waiters did not share one payload allocation".into());
    }
    Ok(())
}

fn run_opc_source_cache_contention_cell(
    config: OpcCacheCellConfig,
    corpus: &Corpus,
    parts: &[PackURI],
    ranges: &[Range<u64>],
) -> Result<CaseResult, Box<dyn Error>> {
    let OpcCacheCellConfig {
        case,
        mode,
        scenario,
        capacity,
        working_set_parts,
        worker_count,
        warmup_iterations,
        samples,
    } = config;
    let entry_bytes = corpus.manifest.entry_bytes;
    let working_set_bytes = u64::try_from(working_set_parts)?
        .checked_mul(u64::try_from(entry_bytes)?)
        .ok_or("OPC cache working-set byte count overflows")?;
    let capacity_parts = capacity.parts(working_set_parts);
    let capacity_entries = capacity_parts.max(1);
    let capacity_bytes = match (scenario, capacity) {
        (OpcCacheScenario::SamePart, OpcCacheCapacity::Half) => entry_bytes / 2,
        _ => capacity_entries
            .checked_mul(entry_bytes)
            .ok_or("OPC cache capacity byte count overflows")?,
    };
    let cache_limits = SourceCacheLimits::new(capacity_bytes, capacity_entries)?;
    let memory_limit = working_set_bytes
        .checked_mul(2)
        .ok_or("OPC managed cache memory limit overflows")?;
    let prefill_start = 0usize;
    let timed_start = working_set_parts;
    let timed_end = timed_start
        .checked_add(working_set_parts)
        .ok_or("OPC cache timed working set overflows")?;
    if timed_end > parts.len() || ranges.len() != parts.len() {
        return Err("OPC cache working sets exceed the fixed corpus".into());
    }
    let expected_outputs = match scenario {
        OpcCacheScenario::SamePart => worker_count,
        OpcCacheScenario::DisjointParts => working_set_parts,
    };
    let timed_parts = match scenario {
        OpcCacheScenario::SamePart => &parts[timed_start..timed_start + 1],
        OpcCacheScenario::DisjointParts => &parts[timed_start..timed_end],
    };
    let expected_initial = match scenario {
        OpcCacheScenario::SamePart => 1,
        OpcCacheScenario::DisjointParts => worker_count,
    };
    let expected_pre_release_flights = expected_initial;
    let expected_pre_release_waiters = match scenario {
        OpcCacheScenario::SamePart => u64::try_from(worker_count.saturating_sub(1))?,
        OpcCacheScenario::DisjointParts => 0,
    };
    let expected_delta = opc_cache_expected_delta(
        scenario,
        capacity,
        worker_count,
        working_set_parts,
        capacity_parts,
    )?;
    let expected_retained_entries = match (scenario, capacity) {
        (OpcCacheScenario::SamePart, OpcCacheCapacity::Half) => 0,
        (OpcCacheScenario::SamePart, OpcCacheCapacity::One) => 1,
        (OpcCacheScenario::SamePart, OpcCacheCapacity::Two) => 2,
        (OpcCacheScenario::DisjointParts, OpcCacheCapacity::Half) => capacity_parts,
        (OpcCacheScenario::DisjointParts, OpcCacheCapacity::One) => working_set_parts,
        (OpcCacheScenario::DisjointParts, OpcCacheCapacity::Two) => working_set_parts * 2,
    };
    let expected_retained_bytes = expected_retained_entries
        .checked_mul(entry_bytes)
        .ok_or("OPC cache retained byte count overflows")?;
    let expected_live_budget = if mode.is_managed() {
        match (scenario, capacity) {
            (OpcCacheScenario::SamePart, OpcCacheCapacity::Two) => {
                u64::try_from(entry_bytes)?.checked_mul(2)
            },
            (OpcCacheScenario::SamePart, _) => Some(u64::try_from(entry_bytes)?),
            (OpcCacheScenario::DisjointParts, OpcCacheCapacity::Two) => {
                working_set_bytes.checked_mul(2)
            },
            (OpcCacheScenario::DisjointParts, _) => Some(working_set_bytes),
        }
        .ok_or("OPC cache live Budget use overflows")?
    } else {
        0
    };
    let expected_cache_budget = if mode.is_managed() {
        u64::try_from(expected_retained_bytes)?
    } else {
        0
    };

    let worker_team = OpcCacheWorkerTeam::new(worker_count)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut diagnostics_summary = OpcCacheDiagnosticsSummary::default();
    let mut gate_summary = OpcCacheGateSummary::default();
    let mut after_handles = Vec::with_capacity(samples);
    let mut after_packages = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let backing = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            ranges.to_vec(),
        ));
        let source = Arc::new(CoordinatedSlowSource::new(
            Arc::clone(&backing),
            ranges.to_vec(),
            Duration::from_micros(OPC_CACHE_SLOW_SOURCE_DELAY_US),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let (package, budget) =
            opc_cache_package(mode, read_at, cache_limits, memory_limit, worker_count)?;

        for (index, part) in parts
            .iter()
            .enumerate()
            .take(timed_start)
            .skip(prefill_start)
        {
            let data = package.part(part)?.data()?;
            if data.as_bytes()
                != payload_bytes(
                    PayloadKind::Incompressible,
                    index,
                    corpus.manifest.entry_bytes,
                )
            {
                return Err("OPC cache prefill differs from deterministic payload".into());
            }
        }
        backing.reset();
        source.arm(expected_initial)?;
        let before = package.cache_diagnostics();
        opc_cache_assert_diagnostics_mode(mode, before, memory_limit)?;
        let assignments =
            opc_cache_assignments(scenario, worker_count, working_set_parts, timed_parts)?;
        worker_team.dispatch(Arc::clone(&package), assignments)?;
        if let Err(error) = source.wait_for_initial_arrivals(OPC_CACHE_COHORT_TIMEOUT) {
            source.release();
            let _ = worker_team.collect();
            return Err(error);
        }
        let pre_release = match opc_cache_wait_for_cohort(
            &package,
            expected_pre_release_flights,
            before
                .waiter_joins
                .checked_add(expected_pre_release_waiters)
                .ok_or("OPC cache waiter count overflows")?,
            OPC_CACHE_COHORT_TIMEOUT,
        ) {
            Ok(diagnostics) => diagnostics,
            Err(error) => {
                source.release();
                let _ = worker_team.collect();
                return Err(error);
            },
        };
        let started = Instant::now();
        source.release();
        let mut outputs = worker_team.collect()?;
        let duration = started.elapsed();

        opc_cache_verify_outputs(
            scenario,
            corpus,
            timed_start,
            &mut outputs,
            expected_outputs,
        )?;
        let after = package.cache_diagnostics();
        opc_cache_assert_diagnostics_mode(mode, after, memory_limit)?;
        let delta = opc_cache_counter_delta(before, after)?;
        let source_snapshot = backing.snapshot();
        let gate = source.snapshot()?;
        if delta != expected_delta
            || after.retained_entries != expected_retained_entries
            || after.retained_bytes != expected_retained_bytes
            || after.in_flight_loads != 0
            || after.budget_memory_used != expected_live_budget
            || after.budget_cache_reserved_bytes != expected_cache_budget
            || opc_cache_budget_used(budget.as_ref()) != expected_live_budget
            || gate.initial_arrivals != u64::try_from(expected_initial)?
            || gate.delayed_payload_arrivals != u64::try_from(working_set_parts)?
            || gate.max_concurrent_delays != u64::try_from(expected_initial)?
            || pre_release.in_flight_loads != expected_pre_release_flights
            || pre_release.waiter_joins.checked_sub(before.waiter_joins)
                != Some(expected_pre_release_waiters)
            || source_snapshot.ordinary_payload_read_calls == 0
            || source_snapshot.ordinary_payload_read_bytes == 0
        {
            return Err(format!(
                "OPC cache contention invariant failed for {}/{}/{}/workers={worker_count}: delta={delta:?}, after={after:?}, gate={gate:?}",
                mode.name(), scenario.name(), capacity.name()
            )
            .into());
        }

        drop(outputs);
        let after_handle_diagnostics = package.cache_diagnostics();
        let used_after_handle = opc_cache_budget_used(budget.as_ref());
        if used_after_handle != expected_cache_budget
            || after_handle_diagnostics.budget_memory_used != expected_cache_budget
            || after_handle_diagnostics.budget_cache_reserved_bytes != expected_cache_budget
        {
            return Err(
                "OPC cache handle drop did not leave exactly the retained reservation".into(),
            );
        }
        drop(package);
        let used_after_package = opc_cache_budget_used(budget.as_ref());
        if used_after_package != 0 {
            return Err("OPC cache package drop leaked a memory reservation".into());
        }

        if iteration >= warmup_iterations {
            record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            source_summary.record_opc(source_snapshot, delta.successful_loads);
            diagnostics_summary.record(delta, after);
            gate_summary.initial_arrivals.push(gate.initial_arrivals);
            gate_summary
                .delayed_payload_arrivals
                .push(gate.delayed_payload_arrivals);
            gate_summary
                .max_concurrent_delays
                .push(gate.max_concurrent_delays);
            gate_summary
                .pre_release_flights
                .push(pre_release.in_flight_loads);
            gate_summary.pre_release_waiters.push(
                pre_release
                    .waiter_joins
                    .checked_sub(before.waiter_joins)
                    .ok_or("OPC cache waiter count moved backwards")?,
            );
            after_handles.push(used_after_handle);
            after_packages.push(used_after_package);
        }
    }

    Ok(opc_cache_result(
        case,
        corpus,
        elapsed,
        source_summary,
        OpcCacheEvidenceSummary {
            cache_mode: mode.name(),
            scenario: scenario.name(),
            capacity_ratio: capacity.name(),
            capacity_entries,
            capacity_bytes,
            working_set_parts,
            working_set_bytes,
            worker_count,
            persistent_worker_teams_created: 1,
            fixed_source_delay_us: OPC_CACHE_SLOW_SOURCE_DELAY_US,
            timing_scope: "post-admission service completion; worker-team creation, package open, prefill, rendezvous, and verification excluded",
            diagnostics: diagnostics_summary,
            gate: Some(gate_summary),
            budget_used_after_handles_drop: after_handles,
            budget_used_after_package_drop: after_packages,
            scaling: opc_cache_empty_scaling("pending-group-classification"),
        },
    ))
}

fn opc_cache_set_scaling(
    result: &mut CaseResult,
    scenario: OpcCacheScenario,
    baseline_p50_ns: u64,
    worker_count: usize,
    request_count: usize,
) -> Result<(), Box<dyn Error>> {
    if baseline_p50_ns == 0 || result.elapsed_ns.p50 == 0 {
        return Err("OPC cache scaling sample has a zero duration".into());
    }
    let duration_ns = result.elapsed_ns.p50 as f64;
    let baseline_duration_ns = baseline_p50_ns as f64;
    let speedup = baseline_duration_ns / duration_ns;
    let throughput = request_count as f64 * 1_000_000_000.0 / duration_ns;
    let baseline_request_count = match scenario {
        OpcCacheScenario::SamePart => 1,
        OpcCacheScenario::DisjointParts => request_count,
    };
    let baseline_throughput =
        baseline_request_count as f64 * 1_000_000_000.0 / baseline_duration_ns;
    let (model, classification, p50_speedup, efficiency, serial_fraction) = match scenario {
        OpcCacheScenario::SamePart => (
            "throughput-only-variable-request-count",
            "amdahl-not-applicable-variable-request-count",
            None,
            None,
            None,
        ),
        OpcCacheScenario::DisjointParts if worker_count == 1 => (
            "amdahl-fixed-request-count",
            "baseline",
            Some(1.0),
            Some(1.0),
            None,
        ),
        OpcCacheScenario::DisjointParts => {
            let workers = worker_count as f64;
            let fraction = (1.0 / speedup - 1.0 / workers) / (1.0 - 1.0 / workers);
            let classification = if speedup > workers {
                "superlinear-observed-model-invalid"
            } else if speedup < 1.0 {
                "slowdown-observed-model-invalid"
            } else if !(0.0..=1.0).contains(&fraction) {
                "serial-fraction-outside-model"
            } else {
                "valid-amdahl-observation"
            };
            (
                "amdahl-fixed-request-count",
                classification,
                Some(speedup),
                Some(speedup / workers),
                Some(fraction),
            )
        },
    };
    let scaling = &mut result
        .source
        .as_mut()
        .and_then(|source| source.opc_cache.as_mut())
        .ok_or("OPC cache result is missing structured evidence")?
        .scaling;
    *scaling = OpcCacheScalingSummary {
        model,
        classification,
        baseline_worker_count: 1,
        p50_speedup,
        p50_efficiency: efficiency,
        amdahl_serial_fraction: serial_fraction,
        p50_requests_per_second: throughput,
        relative_request_throughput: throughput / baseline_throughput,
    };
    Ok(())
}

fn run_opc_source_cache_contention(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    worker_counts: &[usize],
    mode: OpcCacheMode,
) -> Result<Vec<CaseResult>, Box<dyn Error>> {
    validate_opc_cache_corpus(corpus)?;
    if !matches!(
        (case, mode),
        (Case::OpcSourceCacheControlContention, OpcCacheMode::Control)
            | (Case::OpcSourceCacheManagedContention, OpcCacheMode::Managed)
    ) {
        return Err("OPC cache contention case disagrees with its selected mode".into());
    }
    if worker_counts.first().copied() != Some(1)
        || worker_counts.contains(&0)
        || !worker_counts.windows(2).all(|pair| pair[0] < pair[1])
    {
        return Err(
            "OPC cache contention requires sorted, unique worker widths including one".into(),
        );
    }
    let largest_worker_count = worker_counts
        .last()
        .copied()
        .ok_or("OPC cache contention requires a worker width")?;
    let minimum_disjoint_parts = largest_worker_count.max(2);
    let disjoint_parts = minimum_disjoint_parts
        .checked_add(minimum_disjoint_parts % 2)
        .ok_or("OPC cache disjoint working-set size overflows")?;
    if disjoint_parts
        .checked_mul(2)
        .is_none_or(|required| required > corpus.manifest.entry_count)
    {
        return Err("OPC cache contention worker widths exceed the fixed corpus".into());
    }
    let (parts, ranges) = opc_cache_parts_and_ranges(corpus)?;
    let mut results = Vec::with_capacity(
        worker_counts
            .len()
            .checked_mul(2)
            .and_then(|count| count.checked_mul(OpcCacheCapacity::ALL.len()))
            .ok_or("OPC cache contention matrix size overflows")?,
    );

    for scenario in [OpcCacheScenario::SamePart, OpcCacheScenario::DisjointParts] {
        let working_set_parts = match scenario {
            OpcCacheScenario::SamePart => 1,
            OpcCacheScenario::DisjointParts => disjoint_parts,
        };
        let request_count = match scenario {
            OpcCacheScenario::SamePart => None,
            OpcCacheScenario::DisjointParts => Some(disjoint_parts),
        };
        for capacity in OpcCacheCapacity::ALL {
            let group_start = results.len();
            for &worker_count in worker_counts {
                results.push(run_opc_source_cache_contention_cell(
                    OpcCacheCellConfig {
                        case,
                        mode,
                        scenario,
                        capacity,
                        working_set_parts,
                        worker_count,
                        warmup_iterations,
                        samples,
                    },
                    corpus,
                    &parts,
                    &ranges,
                )?);
            }
            let baseline_p50 = results
                .get(group_start)
                .ok_or("OPC cache contention matrix has no baseline")?
                .elapsed_ns
                .p50;
            for result in &mut results[group_start..] {
                let worker_count = result
                    .source
                    .as_ref()
                    .and_then(|source| source.opc_cache.as_ref())
                    .ok_or("OPC cache result is missing structured evidence")?
                    .worker_count;
                opc_cache_set_scaling(
                    result,
                    scenario,
                    baseline_p50,
                    worker_count,
                    request_count.unwrap_or(worker_count),
                )?;
            }
        }
    }
    Ok(results)
}

fn opc_source_cache_limits(corpus: &Corpus) -> Result<SourceCacheLimits, Box<dyn Error>> {
    Ok(SourceCacheLimits::new(
        corpus.manifest.uncompressed_payload_bytes.max(1),
        corpus.manifest.entry_count.max(1),
    )?)
}

fn opc_instrumented_source(
    corpus: &Corpus,
) -> Result<(Arc<InstrumentedSource>, Range<u64>), Box<dyn Error>> {
    let (ordinary_payload_ranges, target_range) = opc_payload_ranges(corpus)?;
    Ok((
        Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            ordinary_payload_ranges,
        )),
        target_range,
    ))
}

fn simulated_source(
    backing: Arc<InstrumentedSource>,
    config: RangeSimulationConfig,
) -> Arc<SimulatedRangeSource> {
    Arc::new(SimulatedRangeSource::new(backing, config))
}

fn verify_simulation_snapshot(
    snapshot: &RangeSimulationSnapshot,
    config: RangeSimulationConfig,
    context: &str,
) -> Result<(), Box<dyn Error>> {
    if snapshot.logical_read_calls == 0
        || snapshot.logical_read_bytes == 0
        || snapshot.physical_request_count == 0
        || snapshot.physical_request_bytes == 0
    {
        return Err(format!("{context} performed no simulated source I/O").into());
    }
    if snapshot.logical_read_bytes != snapshot.physical_request_bytes {
        return Err(format!("{context} logical and physical byte totals differ").into());
    }
    if snapshot.physical_request_count != u64::try_from(snapshot.physical_request_sizes.len())?
        || snapshot
            .physical_request_sizes
            .iter()
            .any(|&bytes| bytes == 0 || bytes > config.max_physical_range_bytes as u64)
    {
        return Err(format!("{context} physical request distribution is invalid").into());
    }
    Ok(())
}

fn run_opc_range_source_open(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    config: RangeSimulationConfig,
    read_main: bool,
) -> Result<CaseResult, Box<dyn Error>> {
    let case = if read_main {
        Case::OpcRangeSourceOpenMainRead
    } else {
        Case::OpcRangeSourceOpen
    };
    let cache_limits = opc_source_cache_limits(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let (backing, _target_range) = opc_instrumented_source(corpus)?;
        let source = simulated_source(backing.clone(), config);
        let started = Instant::now();
        let package =
            SourceBackedPackage::from_read_at_with_cache_limits(source.clone(), cache_limits)?;
        if package.iter_parts().count() != corpus.manifest.entry_count {
            return Err("simulated OPC source open part count differs from manifest".into());
        }
        if read_main {
            std::hint::black_box(verify_opc_main_payload(&package, corpus)?);
        }
        let duration = started.elapsed();
        let metrics = backing.snapshot();
        let simulation = source.snapshot()?;
        verify_simulation_snapshot(&simulation, config, "simulated OPC source open")?;

        if !read_main {
            let before_proof = source.snapshot()?;
            std::hint::black_box(verify_opc_main_payload(&package, corpus)?);
            let after_proof = source.snapshot()?;
            if after_proof.physical_request_count <= before_proof.physical_request_count
                || backing.snapshot().ordinary_payload_read_calls
                    <= metrics.ordinary_payload_read_calls
            {
                return Err(
                    "simulated OPC structural open had already materialized the main part".into(),
                );
            }
        }

        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::from(read_main));
            source_summary.record_simulation(simulation);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(case, corpus, elapsed, source_summary))
}

fn verify_opc_main_payload(
    package: &SourceBackedPackage,
    corpus: &Corpus,
) -> Result<Arc<Vec<u8>>, Box<dyn Error>> {
    let main = package.main_document_part()?;
    if main.partname().membername() != corpus.target_name {
        return Err("source-backed OPC main relationship resolved the wrong part".into());
    }
    let bytes = main.data()?.into_arc()?;
    if bytes.as_slice() != corpus.target_payload {
        return Err("source-backed OPC main payload differs from deterministic corpus".into());
    }
    Ok(bytes)
}

fn run_opc_source_open(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    read_main: bool,
) -> Result<CaseResult, Box<dyn Error>> {
    let case = if read_main {
        Case::OpcSourceOpenMainRead
    } else {
        Case::OpcSourceOpen
    };
    let cache_limits = opc_source_cache_limits(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let (source, _target_range) = opc_instrumented_source(corpus)?;
        let started = Instant::now();
        let package =
            SourceBackedPackage::from_read_at_with_cache_limits(source.clone(), cache_limits)?;
        if package.iter_parts().count() != corpus.manifest.entry_count {
            return Err("source-backed OPC part count differs from corpus manifest".into());
        }
        let after_open = source.snapshot();
        if after_open.read_calls == 0 || after_open.read_bytes == 0 {
            return Err("source-backed OPC open performed no positional reads".into());
        }
        let metrics = if read_main {
            std::hint::black_box(verify_opc_main_payload(&package, corpus)?);
            let metrics = source.snapshot();
            if metrics.read_calls <= after_open.read_calls
                || metrics.ordinary_payload_read_calls <= after_open.ordinary_payload_read_calls
            {
                return Err("source-backed OPC main payload was materialized during open".into());
            }
            metrics
        } else {
            after_open
        };
        let duration = started.elapsed();
        if !read_main {
            std::hint::black_box(verify_opc_main_payload(&package, corpus)?);
            let after_proof_read = source.snapshot();
            if after_proof_read.read_calls <= metrics.read_calls
                || after_proof_read.ordinary_payload_read_calls
                    <= metrics.ordinary_payload_read_calls
            {
                return Err(
                    "source-backed OPC open unexpectedly cached an ordinary payload".into(),
                );
            }
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::from(read_main));
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(case, corpus, elapsed, source_summary))
}

fn run_opc_source_cached_main_read(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let cache_limits = opc_source_cache_limits(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let (source, _target_range) = opc_instrumented_source(corpus)?;
        let package =
            SourceBackedPackage::from_read_at_with_cache_limits(source.clone(), cache_limits)?;
        let cold = verify_opc_main_payload(&package, corpus)?;
        source.reset();
        let started = Instant::now();
        let cached = verify_opc_main_payload(&package, corpus)?;
        let duration = started.elapsed();
        if !Arc::ptr_eq(&cold, &cached) {
            return Err("source-backed OPC cache hit returned a different allocation".into());
        }
        let metrics = source.snapshot();
        if metrics != SourceSnapshot::default() {
            return Err("source-backed OPC cache hit performed a source read".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, 0);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::OpcSourceCachedMainRead,
        corpus,
        elapsed,
        source_summary,
    ))
}

fn run_opc_source_concurrent_same_part(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let cache_limits = opc_source_cache_limits(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let (source, _target_range) = opc_instrumented_source(corpus)?;
        let package =
            SourceBackedPackage::from_read_at_with_cache_limits(source.clone(), cache_limits)?;
        if package.main_document_part()?.partname().membername() != corpus.target_name {
            return Err("source-backed OPC main relationship resolved the wrong part".into());
        }
        source.reset();
        let start = Arc::new(Barrier::new(3));
        let started = Instant::now();
        let (first, second) = std::thread::scope(|scope| {
            let first_start = Arc::clone(&start);
            let first_package = &package;
            let first_task = scope.spawn(move || {
                first_start.wait();
                first_package
                    .main_document_part()
                    .and_then(|part| part.data())
                    .and_then(|data| data.into_arc())
            });
            let second_start = Arc::clone(&start);
            let second_package = &package;
            let second_task = scope.spawn(move || {
                second_start.wait();
                second_package
                    .main_document_part()
                    .and_then(|part| part.data())
                    .and_then(|data| data.into_arc())
            });
            start.wait();
            (first_task.join(), second_task.join())
        });
        let first = first.map_err(|_panic| "first OPC source worker panicked")??;
        let second = second.map_err(|_panic| "second OPC source worker panicked")??;
        let duration = started.elapsed();
        if first.as_slice() != corpus.target_payload || second.as_slice() != corpus.target_payload {
            return Err("concurrent same-part OPC reads returned unexpected bytes".into());
        }
        if !Arc::ptr_eq(&first, &second) {
            return Err("concurrent same-part OPC reads did not share one allocation".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("concurrent same-part OPC reads loaded no payload source bytes".into());
        }
        let cached = verify_opc_main_payload(&package, corpus)?;
        if !Arc::ptr_eq(&first, &cached) || source.snapshot() != metrics {
            return Err("concurrent same-part OPC load did not leave a shared cache hit".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, 1);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::OpcSourceConcurrentSamePart,
        corpus,
        elapsed,
        source_summary,
    ))
}

fn opc_overlay_replacement_payload(corpus: &Corpus) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut payload = corpus.target_payload.clone();
    let first = payload
        .first_mut()
        .ok_or("OPC overlay replacement target is empty")?;
    *first ^= 0xff;
    Ok(payload)
}

fn expected_opc_overlay_output(
    corpus: &Corpus,
    replacement: &[u8],
) -> Result<Vec<u8>, Box<dyn Error>> {
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    let mut package = OpcPackage::from_bytes(&corpus.archive)?;
    package
        .get_part_mut(&target_uri)?
        .set_blob(replacement.to_vec());
    Ok(PackageWriter::to_bytes(&package)?)
}

fn verify_opc_overlay_output(
    corpus: &Corpus,
    output: &[u8],
    replacement: &[u8],
) -> Result<(), Box<dyn Error>> {
    let package = OpcPackage::from_bytes(output)?;
    if package.part_count() != corpus.manifest.entry_count {
        return Err("OPC overlay output part count differs from source corpus".into());
    }
    for index in 0..corpus.manifest.entry_count {
        let name = entry_name(index);
        let uri = PackURI::new(format!("/{name}"))?;
        let part = package.get_part(&uri)?;
        let expected = if name == corpus.target_name {
            replacement.to_vec()
        } else {
            payload_bytes(
                PayloadKind::Incompressible,
                index,
                corpus.manifest.entry_bytes,
            )
        };
        if part.content_type() != CONTENT_TYPE || part.blob() != expected {
            return Err(format!("OPC overlay output Part {name} differs from expectation").into());
        }
    }
    let main = package.main_document_part()?;
    if main.partname().membername() != corpus.target_name || main.blob() != replacement {
        return Err("OPC overlay output main relationship or payload differs".into());
    }
    Ok(())
}

fn run_opc_source_overlay_one_part_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.shape != CorpusShape::FewLarge.name()
        || corpus.manifest.payload_kind != PayloadKind::Incompressible.name()
    {
        return Err("OPC source overlay case requires few-large incompressible corpus".into());
    }
    let replacement = opc_overlay_replacement_payload(corpus)?;
    let expected = expected_opc_overlay_output(corpus, &replacement)?;
    if expected == corpus.archive {
        return Err("OPC overlay expected output did not change source bytes".into());
    }
    verify_opc_overlay_output(corpus, &expected, &replacement)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("OPC overlay sequential output ceiling overflows u64")?;
    let cache_limits = opc_source_cache_limits(corpus)?;
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let (source, _target_range) = opc_instrumented_source(corpus)?;
        let replacement_part = replacement.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let source_package =
            SourceBackedPackage::from_read_at_with_cache_limits(source.clone(), cache_limits)?;
        source_package.write_part_overlay_to_stream(&mut sink, &target_uri, replacement_part)?;
        let duration = started.elapsed();

        let metrics = source.snapshot();
        if metrics.read_calls == 0
            || metrics.read_bytes == 0
            || metrics.ordinary_payload_read_calls == 0
            || metrics.ordinary_payload_read_bytes == 0
        {
            return Err("OPC overlay save performed no ordinary source reads".into());
        }
        if sink.bytes != expected {
            return Err("OPC overlay save differs from deterministic expected output".into());
        }
        verify_opc_overlay_output(corpus, &sink.bytes, &replacement)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("OPC overlay output digest differs between iterations".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, 1);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, "OPC source overlay save")?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("OPC overlay measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: Case::OpcSourceOverlayOnePartSave.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn relationship_signatures(relationships: &Relationships) -> Vec<(String, String, String, bool)> {
    let mut signatures = relationships
        .iter()
        .map(|relationship| {
            (
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.is_external(),
            )
        })
        .collect::<Vec<_>>();
    signatures.sort();
    signatures
}

fn verify_docx_source_edit_output(corpus: &Corpus, output: &[u8]) -> Result<(), Box<dyn Error>> {
    let target = SemanticShape::Medium.docx_paragraphs() / 2;
    let document = litchi_docx::Package::from_reader(Cursor::new(output))?;
    verify_semantic_docx(&document, SemanticShape::Medium, &[target])?;

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("DOCX source-edit package topology differs from source".into());
    }
    let main_uri = PackURI::new("/word/document.xml")?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("DOCX source-edit Part metadata differs from source".into());
        }
        if source_part.partname() == &main_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("DOCX source-edit main document did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("DOCX source-edit changed an unselected Part payload".into());
        }
    }
    for index in 0..DOCX_SOURCE_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/word/media/image{}.png", index + 1))?;
        if candidate.get_part(&uri)?.blob() != docx_source_media_payload(index) {
            return Err("DOCX source-edit media readback differs from specification".into());
        }
    }
    Ok(())
}

fn publish_docx_source_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: W,
) -> Result<(usize, litchi_docx::document::Commit), Box<dyn Error>> {
    let package = litchi_docx::source_backed::Package::from_read_at(source)?;
    let target = SemanticShape::Medium.docx_paragraphs() / 2;
    let mut edit = package.edit_document()?;
    edit.replace_paragraph_text(Position::new(target), semantic_docx_text(target, true))?;
    let commit = edit.commit()?;
    if !commit.patch().changed() || commit.diagnostics().operations() != 1 {
        return Err("DOCX source edit produced unexpected commit diagnostics".into());
    }
    let materializations = usize::try_from(package.cache_diagnostics().successful_loads)?;
    let published = package.publish_document_commit_to_stream(writer, &commit)?;
    if published.xml_bytes() != commit.snapshot().xml_bytes() {
        return Err("DOCX source edit published a different snapshot".into());
    }
    Ok((materializations, commit))
}

fn run_docx_source_backed_one_edit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != DOCX_SOURCE_EDIT_CORPUS_GENERATOR {
        return Err("DOCX source-edit case requires its fixed media-rich corpus".into());
    }
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let (expected_materializations, expected_commit) =
        publish_docx_source_edit(expected_source, &mut expected)?;
    if expected == corpus.archive
        || expected_materializations != 1
        || expected_commit.diagnostics().operations() != 1
    {
        return Err("DOCX source edit did not materialize exactly its main Part".into());
    }
    verify_docx_source_edit_output(corpus, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("DOCX source-edit sequential output ceiling overflows u64")?;
    let (payload_ranges, _target_range) = docx_source_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let (materializations, commit) = publish_docx_source_edit(read_at, &mut sink)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("DOCX source edit differs between iterations".into());
        }
        let replayed = commit.patch().apply(commit.patch().source())?;
        if replayed.xml_bytes() != commit.snapshot().xml_bytes() {
            return Err("DOCX source-edit patch replay differs from commit".into());
        }
        let restored = commit.patch().inverse().apply(commit.snapshot())?;
        if restored.xml_bytes() != commit.patch().source().xml_bytes() {
            return Err("DOCX source-edit inverse did not restore main XML".into());
        }
        if commit.patch().apply(commit.snapshot()).is_ok() {
            return Err("DOCX source-edit patch accepted a stale target".into());
        }
        verify_docx_source_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("DOCX source-edit output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("DOCX source edit performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, "DOCX source-backed edit/save")?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("DOCX source-edit measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: Case::DocxSourceBackedOneEditSave.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn verify_pptx_source_edit_output(
    corpus: &Corpus,
    output: &[u8],
    updated_shapes: usize,
) -> Result<(), Box<dyn Error>> {
    let reopened = litchi_pptx::Package::from_bytes(output)?;
    verify_pptx_source_edit_semantics(&reopened, updated_shapes)?;

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("PPTX source-edit package topology differs from source".into());
    }
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("PPTX source-edit Part metadata differs from source".into());
        }
        if source_part.partname() == &target_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("PPTX source-edit selected slide did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("PPTX source-edit changed an unselected Part payload".into());
        }
    }
    for index in 0..PPTX_SOURCE_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!(
            "/ppt/media/litchi-perf-source-media-{index:02}.png"
        ))?;
        if candidate.get_part(&uri)?.blob() != pptx_source_media_payload(index) {
            return Err("PPTX source-edit media readback differs from specification".into());
        }
    }
    Ok(())
}

const fn pptx_multi_slide_positions() -> [usize; PPTX_MULTI_SLIDE_BATCH_COUNT] {
    [0, 28, 57, 85, 114, 142, 171, 199]
}

fn verify_pptx_multi_slide_edit_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let positions = pptx_multi_slide_positions();
    let updated = positions.map(|position| (position, PPTX_SOURCE_TEXT_BOXES_PER_SLIDE));
    let reopened = litchi_pptx::Package::from_bytes(output)?;
    verify_pptx_source_edit_semantics_for(&reopened, &updated)?;

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("PPTX multi-slide package topology differs from source".into());
    }
    let selected = positions
        .map(|position| PackURI::new(format!("/ppt/slides/slide{}.xml", position + 1)))
        .into_iter()
        .collect::<Result<Vec<_>, _>>()?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("PPTX multi-slide Part metadata differs from source".into());
        }
        let is_selected = selected.contains(source_part.partname());
        if is_selected == (source_part.blob() == candidate_part.blob()) {
            return Err("PPTX multi-slide changed the wrong Part payload set".into());
        }
    }
    for index in 0..PPTX_SOURCE_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!(
            "/ppt/media/litchi-perf-source-media-{index:02}.png"
        ))?;
        if candidate.get_part(&uri)?.blob() != pptx_source_media_payload(index) {
            return Err("PPTX multi-slide media readback differs from specification".into());
        }
    }
    Ok(())
}

fn publish_pptx_source_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: &mut W,
) -> Result<usize, Box<dyn Error>> {
    let editor = litchi_pptx::SourceBackedPresentationEditor::from_read_at(source)?;
    let target_slide = PPTX_SOURCE_SLIDE_COUNT / 2;
    let mut edit = editor.edit_slide(target_slide)?;
    if !edit.set_shape_text(0, semantic_pptx_text(target_slide, 0, true))? {
        return Err("PPTX source edit unexpectedly reported no change".into());
    }
    let commit = edit.commit();
    if !commit.is_changed() {
        return Err("PPTX source edit produced an unchanged commit".into());
    }
    let replayed = commit.patch().apply(commit.patch().source())?;
    if commit.patch().inverse().apply(&replayed).is_err() {
        return Err("PPTX source edit inverse rejected its candidate".into());
    }
    let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
    let published = editor.publish_slide_commit_to_stream(writer, &commit)?;
    if commit.patch().inverse().apply(&published).is_err()
        || commit.patch().apply(&published).is_ok()
    {
        return Err("PPTX source edit published a different slide snapshot".into());
    }
    Ok(materializations)
}

fn run_pptx_source_backed_one_edit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != PPTX_SOURCE_EDIT_CORPUS_GENERATOR {
        return Err("PPTX source-edit case requires its fixed media-rich corpus".into());
    }
    if sha256_hex(&corpus.archive) != corpus.manifest.archive_sha256 {
        return Err("PPTX source-edit source digest differs from corpus manifest".into());
    }
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations = publish_pptx_source_edit(expected_source, &mut expected)?;
    if expected == corpus.archive || expected_materializations != 2 {
        return Err(
            "PPTX source edit did not materialize exactly its root and selected slide".into(),
        );
    }
    verify_pptx_source_edit_output(corpus, &expected, 1)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("PPTX source-edit sequential output ceiling overflows u64")?;
    let payload_ranges = pptx_source_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations = publish_pptx_source_edit(read_at, &mut sink)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("PPTX source edit differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("PPTX source edit exceeded the sequential sink write bound".into());
        }
        verify_pptx_source_edit_output(corpus, &sink.bytes, 1)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("PPTX source-edit output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("PPTX source edit performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, "PPTX source-backed edit/save")?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("PPTX source-edit measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: Case::PptxSourceBackedOneEditSave.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn pptx_batch_replacements(texts: &[String]) -> Vec<litchi_pptx::ShapeTextReplacement<'_>> {
    texts
        .iter()
        .enumerate()
        .map(|(index, text)| litchi_pptx::ShapeTextReplacement::at(index, text))
        .collect()
}

fn publish_pptx_batch_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: &mut W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    let target_slide = PPTX_SOURCE_SLIDE_COUNT / 2;
    let texts = (0..PPTX_SOURCE_TEXT_BOXES_PER_SLIDE)
        .map(|shape| semantic_pptx_text(target_slide, shape, true))
        .collect::<Vec<_>>();
    let replacements = pptx_batch_replacements(&texts);
    if source_backed {
        let editor = litchi_pptx::SourceBackedPresentationEditor::from_read_at(source)?;
        let mut edit = editor.edit_slide(target_slide)?;
        if edit.set_shape_texts(&replacements)? != PPTX_SOURCE_TEXT_BOXES_PER_SLIDE {
            return Err("PPTX source-backed batch changed an unexpected shape count".into());
        }
        let commit = edit.commit();
        if !commit.is_changed() {
            return Err("PPTX source-backed batch produced an unchanged commit".into());
        }
        let replayed = commit.patch().apply(commit.patch().source())?;
        if commit.patch().inverse().apply(&replayed).is_err() {
            return Err("PPTX source-backed batch inverse rejected its candidate".into());
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_slide_commit_to_stream(writer, &commit)?;
        if commit.patch().inverse().apply(&published).is_err()
            || commit.patch().apply(&published).is_ok()
        {
            return Err("PPTX source-backed batch published another snapshot".into());
        }
        Ok(materializations)
    } else {
        let package = SourceBackedPackage::from_read_at(source)?;
        let opc = package.into_opc_package()?;
        let materializations = opc.part_count();
        let mut package = litchi_pptx::Package::from_opc_package(opc)?;
        let mut edit = package.opened_presentation_transaction()?;
        if edit.set_shape_texts(target_slide, &replacements)? != PPTX_SOURCE_TEXT_BOXES_PER_SLIDE {
            return Err("PPTX eager batch changed an unexpected shape count".into());
        }
        let commit = edit.commit()?;
        if !commit.is_changed() {
            return Err("PPTX eager batch produced an unchanged commit".into());
        }
        let inverse = commit.patch().inverse();
        let published = package.apply_opened_presentation_commit(commit)?;
        if package
            .apply_opened_presentation_patch(&inverse)
            .and_then(|_| package.apply_opened_presentation_patch(&inverse.inverse()))
            .is_err()
        {
            return Err("PPTX eager batch inverse/forward replay failed".into());
        }
        if published.slides().len() != PPTX_SOURCE_SLIDE_COUNT {
            return Err("PPTX eager batch changed the slide graph".into());
        }
        let output = package.to_bytes()?;
        for chunk in output.chunks(64 * 1024) {
            writer.write_all(chunk)?;
        }
        Ok(materializations)
    }
}

fn run_pptx_batch_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != PPTX_SOURCE_EDIT_CORPUS_GENERATOR
        || !matches!(
            case,
            Case::PptxEagerBatchEditSave | Case::PptxSourceBackedBatchEditSave
        )
    {
        return Err("PPTX batch case requires its fixed media-rich corpus".into());
    }
    if sha256_hex(&corpus.archive) != corpus.manifest.archive_sha256 {
        return Err("PPTX batch source digest differs from corpus manifest".into());
    }
    let source_backed = case == Case::PptxSourceBackedBatchEditSave;
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations =
        publish_pptx_batch_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        2
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err("PPTX batch materialized an unexpected Part count".into());
    }
    verify_pptx_source_edit_output(corpus, &expected, PPTX_SOURCE_TEXT_BOXES_PER_SLIDE)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("PPTX batch sequential output ceiling overflows u64")?;
    let payload_ranges = pptx_source_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations = publish_pptx_batch_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("PPTX batch differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("PPTX batch exceeded the sequential sink write bound".into());
        }
        verify_pptx_source_edit_output(corpus, &sink.bytes, PPTX_SOURCE_TEXT_BOXES_PER_SLIDE)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("PPTX batch output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("PPTX batch performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("PPTX batch measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn publish_pptx_multi_slide_batch_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: &mut W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    if source_backed {
        let editor = litchi_pptx::SourceBackedPresentationEditor::from_read_at(source)?;
        let mut edit = editor.edit_slides();
        for slide_position in pptx_multi_slide_positions() {
            let texts = (0..PPTX_SOURCE_TEXT_BOXES_PER_SLIDE)
                .map(|shape| semantic_pptx_text(slide_position, shape, true))
                .collect::<Vec<_>>();
            let replacements = pptx_batch_replacements(&texts);
            if edit.set_shape_texts(slide_position, &replacements)?
                != PPTX_SOURCE_TEXT_BOXES_PER_SLIDE
            {
                return Err(
                    "PPTX source-backed multi-slide batch changed an unexpected shape count".into(),
                );
            }
        }
        let commit = edit.commit()?;
        if !commit.is_changed() {
            return Err("PPTX source-backed multi-slide batch produced no changes".into());
        }
        let replayed = commit.patch().apply(commit.patch().source())?;
        if commit.patch().inverse().apply(&replayed).is_err() {
            return Err("PPTX source-backed multi-slide inverse rejected its candidate".into());
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_slide_batch_commit_to_stream(writer, &commit)?;
        if commit.patch().inverse().apply(&published).is_err()
            || commit.patch().apply(&published).is_ok()
        {
            return Err("PPTX source-backed multi-slide published another snapshot".into());
        }
        return Ok(materializations);
    }

    let package = SourceBackedPackage::from_read_at(source)?;
    let opc = package.into_opc_package()?;
    let materializations = opc.part_count();
    let mut package = litchi_pptx::Package::from_opc_package(opc)?;
    let mut edit = package.opened_presentation_transaction()?;
    for slide_position in pptx_multi_slide_positions() {
        let texts = (0..PPTX_SOURCE_TEXT_BOXES_PER_SLIDE)
            .map(|shape| semantic_pptx_text(slide_position, shape, true))
            .collect::<Vec<_>>();
        let replacements = pptx_batch_replacements(&texts);
        if edit.set_shape_texts(slide_position, &replacements)? != PPTX_SOURCE_TEXT_BOXES_PER_SLIDE
        {
            return Err("PPTX eager multi-slide batch changed an unexpected shape count".into());
        }
    }
    let commit = edit.commit()?;
    if !commit.is_changed() {
        return Err("PPTX eager multi-slide batch produced an unchanged commit".into());
    }
    let inverse = commit.patch().inverse();
    let published = package.apply_opened_presentation_commit(commit)?;
    if package
        .apply_opened_presentation_patch(&inverse)
        .and_then(|_| package.apply_opened_presentation_patch(&inverse.inverse()))
        .is_err()
    {
        return Err("PPTX eager multi-slide batch inverse/forward replay failed".into());
    }
    if published.slides().len() != PPTX_SOURCE_SLIDE_COUNT {
        return Err("PPTX eager multi-slide batch changed the slide graph".into());
    }
    let output = package.to_bytes()?;
    for chunk in output.chunks(64 * 1024) {
        writer.write_all(chunk)?;
    }
    Ok(materializations)
}

fn run_pptx_multi_slide_batch_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != PPTX_SOURCE_EDIT_CORPUS_GENERATOR
        || !matches!(
            case,
            Case::PptxEagerMultiSlideBatchEditSave | Case::PptxSourceBackedMultiSlideBatchEditSave
        )
    {
        return Err("PPTX multi-slide batch requires its fixed media-rich corpus".into());
    }
    if sha256_hex(&corpus.archive) != corpus.manifest.archive_sha256 {
        return Err("PPTX multi-slide source digest differs from corpus manifest".into());
    }
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let source_backed = case == Case::PptxSourceBackedMultiSlideBatchEditSave;
    let expected_materializations =
        publish_pptx_multi_slide_batch_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        PPTX_MULTI_SLIDE_BATCH_COUNT + 1
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err("PPTX multi-slide materialized an unexpected Part count".into());
    }
    verify_pptx_multi_slide_edit_output(corpus, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("PPTX multi-slide sequential output ceiling overflows u64")?;
    let payload_ranges = pptx_source_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations =
            publish_pptx_multi_slide_batch_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("PPTX multi-slide batch differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("PPTX multi-slide batch exceeded the sink write bound".into());
        }
        verify_pptx_multi_slide_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("PPTX multi-slide output digest differs from expected".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("PPTX multi-slide batch performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("PPTX multi-slide measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn xlsx_defined_names_target() -> Vec<litchi_xlsx::raw::DefinedName> {
    vec![
        litchi_xlsx::raw::DefinedName {
            name: "PerformanceRange".to_owned(),
            reference: "'Sheet1'!$A$1:$B$2".to_owned(),
            comment: Some("deterministic source-backed benchmark".to_owned()),
            ..litchi_xlsx::raw::DefinedName::default()
        },
        litchi_xlsx::raw::DefinedName {
            name: "LocalPerformanceCell".to_owned(),
            reference: "'Sheet1'!$C$3".to_owned(),
            local_sheet_id: Some(0),
            hidden: true,
            ..litchi_xlsx::raw::DefinedName::default()
        },
    ]
}

fn verify_xlsx_defined_names_edit_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let reopened = litchi_xlsx::Package::from_slice(output)?;
    if reopened.workbook()?.defined_names() != xlsx_defined_names_target() {
        return Err("XLSX defined-name output has unexpected catalog semantics".into());
    }
    if reopened
        .calculation_metadata()?
        .properties()
        .ok_or("XLSX defined-name output has no calcPr")?
        .calculation_id()
        != 7
    {
        return Err("XLSX defined-name output changed calculation metadata".into());
    }

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("XLSX defined-name package topology differs from source".into());
    }
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("XLSX defined-name Part metadata differs from source".into());
        }
        if source_part.partname() == &target_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("XLSX defined-name workbook XML did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("XLSX defined-name edit changed an unselected Part payload".into());
        }
    }
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        if candidate.get_part(&uri)?.blob() != xlsx_calculation_media_payload(index) {
            return Err("XLSX defined-name media readback differs from specification".into());
        }
    }
    Ok(())
}

fn publish_xlsx_defined_names_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    if source_backed {
        let editor = litchi_xlsx::defined_names::SourceBackedEditor::from_read_at(source)?;
        let mut edit = editor.edit();
        if !edit.replace(xlsx_defined_names_target())? {
            return Err("XLSX source-backed defined-name edit reported an exact no-op".into());
        }
        let commit = edit.commit()?;
        if !commit.changed()
            || commit.patch().is_empty()
            || commit.patch().inverse().after() != commit.patch().before()
        {
            return Err("XLSX source-backed defined-name edit produced an invalid patch".into());
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_commit_to_stream(writer, &commit)?;
        if published != *commit.snapshot() {
            return Err("XLSX source-backed defined-name edit published another snapshot".into());
        }
        Ok(materializations)
    } else {
        let source_backed = SourceBackedPackage::from_read_at(source)?;
        let opc = source_backed.into_opc_package()?;
        let materializations = opc.part_count();
        let package = litchi_xlsx::Package::from_opc(opc)?;
        let workbook = package.into_workbook()?;
        let mut edit = workbook.edit()?;
        edit.replace_defined_names(xlsx_defined_names_target())?;
        let commit = edit.commit()?;
        if commit.patch().is_empty() {
            return Err("XLSX eager defined-name edit produced an empty patch".into());
        }
        commit.workbook().write_to(writer)?;
        Ok(materializations)
    }
}

fn run_xlsx_defined_names_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_DEFINED_NAMES_SOURCE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_defined_names_edit_save()
    {
        return Err("XLSX defined-name case requires its fixed media-rich corpus".into());
    }
    let source_backed = case == Case::XlsxSourceBackedDefinedNamesEditSave;
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations =
        publish_xlsx_defined_names_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        1
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err("XLSX defined-name edit materialized an unexpected Part count".into());
    }
    verify_xlsx_defined_names_edit_output(corpus, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("XLSX defined-name sequential output ceiling overflows u64")?;
    let payload_ranges = xlsx_calculation_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations = publish_xlsx_defined_names_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("XLSX defined-name edit differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("XLSX defined-name edit exceeded the sequential sink write bound".into());
        }
        verify_xlsx_defined_names_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("XLSX defined-name output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("XLSX defined-name edit performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLSX defined-name measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn verify_xlsx_calculation_metadata_edit_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let reopened = litchi_xlsx::Package::from_slice(output)?;
    if reopened
        .calculation_metadata()?
        .properties()
        .ok_or("XLSX calculation-edit output has no calcPr")?
        .calculation_id()
        != 91
    {
        return Err("XLSX calculation-edit output has unexpected calculation ID".into());
    }

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("XLSX calculation-edit package topology differs from source".into());
    }
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("XLSX calculation-edit Part metadata differs from source".into());
        }
        if source_part.partname() == &target_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("XLSX calculation-edit workbook XML did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("XLSX calculation-edit changed an unselected Part payload".into());
        }
    }
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        if candidate.get_part(&uri)?.blob() != xlsx_calculation_media_payload(index) {
            return Err("XLSX calculation-edit media readback differs from specification".into());
        }
    }
    Ok(())
}

fn publish_xlsx_calculation_metadata_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    let properties =
        litchi_xlsx::calculation_properties::Properties::new().with_calculation_id(Some(91));
    if source_backed {
        let editor = litchi_xlsx::calculation_properties::SourceBackedEditor::from_read_at(source)?;
        let mut edit = editor.edit();
        if !edit.set_properties(properties) {
            return Err("XLSX source-backed calculation edit reported an exact no-op".into());
        }
        let commit = edit.commit()?;
        if !commit.changed()
            || commit.patch().is_empty()
            || commit.patch().inverse().after() != commit.patch().before()
        {
            return Err("XLSX source-backed calculation edit produced an invalid patch".into());
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_commit_to_stream(writer, &commit)?;
        if published != *commit.snapshot() {
            return Err("XLSX source-backed calculation edit published another snapshot".into());
        }
        Ok(materializations)
    } else {
        let package = SourceBackedPackage::from_read_at(source)?;
        let opc = package.into_opc_package()?;
        let materializations = opc.part_count();
        let mut package = litchi_xlsx::Package::from_opc(opc)?;
        let mut edit = package.edit_calculation_metadata()?;
        if !edit.set_properties(properties) {
            return Err("XLSX eager calculation edit reported an exact no-op".into());
        }
        let commit = edit.commit()?;
        if !commit.changed() || commit.patch().is_empty() {
            return Err("XLSX eager calculation edit produced an invalid patch".into());
        }
        package.write_to(writer)?;
        Ok(materializations)
    }
}

fn run_xlsx_calculation_metadata_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_CALC_SOURCE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_calculation_metadata_edit_save()
    {
        return Err("XLSX calculation-edit case requires its fixed media-rich corpus".into());
    }
    let source_backed = case == Case::XlsxSourceBackedCalculationMetadataEditSave;
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations =
        publish_xlsx_calculation_metadata_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        1
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err("XLSX calculation edit materialized an unexpected Part count".into());
    }
    verify_xlsx_calculation_metadata_edit_output(corpus, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("XLSX calculation-edit sequential output ceiling overflows u64")?;
    let payload_ranges = xlsx_calculation_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations =
            publish_xlsx_calculation_metadata_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("XLSX calculation edit differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("XLSX calculation edit exceeded the sequential sink write bound".into());
        }
        verify_xlsx_calculation_metadata_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("XLSX calculation-edit output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("XLSX calculation edit performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLSX calculation-edit measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn xlsx_page_break_target() -> Result<litchi_xlsx::page_breaks::Collection, Box<dyn Error>> {
    Ok(litchi_xlsx::page_breaks::Collection::horizontal([
        litchi_xlsx::page_breaks::Break::new(100, 0, 16_383)?.with_manual(true),
    ])?)
}

fn verify_xlsx_page_break_edit_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let reopened = litchi_xlsx::Package::from_slice(output)?;
    let page_breaks = reopened.page_breaks("Sheet1")?;
    if page_breaks.page_breaks().horizontal() != Some(&xlsx_page_break_target()?)
        || page_breaks.page_breaks().vertical().is_some()
    {
        return Err("XLSX page-break output has unexpected authored breaks".into());
    }
    if reopened
        .calculation_metadata()?
        .properties()
        .ok_or("XLSX page-break output has no calcPr")?
        .calculation_id()
        != 7
    {
        return Err("XLSX page-break output changed calculation metadata".into());
    }

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("XLSX page-break package topology differs from source".into());
    }
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("XLSX page-break Part metadata differs from source".into());
        }
        if source_part.partname() == &target_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("XLSX page-break worksheet XML did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("XLSX page-break edit changed an unselected Part payload".into());
        }
    }
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        if candidate.get_part(&uri)?.blob() != xlsx_calculation_media_payload(index) {
            return Err("XLSX page-break media readback differs from specification".into());
        }
    }
    Ok(())
}

fn publish_xlsx_page_break_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    let horizontal = xlsx_page_break_target()?;
    if source_backed {
        let editor = litchi_xlsx::page_breaks::SourceBackedEditor::from_read_at(source)?;
        let mut edit = editor.edit("Sheet1")?;
        if !edit.set_horizontal(horizontal)? {
            return Err("XLSX source-backed page-break edit reported an exact no-op".into());
        }
        let commit = edit.commit()?;
        if !commit.changed() || commit.patch().is_empty() {
            return Err("XLSX source-backed page-break edit produced an invalid patch".into());
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_commit_to_stream(writer, &commit)?;
        if published.source_xml() != commit.snapshot().source_xml()
            || published.page_breaks() != commit.snapshot().page_breaks()
        {
            return Err("XLSX source-backed page-break edit published another snapshot".into());
        }
        Ok(materializations)
    } else {
        let package = SourceBackedPackage::from_read_at(source)?;
        let opc = package.into_opc_package()?;
        let materializations = opc.part_count();
        let mut package = litchi_xlsx::Package::from_opc(opc)?;
        let mut edit = package.edit_page_breaks("Sheet1")?;
        if !edit.set_horizontal(horizontal)? {
            return Err("XLSX eager page-break edit reported an exact no-op".into());
        }
        let commit = edit.commit()?;
        if !commit.changed() || commit.patch().is_empty() {
            return Err("XLSX eager page-break edit produced an invalid patch".into());
        }
        package.write_to(writer)?;
        Ok(materializations)
    }
}

fn run_xlsx_page_break_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_PAGE_BREAK_SOURCE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_page_break_edit_save()
    {
        return Err("XLSX page-break case requires its fixed media-rich corpus".into());
    }
    let source_backed = case == Case::XlsxSourceBackedPageBreakEditSave;
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations =
        publish_xlsx_page_break_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        2
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err("XLSX page-break edit materialized an unexpected Part count".into());
    }
    verify_xlsx_page_break_edit_output(corpus, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("XLSX page-break sequential output ceiling overflows u64")?;
    let payload_ranges = xlsx_calculation_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations = publish_xlsx_page_break_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("XLSX page-break edit differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("XLSX page-break edit exceeded the sequential sink write bound".into());
        }
        verify_xlsx_page_break_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("XLSX page-break output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("XLSX page-break edit performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLSX page-break measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn xlsx_page_margin_target() -> Result<litchi_xlsx::page_margins::Margins, Box<dyn Error>> {
    use litchi_xlsx::page_margins::{Margins, PageMargin};

    Ok(Margins::new(
        PageMargin::from_inches(0.7)?,
        PageMargin::from_inches(0.8)?,
        PageMargin::from_inches(1.0)?,
        PageMargin::from_inches(1.1)?,
        PageMargin::from_inches(0.3)?,
        PageMargin::from_inches(0.4)?,
    ))
}

fn verify_xlsx_page_margin_edit_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let reopened = litchi_xlsx::Package::from_slice(output)?;
    let workbook = reopened.workbook()?;
    let sheet = workbook
        .sheet("Sheet1")?
        .ok_or("XLSX page-margin output has no Sheet1")?;
    if sheet.page_margins()? != Some(xlsx_page_margin_target()?) {
        return Err("XLSX page-margin output has unexpected authored margins".into());
    }
    if reopened
        .calculation_metadata()?
        .properties()
        .ok_or("XLSX page-margin output has no calcPr")?
        .calculation_id()
        != 7
    {
        return Err("XLSX page-margin output changed calculation metadata".into());
    }

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("XLSX page-margin package topology differs from source".into());
    }
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("XLSX page-margin Part metadata differs from source".into());
        }
        if source_part.partname() == &target_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("XLSX page-margin worksheet XML did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("XLSX page-margin edit changed an unselected Part payload".into());
        }
    }
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        if candidate.get_part(&uri)?.blob() != xlsx_calculation_media_payload(index) {
            return Err("XLSX page-margin media readback differs from specification".into());
        }
    }
    Ok(())
}

fn publish_xlsx_page_margin_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    let margins = xlsx_page_margin_target()?;
    if source_backed {
        let editor = litchi_xlsx::page_margins::SourceBackedEditor::from_read_at(source)?;
        let mut edit = editor.edit("Sheet1")?;
        if !edit.set(margins) {
            return Err("XLSX source-backed page-margin edit reported an exact no-op".into());
        }
        let commit = edit.commit()?;
        if !commit.changed() || commit.patch().is_empty() {
            return Err("XLSX source-backed page-margin edit produced an invalid patch".into());
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_commit_to_stream(writer, &commit)?;
        if published.source_xml() != commit.snapshot().source_xml()
            || published.page_margins() != commit.snapshot().page_margins()
        {
            return Err("XLSX source-backed page-margin edit published another snapshot".into());
        }
        Ok(materializations)
    } else {
        let package = SourceBackedPackage::from_read_at(source)?;
        let opc = package.into_opc_package()?;
        let materializations = opc.part_count();
        let package = litchi_xlsx::Package::from_opc(opc)?;
        let workbook = package.into_workbook()?;
        let mut edit = workbook.edit()?;
        if edit.put_page_margins("Sheet1", margins)?.is_none() {
            return Err("XLSX eager page-margin worksheet selector did not resolve".into());
        }
        let commit = edit.commit()?;
        if commit.patch().is_empty() {
            return Err("XLSX eager page-margin edit produced an empty patch".into());
        }
        commit.workbook().write_to(writer)?;
        Ok(materializations)
    }
}

fn run_xlsx_page_margin_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_PAGE_MARGIN_SOURCE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_page_margin_edit_save()
    {
        return Err("XLSX page-margin case requires its fixed media-rich corpus".into());
    }
    let source_backed = case == Case::XlsxSourceBackedPageMarginEditSave;
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations =
        publish_xlsx_page_margin_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        2
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err("XLSX page-margin edit materialized an unexpected Part count".into());
    }
    verify_xlsx_page_margin_edit_output(corpus, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("XLSX page-margin sequential output ceiling overflows u64")?;
    let payload_ranges = xlsx_calculation_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations = publish_xlsx_page_margin_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("XLSX page-margin edit differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("XLSX page-margin edit exceeded the sequential sink write bound".into());
        }
        verify_xlsx_page_margin_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("XLSX page-margin output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("XLSX page-margin edit performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLSX page-margin measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn xlsx_page_setup_target() -> Result<litchi_xlsx::page_setup::Setup, Box<dyn Error>> {
    use litchi_xlsx::page_setup::{Order, Orientation, Paper, Scale, Setup};

    let mut setup = Setup::new(Paper::A4);
    setup.scale = Some(Scale::new(85)?);
    setup.order = Some(Order::OverThenDown);
    setup.orientation = Some(Orientation::Landscape);
    setup.use_printer_defaults = Some(false);
    Ok(setup)
}

fn verify_xlsx_page_setup_edit_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let reopened = litchi_xlsx::Package::from_slice(output)?;
    let workbook = reopened.workbook()?;
    let sheet = workbook
        .sheet("Sheet1")?
        .ok_or("XLSX page-setup output has no Sheet1")?;
    if sheet.page_setup()? != Some(xlsx_page_setup_target()?) {
        return Err("XLSX page-setup output has unexpected authored settings".into());
    }
    if reopened
        .calculation_metadata()?
        .properties()
        .ok_or("XLSX page-setup output has no calcPr")?
        .calculation_id()
        != 7
    {
        return Err("XLSX page-setup output changed calculation metadata".into());
    }

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("XLSX page-setup package topology differs from source".into());
    }
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    if litchi_xlsx::parse_worksheet_page_setup_relationship_id(
        candidate.get_part(&target_uri)?.blob(),
    )?
    .is_some()
    {
        return Err("XLSX page-setup output unexpectedly references printer settings".into());
    }
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("XLSX page-setup Part metadata differs from source".into());
        }
        if source_part.partname() == &target_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("XLSX page-setup worksheet XML did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("XLSX page-setup edit changed an unselected Part payload".into());
        }
    }
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        if candidate.get_part(&uri)?.blob() != xlsx_calculation_media_payload(index) {
            return Err("XLSX page-setup media readback differs from specification".into());
        }
    }
    Ok(())
}

fn publish_xlsx_page_setup_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    let setup = xlsx_page_setup_target()?;
    if source_backed {
        let editor = litchi_xlsx::page_setup::SourceBackedEditor::from_read_at(source)?;
        let mut edit = editor.edit("Sheet1")?;
        if !edit.set(setup) {
            return Err("XLSX source-backed page-setup edit reported an exact no-op".into());
        }
        let commit = edit.commit()?;
        if !commit.changed() || commit.patch().is_empty() {
            return Err("XLSX source-backed page-setup edit produced an invalid patch".into());
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_commit_to_stream(writer, &commit)?;
        if published.source_xml() != commit.snapshot().source_xml()
            || published.page_setup() != commit.snapshot().page_setup()
        {
            return Err("XLSX source-backed page-setup edit published another snapshot".into());
        }
        Ok(materializations)
    } else {
        let package = SourceBackedPackage::from_read_at(source)?;
        let opc = package.into_opc_package()?;
        let materializations = opc.part_count();
        let package = litchi_xlsx::Package::from_opc(opc)?;
        let workbook = package.into_workbook()?;
        let mut edit = workbook.edit()?;
        if edit.put_page_setup("Sheet1", setup)?.is_none() {
            return Err("XLSX eager page-setup worksheet selector did not resolve".into());
        }
        let commit = edit.commit()?;
        if commit.patch().is_empty() {
            return Err("XLSX eager page-setup edit produced an empty patch".into());
        }
        commit.workbook().write_to(writer)?;
        Ok(materializations)
    }
}

fn run_xlsx_page_setup_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_PAGE_SETUP_SOURCE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_page_setup_edit_save()
    {
        return Err("XLSX page-setup case requires its fixed media-rich corpus".into());
    }
    let source_backed = case == Case::XlsxSourceBackedPageSetupEditSave;
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations =
        publish_xlsx_page_setup_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        2
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err("XLSX page-setup edit materialized an unexpected Part count".into());
    }
    verify_xlsx_page_setup_edit_output(corpus, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("XLSX page-setup sequential output ceiling overflows u64")?;
    let payload_ranges = xlsx_calculation_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations = publish_xlsx_page_setup_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("XLSX page-setup edit differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("XLSX page-setup edit exceeded the sequential sink write bound".into());
        }
        verify_xlsx_page_setup_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("XLSX page-setup output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("XLSX page-setup edit performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLSX page-setup measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn xlsx_print_options_target() -> litchi_xlsx::print_options::PrintOptions {
    let mut options = litchi_xlsx::print_options::PrintOptions::new();
    options
        .set_horizontal_centered(true)
        .set_print_headings(true)
        .set_print_grid_lines(true);
    options
}

fn verify_xlsx_print_options_edit_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let reopened = litchi_xlsx::Package::from_slice(output)?;
    let workbook = reopened.workbook()?;
    let sheet = workbook
        .sheet("Sheet1")?
        .ok_or("XLSX print-options output has no Sheet1")?;
    if sheet.print_options()? != Some(xlsx_print_options_target()) {
        return Err("XLSX print-options output has unexpected authored options".into());
    }
    if reopened
        .calculation_metadata()?
        .properties()
        .ok_or("XLSX print-options output has no calcPr")?
        .calculation_id()
        != 7
    {
        return Err("XLSX print-options output changed calculation metadata".into());
    }

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("XLSX print-options package topology differs from source".into());
    }
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("XLSX print-options Part metadata differs from source".into());
        }
        if source_part.partname() == &target_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("XLSX print-options worksheet XML did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("XLSX print-options edit changed an unselected Part payload".into());
        }
    }
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        if candidate.get_part(&uri)?.blob() != xlsx_calculation_media_payload(index) {
            return Err("XLSX print-options media readback differs from specification".into());
        }
    }
    Ok(())
}

fn publish_xlsx_print_options_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    let options = xlsx_print_options_target();
    if source_backed {
        let editor = litchi_xlsx::print_options::SourceBackedEditor::from_read_at(source)?;
        let mut edit = editor.edit("Sheet1")?;
        if !edit.set(options) {
            return Err("XLSX source-backed print-options edit reported an exact no-op".into());
        }
        let commit = edit.commit()?;
        if !commit.changed() || commit.patch().is_empty() {
            return Err("XLSX source-backed print-options edit produced an invalid patch".into());
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_commit_to_stream(writer, &commit)?;
        if published.source_xml() != commit.snapshot().source_xml()
            || published.print_options() != commit.snapshot().print_options()
        {
            return Err("XLSX source-backed print-options edit published another snapshot".into());
        }
        Ok(materializations)
    } else {
        let package = SourceBackedPackage::from_read_at(source)?;
        let opc = package.into_opc_package()?;
        let materializations = opc.part_count();
        let package = litchi_xlsx::Package::from_opc(opc)?;
        let workbook = package.into_workbook()?;
        let mut edit = workbook.edit()?;
        if edit.put_print_options("Sheet1", options)?.is_none() {
            return Err("XLSX eager print-options worksheet selector did not resolve".into());
        }
        let commit = edit.commit()?;
        if commit.patch().is_empty() {
            return Err("XLSX eager print-options edit produced an empty patch".into());
        }
        commit.workbook().write_to(writer)?;
        Ok(materializations)
    }
}

fn run_xlsx_print_options_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_PRINT_OPTIONS_SOURCE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_print_options_edit_save()
    {
        return Err("XLSX print-options case requires its fixed media-rich corpus".into());
    }
    let source_backed = case == Case::XlsxSourceBackedPrintOptionsEditSave;
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations =
        publish_xlsx_print_options_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        2
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err("XLSX print-options edit materialized an unexpected Part count".into());
    }
    verify_xlsx_print_options_edit_output(corpus, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("XLSX print-options sequential output ceiling overflows u64")?;
    let payload_ranges = xlsx_calculation_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations = publish_xlsx_print_options_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("XLSX print-options edit differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("XLSX print-options edit exceeded the sequential sink write bound".into());
        }
        verify_xlsx_print_options_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("XLSX print-options output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("XLSX print-options edit performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLSX print-options measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn xlsx_sheet_protection_target() -> Result<litchi_xlsx::sheet_protection::Metadata, Box<dyn Error>>
{
    use litchi_xlsx::sheet_protection::{
        Metadata, ProtectedRange, ProtectedRangeCollection, ProtectedRangeSource, Protection,
        ProtectionPasswordVerifier, ProtectionRangeSqref, StrongProtectionPasswordVerifier,
    };

    let mut protection = Protection::new();
    protection.set_verifier(Some(ProtectionPasswordVerifier::Legacy(0x83af)))?;
    protection.set_sheet_locked(true);
    protection.set_objects_locked(true);
    protection.set_scenarios_locked(true);
    protection.set_auto_filter_locked(false);
    protection.set_sort_locked(false);

    let mut input_range = ProtectedRange::new(
        ProtectedRangeSource::Core,
        "Input cells",
        ProtectionRangeSqref::parse("A1:B4 D6")?,
    )?;
    input_range.set_verifier(Some(ProtectionPasswordVerifier::Legacy(0x1a2b)))?;
    let mut review_range = ProtectedRange::new(
        ProtectedRangeSource::Office2010,
        "Review cells",
        ProtectionRangeSqref::parse("C3:E7")?,
    )?;
    review_range.set_verifier(Some(ProtectionPasswordVerifier::Strong(
        StrongProtectionPasswordVerifier::new(
            "SHA-512",
            (0_u8..64).collect(),
            (64_u8..80).collect(),
            100_000,
        )?,
    )))?;
    review_range.set_security_descriptor(Some("D:(A;;FA;;;SY)".to_owned()))?;

    let mut metadata = Metadata::new();
    metadata.set_sheet_protection(Some(protection))?;
    metadata.set_protected_range_collections(vec![
        ProtectedRangeCollection::new(ProtectedRangeSource::Core, vec![input_range])?,
        ProtectedRangeCollection::new(ProtectedRangeSource::Office2010, vec![review_range])?,
    ])?;
    Ok(metadata)
}

fn verify_xlsx_sheet_protection_edit_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let reopened = litchi_xlsx::Package::from_slice(output)?;
    let workbook = reopened.workbook()?;
    let sheet = workbook
        .sheet("Sheet1")?
        .ok_or("XLSX sheet-protection output has no Sheet1")?;
    if sheet.protection()? != xlsx_sheet_protection_target()? {
        return Err("XLSX sheet-protection output has unexpected authored metadata".into());
    }
    if reopened
        .calculation_metadata()?
        .properties()
        .ok_or("XLSX sheet-protection output has no calcPr")?
        .calculation_id()
        != 7
    {
        return Err("XLSX sheet-protection output changed calculation metadata".into());
    }

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("XLSX sheet-protection package topology differs from source".into());
    }
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("XLSX sheet-protection Part metadata differs from source".into());
        }
        if source_part.partname() == &target_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("XLSX sheet-protection worksheet XML did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("XLSX sheet-protection edit changed an unselected Part payload".into());
        }
    }
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        if candidate.get_part(&uri)?.blob() != xlsx_calculation_media_payload(index) {
            return Err("XLSX sheet-protection media readback differs from specification".into());
        }
    }
    Ok(())
}

fn publish_xlsx_sheet_protection_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    let metadata = xlsx_sheet_protection_target()?;
    if source_backed {
        let editor = litchi_xlsx::sheet_protection::SourceBackedEditor::from_read_at(source)?;
        let mut edit = editor.edit("Sheet1")?;
        if !edit.set(metadata)? {
            return Err("XLSX source-backed sheet-protection edit reported an exact no-op".into());
        }
        let commit = edit.commit()?;
        if !commit.changed() || commit.patch().is_empty() {
            return Err(
                "XLSX source-backed sheet-protection edit produced an invalid patch".into(),
            );
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_commit_to_stream(writer, &commit)?;
        if published.source_xml() != commit.snapshot().source_xml()
            || published.metadata() != commit.snapshot().metadata()
        {
            return Err(
                "XLSX source-backed sheet-protection edit published another snapshot".into(),
            );
        }
        Ok(materializations)
    } else {
        let package = SourceBackedPackage::from_read_at(source)?;
        let mut opc = package.into_opc_package()?;
        let materializations = opc.part_count();
        let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
        let updated = litchi_xlsx::sheet_protection::replace_protection(
            opc.get_part(&worksheet_uri)?.blob(),
            &metadata,
        )?;
        opc.get_part_mut(&worksheet_uri)?.set_blob(updated);
        PackageWriter::write_to_stream(writer, &opc)?;
        Ok(materializations)
    }
}

fn run_xlsx_sheet_protection_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_SHEET_PROTECTION_SOURCE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_sheet_protection_edit_save()
    {
        return Err("XLSX sheet-protection case requires its fixed media-rich corpus".into());
    }
    let source_backed = case == Case::XlsxSourceBackedSheetProtectionEditSave;
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations =
        publish_xlsx_sheet_protection_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        2
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err("XLSX sheet-protection edit materialized an unexpected Part count".into());
    }
    verify_xlsx_sheet_protection_edit_output(corpus, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("XLSX sheet-protection sequential output ceiling overflows u64")?;
    let payload_ranges = xlsx_calculation_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations =
            publish_xlsx_sheet_protection_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("XLSX sheet-protection edit differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err(
                "XLSX sheet-protection edit exceeded the sequential sink write bound".into(),
            );
        }
        verify_xlsx_sheet_protection_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("XLSX sheet-protection output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("XLSX sheet-protection edit performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLSX sheet-protection measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn xlsx_data_validation_values(
    updated: bool,
) -> Result<Vec<litchi_xlsx::data_validation::Collection>, Box<dyn Error>> {
    use litchi_xlsx::data_validation::{
        Collection, Formula, ListSource, Source, Sqref, Validation, ValidationOperator,
        ValidationType,
    };

    let mut core = Validation::new(
        Source::Core,
        ValidationType::Whole,
        Sqref::parse(if updated { "B2:C8" } else { "A1:B4" })?,
    );
    core.set_operator(ValidationOperator::Between);
    core.set_allow_blank(true);
    core.set_show_error_message(true);
    core.set_error_title(Some("Whole numbers".to_owned()))?;
    core.set_formula1(Some(ListSource::Formula(Formula::new(if updated {
        "2"
    } else {
        "1"
    })?)))?;
    core.set_formula2(Some(Formula::new(if updated { "20" } else { "10" })?))?;

    let mut extension = Validation::new(
        Source::Office2010,
        ValidationType::List,
        Sqref::parse(if updated { "E2:E12" } else { "D1:D8" })?
            .with_office2010_flags(true, false, true, true)?,
    );
    extension.set_show_drop_down(true);
    extension.set_show_input_message(true);
    extension.set_prompt(Some(
        if updated {
            "Choose amber or violet"
        } else {
            "Choose red or blue"
        }
        .to_owned(),
    ))?;
    extension.set_formula1(Some(ListSource::QuotedList(
        if updated {
            "\"amber,violet\""
        } else {
            "\"red,blue\""
        }
        .to_owned(),
    )))?;
    extension.set_uid(Some("{12345678-1234-1234-1234-123456789ABC}".to_owned()))?;

    let core = Collection::new(Source::Core, vec![core])?;
    let mut office2010 = Collection::new(Source::Office2010, vec![extension])?;
    office2010.set_disable_prompts(updated);
    office2010.set_window(
        Some(if updated { 11 } else { 7 }),
        Some(if updated { 13 } else { 9 }),
    )?;
    Ok(vec![core, office2010])
}

fn xlsx_auto_filter_value(
    updated: bool,
) -> Result<litchi_xlsx::auto_filter::Definition, Box<dyn Error>> {
    use litchi_xlsx::auto_filter::{
        Calendar, Column, Condition, Definition, Item, Payload, Range, State, Values,
    };
    use litchi_xlsx::sort::{SortBy, SortMethod};

    let mut definition = Definition::new(Some(Range::new(if updated {
        "A1:C100"
    } else {
        "A1:C80"
    })?));
    let mut column = Column::new(0)?;
    column.set_payload(Some(Payload::Values(Values::new(
        false,
        Calendar::None,
        vec![
            Item::Value(if updated { "amber" } else { "red" }.to_owned()),
            Item::Value(if updated { "violet" } else { "blue" }.to_owned()),
        ],
    )?)));
    definition.columns.push(column);
    let condition = Condition::new(
        Range::new(if updated { "B2:B100" } else { "B2:B80" })?,
        updated,
        SortBy::Value,
    );
    let sort = State::new(
        Range::new(if updated { "A2:C100" } else { "A2:C80" })?,
        false,
        updated,
        Some(SortMethod::None),
        vec![condition],
    )?;
    definition.set_sort_state(Some(sort))?;
    Ok(definition)
}

fn verify_xlsx_auto_filter_edit_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let reopened = litchi_xlsx::Package::from_slice(output)?;
    let workbook = reopened.workbook()?;
    let sheet = workbook
        .sheet("Sheet1")?
        .ok_or("XLSX auto-filter output has no Sheet1")?;
    if sheet.auto_filter()? != Some(xlsx_auto_filter_value(true)?) {
        return Err("XLSX auto-filter output has unexpected authored definition".into());
    }
    if reopened
        .calculation_metadata()?
        .properties()
        .ok_or("XLSX auto-filter output has no calcPr")?
        .calculation_id()
        != 7
    {
        return Err("XLSX auto-filter output changed calculation metadata".into());
    }

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("XLSX auto-filter package topology differs from source".into());
    }
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("XLSX auto-filter Part metadata differs from source".into());
        }
        if source_part.partname() == &target_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("XLSX auto-filter worksheet XML did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("XLSX auto-filter edit changed an unselected Part payload".into());
        }
    }
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        if candidate.get_part(&uri)?.blob() != xlsx_calculation_media_payload(index) {
            return Err("XLSX auto-filter media readback differs from specification".into());
        }
    }
    Ok(())
}

fn publish_xlsx_auto_filter_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    let value = xlsx_auto_filter_value(true)?;
    if source_backed {
        let editor = litchi_xlsx::auto_filter::SourceBackedEditor::from_read_at(source)?;
        let mut edit = editor.edit("Sheet1")?;
        if !edit.set(value)? {
            return Err("XLSX source-backed auto-filter edit reported an exact no-op".into());
        }
        let commit = edit.commit()?;
        if !commit.changed() || commit.patch().is_empty() {
            return Err("XLSX source-backed auto-filter edit produced an invalid patch".into());
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_commit_to_stream(writer, &commit)?;
        if published.source_xml() != commit.snapshot().source_xml()
            || published.auto_filter() != commit.snapshot().auto_filter()
        {
            return Err("XLSX source-backed auto-filter edit published another snapshot".into());
        }
        Ok(materializations)
    } else {
        let package = SourceBackedPackage::from_read_at(source)?;
        let mut opc = package.into_opc_package()?;
        let materializations = opc.part_count();
        let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
        let snapshot = litchi_xlsx::auto_filter::Snapshot::load(&opc, "Sheet1")?;
        if snapshot.auto_filter() != Some(&xlsx_auto_filter_value(false)?) {
            return Err("XLSX eager auto-filter source closure differs from specification".into());
        }
        let updated = litchi_xlsx::auto_filter::replace_auto_filter(
            opc.get_part(&worksheet_uri)?.blob(),
            Some(&value),
        )?;
        opc.get_part_mut(&worksheet_uri)?.set_blob(updated);
        PackageWriter::write_to_stream(writer, &opc)?;
        Ok(materializations)
    }
}

fn run_xlsx_auto_filter_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_AUTO_FILTER_SOURCE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_auto_filter_edit_save()
    {
        return Err("XLSX auto-filter case requires its fixed media-rich corpus".into());
    }
    let source_backed = case == Case::XlsxSourceBackedAutoFilterEditSave;
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations =
        publish_xlsx_auto_filter_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        3
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err("XLSX auto-filter edit materialized an unexpected Part count".into());
    }
    verify_xlsx_auto_filter_edit_output(corpus, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("XLSX auto-filter sequential output ceiling overflows u64")?;
    let payload_ranges = xlsx_calculation_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations = publish_xlsx_auto_filter_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("XLSX auto-filter edit differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err("XLSX auto-filter edit exceeded the sequential sink write bound".into());
        }
        verify_xlsx_auto_filter_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("XLSX auto-filter output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("XLSX auto-filter edit performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLSX auto-filter measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn xlsx_conditional_formatting_values(
    updated: bool,
) -> Result<Vec<litchi_xlsx::conditional_formatting::Formatting>, Box<dyn Error>> {
    use litchi_xlsx::conditional_formatting::{Formatting, Kind, Range, Rule};

    let specifications = if updated {
        [
            ("B2:B256", 1, "$B2>=50", true),
            ("D2:F256", 2, "MOD(ROW(),2)=0", false),
            ("H2:H256", 3, "H2=\"ready\"", false),
        ]
    } else {
        [
            ("A1:A128", 1, "$A1>10", false),
            ("C1:E128", 2, "MOD(COLUMN(),2)=0", true),
            ("G1:G128", 3, "G1=\"pending\"", false),
        ]
    };
    specifications
        .into_iter()
        .map(|(range, priority, formula, stop_if_true)| {
            let mut rule = Rule::new(Kind::Expression, priority)?;
            rule.push_formula(formula)?;
            rule.stop_if_true = stop_if_true;
            Formatting::new(vec![Range::new(range)?], vec![rule]).map_err(Into::into)
        })
        .collect()
}

#[derive(Debug, PartialEq, Eq)]
struct RawZipMember {
    local: Vec<u8>,
    central_without_offset: Vec<u8>,
}

fn raw_zip_members(bytes: &[u8]) -> Result<BTreeMap<String, RawZipMember>, Box<dyn Error>> {
    let archive = ZipArchive::from_slice(bytes)?.into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer)?;
    let mut records = archive.entries(&mut buffer);
    index
        .entries()
        .iter()
        .map(|preserved| {
            let record = records
                .next_entry()?
                .ok_or("ZIP preservation index has no matching central record")?;
            let name = record.file_path().try_normalize()?.as_ref().to_owned();
            let local = bytes
                [preserved.local_span().start as usize..preserved.local_span().end as usize]
                .to_vec();
            let central = preserved.central_record();
            let mut central_without_offset =
                bytes[central.start as usize..central.end as usize].to_vec();
            central_without_offset[42..46].fill(0);
            Ok((
                name,
                RawZipMember {
                    local,
                    central_without_offset,
                },
            ))
        })
        .collect()
}

fn verify_xlsx_conditional_formatting_raw_preservation(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let source = raw_zip_members(&corpus.archive)?;
    let candidate = raw_zip_members(output)?;
    if source.keys().ne(candidate.keys()) {
        return Err("XLSX conditional-formatting raw ZIP member set differs from source".into());
    }
    for (name, source_member) in source {
        if name != corpus.target_name && candidate.get(&name) != Some(&source_member) {
            return Err(format!(
                "XLSX conditional-formatting changed raw unselected ZIP member {name}"
            )
            .into());
        }
    }
    Ok(())
}

fn verify_xlsx_conditional_formatting_patch(corpus: &Corpus) -> Result<(), Box<dyn Error>> {
    let editor = litchi_xlsx::conditional_formatting::SourceBackedEditor::from_read_at(Arc::new(
        OwnedSource::new(corpus.archive.clone()),
    ))?;
    let mut edit = editor.edit("Sheet1")?;
    if edit.collections() != xlsx_conditional_formatting_values(false)? {
        return Err("XLSX conditional-formatting patch source differs from specification".into());
    }
    if !edit.set_collections(xlsx_conditional_formatting_values(true)?)? {
        return Err("XLSX conditional-formatting patch operation reported a no-op".into());
    }
    let commit = edit.commit()?;
    let inverse = commit.patch().inverse();
    if !commit.changed()
        || commit.patch().is_empty()
        || inverse.before().source_xml() != commit.patch().after().source_xml()
        || inverse.after().source_xml() != commit.patch().before().source_xml()
        || inverse.before().collections() != commit.patch().after().collections()
        || inverse.after().collections() != commit.patch().before().collections()
    {
        return Err("XLSX conditional-formatting patch/inverse is not exact".into());
    }

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let mut replay = source.clone();
    commit.patch().apply(&mut replay)?;
    if litchi_xlsx::conditional_formatting::Snapshot::load(&replay, "Sheet1")?.collections()
        != xlsx_conditional_formatting_values(true)?
    {
        return Err("XLSX conditional-formatting patch replay changed its target".into());
    }
    inverse.apply(&mut replay)?;
    for source_part in source.iter_parts() {
        let restored = replay.get_part(source_part.partname())?;
        if restored.content_type() != source_part.content_type()
            || relationship_signatures(restored.rels())
                != relationship_signatures(source_part.rels())
            || restored.blob() != source_part.blob()
        {
            return Err("XLSX conditional-formatting inverse did not restore exact Parts".into());
        }
    }
    Ok(())
}

fn verify_xlsx_conditional_formatting_edit_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let reopened = litchi_xlsx::Package::from_slice(output)?;
    let workbook = reopened.workbook()?;
    let sheet = workbook
        .sheet("Sheet1")?
        .ok_or("XLSX conditional-formatting output has no Sheet1")?;
    if sheet.conditional_formattings()? != xlsx_conditional_formatting_values(true)? {
        return Err(
            "XLSX conditional-formatting output has unexpected authored collections".into(),
        );
    }
    if reopened
        .calculation_metadata()?
        .properties()
        .ok_or("XLSX conditional-formatting output has no calcPr")?
        .calculation_id()
        != 7
    {
        return Err("XLSX conditional-formatting output changed calculation metadata".into());
    }

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("XLSX conditional-formatting package topology differs from source".into());
    }
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("XLSX conditional-formatting Part metadata differs from source".into());
        }
        if source_part.partname() == &target_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("XLSX conditional-formatting worksheet XML did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("XLSX conditional-formatting changed an unselected Part payload".into());
        }
    }
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        if candidate.get_part(&uri)?.blob() != xlsx_calculation_media_payload(index) {
            return Err(
                "XLSX conditional-formatting media readback differs from specification".into(),
            );
        }
    }
    Ok(())
}

fn publish_xlsx_conditional_formatting_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    let values = xlsx_conditional_formatting_values(true)?;
    if source_backed {
        let editor = litchi_xlsx::conditional_formatting::SourceBackedEditor::from_read_at(source)?;
        let mut edit = editor.edit("Sheet1")?;
        if edit.collections() != xlsx_conditional_formatting_values(false)? {
            return Err(
                "XLSX source-backed conditional-formatting source differs from specification"
                    .into(),
            );
        }
        if !edit.set_collections(values)? {
            return Err(
                "XLSX source-backed conditional-formatting edit reported an exact no-op".into(),
            );
        }
        let commit = edit.commit()?;
        if !commit.changed() || commit.patch().is_empty() {
            return Err(
                "XLSX source-backed conditional-formatting edit produced an invalid patch".into(),
            );
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_commit_to_stream(writer, &commit)?;
        if published.source_xml() != commit.snapshot().source_xml()
            || published.collections() != commit.snapshot().collections()
        {
            return Err(
                "XLSX source-backed conditional-formatting edit published another snapshot".into(),
            );
        }
        Ok(materializations)
    } else {
        let package = SourceBackedPackage::from_read_at(source)?;
        let mut opc = package.into_opc_package()?;
        let materializations = opc.part_count();
        let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
        let snapshot = litchi_xlsx::conditional_formatting::Snapshot::load(&opc, "Sheet1")?;
        if snapshot.collections() != xlsx_conditional_formatting_values(false)? {
            return Err(
                "XLSX eager conditional-formatting source differs from specification".into(),
            );
        }
        let updated = litchi_xlsx::conditional_formatting::replace_conditional_formattings(
            opc.get_part(&worksheet_uri)?.blob(),
            &values,
            0,
        )?;
        opc.get_part_mut(&worksheet_uri)?.set_blob(updated);
        PackageWriter::write_to_stream(writer, &opc)?;
        Ok(materializations)
    }
}

fn run_xlsx_conditional_formatting_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_CONDITIONAL_FORMATTING_SOURCE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_conditional_formatting_edit_save()
    {
        return Err("XLSX conditional-formatting case requires its fixed media-rich corpus".into());
    }
    let source_backed = case == Case::XlsxSourceBackedConditionalFormattingEditSave;
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations =
        publish_xlsx_conditional_formatting_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        3
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err(
            "XLSX conditional-formatting edit materialized an unexpected Part count".into(),
        );
    }
    verify_xlsx_conditional_formatting_edit_output(corpus, &expected)?;
    verify_xlsx_conditional_formatting_raw_preservation(corpus, &expected)?;
    verify_xlsx_conditional_formatting_patch(corpus)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("XLSX conditional-formatting sequential output ceiling overflows u64")?;
    let maximum_source_bytes = u64::try_from(corpus.archive.len())?
        .checked_add(64 * 1024)
        .ok_or("XLSX conditional-formatting source-read ceiling overflows u64")?;
    let payload_ranges = xlsx_calculation_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations =
            publish_xlsx_conditional_formatting_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("XLSX conditional-formatting edit differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err(
                "XLSX conditional-formatting edit exceeded the sequential sink write bound".into(),
            );
        }
        verify_xlsx_conditional_formatting_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err(
                "XLSX conditional-formatting output digest differs from expected output".into(),
            );
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0
            || metrics.ordinary_payload_read_bytes == 0
            || metrics.read_bytes > maximum_source_bytes
        {
            return Err(
                "XLSX conditional-formatting edit exceeded its source-read contract".into(),
            );
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLSX conditional-formatting measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn verify_xlsx_data_validation_edit_output(
    corpus: &Corpus,
    output: &[u8],
) -> Result<(), Box<dyn Error>> {
    let reopened = litchi_xlsx::Package::from_slice(output)?;
    let workbook = reopened.workbook()?;
    let sheet = workbook
        .sheet("Sheet1")?
        .ok_or("XLSX data-validation output has no Sheet1")?;
    if sheet.data_validations()? != xlsx_data_validation_values(true)? {
        return Err("XLSX data-validation output has unexpected authored collections".into());
    }
    if reopened
        .calculation_metadata()?
        .properties()
        .ok_or("XLSX data-validation output has no calcPr")?
        .calculation_id()
        != 7
    {
        return Err("XLSX data-validation output changed calculation metadata".into());
    }

    let source = OpcPackage::from_bytes(&corpus.archive)?;
    let candidate = OpcPackage::from_bytes(output)?;
    if source.part_count() != corpus.manifest.entry_count
        || candidate.part_count() != source.part_count()
        || relationship_signatures(source.rels()) != relationship_signatures(candidate.rels())
    {
        return Err("XLSX data-validation package topology differs from source".into());
    }
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    for source_part in source.iter_parts() {
        let candidate_part = candidate.get_part(source_part.partname())?;
        if candidate_part.content_type() != source_part.content_type()
            || relationship_signatures(candidate_part.rels())
                != relationship_signatures(source_part.rels())
        {
            return Err("XLSX data-validation Part metadata differs from source".into());
        }
        if source_part.partname() == &target_uri {
            if source_part.blob() == candidate_part.blob() {
                return Err("XLSX data-validation worksheet XML did not change".into());
            }
        } else if source_part.blob() != candidate_part.blob() {
            return Err("XLSX data-validation edit changed an unselected Part payload".into());
        }
    }
    for index in 0..XLSX_CALC_MEDIA_ENTRY_COUNT {
        let uri = PackURI::new(format!("/xl/media/image{}.png", index + 1))?;
        if candidate.get_part(&uri)?.blob() != xlsx_calculation_media_payload(index) {
            return Err("XLSX data-validation media readback differs from specification".into());
        }
    }
    Ok(())
}

fn publish_xlsx_data_validation_edit<W: Write>(
    source: Arc<dyn ReadAt>,
    writer: W,
    source_backed: bool,
) -> Result<usize, Box<dyn Error>> {
    let values = xlsx_data_validation_values(true)?;
    if source_backed {
        let editor = litchi_xlsx::data_validation::SourceBackedEditor::from_read_at(source)?;
        let mut edit = editor.edit("Sheet1")?;
        if !edit.set_collections(values)? {
            return Err("XLSX source-backed data-validation edit reported an exact no-op".into());
        }
        let commit = edit.commit()?;
        if !commit.changed() || commit.patch().is_empty() {
            return Err("XLSX source-backed data-validation edit produced an invalid patch".into());
        }
        let materializations = usize::try_from(editor.cache_diagnostics().successful_loads)?;
        let published = editor.publish_commit_to_stream(writer, &commit)?;
        if published.source_xml() != commit.snapshot().source_xml()
            || published.collections() != commit.snapshot().collections()
        {
            return Err(
                "XLSX source-backed data-validation edit published another snapshot".into(),
            );
        }
        Ok(materializations)
    } else {
        let package = SourceBackedPackage::from_read_at(source)?;
        let mut opc = package.into_opc_package()?;
        let materializations = opc.part_count();
        let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
        let updated = litchi_xlsx::data_validation::replace_data_validation_collections(
            opc.get_part(&worksheet_uri)?.blob(),
            &values,
        )?;
        opc.get_part_mut(&worksheet_uri)?.set_blob(updated);
        PackageWriter::write_to_stream(writer, &opc)?;
        Ok(materializations)
    }
}

fn run_xlsx_data_validation_edit_save(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if corpus.manifest.generator != XLSX_DATA_VALIDATION_SOURCE_EDIT_CORPUS_GENERATOR
        || !case.is_xlsx_data_validation_edit_save()
    {
        return Err("XLSX data-validation case requires its fixed media-rich corpus".into());
    }
    let source_backed = case == Case::XlsxSourceBackedDataValidationEditSave;
    let expected_source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let mut expected = Vec::new();
    let expected_materializations =
        publish_xlsx_data_validation_edit(expected_source, &mut expected, source_backed)?;
    let required_materializations = if source_backed {
        2
    } else {
        corpus.manifest.entry_count
    };
    if expected == corpus.archive || expected_materializations != required_materializations {
        return Err("XLSX data-validation edit materialized an unexpected Part count".into());
    }
    verify_xlsx_data_validation_edit_output(corpus, &expected)?;
    let expected_digest = sha256_hex(&expected);
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("XLSX data-validation sequential output ceiling overflows u64")?;
    let payload_ranges = xlsx_calculation_payload_ranges(corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    let mut measured_digests = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = Arc::new(InstrumentedSource::new(
            corpus.archive.clone(),
            payload_ranges.clone(),
        ));
        let read_at: Arc<dyn ReadAt> = source.clone();
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        let materializations =
            publish_xlsx_data_validation_edit(read_at, &mut sink, source_backed)?;
        let duration = started.elapsed();

        if materializations != expected_materializations || sink.bytes != expected {
            return Err("XLSX data-validation edit differs between iterations".into());
        }
        if sink.summary().largest_write > 64 * 1024 {
            return Err(
                "XLSX data-validation edit exceeded the sequential sink write bound".into(),
            );
        }
        verify_xlsx_data_validation_edit_output(corpus, &sink.bytes)?;
        let digest = sha256_hex(&sink.bytes);
        if digest != expected_digest {
            return Err("XLSX data-validation output digest differs from expected output".into());
        }
        let metrics = source.snapshot();
        if metrics.ordinary_payload_read_calls == 0 || metrics.ordinary_payload_read_bytes == 0 {
            return Err("XLSX data-validation edit performed no ordinary source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record_opc(metrics, u64::try_from(materializations)?);
            sink_summaries.push(sink.summary());
            measured_digests.push(digest);
        }
        std::hint::black_box(&sink.bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if measured_digests
        .iter()
        .any(|digest| digest != &expected_digest)
    {
        return Err("XLSX data-validation measured output digests are not stable".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(source_summary),
        execution: None,
        output_sha256: Some(expected_digest),
    })
}

fn run_opc_open(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let package = OpcPackage::from_bytes(&corpus.archive)?;
        let count = package.part_count();
        if count != corpus.manifest.entry_count {
            return Err("OPC open part count differs from generated corpus manifest".into());
        }
        std::hint::black_box(count);
        record_elapsed(
            &mut elapsed,
            iteration,
            warmup_iterations,
            started.elapsed(),
        )?;
    }
    Ok(result(Case::OpcOpen, corpus, elapsed, None))
}

fn run_opc_open_owned(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let owned = corpus.archive.clone();
        let started = Instant::now();
        let package = OpcPackage::from_vec(owned)?;
        let count = package.part_count();
        if count != corpus.manifest.entry_count {
            return Err("owned OPC open part count differs from generated corpus manifest".into());
        }
        std::hint::black_box(&package);
        record_elapsed(
            &mut elapsed,
            iteration,
            warmup_iterations,
            started.elapsed(),
        )?;
    }
    Ok(result(Case::OpcOpenOwned, corpus, elapsed, None))
}

fn run_opc_noop_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let package = OpcPackage::from_vec(corpus.archive.clone())?;
    let maximum = u64::try_from(corpus.archive.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("sequential output ceiling overflows u64")?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        PackageWriter::write_to_stream(&mut sink, &package)?;
        record_elapsed(
            &mut elapsed,
            iteration,
            warmup_iterations,
            started.elapsed(),
        )?;
        let summary = sink.summary();
        if summary.accepted_bytes == 0 || summary.write_calls == 0 {
            return Err("OPC sequential save wrote no bytes".into());
        }
        if sink.bytes != corpus.archive {
            return Err("OPC no-op save bytes differ from deterministic corpus".into());
        }
        if iteration >= warmup_iterations {
            sink_summaries.push(summary);
        }
    }

    let first = *sink_summaries
        .first()
        .ok_or("OPC sequential save produced no sink summary")?;
    if sink_summaries.iter().any(|summary| *summary != first) {
        return Err("deterministic no-op save produced differing sink summaries".into());
    }
    Ok(result(Case::OpcNoopSave, corpus, elapsed, Some(first)))
}

fn run_opc_mutated_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let target_uri = PackURI::new(format!("/{}", corpus.target_name))?;
    let mut changed_payload = corpus.target_payload.clone();
    let first = changed_payload
        .first_mut()
        .ok_or("OPC mutation target is empty")?;
    *first ^= 0xff;

    let mut package = OpcPackage::from_vec(corpus.archive.clone())?;
    package.get_part_mut(&target_uri)?.set_blob(changed_payload);
    let expected = PackageWriter::to_bytes(&package)?;
    let maximum = u64::try_from(expected.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("sequential output ceiling overflows u64")?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let started = Instant::now();
        PackageWriter::write_to_stream(&mut sink, &package)?;
        record_elapsed(
            &mut elapsed,
            iteration,
            warmup_iterations,
            started.elapsed(),
        )?;
        if sink.bytes != expected {
            return Err("mutated OPC save differs from deterministic expected output".into());
        }
        if iteration >= warmup_iterations {
            sink_summaries.push(sink.summary());
        }
    }

    let first = *sink_summaries
        .first()
        .ok_or("mutated OPC save produced no sink summary")?;
    if sink_summaries.iter().any(|summary| *summary != first) {
        return Err("deterministic mutated save produced differing sink summaries".into());
    }
    Ok(result(Case::OpcMutatedSave, corpus, elapsed, Some(first)))
}

fn run_cfb_open(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let expected_file_size = u64::try_from(corpus.archive.len())?;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let ole = OleFile::open(Cursor::new(corpus.archive.as_slice()))?;
        let duration = started.elapsed();
        if ole.file_size() != expected_file_size {
            return Err("CFB open file size differs from generated corpus manifest".into());
        }
        std::hint::black_box(&ole);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(Case::CfbOpen, corpus, elapsed, None))
}

fn run_cfb_list_streams(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let ole = OleFile::open(Cursor::new(corpus.archive.as_slice()))?;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let streams = ole.list_streams();
        let duration = started.elapsed();
        if streams.len() != corpus.manifest.entry_count {
            return Err("CFB list stream count differs from generated corpus manifest".into());
        }
        std::hint::black_box(&streams);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(Case::CfbListStreams, corpus, elapsed, None))
}

fn run_cfb_read_one(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let mut ole = OleFile::open(Cursor::new(corpus.archive.as_slice()))?;
    let path = [corpus.target_name.as_str()];
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let bytes = ole.open_stream(&path)?;
        let duration = started.elapsed();
        if bytes != corpus.target_payload {
            return Err("CFB read result differs from deterministic corpus payload".into());
        }
        std::hint::black_box(&bytes);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(Case::CfbReadOne, corpus, elapsed, None))
}

fn cfb_instrumented_source(corpus: &Corpus) -> Arc<InstrumentedSource> {
    Arc::new(InstrumentedSource::new(corpus.archive.clone(), Vec::new()))
}

fn cfb_shared_limits(corpus: &Corpus) -> Result<SharedOleFileLimits, Box<dyn Error>> {
    Ok(SharedOleFileLimits::new(u64::try_from(
        corpus.archive.len(),
    )?)?)
}

fn run_cfb_shared_open(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let limits = cfb_shared_limits(corpus)?;
    let expected_size = u64::try_from(corpus.archive.len())?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = cfb_instrumented_source(corpus);
        let started = Instant::now();
        let ole = SharedOleFile::open_with_limits(source.clone(), limits)?;
        let duration = started.elapsed();
        if ole.file_size() != expected_size {
            return Err("shared CFB open file size differs from corpus manifest".into());
        }
        std::hint::black_box(&ole);
        let metrics = source.snapshot();
        if metrics.read_calls == 0 || metrics.read_bytes == 0 {
            return Err("shared CFB open performed no positional source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record(metrics);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::CfbSharedOpen,
        corpus,
        elapsed,
        source_summary,
    ))
}

fn run_cfb_shared_read_one(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let limits = cfb_shared_limits(corpus)?;
    let path = [corpus.target_name.as_str()];
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = cfb_instrumented_source(corpus);
        let ole = SharedOleFile::open_with_limits(source.clone(), limits)?;
        source.reset();
        let started = Instant::now();
        let bytes = ole.open_stream(&path)?;
        let duration = started.elapsed();
        if bytes != corpus.target_payload {
            return Err("shared CFB read differs from deterministic stream payload".into());
        }
        std::hint::black_box(&bytes);
        let metrics = source.snapshot();
        if metrics.read_calls == 0 || metrics.read_bytes == 0 {
            return Err("shared CFB stream read performed no positional source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record(metrics);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::CfbSharedReadOne,
        corpus,
        elapsed,
        source_summary,
    ))
}

fn run_cfb_selective_read(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let shared = matches!(
        case,
        Case::CfbSelectiveMiniSharedRead | Case::CfbSelectiveFatSharedRead
    );
    let implementation = if shared {
        "shared-positional-exact-range"
    } else {
        "legacy-cursor-full-stream"
    };
    let mut elapsed = Vec::with_capacity(samples);
    let mut open_ns = Vec::with_capacity(samples);
    let mut read_ns = Vec::with_capacity(samples);
    let mut total_ns = Vec::with_capacity(samples);
    let mut open_read_calls = Vec::with_capacity(samples);
    let mut open_read_bytes = Vec::with_capacity(samples);
    let mut open_range_sizes = Vec::with_capacity(samples);
    let mut read_calls = Vec::with_capacity(samples);
    let mut read_bytes = Vec::with_capacity(samples);
    let mut read_range_sizes = Vec::with_capacity(samples);
    let mut returned_payload_bytes = Vec::with_capacity(samples);
    let mut selected_hash = None;
    let limits = cfb_shared_limits(corpus)?;
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let bytes = Arc::new(corpus.archive.clone());
        let metrics = Arc::new(SelectiveReadMetrics::default());
        let total_started = Instant::now();
        let open_started = Instant::now();
        let selected = if shared {
            let source = Arc::new(SelectiveReadAt::new(bytes, Arc::clone(&metrics)));
            let ole = SharedOleFile::open_with_limits(source, limits)?;
            let open_duration = open_started.elapsed();
            let open_snapshot = metrics.snapshot()?;
            metrics.reset()?;
            let read_started = Instant::now();
            let path = [corpus.target_name.as_str()];
            let mut payload = Vec::new();
            payload.try_reserve_exact(corpus.target_payload.len())?;
            payload.resize(corpus.target_payload.len(), 0);
            ole.read_stream_range(&path, 0, &mut payload)?;
            let read_duration = read_started.elapsed();
            (payload, open_duration, read_duration, open_snapshot)
        } else {
            let cursor = SelectiveCursor::new(bytes, Arc::clone(&metrics));
            let mut ole = OleFile::open(cursor)?;
            let open_duration = open_started.elapsed();
            let open_snapshot = metrics.snapshot()?;
            metrics.reset()?;
            let read_started = Instant::now();
            let path = [corpus.target_name.as_str()];
            let payload = ole.open_stream(&path)?;
            let read_duration = read_started.elapsed();
            (payload, open_duration, read_duration, open_snapshot)
        };
        let (payload, open_duration, read_duration, open_snapshot) = selected;
        let total_duration = total_started.elapsed();
        if payload != corpus.target_payload {
            return Err("CFB selective read differs from deterministic target payload".into());
        }
        let read_snapshot = metrics.snapshot()?;
        if open_snapshot.read_calls == 0
            || open_snapshot.read_bytes == 0
            || read_snapshot.read_calls == 0
            || read_snapshot.read_bytes == 0
        {
            return Err("CFB selective read performed no measured source I/O".into());
        }
        let hash = sha256_hex(&payload);
        if let Some(expected) = selected_hash.as_deref() {
            if expected != hash {
                return Err("CFB selective read hash changed across samples".into());
            }
        } else {
            selected_hash = Some(hash);
        }
        std::hint::black_box(&payload);
        if iteration >= warmup_iterations {
            let open_duration = elapsed_ns(open_duration)?;
            let read_duration = elapsed_ns(read_duration)?;
            let total_duration = elapsed_ns(total_duration)?;
            elapsed.push(total_duration);
            open_ns.push(open_duration);
            read_ns.push(read_duration);
            total_ns.push(total_duration);
            open_read_calls.push(open_snapshot.read_calls);
            open_read_bytes.push(open_snapshot.read_bytes);
            open_range_sizes.push(open_snapshot.range_sizes);
            read_calls.push(read_snapshot.read_calls);
            read_bytes.push(read_snapshot.read_bytes);
            read_range_sizes.push(read_snapshot.range_sizes);
            returned_payload_bytes.push(u64::try_from(payload.len())?);
        }
    }
    let selected_payload_sha256 = selected_hash.ok_or("CFB selective read produced no hash")?;
    let evidence = CfbSelectiveEvidence {
        timing_scope: "open and selected read are separate stages; legacy materializes the full stream, while shared fills a newly allocated exact-length caller range; corpus construction and validation excluded",
        sink: "none",
        selected_target_kind: if corpus.target_payload.len() < 4096 {
            "minifat-36-byte"
        } else {
            "fat-4mib"
        },
        legacy_or_positional: CfbSelectiveImplementationEvidence {
            implementation,
            open_ns,
            read_ns,
            total_ns,
            open_read_calls,
            open_read_bytes,
            open_range_sizes,
            read_calls,
            read_bytes,
            read_range_sizes,
            returned_payload_bytes,
            selected_payload_sha256,
        },
    };
    let source = SourceSummary {
        cfb_selective: Some(evidence),
        ..SourceSummary::default()
    };
    Ok(result_with_source(case, corpus, elapsed, source))
}

fn corpus_payload_kind(corpus: &Corpus) -> Result<PayloadKind, Box<dyn Error>> {
    match corpus.manifest.payload_kind {
        "compressible" => Ok(PayloadKind::Compressible),
        "incompressible" => Ok(PayloadKind::Incompressible),
        _ => Err("container corpus has an unknown payload kind".into()),
    }
}

fn run_cfb_shared_concurrent_reads(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let limits = cfb_shared_limits(corpus)?;
    let first_name = cfb_entry_name(0);
    let first_expected =
        payload_bytes(corpus_payload_kind(corpus)?, 0, corpus.manifest.entry_bytes);
    let mut elapsed = Vec::with_capacity(samples);
    let mut source_summary = SourceSummary::default();
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = cfb_instrumented_source(corpus);
        let ole = SharedOleFile::open_with_limits(source.clone(), limits)?;
        source.reset();
        let start = Arc::new(Barrier::new(3));
        let started = Instant::now();
        let (first, second) = std::thread::scope(|scope| {
            let first_start = Arc::clone(&start);
            let first_ole = &ole;
            let first_name = first_name.as_str();
            let first_task = scope.spawn(move || {
                first_start.wait();
                first_ole.open_stream(&[first_name])
            });
            let second_start = Arc::clone(&start);
            let second_ole = &ole;
            let second_name = corpus.target_name.as_str();
            let second_task = scope.spawn(move || {
                second_start.wait();
                second_ole.open_stream(&[second_name])
            });
            start.wait();
            (first_task.join(), second_task.join())
        });
        let first = first.map_err(|_panic| "first shared CFB worker panicked")??;
        let second = second.map_err(|_panic| "second shared CFB worker panicked")??;
        let duration = started.elapsed();
        if first != first_expected || second != corpus.target_payload {
            return Err("concurrent shared CFB reads returned unexpected stream bytes".into());
        }
        std::hint::black_box((&first, &second));
        let metrics = source.snapshot();
        if metrics.read_calls == 0 || metrics.read_bytes == 0 {
            return Err("concurrent shared CFB reads performed no positional source reads".into());
        }
        if iteration >= warmup_iterations {
            source_summary.record(metrics);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result_with_source(
        Case::CfbSharedConcurrentReads,
        corpus,
        elapsed,
        source_summary,
    ))
}

fn run_cfb_create_stream(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
    owned: bool,
) -> Result<CaseResult, Box<dyn Error>> {
    let case = if owned {
        Case::CfbCreateStreamOwned
    } else {
        Case::CfbCreateStreamBorrowed
    };
    let path = [corpus.target_name.as_str()];
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let prepared = std::hint::black_box(corpus.target_payload.clone());
        let mut writer = OleWriter::new();
        let started = Instant::now();
        if owned {
            writer.create_stream_owned(&path, prepared)?;
        } else {
            writer.create_stream(&path, &prepared)?;
        }
        let duration = started.elapsed();
        std::hint::black_box(&writer);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(case, corpus, elapsed, None))
}

fn run_ole_common_one_edit_save(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let corpus = build_ole_common_corpus(corpus)?;
    let expected = ole_common_changed_output(&corpus)?;
    let path = vec![corpus.target_name.clone()];
    let targets = OleObjectTargets::default();
    let limits = ole_common_limits(&corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut final_output = None;

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = corpus.archive.clone();
        let replacement = OLE_COMMON_REPLACEMENT.to_vec();
        let targets = targets.clone();
        let started = Instant::now();
        let mut editor = OleObjectEditor::open(source, targets, limits)?;
        editor.put_stream(&path, replacement)?;
        let output = editor.finish()?;
        let duration = started.elapsed();
        if output != expected {
            return Err("OLE common edit/save output is not deterministic".into());
        }
        std::hint::black_box(&output);
        final_output = Some(output);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    verify_ole_common_changed_output(
        &corpus,
        final_output
            .as_deref()
            .ok_or("OLE common edit/save produced no final output")?,
    )?;
    Ok(result(Case::OleCommonOneEditSave, &corpus, elapsed, None))
}

fn run_ole_common_open(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let corpus = build_ole_common_corpus(corpus)?;
    let targets = OleObjectTargets::default();
    let limits = ole_common_limits(&corpus)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut final_editor = None;

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = corpus.archive.clone();
        let targets = targets.clone();
        let started = Instant::now();
        let editor = OleObjectEditor::open(source, targets, limits)?;
        let duration = started.elapsed();
        if editor.is_changed() {
            return Err("fresh OLE common editor unexpectedly reports a change".into());
        }
        std::hint::black_box(&editor);
        final_editor = Some(editor);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    verify_ole_common_source_editor(
        &corpus,
        final_editor
            .as_ref()
            .ok_or("OLE common open produced no final editor")?,
    )?;
    Ok(result(Case::OleCommonOpen, &corpus, elapsed, None))
}

fn run_ole_common_put_stream_publish(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let corpus = build_ole_common_corpus(corpus)?;
    let expected = ole_common_changed_output(&corpus)?;
    let path = vec![corpus.target_name.clone()];
    let opened = OleObjectEditor::open(
        corpus.archive.clone(),
        OleObjectTargets::default(),
        ole_common_limits(&corpus)?,
    )?;
    let source = opened.snapshot();
    let mut elapsed = Vec::with_capacity(samples);
    let mut final_editor = None;

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let mut editor = source.edit();
        let replacement = OLE_COMMON_REPLACEMENT.to_vec();
        let started = Instant::now();
        editor.put_stream(&path, replacement)?;
        let duration = started.elapsed();
        if !editor.is_changed()
            || editor
                .stream(&path)
                .ok_or("OLE common target disappeared")?
                != OLE_COMMON_REPLACEMENT
        {
            return Err("OLE common candidate publication did not retain the replacement".into());
        }
        std::hint::black_box(&editor);
        final_editor = Some(editor);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let output = final_editor
        .ok_or("OLE common candidate publication produced no final editor")?
        .finish()?;
    if output != expected {
        return Err("OLE common candidate publication changed deterministic output".into());
    }
    verify_ole_common_changed_output(&corpus, &output)?;
    Ok(result(
        Case::OleCommonPutStreamPublish,
        &corpus,
        elapsed,
        None,
    ))
}

fn run_ole_common_finish_render(
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let corpus = build_ole_common_corpus(corpus)?;
    let expected = ole_common_changed_output(&corpus)?;
    let path = vec![corpus.target_name.clone()];
    let mut editor = OleObjectEditor::open(
        corpus.archive.clone(),
        OleObjectTargets::default(),
        ole_common_limits(&corpus)?,
    )?;
    editor.put_stream(&path, OLE_COMMON_REPLACEMENT.to_vec())?;
    let changed = editor.snapshot();
    let mut elapsed = Vec::with_capacity(samples);
    let mut final_output = None;

    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let editor = changed.edit();
        let started = Instant::now();
        let output = editor.finish()?;
        let duration = started.elapsed();
        std::hint::black_box(&output);
        final_output = Some(output);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }

    let output = final_output
        .as_deref()
        .ok_or("OLE common changed finish produced no final output")?;
    if output != expected {
        return Err("OLE common changed finish changed deterministic output".into());
    }
    verify_ole_common_changed_output(&corpus, output)?;
    Ok(result(Case::OleCommonFinishRender, &corpus, elapsed, None))
}

fn verify_ole_common_source_editor(
    corpus: &Corpus,
    editor: &OleObjectEditor,
) -> Result<(), Box<dyn Error>> {
    let kind = corpus_payload_kind(corpus)?;
    let unchanged_stream_count = corpus
        .manifest
        .entry_count
        .checked_sub(1)
        .ok_or("OLE common corpus has no edit target")?;
    for index in 0..unchanged_stream_count {
        let name = cfb_entry_name(index);
        let path = [name];
        if editor
            .stream(&path)
            .ok_or("OLE common source stream disappeared")?
            != payload_bytes(kind, index, corpus.manifest.entry_bytes)
        {
            return Err("OLE common open changed an opaque source stream".into());
        }
    }
    let target = [corpus.target_name.clone()];
    if editor
        .stream(&target)
        .ok_or("OLE common source target disappeared")?
        != OLE_COMMON_ORIGINAL
    {
        return Err("OLE common open changed its source target".into());
    }
    Ok(())
}

fn ole_common_changed_output(corpus: &Corpus) -> Result<Vec<u8>, Box<dyn Error>> {
    let path = vec![corpus.target_name.clone()];
    let mut editor = OleObjectEditor::open(
        corpus.archive.clone(),
        OleObjectTargets::default(),
        ole_common_limits(corpus)?,
    )?;
    editor.put_stream(&path, OLE_COMMON_REPLACEMENT.to_vec())?;
    let output = editor.finish()?;
    verify_ole_common_changed_output(corpus, &output)?;
    Ok(output)
}

fn ole_common_limits(corpus: &Corpus) -> Result<OleObjectLimits, Box<dyn Error>> {
    Ok(OleObjectLimits {
        max_objects: 1,
        max_storage_depth: 1,
        max_streams_per_object: 1,
        max_streams: corpus.manifest.entry_count,
        max_stream_size: u64::try_from(
            corpus
                .manifest
                .entry_bytes
                .max(OLE_COMMON_ORIGINAL.len())
                .max(OLE_COMMON_REPLACEMENT.len()),
        )?,
        max_object_size: 1,
        max_total_size: u64::try_from(corpus.manifest.uncompressed_payload_bytes)?,
    })
}

fn verify_ole_common_changed_output(corpus: &Corpus, output: &[u8]) -> Result<(), Box<dyn Error>> {
    let kind = corpus_payload_kind(corpus)?;
    let mut ole = OleFile::open(Cursor::new(output))?;
    let streams = ole.list_streams();
    if streams.len() != corpus.manifest.entry_count {
        return Err("OLE common output stream count differs from its corpus".into());
    }
    let unchanged_stream_count = corpus
        .manifest
        .entry_count
        .checked_sub(1)
        .ok_or("OLE common corpus has no edit target")?;
    for index in 0..unchanged_stream_count {
        let name = cfb_entry_name(index);
        let actual = ole.open_stream(&[name.as_str()])?;
        if actual != payload_bytes(kind, index, corpus.manifest.entry_bytes) {
            return Err("OLE common edit changed an untouched stream".into());
        }
    }
    if ole.open_stream(&[corpus.target_name.as_str()])? != OLE_COMMON_REPLACEMENT {
        return Err("OLE common changed stream differs from its replacement".into());
    }
    Ok(())
}

fn streaming_context(
    memory_bytes: u64,
    input_bytes: u64,
    output_bytes: u64,
    objects: u64,
    work: u64,
) -> Result<ExecutionContext, Box<dyn Error>> {
    let one = NonZeroUsize::new(1).ok_or("streaming worker count must be nonzero")?;
    let in_flight = NonZeroU64::new(memory_bytes.max(1))
        .ok_or("streaming in-flight byte limit must be nonzero")?;
    let execution_limits = ExecutionLimits::new(one, one, in_flight, 0)?;
    let (_cancellation, token) = CancellationSource::pair();
    Ok(ExecutionContext::new(
        Budget::root(
            "litchi-perf-streaming-create",
            Limits::new(memory_bytes, input_bytes, output_bytes, objects, 32, work),
        ),
        token,
        execution_limits,
    ))
}

fn streaming_xlsx_text(row: usize) -> String {
    format!("litchi-perf-streaming-xlsx-row-{row:06}-café-<&>")
}

fn streaming_rtf_text(paragraph: usize) -> String {
    format!("litchi-perf-streaming-rtf-paragraph-{paragraph:06}-café-\\{{}}")
}

fn write_streaming_xlsx<W: Write>(
    sink: W,
    shape: SemanticShape,
) -> Result<(W, StreamingMetrics), Box<dyn Error>> {
    let rows = u64::try_from(shape.streaming_units())?;
    let cells = rows
        .checked_mul(4)
        .ok_or("streaming XLSX cell count overflows")?;
    let max_sheet_xml_bytes = rows
        .checked_mul(512)
        .and_then(|value| value.checked_add(4 * 1024))
        .ok_or("streaming XLSX worksheet ceiling overflows")?;
    let max_output_bytes = max_sheet_xml_bytes
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("streaming XLSX output ceiling overflows")?;
    let objects = rows
        .checked_add(cells)
        .and_then(|value| value.checked_add(16))
        .ok_or("streaming XLSX object budget overflows")?;
    let work = objects
        .checked_mul(2)
        .ok_or("streaming XLSX work budget overflows")?;
    let limits = StreamingWorkbookLimits::new(
        u32::try_from(rows)?,
        cells,
        256,
        XLSX_STREAMING_ROW_BYTES,
        max_sheet_xml_bytes,
        max_output_bytes,
    );
    let context = streaming_context(XLSX_STREAMING_ROW_BYTES, 0, max_output_bytes, objects, work)?;
    let mut writer = StreamingWorkbookWriter::new(sink, context, limits)?;
    let mut input_bytes = 0u64;
    for row in 1..=shape.streaming_units() {
        let text = streaming_xlsx_text(row);
        let row_number = u32::try_from(row)?;
        input_bytes = input_bytes
            .checked_add(u64::try_from(text.len())?)
            .ok_or("streaming XLSX input byte count overflows")?;
        writer.write_row(
            row_number,
            [
                StreamingCell::new(1, StreamingCellValue::Number(f64::from(row_number))),
                StreamingCell::new(2, StreamingCellValue::Text(&text)),
                StreamingCell::new(3, StreamingCellValue::Bool(row % 2 == 0)),
                StreamingCell::new(4, StreamingCellValue::Blank),
            ],
        )?;
    }
    if writer.cell_count() != cells {
        return Err("streaming XLSX writer cell count differs from its shape".into());
    }
    let authored_part_bytes = writer
        .worksheet_xml_bytes()
        .checked_add(u64::try_from(b"</sheetData></worksheet>".len())?)
        .ok_or("streaming XLSX authored XML count overflows")?;
    let output = writer.finish()?;
    Ok((
        output,
        StreamingMetrics {
            rows,
            cells,
            paragraphs: 0,
            runs: 0,
            input_bytes,
            authored_part_bytes,
            retained_authoring_window_bytes: XLSX_STREAMING_ROW_BYTES,
        },
    ))
}

fn write_streaming_rtf<W: Write>(
    sink: W,
    shape: SemanticShape,
) -> Result<(W, StreamingMetrics), Box<dyn Error>> {
    let paragraphs = u64::try_from(shape.streaming_units())?;
    let runs = paragraphs;
    let input_bytes = (0..shape.streaming_units()).try_fold(
        0u64,
        |total, paragraph| -> Result<u64, Box<dyn Error>> {
            let text_bytes = u64::try_from(streaming_rtf_text(paragraph).len())?;
            Ok(total
                .checked_add(text_bytes)
                .ok_or("streaming RTF input byte count overflows")?)
        },
    )?;
    let max_output_bytes = input_bytes
        .checked_mul(8)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("streaming RTF output ceiling overflows")?;
    let objects = paragraphs
        .checked_add(runs)
        .and_then(|value| value.checked_add(8))
        .ok_or("streaming RTF object budget overflows")?;
    let work = input_bytes
        .checked_mul(2)
        .and_then(|value| value.checked_add(objects))
        .ok_or("streaming RTF work budget overflows")?;
    let context = streaming_context(
        RTF_STREAMING_SCRATCH_BYTES,
        input_bytes,
        max_output_bytes,
        objects,
        work,
    )?;
    let limits = litchi_rtf::write::StreamingRtfLimits::new(
        input_bytes,
        max_output_bytes,
        paragraphs,
        runs,
        RTF_STREAMING_SCRATCH_BYTES,
    );
    let mut writer = litchi_rtf::write::StreamingRtfWriter::new(sink, context, limits)?;
    for paragraph in 0..shape.streaming_units() {
        let text = streaming_rtf_text(paragraph);
        writer.start_paragraph()?;
        writer.start_run()?;
        writer.write_all(text.as_bytes())?;
        writer.finish_run()?;
        writer.finish_paragraph()?;
    }
    if writer.paragraph_count() != paragraphs
        || writer.run_count() != runs
        || writer.input_bytes() != input_bytes
    {
        return Err("streaming RTF writer counters differ from its shape".into());
    }
    let authored_part_bytes = writer
        .output_bytes()
        .checked_add(1)
        .ok_or("streaming RTF authored byte count overflows")?;
    let output = writer.finish()?;
    Ok((
        output,
        StreamingMetrics {
            rows: 0,
            cells: 0,
            paragraphs,
            runs,
            input_bytes,
            authored_part_bytes,
            retained_authoring_window_bytes: RTF_STREAMING_SCRATCH_BYTES,
        },
    ))
}

fn verify_streaming_xlsx(bytes: Vec<u8>, shape: SemanticShape) -> Result<(), Box<dyn Error>> {
    let workbook = Workbook::from_bytes(bytes)?;
    if workbook.len() != 1 {
        return Err("streaming XLSX workbook does not contain exactly one sheet".into());
    }
    let sheet = workbook
        .sheet("Sheet1")?
        .ok_or("streaming XLSX Sheet1 is missing")?;
    if sheet.rows()?.count() != shape.streaming_units() {
        return Err("streaming XLSX explicit row count differs from its shape".into());
    }
    let area = format!("A1:D{}", shape.streaming_units());
    let mut visited = 0usize;
    for (address, cell) in sheet.cells(area.as_str())? {
        let row = usize::try_from(address.row().get())? + 1;
        let column = address.column().get();
        let matches = match (column, cell) {
            (0, XlsxCell::Value(XlsxValue::Number(value))) => value.as_f64() == Some(row as f64),
            (1, XlsxCell::Value(XlsxValue::Text(value))) => {
                value.as_str() == streaming_xlsx_text(row)
            },
            (2, XlsxCell::Value(XlsxValue::Bool(value))) => *value == (row % 2 == 0),
            (3, XlsxCell::Empty) => true,
            _ => false,
        };
        if !matches {
            return Err(
                format!("streaming XLSX cell differs at row {row}, column {column}").into(),
            );
        }
        visited = visited
            .checked_add(1)
            .ok_or("streaming XLSX visited-cell count overflows")?;
    }
    if visited != shape.streaming_units().saturating_mul(4) {
        return Err("streaming XLSX stored-cell count differs from its shape".into());
    }
    Ok(())
}

fn verify_streaming_rtf(bytes: &[u8], shape: SemanticShape) -> Result<(), Box<dyn Error>> {
    let document = litchi_rtf::Document::from_bytes(bytes)?;
    if document.paragraph_count() != shape.streaming_units() {
        return Err("streaming RTF paragraph count differs from its shape".into());
    }
    for (index, paragraph) in document.body().paragraphs().enumerate() {
        let expected = streaming_rtf_text(index);
        let mut runs = paragraph.runs();
        if runs.next().map(|run| run.text()) != Some(expected.as_str()) || runs.next().is_some() {
            return Err(format!("streaming RTF paragraph {index} differs from its shape").into());
        }
    }
    Ok(())
}

fn build_streaming_corpus(
    case: Case,
    shape: SemanticShape,
) -> Result<StreamingCorpus, Box<dyn Error>> {
    let (artifact, mut metrics, generator, package_format, compression, target_entry, members) =
        match case {
            Case::XlsxStreamingCreate => {
                let (artifact, metrics) = write_streaming_xlsx(Vec::new(), shape)?;
                verify_streaming_xlsx(artifact.clone(), shape)?;
                (
                    artifact,
                    metrics,
                    XLSX_STREAMING_CORPUS_GENERATOR,
                    "xlsx",
                    "deflate",
                    "xl/worksheets/sheet1.xml".to_string(),
                    6,
                )
            },
            Case::RtfStreamingCreate => {
                let (artifact, metrics) = write_streaming_rtf(Vec::new(), shape)?;
                verify_streaming_rtf(&artifact, shape)?;
                (
                    artifact,
                    metrics,
                    RTF_STREAMING_CORPUS_GENERATOR,
                    "rtf",
                    "none",
                    "document".to_string(),
                    0,
                )
            },
            _ => return Err("non-streaming case requested a streaming corpus".into()),
        };
    let target = if case == Case::XlsxStreamingCreate {
        ArchiveReader::new(&artifact)?.read(&target_entry)?
    } else {
        artifact.clone()
    };
    metrics.authored_part_bytes = u64::try_from(target.len())?;
    let archive_sha256 = sha256_hex(&artifact);
    let target_payload_sha256 = sha256_hex(&target);
    let manifest = CorpusManifest {
        name: format!("{package_format}-streaming-create-{}", shape.name()),
        generator,
        package_format,
        shape: shape.name(),
        payload_kind: "deterministic-scalar-text",
        compression,
        entry_count: usize::try_from(metrics.cells.max(metrics.runs))?,
        archive_member_count: members,
        entry_bytes: 0,
        uncompressed_payload_bytes: usize::try_from(metrics.input_bytes)?,
        archive_bytes: artifact.len(),
        archive_sha256,
        target_entry,
        target_payload_bytes: target.len(),
        target_payload_sha256,
        rtf_variant: (case == Case::RtfStreamingCreate).then_some("plain-streaming"),
        xlsx: (case == Case::XlsxStreamingCreate).then_some(XlsxManifest {
            sheet_count: 1,
            rows_per_sheet: shape.streaming_units(),
            columns_per_sheet: 4,
            one_percent_update_count: 0,
            source_members: XlsxSourceMembersManifest {
                workbook: "xl/workbook.xml".to_string(),
                worksheets: vec!["xl/worksheets/sheet1.xml".to_string()],
                shared_strings: None,
                styles: Some("xl/styles.xml".to_string()),
            },
        }),
    };
    drop(target);
    drop(artifact);
    Ok(StreamingCorpus {
        manifest,
        shape,
        metrics,
    })
}

fn run_streaming_creation(
    case: Case,
    corpus: &StreamingCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let maximum = u64::try_from(corpus.manifest.archive_bytes)?
        .checked_add(64 * 1024)
        .ok_or("streaming sink ceiling overflows")?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut summaries = Vec::with_capacity(samples);
    let mut digests = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let sink = HashingDiscardSink::new(maximum, corpus.metrics.retained_authoring_window_bytes);
        let started = Instant::now();
        let (sink, metrics) = match case {
            Case::XlsxStreamingCreate => write_streaming_xlsx(sink, corpus.shape)?,
            Case::RtfStreamingCreate => write_streaming_rtf(sink, corpus.shape)?,
            _ => return Err("non-streaming case reached streaming runner".into()),
        };
        let duration = started.elapsed();
        let (mut summary, digest) = sink.finish();
        if metrics != corpus.metrics
            || summary.accepted_bytes != u64::try_from(corpus.manifest.archive_bytes)?
            || digest != corpus.manifest.archive_sha256
        {
            return Err(
                "streaming creation counters or digest differ from untimed artifact".into(),
            );
        }
        summary.rows = (metrics.rows != 0).then_some(metrics.rows);
        summary.cells = (metrics.cells != 0).then_some(metrics.cells);
        summary.paragraphs = (metrics.paragraphs != 0).then_some(metrics.paragraphs);
        summary.runs = (metrics.runs != 0).then_some(metrics.runs);
        summary.input_bytes = Some(metrics.input_bytes);
        summary.authored_part_bytes = Some(metrics.authored_part_bytes);
        if iteration >= warmup_iterations {
            summaries.push(summary);
            digests.push(digest);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let sink = deterministic_sink_summary(&summaries, "streaming creation")?;
    if sink.retained_output_bytes != Some(0)
        || sink.retained_authoring_window_bytes
            != Some(corpus.metrics.retained_authoring_window_bytes)
    {
        return Err("streaming creation did not prove its fixed retained-window bound".into());
    }
    if digests
        .iter()
        .any(|digest| digest != &corpus.manifest.archive_sha256)
    {
        return Err("streaming creation output digest changed across samples".into());
    }
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: None,
        execution: None,
        output_sha256: Some(corpus.manifest.archive_sha256.clone()),
    })
}

fn run_fresh_writer(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let shape = match corpus.manifest.shape {
        "tiny" => WriterShape::Tiny,
        "large" => WriterShape::Large,
        "payload-heavy" => WriterShape::PayloadHeavy,
        _ => return Err("fresh writer corpus has an unknown writer shape".into()),
    };
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let started = Instant::now();
        let (output, entry_count, content_bytes) = match case {
            Case::DocFreshWriteTo => write_fresh_doc(shape)?,
            Case::XlsFreshWriteTo => write_fresh_xls(shape)?,
            Case::PptFreshWriteTo => write_fresh_ppt(shape)?,
            _ => return Err("non-writer case passed to fresh writer runner".into()),
        };
        let duration = started.elapsed();
        if entry_count != corpus.manifest.entry_count
            || content_bytes != corpus.manifest.uncompressed_payload_bytes
        {
            return Err(
                "fresh writer content differs from deterministic corpus specification".into(),
            );
        }
        if output != corpus.archive {
            return Err("fresh writer package differs from deterministic corpus output".into());
        }
        std::hint::black_box(&output);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(result(case, corpus, elapsed, None))
}

fn result(case: Case, corpus: &Corpus, elapsed: Vec<u64>, sink: Option<SinkSummary>) -> CaseResult {
    CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink,
        source: None,
        execution: None,
        output_sha256: None,
    }
}

fn result_with_source(
    case: Case,
    corpus: &Corpus,
    elapsed: Vec<u64>,
    source: SourceSummary,
) -> CaseResult {
    CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: None,
        source: Some(source),
        execution: None,
        output_sha256: None,
    }
}

fn result_with_execution(
    case: Case,
    corpus: &Corpus,
    elapsed: Vec<u64>,
    execution: ExecutionSummary,
) -> CaseResult {
    CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: None,
        source: None,
        execution: Some(execution),
        output_sha256: None,
    }
}

fn elapsed_ns(duration: Duration) -> Result<u64, Box<dyn Error>> {
    u64::try_from(duration.as_nanos())
        .map_err(|_error| "duration does not fit u64 nanoseconds".into())
}

fn iteration_count(warmup_iterations: usize, samples: usize) -> Result<usize, Box<dyn Error>> {
    warmup_iterations
        .checked_add(samples)
        .ok_or_else(|| "warm-up and sample iteration count overflows usize".into())
}

fn record_elapsed(
    elapsed: &mut Vec<u64>,
    iteration: usize,
    warmup_iterations: usize,
    duration: Duration,
) -> Result<(), Box<dyn Error>> {
    if iteration >= warmup_iterations {
        elapsed.push(elapsed_ns(duration)?);
    }
    Ok(())
}

fn statistics(mut samples: Vec<u64>) -> Statistics {
    samples.sort_unstable();
    let count = samples.len();
    let (mean, squared_deviation_sum) = samples.iter().enumerate().fold(
        (0.0, 0.0),
        |(mean, squared_deviation_sum), (index, value)| {
            let value = *value as f64;
            let next_count = (index + 1) as f64;
            let delta = value - mean;
            let next_mean = mean + delta / next_count;
            let next_sum = squared_deviation_sum + delta * (value - next_mean);
            (next_mean, next_sum)
        },
    );
    let standard_deviation = if count > 1 {
        (squared_deviation_sum / (count - 1) as f64).sqrt()
    } else {
        0.0
    };
    let margin = if count > 1 {
        student_t_critical_95(count - 1) * standard_deviation / (count as f64).sqrt()
    } else {
        0.0
    };

    Statistics {
        unit: "ns",
        min: samples[0],
        p50: midpoint(samples[(count - 1) / 2], samples[count / 2]),
        p95: nearest_rank(&samples, 95),
        p99: nearest_rank(&samples, 99),
        max: samples[count - 1],
        mean,
        standard_deviation,
        confidence_interval_95: ConfidenceInterval {
            method: "two-sided Student's t interval for the mean",
            lower: (mean - margin).max(0.0),
            upper: mean + margin,
        },
        samples,
    }
}

fn midpoint(left: u64, right: u64) -> u64 {
    left / 2 + right / 2 + (left % 2 + right % 2) / 2
}

fn nearest_rank(samples: &[u64], percentile: usize) -> u64 {
    let index = percentile
        .saturating_mul(samples.len())
        .saturating_add(99)
        .saturating_div(100)
        .saturating_sub(1)
        .min(samples.len() - 1);
    samples[index]
}

fn student_t_critical_95(degrees_of_freedom: usize) -> f64 {
    const VALUES: [f64; 30] = [
        12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228, 2.201, 2.179, 2.160,
        2.145, 2.131, 2.120, 2.110, 2.101, 2.093, 2.086, 2.080, 2.074, 2.069, 2.064, 2.060, 2.056,
        2.052, 2.048, 2.045, 2.042,
    ];
    match degrees_of_freedom {
        0 => 0.0,
        1..=30 => VALUES[degrees_of_freedom - 1],
        _ => {
            // Cornish-Fisher expansion of the two-sided 0.975 quantile.
            // The exact table above covers the small sample counts where the
            // correction is largest; this converges rapidly thereafter.
            const Z: f64 = 1.959_963_984_540_054;
            let degrees = degrees_of_freedom as f64;
            let z2 = Z * Z;
            let z3 = z2 * Z;
            let z5 = z3 * z2;
            let z7 = z5 * z2;
            Z + (z3 + Z) / (4.0 * degrees)
                + (5.0 * z5 + 16.0 * z3 + 3.0 * Z) / (96.0 * degrees * degrees)
                + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * Z)
                    / (384.0 * degrees * degrees * degrees)
        },
    }
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn environment(host: filesystem::HostEvidence) -> Environment {
    Environment {
        rustc_version: command_output(
            std::env::var_os("RUSTC").as_deref().unwrap_or_default(),
            &["--version"],
        )
        .or_else(|| command_output(std::ffi::OsStr::new("rustc"), &["--version"])),
        git_revision: git_output(&["rev-parse", "HEAD"]),
        git_worktree_dirty: git_output(&["status", "--porcelain"]).map(|status| !status.is_empty()),
        logical_cpus_available: std::thread::available_parallelism()
            .map_or(1, std::num::NonZeroUsize::get),
        allocator: "Rust system allocator",
        rustflags: std::env::var("RUSTFLAGS")
            .ok()
            .or_else(|| option_env!("RUSTFLAGS").map(str::to_owned)),
        cargo_build_target: std::env::var("CARGO_BUILD_TARGET")
            .ok()
            .or_else(|| option_env!("CARGO_BUILD_TARGET").map(str::to_owned)),
        perf_event_paranoid: fs::read_to_string("/proc/sys/kernel/perf_event_paranoid")
            .ok()
            .map(|value| value.trim().to_owned()),
        os: host.os,
        kernel: host.kernel,
        cpu_model: host.cpu_model,
        total_memory_bytes: host.total_memory_bytes,
        page_size_bytes: host.page_size_bytes,
        filesystem_type: host.filesystem_type,
        source_destination_same_device: host.source_destination_same_device,
        cpu_affinity: host.cpu_affinity,
        storage_identifier: host.storage_identifier,
    }
}

fn git_output(arguments: &[&str]) -> Option<String> {
    let manifest_directory = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let repository = manifest_directory.parent()?.parent()?;
    let output = Command::new("git")
        .arg("-C")
        .arg(repository)
        .args(arguments)
        .output()
        .ok()?;
    if !output.status.success() {
        return None;
    }
    String::from_utf8(output.stdout)
        .ok()
        .map(|value| value.trim().to_owned())
}

fn command_output(program: &std::ffi::OsStr, arguments: &[&str]) -> Option<String> {
    if program.is_empty() {
        return None;
    }
    let output = Command::new(program).args(arguments).output().ok()?;
    if !output.status.success() {
        return None;
    }
    String::from_utf8(output.stdout)
        .ok()
        .map(|value| value.trim().to_owned())
}

fn write_report(report: &Report, output: Option<&PathBuf>) -> Result<(), Box<dyn Error>> {
    match output {
        Some(path) => {
            if let Some(parent) = path
                .parent()
                .filter(|parent| !parent.as_os_str().is_empty())
            {
                fs::create_dir_all(parent)?;
            }
            let file = File::create(path)?;
            serde_json::to_writer_pretty(file, report)?;
        },
        None => {
            let stdout = io::stdout();
            let mut writer = stdout.lock();
            serde_json::to_writer_pretty(&mut writer, report)?;
            writer.write_all(b"\n")?;
        },
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::{io::Write, sync::Arc, time::Duration};

    use litchi_core::ReadAt;

    use super::{
        Case, CfbSelectiveTarget, CorpusShape, CountingSink, InstrumentedSource,
        ODF_REPAIR_LOCAL_EXTRA, ODF_REPAIR_PUBLICATION_SCRATCH_BYTES, ODP_TEXT_BOX_BATCH_COUNT,
        ODT_RESOURCE_BATCH_COUNT, OpcCacheMode, PPTX_MULTI_SLIDE_BATCH_COUNT, PayloadKind,
        RTF_LOGICAL_TAIL_SINK_WINDOW_BYTES, RangeSimulationConfig, RequestSizeBuckets,
        RtfSemanticVariant, SemanticShape, SimulatedRangeSource, SourceBackedPackage, WriterShape,
        XLSX_CELL_VALUES_MEDIA_ENTRY_COUNT, XLSX_CELL_VALUES_SOURCE_EDIT_CORPUS_GENERATOR,
        XlsxCellCrudShape, XlsxShape, build_cfb_corpus, build_cfb_selective_corpus,
        build_docx_source_edit_corpus, build_odf_repair_corpus, build_odp_media_corpus,
        build_odp_text_box_batch_corpus, build_ods_media_corpus, build_odt_media_corpus,
        build_odt_resource_batch_corpus, build_ole_common_corpus, build_opc_corpus,
        build_pptx_source_edit_corpus, build_rtf_lifecycle_corpus, build_semantic_docx_corpus,
        build_semantic_odp_corpus, build_semantic_ods_corpus, build_semantic_odt_corpus,
        build_semantic_pptx_corpus, build_semantic_rtf_corpus, build_streaming_corpus,
        build_writer_corpus, build_xls_comments_edit_corpus, build_xls_visibility_edit_corpus,
        build_xlsx_auto_filter_edit_corpus, build_xlsx_calculation_metadata_edit_corpus,
        build_xlsx_cell_crud_corpus, build_xlsx_conditional_formatting_edit_corpus,
        build_xlsx_corpus, build_xlsx_data_validation_edit_corpus,
        build_xlsx_defined_names_edit_corpus, build_xlsx_merge_edit_corpus,
        build_xlsx_page_break_edit_corpus, build_xlsx_page_margin_edit_corpus,
        build_xlsx_page_setup_edit_corpus, build_xlsx_print_options_edit_corpus,
        build_xlsx_sheet_protection_edit_corpus, expected_opc_overlay_output,
        ole_common_changed_output, opc_overlay_replacement_payload, payload_bytes,
        resolve_execution_workers, run_case, run_case_with_config, run_cfb_selective_read,
        run_docx_source_backed_one_edit_save, run_opc_source_cache_budget_boundary,
        run_opc_source_cache_contention, run_opc_source_overlay_one_part_save,
        run_pptx_batch_edit_save, run_pptx_multi_slide_batch_edit_save,
        run_pptx_source_backed_one_edit_save, run_scaling_case, run_streaming_creation,
        run_xls_comments_edit_save, run_xls_visibility_edit_save, run_xlsx_auto_filter_edit_save,
        run_xlsx_calculation_metadata_edit_save, run_xlsx_conditional_formatting_edit_save,
        run_xlsx_data_validation_edit_save, run_xlsx_defined_names_edit_save,
        run_xlsx_page_break_edit_save, run_xlsx_page_margin_edit_save,
        run_xlsx_page_setup_edit_save, run_xlsx_print_options_edit_save,
        run_xlsx_sheet_protection_edit_save, sha256_hex, simulated_request_delay, statistics,
        xlsx_cell_count,
    };

    #[test]
    fn corpus_generation_is_deterministic() {
        let first = build_opc_corpus(CorpusShape::Tiny, PayloadKind::Compressible).unwrap();
        let second = build_opc_corpus(CorpusShape::Tiny, PayloadKind::Compressible).unwrap();

        assert_eq!(first.archive, second.archive);
        assert_eq!(
            first.manifest.archive_sha256,
            "1e28b8a9049a82f07e8ea88b2d492ef522d2da793d22fa50e2fe7f354dca3e2a"
        );
        let source_package = SourceBackedPackage::from_read_at(Arc::new(InstrumentedSource::new(
            first.archive.clone(),
            Vec::new(),
        )))
        .unwrap();
        assert_eq!(
            source_package
                .main_document_part()
                .unwrap()
                .partname()
                .membername(),
            first.target_name
        );
    }

    #[test]
    fn selective_cfb_corpora_are_bounded_and_deterministic() {
        for shape in [CorpusShape::ManySmall, CorpusShape::WideRoot] {
            for target in [CfbSelectiveTarget::Mini, CfbSelectiveTarget::Fat] {
                let first = build_cfb_selective_corpus(shape, target).unwrap();
                let second = build_cfb_selective_corpus(shape, target).unwrap();
                assert_eq!(first.archive, second.archive);
                assert_eq!(
                    first.manifest.archive_sha256,
                    second.manifest.archive_sha256
                );
                assert_eq!(
                    first.manifest.target_payload_sha256,
                    sha256_hex(&first.target_payload)
                );
                assert_eq!(first.manifest.entry_count, shape.entry_count());
                assert_eq!(first.manifest.target_payload_bytes, target.target_bytes());
                assert_eq!(
                    first.manifest.uncompressed_payload_bytes,
                    (shape.entry_count() - 1) * 1024 + target.target_bytes()
                );
            }
        }
    }

    #[test]
    fn selective_cfb_read_records_exact_payload_and_io_stages() {
        let corpus =
            build_cfb_selective_corpus(CorpusShape::ManySmall, CfbSelectiveTarget::Mini).unwrap();
        for case in [
            Case::CfbSelectiveMiniLegacyRead,
            Case::CfbSelectiveMiniSharedRead,
        ] {
            let result = run_cfb_selective_read(case, &corpus, 0, 1).unwrap();
            let evidence = result
                .source
                .unwrap()
                .cfb_selective
                .unwrap()
                .legacy_or_positional;
            assert_eq!(evidence.returned_payload_bytes, vec![36]);
            assert_eq!(
                evidence.selected_payload_sha256,
                corpus.manifest.target_payload_sha256
            );
            assert_eq!(evidence.open_ns.len(), 1);
            assert_eq!(evidence.read_ns.len(), 1);
            assert_eq!(evidence.total_ns.len(), 1);
            assert!(evidence.open_read_calls[0] > 0);
            assert!(!evidence.open_range_sizes[0].is_empty());
            assert!(evidence.read_calls[0] > 0);
            assert!(!evidence.read_range_sizes[0].is_empty());
        }
    }

    #[test]
    fn default_matrix_case_and_result_counts_are_stable() {
        assert_eq!(Case::DEFAULT.len(), 36);
        assert_eq!(
            Case::DEFAULT
                .iter()
                .filter(|case| case.uses_synthetic_opc())
                .count(),
            10
        );
        assert_eq!(
            Case::DEFAULT
                .iter()
                .filter(|case| case.uses_synthetic_cfb())
                .count(),
            8
        );
        let substrate_results = 18 * CorpusShape::ALL.len() * PayloadKind::ALL.len();
        let writer_results = 3 * WriterShape::ALL.len();
        let xlsx_results = 15 * XlsxShape::ALL.len();
        assert_eq!(substrate_results + writer_results + xlsx_results, 198);
        assert!(!Case::DEFAULT.contains(&Case::OpcSourceOverlayOnePartSave));
        assert!(!Case::DEFAULT.contains(&Case::DocxSourceBackedOneEditSave));
        assert!(!Case::DEFAULT.contains(&Case::PptxSourceBackedOneEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxEagerCalculationMetadataEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxSourceBackedCalculationMetadataEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxEagerDefinedNamesEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxSourceBackedDefinedNamesEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxEagerPageBreakEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxSourceBackedPageBreakEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxEagerPageMarginEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxSourceBackedPageMarginEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxEagerPageSetupEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxSourceBackedPageSetupEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxEagerPrintOptionsEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxSourceBackedPrintOptionsEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxEagerSheetProtectionEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxSourceBackedSheetProtectionEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxEagerDataValidationEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxSourceBackedDataValidationEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxEagerAutoFilterEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxSourceBackedAutoFilterEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxEagerConditionalFormattingEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxSourceBackedConditionalFormattingEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsCommentsEagerEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsCommentsSourceBackedEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsCommentsEagerBatchEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsCommentsSourceBackedBatchEditSave));
        assert!(!Case::DEFAULT.contains(&Case::XlsxStreamingCreate));
        assert!(!Case::DEFAULT.contains(&Case::RtfStreamingCreate));
        assert!(!Case::DEFAULT.contains(&Case::RtfValidationReport));
        assert!(!Case::DEFAULT.contains(&Case::XlsValidationReport));
        assert!(!Case::DEFAULT.contains(&Case::DocxValidationReport));
        assert!(!Case::DEFAULT.contains(&Case::DocxSectionInventory));
        assert!(!Case::DEFAULT.contains(&Case::PptxValidationReport));
        assert!(!Case::DEFAULT.contains(&Case::OdfValidationReport));
        assert!(!Case::DEFAULT.contains(&Case::OdfMimetypeRepairPlan));
    }

    #[test]
    fn odf_repair_case_is_deterministic_reversible_and_forward_only() {
        let first = build_odf_repair_corpus(SemanticShape::Tiny).unwrap();
        let second = build_odf_repair_corpus(SemanticShape::Tiny).unwrap();
        assert_eq!(first.archive, second.archive);
        assert_eq!(
            first.manifest.archive_sha256,
            second.manifest.archive_sha256
        );
        assert_eq!(first.target_payload, ODF_REPAIR_LOCAL_EXTRA);

        let measured = run_case(Case::OdfMimetypeRepairPlan, &first, 1, 2).unwrap();
        assert_eq!(measured.elapsed_ns.samples.len(), 2);
        let sink = measured.sink.expect("ODF repair sink summary");
        assert_eq!(sink.retained_output_bytes, Some(0));
        assert_eq!(sink.retained_authoring_window_bytes, None);
        assert!(sink.write_calls > 1);
        assert!(sink.largest_write <= ODF_REPAIR_PUBLICATION_SCRATCH_BYTES);
        let repair = measured
            .source
            .expect("ODF repair source summary")
            .odf_repair
            .expect("ODF repair evidence");
        assert_eq!(
            repair.repair_id,
            litchi_odf_common::MIMETYPE_LOCAL_EXTRA_REPAIR
        );
        assert_eq!(repair.extra_field_id, 0x5455);
        assert_eq!(repair.extra_field_bytes, 9);
        assert!(repair.member_payloads_preserved);
        assert!(repair.reversible);
        assert!(repair.exact_canonical_recovery_verified);
        assert!(repair.patch_verified);
        assert!(repair.inverse_verified);
        assert!(repair.stale_source_refusal_verified);
        assert!(repair.canonical_no_plan_verified);
        assert!(repair.partial_sink_progress_verified);
        assert_eq!(
            measured.output_sha256.as_deref(),
            Some(repair.output_sha256.as_str())
        );
    }

    #[test]
    fn bounded_validation_cases_exclude_warmups_and_preserve_topology() {
        let cases = [
            (
                Case::RtfValidationReport,
                build_semantic_rtf_corpus(SemanticShape::Tiny, RtfSemanticVariant::Plain).unwrap(),
            ),
            (
                Case::XlsValidationReport,
                build_writer_corpus(Case::XlsFreshWriteTo, WriterShape::Tiny).unwrap(),
            ),
            (
                Case::DocxValidationReport,
                build_semantic_docx_corpus(SemanticShape::Tiny).unwrap(),
            ),
            (
                Case::DocxSectionInventory,
                build_semantic_docx_corpus(SemanticShape::Tiny).unwrap(),
            ),
            (
                Case::PptxValidationReport,
                build_semantic_pptx_corpus(SemanticShape::Tiny).unwrap(),
            ),
            (
                Case::OdfValidationReport,
                build_semantic_odt_corpus(SemanticShape::Tiny).unwrap(),
            ),
        ];

        for (case, corpus) in cases {
            let measured = run_case(case, &corpus, 1, 2).unwrap();
            assert_eq!(measured.elapsed_ns.samples.len(), 2, "{case:?}");
            let source = measured.source.as_ref().expect("validation source summary");
            let validation = source
                .validation
                .as_ref()
                .expect("validation topology summary");
            assert!(!validation.check_ids.is_empty(), "{case:?}");
            assert_eq!(
                validation.source_sha256_before, validation.source_sha256_after,
                "{case:?} mutated its source"
            );
            assert_eq!(
                validation.source_bytes,
                u64::try_from(corpus.archive.len()).unwrap(),
                "{case:?} source byte count"
            );

            if matches!(case, Case::RtfValidationReport | Case::OdfValidationReport) {
                assert!(source.read_calls.is_empty(), "{case:?} is borrowed-input");
                assert!(validation.source_read_calls.is_none(), "{case:?}");
            } else {
                assert_eq!(source.read_calls.len(), 2, "{case:?} warmup leaked");
                assert_eq!(source.read_bytes.len(), 2, "{case:?} warmup leaked");
                assert_eq!(validation.source_read_calls, Some(source.read_calls[0]));
                assert_eq!(validation.source_read_bytes, Some(source.read_bytes[0]));
                assert!(source.read_bytes[0] > 0, "{case:?}");
            }

            if case == Case::DocxSectionInventory {
                let inventory = validation
                    .section_inventory
                    .as_ref()
                    .expect("DOCX section inventory");
                assert_eq!(inventory.section_count, 1);
                assert_eq!(inventory.paragraph_count, 24);
                let descriptor = inventory
                    .descriptors
                    .first()
                    .expect("DOCX body-final section");
                assert_eq!(descriptor.position, 0);
                assert_eq!(descriptor.ownership, "body_final");
                assert_eq!(descriptor.paragraph_start, 0);
                assert_eq!(descriptor.paragraph_end, 24);
            } else {
                assert!(validation.section_inventory.is_none(), "{case:?}");
            }
        }
    }

    #[test]
    fn xls_comment_controls_are_deterministic_bounded_and_source_evidenced() {
        let corpus = build_xls_comments_edit_corpus().unwrap();
        let again = build_xls_comments_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(
            corpus.manifest.generator,
            "litchi-xls-comments-opaque-heavy-v1"
        );
        assert_eq!(corpus.manifest.archive_member_count, 10);

        for (case, updates, source_backed) in [
            (Case::XlsCommentsEagerEditSave, 1, false),
            (Case::XlsCommentsSourceBackedEditSave, 1, true),
            (Case::XlsCommentsEagerBatchEditSave, 256, false),
            (Case::XlsCommentsSourceBackedBatchEditSave, 256, true),
        ] {
            let measured = run_xls_comments_edit_save(case, &corpus, 0, 1).unwrap();
            assert_eq!(measured.case, case.name());
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
            assert!(measured.output_sha256.is_some());
            let sink = measured.sink.unwrap();
            assert!(sink.accepted_bytes > 0);
            if source_backed {
                assert_eq!(sink.accepted_bytes, measured.corpus.archive_bytes as u64);
            }
            assert!(sink.write_calls > 0);
            assert!(sink.largest_write <= 64 * 1024);

            let source = measured.source.unwrap();
            assert_eq!(source.read_calls.len(), 1);
            assert_eq!(source.read_bytes, vec![corpus.archive.len() as u64]);
            let comments = source.xls_comments.unwrap();
            assert_eq!(comments.source_counter_scope, "owned-source-ingress-only");
            assert_eq!(comments.source_backed, source_backed);
            assert_eq!(comments.update_count, updates);
            assert_eq!(comments.changed_comments, vec![updates]);
            assert_eq!(comments.touched_streams, vec![1]);
            assert_eq!(comments.semantic_staging_plan_ns.len(), 1);
            assert_eq!(comments.publication_ns.len(), 1);
            if source_backed {
                assert_eq!(
                    comments.source_workbook_bytes,
                    comments.target_workbook_bytes
                );
                assert_eq!(comments.splice_count, Some(vec![updates * 2]));
                let replacement_bytes = comments
                    .replacement_bytes
                    .as_ref()
                    .expect("source-backed XLS comment replacement-byte evidence");
                assert_eq!(replacement_bytes.len(), 1);
                assert!(replacement_bytes[0] > 0);
                assert!(replacement_bytes[0] < comments.source_workbook_bytes[0]);
                assert!(comments.changed_spans.unwrap()[0] > 0);
                assert_eq!(comments.source_fingerprints.unwrap().len(), 1);
                assert_eq!(comments.target_fingerprints.unwrap().len(), 1);
            } else {
                assert!(comments.splice_count.is_none());
                assert!(comments.replacement_bytes.is_none());
                assert!(comments.changed_spans.is_none());
                assert!(comments.source_fingerprints.is_none());
                assert!(comments.target_fingerprints.is_none());
            }
        }
    }

    #[test]
    fn xls_visibility_controls_are_deterministic_bounded_and_source_evidenced() {
        let corpus = build_xls_visibility_edit_corpus().unwrap();
        let again = build_xls_visibility_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.generator, "litchi-xls-visibility-opaque-v1");
        assert_eq!(corpus.manifest.entry_count, 66);
        assert_eq!(corpus.manifest.archive_member_count, 10);

        for (case, updates, source_backed) in [
            (Case::XlsVisibilityEagerEditSave, 1, false),
            (Case::XlsVisibilitySourceBackedEditSave, 1, true),
            (
                Case::XlsVisibilityEagerBatchEditSave,
                litchi_xls::sheet_visibility::MAX_VISIBILITY_CHANGES,
                false,
            ),
            (
                Case::XlsVisibilitySourceBackedBatchEditSave,
                litchi_xls::sheet_visibility::MAX_VISIBILITY_CHANGES,
                true,
            ),
        ] {
            let measured = run_xls_visibility_edit_save(case, &corpus, 0, 1).unwrap();
            assert_eq!(measured.case, case.name());
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
            assert!(measured.output_sha256.is_some());
            let sink = measured.sink.unwrap();
            assert_eq!(sink.accepted_bytes, corpus.manifest.archive_bytes as u64);
            assert!(sink.write_calls > 0);
            assert!(sink.largest_write <= 64 * 1024);

            let source = measured.source.unwrap();
            assert_eq!(source.read_calls.len(), 1);
            assert_eq!(source.read_bytes, vec![corpus.archive.len() as u64]);
            let visibility = source.xls_visibility.unwrap();
            assert_eq!(visibility.source_counter_scope, "owned-source-ingress-only");
            assert_eq!(visibility.source_backed, source_backed);
            assert_eq!(visibility.update_count, updates);
            assert_eq!(visibility.changed_worksheets, vec![updates]);
            assert_eq!(visibility.touched_streams, vec![1]);
            assert_eq!(visibility.semantic_staging_plan_ns.len(), 1);
            assert_eq!(visibility.publication_ns.len(), 1);
            assert_eq!(
                visibility.source_workbook_bytes,
                visibility.target_workbook_bytes
            );
            if source_backed {
                assert_eq!(visibility.splice_count, Some(vec![updates]));
                assert_eq!(
                    visibility.replacement_bytes,
                    Some(vec![u64::try_from(updates).unwrap()])
                );
                assert_eq!(visibility.changed_spans.unwrap(), vec![updates]);
                assert_eq!(visibility.source_fingerprints.unwrap().len(), 1);
                assert_eq!(visibility.target_fingerprints.unwrap().len(), 1);
            } else {
                assert!(visibility.splice_count.is_none());
                assert!(visibility.replacement_bytes.is_none());
                assert!(visibility.changed_spans.is_none());
                assert!(visibility.source_fingerprints.is_none());
                assert!(visibility.target_fingerprints.is_none());
            }
        }
    }

    #[test]
    fn streaming_creation_evidence_is_fixed_window_and_reopen_verified() {
        for case in [Case::XlsxStreamingCreate, Case::RtfStreamingCreate] {
            let small = build_streaming_corpus(case, SemanticShape::Tiny).unwrap();
            let scaled = build_streaming_corpus(case, SemanticShape::Medium).unwrap();
            assert_eq!(
                small.metrics.retained_authoring_window_bytes,
                scaled.metrics.retained_authoring_window_bytes
            );
            assert!(scaled.manifest.archive_bytes > small.manifest.archive_bytes);
            assert!(scaled.metrics.input_bytes > small.metrics.input_bytes);
            assert!(
                u64::try_from(scaled.manifest.archive_bytes).unwrap()
                    > scaled.metrics.retained_authoring_window_bytes
            );

            let measured = run_streaming_creation(case, &small, 0, 1).unwrap();
            let sink = measured.sink.unwrap();
            assert_eq!(sink.retained_output_bytes, Some(0));
            assert_eq!(
                sink.retained_authoring_window_bytes,
                Some(small.metrics.retained_authoring_window_bytes)
            );
            assert_eq!(sink.accepted_bytes, small.manifest.archive_bytes as u64);
            assert_eq!(
                measured.output_sha256.as_deref(),
                Some(small.manifest.archive_sha256.as_str())
            );
        }
    }

    #[test]
    fn opc_source_overlay_save_is_deterministic_and_emits_complete_evidence() {
        let corpus = build_opc_corpus(CorpusShape::FewLarge, PayloadKind::Incompressible).unwrap();
        let replacement = opc_overlay_replacement_payload(&corpus).unwrap();
        let first = expected_opc_overlay_output(&corpus, &replacement).unwrap();
        let second = expected_opc_overlay_output(&corpus, &replacement).unwrap();
        assert_eq!(first, second);
        assert_eq!(
            corpus.manifest.archive_sha256,
            "a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6"
        );
        assert_eq!(
            sha256_hex(&first),
            "f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009"
        );

        let measured = run_opc_source_overlay_one_part_save(&corpus, 0, 2).unwrap();
        let digest = sha256_hex(&first);
        assert_eq!(measured.case, "opc_source_overlay_one_part_save");
        assert_eq!(measured.output_sha256.as_deref(), Some(digest.as_str()));
        assert_eq!(measured.elapsed_ns.samples.len(), 2);
        assert_eq!(
            measured.sink.unwrap().accepted_bytes,
            u64::try_from(first.len()).unwrap()
        );
        let source = measured.source.unwrap();
        assert_eq!(source.read_calls.len(), 2);
        assert!(source.read_calls.iter().all(|&calls| calls > 0));
        assert!(source.read_calls.windows(2).all(|pair| pair[0] == pair[1]));
        assert!(source.read_bytes.windows(2).all(|pair| pair[0] == pair[1]));
        assert_eq!(source.ordinary_payload_materializations, Some(vec![1, 1]));
    }

    #[test]
    fn opc_source_cache_budget_boundary_emits_exact_accounting_evidence() {
        let corpus = build_opc_corpus(CorpusShape::ManySmall, PayloadKind::Incompressible).unwrap();
        let measured = run_opc_source_cache_budget_boundary(&corpus, 0, 1).unwrap();

        assert_eq!(measured.len(), 2);
        let exact = measured[0]
            .source
            .as_ref()
            .unwrap()
            .opc_cache
            .as_ref()
            .unwrap();
        assert_eq!(exact.scenario, "exact-budget");
        assert_eq!(exact.diagnostics.successful_loads, vec![1]);
        assert_eq!(exact.diagnostics.budget_reservation_failures, vec![0]);
        assert_eq!(exact.budget_used_after_handles_drop, vec![1_024]);
        assert_eq!(exact.budget_used_after_package_drop, vec![0]);

        let refused_source = measured[1].source.as_ref().unwrap();
        let refused = refused_source.opc_cache.as_ref().unwrap();
        assert_eq!(refused.scenario, "one-under-budget");
        assert_eq!(refused_source.read_calls, vec![0]);
        assert_eq!(refused_source.read_bytes, vec![0]);
        assert_eq!(refused.diagnostics.cold_loads, vec![0]);
        assert_eq!(refused.diagnostics.budget_reservation_failures, vec![2]);
        assert_eq!(refused.diagnostics.budget_memory_used, vec![0]);
        assert_eq!(refused.budget_used_after_handles_drop, vec![0]);
        assert_eq!(refused.budget_used_after_package_drop, vec![0]);
    }

    #[test]
    fn opc_source_cache_contention_matrix_is_gated_and_mode_explicit() {
        let corpus = build_opc_corpus(CorpusShape::ManySmall, PayloadKind::Incompressible).unwrap();
        for (case, mode, managed) in [
            (
                Case::OpcSourceCacheControlContention,
                OpcCacheMode::Control,
                false,
            ),
            (
                Case::OpcSourceCacheManagedContention,
                OpcCacheMode::Managed,
                true,
            ),
        ] {
            let measured =
                run_opc_source_cache_contention(case, &corpus, 0, 1, &[1, 2], mode).unwrap();
            assert_eq!(measured.len(), 12);
            for result in &measured {
                assert_eq!(result.elapsed_ns.samples.len(), 1);
                let evidence = result.source.as_ref().unwrap().opc_cache.as_ref().unwrap();
                assert_eq!(evidence.persistent_worker_teams_created, 1);
                assert_eq!(evidence.diagnostics.cold_loads.len(), 1);
                assert_eq!(evidence.cache_mode == "budget-managed", managed);
                assert_eq!(
                    evidence.diagnostics.budget_memory_limit[0].is_some(),
                    managed
                );
                assert_eq!(evidence.budget_used_after_package_drop, vec![0]);
                let gate = evidence.gate.as_ref().unwrap();
                assert_eq!(gate.initial_arrivals.len(), 1);
                assert_eq!(gate.max_concurrent_delays, gate.initial_arrivals);
                if evidence.scenario == "same-part" {
                    assert!(evidence.scaling.p50_speedup.is_none());
                    assert!(evidence.scaling.amdahl_serial_fraction.is_none());
                } else {
                    assert!(evidence.scaling.p50_speedup.is_some());
                }
            }
        }
    }

    #[test]
    fn docx_source_edit_is_deterministic_and_emits_complete_evidence() {
        let corpus = build_docx_source_edit_corpus().unwrap();
        let again = build_docx_source_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(
            corpus.manifest.archive_sha256,
            "a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4"
        );
        let measured = run_docx_source_backed_one_edit_save(&corpus, 0, 1).unwrap();
        assert_eq!(measured.case, "docx_source_backed_one_edit_save");
        assert_eq!(measured.elapsed_ns.samples.len(), 1);
        assert!(measured.output_sha256.is_some());
        let source = measured.source.unwrap();
        assert_eq!(source.read_calls.len(), 1);
        assert_eq!(source.ordinary_payload_materializations, Some(vec![1]));
    }

    #[test]
    fn pptx_source_edit_is_deterministic_and_emits_complete_evidence() {
        let corpus = build_pptx_source_edit_corpus().unwrap();
        let again = build_pptx_source_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&corpus.archive));
        let measured = run_pptx_source_backed_one_edit_save(&corpus, 0, 1).unwrap();
        assert_eq!(measured.case, "pptx_source_backed_one_edit_save");
        assert_eq!(measured.elapsed_ns.samples.len(), 1);
        assert!(measured.output_sha256.is_some());
        let sink = measured.sink.unwrap();
        assert!(sink.largest_write <= 64 * 1024);
        let source = measured.source.unwrap();
        assert_eq!(source.read_calls.len(), 1);
        assert_eq!(source.ordinary_payload_materializations, Some(vec![2]));
    }

    #[test]
    fn pptx_batch_controls_are_deterministic_and_materialization_matched() {
        let corpus = build_pptx_source_edit_corpus().unwrap();
        let eager = run_pptx_batch_edit_save(Case::PptxEagerBatchEditSave, &corpus, 0, 1).unwrap();
        let source_backed =
            run_pptx_batch_edit_save(Case::PptxSourceBackedBatchEditSave, &corpus, 0, 1).unwrap();
        assert_eq!(eager.case, "pptx_eager_batch_edit_save");
        assert_eq!(source_backed.case, "pptx_source_backed_batch_edit_save");
        assert_eq!(eager.elapsed_ns.samples.len(), 1);
        assert_eq!(source_backed.elapsed_ns.samples.len(), 1);
        assert_eq!(eager.output_sha256, source_backed.output_sha256);
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![2])
        );
    }

    #[test]
    fn pptx_multi_slide_controls_are_deterministic_and_equivalent() {
        let corpus = build_pptx_source_edit_corpus().unwrap();
        let eager = run_pptx_multi_slide_batch_edit_save(
            Case::PptxEagerMultiSlideBatchEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        let source_backed = run_pptx_multi_slide_batch_edit_save(
            Case::PptxSourceBackedMultiSlideBatchEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        assert_eq!(eager.case, "pptx_eager_multi_slide_batch_edit_save");
        assert_eq!(
            source_backed.case,
            "pptx_source_backed_multi_slide_batch_edit_save"
        );
        assert_eq!(eager.elapsed_ns.samples.len(), 1);
        assert_eq!(source_backed.elapsed_ns.samples.len(), 1);
        assert_eq!(eager.output_sha256, source_backed.output_sha256);
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![(PPTX_MULTI_SLIDE_BATCH_COUNT + 1) as u64])
        );
    }

    #[test]
    fn xlsx_calculation_edit_controls_are_deterministic_and_equivalent() {
        let corpus = build_xlsx_calculation_metadata_edit_corpus().unwrap();
        let again = build_xlsx_calculation_metadata_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&corpus.archive));

        let eager = run_xlsx_calculation_metadata_edit_save(
            Case::XlsxEagerCalculationMetadataEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        let source_backed = run_xlsx_calculation_metadata_edit_save(
            Case::XlsxSourceBackedCalculationMetadataEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        assert_eq!(eager.case, "xlsx_eager_calculation_metadata_edit_save");
        assert_eq!(
            source_backed.case,
            "xlsx_source_backed_calculation_metadata_edit_save"
        );
        assert_eq!(eager.elapsed_ns.samples.len(), 1);
        assert_eq!(source_backed.elapsed_ns.samples.len(), 1);
        assert!(eager.output_sha256.is_some());
        assert!(source_backed.output_sha256.is_some());
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![1])
        );
    }

    #[test]
    fn xlsx_defined_name_edit_controls_are_deterministic_and_equivalent() {
        let corpus = build_xlsx_defined_names_edit_corpus().unwrap();
        let again = build_xlsx_defined_names_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&corpus.archive));

        let eager =
            run_xlsx_defined_names_edit_save(Case::XlsxEagerDefinedNamesEditSave, &corpus, 0, 1)
                .unwrap();
        let source_backed = run_xlsx_defined_names_edit_save(
            Case::XlsxSourceBackedDefinedNamesEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        assert_eq!(eager.case, "xlsx_eager_defined_names_edit_save");
        assert_eq!(
            source_backed.case,
            "xlsx_source_backed_defined_names_edit_save"
        );
        assert_eq!(eager.output_sha256, source_backed.output_sha256);
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![1])
        );
    }

    #[test]
    fn xlsx_page_break_edit_controls_are_deterministic_and_equivalent() {
        let corpus = build_xlsx_page_break_edit_corpus().unwrap();
        let again = build_xlsx_page_break_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&corpus.archive));

        let eager =
            run_xlsx_page_break_edit_save(Case::XlsxEagerPageBreakEditSave, &corpus, 0, 1).unwrap();
        let source_backed =
            run_xlsx_page_break_edit_save(Case::XlsxSourceBackedPageBreakEditSave, &corpus, 0, 1)
                .unwrap();
        assert_eq!(eager.case, "xlsx_eager_page_break_edit_save");
        assert_eq!(
            source_backed.case,
            "xlsx_source_backed_page_break_edit_save"
        );
        assert_eq!(eager.output_sha256, source_backed.output_sha256);
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![2])
        );
    }

    #[test]
    fn xlsx_page_margin_edit_controls_are_deterministic_and_equivalent() {
        let corpus = build_xlsx_page_margin_edit_corpus().unwrap();
        let again = build_xlsx_page_margin_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&corpus.archive));

        let eager =
            run_xlsx_page_margin_edit_save(Case::XlsxEagerPageMarginEditSave, &corpus, 0, 1)
                .unwrap();
        let source_backed =
            run_xlsx_page_margin_edit_save(Case::XlsxSourceBackedPageMarginEditSave, &corpus, 0, 1)
                .unwrap();
        assert_eq!(eager.case, "xlsx_eager_page_margin_edit_save");
        assert_eq!(
            source_backed.case,
            "xlsx_source_backed_page_margin_edit_save"
        );
        assert_eq!(eager.output_sha256, source_backed.output_sha256);
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![2])
        );
    }

    #[test]
    fn xlsx_print_options_edit_controls_are_deterministic_and_equivalent() {
        let corpus = build_xlsx_print_options_edit_corpus().unwrap();
        let again = build_xlsx_print_options_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&corpus.archive));

        let eager =
            run_xlsx_print_options_edit_save(Case::XlsxEagerPrintOptionsEditSave, &corpus, 0, 1)
                .unwrap();
        let source_backed = run_xlsx_print_options_edit_save(
            Case::XlsxSourceBackedPrintOptionsEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        assert_eq!(eager.case, "xlsx_eager_print_options_edit_save");
        assert_eq!(
            source_backed.case,
            "xlsx_source_backed_print_options_edit_save"
        );
        assert_eq!(eager.output_sha256, source_backed.output_sha256);
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![2])
        );
    }

    #[test]
    fn xlsx_page_setup_edit_controls_are_deterministic_and_equivalent() {
        let corpus = build_xlsx_page_setup_edit_corpus().unwrap();
        let again = build_xlsx_page_setup_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&corpus.archive));

        let eager =
            run_xlsx_page_setup_edit_save(Case::XlsxEagerPageSetupEditSave, &corpus, 0, 1).unwrap();
        let source_backed =
            run_xlsx_page_setup_edit_save(Case::XlsxSourceBackedPageSetupEditSave, &corpus, 0, 1)
                .unwrap();
        assert_eq!(eager.case, "xlsx_eager_page_setup_edit_save");
        assert_eq!(
            source_backed.case,
            "xlsx_source_backed_page_setup_edit_save"
        );
        assert_eq!(eager.output_sha256, source_backed.output_sha256);
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![2])
        );
    }

    #[test]
    fn xlsx_sheet_protection_edit_controls_are_deterministic_and_equivalent() {
        let corpus = build_xlsx_sheet_protection_edit_corpus().unwrap();
        let again = build_xlsx_sheet_protection_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&corpus.archive));

        let eager = run_xlsx_sheet_protection_edit_save(
            Case::XlsxEagerSheetProtectionEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        let source_backed = run_xlsx_sheet_protection_edit_save(
            Case::XlsxSourceBackedSheetProtectionEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        assert_eq!(eager.case, "xlsx_eager_sheet_protection_edit_save");
        assert_eq!(
            source_backed.case,
            "xlsx_source_backed_sheet_protection_edit_save"
        );
        assert_eq!(eager.output_sha256, source_backed.output_sha256);
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![2])
        );
    }

    #[test]
    fn xlsx_data_validation_edit_controls_are_deterministic_and_equivalent() {
        let corpus = build_xlsx_data_validation_edit_corpus().unwrap();
        let again = build_xlsx_data_validation_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&corpus.archive));

        let eager = run_xlsx_data_validation_edit_save(
            Case::XlsxEagerDataValidationEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        let source_backed = run_xlsx_data_validation_edit_save(
            Case::XlsxSourceBackedDataValidationEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        assert_eq!(eager.case, "xlsx_eager_data_validation_edit_save");
        assert_eq!(
            source_backed.case,
            "xlsx_source_backed_data_validation_edit_save"
        );
        assert_eq!(eager.output_sha256, source_backed.output_sha256);
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![2])
        );
    }

    #[test]
    fn xlsx_auto_filter_edit_controls_are_deterministic_and_equivalent() {
        let corpus = build_xlsx_auto_filter_edit_corpus().unwrap();
        let again = build_xlsx_auto_filter_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&corpus.archive));

        let eager =
            run_xlsx_auto_filter_edit_save(Case::XlsxEagerAutoFilterEditSave, &corpus, 0, 1)
                .unwrap();
        let source_backed =
            run_xlsx_auto_filter_edit_save(Case::XlsxSourceBackedAutoFilterEditSave, &corpus, 0, 1)
                .unwrap();
        assert_eq!(eager.case, "xlsx_eager_auto_filter_edit_save");
        assert_eq!(
            source_backed.case,
            "xlsx_source_backed_auto_filter_edit_save"
        );
        assert_eq!(eager.output_sha256, source_backed.output_sha256);
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![3])
        );
    }

    #[test]
    fn xlsx_conditional_formatting_edit_controls_are_deterministic_and_equivalent() {
        let corpus = build_xlsx_conditional_formatting_edit_corpus().unwrap();
        let again = build_xlsx_conditional_formatting_edit_corpus().unwrap();
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&corpus.archive));

        let eager = run_xlsx_conditional_formatting_edit_save(
            Case::XlsxEagerConditionalFormattingEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        let source_backed = run_xlsx_conditional_formatting_edit_save(
            Case::XlsxSourceBackedConditionalFormattingEditSave,
            &corpus,
            0,
            1,
        )
        .unwrap();
        assert_eq!(eager.case, "xlsx_eager_conditional_formatting_edit_save");
        assert_eq!(
            source_backed.case,
            "xlsx_source_backed_conditional_formatting_edit_save"
        );
        assert_eq!(eager.output_sha256, source_backed.output_sha256);
        assert_eq!(
            eager.source.unwrap().ordinary_payload_materializations,
            Some(vec![corpus.manifest.entry_count as u64])
        );
        assert_eq!(
            source_backed
                .source
                .unwrap()
                .ordinary_payload_materializations,
            Some(vec![3])
        );
    }

    #[test]
    fn semantic_docx_and_pptx_tiny_corpora_are_deterministic_and_editable() {
        let docx = build_semantic_docx_corpus(SemanticShape::Tiny).unwrap();
        let docx_again = build_semantic_docx_corpus(SemanticShape::Tiny).unwrap();
        assert_eq!(docx.archive, docx_again.archive);
        assert_eq!(docx.manifest.entry_count, 24);
        let docx_result = run_case(Case::DocxSemanticOnePercentEditSave, &docx, 0, 1).unwrap();
        assert!(docx_result.sink.is_some());

        let pptx = build_semantic_pptx_corpus(SemanticShape::Tiny).unwrap();
        let pptx_again = build_semantic_pptx_corpus(SemanticShape::Tiny).unwrap();
        assert_eq!(pptx.archive, pptx_again.archive);
        assert_eq!(pptx.manifest.entry_count, 12);
        let pptx_result = run_case(Case::PptxSemanticOnePercentEditSave, &pptx, 0, 1).unwrap();
        assert!(pptx_result.sink.is_none());
    }

    #[test]
    fn native_ole2_tiny_corpora_exercise_all_semantic_cases() {
        let families = [
            (
                Case::DocFreshWriteTo,
                [
                    Case::DocSemanticOpen,
                    Case::DocSemanticListParagraphs,
                    Case::DocSemanticOneParagraph,
                    Case::DocSemanticFullText,
                    Case::DocSemanticNoopEditSave,
                    Case::DocSemanticOneEditSave,
                ],
            ),
            (
                Case::XlsFreshWriteTo,
                [
                    Case::XlsSemanticOpen,
                    Case::XlsSemanticListWorksheets,
                    Case::XlsSemanticOneCell,
                    Case::XlsSemanticFullCellScan,
                    Case::XlsSemanticNoopEditSave,
                    Case::XlsSemanticOneEditSave,
                ],
            ),
            (
                Case::PptFreshWriteTo,
                [
                    Case::PptSemanticOpen,
                    Case::PptSemanticListSlides,
                    Case::PptSemanticOneShapeText,
                    Case::PptSemanticFullText,
                    Case::PptSemanticNoopEditSave,
                    Case::PptSemanticOneEditSave,
                ],
            ),
        ];

        for (writer_case, semantic_cases) in families {
            let corpus = build_writer_corpus(writer_case, WriterShape::Tiny).unwrap();
            let again = build_writer_corpus(writer_case, WriterShape::Tiny).unwrap();
            assert_eq!(corpus.archive, again.archive, "{}", writer_case.name());
            for case in semantic_cases {
                let measured = run_case(case, &corpus, 0, 1).unwrap();
                assert_eq!(measured.case, case.name());
                assert_eq!(measured.elapsed_ns.samples.len(), 1);
                assert!(measured.sink.is_none());
            }
        }

        let ppt = build_writer_corpus(Case::PptFreshWriteTo, WriterShape::Tiny).unwrap();
        for case in [
            Case::PptSlideOrderSnapshotOpen,
            Case::PptTextEditOneEditSave,
        ] {
            let measured = run_case(case, &ppt, 0, 1).unwrap();
            assert_eq!(measured.case, case.name());
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
            assert!(measured.sink.is_none());
        }
    }

    #[test]
    fn doc_body_snapshot_case_is_deterministic_and_semantic() {
        let corpus = build_writer_corpus(Case::DocFreshWriteTo, WriterShape::Tiny).unwrap();
        let measured = run_case(Case::DocBodySnapshotListParagraphs, &corpus, 0, 2).unwrap();
        assert_eq!(measured.case, "doc_body_snapshot_list_paragraphs");
        assert_eq!(measured.elapsed_ns.samples.len(), 2);
        assert!(measured.sink.is_none());
    }

    #[test]
    fn semantic_docx_medium_one_percent_edit_exercises_batch_publication() {
        let docx = build_semantic_docx_corpus(SemanticShape::Medium).unwrap();
        assert_eq!(docx.manifest.entry_count, 200);

        let result = run_case(Case::DocxSemanticOnePercentEditSave, &docx, 0, 1).unwrap();

        assert_eq!(result.case, "docx_semantic_one_percent_edit_save");
        assert!(result.sink.is_some());
    }

    #[test]
    fn semantic_rtf_tiny_variants_are_deterministic_and_capability_bounded() {
        let cases = [
            Case::RtfSemanticOpen,
            Case::RtfSemanticParagraphCount,
            Case::RtfSemanticListParagraphs,
            Case::RtfSemanticCollectParagraphs,
            Case::RtfSemanticOneParagraph,
            Case::RtfSemanticFullText,
            Case::RtfSemanticTextToSink,
            Case::RtfSemanticStreamSave,
            Case::RtfSemanticNoopEditSave,
            Case::RtfSemanticOneEditSave,
            Case::RtfSemanticOnePercentEditSave,
            Case::RtfSemanticRemoveParagraphSave,
            Case::RtfSemanticMoveParagraphSave,
            Case::RtfLogicalTailAppend,
            Case::RtfLogicalTailNoopSave,
        ];

        for variant in RtfSemanticVariant::ALL {
            let rtf = build_semantic_rtf_corpus(SemanticShape::Tiny, variant).unwrap();
            let again = build_semantic_rtf_corpus(SemanticShape::Tiny, variant).unwrap();
            let lifecycle = (variant == RtfSemanticVariant::Plain)
                .then(|| build_rtf_lifecycle_corpus(SemanticShape::Tiny).unwrap());
            assert_eq!(rtf.archive, again.archive, "{}", variant.name());
            assert_eq!(rtf.manifest.rtf_variant, Some(variant.name()));

            for case in cases {
                let selected_corpus = if case.is_rtf_lifecycle() || case.is_rtf_logical_tail() {
                    lifecycle.as_ref().unwrap_or(&rtf)
                } else {
                    &rtf
                };
                let result = run_case(case, selected_corpus, 0, 1);
                if variant.supports_case(case) {
                    let result = result.unwrap();
                    assert_eq!(
                        result.sink.is_some(),
                        case.name().contains("save")
                            || case == Case::RtfSemanticTextToSink
                            || case.is_rtf_logical_tail()
                    );
                } else {
                    assert!(
                        result.is_err(),
                        "{} unexpectedly supports {}",
                        variant.name(),
                        case.name()
                    );
                }
            }
        }

        let plain =
            build_semantic_rtf_corpus(SemanticShape::Tiny, RtfSemanticVariant::Plain).unwrap();
        assert_eq!(plain.manifest.entry_count, 24);
        assert_eq!(
            plain.manifest.archive_sha256,
            "ee4a5c5b5d1c97d5fb4f1e862c2787a859136b237addd0d14a7d52ddc9e62328"
        );
        let lifecycle = build_rtf_lifecycle_corpus(SemanticShape::Tiny).unwrap();
        assert_eq!(lifecycle.manifest.entry_count, 24);
        assert_eq!(
            lifecycle.manifest.archive_sha256,
            "73641cf09b630632deabce8585c67f395a6bd3ac01eedcca6a8b7224ef00d252"
        );

        let byte1252 =
            build_semantic_rtf_corpus(SemanticShape::Tiny, RtfSemanticVariant::Byte1252).unwrap();
        assert!(byte1252.archive.contains(&0xe9));
        assert_eq!(
            byte1252.manifest.archive_sha256,
            "47a20904dfb8107bb1cd9ad099decfed13c76cbde993fdd93eda3d919a9bb1aa"
        );
        let lzfu =
            build_semantic_rtf_corpus(SemanticShape::Tiny, RtfSemanticVariant::Lzfu).unwrap();
        assert!(litchi_rtf::transport::is_compressed_rtf(&lzfu.archive));
        assert_eq!(
            lzfu.manifest.archive_sha256,
            "bf755db7d4afc26a66ffab476884431e6e585f3259df5b6469e2d4fadfc51baf"
        );
        let watermark =
            build_semantic_rtf_corpus(SemanticShape::Tiny, RtfSemanticVariant::Watermark).unwrap();
        assert_eq!(
            watermark.manifest.archive_sha256,
            "48d62dcd959e737b06ebb8255780bcaaf1e88056ff9c3d5a21d3ff5cd3ddf9cb"
        );
        assert!(
            build_semantic_rtf_corpus(SemanticShape::Large, RtfSemanticVariant::Watermark).is_err()
        );
        assert_eq!(
            RtfSemanticVariant::ALL
                .iter()
                .flat_map(|variant| cases.iter().map(move |case| (*variant, *case)))
                .filter(|(variant, case)| variant.supports_case(*case))
                .count(),
            41
        );
    }

    #[test]
    fn semantic_rtf_logical_tail_is_bounded_reopenable_and_reversible() {
        for shape in [SemanticShape::Tiny, SemanticShape::Large] {
            let corpus = build_rtf_lifecycle_corpus(shape).unwrap();
            let append = run_case(Case::RtfLogicalTailAppend, &corpus, 0, 1).unwrap();
            let noop = run_case(Case::RtfLogicalTailNoopSave, &corpus, 0, 1).unwrap();
            for result in [append, noop] {
                assert_eq!(result.elapsed_ns.samples.len(), 1);
                assert!(result.output_sha256.is_some());
                let sink = result.sink.unwrap();
                let tail = sink.rtf_tail_append.unwrap();
                assert!(tail.exact_noop_verified);
                assert!(tail.in_memory_patch_verified);
                assert!(tail.durable_patch_verified);
                assert!(tail.reopen_verified);
                assert!(tail.source_conflict_verified);
                assert_eq!(sink.retained_output_bytes, Some(0));
                assert_eq!(
                    sink.retained_authoring_window_bytes,
                    Some(RTF_LOGICAL_TAIL_SINK_WINDOW_BYTES as u64)
                );
                assert!(sink.largest_write <= RTF_LOGICAL_TAIL_SINK_WINDOW_BYTES as u64);
            }
        }
    }

    #[test]
    fn semantic_rtf_lifecycle_cases_are_matched_durable_and_transport_bounded() {
        let corpus = build_rtf_lifecycle_corpus(SemanticShape::Tiny).unwrap();
        let source = litchi_rtf::Document::from_bytes(&corpus.archive).unwrap();

        for case in [
            Case::RtfSemanticRemoveParagraphSave,
            Case::RtfSemanticMoveParagraphSave,
        ] {
            let result = run_case(case, &corpus, 0, 1).unwrap();
            assert_eq!(result.case, case.name());
            assert_eq!(result.elapsed_ns.samples.len(), 1);
            let (expected_bytes, expected_sha256) = match case {
                Case::RtfSemanticRemoveParagraphSave => (
                    1_250,
                    "49ef949a6ee85cc3a1bce19026e10a3b953136c73997eec6f719940e2c0b37a2",
                ),
                Case::RtfSemanticMoveParagraphSave => (
                    1_304,
                    "9c7e42060e71be8cedf54fed9907d6a189efa45f1fec0f57d483e02af756f1fd",
                ),
                _ => unreachable!(),
            };
            assert_eq!(result.output_sha256.as_deref(), Some(expected_sha256));
            let sink = result.sink.unwrap();
            assert_eq!(sink.accepted_bytes, expected_bytes);
            assert!(sink.write_calls > 0);
            assert!(sink.largest_write <= sink.accepted_bytes);
        }

        let mut noop = source.edit();
        noop.move_paragraph(0, 0).unwrap();
        let noop = noop.commit().unwrap();
        assert!(!noop.diagnostics().changed());
        assert!(noop.snapshot().same_snapshot(&source));
        assert_eq!(noop.snapshot().to_bytes().unwrap(), corpus.archive);

        for variant in [RtfSemanticVariant::Byte1252, RtfSemanticVariant::Lzfu] {
            let corpus = build_semantic_rtf_corpus(SemanticShape::Tiny, variant).unwrap();
            let document = litchi_rtf::Document::from_bytes(&corpus.archive).unwrap();
            let exact = document.to_bytes().unwrap();
            let selected = document.paragraph_count() / 2;

            let mut remove = document.edit();
            remove.remove_paragraph(selected).unwrap();
            assert!(matches!(
                remove.commit(),
                Err(litchi_rtf::edit::Error::UnsupportedSource(_))
            ));
            assert_eq!(document.to_bytes().unwrap(), exact);

            let mut reorder = document.edit();
            reorder
                .move_paragraph(0, document.paragraph_count() - 1)
                .unwrap();
            assert!(matches!(
                reorder.commit(),
                Err(litchi_rtf::edit::Error::UnsupportedSource(_))
            ));
            assert_eq!(document.to_bytes().unwrap(), exact);

            let mut noop = document.edit();
            noop.move_paragraph(0, 0).unwrap();
            let noop = noop.commit().unwrap();
            assert!(!noop.diagnostics().changed());
            assert!(noop.snapshot().same_snapshot(&document));
            assert_eq!(noop.snapshot().to_bytes().unwrap(), exact);
        }

        let watermark =
            build_semantic_rtf_corpus(SemanticShape::Tiny, RtfSemanticVariant::Watermark).unwrap();
        let watermark = litchi_rtf::Document::from_bytes(&watermark.archive).unwrap();
        let exact = watermark.to_bytes().unwrap();
        let mut remove = watermark.edit();
        remove.remove_paragraph(0).unwrap();
        assert!(matches!(
            remove.commit(),
            Err(litchi_rtf::edit::Error::UnsupportedSource(_))
        ));
        assert_eq!(watermark.to_bytes().unwrap(), exact);
        let mut noop = watermark.edit();
        noop.move_paragraph(0, 0).unwrap();
        let noop = noop.commit().unwrap();
        assert!(!noop.diagnostics().changed());
        assert!(noop.snapshot().same_snapshot(&watermark));
        assert_eq!(noop.snapshot().to_bytes().unwrap(), exact);

        let opaque =
            litchi_rtf::Document::parse(r"{\rtf1\ansi First\par \b Opaque\b0\par Third}").unwrap();
        let exact = opaque.to_bytes().unwrap();
        for case in [
            Case::RtfSemanticRemoveParagraphSave,
            Case::RtfSemanticMoveParagraphSave,
        ] {
            let mut edit = opaque.edit();
            match case {
                Case::RtfSemanticRemoveParagraphSave => {
                    edit.remove_paragraph(1).unwrap();
                },
                Case::RtfSemanticMoveParagraphSave => {
                    edit.move_paragraph(0, 2).unwrap();
                },
                _ => unreachable!(),
            }
            assert!(matches!(
                edit.commit(),
                Err(litchi_rtf::edit::Error::UnsupportedSource(_))
            ));
            assert_eq!(opaque.to_bytes().unwrap(), exact);
        }
        let mut opaque_noop = opaque.edit();
        opaque_noop.move_paragraph(0, 0).unwrap();
        let opaque_noop = opaque_noop.commit().unwrap();
        assert!(!opaque_noop.diagnostics().changed());
        assert!(opaque_noop.snapshot().same_snapshot(&opaque));
        assert_eq!(opaque_noop.snapshot().to_bytes().unwrap(), exact);
    }

    #[test]
    fn semantic_rtf_text_sink_is_utf8_bounded_and_transport_independent() {
        for variant in [
            RtfSemanticVariant::Plain,
            RtfSemanticVariant::Byte1252,
            RtfSemanticVariant::Lzfu,
        ] {
            let corpus = build_semantic_rtf_corpus(SemanticShape::Tiny, variant).unwrap();
            let result = run_case(Case::RtfSemanticTextToSink, &corpus, 0, 1).unwrap();
            let sink = result.sink.unwrap();
            assert_eq!(result.case, "rtf_semantic_text_to_sink");
            assert_eq!(
                sink.accepted_bytes,
                u64::try_from(
                    super::semantic_rtf_expected_text(SemanticShape::Tiny, variant, &[]).len()
                )
                .unwrap()
            );
            assert!(sink.write_calls > 0);
            assert!(sink.largest_write <= sink.accepted_bytes);
        }

        let watermark =
            build_semantic_rtf_corpus(SemanticShape::Tiny, RtfSemanticVariant::Watermark).unwrap();
        assert!(run_case(Case::RtfSemanticTextToSink, &watermark, 0, 1).is_err());
    }

    #[test]
    fn semantic_rtf_medium_one_percent_batch_is_atomic_reopenable_and_reversible() {
        let corpus =
            build_semantic_rtf_corpus(SemanticShape::Medium, RtfSemanticVariant::Plain).unwrap();
        let result = run_case(Case::RtfSemanticOnePercentEditSave, &corpus, 0, 1).unwrap();
        assert_eq!(result.case, "rtf_semantic_one_percent_edit_save");
        assert!(result.sink.is_some());

        let source = litchi_rtf::Document::from_bytes(&corpus.archive).unwrap();
        let mut invalid = super::semantic_update_indices(source.paragraph_count())
            .unwrap()
            .into_iter()
            .map(|position| {
                litchi_rtf::edit::ParagraphTextReplacement::new(
                    position,
                    super::semantic_rtf_variant_text(RtfSemanticVariant::Plain, position, true),
                )
            })
            .collect::<Vec<_>>();
        invalid.push(litchi_rtf::edit::ParagraphTextReplacement::new(
            source.paragraph_count(),
            "out-of-range",
        ));
        let mut edit = source.edit();
        assert!(edit.replace_body_paragraph_texts(&invalid).is_err());
        assert_eq!(edit.operation_count(), 0);
        assert!(edit.commit().unwrap().snapshot().same_snapshot(&source));
    }

    #[test]
    fn semantic_rtf_large_evidence_output_hashes_are_stable() {
        let shape = SemanticShape::Large;
        let variant = RtfSemanticVariant::Plain;
        let corpus = build_semantic_rtf_corpus(shape, variant).unwrap();
        assert_eq!(
            corpus.manifest.archive_sha256,
            "957645f9109433d8dc25a66e384a496b19a97ed5ff4fab4bb981f8cda3c6e02e"
        );

        let source = litchi_rtf::Document::from_bytes(&corpus.archive).unwrap();
        let replacements = super::semantic_update_indices(source.paragraph_count())
            .unwrap()
            .into_iter()
            .map(|position| {
                litchi_rtf::edit::ParagraphTextReplacement::new(
                    position,
                    super::semantic_rtf_variant_text(variant, position, true),
                )
            })
            .collect::<Vec<_>>();
        let mut edit = source.edit();
        edit.replace_body_paragraph_texts(&replacements).unwrap();
        let output = edit.commit().unwrap().snapshot().to_bytes().unwrap();
        assert_eq!(output.len(), 540_151);
        assert_eq!(
            sha256_hex(&output),
            "d040328cb691fc5ec65192477688f4a9a4275a8b62fa354a2fdb68d739786d8f"
        );

        for (variant, expected_bytes, expected_sha256) in [
            (
                RtfSemanticVariant::Plain,
                499_999,
                "f122900c5e612d3218061ddd766889f8c2520c4e3b51e6aa3df6158fc736abdc",
            ),
            (
                RtfSemanticVariant::Byte1252,
                529_999,
                "b6fb32c1fb727d747630e2667091e1c119c7e4ae4653036f702078580e4a5872",
            ),
            (
                RtfSemanticVariant::Lzfu,
                499_999,
                "f122900c5e612d3218061ddd766889f8c2520c4e3b51e6aa3df6158fc736abdc",
            ),
        ] {
            let text = super::semantic_rtf_expected_text(shape, variant, &[]);
            assert_eq!(text.len(), expected_bytes);
            assert_eq!(sha256_hex(text.as_bytes()), expected_sha256);
        }
    }

    #[test]
    fn semantic_odf_tiny_corpora_are_deterministic_and_editable() {
        let odt = build_semantic_odt_corpus(SemanticShape::Tiny).unwrap();
        assert_eq!(
            odt.archive,
            build_semantic_odt_corpus(SemanticShape::Tiny)
                .unwrap()
                .archive
        );
        assert_eq!(odt.manifest.entry_count, 24);
        run_case(Case::OdtSemanticOneEditSave, &odt, 0, 1).unwrap();
        run_case(Case::OdtSemanticOnePercentEditSave, &odt, 0, 1).unwrap();

        let ods = build_semantic_ods_corpus(SemanticShape::Tiny).unwrap();
        assert_eq!(
            ods.archive,
            build_semantic_ods_corpus(SemanticShape::Tiny)
                .unwrap()
                .archive
        );
        assert_eq!(ods.manifest.entry_count, 64);
        run_case(Case::OdsSemanticCellSweep, &ods, 0, 1).unwrap();
        run_case(Case::OdsSemanticOneEditSave, &ods, 0, 1).unwrap();
        run_case(Case::OdsSemanticOnePercentEditSave, &ods, 0, 1).unwrap();

        let odp = build_semantic_odp_corpus(SemanticShape::Tiny).unwrap();
        assert_eq!(
            odp.archive,
            build_semantic_odp_corpus(SemanticShape::Tiny)
                .unwrap()
                .archive
        );
        assert_eq!(odp.manifest.entry_count, 3);
        run_case(Case::OdpSemanticOneEditSave, &odp, 0, 1).unwrap();
    }

    #[test]
    fn semantic_odt_medium_one_percent_edit_exercises_repeated_publication() {
        let odt = build_semantic_odt_corpus(SemanticShape::Medium).unwrap();
        assert_eq!(odt.manifest.entry_count, 200);

        let result = run_case(Case::OdtSemanticOnePercentEditSave, &odt, 0, 1).unwrap();

        assert_eq!(result.case, "odt_semantic_one_percent_edit_save");
        assert!(result.sink.is_none());
    }

    #[test]
    fn semantic_ods_medium_one_percent_edit_exercises_atomic_cell_batches() {
        let ods = build_semantic_ods_corpus(SemanticShape::Medium).unwrap();
        assert_eq!(ods.manifest.entry_count, 2_048);

        let result = run_case(Case::OdsSemanticOnePercentEditSave, &ods, 0, 1).unwrap();

        assert_eq!(result.case, "ods_semantic_one_percent_edit_save");
        assert!(result.sink.is_none());
    }

    #[test]
    fn media_rich_ods_corpus_is_deterministic_and_preserved_by_one_cell_edit() {
        let first = build_ods_media_corpus().unwrap();
        let second = build_ods_media_corpus().unwrap();

        assert_eq!(first.archive, second.archive);
        assert_eq!(first.manifest.generator, "litchi-ods-media-publication-v1");
        assert_eq!(first.manifest.shape, "media-rich");
        assert_eq!(first.manifest.entry_bytes, 2 * 1024 * 1024);
        let result = run_case(Case::OdsMediaOneEditSave, &first, 0, 1).unwrap();
        assert_eq!(result.case, "ods_media_one_edit_save");
        assert_eq!(result.elapsed_ns.samples.len(), 1);

        let snapshot = litchi_ods::document::Snapshot::from_bytes(first.archive.clone()).unwrap();
        let mut edit = snapshot.edit();
        edit.worksheets(|worksheets| {
            let text = super::semantic_ods_text(1, 16, 16, true);
            worksheets
                .set_cell(
                    "Sheet 1",
                    16,
                    16,
                    litchi_ods::Cell::new(litchi_ods::CellValue::Text(text.clone()), text),
                )?
                .ok_or_else(|| {
                    litchi_core::Error::InvalidFormat(
                        "media-rich ODS test sheet is missing".to_owned(),
                    )
                })?;
            Ok(())
        })
        .unwrap();
        let output = edit.commit().unwrap();
        let identical = litchi_odf_common::package::raw_identical_members(
            &first.archive,
            output.snapshot().as_bytes(),
        )
        .unwrap();
        assert!(!identical.contains("content.xml"));
        assert!(identical.contains("META-INF/manifest.xml"));
        for index in 0..super::ODS_MEDIA_ENTRY_COUNT {
            assert!(identical.contains(&super::ods_media_path(index)));
        }
    }

    #[test]
    fn media_rich_odt_corpus_is_deterministic_and_preserved_by_paragraph_edit() {
        let first = build_odt_media_corpus().unwrap();
        let second = build_odt_media_corpus().unwrap();

        assert_eq!(first.archive, second.archive);
        assert_eq!(
            first.manifest.generator,
            "litchi-odt-media-paragraph-publication-v1"
        );
        assert_eq!(first.manifest.shape, "media-rich");
        assert_eq!(first.manifest.entry_count, 208);
        assert_eq!(first.manifest.entry_bytes, 2 * 1024 * 1024);
        let result = run_case(Case::OdtMediaParagraphEditSave, &first, 0, 1).unwrap();
        assert_eq!(result.case, "odt_media_paragraph_edit_save");
        assert_eq!(result.elapsed_ns.samples.len(), 1);

        let source = litchi_odt::transaction::Snapshot::from_bytes(first.archive.clone()).unwrap();
        let target = SemanticShape::Medium.docx_paragraphs() / 2;
        let mut edit = source.edit();
        edit.replace_paragraph(
            litchi_odt::transaction::Position::new(target),
            super::semantic_odt_text(target, true),
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        let identical = litchi_odf_common::package::raw_identical_members(
            &first.archive,
            commit.snapshot().as_bytes(),
        )
        .unwrap();
        assert!(!identical.contains("content.xml"));
        for path in [
            "mimetype",
            "styles.xml",
            "meta.xml",
            "META-INF/manifest.xml",
        ] {
            assert!(identical.contains(path), "{path}");
        }
        for index in 0..super::ODS_MEDIA_ENTRY_COUNT {
            assert!(identical.contains(&super::odt_media_path(index)));
        }
    }

    #[test]
    fn media_rich_odt_line_break_is_deterministic_and_preserves_untouched_members() {
        let corpus = build_odt_media_corpus().unwrap();
        let result = run_case(Case::OdtMediaLineBreakEditSave, &corpus, 0, 1).unwrap();
        assert_eq!(result.case, "odt_media_line_break_edit_save");
        assert_eq!(result.elapsed_ns.samples.len(), 1);
        assert!(result.output_sha256.is_some());

        let source = litchi_odt::transaction::Snapshot::from_bytes(corpus.archive.clone()).unwrap();
        let target = SemanticShape::Medium.docx_paragraphs() / 2;
        let mut edit = source.edit();
        edit.append_line_break(litchi_odt::transaction::ParagraphSelector::position(
            litchi_core::Position::new(target),
        ))
        .unwrap();
        let commit = edit.commit().unwrap();
        super::verify_odt_media_line_break_archive(commit.snapshot().as_bytes()).unwrap();
        let identical = litchi_odf_common::package::raw_identical_members(
            &corpus.archive,
            commit.snapshot().as_bytes(),
        )
        .unwrap();
        assert!(!identical.contains("content.xml"));
        for path in [
            "mimetype",
            "styles.xml",
            "meta.xml",
            "META-INF/manifest.xml",
        ] {
            assert!(identical.contains(path), "{path}");
        }
        for index in 0..super::ODS_MEDIA_ENTRY_COUNT {
            assert!(identical.contains(&super::odt_media_path(index)));
        }
    }

    #[test]
    fn media_rich_odt_append_run_is_deterministic_and_preserves_untouched_members() {
        let corpus = build_odt_media_corpus().unwrap();
        let result = run_case(Case::OdtMediaAppendRunEditSave, &corpus, 0, 1).unwrap();
        assert_eq!(result.case, "odt_media_append_run_edit_save");
        assert_eq!(result.elapsed_ns.samples.len(), 1);
        assert!(result.output_sha256.is_some());

        let source = litchi_odt::transaction::Snapshot::from_bytes(corpus.archive.clone()).unwrap();
        let target = SemanticShape::Medium.docx_paragraphs() / 2;
        let mut edit = source.edit();
        edit.append_run(
            litchi_core::Position::new(target),
            super::ODT_MEDIA_APPEND_RUN_TEXT,
            None,
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        super::verify_odt_media_append_run_archive(commit.snapshot().as_bytes()).unwrap();
        let identical = litchi_odf_common::package::raw_identical_members(
            &corpus.archive,
            commit.snapshot().as_bytes(),
        )
        .unwrap();
        assert!(!identical.contains("content.xml"));
        for path in [
            "mimetype",
            "styles.xml",
            "meta.xml",
            "META-INF/manifest.xml",
        ] {
            assert!(identical.contains(path), "{path}");
        }
        for index in 0..super::ODS_MEDIA_ENTRY_COUNT {
            assert!(identical.contains(&super::odt_media_path(index)));
        }
    }

    #[test]
    fn media_rich_odt_append_hyperlink_is_deterministic_and_preserves_untouched_members() {
        let corpus = build_odt_media_corpus().unwrap();
        let result = run_case(Case::OdtMediaAppendHyperlinkEditSave, &corpus, 0, 1).unwrap();
        assert_eq!(result.case, "odt_media_append_hyperlink_edit_save");
        assert_eq!(result.elapsed_ns.samples.len(), 1);
        assert!(result.output_sha256.is_some());

        let source = litchi_odt::transaction::Snapshot::from_bytes(corpus.archive.clone()).unwrap();
        let target = SemanticShape::Medium.docx_paragraphs() / 2;
        let mut edit = source.edit();
        edit.append_hyperlink(
            litchi_core::Position::new(target),
            super::ODT_MEDIA_APPEND_HYPERLINK_HREF,
            super::ODT_MEDIA_APPEND_HYPERLINK_TEXT,
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        super::verify_odt_media_append_hyperlink_archive(commit.snapshot().as_bytes()).unwrap();
        let identical = litchi_odf_common::package::raw_identical_members(
            &corpus.archive,
            commit.snapshot().as_bytes(),
        )
        .unwrap();
        assert!(!identical.contains("content.xml"));
        for path in [
            "mimetype",
            "styles.xml",
            "meta.xml",
            "META-INF/manifest.xml",
        ] {
            assert!(identical.contains(path), "{path}");
        }
        for index in 0..super::ODS_MEDIA_ENTRY_COUNT {
            assert!(identical.contains(&super::odt_media_path(index)));
        }
    }

    #[test]
    fn media_rich_odt_structural_paragraph_edits_are_deterministic_and_preserve_members() {
        let corpus = build_odt_media_corpus().unwrap();
        let target = SemanticShape::Medium.docx_paragraphs() / 2;
        for (case, inserted) in [
            (Case::OdtMediaInsertParagraphEditSave, true),
            (Case::OdtMediaRemoveParagraphEditSave, false),
        ] {
            let result = run_case(case, &corpus, 0, 1).unwrap();
            assert_eq!(result.case, case.name());
            assert_eq!(result.elapsed_ns.samples.len(), 1);
            assert!(result.output_sha256.is_some());

            let source =
                litchi_odt::transaction::Snapshot::from_bytes(corpus.archive.clone()).unwrap();
            let mut edit = source.edit();
            if inserted {
                edit.insert_paragraph(
                    litchi_core::Position::new(target),
                    super::ODT_MEDIA_INSERT_PARAGRAPH_TEXT,
                )
                .unwrap();
            } else {
                edit.remove_paragraph(litchi_core::Position::new(target))
                    .unwrap();
            }
            let commit = edit.commit().unwrap();
            super::verify_odt_media_structural_paragraph_archive(
                commit.snapshot().as_bytes(),
                inserted,
            )
            .unwrap();
            let identical = litchi_odf_common::package::raw_identical_members(
                &corpus.archive,
                commit.snapshot().as_bytes(),
            )
            .unwrap();
            assert!(!identical.contains("content.xml"));
            for path in [
                "mimetype",
                "styles.xml",
                "meta.xml",
                "META-INF/manifest.xml",
            ] {
                assert!(identical.contains(path), "{path}");
            }
            for index in 0..super::ODS_MEDIA_ENTRY_COUNT {
                assert!(identical.contains(&super::odt_media_path(index)));
            }
        }
    }

    #[test]
    fn media_rich_odt_scalar_and_batch_resource_replacements_are_matched() {
        let first = build_odt_resource_batch_corpus().unwrap();
        let second = build_odt_resource_batch_corpus().unwrap();
        assert_eq!(first.archive, second.archive);
        assert_eq!(first.manifest.entry_count, 272);
        assert_eq!(first.manifest.archive_member_count, 77);
        assert_eq!(first.manifest.archive_bytes, 17_061_898);
        assert_eq!(
            first.manifest.archive_sha256,
            "7b0ddd1c00ef91d24e60f30bf4a0ca0045807d537329e213f2f03020dfb0750b"
        );
        assert_eq!(first.manifest.shape, "media-rich-64-image-owners");
        assert_eq!(
            first.manifest.generator,
            "litchi-odt-embedded-resource-batch-publication-v1"
        );

        let scalar = run_case(Case::OdtEmbeddedResourceScalarReplaceSave, &first, 0, 1).unwrap();
        let batch = run_case(Case::OdtEmbeddedResourceBatchReplaceSave, &first, 0, 1).unwrap();
        assert_eq!(
            scalar.output_sha256.as_deref(),
            Some("2da19ec3aff1f8cf76a2690a498bb9582b604c0aab25cd40c3b688efa5888a1d")
        );
        assert_eq!(
            batch.output_sha256.as_deref(),
            Some("fa71c846111de90d5cfed8e6a95493126baad291f4ef4d9f4905bf65fc54e896")
        );
        for (measured, expected_bytes) in [(scalar, 17_336_931), (batch, 17_336_924)] {
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
            assert!(measured.source.is_none());
            assert!(measured.output_sha256.is_some());
            let sink = measured.sink.unwrap();
            assert_eq!(sink.write_calls, 1);
            assert_eq!(sink.accepted_bytes, expected_bytes);
            assert_eq!(sink.largest_write, expected_bytes);
        }

        let document = litchi_odt::Document::from_bytes(first.archive.clone()).unwrap();
        assert_eq!(document.images().unwrap().len(), ODT_RESOURCE_BATCH_COUNT);
        assert!(document.images().unwrap().iter().all(|image| {
            image
                .frame
                .as_ref()
                .and_then(|frame| frame.name.as_deref())
                .is_some_and(|name| name.starts_with("litchi-perf-odt-resource-owner-"))
        }));
    }

    #[test]
    fn media_rich_odp_corpus_is_deterministic_and_preserved_by_text_box_edit() {
        let first = build_odp_media_corpus().unwrap();
        let second = build_odp_media_corpus().unwrap();

        assert_eq!(first.archive, second.archive);
        assert_eq!(
            first.manifest.generator,
            "litchi-odp-media-textbox-publication-v1"
        );
        assert_eq!(first.manifest.shape, "media-rich");
        assert_eq!(first.manifest.entry_bytes, 2 * 1024 * 1024);
        let result = run_case(Case::OdpMediaTextBoxEditSave, &first, 0, 1).unwrap();
        assert_eq!(result.case, "odp_media_textbox_edit_save");
        assert_eq!(result.elapsed_ns.samples.len(), 1);

        let source =
            litchi_odp::authoring::edit::Snapshot::from_bytes(first.archive.clone()).unwrap();
        let mut transaction = source.transaction().unwrap();
        transaction
            .add_text_box(
                0usize,
                &litchi_odp::content::TextBox::new(
                    super::ODP_MEDIA_TEXT_BOX_NAME,
                    litchi_odp::content::RichText::plain(super::odp_media_text()).unwrap(),
                )
                .unwrap(),
            )
            .unwrap();
        let commit = transaction.commit().unwrap();
        let identical = litchi_odf_common::package::raw_identical_members(
            &first.archive,
            commit.snapshot().bytes(),
        )
        .unwrap();
        assert!(!identical.contains("content.xml"));
        for path in [
            "mimetype",
            "styles.xml",
            "meta.xml",
            "META-INF/manifest.xml",
        ] {
            assert!(identical.contains(path), "{path}");
        }
        for index in 0..super::ODS_MEDIA_ENTRY_COUNT {
            assert!(identical.contains(&super::odp_media_path(index)));
        }
    }

    #[test]
    fn media_rich_odp_scalar_and_batch_text_box_replacements_are_matched() {
        let first = build_odp_text_box_batch_corpus().unwrap();
        let second = build_odp_text_box_batch_corpus().unwrap();
        assert_eq!(first.archive, second.archive);
        assert_eq!(first.manifest.entry_count, 28);
        assert_eq!(first.manifest.archive_member_count, 13);
        assert_eq!(first.manifest.uncompressed_payload_bytes, 16_778_604);
        assert_eq!(first.manifest.archive_bytes, 16_786_244);
        assert_eq!(
            first.manifest.archive_sha256,
            "dcbb1f88da9366f2eab8eb6029dcc73930ea2fc03552b78dd4922689f8a9655d"
        );
        assert_eq!(first.manifest.shape, "media-rich-cross-slide");
        assert_eq!(
            first.manifest.generator,
            "litchi-odp-cross-slide-textbox-publication-v1"
        );

        let scalar = run_case(Case::OdpMediaTextBoxScalarReplaceSave, &first, 0, 1).unwrap();
        let batch = run_case(Case::OdpMediaTextBoxBatchReplaceSave, &first, 0, 1).unwrap();
        assert_eq!(
            scalar.output_sha256.as_deref(),
            Some("ee31f8c046af7b99819b183ca4fc56e00b97d2f97b36fa776c7d4c96dee3614b")
        );
        assert_eq!(
            batch.output_sha256.as_deref(),
            Some("fb4243a5433028d050ea97a5cb8db18c1af2ef66bb0d75071c95c2d9e83ec3cf")
        );
        for (measured, expected_bytes) in [(scalar, 16_786_370), (batch, 16_786_368)] {
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
            assert!(measured.source.is_none());
            let sink = measured.sink.unwrap();
            assert_eq!(sink.write_calls, 1);
            assert_eq!(sink.accepted_bytes, expected_bytes);
            assert_eq!(sink.accepted_bytes, sink.largest_write);
        }

        let source =
            litchi_odp::authoring::edit::Snapshot::from_bytes(first.archive.clone()).unwrap();
        let inventory = source.rich_content().unwrap();
        assert_eq!(
            inventory
                .text_boxes()
                .iter()
                .filter(|model| model.name().starts_with("litchi-perf-odp-batch-text-box-"))
                .count(),
            ODP_TEXT_BOX_BATCH_COUNT
        );
        assert!(inventory.text_boxes().iter().all(|model| {
            !model.name().starts_with("litchi-perf-odp-batch-text-box-")
                || model.xml().contains("-source-")
        }));
    }

    #[test]
    fn cfb_corpus_generation_is_deterministic_and_targets_last_root_stream() {
        let first = build_cfb_corpus(CorpusShape::Tiny, PayloadKind::Compressible).unwrap();
        let second = build_cfb_corpus(CorpusShape::Tiny, PayloadKind::Compressible).unwrap();

        assert_eq!(first.archive, second.archive);
        assert_eq!(first.target_name, "benchmark_stream_00002.bin");
        assert_eq!(first.manifest.archive_member_count, 3);
        assert_eq!(first.manifest.package_format, "CFB/OLE2");
        assert_eq!(
            first.manifest.archive_sha256,
            "84f1dcb6b2f35b87a9abb4aa783cc9988172b368cb72956d72b11c0a0ec5f282"
        );
    }

    #[test]
    fn cfb_cases_run_against_tiny_smoke_corpus() {
        let corpus = build_cfb_corpus(CorpusShape::Tiny, PayloadKind::Incompressible).unwrap();
        let cases = [
            Case::CfbOpen,
            Case::CfbListStreams,
            Case::CfbReadOne,
            Case::CfbCreateStreamBorrowed,
            Case::CfbCreateStreamOwned,
            Case::CfbSharedOpen,
            Case::CfbSharedReadOne,
            Case::CfbSharedConcurrentReads,
        ];

        for case in cases {
            let measured = run_case(case, &corpus, 0, 1).unwrap();
            assert_eq!(measured.case, case.name());
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
            if matches!(
                case,
                Case::CfbSharedOpen | Case::CfbSharedReadOne | Case::CfbSharedConcurrentReads
            ) {
                assert_eq!(measured.source.as_ref().unwrap().read_calls.len(), 1);
            }
        }
    }

    #[test]
    fn ole_common_heavy_edit_is_deterministic_and_preserves_unchanged_streams() {
        let base = build_cfb_corpus(CorpusShape::FewLarge, PayloadKind::Incompressible).unwrap();
        let corpus = build_ole_common_corpus(&base).unwrap();
        assert_eq!(corpus.manifest.entry_count, 5);
        assert_eq!(corpus.manifest.entry_bytes, 4 * 1024 * 1024);
        assert_eq!(
            corpus.manifest.archive_sha256,
            "7ffbd37c3e472a21b382bcbb02e430a62164e58d2270bbee0deaa584ff47a94d"
        );

        let first = ole_common_changed_output(&corpus).unwrap();
        let second = ole_common_changed_output(&corpus).unwrap();
        assert_eq!(first, second);
        assert_eq!(
            super::sha256_hex(&first),
            "b9323eeace80e2c9c88801879265bfdfac83690bb2550880f5ef6bf87b48d131"
        );

        for case in [
            Case::OleCommonOpen,
            Case::OleCommonPutStreamPublish,
            Case::OleCommonFinishRender,
            Case::OleCommonOneEditSave,
        ] {
            let measured = run_case(case, &base, 0, 1).unwrap();
            assert_eq!(measured.case, case.name());
            assert_eq!(measured.corpus.entry_count, 5);
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
        }
    }

    #[test]
    fn source_backed_opc_cases_prove_deferred_cached_and_single_flight_reads() {
        let corpus = build_opc_corpus(CorpusShape::Tiny, PayloadKind::Compressible).unwrap();
        for case in [
            Case::OpcSourceOpen,
            Case::OpcSourceOpenMainRead,
            Case::OpcSourceCachedMainRead,
            Case::OpcSourceConcurrentSamePart,
        ] {
            let measured = run_case(case, &corpus, 0, 1).unwrap();
            let source = measured.source.unwrap();
            assert_eq!(source.read_calls.len(), 1, "{}", case.name());
            assert_eq!(source.read_bytes.len(), 1, "{}", case.name());
            assert_eq!(
                source.ordinary_payload_materializations.unwrap(),
                vec![u64::from(matches!(
                    case,
                    Case::OpcSourceOpenMainRead | Case::OpcSourceConcurrentSamePart
                ))],
                "{}",
                case.name()
            );
            if case == Case::OpcSourceCachedMainRead {
                assert_eq!(source.read_calls, vec![0]);
                assert_eq!(source.read_bytes, vec![0]);
            }
        }
    }

    #[test]
    fn xlsx_tiny_corpus_is_deterministic_and_all_xlsx_cases_smoke() {
        let first = build_xlsx_corpus(XlsxShape::Tiny).unwrap();
        let second = build_xlsx_corpus(XlsxShape::Tiny).unwrap();

        assert_eq!(first.archive, second.archive);
        assert_eq!(first.manifest.generator, "litchi-xlsx-synthetic-v1");
        assert_eq!(first.manifest.package_format, "XLSX/OPC/ZIP");
        assert_eq!(first.manifest.entry_count, 192);
        assert_eq!(
            first.manifest.archive_sha256,
            "69ef199769a316eaa465a41ebf08f7a1b501f708775fabd7a084a90dc6a9b428"
        );
        let xlsx = first.manifest.xlsx.as_ref().unwrap();
        assert_eq!(xlsx.sheet_count, 3);
        assert_eq!(xlsx.rows_per_sheet, 8);
        assert_eq!(xlsx.columns_per_sheet, 8);
        assert_eq!(xlsx.one_percent_update_count, 2);
        assert_eq!(xlsx.source_members.workbook, "xl/workbook.xml");
        assert_eq!(
            xlsx.source_members.worksheets,
            [
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/sheet2.xml",
                "xl/worksheets/sheet3.xml",
            ]
        );
        assert_eq!(xlsx.source_members.shared_strings, None);
        assert_eq!(xlsx.source_members.styles.as_deref(), Some("xl/styles.xml"));

        for case in Case::DEFAULT.into_iter().filter(|case| case.uses_xlsx()) {
            let measured = run_case(case, &first, 0, 1).unwrap();
            assert_eq!(measured.case, case.name());
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
            if matches!(
                case,
                Case::XlsxSourceOpen
                    | Case::XlsxSourceListSheets
                    | Case::XlsxSourceFirstCell
                    | Case::XlsxSourceNarrowColumnRangeScan
            ) {
                let xlsx_source = measured.source.unwrap().xlsx.unwrap();
                assert_eq!(xlsx_source.workbook_read_calls.len(), 1);
                assert_eq!(xlsx_source.selected_worksheet_read_calls.len(), 1);
                assert_eq!(xlsx_source.unselected_worksheet_read_calls.len(), 1);
            }
        }

        let measured = run_case(Case::XlsxOneCellCommitFirstRead, &first, 0, 1).unwrap();
        assert_eq!(measured.case, "xlsx_one_cell_commit_first_read");
        assert_eq!(measured.elapsed_ns.samples.len(), 1);
    }

    #[test]
    fn xlsx_cell_values_matched_controls_are_deterministic_and_bounded() {
        let first = build_xlsx_cell_crud_corpus(XlsxCellCrudShape::Medium).unwrap();
        let second = build_xlsx_cell_crud_corpus(XlsxCellCrudShape::Medium).unwrap();
        assert_eq!(first.archive, second.archive);
        assert_eq!(
            first.manifest.generator,
            XLSX_CELL_VALUES_SOURCE_EDIT_CORPUS_GENERATOR
        );
        assert!(first.manifest.archive_member_count >= XLSX_CELL_VALUES_MEDIA_ENTRY_COUNT);
        let spec = first.xlsx.as_ref().unwrap();
        assert_eq!(spec.sheet_count, 4);
        assert_eq!(xlsx_cell_count(spec).unwrap(), 9_216);
        assert_eq!(spec.one_percent_updates.len(), 93);
        assert_eq!(XlsxCellCrudShape::ALL.len(), 2);

        for case in [
            Case::XlsxEagerCellValuesOneEditSave,
            Case::XlsxSourceBackedCellValuesOneEditSave,
            Case::XlsxEagerCellValuesOnePercentEditSave,
            Case::XlsxSourceBackedCellValuesOnePercentEditSave,
            Case::XlsxEagerCellValuesBatchEditSave,
            Case::XlsxSourceBackedCellValuesBatchEditSave,
        ] {
            let measured = run_case(case, &first, 0, 1).unwrap();
            assert_eq!(measured.case, case.name());
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
            assert!(measured.output_sha256.is_some());
            let sink = measured.sink.unwrap();
            assert!(sink.largest_write <= 64 * 1024, "{}", case.name());
            if case.is_xlsx_cell_values_edit_save() {
                assert_eq!(
                    measured.corpus.generator,
                    XLSX_CELL_VALUES_SOURCE_EDIT_CORPUS_GENERATOR
                );
            }
        }

        let dense_first = build_xlsx_cell_crud_corpus(XlsxCellCrudShape::DenseSparse).unwrap();
        let dense_second = build_xlsx_cell_crud_corpus(XlsxCellCrudShape::DenseSparse).unwrap();
        assert_eq!(dense_first.archive, dense_second.archive);
        let dense_spec = dense_first.xlsx.as_ref().unwrap();
        assert_eq!(dense_spec.row_count, 128);
        assert_eq!(dense_spec.column_count, 128);
        assert_eq!(xlsx_cell_count(dense_spec).unwrap(), 17_792);
        assert_eq!(dense_spec.one_percent_updates.len(), 178);
        assert!(dense_first.manifest.archive_member_count >= XLSX_CELL_VALUES_MEDIA_ENTRY_COUNT);
    }

    #[test]
    fn xlsx_merge_and_unmerge_cases_are_deterministic_and_reversible() {
        for case in [
            Case::XlsxEagerMergeCommitSave,
            Case::XlsxEagerUnmergeCommitSave,
        ] {
            let first = build_xlsx_merge_edit_corpus(case).unwrap();
            let second = build_xlsx_merge_edit_corpus(case).unwrap();
            assert_eq!(first.archive, second.archive, "{}", case.name());
            let measured = run_case(case, &first, 0, 1).unwrap();
            assert_eq!(measured.case, case.name());
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
            assert!(measured.output_sha256.is_some());
        }
    }

    #[test]
    fn xlsx_dense_wide_shape_preserves_the_narrow_range_contrast() {
        assert_eq!(XlsxShape::DenseWide.sheet_count(), 2);
        assert_eq!(XlsxShape::DenseWide.row_count(), 256);
        assert_eq!(XlsxShape::DenseWide.column_count(), 256);
        assert_eq!(
            XlsxShape::DenseWide.sheet_count()
                * XlsxShape::DenseWide.row_count()
                * XlsxShape::DenseWide.column_count(),
            131_072
        );

        let corpus = build_xlsx_corpus(XlsxShape::DenseWide).unwrap();
        assert_eq!(
            corpus.manifest.archive_sha256,
            "5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714"
        );
        assert_eq!(
            corpus
                .manifest
                .xlsx
                .as_ref()
                .unwrap()
                .one_percent_update_count,
            1_311
        );
        let measured = run_case(Case::XlsxNarrowColumnRangeScan, &corpus, 0, 1).unwrap();
        assert_eq!(measured.case, "xlsx_narrow_column_range_scan");
        assert_eq!(measured.elapsed_ns.samples.len(), 1);
    }

    #[test]
    fn fresh_writer_corpora_are_deterministic_and_identify_the_packaged_stream() {
        let cases = [
            (
                Case::DocFreshWriteTo,
                "DOC/CFB",
                "WordDocument",
                "ec7824ca46413dbdb6c96ee01abf2d49ffa702046d675c04518eebf0ab3e4e3b",
            ),
            (
                Case::XlsFreshWriteTo,
                "XLS/CFB",
                "Workbook",
                "cdc133bd87aaa60a91ea5e94df6ff8da0eb6bb0f2432afa4bfdb13cf70c0298b",
            ),
            (
                Case::PptFreshWriteTo,
                "PPT/CFB",
                "PowerPoint Document",
                "e233c6b63928578c2429178c3ac8589b32d73d44df9953ac18c9d27f6968d8b4",
            ),
        ];

        for (case, package_format, target_entry, archive_sha256) in cases {
            let first = build_writer_corpus(case, WriterShape::Tiny).unwrap();
            let second = build_writer_corpus(case, WriterShape::Tiny).unwrap();

            assert_eq!(first.archive, second.archive, "{}", case.name());
            assert_eq!(first.manifest.generator, "litchi-legacy-writer-v1");
            assert_eq!(first.manifest.package_format, package_format);
            assert_eq!(first.manifest.target_entry, target_entry);
            assert_eq!(first.manifest.archive_sha256, archive_sha256);
            assert!(!first.target_payload.is_empty());

            let measured = run_case(case, &first, 0, 1).unwrap();
            assert_eq!(measured.case, case.name());
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
        }
    }

    #[test]
    fn large_writer_corpora_are_bounded_and_deterministic() {
        for case in [
            Case::DocFreshWriteTo,
            Case::XlsFreshWriteTo,
            Case::PptFreshWriteTo,
        ] {
            let first = build_writer_corpus(case, WriterShape::Large).unwrap();
            let second = build_writer_corpus(case, WriterShape::Large).unwrap();

            assert_eq!(first.archive, second.archive, "{}", case.name());
            assert_eq!(first.manifest.shape, "large");
            assert!(first.manifest.entry_count > 100, "{}", case.name());
            assert!(
                first.manifest.archive_bytes < 16 * 1024 * 1024,
                "{}",
                case.name()
            );
        }
    }

    #[test]
    fn payload_heavy_writer_corpora_fill_the_primary_stream_deterministically() {
        for case in [
            Case::DocFreshWriteTo,
            Case::XlsFreshWriteTo,
            Case::PptFreshWriteTo,
        ] {
            let first = build_writer_corpus(case, WriterShape::PayloadHeavy).unwrap();
            let second = build_writer_corpus(case, WriterShape::PayloadHeavy).unwrap();

            assert_eq!(first.archive, second.archive, "{}", case.name());
            assert_eq!(first.manifest.shape, "payload-heavy");
            assert!(
                (4 * 1024 * 1024..=8 * 1024 * 1024).contains(&first.target_payload.len()),
                "{} primary stream is {} bytes",
                case.name(),
                first.target_payload.len()
            );
            let measured = run_case(case, &first, 0, 1).unwrap();
            assert_eq!(measured.elapsed_ns.samples.len(), 1);
        }
    }

    #[test]
    fn payload_families_are_distinct_and_repeatable() {
        let compressible = payload_bytes(PayloadKind::Compressible, 7, 4096);
        let incompressible = payload_bytes(PayloadKind::Incompressible, 7, 4096);

        assert_eq!(
            incompressible,
            payload_bytes(PayloadKind::Incompressible, 7, 4096)
        );
        assert_ne!(compressible, incompressible);
    }

    #[test]
    fn statistics_include_tails_dispersion_and_confidence_interval() {
        let measured = statistics(vec![5, 1, 4, 2, 3]);

        assert_eq!(measured.samples, vec![1, 2, 3, 4, 5]);
        assert_eq!(measured.p50, 3);
        assert_eq!(measured.p95, 5);
        assert_eq!(measured.p99, 5);
        assert_eq!(measured.mean, 3.0);
        assert!((measured.standard_deviation - 1.581_138_830_084_189_8).abs() < f64::EPSILON);
        assert!(measured.confidence_interval_95.lower < measured.mean);
        assert!(measured.confidence_interval_95.upper > measured.mean);
    }

    #[test]
    fn range_simulator_delay_arithmetic_is_exact() {
        let config = RangeSimulationConfig {
            fixed_latency_us: 10,
            request_overhead_us: 5,
            bandwidth_bytes_per_second: 1_000,
            max_physical_range_bytes: 4_096,
        };

        assert_eq!(
            simulated_request_delay(config, 1_000),
            Duration::from_micros(1_000_015)
        );
        assert_eq!(
            simulated_request_delay(config, 1),
            Duration::from_micros(1_015)
        );
    }

    #[test]
    fn range_simulator_chunks_and_reports_a_deterministic_distribution() {
        let bytes = (0..10_000).map(|index| index as u8).collect::<Vec<_>>();
        let backing = Arc::new(InstrumentedSource::new(bytes.clone(), Vec::new()));
        let config = RangeSimulationConfig {
            fixed_latency_us: 0,
            request_overhead_us: 0,
            bandwidth_bytes_per_second: 1_000_000_000,
            max_physical_range_bytes: 4_096,
        };
        let source = SimulatedRangeSource::new(backing, config);

        let mut output = vec![0; bytes.len()];
        assert_eq!(source.read_at(0, &mut output).unwrap(), bytes.len());
        assert_eq!(output, bytes);
        let first = source.snapshot().unwrap();
        assert_eq!(first.logical_read_calls, 1);
        assert_eq!(first.logical_read_bytes, 10_000);
        assert_eq!(first.physical_request_count, 3);
        assert_eq!(first.physical_request_bytes, 10_000);
        assert_eq!(first.physical_request_sizes, vec![1_808, 4_096, 4_096]);
        assert_eq!(
            first.physical_request_size_buckets,
            RequestSizeBuckets {
                bytes_513_to_4096: 3,
                ..RequestSizeBuckets::default()
            }
        );

        source.reset().unwrap();
        assert_eq!(source.read_at(0, &mut output).unwrap(), bytes.len());
        assert_eq!(source.snapshot().unwrap(), first);
    }

    #[test]
    fn range_source_cases_emit_physical_metrics_and_preserve_xlsx_deferral() {
        let config = RangeSimulationConfig {
            fixed_latency_us: 0,
            request_overhead_us: 0,
            bandwidth_bytes_per_second: 1_000_000_000,
            max_physical_range_bytes: 512,
        };
        let opc = build_opc_corpus(CorpusShape::Tiny, PayloadKind::Compressible).unwrap();
        for case in [Case::OpcRangeSourceOpen, Case::OpcRangeSourceOpenMainRead] {
            let measured = run_case_with_config(case, &opc, 0, 1, config).unwrap();
            let simulation = measured.source.unwrap().simulation.unwrap();
            assert_eq!(simulation.logical_read_calls.len(), 1);
            assert!(simulation.physical_request_count[0] > 0);
            assert!(
                simulation.physical_request_sizes[0]
                    .iter()
                    .all(|&bytes| bytes <= 512)
            );
        }

        let xlsx = build_xlsx_corpus(XlsxShape::Tiny).unwrap();
        for case in [
            Case::XlsxRangeSourceOpen,
            Case::XlsxRangeSourceListSheets,
            Case::XlsxRangeSourceFirstCell,
            Case::XlsxRangeSourceNarrowColumnRangeScan,
        ] {
            let measured = run_case_with_config(case, &xlsx, 0, 1, config).unwrap();
            let source = measured.source.unwrap();
            let simulation = source.simulation.unwrap();
            assert_eq!(simulation.logical_read_calls.len(), 1);
            assert_eq!(
                source.xlsx.unwrap().unselected_worksheet_read_calls,
                vec![0]
            );
            if case == Case::XlsxRangeSourceListSheets {
                assert_eq!(simulation.logical_read_calls, vec![0]);
                assert_eq!(simulation.physical_request_count, vec![0]);
            } else {
                assert!(simulation.physical_request_count[0] > 0);
            }
        }
    }

    #[test]
    fn execution_worker_selection_is_capped_deduplicated_and_sorted() {
        let selected = resolve_execution_workers(["8", "1", "available", "2", "2"]).unwrap();
        let available = std::thread::available_parallelism().map_or(1, std::num::NonZeroUsize::get);

        assert!(!selected.is_empty());
        assert!(selected.windows(2).all(|pair| pair[0] < pair[1]));
        assert!(selected.iter().all(|&workers| workers <= available));
        assert_eq!(
            selected,
            resolve_execution_workers(["8", "1", "available", "2", "2"]).unwrap()
        );
    }

    #[test]
    fn explicit_scaling_cases_record_exact_work_units() {
        let opc = build_opc_corpus(CorpusShape::Tiny, PayloadKind::Compressible).unwrap();
        let cfb = build_cfb_corpus(CorpusShape::Tiny, PayloadKind::Compressible).unwrap();
        let results = [
            run_scaling_case(Case::OpcOpenSessionScaling, &opc, 0, 1, 1).unwrap(),
            run_scaling_case(Case::CfbBulkReadScaling, &cfb, 0, 1, 1).unwrap(),
        ];

        assert_eq!(results[0].execution.unwrap().worker_count, 1);
        assert_eq!(
            results[0].execution.unwrap().logical_tasks,
            opc.manifest.archive_member_count
        );
        assert_eq!(results[1].execution.unwrap().worker_count, 1);
        assert_eq!(
            results[1].execution.unwrap().logical_tasks,
            cfb.manifest.entry_count
        );
        assert_eq!(
            results[1].execution.unwrap().logical_bytes,
            cfb.manifest.uncompressed_payload_bytes as u64
        );
    }

    #[test]
    fn counting_sink_rejects_writes_before_mutating_counters() {
        let mut sink = CountingSink::bounded(4, 3);

        assert!(sink.write(b"abcd").is_err());
        assert_eq!(sink.summary().accepted_bytes, 0);
        sink.write_all(b"abc").unwrap();
        assert!(sink.write_all(b"de").is_err());
        assert_eq!(sink.summary().accepted_bytes, 3);
    }
}
