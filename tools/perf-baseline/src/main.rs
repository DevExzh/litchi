//! Reproducible, content-free baseline measurements for Litchi's OPC and CFB substrates.
//!
//! This is deliberately a standalone tool rather than a public crate dependency.
//! It creates all inputs in memory from fixed specifications and writes JSON that
//! identifies the exact generated corpus by SHA-256.

#![forbid(unsafe_code)]

use std::{
    collections::BTreeSet,
    error::Error,
    fs::{self, File},
    io::{self, Cursor, Seek, SeekFrom, Write},
    num::{NonZeroU64, NonZeroUsize},
    ops::Range,
    path::PathBuf,
    process::Command,
    sync::{
        Arc, Barrier, Mutex,
        atomic::{AtomicU64, Ordering},
    },
    time::{Duration, Instant},
};

use litchi_cfb::{OleFile, OleWriter, SharedOleFile, SharedOleFileLimits};
use litchi_core::Position;
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, ReadAt, SourceVersion,
};
use litchi_opc::{
    BlobPart, OpcPackage, OpenSession, PackURI, PackageWriter, ReadLimits, SourceBackedPackage,
    SourceCacheLimits, TargetMode, constants::relationship_type,
};
use litchi_xlsx::{Cell as XlsxCell, SourceBackedWorkbook, Value as XlsxValue, Workbook};
use serde::Serialize;
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;

const SCHEMA_VERSION: u32 = 1;
const DEFAULT_SAMPLES: usize = 15;
const DEFAULT_WARMUP_ITERATIONS: usize = 3;
const DEFAULT_RANGE_FIXED_LATENCY_US: u64 = 100;
const DEFAULT_RANGE_REQUEST_OVERHEAD_US: u64 = 25;
const DEFAULT_RANGE_BANDWIDTH_BYTES_PER_SECOND: u64 = 50 * 1024 * 1024;
const DEFAULT_RANGE_MAX_PHYSICAL_BYTES: usize = 4 * 1024;
const CONTENT_TYPE: &str = "application/octet-stream";
const OPC_CORPUS_GENERATOR: &str = "litchi-opc-synthetic-v2";
const CFB_CORPUS_GENERATOR: &str = "litchi-cfb-synthetic-v1";
const LEGACY_WRITER_CORPUS_GENERATOR: &str = "litchi-legacy-writer-v1";
const XLSX_CORPUS_GENERATOR: &str = "litchi-xlsx-synthetic-v1";
const SEMANTIC_DOCX_CORPUS_GENERATOR: &str = "litchi-docx-semantic-v1";
const SEMANTIC_PPTX_CORPUS_GENERATOR: &str = "litchi-pptx-semantic-v1";
const SEMANTIC_ODT_CORPUS_GENERATOR: &str = "litchi-odt-semantic-v1";
const SEMANTIC_ODS_CORPUS_GENERATOR: &str = "litchi-ods-semantic-v1";
const SEMANTIC_ODP_CORPUS_GENERATOR: &str = "litchi-odp-semantic-v1";
const SEMANTIC_RTF_CORPUS_GENERATOR: &str = "litchi-rtf-semantic-v1";
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

/// Small, complete public-API DOCX/PPTX corpora.  These cases are opt-in so
/// their intentionally semantic workload does not alter the substrate matrix.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SemanticShape {
    Tiny,
    Medium,
    Large,
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
    CfbOpen,
    CfbListStreams,
    CfbReadOne,
    CfbCreateStreamBorrowed,
    CfbCreateStreamOwned,
    CfbSharedOpen,
    CfbSharedReadOne,
    CfbSharedConcurrentReads,
    DocFreshWriteTo,
    XlsFreshWriteTo,
    PptFreshWriteTo,
    DocSemanticOpen,
    DocSemanticListParagraphs,
    DocSemanticOneParagraph,
    DocSemanticFullText,
    DocSemanticNoopEditSave,
    DocSemanticOneEditSave,
    XlsSemanticOpen,
    XlsSemanticListWorksheets,
    XlsSemanticOneCell,
    XlsSemanticFullCellScan,
    XlsSemanticNoopEditSave,
    XlsSemanticOneEditSave,
    PptSemanticOpen,
    PptSemanticListSlides,
    PptSemanticOneShapeText,
    PptSemanticFullText,
    PptSlideOrderSnapshotOpen,
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
    OpcRangeSourceOpen,
    OpcRangeSourceOpenMainRead,
    XlsxRangeSourceOpen,
    XlsxRangeSourceListSheets,
    XlsxRangeSourceFirstCell,
    XlsxRangeSourceNarrowColumnRangeScan,
    OpcOpenSessionScaling,
    CfbBulkReadScaling,
    RtfSemanticOpen,
    RtfSemanticListParagraphs,
    RtfSemanticOneParagraph,
    RtfSemanticFullText,
    RtfSemanticStreamSave,
    RtfSemanticNoopEditSave,
    RtfSemanticOneEditSave,
    DocxSemanticOpen,
    DocxSemanticListParagraphs,
    DocxSemanticOneParagraph,
    DocxSemanticFullText,
    DocxSemanticCreateSmall,
    DocxSemanticNoopEditSave,
    DocxSemanticOneEditSave,
    DocxSemanticOnePercentEditSave,
    PptxSemanticOpen,
    PptxSemanticListSlides,
    PptxSemanticOneSlide,
    PptxSemanticFullText,
    PptxSemanticCreateSmall,
    PptxSemanticNoopEditSave,
    PptxSemanticOneEditSave,
    PptxSemanticOnePercentEditSave,
    OdtSemanticOpen,
    OdtSemanticListParagraphs,
    OdtSemanticOneParagraph,
    OdtSemanticFullText,
    OdtSemanticCreateSmall,
    OdtSemanticNoopEditSave,
    OdtSemanticOneEditSave,
    OdsSemanticOpen,
    OdsSemanticListSheets,
    OdsSemanticOneCell,
    OdsSemanticFullCellText,
    OdsSemanticCreateSmall,
    OdsSemanticNoopEditSave,
    OdsSemanticOneEditSave,
    OdpSemanticOpen,
    OdpSemanticListSlides,
    OdpSemanticOneSlide,
    OdpSemanticFullText,
    OdpSemanticCreateSmall,
    OdpSemanticNoopEditSave,
    OdpSemanticOneEditSave,
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
            Self::CfbOpen => "cfb_open",
            Self::CfbListStreams => "cfb_list_streams",
            Self::CfbReadOne => "cfb_read_one",
            Self::CfbCreateStreamBorrowed => "cfb_create_stream_borrowed",
            Self::CfbCreateStreamOwned => "cfb_create_stream_owned",
            Self::CfbSharedOpen => "cfb_shared_open",
            Self::CfbSharedReadOne => "cfb_shared_read_one",
            Self::CfbSharedConcurrentReads => "cfb_shared_concurrent_reads",
            Self::DocFreshWriteTo => "doc_fresh_write_to",
            Self::XlsFreshWriteTo => "xls_fresh_write_to",
            Self::PptFreshWriteTo => "ppt_fresh_write_to",
            Self::DocSemanticOpen => "doc_semantic_open",
            Self::DocSemanticListParagraphs => "doc_semantic_list_paragraphs",
            Self::DocSemanticOneParagraph => "doc_semantic_one_paragraph",
            Self::DocSemanticFullText => "doc_semantic_full_text",
            Self::DocSemanticNoopEditSave => "doc_semantic_noop_edit_save",
            Self::DocSemanticOneEditSave => "doc_semantic_one_edit_save",
            Self::XlsSemanticOpen => "xls_semantic_open",
            Self::XlsSemanticListWorksheets => "xls_semantic_list_worksheets",
            Self::XlsSemanticOneCell => "xls_semantic_one_cell",
            Self::XlsSemanticFullCellScan => "xls_semantic_full_cell_scan",
            Self::XlsSemanticNoopEditSave => "xls_semantic_noop_edit_save",
            Self::XlsSemanticOneEditSave => "xls_semantic_one_edit_save",
            Self::PptSemanticOpen => "ppt_semantic_open",
            Self::PptSemanticListSlides => "ppt_semantic_list_slides",
            Self::PptSemanticOneShapeText => "ppt_semantic_one_shape_text",
            Self::PptSemanticFullText => "ppt_semantic_full_text",
            Self::PptSlideOrderSnapshotOpen => "ppt_slide_order_snapshot_open",
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
            Self::RtfSemanticListParagraphs => "rtf_semantic_list_paragraphs",
            Self::RtfSemanticOneParagraph => "rtf_semantic_one_paragraph",
            Self::RtfSemanticFullText => "rtf_semantic_full_text",
            Self::RtfSemanticStreamSave => "rtf_semantic_stream_save",
            Self::RtfSemanticNoopEditSave => "rtf_semantic_noop_edit_save",
            Self::RtfSemanticOneEditSave => "rtf_semantic_one_edit_save",
            Self::DocxSemanticOpen => "docx_semantic_open",
            Self::DocxSemanticListParagraphs => "docx_semantic_list_paragraphs",
            Self::DocxSemanticOneParagraph => "docx_semantic_one_paragraph",
            Self::DocxSemanticFullText => "docx_semantic_full_text",
            Self::DocxSemanticCreateSmall => "docx_semantic_create_small",
            Self::DocxSemanticNoopEditSave => "docx_semantic_noop_edit_save",
            Self::DocxSemanticOneEditSave => "docx_semantic_one_edit_save",
            Self::DocxSemanticOnePercentEditSave => "docx_semantic_one_percent_edit_save",
            Self::PptxSemanticOpen => "pptx_semantic_open",
            Self::PptxSemanticListSlides => "pptx_semantic_list_slides",
            Self::PptxSemanticOneSlide => "pptx_semantic_one_slide",
            Self::PptxSemanticFullText => "pptx_semantic_full_text",
            Self::PptxSemanticCreateSmall => "pptx_semantic_create_small",
            Self::PptxSemanticNoopEditSave => "pptx_semantic_noop_edit_save",
            Self::PptxSemanticOneEditSave => "pptx_semantic_one_edit_save",
            Self::PptxSemanticOnePercentEditSave => "pptx_semantic_one_percent_edit_save",
            Self::OdtSemanticOpen => "odt_semantic_open",
            Self::OdtSemanticListParagraphs => "odt_semantic_list_paragraphs",
            Self::OdtSemanticOneParagraph => "odt_semantic_one_paragraph",
            Self::OdtSemanticFullText => "odt_semantic_full_text",
            Self::OdtSemanticCreateSmall => "odt_semantic_create_small",
            Self::OdtSemanticNoopEditSave => "odt_semantic_noop_edit_save",
            Self::OdtSemanticOneEditSave => "odt_semantic_one_edit_save",
            Self::OdsSemanticOpen => "ods_semantic_open",
            Self::OdsSemanticListSheets => "ods_semantic_list_sheets",
            Self::OdsSemanticOneCell => "ods_semantic_one_cell",
            Self::OdsSemanticFullCellText => "ods_semantic_full_cell_text",
            Self::OdsSemanticCreateSmall => "ods_semantic_create_small",
            Self::OdsSemanticNoopEditSave => "ods_semantic_noop_edit_save",
            Self::OdsSemanticOneEditSave => "ods_semantic_one_edit_save",
            Self::OdpSemanticOpen => "odp_semantic_open",
            Self::OdpSemanticListSlides => "odp_semantic_list_slides",
            Self::OdpSemanticOneSlide => "odp_semantic_one_slide",
            Self::OdpSemanticFullText => "odp_semantic_full_text",
            Self::OdpSemanticCreateSmall => "odp_semantic_create_small",
            Self::OdpSemanticNoopEditSave => "odp_semantic_noop_edit_save",
            Self::OdpSemanticOneEditSave => "odp_semantic_one_edit_save",
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
                | Self::CfbSharedOpen
                | Self::CfbSharedReadOne
                | Self::CfbSharedConcurrentReads
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

    const fn uses_semantic_ppt(self) -> bool {
        matches!(
            self,
            Self::PptSemanticOpen
                | Self::PptSemanticListSlides
                | Self::PptSemanticOneShapeText
                | Self::PptSemanticFullText
                | Self::PptSlideOrderSnapshotOpen
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
                | Self::RtfSemanticListParagraphs
                | Self::RtfSemanticOneParagraph
                | Self::RtfSemanticFullText
                | Self::RtfSemanticStreamSave
                | Self::RtfSemanticNoopEditSave
                | Self::RtfSemanticOneEditSave
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
        )
    }

    const fn uses_semantic_ods(self) -> bool {
        matches!(
            self,
            Self::OdsSemanticOpen
                | Self::OdsSemanticListSheets
                | Self::OdsSemanticOneCell
                | Self::OdsSemanticFullCellText
                | Self::OdsSemanticCreateSmall
                | Self::OdsSemanticNoopEditSave
                | Self::OdsSemanticOneEditSave
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
    cases: Vec<Case>,
    shapes: Vec<CorpusShape>,
    payloads: Vec<PayloadKind>,
    writer_shapes: Vec<WriterShape>,
    xlsx_shapes: Vec<XlsxShape>,
    semantic_shapes: Vec<SemanticShape>,
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
struct XlsxCorpus {
    sheet_count: usize,
    row_count: usize,
    column_count: usize,
    one_percent_updates: Vec<XlsxCoordinate>,
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
}

#[derive(Serialize)]
struct Configuration {
    samples_per_case: usize,
    warmup_iterations_per_case: usize,
    cases: Vec<&'static str>,
    corpus_shapes: Vec<&'static str>,
    payload_kinds: Vec<&'static str>,
    writer_shapes: Vec<&'static str>,
    xlsx_shapes: Vec<&'static str>,
    semantic_shapes: Vec<&'static str>,
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
    corpus: CorpusManifest,
    elapsed_ns: Statistics,
    sink: Option<SinkSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    source: Option<SourceSummary>,
    #[serde(skip_serializing_if = "Option::is_none")]
    execution: Option<ExecutionSummary>,
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
    simulation: Option<RangeSimulationSummary>,
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

fn main() -> Result<(), Box<dyn Error>> {
    let options = parse_options()?;
    let mut results = Vec::new();

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
                    && !case.uses_semantic_rtf()
                    && !case.uses_semantic_docx()
                    && !case.uses_semantic_pptx()
                    && !case.uses_semantic_odt()
                    && !case.uses_semantic_ods()
                    && !case.uses_semantic_odp()
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
        for shape in &options.semantic_shapes {
            let corpus = build_semantic_rtf_corpus(*shape)?;
            for case in options.cases.iter().filter(|case| case.uses_semantic_rtf()) {
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
        environment: environment(),
        configuration: Configuration {
            samples_per_case: options.samples,
            warmup_iterations_per_case: options.warmup_iterations,
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
            semantic_shapes: options
                .semantic_shapes
                .iter()
                .map(|shape| shape.name())
                .collect(),
            range_simulation: options.range_simulation,
            execution_workers: options.execution_workers,
        },
        results,
    };

    write_report(&report, options.output.as_ref())
}

fn parse_options() -> Result<Options, Box<dyn Error>> {
    let mut samples = DEFAULT_SAMPLES;
    let mut warmup_iterations = DEFAULT_WARMUP_ITERATIONS;
    let mut cases = Case::DEFAULT.to_vec();
    let mut shapes = CorpusShape::ALL.to_vec();
    let mut payloads = PayloadKind::ALL.to_vec();
    let mut writer_shapes = WriterShape::ALL.to_vec();
    let mut xlsx_shapes = XlsxShape::ALL.to_vec();
    let mut semantic_shapes = SemanticShape::ALL.to_vec();
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
            "--semantic-shape" => {
                semantic_shapes =
                    parse_selection(arguments.next(), "--semantic-shape", parse_semantic_shape)?;
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
        cases,
        shapes,
        payloads,
        writer_shapes,
        xlsx_shapes,
        semantic_shapes,
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
        "cfb_open" => Some(Case::CfbOpen),
        "cfb_list_streams" => Some(Case::CfbListStreams),
        "cfb_read_one" => Some(Case::CfbReadOne),
        "cfb_create_stream_borrowed" => Some(Case::CfbCreateStreamBorrowed),
        "cfb_create_stream_owned" => Some(Case::CfbCreateStreamOwned),
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
        "xls_semantic_open" => Some(Case::XlsSemanticOpen),
        "xls_semantic_list_worksheets" => Some(Case::XlsSemanticListWorksheets),
        "xls_semantic_one_cell" => Some(Case::XlsSemanticOneCell),
        "xls_semantic_full_cell_scan" => Some(Case::XlsSemanticFullCellScan),
        "xls_semantic_noop_edit_save" => Some(Case::XlsSemanticNoopEditSave),
        "xls_semantic_one_edit_save" => Some(Case::XlsSemanticOneEditSave),
        "ppt_semantic_open" => Some(Case::PptSemanticOpen),
        "ppt_semantic_list_slides" => Some(Case::PptSemanticListSlides),
        "ppt_semantic_one_shape_text" => Some(Case::PptSemanticOneShapeText),
        "ppt_semantic_full_text" => Some(Case::PptSemanticFullText),
        "ppt_slide_order_snapshot_open" => Some(Case::PptSlideOrderSnapshotOpen),
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
        "rtf_semantic_list_paragraphs" => Some(Case::RtfSemanticListParagraphs),
        "rtf_semantic_one_paragraph" => Some(Case::RtfSemanticOneParagraph),
        "rtf_semantic_full_text" => Some(Case::RtfSemanticFullText),
        "rtf_semantic_stream_save" => Some(Case::RtfSemanticStreamSave),
        "rtf_semantic_noop_edit_save" => Some(Case::RtfSemanticNoopEditSave),
        "rtf_semantic_one_edit_save" => Some(Case::RtfSemanticOneEditSave),
        "docx_semantic_open" => Some(Case::DocxSemanticOpen),
        "docx_semantic_list_paragraphs" => Some(Case::DocxSemanticListParagraphs),
        "docx_semantic_one_paragraph" => Some(Case::DocxSemanticOneParagraph),
        "docx_semantic_full_text" => Some(Case::DocxSemanticFullText),
        "docx_semantic_create_small" => Some(Case::DocxSemanticCreateSmall),
        "docx_semantic_noop_edit_save" => Some(Case::DocxSemanticNoopEditSave),
        "docx_semantic_one_edit_save" => Some(Case::DocxSemanticOneEditSave),
        "docx_semantic_one_percent_edit_save" => Some(Case::DocxSemanticOnePercentEditSave),
        "pptx_semantic_open" => Some(Case::PptxSemanticOpen),
        "pptx_semantic_list_slides" => Some(Case::PptxSemanticListSlides),
        "pptx_semantic_one_slide" => Some(Case::PptxSemanticOneSlide),
        "pptx_semantic_full_text" => Some(Case::PptxSemanticFullText),
        "pptx_semantic_create_small" => Some(Case::PptxSemanticCreateSmall),
        "pptx_semantic_noop_edit_save" => Some(Case::PptxSemanticNoopEditSave),
        "pptx_semantic_one_edit_save" => Some(Case::PptxSemanticOneEditSave),
        "pptx_semantic_one_percent_edit_save" => Some(Case::PptxSemanticOnePercentEditSave),
        "odt_semantic_open" => Some(Case::OdtSemanticOpen),
        "odt_semantic_list_paragraphs" => Some(Case::OdtSemanticListParagraphs),
        "odt_semantic_one_paragraph" => Some(Case::OdtSemanticOneParagraph),
        "odt_semantic_full_text" => Some(Case::OdtSemanticFullText),
        "odt_semantic_create_small" => Some(Case::OdtSemanticCreateSmall),
        "odt_semantic_noop_edit_save" => Some(Case::OdtSemanticNoopEditSave),
        "odt_semantic_one_edit_save" => Some(Case::OdtSemanticOneEditSave),
        "ods_semantic_open" => Some(Case::OdsSemanticOpen),
        "ods_semantic_list_sheets" => Some(Case::OdsSemanticListSheets),
        "ods_semantic_one_cell" => Some(Case::OdsSemanticOneCell),
        "ods_semantic_full_cell_text" => Some(Case::OdsSemanticFullCellText),
        "ods_semantic_create_small" => Some(Case::OdsSemanticCreateSmall),
        "ods_semantic_noop_edit_save" => Some(Case::OdsSemanticNoopEditSave),
        "ods_semantic_one_edit_save" => Some(Case::OdsSemanticOneEditSave),
        "odp_semantic_open" => Some(Case::OdpSemanticOpen),
        "odp_semantic_list_slides" => Some(Case::OdpSemanticListSlides),
        "odp_semantic_one_slide" => Some(Case::OdpSemanticOneSlide),
        "odp_semantic_full_text" => Some(Case::OdpSemanticFullText),
        "odp_semantic_create_small" => Some(Case::OdpSemanticCreateSmall),
        "odp_semantic_noop_edit_save" => Some(Case::OdpSemanticNoopEditSave),
        "odp_semantic_one_edit_save" => Some(Case::OdpSemanticOneEditSave),
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

fn parse_semantic_shape(value: &str) -> Option<SemanticShape> {
    match value {
        "tiny" => Some(SemanticShape::Tiny),
        "medium" => Some(SemanticShape::Medium),
        "large" => Some(SemanticShape::Large),
        _ => None,
    }
}

fn print_usage() {
    println!(
        "Usage: cargo run --release --manifest-path tools/perf-baseline/Cargo.toml -- [OPTIONS]\n\n\
         Options:\n\
           --samples N                 Samples per case (default: {DEFAULT_SAMPLES})\n\
           --warmup N                  Untimed iterations per case (default: {DEFAULT_WARMUP_ITERATIONS})\n\
           --case LIST                 zip_index,zip_read_one,opc_open,opc_open_owned,\n\
                                       opc_noop_save,opc_mutated_save,opc_source_open,\n\
                                       opc_source_open_main_read,opc_source_cached_main_read,\n\
                                       opc_source_concurrent_same_part,\n\
                                       cfb_open,cfb_list_streams,cfb_read_one,\n\
                                       cfb_create_stream_borrowed,cfb_create_stream_owned,\n\
                                       cfb_shared_open,cfb_shared_read_one,\n\
                                       cfb_shared_concurrent_reads,\n\
                                       doc_fresh_write_to,xls_fresh_write_to,ppt_fresh_write_to,\n\
                                       doc_semantic_open,doc_semantic_list_paragraphs,\n\
                                       doc_semantic_one_paragraph,doc_semantic_full_text,\n\
                                       doc_semantic_noop_edit_save,doc_semantic_one_edit_save,\n\
                                       xls_semantic_open,xls_semantic_list_worksheets,\n\
                                       xls_semantic_one_cell,xls_semantic_full_cell_scan,\n\
                                       xls_semantic_noop_edit_save,xls_semantic_one_edit_save,\n\
                                       ppt_semantic_open,ppt_semantic_list_slides,\n\
                                       ppt_semantic_one_shape_text,ppt_semantic_full_text,\n\
                                       ppt_slide_order_snapshot_open,\n\
                                       ppt_semantic_noop_edit_save,ppt_semantic_one_edit_save,\n\
                                       xlsx_open_owned,xlsx_list_sheets,xlsx_first_cell,\n\
                                       xlsx_full_cell_scan,xlsx_narrow_column_range_scan,\n\
                                       xlsx_noop_commit,xlsx_noop_commit_save,\n\
                                       xlsx_one_cell_commit,xlsx_one_cell_commit_first_read,\n\
                                       xlsx_one_cell_commit_save,\n\
                                       xlsx_one_percent_commit,xlsx_one_percent_commit_save,\n\
                                       xlsx_source_open,xlsx_source_list_sheets,\n\
                                       xlsx_source_first_cell,\n\
                                       xlsx_source_narrow_column_range_scan,\n\
                                       opc_range_source_open,opc_range_source_open_main_read,\n\
                                       xlsx_range_source_open,xlsx_range_source_list_sheets,\n\
                                       xlsx_range_source_first_cell,\n\
                                       xlsx_range_source_narrow_column_range_scan,\n\
                                       opc_open_session_scaling,cfb_bulk_read_scaling,\n\
                                       rtf_semantic_open,rtf_semantic_list_paragraphs,\n\
                                       rtf_semantic_one_paragraph,rtf_semantic_full_text,\n\
                                       rtf_semantic_stream_save,rtf_semantic_noop_edit_save,\n\
                                       rtf_semantic_one_edit_save,\n\
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
                                       odt_semantic_one_edit_save,ods_semantic_open,\n\
                                       ods_semantic_list_sheets,ods_semantic_one_cell,\n\
                                       ods_semantic_full_cell_text,ods_semantic_create_small,\n\
                                       ods_semantic_noop_edit_save,ods_semantic_one_edit_save,\n\
                                       odp_semantic_open,odp_semantic_list_slides,\n\
                                       odp_semantic_one_slide,odp_semantic_full_text,\n\
                                       odp_semantic_create_small,odp_semantic_noop_edit_save,\n\
                                       odp_semantic_one_edit_save\n\
           --shape LIST                tiny,many-small,few-large,wide-root\n\
           --payload LIST              compressible,incompressible\n\
           --writer-shape LIST         tiny,large,payload-heavy\n\
           --xlsx-shape LIST           tiny,medium,dense-wide\n\
           --semantic-shape LIST       tiny,medium,large (only used by opt-in Office semantic cases)\n\
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
            xlsx: None,
        },
        archive,
        target_name,
        target_payload,
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

fn semantic_rtf_expected_text(shape: SemanticShape, updated: Option<usize>) -> String {
    (0..shape.rtf_paragraphs())
        .map(|index| semantic_rtf_text(index, updated == Some(index)))
        .collect::<Vec<_>>()
        .join("\n")
}

fn semantic_rtf_bytes(shape: SemanticShape) -> Result<Vec<u8>, Box<dyn Error>> {
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

fn build_semantic_rtf_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let archive = semantic_rtf_bytes(shape)?;
    let document = litchi_rtf::Document::from_bytes(&archive)?;
    verify_semantic_rtf(&document, shape, None)?;
    let target_payload = semantic_rtf_text(0, false).into_bytes();
    let content_bytes = semantic_rtf_expected_text(shape, None).len();
    Ok(Corpus {
        manifest: CorpusManifest {
            name: format!("rtf-semantic-{}", shape.name()),
            generator: SEMANTIC_RTF_CORPUS_GENERATOR,
            package_format: "RTF",
            shape: shape.name(),
            payload_kind: "deterministic-semantic-text",
            compression: "none",
            entry_count: shape.rtf_paragraphs(),
            archive_member_count: 1,
            entry_bytes: semantic_rtf_text(0, false).len(),
            uncompressed_payload_bytes: content_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: "paragraph:0".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            xlsx: None,
        },
        archive,
        target_name: "paragraph:0".to_owned(),
        target_payload,
        xlsx: None,
    })
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
            xlsx: None,
        },
        archive,
        target_name: "paragraph:0".to_owned(),
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

fn build_semantic_odt_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let archive = semantic_odt_bytes(shape)?;
    let document = litchi_odt::Document::from_bytes(archive.clone())?;
    verify_semantic_odt(&document, shape, false)?;
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
            xlsx: None,
        },
        archive,
        target_name: "paragraph:0".to_owned(),
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
            xlsx: None,
        },
        archive,
        target_name: "Sheet 0!R0C0".to_owned(),
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
            xlsx: None,
        },
        archive,
        target_name: "slide:0".to_owned(),
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
        let observed = sheet.cells("A1:XFD1048576")?.count();
        let expected = spec
            .row_count
            .checked_mul(spec.column_count)
            .ok_or("XLSX per-sheet cell count overflows usize")?;
        if observed != expected {
            return Err("XLSX stored cell count differs from corpus specification".into());
        }
        for row in 0..spec.row_count {
            for column in 0..spec.column_count {
                let coordinate = XlsxCoordinate {
                    sheet: sheet_index,
                    row,
                    column,
                };
                let expected = xlsx_value(coordinate) + i32::from(updated.contains(&coordinate));
                let address = xlsx_address(row, column)?;
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
        Case::CfbOpen => run_cfb_open(corpus, warmup_iterations, samples),
        Case::CfbListStreams => run_cfb_list_streams(corpus, warmup_iterations, samples),
        Case::CfbReadOne => run_cfb_read_one(corpus, warmup_iterations, samples),
        Case::CfbCreateStreamBorrowed => {
            run_cfb_create_stream(corpus, warmup_iterations, samples, false)
        },
        Case::CfbCreateStreamOwned => {
            run_cfb_create_stream(corpus, warmup_iterations, samples, true)
        },
        Case::CfbSharedOpen => run_cfb_shared_open(corpus, warmup_iterations, samples),
        Case::CfbSharedReadOne => run_cfb_shared_read_one(corpus, warmup_iterations, samples),
        Case::CfbSharedConcurrentReads => {
            run_cfb_shared_concurrent_reads(corpus, warmup_iterations, samples)
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
        Case::XlsSemanticOpen
        | Case::XlsSemanticListWorksheets
        | Case::XlsSemanticOneCell
        | Case::XlsSemanticFullCellScan
        | Case::XlsSemanticNoopEditSave
        | Case::XlsSemanticOneEditSave => {
            run_semantic_xls(case, corpus, warmup_iterations, samples)
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
        Case::XlsxSourceOpen => run_xlsx_source_open(corpus, warmup_iterations, samples),
        Case::XlsxSourceListSheets => {
            run_xlsx_source_list_sheets(corpus, warmup_iterations, samples)
        },
        Case::XlsxSourceFirstCell => run_xlsx_source_first_cell(corpus, warmup_iterations, samples),
        Case::XlsxSourceNarrowColumnRangeScan => {
            run_xlsx_source_narrow_column_range_scan(corpus, warmup_iterations, samples)
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
        | Case::RtfSemanticListParagraphs
        | Case::RtfSemanticOneParagraph
        | Case::RtfSemanticFullText
        | Case::RtfSemanticStreamSave
        | Case::RtfSemanticNoopEditSave
        | Case::RtfSemanticOneEditSave => {
            run_semantic_rtf(case, corpus, warmup_iterations, samples)
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
        Case::OdtSemanticOpen
        | Case::OdtSemanticListParagraphs
        | Case::OdtSemanticOneParagraph
        | Case::OdtSemanticFullText
        | Case::OdtSemanticCreateSmall
        | Case::OdtSemanticNoopEditSave
        | Case::OdtSemanticOneEditSave => {
            run_semantic_odt(case, corpus, warmup_iterations, samples)
        },
        Case::OdsSemanticOpen
        | Case::OdsSemanticListSheets
        | Case::OdsSemanticOneCell
        | Case::OdsSemanticFullCellText
        | Case::OdsSemanticCreateSmall
        | Case::OdsSemanticNoopEditSave
        | Case::OdsSemanticOneEditSave => {
            run_semantic_ods(case, corpus, warmup_iterations, samples)
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
    updated: Option<usize>,
) -> Result<(), Box<dyn Error>> {
    if document.paragraph_count() != shape.rtf_paragraphs() {
        return Err("semantic RTF paragraph count differs from specification".into());
    }
    let mut count = 0usize;
    for (index, paragraph) in document.body().paragraphs().enumerate() {
        if paragraph.to_text() != semantic_rtf_text(index, updated == Some(index)) {
            return Err("semantic RTF paragraph text differs from specification".into());
        }
        count = count
            .checked_add(1)
            .ok_or("semantic RTF paragraph count overflows usize")?;
    }
    if count != shape.rtf_paragraphs() {
        return Err("semantic RTF paragraph traversal differs from specification".into());
    }
    if document.text() != semantic_rtf_expected_text(shape, updated) {
        return Err("semantic RTF full text differs from specification".into());
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

fn verify_semantic_odt(
    document: &litchi_odt::Document,
    shape: SemanticShape,
    updated: bool,
) -> Result<(), Box<dyn Error>> {
    let paragraphs = document.paragraphs()?;
    if paragraphs.len() != shape.docx_paragraphs() {
        return Err("semantic ODT paragraph count differs from specification".into());
    }
    for (index, paragraph) in paragraphs.iter().enumerate() {
        let is_updated = updated && index == shape.docx_paragraphs() / 2;
        if paragraph.text()? != semantic_odt_text(index, is_updated) {
            return Err("semantic ODT paragraph text differs from specification".into());
        }
    }
    let expected = (0..shape.docx_paragraphs())
        .map(|index| semantic_odt_text(index, updated && index == shape.docx_paragraphs() / 2))
        .collect::<Vec<_>>()
        .join("\n");
    if document.text()? != expected {
        return Err("semantic ODT full text differs from paragraph scan".into());
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

fn expected_semantic_ods_full_cell_text(shape: SemanticShape, updated: bool) -> String {
    let middle_sheet = shape.ods_sheet_count() / 2;
    let middle_row = shape.ods_rows_per_sheet() / 2;
    let middle_column = shape.ods_columns_per_sheet() / 2;
    let mut values = Vec::with_capacity(shape.ods_cell_count());
    for sheet in 0..shape.ods_sheet_count() {
        for row in 0..shape.ods_rows_per_sheet() {
            for column in 0..shape.ods_columns_per_sheet() {
                values.push(semantic_ods_text(
                    sheet,
                    row,
                    column,
                    updated
                        && sheet == middle_sheet
                        && row == middle_row
                        && column == middle_column,
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
    if spreadsheet.sheets().len() != shape.ods_sheet_count() {
        return Err("semantic ODS sheet count differs from specification".into());
    }
    let middle_sheet = shape.ods_sheet_count() / 2;
    let middle_row = shape.ods_rows_per_sheet() / 2;
    let middle_column = shape.ods_columns_per_sheet() / 2;
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
                let is_updated = updated
                    && sheet == middle_sheet
                    && row == middle_row
                    && column == middle_column;
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
        != expected_semantic_ods_full_cell_text(shape, updated)
    {
        return Err("semantic ODS full cell text differs from specification".into());
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

fn run_semantic_rtf(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let shape = semantic_shape(corpus)?;
    let selected = shape.rtf_paragraphs() / 2;
    let expected_changed = if case == Case::RtfSemanticOneEditSave {
        let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
        let mut edit = document.edit();
        edit.replace_paragraph_text(selected, semantic_rtf_text(selected, true))?;
        edit.commit()?.snapshot().to_bytes()?
    } else {
        corpus.archive.clone()
    };
    let sink_ceiling = u64::try_from(expected_changed.len())?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut sinks = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        match case {
            Case::RtfSemanticOpen => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let document = litchi_rtf::Document::from_bytes(&owned)?;
                let duration = started.elapsed();
                verify_semantic_rtf(&document, shape, None)?;
                std::hint::black_box(document);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::RtfSemanticListParagraphs => {
                let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
                let started = Instant::now();
                let count = document.body().paragraphs().count();
                let duration = started.elapsed();
                if count != shape.rtf_paragraphs() {
                    return Err("semantic RTF paragraph list differs from specification".into());
                }
                verify_semantic_rtf(&document, shape, None)?;
                std::hint::black_box(count);
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
                if paragraph != semantic_rtf_text(selected, false) {
                    return Err("semantic RTF selected paragraph differs from specification".into());
                }
                verify_semantic_rtf(&document, shape, None)?;
                std::hint::black_box(paragraph);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::RtfSemanticFullText => {
                let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
                let started = Instant::now();
                let text = document.text();
                let duration = started.elapsed();
                if text != semantic_rtf_expected_text(shape, None) {
                    return Err("semantic RTF full text differs from specification".into());
                }
                verify_semantic_rtf(&document, shape, None)?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::RtfSemanticStreamSave
            | Case::RtfSemanticNoopEditSave
            | Case::RtfSemanticOneEditSave => {
                let document = litchi_rtf::Document::from_bytes(&corpus.archive)?;
                let mut sink = CountingSink::bounded(sink_ceiling, sink_ceiling);
                sink.reserve_budget()?;
                let started = Instant::now();
                let (published, expected_update) = match case {
                    Case::RtfSemanticStreamSave => (document.clone(), None),
                    Case::RtfSemanticNoopEditSave => {
                        let commit = document.edit().commit()?;
                        if commit.diagnostics().changed()
                            || commit.diagnostics().operation_count() != 0
                            || !commit.snapshot().same_snapshot(&document)
                        {
                            return Err("semantic RTF no-op commit changed its source".into());
                        }
                        (commit.into_snapshot(), None)
                    },
                    Case::RtfSemanticOneEditSave => {
                        let mut edit = document.edit();
                        edit.replace_paragraph_text(selected, semantic_rtf_text(selected, true))?;
                        let commit = edit.commit()?;
                        if !commit.diagnostics().changed()
                            || commit.diagnostics().operation_count() != 1
                        {
                            return Err(
                                "semantic RTF changed commit has unexpected diagnostics".into()
                            );
                        }
                        (commit.into_snapshot(), Some(selected))
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
                verify_semantic_rtf(&reopened, shape, expected_update)?;
                if iteration >= warmup_iterations {
                    sinks.push(summary);
                }
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-RTF semantic case passed to RTF runner".into()),
        }
    }
    let sink = (!sinks.is_empty())
        .then(|| deterministic_sink_summary(&sinks, "semantic RTF stream/save"))
        .transpose()?;
    Ok(result(case, corpus, elapsed, sink))
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
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        match case {
            Case::OdtSemanticCreateSmall => {
                let started = Instant::now();
                let bytes = semantic_odt_bytes(SemanticShape::Tiny)?;
                let duration = started.elapsed();
                let reopened = litchi_odt::Document::from_bytes(bytes.clone())?;
                verify_semantic_odt(&reopened, SemanticShape::Tiny, false)?;
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
                verify_semantic_odt(&document, shape, false)?;
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
                verify_semantic_odt(&document, shape, false)?;
                std::hint::black_box(paragraphs);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdtSemanticOneParagraph => {
                let document = litchi_odt::Document::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let paragraphs = document.paragraphs()?;
                let text = paragraphs
                    .get(index)
                    .ok_or("semantic ODT selected paragraph is missing")?
                    .text()?;
                let duration = started.elapsed();
                if text != semantic_odt_text(index, false) {
                    return Err("semantic ODT selected paragraph differs from specification".into());
                }
                verify_semantic_odt(&document, shape, false)?;
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdtSemanticFullText => {
                let document = litchi_odt::Document::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let text = document.text()?;
                let duration = started.elapsed();
                verify_semantic_odt(&document, shape, false)?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdtSemanticNoopEditSave | Case::OdtSemanticOneEditSave => {
                let document = litchi_odt::Document::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let mut edit = document.edit()?;
                let updated = matches!(case, Case::OdtSemanticOneEditSave);
                if updated {
                    edit.replace_paragraph(Position::new(index), semantic_odt_text(index, true))?;
                }
                let commit = edit.commit()?;
                let bytes = commit.snapshot().as_bytes().to_vec();
                let duration = started.elapsed();
                if (bytes != corpus.archive) != updated {
                    return Err(
                        "semantic ODT edit/save changed-state differs from specification".into(),
                    );
                }
                let reopened = litchi_odt::Document::from_bytes(bytes.clone())?;
                verify_semantic_odt(&reopened, shape, updated)?;
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-ODT semantic case passed to ODT runner".into()),
        }
    }
    Ok(result(case, corpus, elapsed, None))
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
            Case::OdsSemanticFullCellText => {
                let spreadsheet = litchi_ods::Spreadsheet::from_bytes(corpus.archive.clone())?;
                let started = Instant::now();
                let text = semantic_ods_full_cell_text(&spreadsheet, shape)?;
                let duration = started.elapsed();
                verify_semantic_ods(&spreadsheet, shape, false)?;
                std::hint::black_box(text);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            Case::OdsSemanticNoopEditSave | Case::OdsSemanticOneEditSave => {
                let owned = corpus.archive.clone();
                let started = Instant::now();
                let snapshot = litchi_ods::document::Snapshot::from_bytes(owned)?;
                let mut edit = snapshot.edit();
                let updated = matches!(case, Case::OdsSemanticOneEditSave);
                if updated {
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
                }
                let commit = edit.commit()?;
                let bytes = commit.snapshot().as_bytes().to_vec();
                let duration = started.elapsed();
                if (bytes != corpus.archive) != updated || commit.changed() != updated {
                    return Err(
                        "semantic ODS edit/save changed-state differs from specification".into(),
                    );
                }
                let reopened = litchi_ods::Spreadsheet::from_bytes(bytes.clone())?;
                verify_semantic_ods(&reopened, shape, updated)?;
                std::hint::black_box(bytes);
                record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
            },
            _ => return Err("non-ODS semantic case passed to ODS runner".into()),
        }
    }
    Ok(result(case, corpus, elapsed, None))
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
                let slides = presentation.slides()?;
                let slide = slides
                    .get(index)
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
    let bytes = main.data()?.into_arc();
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
                    .map(|data| data.into_arc())
            });
            let second_start = Arc::clone(&start);
            let second_package = &package;
            let second_task = scope.spawn(move || {
                second_start.wait();
                second_package
                    .main_document_part()
                    .and_then(|part| part.data())
                    .map(|data| data.into_arc())
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
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink,
        source: None,
        execution: None,
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
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: None,
        source: Some(source),
        execution: None,
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
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: None,
        source: None,
        execution: Some(execution),
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

fn environment() -> Environment {
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
        Case, CorpusShape, CountingSink, InstrumentedSource, PayloadKind, RangeSimulationConfig,
        RequestSizeBuckets, SemanticShape, SimulatedRangeSource, SourceBackedPackage, WriterShape,
        XlsxShape, build_cfb_corpus, build_opc_corpus, build_semantic_docx_corpus,
        build_semantic_odp_corpus, build_semantic_ods_corpus, build_semantic_odt_corpus,
        build_semantic_pptx_corpus, build_semantic_rtf_corpus, build_writer_corpus,
        build_xlsx_corpus, payload_bytes, resolve_execution_workers, run_case,
        run_case_with_config, run_scaling_case, simulated_request_delay, statistics,
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
        let measured = run_case(Case::PptSlideOrderSnapshotOpen, &ppt, 0, 1).unwrap();
        assert_eq!(measured.case, "ppt_slide_order_snapshot_open");
        assert_eq!(measured.elapsed_ns.samples.len(), 1);
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
    fn semantic_rtf_tiny_corpus_is_deterministic_and_exercises_native_crud() {
        let rtf = build_semantic_rtf_corpus(SemanticShape::Tiny).unwrap();
        let rtf_again = build_semantic_rtf_corpus(SemanticShape::Tiny).unwrap();
        assert_eq!(rtf.archive, rtf_again.archive);
        assert_eq!(rtf.manifest.entry_count, 24);
        assert_eq!(
            rtf.manifest.archive_sha256,
            "ee4a5c5b5d1c97d5fb4f1e862c2787a859136b237addd0d14a7d52ddc9e62328"
        );

        for case in [
            Case::RtfSemanticOpen,
            Case::RtfSemanticListParagraphs,
            Case::RtfSemanticOneParagraph,
            Case::RtfSemanticFullText,
            Case::RtfSemanticStreamSave,
            Case::RtfSemanticNoopEditSave,
            Case::RtfSemanticOneEditSave,
        ] {
            let result = run_case(case, &rtf, 0, 1).unwrap();
            assert_eq!(result.sink.is_some(), case.name().contains("save"));
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

        let ods = build_semantic_ods_corpus(SemanticShape::Tiny).unwrap();
        assert_eq!(
            ods.archive,
            build_semantic_ods_corpus(SemanticShape::Tiny)
                .unwrap()
                .archive
        );
        assert_eq!(ods.manifest.entry_count, 64);
        run_case(Case::OdsSemanticOneEditSave, &ods, 0, 1).unwrap();

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
