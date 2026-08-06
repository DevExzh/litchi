//! Semantic facade for the paired caption metadata tables.

use super::codec::{AUTO_CAPTION_FIB_INDEX, CAPTION_FIB_INDEX};
use super::model::{AutoTable, LabelTable};
use super::validation::validate_references;
use crate::package::Result;
use crate::parts::fib::FileInformationBlock;

/// Caption metadata selected from one Word document's FIB/table stream.
///
/// The two tables are optional because Word only defines their pointers for
/// the Normal template. The facade keeps the tables separate and validates
/// every `SttbfAutoCaption` index against the label table when both are read.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Tables {
    labels: Option<LabelTable>,
    auto: Option<AutoTable>,
}

impl Tables {
    /// Construct a validated pair of detached metadata tables.
    pub fn try_new(labels: Option<LabelTable>, auto: Option<AutoTable>) -> Result<Self> {
        validate_references(labels.as_ref(), auto.as_ref())?;
        Ok(Self { labels, auto })
    }

    /// Parse both caption tables from their Word FIB/table-stream pointers.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        // MS-DOC defines these pointers only for the Normal template. A
        // non-template document must ignore them even if a malformed producer
        // left nonzero ranges behind.
        if !fib.is_template() {
            return Ok(Self::default());
        }
        let labels =
            super::codec::parse_fib_table(fib, table_stream, CAPTION_FIB_INDEX, "SttbfCaption")?
                .map(LabelTable::parse_bytes)
                .transpose()?;
        let auto = super::codec::parse_fib_table(
            fib,
            table_stream,
            AUTO_CAPTION_FIB_INDEX,
            "SttbfAutoCaption",
        )?
        .map(AutoTable::parse_bytes)
        .transpose()?;
        Self::try_new(labels, auto)
    }

    /// Caption label definitions, when `SttbfCaption` is present.
    pub fn labels(&self) -> Option<&LabelTable> {
        self.labels.as_ref()
    }

    /// Caption label definitions, using the protocol-oriented terminology.
    pub fn captions(&self) -> Option<&LabelTable> {
        self.labels()
    }

    /// Automatic-caption ProgID mappings, when `SttbfAutoCaption` is present.
    pub fn auto(&self) -> Option<&AutoTable> {
        self.auto.as_ref()
    }

    /// Alias phrased in the protocol vocabulary for callers inspecting a
    /// document's AutoCaption settings.
    pub fn auto_captions(&self) -> Option<&AutoTable> {
        self.auto()
    }

    /// Whether either optional caption range is present.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.labels.is_some() || self.auto.is_some()
    }

    /// Creates an empty pair of absent caption ranges.
    #[must_use]
    pub const fn empty() -> Self {
        Self {
            labels: None,
            auto: None,
        }
    }
}
