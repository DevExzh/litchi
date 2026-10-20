//! Options shared by the line-oriented formats (SYLK, DIF, fixed width).

use std::io::{Read, Write};

use litchi_core::BomKind;
use litchi_core::sheet::CellValue;

use crate::encoding::Encoding;
use crate::error::{Error, Format, OptionKind, Result};
use crate::limits::{CellBudget, SemanticLimits, SourceLimits};

/// Encoding, byte-order mark and limits of one format's options.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub(crate) struct Common {
    pub(crate) encoding: Encoding,
    pub(crate) write_bom: Option<BomKind>,
    pub(crate) source_limits: SourceLimits,
    pub(crate) semantic_limits: SemanticLimits,
}

impl Common {
    /// Read and decode `reader`, then hand its lines to `parse`.
    pub(crate) fn read<T>(
        &self,
        reader: impl Read,
        parse: impl FnOnce(&str, &mut CellBudget) -> Result<T>,
    ) -> Result<T> {
        let bytes = self.source_limits.read(reader)?;
        let (text, _) = self.encoding.decode(&bytes)?;
        let mut budget = CellBudget::new(self.semantic_limits);
        parse(&text, &mut budget)
    }
}

/// Grow `rows` so that `(row, column)` (zero-based) exists, charging every
/// new cell, and refuse a coordinate outside the spreadsheet grid.
pub(crate) fn place(
    rows: &mut Vec<Vec<CellValue>>,
    row: usize,
    column: usize,
    budget: &mut CellBudget,
    format: Format,
    line: usize,
) -> Result<()> {
    if row >= litchi_sheet::ROWS as usize || column >= litchi_sheet::COLUMNS as usize {
        return Err(Error::OutOfGrid { format, line });
    }
    if rows.len() <= row {
        let added = row + 1 - rows.len();
        rows.try_reserve(added)?;
        rows.resize_with(row + 1, Vec::new);
    }
    let cells = &mut rows[row];
    if cells.len() <= column {
        let added = column + 1 - cells.len();
        budget.charge(added)?;
        cells.try_reserve(added)?;
        cells.resize(column + 1, CellValue::Empty);
    }
    Ok(())
}

/// Generate the shared setters of a format's `Options`.
macro_rules! common_setters {
    () => {
        /// Decode the source with this encoding.
        #[must_use]
        pub const fn encoding(mut self, encoding: $crate::Encoding) -> Self {
            self.common.encoding = encoding;
            self
        }

        /// Write this byte-order mark before the output, or none.
        #[must_use]
        pub const fn write_bom(mut self, bom: Option<litchi_core::BomKind>) -> Self {
            self.common.write_bom = bom;
            self
        }

        /// Bound the bytes read.
        #[must_use]
        pub const fn source_limits(mut self, limits: $crate::SourceLimits) -> Self {
            self.common.source_limits = limits;
            self
        }

        /// Bound the cells materialized.
        #[must_use]
        pub const fn semantic_limits(mut self, limits: $crate::SemanticLimits) -> Self {
            self.common.semantic_limits = limits;
            self
        }
    };
}

pub(crate) use common_setters;

/// Write `body`'s UTF-8 output to `writer` in the encoding `bom` names,
/// preceded by that mark: UTF-8 (or no mark) streams directly, UTF-16 is
/// transcoded. UTF-32 is [`Error::InvalidOption`] naming
/// [`OptionKind::Bom`]: no reader in this crate decodes it.
pub(crate) fn emit(
    mut writer: impl Write,
    bom: Option<BomKind>,
    body: impl FnOnce(&mut dyn Write) -> Result<()>,
) -> Result<()> {
    match bom {
        None => body(&mut writer),
        Some(BomKind::Utf8) => {
            writer.write_all(&litchi_core::UTF8_BOM)?;
            body(&mut writer)
        },
        Some(kind @ (BomKind::Utf16Le | BomKind::Utf16Be)) => {
            let mut utf8 = Vec::new();
            body(&mut utf8)?;
            let text = String::from_utf8(utf8).map_err(|error| Error::Encoding {
                offset: error.utf8_error().valid_up_to(),
            })?;
            let mut bytes = Vec::new();
            bytes.try_reserve(2 + text.len() * 2)?;
            bytes.extend_from_slice(kind.as_bytes());
            for unit in text.encode_utf16() {
                bytes.extend_from_slice(&if kind == BomKind::Utf16Le {
                    unit.to_le_bytes()
                } else {
                    unit.to_be_bytes()
                });
            }
            writer.write_all(&bytes)?;
            Ok(())
        },
        Some(_) => Err(Error::InvalidOption {
            option: OptionKind::Bom,
        }),
    }
}
