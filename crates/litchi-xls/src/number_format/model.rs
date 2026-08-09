//! Semantic BIFF8 workbook-formatting values.

use std::collections::HashMap;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DateSystem {
    #[default]
    Excel1900,
    Excel1904,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberFormat {
    pub(super) id: u16,
    pub(super) code: String,
    pub(super) date_time: bool,
}

impl NumberFormat {
    #[must_use]
    pub fn id(&self) -> u16 {
        self.id
    }

    #[must_use]
    pub fn code(&self) -> &str {
        &self.code
    }

    #[must_use]
    pub fn is_builtin_override(&self) -> bool {
        self.id < 164
    }

    #[must_use]
    pub fn is_date_time(&self) -> bool {
        self.date_time
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExtendedFormatKind {
    Cell { parent_style_xf: u16 },
    Style,
}

/// Local-application versus parent-inheritance semantics for the six XF property families.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExtendedFormatApplications {
    number_format: bool,
    font: bool,
    alignment: bool,
    border: bool,
    fill: bool,
    protection: bool,
}

impl ExtendedFormatApplications {
    #[must_use]
    pub fn applies_number_format(&self) -> bool {
        self.number_format
    }
    #[must_use]
    pub fn applies_font(&self) -> bool {
        self.font
    }
    #[must_use]
    pub fn applies_alignment(&self) -> bool {
        self.alignment
    }
    #[must_use]
    pub fn applies_border(&self) -> bool {
        self.border
    }
    #[must_use]
    pub fn applies_fill(&self) -> bool {
        self.fill
    }
    #[must_use]
    pub fn applies_protection(&self) -> bool {
        self.protection
    }
    #[must_use]
    pub fn inherits_number_format(&self) -> bool {
        !self.number_format
    }
    #[must_use]
    pub fn inherits_font(&self) -> bool {
        !self.font
    }
    #[must_use]
    pub fn inherits_alignment(&self) -> bool {
        !self.alignment
    }
    #[must_use]
    pub fn inherits_border(&self) -> bool {
        !self.border
    }
    #[must_use]
    pub fn inherits_fill(&self) -> bool {
        !self.fill
    }
    #[must_use]
    pub fn inherits_protection(&self) -> bool {
        !self.protection
    }

    pub(super) fn all_local() -> Self {
        Self {
            number_format: true,
            font: true,
            alignment: true,
            border: true,
            fill: true,
            protection: true,
        }
    }

    pub(super) fn from_cell_bits(bits: u8) -> Self {
        Self {
            number_format: bits & 0x01 != 0,
            font: bits & 0x02 != 0,
            alignment: bits & 0x04 != 0,
            border: bits & 0x08 != 0,
            fill: bits & 0x10 != 0,
            protection: bits & 0x20 != 0,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtendedFormat {
    pub(super) index: u16,
    pub(super) font_index: u16,
    pub(super) number_format_id: u16,
    pub(super) kind: ExtendedFormatKind,
    pub(super) applications: ExtendedFormatApplications,
    pub(super) quote_prefix: bool,
    pub(super) pivot_button: bool,
    pub(super) has_xf_extension: bool,
    pub(super) locked: bool,
    pub(super) hidden: bool,
    pub(super) alignment: crate::alignment::CellAlignment,
    pub(super) borders: crate::border_fill::CellBorders,
    pub(super) fill: crate::border_fill::CellFill,
}

impl ExtendedFormat {
    #[must_use]
    pub fn index(&self) -> u16 {
        self.index
    }

    #[must_use]
    pub fn number_format_id(&self) -> u16 {
        self.number_format_id
    }

    /// Returns the logical index of the global Font record used by this XF.
    #[must_use]
    pub fn font_index(&self) -> u16 {
        self.font_index
    }

    #[must_use]
    pub fn kind(&self) -> ExtendedFormatKind {
        self.kind
    }

    #[must_use]
    pub fn parent_style_xf_index(&self) -> Option<u16> {
        match self.kind {
            ExtendedFormatKind::Cell { parent_style_xf } => Some(parent_style_xf),
            ExtendedFormatKind::Style => None,
        }
    }

    #[must_use]
    pub fn applications(&self) -> ExtendedFormatApplications {
        self.applications
    }

    #[must_use]
    pub fn quote_prefix(&self) -> bool {
        self.quote_prefix
    }
    #[must_use]
    pub fn pivot_button(&self) -> bool {
        self.pivot_button
    }
    #[must_use]
    pub fn has_xf_extension(&self) -> bool {
        self.has_xf_extension
    }

    #[must_use]
    pub fn is_cell_format(&self) -> bool {
        matches!(self.kind, ExtendedFormatKind::Cell { .. })
    }

    #[must_use]
    pub fn locked(&self) -> bool {
        self.locked
    }

    #[must_use]
    pub fn hidden(&self) -> bool {
        self.hidden
    }

    #[must_use]
    pub fn alignment(&self) -> &crate::alignment::CellAlignment {
        &self.alignment
    }

    /// Returns the border metadata stored by this XF record.
    #[must_use]
    pub fn borders(&self) -> &crate::border_fill::CellBorders {
        &self.borders
    }

    /// Returns the fill pattern and colors stored by this XF record.
    #[must_use]
    pub fn fill(&self) -> &crate::border_fill::CellFill {
        &self.fill
    }
}

/// Borrowed effective formatting after applying a `CellXF`'s parent `StyleXF`.
#[derive(Debug, Clone, Copy)]
pub struct EffectiveExtendedFormat<'a> {
    direct: &'a ExtendedFormat,
    parent: Option<&'a ExtendedFormat>,
}

impl<'a> EffectiveExtendedFormat<'a> {
    #[must_use]
    pub fn direct(&self) -> &'a ExtendedFormat {
        self.direct
    }
    #[must_use]
    pub fn parent_style(&self) -> Option<&'a ExtendedFormat> {
        self.parent
    }

    fn source(&self, local: bool) -> &'a ExtendedFormat {
        if local {
            self.direct
        } else {
            self.parent.unwrap_or(self.direct)
        }
    }

    #[must_use]
    pub fn number_format_source(&self) -> &'a ExtendedFormat {
        self.source(self.direct.applications.applies_number_format())
    }
    #[must_use]
    pub fn font_source(&self) -> &'a ExtendedFormat {
        self.source(self.direct.applications.applies_font())
    }
    #[must_use]
    pub fn alignment_source(&self) -> &'a ExtendedFormat {
        self.source(self.direct.applications.applies_alignment())
    }
    #[must_use]
    pub fn border_source(&self) -> &'a ExtendedFormat {
        self.source(self.direct.applications.applies_border())
    }
    #[must_use]
    pub fn fill_source(&self) -> &'a ExtendedFormat {
        self.source(self.direct.applications.applies_fill())
    }
    #[must_use]
    pub fn protection_source(&self) -> &'a ExtendedFormat {
        self.source(self.direct.applications.applies_protection())
    }

    #[must_use]
    pub fn number_format_id(&self) -> u16 {
        self.number_format_source().number_format_id
    }
    #[must_use]
    pub fn font_index(&self) -> u16 {
        self.font_source().font_index
    }
    #[must_use]
    pub fn alignment(&self) -> &'a crate::alignment::CellAlignment {
        &self.alignment_source().alignment
    }
    #[must_use]
    pub fn borders(&self) -> &'a crate::border_fill::CellBorders {
        &self.border_source().borders
    }
    #[must_use]
    pub fn fill(&self) -> &'a crate::border_fill::CellFill {
        &self.fill_source().fill
    }
    #[must_use]
    pub fn locked(&self) -> bool {
        self.protection_source().locked
    }
    #[must_use]
    pub fn hidden(&self) -> bool {
        self.protection_source().hidden
    }
    #[must_use]
    pub fn quote_prefix(&self) -> bool {
        self.direct.quote_prefix
    }
    #[must_use]
    pub fn pivot_button(&self) -> bool {
        self.direct.pivot_button
    }
    #[must_use]
    pub fn has_xf_extension(&self) -> bool {
        self.direct.has_xf_extension
    }
}

#[derive(Debug, Clone, Default)]
pub struct Formatting {
    pub(super) date_system: DateSystem,
    pub(super) number_formats: Vec<NumberFormat>,
    pub(super) extended_formats: Vec<ExtendedFormat>,
    pub(super) differential_formats: Vec<crate::differential_format::DifferentialFormat>,
    pub(super) xf_extensions: Vec<crate::xf_ext::XfExt>,
    pub(super) format_by_id: HashMap<u16, usize>,
}

impl Formatting {
    #[must_use]
    pub fn date_system(&self) -> DateSystem {
        self.date_system
    }

    /// Explicit BIFF `Format` records in their original workbook order.
    #[must_use]
    pub fn number_formats(&self) -> &[NumberFormat] {
        &self.number_formats
    }

    /// BIFF `XF` records in index order, including style-XF slots.
    #[must_use]
    pub fn extended_formats(&self) -> &[ExtendedFormat] {
        &self.extended_formats
    }

    /// Global `DXF` records in zero-based reference order.
    #[must_use]
    pub fn differential_formats(&self) -> &[crate::differential_format::DifferentialFormat] {
        &self.differential_formats
    }

    /// `XFExt` formatting property extensions (MS-XLS 2.4.355), in record order.
    #[must_use]
    pub fn xf_extensions(&self) -> &[crate::xf_ext::XfExt] {
        &self.xf_extensions
    }

    #[must_use]
    pub fn differential_format(
        &self,
        id: crate::table_styles::DifferentialFormatId,
    ) -> Option<&crate::differential_format::DifferentialFormat> {
        self.differential_formats.get(id.index() as usize)
    }

    #[must_use]
    pub fn number_format(&self, id: u16) -> Option<&NumberFormat> {
        self.format_by_id
            .get(&id)
            .and_then(|index| self.number_formats.get(*index))
    }

    #[must_use]
    pub fn extended_format(&self, index: u16) -> Option<&ExtendedFormat> {
        self.extended_formats.get(index as usize)
    }

    #[must_use]
    pub fn effective_extended_format(&self, index: u16) -> Option<EffectiveExtendedFormat<'_>> {
        let direct = self.extended_format(index)?;
        let parent = direct
            .parent_style_xf_index()
            .and_then(|parent| self.extended_format(parent));
        Some(EffectiveExtendedFormat { direct, parent })
    }

    pub fn is_date_time_format(&self, id: u16) -> bool {
        self.number_format(id)
            .map_or_else(|| is_builtin_date_time(id), NumberFormat::is_date_time)
    }
}

fn is_builtin_date_time(id: u16) -> bool {
    matches!(id, 14..=22 | 27..=36 | 45..=47 | 50..=58)
}
