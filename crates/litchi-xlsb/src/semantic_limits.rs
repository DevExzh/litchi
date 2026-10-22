//! The semantic budgets an XLSB package keeps (record 0351 §2.5).
//!
//! A [`crate::Package`], its [`crate::Workbook`] projection and a
//! [`crate::SourceBackedWorkbook`] keep the [`SemanticLimits`] they were
//! opened with ([`crate::OpenOptions::semantic_limits`]); every family read
//! uses them, and a family edit defaults to them unless its
//! `edit_x_with_options` configures another budget.

use litchi_ooxml_common::embedded;

/// The finite semantic budgets of the XLSB families.
///
/// Every dimension defaults to its family's standard finite policy. Setters
/// consume already validated family limits and perform no I/O.
///
/// ```
/// use litchi_xlsb::{OpenOptions, SemanticLimits, theme};
///
/// let limits = SemanticLimits::new().theme(theme::Limits::DEFAULT);
/// let options = OpenOptions::new().semantic_limits(limits);
/// assert_eq!(options.semantic_limits_value(), limits);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SemanticLimits {
    custom_data: crate::custom_data::Limits,
    data_model: crate::data_model::SemanticLimits,
    calculation_chain: crate::calculation_chain::SemanticLimits,
    xml_maps: crate::xml_maps::SemanticLimits,
    volatile_dependencies: crate::volatile_dependencies::SemanticLimits,
    theme: crate::theme::Limits,
    sparklines: crate::sparkline::Limits,
    cell_values: crate::cell_values::Limits,
    binary_index: crate::binary_index::Limits,
    embedded: embedded::Limits,
    xml_parts: litchi_core::Limits,
}

impl Default for SemanticLimits {
    fn default() -> Self {
        Self::new()
    }
}

impl SemanticLimits {
    /// The standard finite policy of every family.
    #[must_use]
    pub fn new() -> Self {
        Self {
            custom_data: crate::custom_data::Limits::default(),
            data_model: crate::data_model::SemanticLimits::DEFAULT,
            calculation_chain: crate::calculation_chain::SemanticLimits::DEFAULT,
            xml_maps: crate::xml_maps::SemanticLimits::DEFAULT,
            volatile_dependencies: crate::volatile_dependencies::SemanticLimits::DEFAULT,
            theme: crate::theme::Limits::DEFAULT,
            sparklines: crate::sparkline::Limits::DEFAULT,
            cell_values: crate::cell_values::Limits::DEFAULT,
            binary_index: crate::binary_index::Limits::DEFAULT,
            embedded: embedded::Limits::default(),
            xml_parts: litchi_core::Limits::for_profile(litchi_core::Profile::Server),
        }
    }

    /// Select the Custom Data budget.
    #[must_use]
    pub const fn custom_data(mut self, limits: crate::custom_data::Limits) -> Self {
        self.custom_data = limits;
        self
    }

    /// The Custom Data budget.
    #[must_use]
    pub const fn custom_data_value(&self) -> crate::custom_data::Limits {
        self.custom_data
    }

    /// Select the Data Model budget.
    #[must_use]
    pub const fn data_model(mut self, limits: crate::data_model::SemanticLimits) -> Self {
        self.data_model = limits;
        self
    }

    /// The Data Model budget.
    #[must_use]
    pub const fn data_model_value(&self) -> crate::data_model::SemanticLimits {
        self.data_model
    }

    /// Select the Calculation Chain budget.
    #[must_use]
    pub const fn calculation_chain(
        mut self,
        limits: crate::calculation_chain::SemanticLimits,
    ) -> Self {
        self.calculation_chain = limits;
        self
    }

    /// The Calculation Chain budget.
    #[must_use]
    pub const fn calculation_chain_value(&self) -> crate::calculation_chain::SemanticLimits {
        self.calculation_chain
    }

    /// Select the Custom XML Maps budget.
    #[must_use]
    pub const fn xml_maps(mut self, limits: crate::xml_maps::SemanticLimits) -> Self {
        self.xml_maps = limits;
        self
    }

    /// The Custom XML Maps budget.
    #[must_use]
    pub const fn xml_maps_value(&self) -> crate::xml_maps::SemanticLimits {
        self.xml_maps
    }

    /// Select the Volatile Dependencies budget.
    #[must_use]
    pub const fn volatile_dependencies(
        mut self,
        limits: crate::volatile_dependencies::SemanticLimits,
    ) -> Self {
        self.volatile_dependencies = limits;
        self
    }

    /// The Volatile Dependencies budget.
    #[must_use]
    pub const fn volatile_dependencies_value(
        &self,
    ) -> crate::volatile_dependencies::SemanticLimits {
        self.volatile_dependencies
    }

    /// Select the Theme XML budget.
    #[must_use]
    pub const fn theme(mut self, limits: crate::theme::Limits) -> Self {
        self.theme = limits;
        self
    }

    /// The Theme XML budget.
    #[must_use]
    pub const fn theme_value(&self) -> crate::theme::Limits {
        self.theme
    }

    /// Select the sparkline budget.
    #[must_use]
    pub const fn sparklines(mut self, limits: crate::sparkline::Limits) -> Self {
        self.sparklines = limits;
        self
    }

    /// The sparkline budget.
    #[must_use]
    pub const fn sparklines_value(&self) -> crate::sparkline::Limits {
        self.sparklines
    }

    /// Select the cell-value snapshot budget.
    #[must_use]
    pub const fn cell_values(mut self, limits: crate::cell_values::Limits) -> Self {
        self.cell_values = limits;
        self
    }

    /// The cell-value snapshot budget.
    #[must_use]
    pub const fn cell_values_value(&self) -> crate::cell_values::Limits {
        self.cell_values
    }

    /// Select the Worksheet Binary Index budget.
    #[must_use]
    pub const fn binary_index(mut self, limits: crate::binary_index::Limits) -> Self {
        self.binary_index = limits;
        self
    }

    /// The Worksheet Binary Index budget.
    #[must_use]
    pub const fn binary_index_value(&self) -> crate::binary_index::Limits {
        self.binary_index
    }

    /// Select the embedded-object inventory budget.
    #[must_use]
    pub const fn embedded(mut self, limits: embedded::Limits) -> Self {
        self.embedded = limits;
        self
    }

    /// The embedded-object inventory budget.
    #[must_use]
    pub const fn embedded_value(&self) -> embedded::Limits {
        self.embedded
    }

    /// Select the budget of the XML parts an XLSB package carries
    /// (SpreadsheetDrawing, timeline and threaded-comment parts): input bytes,
    /// output bytes, elements (objects) and depth. Each family's format
    /// ceiling still applies.
    #[must_use]
    pub const fn xml_parts(mut self, limits: litchi_core::Limits) -> Self {
        self.xml_parts = limits;
        self
    }

    /// The XML-part budget.
    #[must_use]
    pub const fn xml_parts_value(&self) -> litchi_core::Limits {
        self.xml_parts
    }
}

/// The smaller of a family's format `ceiling` and the caller's `resource`
/// budget in `limits`.
pub(crate) fn xml_ceiling(
    ceiling: usize,
    limits: litchi_core::Limits,
    resource: litchi_core::Resource,
) -> usize {
    ceiling.min(usize::try_from(limits.get(resource)).unwrap_or(usize::MAX))
}
