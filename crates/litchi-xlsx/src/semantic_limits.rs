//! The semantic budgets an XLSX package keeps (record 0360).
//!
//! [`OpenOptions::semantic_limits`](crate::OpenOptions::semantic_limits)
//! configures them when a package is opened; a package keeps the limits it
//! was opened with, every bounded semantic read uses them, and the edits of
//! the record families default to them.
//! [`Package::with_semantic_limits`](crate::Package::with_semantic_limits)
//! returns the same package under other limits without copying it. Each
//! budget defaults to its family's finite standard limits.
//!
//! ```
//! use litchi_xlsx::{OpenOptions, Package, SemanticLimits, survey};
//!
//! let limits = SemanticLimits::default().survey(survey::Limits::new().with_max_surveys(4));
//! let bytes = Package::new()?.to_bytes()?;
//! let package = Package::from_bytes_with_options(
//!     &bytes,
//!     OpenOptions::new().semantic_limits(limits),
//! )?;
//! assert_eq!(package.semantic_limits().survey_value().max_surveys(), 4);
//! assert!(package.surveys()?.is_empty());
//! # Ok::<(), litchi_xlsx::Error>(())
//! ```

/// The semantic budgets of one XLSX package.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SemanticLimits {
    custom_data: crate::custom_data::Limits,
    survey: crate::survey::Limits,
    calculation_properties: crate::calculation_properties::Limits,
    pivot_table_data: crate::pivot::table_data::Limits,
    form_controls: Option<crate::form_control::OwnerLimits>,
}

macro_rules! copy_limit {
    ($field:ident, $value:ident, $ty:ty, $what:literal) => {
        #[doc = concat!("Bound ", $what, " by `limits`.")]
        #[must_use]
        pub const fn $field(mut self, limits: $ty) -> Self {
            self.$field = limits;
            self
        }

        #[doc = concat!("The budget of ", $what, ".")]
        #[must_use]
        pub const fn $value(&self) -> $ty {
            self.$field
        }
    };
}

impl SemanticLimits {
    copy_limit!(
        custom_data,
        custom_data_value,
        crate::custom_data::Limits,
        "the Custom Data catalog and its edits"
    );
    copy_limit!(
        survey,
        survey_value,
        crate::survey::Limits,
        "the Survey catalog and its edits"
    );
    copy_limit!(
        calculation_properties,
        calculation_properties_value,
        crate::calculation_properties::Limits,
        "the workbook calculation metadata and its edits"
    );
    copy_limit!(
        pivot_table_data,
        pivot_table_data_value,
        crate::pivot::table_data::Limits,
        "the `pivotTableData` reads and edits"
    );

    /// Bound worksheet form-control reads by `limits`.
    #[must_use]
    pub const fn form_controls(mut self, limits: crate::form_control::OwnerLimits) -> Self {
        self.form_controls = Some(limits);
        self
    }

    /// The configured budget of worksheet form-control reads; `None` reads
    /// under the default policy (the finite owner defaults for a loaded
    /// package, the source read limits for a source-backed workbook).
    #[must_use]
    pub const fn form_controls_value(&self) -> Option<crate::form_control::OwnerLimits> {
        self.form_controls
    }
}
