//! The semantic read budgets a DOCX package keeps (record 0355).
//!
//! [`OpenOptions::semantic_limits`](crate::OpenOptions::semantic_limits)
//! configures them when a package is opened; a package keeps the limits it
//! was opened (or built) with and every bounded semantic read uses them.
//! [`Package::with_semantic_limits`](crate::Package::with_semantic_limits)
//! (and the source-backed owner's equivalent) returns the same package under
//! other limits without copying it. Each budget defaults to its family's
//! finite standard limits.
//!
//! ```
//! use litchi_docx::{OpenOptions, Package, SemanticLimits, ink};
//!
//! let limits = SemanticLimits::default().ink(ink::Limits::default());
//! let bytes = Package::new()?.to_bytes()?;
//! let package = Package::from_bytes_with_options(
//!     &bytes,
//!     OpenOptions::default().semantic_limits(limits),
//! )?;
//! assert!(package.ink()?.annotations().is_empty());
//! # Ok::<(), litchi_docx::Error>(())
//! ```

/// The semantic read budgets of one DOCX package.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SemanticLimits {
    revisions: crate::revision::Limits,
    sections: crate::section::Limits,
    stories: crate::package::StoryLimits,
    conflicts: crate::revision::conflict::Limits,
    tracked_revisions: crate::revision::authoring::Limits,
    content_controls: crate::content_control::PackageLimits,
    ink: crate::ink::Limits,
    task_panes: Option<crate::web_extensions::Limits>,
    hyperlinks: crate::sanitize::Limits,
    story_hyperlinks: crate::story_hyperlinks::Limits,
    story_text: crate::source_backed::story_text::Limits,
    paragraph_copy: crate::source_backed::paragraph_copy::raw::Limits,
    tail_append: crate::source_backed::tail_append::Limits,
    embedded: litchi_ooxml_common::embedded::Limits,
    glossary: crate::glossary::Limits,
    #[cfg(feature = "automatic-fonts")]
    fonts: litchi_fonts::PreparationLimits,
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

macro_rules! clone_limit {
    ($field:ident, $value:ident, $ty:ty, $what:literal) => {
        #[doc = concat!("Bound ", $what, " by `limits`.")]
        #[must_use]
        pub fn $field(mut self, limits: $ty) -> Self {
            self.$field = limits;
            self
        }

        #[doc = concat!("The budget of ", $what, ".")]
        #[must_use]
        pub const fn $value(&self) -> &$ty {
            &self.$field
        }
    };
}

impl SemanticLimits {
    copy_limit!(
        revisions,
        revisions_value,
        crate::revision::Limits,
        "tracked-revision reads"
    );
    clone_limit!(
        sections,
        sections_value,
        crate::section::Limits,
        "section inventories and layouts"
    );
    copy_limit!(
        stories,
        stories_value,
        crate::package::StoryLimits,
        "story inventories"
    );
    copy_limit!(
        conflicts,
        conflicts_value,
        crate::revision::conflict::Limits,
        "conflict-markup reads"
    );
    copy_limit!(
        tracked_revisions,
        tracked_revisions_value,
        crate::revision::authoring::Limits,
        "tracked-revision dispositions"
    );
    clone_limit!(
        content_controls,
        content_controls_value,
        crate::content_control::PackageLimits,
        "content-control reads"
    );
    copy_limit!(ink, ink_value, crate::ink::Limits, "InkML annotation reads");
    copy_limit!(
        hyperlinks,
        hyperlinks_value,
        crate::sanitize::Limits,
        "external-hyperlink sanitization and redaction"
    );
    copy_limit!(
        story_hyperlinks,
        story_hyperlinks_value,
        crate::story_hyperlinks::Limits,
        "story-hyperlink redaction"
    );
    copy_limit!(
        story_text,
        story_text_value,
        crate::source_backed::story_text::Limits,
        "source-backed story and glossary text edits"
    );
    copy_limit!(
        paragraph_copy,
        paragraph_copy_value,
        crate::source_backed::paragraph_copy::raw::Limits,
        "source-backed plain-paragraph copy and removal"
    );
    copy_limit!(
        tail_append,
        tail_append_value,
        crate::source_backed::tail_append::Limits,
        "source-backed tail appends"
    );
    copy_limit!(
        embedded,
        embedded_value,
        litchi_ooxml_common::embedded::Limits,
        "embedded-object discovery"
    );
    copy_limit!(
        glossary,
        glossary_value,
        crate::glossary::Limits,
        "glossary and building-block reads and edits"
    );
    #[cfg(feature = "automatic-fonts")]
    copy_limit!(
        fonts,
        fonts_value,
        litchi_fonts::PreparationLimits,
        "automatic font preparation"
    );

    /// Bound task-pane and web-extension reads by `limits`.
    #[must_use]
    pub const fn task_panes(mut self, limits: crate::web_extensions::Limits) -> Self {
        self.task_panes = Some(limits);
        self
    }

    /// The budget of task-pane and web-extension reads.
    #[must_use]
    pub fn task_panes_value(&self) -> crate::web_extensions::Limits {
        self.task_panes
            .unwrap_or_else(crate::web_extensions::Limits::standard)
    }
}
