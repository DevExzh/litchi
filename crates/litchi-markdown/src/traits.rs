//! Core trait for Markdown conversion.
//!
//! This module defines the [`ToMarkdown`] trait that enables types to be
//! converted to Markdown format.

use crate::config::MarkdownOptions;

/// Core trait for types that can be converted to Markdown.
///
/// Each implementor names the error its conversion reports, so a format
/// adapter keeps its own typed error instead of flattening it into another
/// crate's: the `litchi` facade's unified `File` reports `litchi::Error`,
/// for example.
///
/// # Examples
///
/// ```
/// use litchi_markdown::{MarkdownOptions, ToMarkdown};
///
/// struct Title(&'static str);
///
/// impl ToMarkdown for Title {
///     type Error = std::fmt::Error;
///
///     fn to_markdown_with_options(
///         &self,
///         _options: &MarkdownOptions,
///     ) -> Result<String, Self::Error> {
///         Ok(format!("# {}", self.0))
///     }
/// }
///
/// assert_eq!(Title("Report").to_markdown().unwrap(), "# Report");
/// ```
pub trait ToMarkdown {
    /// The error a conversion reports.
    type Error: std::error::Error + Send + Sync + 'static;

    /// Convert this item to Markdown with default options.
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying document content cannot be read
    /// or converted to Markdown.
    fn to_markdown(&self) -> Result<String, Self::Error> {
        self.to_markdown_with_options(&MarkdownOptions::default())
    }

    /// Convert this item to Markdown with custom options.
    ///
    /// # Arguments
    ///
    /// * `options` - Configuration for the conversion
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying document content cannot be read
    /// or converted to Markdown.
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String, Self::Error>;
}
