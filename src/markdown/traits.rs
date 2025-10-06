/// Core trait for Markdown conversion.
///
/// This module defines the `ToMarkdown` trait that enables types to be
/// converted to Markdown format.

use crate::common::Result;
use super::config::MarkdownOptions;

/// Core trait for types that can be converted to Markdown.
///
/// This trait is implemented for Document, Presentation, and their constituent
/// parts (paragraphs, runs, tables, etc.).
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::{Document, markdown::ToMarkdown};
///
/// # fn main() -> Result<(), litchi::Error> {
/// let doc = Document::open("document.docx")?;
///
/// // Convert entire document
/// let markdown = doc.to_markdown()?;
///
/// // Or convert individual parts
/// for para in doc.paragraphs()? {
///     let para_md = para.to_markdown()?;
///     println!("{}", para_md);
/// }
/// # Ok(())
/// # }
/// ```
pub trait ToMarkdown {
    /// Convert this item to Markdown with default options.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::{Document, markdown::ToMarkdown};
    ///
    /// # fn main() -> Result<(), litchi::Error> {
    /// let doc = Document::open("document.docx")?;
    /// let markdown = doc.to_markdown()?;
    /// # Ok(())
    /// # }
    /// ```
    fn to_markdown(&self) -> Result<String> {
        self.to_markdown_with_options(&MarkdownOptions::default())
    }

    /// Convert this item to Markdown with custom options.
    ///
    /// # Arguments
    ///
    /// * `options` - Configuration for the conversion
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::{Document, markdown::{ToMarkdown, MarkdownOptions}};
    ///
    /// # fn main() -> Result<(), litchi::Error> {
    /// let doc = Document::open("document.docx")?;
    /// let options = MarkdownOptions::new().with_styles(true);
    /// let markdown = doc.to_markdown_with_options(&options)?;
    /// # Ok(())
    /// # }
    /// ```
    fn to_markdown_with_options(&self, options: &MarkdownOptions) -> Result<String>;
}

