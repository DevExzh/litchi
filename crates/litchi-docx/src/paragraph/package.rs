//! Package relationship integration for WordprocessingML paragraphs.

use crate::error::Result;
use crate::hyperlink::Hyperlink;
use litchi_opc::rel::Relationships;

use super::model::Paragraph;

impl Paragraph {
    /// Get all hyperlinks in this paragraph.
    ///
    /// Returns a vector of `Hyperlink` objects representing all hyperlinks
    /// found in this paragraph. Requires relationships to resolve external URLs.
    ///
    /// # Arguments
    ///
    /// * `rels` - Relationships for resolving relationship IDs to URLs
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let para = doc.paragraph(0)?.unwrap();
    /// let hyperlinks = para.hyperlinks(&main_part.rels())?;
    /// for link in hyperlinks {
    ///     println!("Link: {} -> {:?}", link.text(), link.url());
    /// }
    /// ```
    pub fn hyperlinks(&self, rels: &Relationships) -> Result<Vec<Hyperlink>> {
        Ok(Hyperlink::extract_from_paragraph(self.xml_bytes(), rels)?)
    }
}
