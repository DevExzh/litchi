//! Well-known composite values from [MS-OSHARED].

use super::{
    MAX_COMPOSITE_ELEMENTS, UNICODE_CODEPAGE, invalid, try_clone_string, try_vec_with_capacity,
};
use litchi_cfb::OleError;

/// The string representation selected by a property-set composite value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextEncoding {
    /// A code-page string using the enclosing section's code page.
    Ansi,
    /// A UTF-16LE string using the enclosing section's Unicode code page.
    Unicode,
}

/// One heading and the number of document parts assigned to it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HeadingPair {
    heading: String,
    part_count: u32,
}

impl HeadingPair {
    /// Creates a heading pair with a nonnegative document-part count.
    ///
    /// # Errors
    ///
    /// Returns an error if the heading contains a NUL character or `part_count`
    /// does not fit the `VT_I4` representation.
    pub fn new(heading: impl Into<String>, part_count: u32) -> Result<Self, OleError> {
        let heading_text = heading.into();
        validate_text(&heading_text, "heading")?;
        validate_part_count(part_count)?;
        Ok(Self {
            heading: heading_text,
            part_count,
        })
    }

    /// Returns the heading text.
    #[must_use]
    pub fn heading(&self) -> &str {
        &self.heading
    }

    /// Returns the number of document parts assigned to the heading.
    #[must_use]
    pub const fn part_count(&self) -> u32 {
        self.part_count
    }

    pub(crate) fn validate(&self) -> Result<(), OleError> {
        validate_text(&self.heading, "heading")?;
        validate_part_count(self.part_count)
    }
}

/// The ordered `GKPIDDSI_HEADINGPAIR` value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HeadingPairs {
    pairs: Vec<HeadingPair>,
}

impl HeadingPairs {
    /// Creates an ordered heading-pair collection.
    ///
    /// # Errors
    ///
    /// Returns an error if there are too many pairs or an individual pair is
    /// invalid.
    pub fn new(pairs: Vec<HeadingPair>) -> Result<Self, OleError> {
        let value = Self { pairs };
        value.validate()?;
        Ok(value)
    }

    /// Returns the headings in document order.
    #[must_use]
    pub fn pairs(&self) -> &[HeadingPair] {
        &self.pairs
    }

    /// Returns one heading pair by checked zero-based position.
    #[must_use]
    pub fn pair(&self, index: usize) -> Option<&HeadingPair> {
        self.pairs.get(index)
    }

    /// Returns the number of heading pairs.
    #[must_use]
    pub fn len(&self) -> usize {
        self.pairs.len()
    }

    /// Returns whether the collection has no heading pairs.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.pairs.is_empty()
    }

    /// Returns the total number of document parts named by all headings.
    #[must_use]
    pub fn document_part_count(&self) -> u64 {
        self.pairs
            .iter()
            .map(|pair| u64::from(pair.part_count))
            .sum()
    }

    pub(crate) fn validate(&self) -> Result<(), OleError> {
        if self.pairs.len() > MAX_COMPOSITE_ELEMENTS {
            return Err(invalid("Heading pair count exceeds the safety limit"));
        }
        for pair in &self.pairs {
            pair.validate()?;
        }
        Ok(())
    }

    pub(crate) fn try_clone(&self) -> Result<Self, OleError> {
        let mut pairs = try_vec_with_capacity(self.pairs.len(), "heading pairs")?;
        for pair in &self.pairs {
            pairs.push(HeadingPair {
                heading: try_clone_string(&pair.heading, "heading text")?,
                part_count: pair.part_count,
            });
        }
        Self::new(pairs)
    }
}

/// The ordered `GKPIDDSI_DOCPARTS` value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocParts {
    encoding: TextEncoding,
    values: Vec<String>,
}

impl DocParts {
    /// Creates document parts with an explicit wire string representation.
    ///
    /// # Errors
    ///
    /// Returns an error if there are too many values or a value contains a NUL
    /// character.
    pub fn new(encoding: TextEncoding, values: Vec<String>) -> Result<Self, OleError> {
        let value = Self { encoding, values };
        value.validate()?;
        Ok(value)
    }

    /// Creates a code-page (`VT_VECTOR|VT_LPSTR`) document-parts value.
    ///
    /// # Errors
    ///
    /// Returns an error if there are too many values or a value contains a NUL
    /// character.
    pub fn ansi(values: Vec<String>) -> Result<Self, OleError> {
        Self::new(TextEncoding::Ansi, values)
    }

    /// Creates a UTF-16LE (`VT_VECTOR|VT_LPWSTR`) document-parts value.
    ///
    /// # Errors
    ///
    /// Returns an error if there are too many values or a value contains a NUL
    /// character.
    pub fn unicode(values: Vec<String>) -> Result<Self, OleError> {
        Self::new(TextEncoding::Unicode, values)
    }

    /// Returns the wire string representation.
    #[must_use]
    pub const fn encoding(&self) -> TextEncoding {
        self.encoding
    }

    /// Returns the ordered document-part names.
    #[must_use]
    pub fn values(&self) -> &[String] {
        &self.values
    }

    /// Returns one document-part name by checked zero-based position.
    pub fn value(&self, index: usize) -> Option<&str> {
        self.values.get(index).map(String::as_str)
    }

    /// Returns the number of document parts.
    #[must_use]
    pub fn len(&self) -> usize {
        self.values.len()
    }

    /// Returns whether the collection has no document parts.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    pub(crate) fn validate(&self) -> Result<(), OleError> {
        if self.values.len() > MAX_COMPOSITE_ELEMENTS {
            return Err(invalid("Document-part count exceeds the safety limit"));
        }
        for value in &self.values {
            validate_text(value, "document-part name")?;
        }
        Ok(())
    }

    pub(crate) fn validate_for_codepage(&self, codepage: u16) -> Result<(), OleError> {
        self.validate()?;
        let expected = if codepage == UNICODE_CODEPAGE {
            TextEncoding::Unicode
        } else {
            TextEncoding::Ansi
        };
        if self.encoding != expected {
            return Err(invalid(
                "Document-part string type does not match the section code page",
            ));
        }
        Ok(())
    }

    pub(crate) fn try_clone(&self) -> Result<Self, OleError> {
        let mut values = try_vec_with_capacity(self.values.len(), "document parts")?;
        for value in &self.values {
            values.push(try_clone_string(value, "document-part name")?);
        }
        Self::new(self.encoding, values)
    }
}

fn validate_text(value: &str, description: &str) -> Result<(), OleError> {
    if value.contains('\0') {
        return Err(invalid(format!("{description} must not contain NUL")));
    }
    Ok(())
}

fn validate_part_count(value: u32) -> Result<(), OleError> {
    if value > i32::MAX.unsigned_abs() {
        return Err(invalid("Heading pair part count exceeds VT_I4"));
    }
    Ok(())
}
