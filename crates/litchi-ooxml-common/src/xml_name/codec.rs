use super::{NameError, NcName, QualifiedName, is_qualified_name};

/// Parse one XML Schema `QName` lexical value.
///
/// # Errors
///
/// Returns `NameError::InvalidQualifiedName` for an invalid lexical value.
pub fn parse(value: &str) -> Result<QualifiedName, NameError> {
    if !is_qualified_name(value) {
        return Err(NameError::InvalidQualifiedName(value.to_owned()));
    }
    let mut components = value.split(':');
    let first = components.next().unwrap_or_default();
    let second = components.next();
    if let Some(local_lexeme) = second {
        let prefix = NcName::new(first)
            .map_err(|_error| NameError::InvalidQualifiedName(value.to_owned()))?;
        let local = NcName::new(local_lexeme)
            .map_err(|_error| NameError::InvalidQualifiedName(value.to_owned()))?;
        Ok(QualifiedName::from_parts(Some(&prefix), &local))
    } else {
        let local = NcName::new(first)
            .map_err(|_error| NameError::InvalidQualifiedName(value.to_owned()))?;
        Ok(QualifiedName::from_parts(None, &local))
    }
}

/// Write a `QName` without re-encoding or allocating its lexical value.
///
/// # Errors
///
/// Returns the formatter's error when the destination rejects the write.
pub fn write<W>(value: &QualifiedName, output: &mut W) -> std::fmt::Result
where
    W: std::fmt::Write,
{
    output.write_str(value.as_str())
}
