//! Structural validation for the fixed-size writer FIB.

use std::io::{Error, ErrorKind};

use super::IoError;
use super::codec::FIB_SIZE;

/// Validate the fixed layout before the section codecs write into it.
pub(super) fn validate_layout(buffer: &[u8]) -> Result<(), IoError> {
    if buffer.len() == FIB_SIZE {
        Ok(())
    } else {
        Err(Error::new(
            ErrorKind::InvalidData,
            "writer FIB buffer has an unexpected size",
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_the_word_2002_layout() {
        assert!(validate_layout(&[0; FIB_SIZE]).is_ok());
    }

    #[test]
    fn rejects_a_truncated_layout() {
        assert!(validate_layout(&[0; FIB_SIZE - 1]).is_err());
    }
}
