//! Focused validation rules shared by the package editing layers.

use super::super::model::{Error, Result};

pub(super) fn validate_mail_merge_external_uri(uri: &str) -> Result<()> {
    if uri.is_empty() || uri.len() > 32 * 1024 || uri.chars().any(char::is_control) {
        return Err(Error::InvalidFormat(
            "mail-merge external target is empty or exceeds URI limits".into(),
        ));
    }
    Ok(())
}

pub(super) fn validate_mail_merge_internal_source(
    bytes: &[u8],
    content_type: &str,
    extension: &str,
) -> Result<()> {
    if bytes.len() > 128 * 1024 * 1024 {
        return Err(Error::InvalidFormat(
            "mail-merge source exceeds the 128 MiB authoring limit".into(),
        ));
    }
    if content_type.is_empty()
        || content_type.len() > 1024
        || content_type.chars().any(char::is_control)
    {
        return Err(Error::InvalidFormat(
            "mail-merge source content type is invalid".into(),
        ));
    }
    if extension.is_empty()
        || extension.len() > 16
        || !extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
    {
        return Err(Error::InvalidFormat(
            "mail-merge source extension is invalid".into(),
        ));
    }
    Ok(())
}
