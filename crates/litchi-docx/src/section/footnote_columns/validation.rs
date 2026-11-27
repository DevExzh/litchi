#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Lexical, namespace, and resource validation for `footnoteColumns`.

use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use quick_xml::name::{Namespace, ResolveResult};

use super::model::Layout;

/// Word 2012 `WordprocessingML` namespace used by this extension.
pub(crate) const WORD_2012_NAMESPACE: &[u8] =
    b"http://schemas.microsoft.com/office/word/2012/wordml";
/// Markup-compatibility namespace used by `mc:Ignorable`.
pub(crate) const MC_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/markup-compatibility/2006";
/// Maximum section fragment accepted by this focused owner.
pub(crate) const MAX_XML_BYTES: usize = 4 * 1024 * 1024;
/// Maximum XML nesting accepted by this focused owner.
pub(crate) const MAX_XML_DEPTH: usize = 64;
/// Maximum XML elements accepted by this focused owner.
pub(crate) const MAX_XML_NODES: usize = 4096;
/// Maximum number of namespace bindings retained from an enclosing document
/// when a section is detached from its package part.
pub(crate) const MAX_CONTEXT_BINDINGS: usize = 4096;
/// Maximum bytes retained by a detached namespace/markup-compatibility
/// context.
pub(crate) const MAX_CONTEXT_BYTES: usize = 256 * 1024;

pub(crate) fn validate_layout(value: Option<Layout>) -> Result<()> {
    if value.is_some_and(|value| value.columns() < 0) {
        return Err(Error::InvalidFormat(
            "footnote column count cannot be negative".into(),
        ));
    }
    Ok(())
}

pub(crate) fn parse_columns(value: &str) -> Result<Layout> {
    let value = value.trim();
    if value.is_empty() {
        return Err(Error::InvalidFormat(
            "footnoteColumns requires a decimal column count".into(),
        ));
    }
    let columns = value.parse::<i32>().map_err(|_source_error| {
        Error::InvalidFormat(format!(
            "invalid footnoteColumns decimal column count '{value}'"
        ))
    })?;
    Layout::new(columns)
}

#[inline]
pub(crate) fn is_word_2012(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == WORD_2012_NAMESPACE
    )
}

/// Match an extension element whose namespace declaration was inherited by a
/// caller that extracted a `sectPr` fragment from its document part.
pub(crate) fn is_word_2012_element(namespace: &ResolveResult<'_>, prefix: Option<&[u8]>) -> bool {
    is_word_2012(namespace)
        || matches!(
            (namespace, prefix),
            (ResolveResult::Unknown(value), Some(prefix))
                if value.as_slice() == b"w12" && prefix == b"w12"
        )
}

#[inline]
pub(crate) fn is_markup_compatibility(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == MC_NAMESPACE
    )
}

pub(crate) fn is_inherited_markup_compatibility(namespace: &ResolveResult<'_>) -> bool {
    is_markup_compatibility(namespace)
        || matches!(namespace, ResolveResult::Unknown(value) if value.as_slice() == b"mc")
}

pub(crate) fn is_word_value_attribute(
    namespace: &ResolveResult<'_>,
    prefix: Option<&[u8]>,
) -> bool {
    is_wordprocessing_namespace(namespace)
        || matches!(
            (namespace, prefix),
            (ResolveResult::Unknown(value), Some(prefix))
                if value.as_slice() == b"w" && prefix == b"w"
        )
}

pub(crate) fn has_ignorable_prefix(value: &str, prefix: &[u8]) -> bool {
    let Ok(prefix) = std::str::from_utf8(prefix) else {
        return false;
    };
    value.split_whitespace().any(|item| item == prefix)
}
