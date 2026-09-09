//! Package-side settings-extension preprocessing.

use std::borrow::Cow;

use crate::Result;
use litchi_ooxml_common::mce::{Capabilities, Limits as MceLimits, process_markup_compatibility};
use litchi_opc::part::Part;

use super::{WORD_2010_NAMESPACE, WORD_2012_NAMESPACE};

/// Apply the DOCX settings MCE profile while retaining known Word extension
/// namespaces for the typed settings codec.
pub(crate) fn process_part(part: &dyn Part) -> Result<Cow<'_, [u8]>> {
    process_bytes_with_limits(part.blob(), &MceLimits::default())
}

/// Apply the settings extension profile to borrowed XML with an explicit
/// operation ceiling.  The capabilities and fail-closed MCE behavior are the
/// same as [`process_part`]; only the caller-owned finite resource policy is
/// narrowed for source-backed topology checks.
pub(crate) fn process_bytes_with_limits<'a>(
    bytes: &'a [u8],
    limits: &MceLimits,
) -> Result<Cow<'a, [u8]>> {
    let mut capabilities = Capabilities::default();
    capabilities
        .understand_namespace(WORD_2010_NAMESPACE)
        .understand_namespace(WORD_2012_NAMESPACE);
    Ok(process_markup_compatibility(bytes, &capabilities, limits)?.xml)
}
