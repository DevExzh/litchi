//! Package-side settings-extension preprocessing.

use std::borrow::Cow;

use crate::Result;
use litchi_ooxml_common::mce::{Capabilities, Limits, process_markup_compatibility};
use litchi_opc::part::Part;

use super::{WORD_2010_NAMESPACE, WORD_2012_NAMESPACE};

/// Apply the DOCX settings MCE profile while retaining known Word extension
/// namespaces for the typed settings codec.
pub(crate) fn process_part(part: &dyn Part) -> Result<Cow<'_, [u8]>> {
    let mut capabilities = Capabilities::default();
    capabilities
        .understand_namespace(WORD_2010_NAMESPACE)
        .understand_namespace(WORD_2012_NAMESPACE);
    Ok(process_markup_compatibility(part.blob(), &capabilities, &Limits::default())?.xml)
}
