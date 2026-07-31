//! Migration adapters for canonical SpreadsheetML namespace helpers.

use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::NamespaceResolver;

use crate::error::{OoxmlError, Result};

pub(crate) use litchi_xlsx::raw::namespace::{SPREADSHEETML_NAMESPACE, is_spreadsheetml_name};

pub(crate) fn relationship_attribute_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Option<String>> {
    litchi_xlsx::raw::namespace::relationship_attribute_value(element, name, decoder, resolver)
        .map_err(|error| OoxmlError::InvalidFormat(error.to_string()))
}
