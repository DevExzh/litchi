//! Typed external-workbook relationship vocabulary.

use litchi_opc::constants::relationship_type;

/// Every relationship type accepted as an external-workbook target.
///
/// Ordered standard-first so the common case matches on the first comparison.
pub const EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES: &[&str] = &[
    // ECMA-376 transitional and strict.
    relationship_type::EXTERNAL_LINK_PATH,
    relationship_type::STRICT_EXTERNAL_LINK_PATH,
    // MS-OE376 resolution-origin families.
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup",
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlAlternateStartup",
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary",
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing",
    // Long-path families for targets beyond the classic path limit.
    "http://schemas.microsoft.com/office/2019/04/relationships/externalLinkLongPath",
    "http://schemas.microsoft.com/office/2019/04/relationships/xlExternalLinkLongPath/xlStartup",
    "http://schemas.microsoft.com/office/2019/04/relationships/xlExternalLinkLongPath/xlAlternateStartup",
    "http://schemas.microsoft.com/office/2009/04/relationships/xlExternalLinkLongPath/xlPathMissing",
    "http://schemas.microsoft.com/office/2009/04/relationships/xlExternalLinkLongPath/xlLibrary",
];
