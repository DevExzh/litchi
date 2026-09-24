#![allow(
    clippy::unwrap_used,
    reason = "focused integration fixtures deliberately fail fast"
)]

//! Source-bound tests for the diagnostic `cachedUniqueNames` owner.
//!
//! The fixtures in this file are intentionally synthetic OPC packages.  They
//! exercise the semantic cache/field selectors, the exact extension owner,
//! and the two model-connection routes recorded in the design document.  The
//! `cacheField` host contains schema-valid `sharedItems`; there is deliberately
//! no invented `cacheField/items` element and these tests make no native
//! Office interoperability claim.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, ReadLimits, TargetMode};
use litchi_xlsx::Error;
use litchi_xlsx::pivot::cached_unique_names::{
    CacheSelector, CachedUniqueName, DiagnosticStatus, FieldSelector, PivotCacheId, Snapshot,
    Transaction,
};
use litchi_xlsx::workbook::Workbook;

const TRANSITIONAL_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_MAIN: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const TRANSITIONAL_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const X15_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const X14_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MCE_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const FOREIGN_NS: &str = "http://example.com/cached-unique-names-foreign";
const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";

const CACHED_UNIQUE_NAMES_URI: &str = "{4F2E5C28-24EA-4EB8-9CBF-B6C8F9C3D259}";
const PIVOT_CACHE_ID_VERSION_URI: &str = "{ABF5C744-AB39-4b91-8756-CFA1BBC848D5}";
const PIVOT_CACHE_DEFINITION_URI: &str = "{725AE2AE-9491-48BE-B2B4-4EB974FC3084}";
const F057_URI: &str = "{F057638F-6D5F-4E77-A914-E7F072B9BCA8}";
const DE250_URI: &str = "{DE250136-89BD-433C-8126-D09CA5730AF9}";
const WRONG_URI: &str = "{00000000-0000-0000-0000-000000000000}";
const D799_URI: &str = "{D79990A0-CA42-45E3-83F4-45C500A0EAA5}";
const CACHE_ID: u32 = 42;

const WORKBOOK_URI: &str = "/xl/workbook.xml";
const SHEET_URI: &str = "/xl/worksheets/sheet1.xml";
const CACHE_URI: &str = "/xl/pivotCache/pivotCacheDefinition1.xml";
const CONNECTIONS_URI: &str = "/xl/connections.xml";
const CONNECTIONS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CacheIdVersionMode {
    Valid,
    Version255,
    VersionPlus255,
    Version256,
    UnknownChild,
    NonWhitespaceText,
    CData,
    UnknownAttribute,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum MceRequiresMode {
    Valid,
    XmlWhitespace,
    Nbsp,
    DuplicatePrefix,
    SameUriAlias,
    MceNamespace,
    XmlReservedPrefix,
    EncodedXmlReservedPrefix,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum MceAttributeMode {
    None,
    QualifiedBranchAttributes,
    UnqualifiedBranchAttribute,
    ForeignBranchAttributeWithoutIgnorable,
    ForeignBranchAttributeWithIgnorable,
    ForeignChildOwnIgnorable,
    ForeignChildAncestorIgnorable,
    IgnorableNamesMceNamespace,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ExtensionUriMode {
    Canonical,
    XmlWhitespacePadding,
    NumericWhitespacePadding,
}

fn namespace_uri(value: &str, escaped: bool) -> String {
    if escaped {
        value.replace('/', "&#47;")
    } else {
        value.to_owned()
    }
}

fn extension_uri(value: &str, mode: ExtensionUriMode) -> String {
    match mode {
        ExtensionUriMode::Canonical => value.to_owned(),
        ExtensionUriMode::XmlWhitespacePadding => format!(" \t{value}\r\n "),
        ExtensionUriMode::NumericWhitespacePadding => format!("&#x9;{value}&#xA;"),
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Dialect {
    Transitional,
    Strict,
}

impl Dialect {
    const fn main(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_MAIN,
            Self::Strict => STRICT_MAIN,
        }
    }

    const fn rel(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_REL,
            Self::Strict => STRICT_REL,
        }
    }

    const fn office_document_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::OFFICE_DOCUMENT,
            Self::Strict => rt::STRICT_OFFICE_DOCUMENT,
        }
    }

    const fn worksheet_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::WORKSHEET,
            Self::Strict => rt::STRICT_WORKSHEET,
        }
    }

    const fn pivot_cache_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::PIVOT_CACHE_DEFINITION,
            Self::Strict => rt::STRICT_PIVOT_CACHE_DEFINITION,
        }
    }

    const fn connections_rel(self) -> &'static str {
        match self {
            Self::Transitional => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections"
            },
            Self::Strict => "http://purl.oclc.org/ooxml/officeDocument/relationships/connections",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Route {
    F057Name,
    ConnectionIdOnly,
    BothConsistent,
    BothMismatch,
    F057WithZero,
    Missing,
    MissingSourceType,
    DefaultConnectionId,
    WorksheetConnectionId,
    DuplicateF057,
    UnresolvedF057,
    MalformedF057,
    WrongF057QName,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Payload {
    Valid,
    WrongUri,
    WrongQName,
    WrongNamespace,
    Empty,
    MceDirect,
    MceNestedDirect,
    MceNestedChoice,
    UnknownUniqueNameAttribute,
    UnknownUniqueNamesChild,
    MissingIndex,
    MissingName,
    DuplicateIndex,
    PositiveIndex,
    NegativeZeroIndex,
    NegativeIndex,
    OverflowIndex,
    UnpairedSurrogate,
    Mce,
    MceIdentical,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ConnectionMode {
    Valid,
    WrongType,
    ModelFalse,
    NonEmptyModelId,
    MissingModelId,
    WrongModelUri,
    D799,
}

#[derive(Clone, Debug)]
struct FixtureOptions {
    dialect: Dialect,
    root_prefix: &'static str,
    x15_prefix: &'static str,
    x14_prefix: &'static str,
    route: Route,
    payload: Payload,
    connection: ConnectionMode,
    names: Vec<(u32, String)>,
    include_opaque_sibling: bool,
    signed: bool,
    duplicate_field_name: bool,
    cache_id_version: CacheIdVersionMode,
    escaped_namespace_uris: bool,
    mce_requires: MceRequiresMode,
    mce_attributes: MceAttributeMode,
    extension_uris: ExtensionUriMode,
}

impl Default for FixtureOptions {
    fn default() -> Self {
        Self {
            dialect: Dialect::Transitional,
            root_prefix: "",
            x15_prefix: "x15",
            x14_prefix: "x14",
            route: Route::F057Name,
            payload: Payload::Valid,
            connection: ConnectionMode::Valid,
            names: vec![
                (0, "North &amp; South".to_owned()),
                (7, "North &amp; South".to_owned()),
            ],
            include_opaque_sibling: true,
            signed: false,
            duplicate_field_name: false,
            cache_id_version: CacheIdVersionMode::Valid,
            escaped_namespace_uris: false,
            mce_requires: MceRequiresMode::Valid,
            mce_attributes: MceAttributeMode::None,
            extension_uris: ExtensionUriMode::Canonical,
        }
    }
}

#[derive(Clone, Debug)]
struct Fixture {
    bytes: Vec<u8>,
    cache_xml: String,
    connections_xml: String,
}

impl Fixture {
    fn build(options: FixtureOptions) -> Self {
        let main = options.dialect.main();
        let rel = options.dialect.rel();
        let root_prefix = options.root_prefix;
        let root = if root_prefix.is_empty() {
            String::new()
        } else {
            format!("{root_prefix}:")
        };
        let x15 = options.x15_prefix;
        let x14 = options.x14_prefix;
        let main_uri = namespace_uri(main, options.escaped_namespace_uris);
        let rel_uri = namespace_uri(rel, options.escaped_namespace_uris);
        let x15_uri = namespace_uri(X15_NS, options.escaped_namespace_uris);
        let x14_uri = namespace_uri(X14_NS, options.escaped_namespace_uris);
        let cached_unique_names_uri =
            extension_uri(CACHED_UNIQUE_NAMES_URI, options.extension_uris);
        let pivot_cache_id_version_uri =
            extension_uri(PIVOT_CACHE_ID_VERSION_URI, options.extension_uris);
        let f057_uri = extension_uri(F057_URI, options.extension_uris);
        let de250_uri = extension_uri(DE250_URI, options.extension_uris);
        let wrong_uri = extension_uri(WRONG_URI, options.extension_uris);
        let d799_uri = extension_uri(D799_URI, options.extension_uris);
        let main_ns = if root_prefix.is_empty() {
            format!(r#"xmlns="{main_uri}""#)
        } else {
            format!(r#"xmlns:{root_prefix}="{main_uri}""#)
        };

        let workbook_xml = format!(
            r#"<{root}workbook {main_ns} xmlns:r="{rel_uri}"><{root}sheets><{root}sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></{root}sheets><{root}pivotCaches><{root}pivotCache cacheId="{CACHE_ID}" r:id="rIdCache"/></{root}pivotCaches></{root}workbook>"#,
            root = root,
            main_ns = main_ns,
            rel_uri = rel_uri,
        );

        let cache_source = match options.route {
            Route::F057Name
            | Route::BothConsistent
            | Route::BothMismatch
            | Route::F057WithZero
            | Route::DuplicateF057
            | Route::UnresolvedF057
            | Route::MalformedF057
            | Route::WrongF057QName => {
                let connection_id = match options.route {
                    Route::BothConsistent => r#" connectionId="5""#,
                    Route::BothMismatch => r#" connectionId="6""#,
                    Route::F057WithZero => r#" connectionId="0""#,
                    Route::UnresolvedF057 => r#" connectionId="5""#,
                    _ => "",
                };
                let source_name = if options.route == Route::UnresolvedF057 {
                    "missing"
                } else {
                    "canonical"
                };
                let source_connection = match options.route {
                    Route::MalformedF057 => {
                        format!(r#"<{x14}:sourceConnection/>"#, x14 = x14)
                    },
                    Route::WrongF057QName => {
                        format!(r#"<{x15}:sourceConnection name="canonical"/>"#, x15 = x15)
                    },
                    _ => format!(
                        r#"<{x14}:sourceConnection name="{source_name}"/>"#,
                        x14 = x14,
                        source_name = source_name,
                    ),
                };
                let duplicate = if options.route == Route::DuplicateF057 {
                    format!(
                        r#"<{root}ext uri="{f057_uri}"><{x14}:sourceConnection name="canonical"/></{root}ext>"#,
                        root = root,
                        x14 = x14,
                        f057_uri = f057_uri,
                    )
                } else {
                    String::new()
                };
                format!(
                    r#"<{root}cacheSource type="external"{connection_id}><{root}extLst><{root}ext uri="{f057_uri}">{source_connection}</{root}ext>{duplicate}</{root}extLst></{root}cacheSource>"#,
                    root = root,
                    connection_id = connection_id,
                    source_connection = source_connection,
                    duplicate = duplicate,
                    f057_uri = f057_uri,
                )
            },
            Route::ConnectionIdOnly => format!(
                r#"<{root}cacheSource type="external" connectionId="5"/>"#,
                root = root
            ),
            Route::DefaultConnectionId => format!(
                r#"<{root}cacheSource type="external" connectionId="0"/>"#,
                root = root
            ),
            Route::Missing => format!(r#"<{root}cacheSource type="external"/>"#, root = root),
            Route::MissingSourceType => {
                format!(r#"<{root}cacheSource connectionId="5"/>"#, root = root)
            },
            Route::WorksheetConnectionId => format!(
                r#"<{root}cacheSource type="worksheet" connectionId="5"/>"#,
                root = root
            ),
        };

        let payload_uri = if options.payload == Payload::WrongUri {
            wrong_uri.as_str()
        } else {
            cached_unique_names_uri.as_str()
        };
        let mce_requires = match options.mce_requires {
            MceRequiresMode::Valid => x15.to_owned(),
            MceRequiresMode::XmlWhitespace => format!("{x15}&#x9;&#xA;&#xD;{x14}"),
            MceRequiresMode::Nbsp => format!("{x15}&#xA0;"),
            MceRequiresMode::DuplicatePrefix => format!("{x15} {x15}"),
            MceRequiresMode::SameUriAlias => format!("{x15} x15alias"),
            MceRequiresMode::MceNamespace
            | MceRequiresMode::XmlReservedPrefix
            | MceRequiresMode::EncodedXmlReservedPrefix => match options.mce_requires {
                MceRequiresMode::MceNamespace => "mc".to_owned(),
                MceRequiresMode::XmlReservedPrefix | MceRequiresMode::EncodedXmlReservedPrefix => {
                    "xml".to_owned()
                },
                _ => unreachable!(),
            },
        };
        let choice_attributes = match options.mce_attributes {
            MceAttributeMode::QualifiedBranchAttributes => {
                format!(
                    r#" mc:Ignorable="{x15}" mc:MustUnderstand="{x15}" mc:ProcessContent="{x15}""#
                )
            },
            MceAttributeMode::UnqualifiedBranchAttribute => r#" future="yes""#.to_owned(),
            MceAttributeMode::ForeignBranchAttributeWithoutIgnorable
            | MceAttributeMode::ForeignBranchAttributeWithIgnorable => {
                let ignorable = if options.mce_attributes
                    == MceAttributeMode::ForeignBranchAttributeWithIgnorable
                {
                    r#" mc:Ignorable="foo""#
                } else {
                    ""
                };
                format!(r#" foo:metadata="yes"{ignorable}"#)
            },
            _ => String::new(),
        };
        let fallback_attributes = match options.mce_attributes {
            MceAttributeMode::QualifiedBranchAttributes => {
                format!(
                    r#" mc:Ignorable="{x15}" mc:MustUnderstand="{x15}" mc:ProcessContent="{x15}""#
                )
            },
            _ => String::new(),
        };
        let alternate_attributes = match options.mce_attributes {
            MceAttributeMode::ForeignChildAncestorIgnorable => {
                r#" mc:Ignorable="x15 foo""#.to_owned()
            },
            _ => String::new(),
        };
        let foreign_child = match options.mce_attributes {
            MceAttributeMode::ForeignChildOwnIgnorable => {
                r#"<foo:foreign mc:Ignorable="foo"/>"#.to_owned()
            },
            MceAttributeMode::ForeignChildAncestorIgnorable => r#"<foo:foreign/>"#.to_owned(),
            _ => String::new(),
        };
        let leaves = unique_name_leaves(&options);
        let payload = match options.payload {
            Payload::WrongQName => format!(
                r#"<{x15}:cachedUniqueNamesWrong>{leaves}</{x15}:cachedUniqueNamesWrong>"#,
                x15 = x15,
                leaves = leaves,
            ),
            Payload::WrongNamespace => {
                let leaves = leaves.replace(&format!("{x15}:"), &format!("{x14}:"));
                format!(
                    r#"<{x14}:cachedUniqueNames>{leaves}</{x14}:cachedUniqueNames>"#,
                    x14 = x14,
                    leaves = leaves,
                )
            },
            Payload::Valid => format!(
                r#"<{x15}:cachedUniqueNames>{leaves}</{x15}:cachedUniqueNames>"#,
                x15 = x15,
                leaves = leaves,
            ),
            Payload::Empty => format!(r#"<{x15}:cachedUniqueNames/>"#, x15 = x15),
            Payload::MceDirect => format!(
                r#"<mc:AlternateContent><{x15}:cachedUniqueNames>{leaves}</{x15}:cachedUniqueNames></mc:AlternateContent>"#,
                x15 = x15,
                leaves = leaves,
            ),
            Payload::MceNestedDirect => format!(
                r#"<mc:AlternateContent><mc:Choice Requires="{mce_requires}"><mc:AlternateContent><{x15}:cachedUniqueNames>{leaves}</{x15}:cachedUniqueNames></mc:AlternateContent></mc:Choice><mc:Fallback><{x15}:futureFallback/></mc:Fallback></mc:AlternateContent>"#,
                x15 = x15,
                mce_requires = mce_requires,
                leaves = leaves,
            ),
            Payload::MceNestedChoice => format!(
                r#"<mc:AlternateContent><mc:Choice Requires="{mce_requires}"><mc:AlternateContent><mc:Choice Requires="{mce_requires}"><{x15}:cachedUniqueNames>{leaves}</{x15}:cachedUniqueNames></mc:Choice><mc:Fallback><{x15}:futureFallback/></mc:Fallback></mc:AlternateContent></mc:Choice><mc:Fallback><{x15}:futureFallback/></mc:Fallback></mc:AlternateContent>"#,
                x15 = x15,
                mce_requires = mce_requires,
                leaves = leaves,
            ),
            Payload::Mce => format!(
                r#"<mc:AlternateContent{alternate_attributes}><mc:Choice Requires="{mce_requires}"{choice_attributes}><{x15}:cachedUniqueNames>{leaves}</{x15}:cachedUniqueNames></mc:Choice>{foreign_child}<mc:Fallback{fallback_attributes}><{x15}:futureFallback/></mc:Fallback></mc:AlternateContent>"#,
                x15 = x15,
                mce_requires = mce_requires,
                choice_attributes = choice_attributes,
                fallback_attributes = fallback_attributes,
                alternate_attributes = alternate_attributes,
                foreign_child = foreign_child,
                leaves = leaves,
            ),
            Payload::MceIdentical => format!(
                r#"<mc:AlternateContent{alternate_attributes}><mc:Choice Requires="{mce_requires}"{choice_attributes}><{x15}:cachedUniqueNames>{leaves}</{x15}:cachedUniqueNames></mc:Choice>{foreign_child}<mc:Fallback{fallback_attributes}><{x15}:cachedUniqueNames>{leaves}</{x15}:cachedUniqueNames></mc:Fallback></mc:AlternateContent>"#,
                x15 = x15,
                mce_requires = mce_requires,
                choice_attributes = choice_attributes,
                fallback_attributes = fallback_attributes,
                alternate_attributes = alternate_attributes,
                foreign_child = foreign_child,
                leaves = leaves,
            ),
            _ => format!(
                r#"<{x15}:cachedUniqueNames>{leaves}</{x15}:cachedUniqueNames>"#,
                x15 = x15,
                leaves = leaves,
            ),
        };
        let opaque = if options.include_opaque_sibling {
            format!(
                r#"<{root}ext uri="{{11111111-1111-1111-1111-111111111111}}"><{x15}:opaque keep="yes"/></{root}ext>"#,
                root = root,
                x15 = x15,
            )
        } else {
            String::new()
        };
        let ext = format!(
            r#"<{root}ext uri="{payload_uri}">{payload}</{root}ext>{opaque}"#,
            root = root,
            payload_uri = payload_uri,
            payload = payload,
            opaque = opaque,
        );
        let mce_decl = if matches!(
            options.payload,
            Payload::Mce
                | Payload::MceIdentical
                | Payload::MceDirect
                | Payload::MceNestedDirect
                | Payload::MceNestedChoice
        ) {
            let mce_uri = namespace_uri(MCE_NS, options.escaped_namespace_uris);
            let extra_namespace = match options.mce_requires {
                MceRequiresMode::SameUriAlias => {
                    format!(r#" xmlns:x15alias="{x15_uri}""#)
                },
                MceRequiresMode::XmlReservedPrefix => {
                    format!(r#" xmlns:xml="{XML_NS}""#)
                },
                MceRequiresMode::EncodedXmlReservedPrefix => {
                    format!(r#" xmlns:xml="{}""#, namespace_uri(XML_NS, true))
                },
                _ => String::new(),
            };
            let foreign_namespace = match options.mce_attributes {
                MceAttributeMode::ForeignBranchAttributeWithoutIgnorable
                | MceAttributeMode::ForeignBranchAttributeWithIgnorable
                | MceAttributeMode::ForeignChildOwnIgnorable
                | MceAttributeMode::ForeignChildAncestorIgnorable => {
                    format!(r#" xmlns:foo="{FOREIGN_NS}""#)
                },
                _ => String::new(),
            };
            let ignorable =
                if options.mce_attributes == MceAttributeMode::IgnorableNamesMceNamespace {
                    format!("{x15} mc")
                } else {
                    x15.to_owned()
                };
            format!(
                r#" xmlns:mc="{mce_uri}" mc:Ignorable="{ignorable}"{extra_namespace}{foreign_namespace}"#
            )
        } else {
            String::new()
        };
        let second_name = if options.duplicate_field_name {
            "Region"
        } else {
            "Amount"
        };
        let root_opaque = format!(
            r#"<{root}extLst><{root}ext uri="{pivot_cache_id_version_uri}">{pivot_cache_id_version}</{root}ext><{root}ext uri="{{22222222-2222-2222-2222-222222222222}}"><{x15}:cacheSibling keep="yes"/></{root}ext></{root}extLst>"#,
            root = root,
            x15 = x15,
            pivot_cache_id_version_uri = pivot_cache_id_version_uri,
            pivot_cache_id_version = match options.cache_id_version {
                CacheIdVersionMode::Valid => format!(
                    r#"<{x15}:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"/>"#,
                    x15 = x15,
                ),
                CacheIdVersionMode::Version255 => format!(
                    r#"<{x15}:pivotCacheIdVersion cacheIdSupportedVersion="255" cacheIdCreatedVersion="255"/>"#,
                    x15 = x15,
                ),
                CacheIdVersionMode::VersionPlus255 => format!(
                    r#"<{x15}:pivotCacheIdVersion cacheIdSupportedVersion="+255" cacheIdCreatedVersion="+255"/>"#,
                    x15 = x15,
                ),
                CacheIdVersionMode::Version256 => format!(
                    r#"<{x15}:pivotCacheIdVersion cacheIdSupportedVersion="256" cacheIdCreatedVersion="256"/>"#,
                    x15 = x15,
                ),
                CacheIdVersionMode::UnknownChild => format!(
                    r#"<{x15}:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"><{x15}:future/></{x15}:pivotCacheIdVersion>"#,
                    x15 = x15,
                ),
                CacheIdVersionMode::NonWhitespaceText => format!(
                    r#"<{x15}:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15">payload</{x15}:pivotCacheIdVersion>"#,
                    x15 = x15,
                ),
                CacheIdVersionMode::CData => format!(
                    r#"<{x15}:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"><![CDATA[payload]]></{x15}:pivotCacheIdVersion>"#,
                    x15 = x15,
                ),
                CacheIdVersionMode::UnknownAttribute => format!(
                    r#"<{x15}:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15" future="nope"/>"#,
                    x15 = x15,
                ),
            },
        );
        let cache_fields = format!(
            r#"<{root}cacheFields count="2"><{root}cacheField name="Region"><{root}sharedItems count="2"><{root}s v="North"/><{root}s v="South"/></{root}sharedItems><{root}extLst>{ext}</{root}extLst></{root}cacheField><{root}cacheField name="{second_name}"><{root}sharedItems count="0"/></{root}cacheField></{root}cacheFields>"#,
            root = root,
            second_name = second_name,
            ext = ext,
        );
        let cache_xml = format!(
            r#"<{root}pivotCacheDefinition {main_ns} xmlns:{x15}="{x15_uri}" xmlns:{x14}="{x14_uri}"{mce_decl}>{cache_source}{cache_fields}{root_opaque}</{root}pivotCacheDefinition>"#,
            root = root,
            main_ns = main_ns,
            x15 = x15,
            x14 = x14,
            x15_uri = x15_uri,
            x14_uri = x14_uri,
            mce_decl = mce_decl,
            cache_source = cache_source,
            cache_fields = cache_fields,
            root_opaque = root_opaque,
        );

        let (model_uri, model) = match options.connection {
            ConnectionMode::WrongModelUri => (
                wrong_uri.as_str(),
                format!(r#"<{x15}:connection model="true" id=""/>"#, x15 = x15),
            ),
            ConnectionMode::D799 => (
                d799_uri.as_str(),
                format!(r#"<{x15}:connection model="true" id=""/>"#, x15 = x15),
            ),
            ConnectionMode::ModelFalse => (
                de250_uri.as_str(),
                format!(r#"<{x15}:connection model="false" id=""/>"#, x15 = x15),
            ),
            ConnectionMode::NonEmptyModelId => (
                de250_uri.as_str(),
                format!(
                    r#"<{x15}:connection model="true" id="model-id"/>"#,
                    x15 = x15
                ),
            ),
            ConnectionMode::MissingModelId => (
                de250_uri.as_str(),
                format!(r#"<{x15}:connection model="true"/>"#, x15 = x15),
            ),
            ConnectionMode::Valid | ConnectionMode::WrongType => (
                de250_uri.as_str(),
                format!(r#"<{x15}:connection model="true" id=""/>"#, x15 = x15),
            ),
        };
        let standard_type = if options.connection == ConnectionMode::WrongType {
            "4"
        } else {
            "5"
        };
        let model_connection = format!(
            r#"<{root}extLst><{root}ext uri="{model_uri}">{model}</{root}ext></{root}extLst>"#,
            root = root,
            model_uri = model_uri,
            model = model,
        );
        let connections_xml = format!(
            r#"<{root}connections {main_ns} xmlns:{x15}="{x15_uri}"><{root}connection id="5" name="canonical" type="{standard_type}" refreshedVersion="7">{model_connection}</{root}connection><{root}connection id="6" name="other" type="5" refreshedVersion="7">{model_connection}</{root}connection></{root}connections>"#,
            root = root,
            main_ns = main_ns,
            x15 = x15,
            x15_uri = x15_uri,
            standard_type = standard_type,
            model_connection = model_connection,
        );

        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(WORKBOOK_URI).unwrap(),
                ct::SML_SHEET_MAIN.to_owned(),
                workbook_xml.as_bytes().to_vec(),
            )))
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(SHEET_URI).unwrap(),
                ct::SML_WORKSHEET.to_owned(),
                format!(r#"<worksheet xmlns="{main_uri}"><sheetData/></worksheet>"#).into_bytes(),
            )))
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(CACHE_URI).unwrap(),
                ct::SML_PIVOT_CACHE_DEFINITION.to_owned(),
                cache_xml.as_bytes().to_vec(),
            )))
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new(CONNECTIONS_URI).unwrap(),
                CONNECTIONS_CONTENT_TYPE.to_owned(),
                connections_xml.as_bytes().to_vec(),
            )))
            .unwrap();

        let workbook_uri = PackURI::new(WORKBOOK_URI).unwrap();
        let workbook = package.get_part_mut(&workbook_uri).unwrap();
        workbook
            .rels_mut()
            .try_add_relationship(
                options.dialect.worksheet_rel().to_owned(),
                "worksheets/sheet1.xml".to_owned(),
                "rIdSheet".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        workbook
            .rels_mut()
            .try_add_relationship(
                options.dialect.pivot_cache_rel().to_owned(),
                "pivotCache/pivotCacheDefinition1.xml".to_owned(),
                "rIdCache".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        workbook
            .rels_mut()
            .try_add_relationship(
                options.dialect.connections_rel().to_owned(),
                "connections.xml".to_owned(),
                "rIdConnections".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        package.relate_to("xl/workbook.xml", options.dialect.office_document_rel());

        if options.signed {
            package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
                    ct::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
                    b"<origin/>".to_vec(),
                )))
                .unwrap();
            package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        }

        Self {
            bytes: PackageWriter::to_bytes(&package).unwrap(),
            cache_xml,
            connections_xml,
        }
    }

    fn package(&self) -> OpcPackage {
        OpcPackage::from_bytes(&self.bytes).unwrap()
    }
}

fn unique_name_leaves(options: &FixtureOptions) -> String {
    match options.payload {
        Payload::MissingIndex => {
            r#"<x15:cachedUniqueName name="North"/>"#.replace("x15", options.x15_prefix)
        },
        Payload::MissingName => {
            r#"<x15:cachedUniqueName index="0"/>"#.replace("x15", options.x15_prefix)
        },
        Payload::DuplicateIndex => {
            format!(
                r#"<{x15}:cachedUniqueName index="0" name="North"/><{x15}:cachedUniqueName index="0" name="South"/>"#,
                x15 = options.x15_prefix
            )
        },
        Payload::PositiveIndex => {
            format!(
                r#"<{x15}:cachedUniqueName index="+1" name="North"/>"#,
                x15 = options.x15_prefix
            )
        },
        Payload::NegativeZeroIndex => {
            format!(
                r#"<{x15}:cachedUniqueName index="-0" name="North"/>"#,
                x15 = options.x15_prefix
            )
        },
        Payload::NegativeIndex => {
            format!(
                r#"<{x15}:cachedUniqueName index="-1" name="North"/>"#,
                x15 = options.x15_prefix
            )
        },
        Payload::OverflowIndex => {
            format!(
                r#"<{x15}:cachedUniqueName index="4294967296" name="North"/>"#,
                x15 = options.x15_prefix
            )
        },
        Payload::UnpairedSurrogate => {
            format!(
                r#"<{x15}:cachedUniqueName index="0" name="_xD800_"/>"#,
                x15 = options.x15_prefix
            )
        },
        Payload::UnknownUniqueNameAttribute => {
            format!(
                r#"<{x15}:cachedUniqueName index="0" name="North" future="nope"/>"#,
                x15 = options.x15_prefix
            )
        },
        Payload::UnknownUniqueNamesChild => {
            format!(
                r#"<{x15}:cachedUniqueName index="0" name="North"/><{x15}:futureCachedUniqueName/>"#,
                x15 = options.x15_prefix
            )
        },
        _ => options
            .names
            .iter()
            .map(|(index, name)| {
                format!(
                    r#"<{x15}:cachedUniqueName index="{index}" name="{name}"/>"#,
                    x15 = options.x15_prefix,
                )
            })
            .collect(),
    }
}

fn cache_selector() -> CacheSelector {
    CacheSelector::Id(PivotCacheId(CACHE_ID))
}

fn field_selector() -> FieldSelector<'static> {
    FieldSelector::Ordinal(0)
}

fn field_name_selector() -> FieldSelector<'static> {
    FieldSelector::Name("Region")
}

fn cache_blob(package: &OpcPackage) -> Vec<u8> {
    package
        .get_part(&PackURI::new(CACHE_URI).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn formatted_cache_package(fixture: &Fixture) -> OpcPackage {
    let mut package = fixture.package();
    let source = String::from_utf8(cache_blob(&package)).unwrap();
    let owner_open = format!(r#"<ext uri="{CACHED_UNIQUE_NAMES_URI}">"#);
    let owner_open_formatted =
        format!("\n  <!-- retained cached-name extension comment -->\n  {owner_open}");
    let version_open = format!(r#"<ext uri="{PIVOT_CACHE_ID_VERSION_URI}">"#);
    let version_open_formatted =
        format!("\n  <!-- retained pivot-cache extension comment -->\n  {version_open}");
    let formatted = source
        .replace(
            r#"<cacheFields count="2">"#,
            "<cacheFields count=\"2\">\n  <!-- retained field comment -->\n  ",
        )
        .replace(&version_open, &version_open_formatted)
        .replace(&owner_open, &owner_open_formatted)
        .replace(
            r#"<x15:cachedUniqueNames>"#,
            "<x15:cachedUniqueNames>\n  <!-- retained cached-name payload comment -->\n  ",
        )
        .replace(r#"</x15:cachedUniqueNames>"#, "\n</x15:cachedUniqueNames>")
        .replace(
            r#"<x15:cacheSibling keep="yes"/>"#,
            "<!-- retained opaque sibling comment -->\n  <x15:cacheSibling keep=\"yes\"/>",
        );
    package
        .get_part_mut(&PackURI::new(CACHE_URI).unwrap())
        .unwrap()
        .set_blob(formatted.into_bytes());
    package
}

fn cache_package_with_replacements(fixture: &Fixture, replacements: &[(&str, &str)]) -> OpcPackage {
    let mut package = fixture.package();
    let mut source = String::from_utf8(cache_blob(&package)).unwrap();
    for (from, to) in replacements {
        assert!(
            source.contains(from),
            "fixture source missing replacement: {from}"
        );
        source = source.replace(from, to);
    }
    package
        .get_part_mut(&PackURI::new(CACHE_URI).unwrap())
        .unwrap()
        .set_blob(source.into_bytes());
    package
}

fn cache_package_with_appended_owner(
    fixture: &Fixture,
    owner_uri: &str,
    duplicate: &str,
) -> OpcPackage {
    let mut package = fixture.package();
    let mut source = String::from_utf8(cache_blob(&package)).unwrap();
    let owner_open = format!(r#"<ext uri="{owner_uri}">"#);
    let owner_start = source
        .find(&owner_open)
        .unwrap_or_else(|| panic!("fixture source missing owner: {owner_uri}"));
    let owner_end = owner_start
        + source[owner_start..]
            .find("</ext>")
            .expect("fixture owner missing closing ext")
        + "</ext>".len();
    source.insert_str(owner_end, duplicate);
    package
        .get_part_mut(&PackURI::new(CACHE_URI).unwrap())
        .unwrap()
        .set_blob(source.into_bytes());
    package
}

fn cache_package_with_appended_payload(
    fixture: &Fixture,
    owner_uri: &str,
    duplicate_payload: &str,
) -> OpcPackage {
    let mut package = fixture.package();
    append_cache_payload(&mut package, owner_uri, duplicate_payload);
    package
}

fn append_cache_payload(package: &mut OpcPackage, owner_uri: &str, payload: &str) {
    let mut source = String::from_utf8(cache_blob(package)).unwrap();
    let owner_open = format!(r#"<ext uri="{owner_uri}">"#);
    let owner_start = source
        .find(&owner_open)
        .unwrap_or_else(|| panic!("fixture source missing owner: {owner_uri}"));
    let owner_end = owner_start
        + source[owner_start..]
            .find("</ext>")
            .expect("fixture owner missing closing ext");
    source.insert_str(owner_end, payload);
    package
        .get_part_mut(&PackURI::new(CACHE_URI).unwrap())
        .unwrap()
        .set_blob(source.into_bytes());
}

fn cache_package_with_root_extensions(fixture: &Fixture, extensions: &str) -> OpcPackage {
    let mut package = fixture.package();
    let mut source = String::from_utf8(cache_blob(&package)).unwrap();
    let root_close = "</extLst></pivotCacheDefinition>";
    let root_close_start = source
        .rfind(root_close)
        .expect("fixture root extension list missing");
    source.insert_str(root_close_start, extensions);
    package
        .get_part_mut(&PackURI::new(CACHE_URI).unwrap())
        .unwrap()
        .set_blob(source.into_bytes());
    package
}

fn oversized_extension(owner_uri: &str, payload: &str) -> String {
    format!(
        r#"<ext uri="{owner_uri}">{payload}<!--{}--></ext>"#,
        "x".repeat(1024 * 1024)
    )
}

fn oversized_payload(payload: &str) -> String {
    let close = payload
        .rfind("</")
        .expect("payload fixture must have an explicit closing tag");
    format!(
        "{}<!--{}-->{}",
        &payload[..close],
        "x".repeat(1024 * 1024),
        &payload[close..]
    )
}

fn assert_fragment_resource_limit<T>(result: Result<T, Error>, owner: &str) {
    match result {
        Err(Error::ResourceLimit(limit)) => {
            assert!(
                limit.observed > limit.limit,
                "{owner} refusal must report an observed size above its limit: {limit:?}"
            );
            assert_eq!(
                limit.limit,
                1024 * 1024,
                "{owner} refusal must carry the one MiB fragment limit: {limit:?}"
            );
            assert!(
                limit
                    .scope
                    .to_ascii_lowercase()
                    .contains(&owner.to_ascii_lowercase()),
                "{owner} refusal lost its owner scope: {limit:?}"
            );
        },
        Err(error) => {
            panic!("expected typed resource refusal for oversized {owner} extension, got {error:?}")
        },
        Ok(_) => panic!("accepted oversized {owner} extension"),
    }
}

fn connections_package_with_replacements(
    fixture: &Fixture,
    replacements: &[(&str, &str)],
) -> OpcPackage {
    let mut package = fixture.package();
    let mut source = String::from_utf8(connections_blob(&package)).unwrap();
    for (from, to) in replacements {
        assert!(
            source.contains(from),
            "connections source missing replacement: {from}"
        );
        source = source.replace(from, to);
    }
    package
        .get_part_mut(&PackURI::new(CONNECTIONS_URI).unwrap())
        .unwrap()
        .set_blob(source.into_bytes());
    package
}

fn connections_blob(package: &OpcPackage) -> Vec<u8> {
    package
        .get_part(&PackURI::new(CONNECTIONS_URI).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

#[test]
fn fixture_uses_exact_owner_and_unresolved_shared_items_anchor() {
    let fixture = Fixture::build(FixtureOptions::default());
    assert!(fixture.cache_xml.contains(CACHED_UNIQUE_NAMES_URI));
    assert!(fixture.cache_xml.contains(X15_NS));
    assert!(fixture.cache_xml.contains("<x15:cachedUniqueNames>"));
    assert!(fixture.cache_xml.contains("<cacheField name=\"Region\">"));
    assert!(
        fixture.cache_xml.contains("<sharedItems count=\"2\">")
            || fixture
                .cache_xml
                .contains("<cacheField name=\"Region\"><sharedItems")
    );
    assert!(!fixture.cache_xml.contains("<items"));
    assert!(!fixture.cache_xml.contains(":items"));
}

#[test]
fn reads_by_semantic_cache_id_and_field_ordinal_or_name_and_reports_unresolved_bound() {
    let fixture = Fixture::build(FixtureOptions::default());
    let package = fixture.package();
    let ordinal = Snapshot::load(&package, cache_selector(), field_selector()).unwrap();
    let by_name = Snapshot::load(&package, cache_selector(), field_name_selector()).unwrap();

    assert_eq!(ordinal.entries(), by_name.entries());
    assert_eq!(ordinal.entries().len(), 2);
    assert_eq!(ordinal.entries()[0].index, 0);
    assert_eq!(ordinal.entries()[0].name, "North & South");
    assert_eq!(ordinal.entries()[1].index, 7);
    assert_eq!(ordinal.entries()[1].name, "North & South");
    assert_eq!(ordinal.diagnostic_status(), DiagnosticStatus::Unresolved);
    assert!(Snapshot::load(&package, cache_selector(), FieldSelector::Ordinal(1)).is_err());
    assert!(
        Snapshot::load(
            &package,
            CacheSelector::Id(PivotCacheId(41)),
            field_selector()
        )
        .is_err()
    );
}

#[test]
fn strict_and_transitional_roots_and_prefixed_aliases_have_same_semantics() {
    let transitional = Fixture::build(FixtureOptions::default());
    let strict = Fixture::build(FixtureOptions {
        dialect: Dialect::Strict,
        root_prefix: "s",
        x15_prefix: "p",
        x14_prefix: "q",
        ..FixtureOptions::default()
    });
    let first =
        Snapshot::load(&transitional.package(), cache_selector(), field_selector()).unwrap();
    let second = Snapshot::load(&strict.package(), cache_selector(), field_selector()).unwrap();
    assert_eq!(first.entries(), second.entries());
    assert_eq!(first.diagnostic_status(), second.diagnostic_status());
}

#[test]
fn entity_escaped_namespace_uris_normalize_for_direct_and_mce_owners() {
    let direct = Fixture::build(FixtureOptions {
        escaped_namespace_uris: true,
        ..FixtureOptions::default()
    });
    assert!(direct.cache_xml.contains(&format!(
        r#"xmlns="{}""#,
        namespace_uri(TRANSITIONAL_MAIN, true)
    )));
    assert!(
        direct
            .cache_xml
            .contains(&format!(r#"xmlns:x14="{}""#, namespace_uri(X14_NS, true)))
    );
    assert!(
        direct
            .cache_xml
            .contains(&format!(r#"xmlns:x15="{}""#, namespace_uri(X15_NS, true)))
    );
    let direct_package = direct.package();
    let workbook_xml = String::from_utf8(
        direct_package
            .get_part(&PackURI::new(WORKBOOK_URI).unwrap())
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap();
    assert!(workbook_xml.contains(&format!(
        r#"xmlns:r="{}""#,
        namespace_uri(TRANSITIONAL_REL, true)
    )));
    let direct_snapshot =
        Snapshot::load(&direct_package, cache_selector(), field_selector()).unwrap();
    assert_eq!(direct_snapshot.entries()[0].name, "North & South");

    let mut edited_package = direct.package();
    let before = cache_blob(&edited_package);
    let mut transaction =
        Transaction::new(&mut edited_package, cache_selector(), field_selector()).unwrap();
    assert!(transaction.set_cached_unique_name(7, "Escaped").unwrap());
    let commit = transaction.commit().unwrap();
    assert_eq!(commit.snapshot().entries()[1].name, "Escaped");
    let inverse = commit.patch().inverse();
    inverse.apply(&mut edited_package).unwrap();
    assert_eq!(cache_blob(&edited_package), before);
    assert_eq!(
        Snapshot::load(&edited_package, cache_selector(), field_selector())
            .unwrap()
            .entries()[1]
            .name,
        "North & South"
    );

    let mce = Fixture::build(FixtureOptions {
        escaped_namespace_uris: true,
        payload: Payload::Mce,
        ..FixtureOptions::default()
    });
    assert!(mce.cache_xml.contains(r#"Requires="x15""#));
    let mce_snapshot = Snapshot::load(&mce.package(), cache_selector(), field_selector()).unwrap();
    assert!(mce_snapshot.has_ambiguous_mce_owner());
    assert_eq!(mce_snapshot.entries(), direct_snapshot.entries());
}

#[test]
fn extension_uri_token_padding_is_semantic_but_raw_source_bound() {
    for mode in [
        ExtensionUriMode::XmlWhitespacePadding,
        ExtensionUriMode::NumericWhitespacePadding,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            extension_uris: mode,
            ..FixtureOptions::default()
        });
        let mut package = fixture.package();
        let before_cache = cache_blob(&package);
        let before_connections = connections_blob(&package);
        let cache_source = String::from_utf8(before_cache.clone()).unwrap();
        for uri in [
            CACHED_UNIQUE_NAMES_URI,
            F057_URI,
            PIVOT_CACHE_ID_VERSION_URI,
        ] {
            let token = extension_uri(uri, mode);
            assert!(
                cache_source.contains(&format!(r#"uri="{token}""#)),
                "cache source lost padded token for {uri:?}"
            );
        }
        let connections_source = String::from_utf8(before_connections.clone()).unwrap();
        let model_token = extension_uri(DE250_URI, mode);
        assert!(connections_source.contains(&format!(r#"uri="{model_token}""#)));

        let snapshot = Snapshot::load(&package, cache_selector(), field_selector()).unwrap();
        assert_eq!(snapshot.entries()[0].name, "North & South");

        let no_op = Transaction::new(&mut package, cache_selector(), field_selector())
            .unwrap()
            .commit()
            .unwrap();
        assert!(!no_op.changed());
        assert_eq!(cache_blob(&package), before_cache);
        assert_eq!(connections_blob(&package), before_connections);

        let mut transaction =
            Transaction::new(&mut package, cache_selector(), field_selector()).unwrap();
        assert!(transaction.set_cached_unique_name(7, "Padded").unwrap());
        let commit = transaction.commit().unwrap();
        assert_eq!(commit.snapshot().entries()[1].name, "Padded");
        commit.patch().inverse().apply(&mut package).unwrap();
        assert_eq!(cache_blob(&package), before_cache);
        assert_eq!(connections_blob(&package), before_connections);
    }
}

#[test]
fn extension_uri_internal_whitespace_and_wrong_values_do_not_infer_by_qname() {
    let fixture = Fixture::build(FixtureOptions::default());
    let internal_cached = "{4F2E5C28 24EA-4EB8-9CBF-B6C8F9C3D259}";
    let old_cached = format!(r#"uri="{CACHED_UNIQUE_NAMES_URI}""#);
    let new_cached = format!(r#"uri="{internal_cached}""#);
    let package =
        cache_package_with_replacements(&fixture, &[(old_cached.as_str(), new_cached.as_str())]);
    assert!(Snapshot::load(&package, cache_selector(), field_selector()).is_err());

    let internal_f057 = "{F057638F-6D5F-4E77-A914-E7F072B9BCA8 }";
    let old_f057 = format!(r#"uri="{F057_URI}""#);
    let new_f057 = format!(r#"uri="{internal_f057}""#);
    let package =
        cache_package_with_replacements(&fixture, &[(old_f057.as_str(), new_f057.as_str())]);
    assert!(Snapshot::load(&package, cache_selector(), field_selector()).is_err());

    let internal_abf5 = "{ABF5C744-AB39-4b91-8756-CFA1BBC848D5 }";
    let old_abf5 = format!(r#"uri="{PIVOT_CACHE_ID_VERSION_URI}""#);
    let new_abf5 = format!(r#"uri="{internal_abf5}""#);
    let package =
        cache_package_with_replacements(&fixture, &[(old_abf5.as_str(), new_abf5.as_str())]);
    assert!(Snapshot::load(&package, cache_selector(), field_selector()).is_err());

    let internal_de250 = "{DE250136-89BD-433C-8126-D09CA5730AF9 }";
    let old_de250 = format!(r#"uri="{DE250_URI}""#);
    let new_de250 = format!(r#"uri="{internal_de250}""#);
    let package = connections_package_with_replacements(
        &fixture,
        &[(old_de250.as_str(), new_de250.as_str())],
    );
    assert!(Snapshot::load(&package, cache_selector(), field_selector()).is_err());

    let wrong = Fixture::build(FixtureOptions {
        payload: Payload::WrongUri,
        ..FixtureOptions::default()
    });
    assert!(Snapshot::load(&wrong.package(), cache_selector(), field_selector()).is_err());
}

#[test]
fn mce_requires_accepts_only_xml_whitespace_and_legal_prefix_aliases() {
    for mode in [
        MceRequiresMode::XmlWhitespace,
        MceRequiresMode::DuplicatePrefix,
        MceRequiresMode::SameUriAlias,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            payload: Payload::Mce,
            mce_requires: mode,
            ..FixtureOptions::default()
        });
        let package = fixture.package();
        let snapshot = Snapshot::load(&package, cache_selector(), field_selector());
        assert!(
            snapshot.is_ok(),
            "legal MCE Requires form was rejected: {mode:?}"
        );
        let snapshot = snapshot.unwrap();
        assert_eq!(snapshot.entries().len(), 2);
        assert!(snapshot.has_ambiguous_mce_owner());
    }

    for mode in [
        MceRequiresMode::Nbsp,
        MceRequiresMode::MceNamespace,
        MceRequiresMode::XmlReservedPrefix,
        MceRequiresMode::EncodedXmlReservedPrefix,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            payload: Payload::Mce,
            mce_requires: mode,
            ..FixtureOptions::default()
        });
        assert!(
            Snapshot::load(&fixture.package(), cache_selector(), field_selector()).is_err(),
            "invalid MCE Requires form was accepted: {mode:?}"
        );
    }
}

#[test]
fn mce_branch_attributes_require_qualified_or_effectively_ignorable_namespaces() {
    for mode in [
        MceAttributeMode::QualifiedBranchAttributes,
        MceAttributeMode::ForeignBranchAttributeWithIgnorable,
        MceAttributeMode::ForeignChildAncestorIgnorable,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            payload: Payload::Mce,
            mce_attributes: mode,
            ..FixtureOptions::default()
        });
        let snapshot = Snapshot::load(&fixture.package(), cache_selector(), field_selector());
        assert!(
            snapshot.is_ok(),
            "legal MCE attribute form was rejected: {mode:?}"
        );
        assert!(snapshot.unwrap().has_ambiguous_mce_owner());
    }

    for mode in [
        MceAttributeMode::UnqualifiedBranchAttribute,
        MceAttributeMode::ForeignBranchAttributeWithoutIgnorable,
        MceAttributeMode::ForeignChildOwnIgnorable,
        MceAttributeMode::IgnorableNamesMceNamespace,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            payload: Payload::Mce,
            mce_attributes: mode,
            ..FixtureOptions::default()
        });
        assert!(
            Snapshot::load(&fixture.package(), cache_selector(), field_selector()).is_err(),
            "invalid MCE attribute authorization was accepted: {mode:?}"
        );
    }
}

#[test]
fn opaque_cdata_comments_and_entity_encoded_xml_prefix_are_preserved() {
    let fixture = Fixture::build(FixtureOptions::default());
    let opaque_with_misc = cache_package_with_replacements(
        &fixture,
        &[(
            r#"<x15:opaque keep="yes"/>"#,
            r#"<x15:opaque keep="yes"><![CDATA[opaque <& content]]><!-- retained opaque comment --></x15:opaque>"#,
        )],
    );
    let before = cache_blob(&opaque_with_misc);
    let snapshot = Snapshot::load(&opaque_with_misc, cache_selector(), field_selector()).unwrap();
    assert_eq!(snapshot.entries()[0].name, "North & South");

    let mut edited = opaque_with_misc;
    let mut transaction =
        Transaction::new(&mut edited, cache_selector(), field_selector()).unwrap();
    assert!(transaction.set_cached_unique_name(0, "West").unwrap());
    transaction.commit().unwrap();
    let expected = String::from_utf8(before)
        .unwrap()
        .replace(
            r#"index="0" name="North &amp; South""#,
            r#"index="0" name="West""#,
        )
        .into_bytes();
    assert_eq!(cache_blob(&edited), expected);
    let edited_source = String::from_utf8(cache_blob(&edited)).unwrap();
    assert!(edited_source.contains(r#"<![CDATA[opaque <& content]]>"#));
    assert!(edited_source.contains("retained opaque comment"));

    let encoded_xml_prefix = cache_package_with_replacements(
        &fixture,
        &[(
            r#"<x15:opaque keep="yes"/>"#,
            r#"<x15:opaque keep="yes" xmlns:xml="http:&#47;&#47;www.w3.org/XML/1998/namespace"/>"#,
        )],
    );
    let snapshot = Snapshot::load(&encoded_xml_prefix, cache_selector(), field_selector());
    assert!(
        snapshot.is_ok(),
        "legal entity-encoded xmlns:xml was rejected: {snapshot:?}"
    );
}

#[test]
fn same_qname_under_unknown_or_inactive_mce_is_opaque_and_not_fragment_capped() {
    let marker = "opaque-cached-unique-names-marker";
    let unknown_payload = format!(
        r#"<x15:cachedUniqueNames><x15:cachedUniqueName index="99" name="{marker}"/><!--{}--></x15:cachedUniqueNames>"#,
        "u".repeat(1024 * 1024)
    );
    let unknown = cache_package_with_replacements(
        &Fixture::build(FixtureOptions::default()),
        &[(r#"<x15:opaque keep="yes"/>"#, unknown_payload.as_str())],
    );

    let inactive_mce = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="foo"><x15:cachedUniqueNames><x15:cachedUniqueName index="99" name="{marker}"/><!--{}--></x15:cachedUniqueNames></mc:Choice><mc:Fallback><x15:futureFallback/></mc:Fallback></mc:AlternateContent>"#,
        "i".repeat(1024 * 1024)
    );
    let mce_binding_from = format!(r#"xmlns:x14="{X14_NS}""#);
    let mce_binding_to =
        format!(r#"xmlns:x14="{X14_NS}" xmlns:mc="{MCE_NS}" xmlns:foo="urn:inactive""#);
    let inactive_fixture = Fixture::build(FixtureOptions::default());
    let inactive = cache_package_with_replacements(
        &inactive_fixture,
        &[
            (mce_binding_from.as_str(), mce_binding_to.as_str()),
            (r#"<x15:opaque keep="yes"/>"#, inactive_mce.as_str()),
        ],
    );

    for (label, mut package, expected_mce) in [
        ("unknown URI", unknown, false),
        ("inactive MCE", inactive, false),
    ] {
        let before = cache_blob(&package);
        let snapshot = Snapshot::load(&package, cache_selector(), field_selector())
            .unwrap_or_else(|error| panic!("{label} same-QName payload was inferred: {error:?}"));
        assert_eq!(
            snapshot.entries().len(),
            2,
            "{label} changed typed entry count"
        );
        assert_eq!(snapshot.has_ambiguous_mce_owner(), expected_mce);
        assert!(
            before
                .windows(marker.len())
                .any(|window| window == marker.as_bytes()),
            "{label} opaque payload marker was not retained"
        );
        assert_eq!(snapshot.source_xml(), before.as_slice());

        let commit = Transaction::new(&mut package, cache_selector(), field_selector())
            .unwrap()
            .commit()
            .unwrap();
        assert!(
            !commit.changed(),
            "{label} no-op unexpectedly changed source"
        );
        assert_eq!(
            cache_blob(&package),
            before,
            "{label} opaque source changed"
        );
    }
}

#[test]
fn malformed_opaque_qnames_text_and_reserved_namespace_bindings_are_rejected() {
    let fixture = Fixture::build(FixtureOptions::default());
    let cases = [
        (
            "element QName",
            r#"<x15:opaque keep="yes"/>"#,
            r#"<x15:opaque:broken keep="yes"/>"#,
        ),
        (
            "attribute QName",
            r#"<x15:opaque keep="yes"/>"#,
            r#"<x15:opaque bad:attribute="yes"/>"#,
        ),
        (
            "raw less-than attribute value",
            r#"<x15:opaque keep="yes"/>"#,
            r#"<x15:opaque keep="<raw"/>"#,
        ),
        (
            "raw CDATA terminator text",
            r#"<x15:opaque keep="yes"/>"#,
            r#"<x15:opaque keep="yes">]]></x15:opaque>"#,
        ),
        (
            "default XML namespace URI",
            r#"<x15:opaque keep="yes"/>"#,
            r#"<x15:opaque xmlns="http://www.w3.org/XML/1998/namespace" keep="yes"/>"#,
        ),
        (
            "entity-encoded default XML namespace URI",
            r#"<x15:opaque keep="yes"/>"#,
            r#"<x15:opaque xmlns="http:&#47;&#47;www.w3.org/XML/1998/namespace" keep="yes"/>"#,
        ),
        (
            "unused prefixed empty declaration",
            r#"<x15:opaque keep="yes"/>"#,
            r#"<x15:opaque xmlns:unused="" keep="yes"/>"#,
        ),
    ];
    for (label, from, to) in cases {
        let package = cache_package_with_replacements(&fixture, &[(from, to)]);
        assert!(
            Snapshot::load(&package, cache_selector(), field_selector()).is_err(),
            "accepted malformed opaque XML: {label}"
        );
    }
}

#[test]
fn model_connection_routes_require_f057_or_explicit_id_and_allow_both_when_consistent() {
    for route in [
        Route::F057Name,
        Route::ConnectionIdOnly,
        Route::BothConsistent,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            route,
            ..FixtureOptions::default()
        });
        assert!(Snapshot::load(&fixture.package(), cache_selector(), field_selector()).is_ok());
    }
    for route in [
        Route::BothMismatch,
        Route::F057WithZero,
        Route::Missing,
        Route::MissingSourceType,
        Route::DefaultConnectionId,
        Route::WorksheetConnectionId,
        Route::DuplicateF057,
        Route::UnresolvedF057,
        Route::MalformedF057,
        Route::WrongF057QName,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            route,
            ..FixtureOptions::default()
        });
        assert!(Snapshot::load(&fixture.package(), cache_selector(), field_selector()).is_err());
    }
}

#[test]
fn explicit_zero_connection_id_does_not_establish_the_model_owner() {
    let fixture = Fixture::build(FixtureOptions {
        route: Route::F057WithZero,
        ..FixtureOptions::default()
    });
    assert!(
        Snapshot::load(&fixture.package(), cache_selector(), field_selector()).is_err(),
        "an explicit connectionId=0 must not satisfy the owner route"
    );
}

#[test]
fn model_connection_requires_type_five_de250_true_and_empty_id() {
    for connection in [
        ConnectionMode::WrongType,
        ConnectionMode::ModelFalse,
        ConnectionMode::NonEmptyModelId,
        ConnectionMode::MissingModelId,
        ConnectionMode::WrongModelUri,
        ConnectionMode::D799,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            connection,
            ..FixtureOptions::default()
        });
        assert!(Snapshot::load(&fixture.package(), cache_selector(), field_selector()).is_err());
    }
}

#[test]
fn duplicate_names_are_legal_but_indices_must_be_unsigned_and_unique() {
    assert!(
        Snapshot::load(
            &Fixture::build(FixtureOptions::default()).package(),
            cache_selector(),
            field_selector()
        )
        .is_ok()
    );
    for payload in [
        Payload::DuplicateIndex,
        Payload::NegativeIndex,
        Payload::OverflowIndex,
        Payload::MissingIndex,
        Payload::MissingName,
        Payload::UnpairedSurrogate,
        Payload::WrongUri,
        Payload::WrongQName,
        Payload::WrongNamespace,
        Payload::Empty,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            payload,
            ..FixtureOptions::default()
        });
        assert!(Snapshot::load(&fixture.package(), cache_selector(), field_selector()).is_err());
    }
}

#[test]
fn unsigned_indices_accept_plus_one_and_negative_zero_lexicals_source_bound() {
    for (payload, semantic_index, raw_index) in [
        (Payload::PositiveIndex, 1_u32, "+1"),
        (Payload::NegativeZeroIndex, 0_u32, "-0"),
    ] {
        let fixture = Fixture::build(FixtureOptions {
            payload,
            ..FixtureOptions::default()
        });
        let mut package = fixture.package();
        let before = cache_blob(&package);
        let snapshot = Snapshot::load(&package, cache_selector(), field_selector()).unwrap();
        assert_eq!(snapshot.entries().len(), 1);
        assert_eq!(snapshot.entries()[0].index, semantic_index);
        assert!(
            String::from_utf8(before.clone())
                .unwrap()
                .contains(&format!(r#"index="{raw_index}""#))
        );

        let no_op = Transaction::new(&mut package, cache_selector(), field_selector())
            .unwrap()
            .commit()
            .unwrap();
        assert!(!no_op.changed());
        assert_eq!(cache_blob(&package), before);

        let mut transaction =
            Transaction::new(&mut package, cache_selector(), field_selector()).unwrap();
        assert!(
            transaction
                .set_cached_unique_name(semantic_index, "Edited")
                .unwrap()
        );
        let commit = transaction.commit().unwrap();
        assert_eq!(commit.snapshot().entries()[0].index, semantic_index);
        assert_eq!(commit.snapshot().entries()[0].name, "Edited");
        commit.patch().inverse().apply(&mut package).unwrap();
        assert_eq!(cache_blob(&package), before);
        let restored = Snapshot::load(&package, cache_selector(), field_selector()).unwrap();
        assert_eq!(restored.entries()[0].index, semantic_index);
    }
}

#[test]
fn cache_id_version_accepts_unsigned_byte_maximum_but_rejects_overflow() {
    let accepted = Fixture::build(FixtureOptions {
        cache_id_version: CacheIdVersionMode::Version255,
        ..FixtureOptions::default()
    });
    assert!(
        Snapshot::load(&accepted.package(), cache_selector(), field_selector()).is_ok(),
        "cacheId version 255 must be accepted"
    );

    let plus = Fixture::build(FixtureOptions {
        cache_id_version: CacheIdVersionMode::VersionPlus255,
        ..FixtureOptions::default()
    });
    let mut plus_package = plus.package();
    let plus_before = cache_blob(&plus_package);
    assert!(
        String::from_utf8(plus_before.clone())
            .unwrap()
            .contains(r#"cacheIdSupportedVersion="+255" cacheIdCreatedVersion="+255""#)
    );
    assert!(Snapshot::load(&plus_package, cache_selector(), field_selector()).is_ok());
    let plus_no_op = Transaction::new(&mut plus_package, cache_selector(), field_selector())
        .unwrap()
        .commit()
        .unwrap();
    assert!(!plus_no_op.changed());
    assert_eq!(cache_blob(&plus_package), plus_before);

    let overflow = Fixture::build(FixtureOptions {
        cache_id_version: CacheIdVersionMode::Version256,
        ..FixtureOptions::default()
    });
    assert!(
        Snapshot::load(&overflow.package(), cache_selector(), field_selector()).is_err(),
        "cacheId version 256 must overflow unsignedByte"
    );
}

#[test]
fn cached_unique_names_exact_grammar_rejects_unowned_mce_and_unknown_content() {
    for payload in [
        Payload::MceDirect,
        Payload::MceNestedDirect,
        Payload::UnknownUniqueNameAttribute,
        Payload::UnknownUniqueNamesChild,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            payload,
            ..FixtureOptions::default()
        });
        assert!(
            Snapshot::load(&fixture.package(), cache_selector(), field_selector()).is_err(),
            "accepted malformed cachedUniqueNames payload: {payload:?}"
        );
    }
}

#[test]
fn cached_unique_names_parent_sequence_requires_xml_whitespace_between_children() {
    let fixture = Fixture::build(FixtureOptions::default());
    let whitespace = cache_package_with_replacements(
        &fixture,
        &[(
            r#"name="North &amp; South"/><x15:cachedUniqueName"#,
            "name=\"North &amp; South\"/>\n\t  <x15:cachedUniqueName",
        )],
    );
    let snapshot = Snapshot::load(&whitespace, cache_selector(), field_selector())
        .expect("XML whitespace between cachedUniqueName children is legal");
    assert_eq!(snapshot.entries().len(), 2);

    for marker in ["payload", "<![CDATA[&#x20;]]>", "<![CDATA[&#x9;]]>"] {
        let inserted = format!("<x15:cachedUniqueNames>{marker}<x15:cachedUniqueName");
        let package = cache_package_with_replacements(
            &fixture,
            &[(
                "<x15:cachedUniqueNames><x15:cachedUniqueName",
                inserted.as_str(),
            )],
        );
        assert!(
            Snapshot::load(&package, cache_selector(), field_selector()).is_err(),
            "accepted non-whitespace cachedUniqueNames parent content: {marker}"
        );
    }
}

#[test]
fn cached_unique_name_leaves_are_attribute_only_and_empty_name_is_valid() {
    let empty_name = Fixture::build(FixtureOptions {
        names: vec![(0, String::new()), (7, "North &amp; South".to_owned())],
        ..FixtureOptions::default()
    });
    let snapshot = Snapshot::load(&empty_name.package(), cache_selector(), field_selector())
        .expect("an empty cachedUniqueName name attribute is valid");
    assert_eq!(snapshot.entries()[0].name, "");

    for marker in ["payload", "<![CDATA[payload]]>"] {
        let inserted = format!(
            r#"<x15:cachedUniqueName index="0" name="North &amp; South">{marker}</x15:cachedUniqueName>"#
        );
        let package = cache_package_with_replacements(
            &Fixture::build(FixtureOptions::default()),
            &[(
                r#"<x15:cachedUniqueName index="0" name="North &amp; South"/>"#,
                inserted.as_str(),
            )],
        );
        assert!(
            Snapshot::load(&package, cache_selector(), field_selector()).is_err(),
            "accepted cachedUniqueName leaf text: {marker}"
        );
    }
}

#[test]
fn oversized_duplicate_cache_extensions_are_refused_before_duplicate_diagnostics() {
    let base = Fixture::build(FixtureOptions::default());
    let d259 = oversized_extension(
        CACHED_UNIQUE_NAMES_URI,
        r#"<x15:cachedUniqueNames><x15:cachedUniqueName index="99" name="ignored"/></x15:cachedUniqueNames>"#,
    );
    let d259_payload = oversized_payload(
        r#"<x15:cachedUniqueNames><x15:cachedUniqueName index="99" name="ignored"/></x15:cachedUniqueNames>"#,
    );
    let abf5 = oversized_extension(
        PIVOT_CACHE_ID_VERSION_URI,
        r#"<x15:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"/>"#,
    );
    let abf5_payload = oversized_payload(
        r#"<x15:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"></x15:pivotCacheIdVersion>"#,
    );
    let f057 = oversized_extension(F057_URI, r#"<x14:sourceConnection name="canonical"/>"#);
    let f057_payload =
        oversized_payload(r#"<x14:sourceConnection name="canonical"></x14:sourceConnection>"#);
    let definition = format!(
        r#"<ext uri="{PIVOT_CACHE_DEFINITION_URI}"><x14:pivotCacheDefinition pivotCacheId="42"/></ext>{}"#,
        oversized_extension(
            PIVOT_CACHE_DEFINITION_URI,
            r#"<x14:pivotCacheDefinition pivotCacheId="42"/>"#,
        )
    );
    let definition_payload = format!(
        r#"<ext uri="{PIVOT_CACHE_DEFINITION_URI}"><x14:pivotCacheDefinition pivotCacheId="42"></x14:pivotCacheDefinition>{}</ext>"#,
        oversized_payload(
            r#"<x14:pivotCacheDefinition pivotCacheId="42"></x14:pivotCacheDefinition>"#,
        ),
    );

    let cases = [
        (
            "cachedUniqueNames",
            cache_package_with_appended_owner(&base, CACHED_UNIQUE_NAMES_URI, &d259),
        ),
        (
            "cachedUniqueNames",
            cache_package_with_appended_payload(&base, CACHED_UNIQUE_NAMES_URI, &d259_payload),
        ),
        (
            "pivotCacheIdVersion",
            cache_package_with_appended_owner(&base, PIVOT_CACHE_ID_VERSION_URI, &abf5),
        ),
        (
            "pivotCacheIdVersion",
            cache_package_with_appended_payload(&base, PIVOT_CACHE_ID_VERSION_URI, &abf5_payload),
        ),
        (
            "F057",
            cache_package_with_appended_owner(&base, F057_URI, &f057),
        ),
        (
            "F057",
            cache_package_with_appended_payload(&base, F057_URI, &f057_payload),
        ),
        (
            "pivotCacheDefinition",
            cache_package_with_root_extensions(&base, &definition),
        ),
        (
            "pivotCacheDefinition",
            cache_package_with_root_extensions(&base, &definition_payload),
        ),
    ];

    for (owner, package) in cases {
        assert_fragment_resource_limit(
            Snapshot::load(&package, cache_selector(), field_selector()),
            owner,
        );
    }
}

#[test]
fn pivot_cache_id_version_requires_exact_attributes_and_payload_grammar() {
    for cache_id_version in [
        CacheIdVersionMode::UnknownChild,
        CacheIdVersionMode::NonWhitespaceText,
        CacheIdVersionMode::CData,
        CacheIdVersionMode::UnknownAttribute,
    ] {
        let fixture = Fixture::build(FixtureOptions {
            cache_id_version,
            ..FixtureOptions::default()
        });
        assert!(
            Snapshot::load(&fixture.package(), cache_selector(), field_selector()).is_err(),
            "accepted malformed pivotCacheIdVersion payload: {cache_id_version:?}"
        );
    }
}

#[test]
fn entities_and_spreadsheet_escape_suppression_are_decoded_as_xstring() {
    let fixture = Fixture::build(FixtureOptions {
        names: vec![
            (0, "A&amp;B &quot;quoted&quot;".to_owned()),
            (1, "_x005F_x0041_".to_owned()),
        ],
        ..FixtureOptions::default()
    });
    let snapshot = Snapshot::load(&fixture.package(), cache_selector(), field_selector()).unwrap();
    assert_eq!(snapshot.entries()[0].name, "A&B \"quoted\"");
    assert_eq!(snapshot.entries()[1].name, "_x0041_");
}

#[test]
fn scalar_edit_round_trips_decoded_control_as_spreadsheet_escape() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    let mut transaction =
        Transaction::new(&mut package, cache_selector(), field_selector()).unwrap();
    assert!(transaction.set_cached_unique_name(0, "\u{1}").unwrap());
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());

    let source = String::from_utf8(cache_blob(&package)).unwrap();
    assert!(source.contains(r#"name="_x0001_""#));
    let reopened = Snapshot::load(&package, cache_selector(), field_selector()).unwrap();
    assert_eq!(reopened.entries()[0].name, "\u{1}");
    assert_eq!(reopened.entries()[1].name, "North & South");
}

#[test]
fn decoded_utf16_name_limit_counts_supplementary_scalars_and_is_atomic() {
    let at_limit = "a".repeat(65_535);
    let fixture = Fixture::build(FixtureOptions {
        names: vec![(0, at_limit), (1, "other".to_owned())],
        ..FixtureOptions::default()
    });
    let limits = ReadLimits::builder()
        .max_xml_attribute_bytes(256 * 1024)
        .unwrap()
        .build()
        .unwrap();
    let package = OpcPackage::from_bytes_with_limits(&fixture.bytes, limits).unwrap();
    assert_eq!(
        Snapshot::load(&package, cache_selector(), field_selector())
            .unwrap()
            .entries()[0]
            .name
            .encode_utf16()
            .count(),
        65_535
    );

    let too_long = Fixture::build(FixtureOptions {
        names: vec![(0, "a".repeat(65_536)), (1, "other".to_owned())],
        ..FixtureOptions::default()
    });
    let package = OpcPackage::from_bytes_with_limits(&too_long.bytes, limits).unwrap();
    assert!(Snapshot::load(&package, cache_selector(), field_selector()).is_err());

    let supplementary = Fixture::build(FixtureOptions {
        names: vec![(0, "😀".repeat(32_767) + "a"), (1, "other".to_owned())],
        ..FixtureOptions::default()
    });
    let package = OpcPackage::from_bytes_with_limits(&supplementary.bytes, limits).unwrap();
    let supplementary_result = Snapshot::load(&package, cache_selector(), field_selector());
    assert!(supplementary_result.is_ok(), "{supplementary_result:?}");
    let over_supplementary = Fixture::build(FixtureOptions {
        names: vec![(0, "😀".repeat(32_768)), (1, "other".to_owned())],
        ..FixtureOptions::default()
    });
    let package = OpcPackage::from_bytes_with_limits(&over_supplementary.bytes, limits).unwrap();
    assert!(Snapshot::load(&package, cache_selector(), field_selector()).is_err());
}

#[test]
fn scalar_edit_is_existing_name_only_and_preserves_source_geometry() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = formatted_cache_package(&fixture);
    let before = cache_blob(&package);
    let mut transaction =
        Transaction::new(&mut package, cache_selector(), field_selector()).unwrap();
    assert!(transaction.set_cached_unique_name(0, "West").unwrap());
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.snapshot().entries()[0].name, "West");
    assert!(!commit.patch().is_empty());
    let after = cache_blob(&package);
    let expected = String::from_utf8(before.clone())
        .unwrap()
        .replace(
            r#"index="0" name="North &amp; South""#,
            r#"index="0" name="West""#,
        )
        .into_bytes();
    assert_eq!(after, expected);
    let after_text = String::from_utf8(after.clone()).unwrap();
    assert!(after_text.contains("cacheSibling keep=\"yes\""));
    assert!(after_text.contains("opaque keep=\"yes\""));
    assert!(after_text.contains("retained pivot-cache extension comment"));
    assert!(after_text.contains("retained cached-name extension comment"));
    assert!(after_text.contains("retained cached-name payload comment"));
    assert!(after_text.contains(&format!(r#"xmlns:x15="{X15_NS}""#)));
    assert!(after_text.contains(&format!(r#"xmlns:x14="{X14_NS}""#)));
    assert_eq!(
        connections_blob(&package),
        fixture.connections_xml.as_bytes()
    );
    let reopened = Snapshot::load(&package, cache_selector(), field_selector()).unwrap();
    assert_eq!(reopened.entries()[0].name, "West");
}

#[test]
fn no_op_and_same_value_edits_retain_exact_cache_source_bytes() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    let before = cache_blob(&package);
    let commit = Transaction::new(&mut package, cache_selector(), field_selector())
        .unwrap()
        .commit()
        .unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(cache_blob(&package), before);

    let mut package = fixture.package();
    let before = cache_blob(&package);
    let mut transaction =
        Transaction::new(&mut package, cache_selector(), field_selector()).unwrap();
    assert!(
        !transaction
            .set_cached_unique_name(0, "North & South")
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(cache_blob(&package), before);
}

#[test]
fn patch_inverse_restores_exact_bytes_and_stale_or_signed_targets_are_atomic() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut source = fixture.package();
    let before = cache_blob(&source);
    let mut transaction =
        Transaction::new(&mut source, cache_selector(), field_selector()).unwrap();
    assert!(transaction.set_cached_unique_name(7, "South").unwrap());
    let patch = transaction.commit().unwrap().patch().clone();
    let mut target = fixture.package();
    assert!(patch.apply(&mut target).is_ok());
    assert_eq!(
        Snapshot::load(&target, cache_selector(), field_selector())
            .unwrap()
            .entries()[1]
            .name,
        "South"
    );
    assert!(patch.inverse().apply(&mut source).is_ok());
    assert_eq!(cache_blob(&source), before);

    let mut stale = fixture.package();
    let stale_before = cache_blob(&stale);
    let changed = String::from_utf8(stale_before.clone())
        .unwrap()
        .replace("cacheSibling keep=\"yes\"", "cacheSibling keep=\"stale\"");
    stale
        .get_part_mut(&PackURI::new(CACHE_URI).unwrap())
        .unwrap()
        .set_blob(changed.into_bytes());
    let stale_before_apply = cache_blob(&stale);
    assert!(patch.apply(&mut stale).is_err());
    assert_eq!(cache_blob(&stale), stale_before_apply);

    let mut stale_connections = fixture.package();
    let stale_connection_xml = String::from_utf8(connections_blob(&stale_connections))
        .unwrap()
        .replace(
            r#"name="canonical""#,
            r#"name="stale" description="changed""#,
        );
    stale_connections
        .get_part_mut(&PackURI::new(CONNECTIONS_URI).unwrap())
        .unwrap()
        .set_blob(stale_connection_xml.into_bytes());
    let stale_connections_before_apply = cache_blob(&stale_connections);
    assert!(patch.apply(&mut stale_connections).is_err());
    assert_eq!(
        cache_blob(&stale_connections),
        stale_connections_before_apply
    );

    let mut stale_workbook = fixture.package();
    let stale_workbook_xml = String::from_utf8(
        stale_workbook
            .get_part(&PackURI::new(WORKBOOK_URI).unwrap())
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap()
    .replace("<workbook", "<workbook stale=\"yes\"");
    stale_workbook
        .get_part_mut(&PackURI::new(WORKBOOK_URI).unwrap())
        .unwrap()
        .set_blob(stale_workbook_xml.into_bytes());
    let stale_workbook_before_apply = cache_blob(&stale_workbook);
    assert!(patch.apply(&mut stale_workbook).is_err());
    assert_eq!(cache_blob(&stale_workbook), stale_workbook_before_apply);

    let signed = Fixture::build(FixtureOptions {
        signed: true,
        ..FixtureOptions::default()
    });
    let mut signed_package = signed.package();
    let signed_before = cache_blob(&signed_package);
    let mut signed_transaction =
        Transaction::new(&mut signed_package, cache_selector(), field_selector()).unwrap();
    signed_transaction
        .set_cached_unique_name(0, "Signed edit")
        .unwrap();
    assert!(signed_transaction.commit().is_err());
    assert_eq!(cache_blob(&signed_package), signed_before);
}

#[test]
fn existing_name_edits_allow_empty_xstring_but_reject_structure_and_unknown_index() {
    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    let original_cache = cache_blob(&package);
    let mut transaction =
        Transaction::new(&mut package, cache_selector(), field_selector()).unwrap();
    assert!(transaction.set_cached_unique_name(0, "").unwrap());
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let inverse = commit.patch().inverse();
    let reopened = Snapshot::load(&package, cache_selector(), field_selector()).unwrap();
    assert_eq!(reopened.entries()[0].name, "");
    assert_eq!(reopened.entries()[1].name, "North & South");
    assert_ne!(cache_blob(&package), original_cache);
    inverse.apply(&mut package).unwrap();
    assert_eq!(cache_blob(&package), original_cache);
    let reverted = Snapshot::load(&package, cache_selector(), field_selector()).unwrap();
    assert_eq!(reverted.entries()[0].name, "North & South");

    let fixture = Fixture::build(FixtureOptions::default());
    let mut package = fixture.package();
    let before = cache_blob(&package);
    let mut transaction =
        Transaction::new(&mut package, cache_selector(), field_selector()).unwrap();
    assert!(transaction.set_cached_unique_name(999, "invented").is_err());
    assert!(
        transaction
            .insert_cached_unique_name(
                0,
                CachedUniqueName {
                    index: 999,
                    name: "invented".to_owned(),
                },
            )
            .is_err()
    );
    assert!(transaction.remove_cached_unique_name(0).is_err());
    assert!(transaction.set_cached_unique_name_index(0, 1).is_err());
    assert_eq!(cache_blob(&package), before);
}

#[test]
fn caller_limits_reject_large_source_or_candidate_before_publication() {
    let fixture = Fixture::build(FixtureOptions {
        names: vec![(0, "l".repeat(20_000)), (1, "other".to_owned())],
        ..FixtureOptions::default()
    });
    let tiny_attribute_limit = ReadLimits::builder()
        .max_xml_attribute_bytes(4096)
        .unwrap()
        .build()
        .unwrap();
    let tiny_package =
        OpcPackage::from_bytes_with_limits(&fixture.bytes, tiny_attribute_limit).unwrap();
    assert!(Snapshot::load(&tiny_package, cache_selector(), field_selector()).is_err());

    let generous_limits = ReadLimits::builder()
        .max_xml_attribute_bytes(256 * 1024)
        .unwrap()
        .build()
        .unwrap();
    let mut package = OpcPackage::from_bytes_with_limits(&fixture.bytes, generous_limits).unwrap();
    let before = cache_blob(&package);
    let mut transaction =
        Transaction::new(&mut package, cache_selector(), field_selector()).unwrap();
    assert!(
        transaction
            .set_cached_unique_name(0, "&".repeat(60_000))
            .is_err()
    );
    assert_eq!(cache_blob(&package), before);
}

#[test]
fn mce_selected_branch_is_readable_but_ambiguous_scalar_publication_is_refused() {
    let fixture = Fixture::build(FixtureOptions {
        payload: Payload::Mce,
        ..FixtureOptions::default()
    });
    let package = fixture.package();
    let mce_result = Snapshot::load(&package, cache_selector(), field_selector());
    assert!(mce_result.is_ok(), "{mce_result:?}");
    assert!(mce_result.unwrap().has_ambiguous_mce_owner());
    let mut package = fixture.package();
    let before = cache_blob(&package);
    if let Ok(mut transaction) = Transaction::new(&mut package, cache_selector(), field_selector())
    {
        assert!(transaction.set_cached_unique_name(0, "Changed").is_ok());
        assert!(transaction.commit().is_err());
    }
    assert_eq!(cache_blob(&package), before);
}

#[test]
fn mce_inactive_fallback_with_identical_payload_is_not_a_duplicate() {
    let fixture = Fixture::build(FixtureOptions {
        payload: Payload::MceIdentical,
        ..FixtureOptions::default()
    });
    let package = fixture.package();
    let snapshot = Snapshot::load(&package, cache_selector(), field_selector()).unwrap();
    assert_eq!(snapshot.entries().len(), 2);
    assert!(snapshot.has_ambiguous_mce_owner());

    let mut package = fixture.package();
    let before = cache_blob(&package);
    if let Ok(mut transaction) = Transaction::new(&mut package, cache_selector(), field_selector())
    {
        assert!(transaction.set_cached_unique_name(0, "Changed").is_ok());
        assert!(transaction.commit().is_err());
    }
    assert_eq!(cache_blob(&package), before);
}

#[test]
fn mce_nested_selected_choice_is_readable_but_edit_refused() {
    let fixture = Fixture::build(FixtureOptions {
        payload: Payload::MceNestedChoice,
        ..FixtureOptions::default()
    });
    let package = fixture.package();
    let snapshot = Snapshot::load(&package, cache_selector(), field_selector()).unwrap();
    assert_eq!(snapshot.entries().len(), 2);
    assert!(snapshot.has_ambiguous_mce_owner());

    let mut package = fixture.package();
    let before = cache_blob(&package);
    if let Ok(mut transaction) = Transaction::new(&mut package, cache_selector(), field_selector())
    {
        assert!(transaction.set_cached_unique_name(0, "Changed").is_ok());
        assert!(transaction.commit().is_err());
    }
    assert_eq!(cache_blob(&package), before);
}

#[test]
fn duplicate_field_names_are_not_a_semantic_name_selector() {
    let fixture = Fixture::build(FixtureOptions {
        duplicate_field_name: true,
        ..FixtureOptions::default()
    });
    let package = fixture.package();
    assert!(Snapshot::load(&package, cache_selector(), field_selector()).is_ok());
    assert!(Snapshot::load(&package, cache_selector(), field_name_selector()).is_err());
}

#[test]
fn ordinary_workbook_facade_reads_semantic_cache_and_field_selectors() {
    let fixture = Fixture::build(FixtureOptions::default());
    let workbook = Workbook::from_bytes(fixture.bytes.clone()).unwrap();
    let by_ordinal = workbook
        .pivot_cache(cache_selector())
        .field(field_selector())
        .cached_unique_names()
        .unwrap();
    let by_name = workbook
        .pivot_cache(cache_selector())
        .field(field_name_selector())
        .cached_unique_names()
        .unwrap();

    assert_eq!(by_ordinal.cache_id(), PivotCacheId(CACHE_ID));
    assert_eq!(by_ordinal.field_ordinal(), 0);
    assert_eq!(by_ordinal.field_name(), "Region");
    assert_eq!(by_ordinal.entries(), by_name.entries());
    assert_eq!(by_ordinal.entries()[0].name, "North & South");
    assert_eq!(by_ordinal.entries()[1].index, 7);
    assert_eq!(by_ordinal.diagnostic_status(), DiagnosticStatus::Unresolved);

    let source = workbook
        .pivot_cache(cache_selector())
        .field(field_selector())
        .cached_unique_names_source()
        .unwrap();
    assert_eq!(source.entries(), by_ordinal.entries());
}

#[test]
fn ordinary_workbook_edit_returns_new_snapshot_preserves_original_and_reopens() {
    let fixture = Fixture::build(FixtureOptions::default());
    let original = Workbook::from_bytes(fixture.bytes.clone()).unwrap();
    let original_bytes = original.to_bytes().unwrap();
    let mut edit = original
        .edit_pivot_cache(cache_selector())
        .field(field_selector())
        .unwrap();
    assert!(edit.set_cached_unique_name(7, "South").unwrap());
    let committed = edit.commit().unwrap();
    assert!(committed.changed());
    assert_eq!(committed.snapshot().entries()[1].name, "South");

    assert_eq!(original.to_bytes().unwrap(), original_bytes);
    assert_eq!(
        original
            .pivot_cache(cache_selector())
            .field(field_selector())
            .cached_unique_names()
            .unwrap()
            .entries()[1]
            .name,
        "North & South"
    );
    let edited = committed.workbook();
    let edited_bytes = edited.to_bytes().unwrap();
    assert_ne!(edited_bytes, original_bytes);
    let reopened = Workbook::from_bytes(edited_bytes).unwrap();
    assert_eq!(
        reopened
            .pivot_cache(cache_selector())
            .field(field_selector())
            .cached_unique_names()
            .unwrap()
            .entries()[1]
            .name,
        "South"
    );
}

#[test]
fn ordinary_workbook_no_op_preserves_bytes_and_reports_empty_patch() {
    let fixture = Fixture::build(FixtureOptions::default());
    let original = Workbook::from_bytes(fixture.bytes).unwrap();
    let before = original.to_bytes().unwrap();
    let mut edit = original
        .edit_pivot_cache(cache_selector())
        .field(field_selector())
        .unwrap();
    assert!(!edit.set_cached_unique_name(0, "North & South").unwrap());
    let committed = edit.commit().unwrap();

    assert!(!committed.changed());
    assert!(committed.patch().is_empty());
    assert_eq!(committed.workbook().to_bytes().unwrap(), before);
}

#[test]
fn ordinary_workbook_patch_is_fresh_source_bound_stale_atomic_and_exactly_reversible() {
    let fixture = Fixture::build(FixtureOptions::default());
    let base = Workbook::from_bytes(fixture.bytes.clone()).unwrap();
    let base_bytes = base.to_bytes().unwrap();
    let mut edit = base
        .edit_pivot_cache(cache_selector())
        .field(field_selector())
        .unwrap();
    assert!(edit.set_cached_unique_name(7, "South").unwrap());
    let committed = edit.commit().unwrap();
    let patch = committed.patch().clone();
    assert!(!patch.is_empty());

    let fresh = Workbook::from_bytes(fixture.bytes.clone()).unwrap();
    let forwarded = patch.apply(&fresh).unwrap();
    assert!(forwarded.changed());
    assert_eq!(
        forwarded.workbook().to_bytes().unwrap(),
        committed.workbook().to_bytes().unwrap()
    );

    let mut stale_package = fixture.package();
    let stale_cache_xml = String::from_utf8(cache_blob(&stale_package))
        .unwrap()
        .replace("cacheSibling keep=\"yes\"", "cacheSibling keep=\"stale\"");
    stale_package
        .get_part_mut(&PackURI::new(CACHE_URI).unwrap())
        .unwrap()
        .set_blob(stale_cache_xml.into_bytes());
    let stale = Workbook::from_bytes(PackageWriter::to_bytes(&stale_package).unwrap()).unwrap();
    let stale_before = stale.to_bytes().unwrap();
    assert!(patch.apply(&stale).is_err());
    assert_eq!(stale.to_bytes().unwrap(), stale_before);

    let restored = patch.inverse().apply(committed.workbook()).unwrap();
    let restored_bytes = restored.workbook().to_bytes().unwrap();
    assert_eq!(restored_bytes, base_bytes);
    let base_package = OpcPackage::from_bytes(&base_bytes).unwrap();
    let restored_package = OpcPackage::from_bytes(&restored_bytes).unwrap();
    assert_eq!(
        base_package
            .get_part(&PackURI::new(WORKBOOK_URI).unwrap())
            .unwrap()
            .blob(),
        restored_package
            .get_part(&PackURI::new(WORKBOOK_URI).unwrap())
            .unwrap()
            .blob()
    );
    assert_eq!(cache_blob(&base_package), cache_blob(&restored_package));
    assert_eq!(
        connections_blob(&base_package),
        connections_blob(&restored_package)
    );
}
