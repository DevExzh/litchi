//! Glossary XML, semantic, and package boundary validation.

use super::super::graph::{is_reserved_physical_part, is_signature_part};
use super::super::model::*;
use super::super::*;
use super::super::{MAX_GRAPH_BYTES, MAX_NAME_KEY, MAX_PARTS, MAX_STRING, R, RS, VML, W, WS};
use super::xml::{Content, Node};

pub(in crate::glossary) fn validate_catalog_fields(v: &Catalog) -> Result<()> {
    if v.entries.len() > MAX_PARTS {
        return Err(invalid("glossary entry limit exceeded"));
    }
    for e in &v.entries {
        validate_entry_fields(e)?;
    }
    Ok(())
}

pub(in crate::glossary) fn validate_entry_fields(entry: &Entry) -> Result<()> {
    let Some(props) = &entry.props else {
        return Ok(());
    };
    if let Some(name) = &props.name {
        bounded(name.as_str())?;
    }
    for value in [
        props.style.as_deref(),
        props.category.as_ref().map(Category::name),
        props.description.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        bounded(value)?;
    }
    if !Kind::all().contains(props.kinds) || !Insert::all().contains(props.inserts) {
        return Err(invalid("unknown glossary option flag"));
    }
    if props.all_kinds.is_some() && props.kinds.is_empty() {
        return Err(invalid(
            "document-part types requires at least one kind when present",
        ));
    }
    Ok(())
}

pub(in crate::glossary) fn validate_name(value: &str) -> Result<()> {
    bounded(value)?;
    if value.trim().is_empty() {
        Err(invalid("building-block name cannot be empty"))
    } else {
        Ok(())
    }
}

pub(in crate::glossary) fn validate_raw_part(
    name: &str,
    content_type: &str,
    len: usize,
) -> Result<()> {
    let uri = validate_physical_part(name, content_type, len)?;
    if uri.as_str() == "/word/glossary/document.xml" {
        return Err(invalid(
            "glossary auxiliary part conflicts with the default root",
        ));
    }
    Ok(())
}

pub(in crate::glossary) fn validate_physical_part(
    name: &str,
    content_type: &str,
    len: usize,
) -> Result<PackURI> {
    let uri = PackURI::new(name).map_err(Error::Uri)?;
    if is_signature_part(&uri) || is_reserved_physical_part(&uri) {
        return Err(invalid(format!(
            "'{}' is reserved OPC package infrastructure",
            uri.as_str()
        )));
    }
    ContentType::new(content_type.to_owned())?;
    let media_type = content_type.split(';').next().unwrap_or_default().trim();
    if [
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
        ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE,
        ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE,
        ct::OPC_RELATIONSHIPS,
    ]
    .iter()
    .any(|reserved| media_type.eq_ignore_ascii_case(reserved))
    {
        return Err(invalid(
            "glossary parts cannot use reserved OPC infrastructure content types",
        ));
    }
    if len > MAX_GRAPH_BYTES {
        return Err(invalid("glossary auxiliary part exceeds 256 MiB"));
    }
    Ok(uri)
}

pub(in crate::glossary) fn name_key(value: &str) -> Result<String> {
    bounded(value)?;
    let mut key = String::new();
    key.try_reserve(value.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary semantic-name key",
            source,
        })?;
    for character in value.chars().nfd().default_case_fold().nfd() {
        let len = key
            .len()
            .checked_add(character.len_utf8())
            .ok_or_else(|| invalid("glossary semantic-name key size overflow"))?;
        if len > MAX_NAME_KEY {
            return Err(invalid(
                "glossary semantic-name key exceeds the normalized size limit",
            ));
        }
        if key.capacity() < len {
            key.try_reserve(character.len_utf8())
                .map_err(|source| Error::Allocation {
                    resource: "glossary semantic-name key",
                    source,
                })?;
        }
        key.push(character);
    }
    Ok(key)
}

pub(in crate::glossary) fn canonical_id(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() == 38
        && bytes.first() == Some(&b'{')
        && bytes.last() == Some(&b'}')
        && bytes.iter().enumerate().all(|(index, byte)| match index {
            0 => *byte == b'{',
            9 | 14 | 19 | 24 => *byte == b'-',
            37 => *byte == b'}',
            _ => byte.is_ascii_digit() || (b'A'..=b'F').contains(byte),
        })
}

pub(in crate::glossary) fn valid_ncname(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_ncname_start) && characters.all(is_ncname_char)
}

pub(in crate::glossary) fn is_ncname_start(character: char) -> bool {
    matches!(
        character,
        'A'..='Z'
            | '_'
            | 'a'..='z'
            | '\u{00c0}'..='\u{00d6}'
            | '\u{00d8}'..='\u{00f6}'
            | '\u{00f8}'..='\u{02ff}'
            | '\u{0370}'..='\u{037d}'
            | '\u{037f}'..='\u{1fff}'
            | '\u{200c}'..='\u{200d}'
            | '\u{2070}'..='\u{218f}'
            | '\u{2c00}'..='\u{2fef}'
            | '\u{3001}'..='\u{d7ff}'
            | '\u{f900}'..='\u{fdcf}'
            | '\u{fdf0}'..='\u{fffd}'
            | '\u{10000}'..='\u{effff}'
    )
}

pub(in crate::glossary) fn is_ncname_char(character: char) -> bool {
    is_ncname_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{00b7}' | '\u{0300}'..='\u{036f}' | '\u{203f}'..='\u{2040}'
        )
}

pub(in crate::glossary) fn kids(n: &Node) -> Result<Vec<&Node>> {
    let mut v = Vec::new();
    for c in &n.content {
        match c {
            Content::Node(x) => v.push(x),
            Content::Text(x) if x.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed glossary metadata")),
        }
    }
    Ok(v)
}
pub(in crate::glossary) fn leaf(n: &Node) -> Result<()> {
    if kids(n)?.is_empty() {
        Ok(())
    } else {
        Err(invalid("glossary metadata leaf has children"))
    }
}
pub(in crate::glossary) fn validate_word_dialect(n: &Node, conformance: Conformance) -> Result<()> {
    if matches!(n.ns.as_ref(), W | WS) && n.ns.as_ref() != conformance.word() {
        return Err(invalid(
            "glossary mixes Strict and Transitional WordprocessingML",
        ));
    }
    if matches!(n.ns.as_ref(), R | RS) && n.ns.as_ref() != conformance.relationships() {
        return Err(invalid(
            "glossary mixes Strict and Transitional relationship namespaces",
        ));
    }
    if conformance == Conformance::Strict && n.ns.as_ref() == VML {
        return Err(invalid("Strict glossary content cannot contain VML"));
    }
    for attribute in &n.attrs {
        if matches!(attribute.ns.as_ref(), W | WS) && attribute.ns.as_ref() != conformance.word() {
            return Err(invalid(
                "glossary mixes Strict and Transitional WordprocessingML attributes",
            ));
        }
        if matches!(attribute.ns.as_ref(), R | RS)
            && attribute.ns.as_ref() != conformance.relationships()
        {
            return Err(invalid(
                "glossary mixes Strict and Transitional relationship attributes",
            ));
        }
        if conformance == Conformance::Strict && attribute.ns.as_ref() == VML {
            return Err(invalid(
                "Strict glossary content cannot contain VML attributes",
            ));
        }
    }
    for content in &n.content {
        if let Content::Node(child) = content {
            validate_word_dialect(child, conformance)?;
        }
    }
    Ok(())
}
pub(in crate::glossary) fn expect(n: &Node, l: &str) -> Result<()> {
    if matches!(n.ns.as_ref(), W | WS) && n.l == l {
        Ok(())
    } else {
        Err(invalid(format!("expected WordprocessingML {l}")))
    }
}
pub(in crate::glossary) fn expect_w(n: &Node) -> Result<()> {
    if matches!(n.ns.as_ref(), W | WS) {
        Ok(())
    } else {
        Err(invalid("expected WordprocessingML metadata"))
    }
}
pub(in crate::glossary) fn wval(n: &Node) -> Result<String> {
    let v = wattr_get(n, "val")?.ok_or_else(|| invalid("missing w:val"))?;
    bounded(&v)?;
    only_w(n, &["val"])?;
    leaf(n)?;
    Ok(v)
}
pub(in crate::glossary) fn wattr_get(n: &Node, l: &str) -> Result<Option<String>> {
    let mut v = None;
    for a in &n.attrs {
        if matches!(a.ns.as_ref(), W | WS) && a.l == l {
            if v.is_some() {
                return Err(invalid("duplicate WordprocessingML attribute"));
            }
            v = Some(a.v.clone());
        }
    }
    Ok(v)
}
pub(in crate::glossary) fn onoff(n: &Node, l: &str) -> Result<Option<bool>> {
    let value = wattr_get(n, l)?;
    match value.as_deref() {
        None => Ok(None),
        Some("1" | "true") => Ok(Some(true)),
        Some("0" | "false") => Ok(Some(false)),
        Some("on") if n.ns.as_ref() == W => Ok(Some(true)),
        Some("off") if n.ns.as_ref() == W => Ok(Some(false)),
        _ => Err(invalid(format!("invalid on/off attribute '{l}'"))),
    }
}
pub(in crate::glossary) fn only_w(n: &Node, allowed: &[&str]) -> Result<()> {
    for a in &n.attrs {
        if !(matches!(a.ns.as_ref(), W | WS) && allowed.contains(&a.l.as_str())) {
            return Err(invalid(format!("unexpected glossary attribute '{}'", a.q)));
        }
    }
    Ok(())
}
pub(in crate::glossary) fn noattrs(n: &Node) -> Result<()> {
    if n.attrs.is_empty() {
        Ok(())
    } else {
        Err(invalid("unexpected glossary attributes"))
    }
}
pub(in crate::glossary) fn parse_type(v: &str) -> Result<Kind> {
    match v {
        "none" => Ok(Kind::NONE),
        "normal" => Ok(Kind::NORMAL),
        "autoExp" => Ok(Kind::AUTO_EXPAND),
        "toolbar" => Ok(Kind::TOOLBAR),
        "speller" => Ok(Kind::SPELLER),
        "formFld" => Ok(Kind::FORM_FIELD),
        "bbPlcHdr" => Ok(Kind::SDT_PLACEHOLDER),
        _ => Err(invalid(format!("invalid document-part type '{v}'"))),
    }
}
pub(in crate::glossary) const KIND_VALUES: [(Kind, &str); 7] = [
    (Kind::NONE, "none"),
    (Kind::NORMAL, "normal"),
    (Kind::AUTO_EXPAND, "autoExp"),
    (Kind::TOOLBAR, "toolbar"),
    (Kind::SPELLER, "speller"),
    (Kind::FORM_FIELD, "formFld"),
    (Kind::SDT_PLACEHOLDER, "bbPlcHdr"),
];
pub(in crate::glossary) fn parse_behavior(v: &str) -> Result<Insert> {
    match v {
        "content" => Ok(Insert::CONTENT),
        "p" => Ok(Insert::PARAGRAPH),
        "pg" => Ok(Insert::PAGE),
        _ => Err(invalid(format!("invalid insertion behavior '{v}'"))),
    }
}
pub(in crate::glossary) const INSERT_VALUES: [(Insert, &str); 3] = [
    (Insert::CONTENT, "content"),
    (Insert::PARAGRAPH, "p"),
    (Insert::PAGE, "pg"),
];
pub(in crate::glossary) const GALLERIES: &[&str] = &[
    "placeholder",
    "any",
    "default",
    "docParts",
    "coverPg",
    "eq",
    "ftrs",
    "hdrs",
    "pgNum",
    "tbls",
    "watermarks",
    "autoTxt",
    "txtBox",
    "pgNumT",
    "pgNumB",
    "pgNumMargins",
    "tblOfContents",
    "bib",
    "custQuickParts",
    "custCoverPg",
    "custEq",
    "custFtrs",
    "custHdrs",
    "custPgNum",
    "custTbls",
    "custWatermarks",
    "custAutoTxt",
    "custTxtBox",
    "custPgNumT",
    "custPgNumB",
    "custPgNumMargins",
    "custTblOfContents",
    "custBib",
    "custom1",
    "custom2",
    "custom3",
    "custom4",
    "custom5",
];
pub(in crate::glossary) fn split(q: &str) -> Result<(&str, &str)> {
    if let Some((p, l)) = q.split_once(':') {
        if l.is_empty() || l.contains(':') {
            return Err(invalid("invalid QName"));
        }
        Ok((p, l))
    } else {
        Ok(("", q))
    }
}
pub(in crate::glossary) fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_STRING {
        return Err(invalid("glossary metadata string exceeds 1 MiB"));
    }
    if !v.chars().all(xml_char) {
        return Err(invalid(
            "glossary metadata contains a character forbidden by XML 1.0",
        ));
    }
    Ok(())
}

pub(in crate::glossary) fn xml_char(character: char) -> bool {
    matches!(
        character,
        '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
    )
}
pub(in crate::glossary) fn invalid(value: impl Into<String>) -> Error {
    Error::Invalid(value.into())
}
pub(in crate::glossary) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
