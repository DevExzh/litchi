//! Authoring and safety validation for ODT text elements.

use super::model::{LinkActuate, LinkShow};
use crate::elements::element::{Element, ElementBase};
use litchi_core::{Error, Result};

/// Validate an XLink target before it is inserted into a document.
pub(super) fn href(href: &str) -> Result<()> {
    if href.is_empty() {
        return Err(Error::InvalidFormat(
            "text:a hyperlink href must not be empty".to_string(),
        ));
    }
    if href.chars().any(|character| character.is_control()) {
        return Err(Error::InvalidFormat(
            "text:a hyperlink href must not contain control characters".to_string(),
        ));
    }
    Ok(())
}

/// Validate a string used in an XML attribute value.
pub(super) fn xml_string(value: &str, label: &str) -> Result<()> {
    if value
        .chars()
        .any(|character| matches!(character, '\0'..='\x08' | '\x0B'..='\x0C' | '\x0E'..='\x1F'))
    {
        return Err(Error::InvalidFormat(format!(
            "{label} must not contain XML control characters"
        )));
    }
    Ok(())
}

/// Validate a `text:a` element for safe simple-hyperlink authoring.
pub(super) fn hyperlink(element: &Element) -> Result<()> {
    let href_value = element
        .get_attribute("xlink:href")
        .ok_or_else(|| Error::InvalidFormat("text:a hyperlink requires xlink:href".to_string()))?;
    href(href_value)?;

    match element.get_attribute("xlink:type") {
        Some("simple") => {},
        None => {
            return Err(Error::InvalidFormat(
                "text:a hyperlink requires xlink:type='simple'".to_string(),
            ));
        },
        Some(value) => {
            return Err(Error::InvalidFormat(format!(
                "text:a hyperlink xlink:type must be 'simple', got '{value}'"
            )));
        },
    }

    if element
        .get_attribute("xlink:show")
        .is_some_and(|value| LinkShow::parse(value).is_none())
    {
        return Err(Error::InvalidFormat(
            "text:a hyperlink xlink:show must be 'new' or 'replace'".to_string(),
        ));
    }
    if element
        .get_attribute("xlink:actuate")
        .is_some_and(|value| LinkActuate::parse(value).is_none())
    {
        return Err(Error::InvalidFormat(
            "text:a hyperlink xlink:actuate must be 'onRequest'".to_string(),
        ));
    }

    for (attribute, label) in [
        ("office:name", "text:a hyperlink name"),
        ("office:title", "text:a hyperlink title"),
        (
            "office:target-frame-name",
            "text:a hyperlink target frame name",
        ),
        ("text:style-name", "text:a hyperlink style name"),
        (
            "text:visited-style-name",
            "text:a hyperlink visited style name",
        ),
    ] {
        if let Some(value) = element.get_attribute(attribute) {
            xml_string(value, label)?;
        }
    }
    Ok(())
}
