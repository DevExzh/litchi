//! Namespace and extension validation for worksheet snapshot edits.

use super::super::super::x14ac;
use super::MCE;
use super::snapshot::{Layout, RootEffect, Tag};
use crate::error::{Result, invalid};

/// Namespace names required when a worksheet edit introduces an extension
/// attribute.  The planner is kept separate from the byte writer so callers
/// can validate all namespace effects before any output is allocated.
#[derive(Debug)]
pub(crate) struct ExtensionNames {
    pub(crate) descent: Box<str>,
    pub(crate) root: Option<RootEffect>,
}

impl ExtensionNames {
    pub(crate) fn plan(layout: &Layout, required: bool) -> Result<Self> {
        if !required {
            return Ok(Self {
                descent: "x14ac:dyDescent".into(),
                root: None,
            });
        }

        let root = &layout.root;
        let x14_prefix = x14_prefix(layout)?;
        let mce_prefix = namespace_prefix(&root.tag, MCE)
            .map_or_else(|| available_prefix(&root.tag, "mc"), str::to_owned);
        let mut appended = Vec::new();
        if namespace_uri(&root.tag, &x14_prefix) != Some(x14ac::NAMESPACE) {
            appended.push((
                format!("xmlns:{x14_prefix}").into_boxed_str(),
                String::from_utf8_lossy(x14ac::NAMESPACE).into_owned(),
            ));
        }
        if namespace_prefix(&root.tag, MCE).is_none() {
            appended.push((
                format!("xmlns:{mce_prefix}").into_boxed_str(),
                String::from_utf8_lossy(MCE).into_owned(),
            ));
        }

        let mut ignorable = None::<(&str, &str)>;
        for attribute in &root.tag.attributes {
            let Some((prefix, local)) = attribute.name.split_once(':') else {
                continue;
            };
            if local != "Ignorable" || namespace_uri(&root.tag, prefix) != Some(MCE) {
                continue;
            }
            if ignorable
                .replace((&attribute.name, &attribute.value))
                .is_some()
            {
                return Err(invalid(
                    "worksheet root has duplicate MCE Ignorable attributes",
                ));
            }
        }
        let (removed, ignorable_value) = match ignorable {
            Some((name, value))
                if !value
                    .split_whitespace()
                    .any(|token| token == x14_prefix.as_str()) =>
            {
                let mut tokens = value.split_whitespace().collect::<Vec<_>>();
                tokens.push(&x14_prefix);
                (Some(name.into()), Some(tokens.join(" ")))
            },
            Some(_) => (None, None),
            None => (None, Some(x14_prefix.clone())),
        };
        if let Some(ignorable_value) = ignorable_value {
            appended.push((
                format!("{mce_prefix}:Ignorable").into_boxed_str(),
                ignorable_value,
            ));
        }

        Ok(Self {
            descent: format!("{x14_prefix}:dyDescent").into_boxed_str(),
            root: (!appended.is_empty()).then_some(RootEffect { removed, appended }),
        })
    }
}

fn x14_prefix(layout: &Layout) -> Result<String> {
    if let Some(prefix) = layout.root.tag.attributes.iter().find_map(|attribute| {
        attribute.name.strip_prefix("xmlns:").filter(|prefix| {
            attribute.value.as_bytes() == x14ac::NAMESPACE && x14_prefix_is_usable(layout, prefix)
        })
    }) {
        return Ok(prefix.to_owned());
    }
    if x14_prefix_is_usable(layout, "x14ac") {
        return Ok("x14ac".to_owned());
    }

    let declarations = layout_tags(layout).try_fold(0usize, |count, tag| {
        count
            .checked_add(
                tag.attributes
                    .iter()
                    .filter(|attribute| attribute.name.starts_with("xmlns:"))
                    .count(),
            )
            .ok_or_else(|| invalid("worksheet namespace declaration count overflow"))
    })?;
    let limit = declarations
        .checked_add(1)
        .ok_or_else(|| invalid("worksheet namespace prefix search overflow"))?;
    for suffix in 1..=limit {
        let candidate = format!("x14ac{suffix}");
        if x14_prefix_is_usable(layout, &candidate) {
            return Ok(candidate);
        }
    }
    Err(invalid("cannot allocate a worksheet extension prefix"))
}

fn x14_prefix_is_usable(layout: &Layout, prefix: &str) -> bool {
    layout_tags(layout)
        .all(|tag| namespace_uri(tag, prefix).is_none_or(|namespace| namespace == x14ac::NAMESPACE))
}

fn layout_tags(layout: &Layout) -> impl Iterator<Item = &Tag> {
    std::iter::once(&layout.root.tag)
        .chain(std::iter::once(&layout.sheet_data.tag))
        .chain(layout.defaults.iter().map(|defaults| &defaults.tag))
        .chain(layout.sheet_data.rows.iter().map(|row| &row.tag))
}

fn namespace_prefix<'a>(tag: &'a Tag, namespace: &[u8]) -> Option<&'a str> {
    tag.attributes.iter().find_map(|attribute| {
        attribute
            .name
            .strip_prefix("xmlns:")
            .filter(|_| attribute.value.as_bytes() == namespace)
    })
}

fn namespace_uri<'a>(tag: &'a Tag, prefix: &str) -> Option<&'a [u8]> {
    let name = format!("xmlns:{prefix}");
    tag.attributes
        .iter()
        .find(|attribute| attribute.name.as_ref() == name)
        .map(|attribute| attribute.value.as_bytes())
}

fn available_prefix(tag: &Tag, base: &str) -> String {
    if namespace_uri(tag, base).is_none() {
        return base.to_owned();
    }
    // At most one candidate can be occupied by each stored attribute, so
    // checking one more suffix than the attribute count guarantees a free
    // prefix without relying on an unbounded iterator.
    for suffix in 1..=tag.attributes.len().saturating_add(1) {
        let candidate = format!("{base}{suffix}");
        if namespace_uri(tag, &candidate).is_none() {
            return candidate;
        }
    }
    format!("{base}Extension")
}
