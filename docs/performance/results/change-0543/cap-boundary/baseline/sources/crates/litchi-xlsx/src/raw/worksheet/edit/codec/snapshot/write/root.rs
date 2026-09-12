//! Root, sheet-format, and default-column semantic writers.

use super::super::super::wire::{sibling_name, write_attribute, write_close, write_tag};
use super::super::model::{DefaultsSlot, RootEffect, RootSlot, Tag};
use crate::raw::worksheet::edit::model::{DefaultsEffects, DescentEffect, OptionalEffect};

pub(crate) fn write_root(output: &mut Vec<u8>, root: &RootSlot, effect: &RootEffect) {
    output.extend_from_slice(b"<");
    output.extend_from_slice(root.tag.name.as_bytes());
    for attribute in &root.tag.attributes {
        if effect
            .removed
            .as_deref()
            .is_some_and(|name| name == attribute.name.as_ref())
        {
            continue;
        }
        write_attribute(output, &attribute.name, &attribute.value);
    }
    for (name, value) in &effect.appended {
        write_attribute(output, name, value);
    }
    output.extend_from_slice(b">");
}

pub(crate) fn write_defaults(
    output: &mut Vec<u8>,
    source: &[u8],
    stored: &DefaultsSlot,
    effects: DefaultsEffects,
    descent_name: &str,
) {
    let stored_descent = stored.descent_attribute.as_deref().unwrap_or(descent_name);
    let mut removed = Vec::new();
    let mut appended = Vec::new();
    defaults_effect_attributes(
        effects,
        stored_descent,
        descent_name,
        &mut removed,
        &mut appended,
    );
    write_tag(output, &stored.tag, stored.empty, &removed, &appended);
    if !stored.empty {
        output.extend_from_slice(&source[stored.tag_end..stored.close_start]);
        write_close(output, &stored.tag.name);
    }
}

pub(crate) fn write_new_defaults(
    output: &mut Vec<u8>,
    sheet_data_name: &str,
    effects: DefaultsEffects,
    descent_name: &str,
) {
    let name = sibling_name(sheet_data_name, "sheetFormatPr");
    let tag = Tag {
        name: name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut removed = Vec::new();
    let mut appended = Vec::new();
    defaults_effect_attributes(
        effects,
        descent_name,
        descent_name,
        &mut removed,
        &mut appended,
    );
    write_tag(output, &tag, true, &removed, &appended);
}

fn defaults_effect_attributes<'a>(
    effects: DefaultsEffects,
    stored_descent_name: &'a str,
    appended_descent_name: &'a str,
    removed: &mut Vec<&'a str>,
    appended: &mut Vec<(&'a str, String)>,
) {
    if let Some(effect) = effects.base_width {
        removed.push("baseColWidth");
        if let OptionalEffect::Set(value) = effect {
            appended.push(("baseColWidth", value.to_string()));
        }
    }
    if let Some(effect) = effects.width {
        removed.push("defaultColWidth");
        if let OptionalEffect::Set(value) = effect {
            appended.push(("defaultColWidth", value.get().to_string()));
        }
    }
    if let Some(height) = effects.height {
        removed.extend(["defaultRowHeight", "customHeight"]);
        appended.push(("defaultRowHeight", height.get().to_string()));
        appended.push(("customHeight", "1".to_owned()));
    }
    for (value, name) in [
        (effects.hidden, "zeroHeight"),
        (effects.thick_top, "thickTop"),
        (effects.thick_bottom, "thickBottom"),
    ] {
        if let Some(value) = value {
            removed.push(name);
            if value {
                appended.push((name, "1".to_owned()));
            }
        }
    }
    if let Some(effect) = effects.descent {
        removed.push(stored_descent_name);
        if let DescentEffect::Set(value) = effect {
            appended.push((appended_descent_name, value.get().to_string()));
        }
    }
}
