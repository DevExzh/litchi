use super::super::codec::{enforce_count_with, invalid, limit, require_nonempty};
use super::super::model::{
    AddIn, Binding, Effect, ExtKind, ExtList, Limits, OperationBudget, Pane, Panes, Reference,
    SnapshotTarget, Store,
};
use super::super::package::fold_part_name;
use super::super::{ADD_IN_RELATIONSHIP, IMAGE_RELATIONSHIP_TYPE};
use crate::{Error, Result};
use std::collections::{HashMap, HashSet};
pub(in crate::web) fn validate_model(extension: &AddIn) -> Result<()> {
    validate_model_with(extension, &Limits::standard())
}

pub(in crate::web) fn validate_model_with(extension: &AddIn, limits: &Limits) -> Result<()> {
    require_nonempty("web extension id", &extension.id)?;
    validate_store_reference(&extension.reference)?;
    enforce_count_with(
        "alternate reference",
        extension.alternate_references.len(),
        limits,
    )?;
    enforce_count_with("property", extension.properties.len(), limits)?;
    enforce_count_with("binding", extension.bindings.len(), limits)?;
    let mut reference_ids = HashSet::new();
    reference_ids.insert(extension.reference.id.as_str());
    for reference in &extension.alternate_references {
        validate_store_reference(reference)?;
        if !reference_ids.insert(reference.id.as_str()) {
            return invalid(format!("duplicate reference id '{}'", reference.id));
        }
    }
    let mut property_names = HashSet::new();
    for property in &extension.properties {
        require_nonempty("property name", &property.name)?;
        if !property_names.insert(property.name.as_str()) {
            return invalid(format!("duplicate property name '{}'", property.name));
        }
    }
    let mut binding_ids = HashSet::new();
    let mut binding_app_refs = HashSet::new();
    for binding in &extension.bindings {
        validate_binding(binding)?;
        if !binding_ids.insert(binding.id.as_str()) {
            return invalid(format!("duplicate binding id '{}'", binding.id));
        }
        if !binding_app_refs.insert(binding.app_ref.as_str()) {
            return invalid(format!("duplicate binding appRef '{}'", binding.app_ref));
        }
    }
    if let Some(snapshot) = &extension.snapshot {
        enforce_count_with("snapshot effect", snapshot.effects.len(), limits)?;
        for effect in &snapshot.effects {
            let reparsed = Effect::from_xml(effect.xml.as_bytes())?;
            if reparsed.kind != effect.kind {
                return invalid("snapshot effect kind does not match its XML root".into());
            }
        }
        validate_extension_list(
            snapshot.extension_list.as_ref(),
            &[ExtKind::DrawingMl, ExtKind::StrictDrawingMl],
        )?;
    }
    validate_extension_list(extension.extension_list.as_ref(), &[ExtKind::AddIn])?;
    Ok(())
}

pub(in crate::web) fn validate_binding(binding: &Binding) -> Result<()> {
    require_nonempty("binding id", &binding.id)?;
    require_nonempty("binding type", binding.kind.as_str())?;
    require_nonempty("binding appref", &binding.app_ref)?;
    validate_extension_list(binding.extension_list.as_ref(), &[ExtKind::AddIn])
}

pub(in crate::web) fn validate_store_reference(reference: &Reference) -> Result<()> {
    require_nonempty("reference id", &reference.id)?;
    require_nonempty("reference version", &reference.version)?;
    if let Some(location) = &reference.location {
        require_nonempty("reference location", location)?;
    } else if reference.store == Store::FileSystem {
        return invalid("FileSystem reference requires a non-empty location".into());
    }
    validate_extension_list(reference.extension_list.as_ref(), &[ExtKind::AddIn])
}

pub(in crate::web) fn validate_task_pane(pane: &Pane) -> Result<()> {
    validate_task_pane_with(pane, &Limits::standard())
}

pub(in crate::web) fn validate_task_pane_with(pane: &Pane, limits: &Limits) -> Result<()> {
    require_nonempty("dock state", pane.dock_state.as_str())?;
    require_nonempty("task-pane relationship id", &pane.relationship_id)?;
    if !pane.width.is_finite() || pane.width <= 0.0 {
        return invalid("task-pane width must be finite and positive".into());
    }
    validate_extension_list(pane.extension_list.as_ref(), &[ExtKind::TaskPane])?;
    validate_model_with(&pane.add_in, limits)?;
    validate_snapshot_resources_with(pane, limits)
}

pub(in crate::web) fn validate_panes(task_panes: &Panes, limits: &Limits) -> Result<()> {
    enforce_count_with("task pane", task_panes.panes.len(), limits)?;
    let mut relationship_ids = HashSet::new();
    let mut extension_ids = HashSet::new();
    let mut total_snapshot_bytes = 0usize;
    let mut snapshot_names = HashSet::new();
    for pane in &task_panes.panes {
        validate_task_pane_with(pane, limits)?;
        if !relationship_ids.insert(pane.relationship_id.as_str()) {
            return invalid(format!(
                "duplicate task-pane relationship ID '{}'",
                pane.relationship_id
            ));
        }
        if !extension_ids.insert(pane.add_in.id.as_str()) {
            return invalid(format!(
                "duplicate web extension instance ID '{}'",
                pane.add_in.id
            ));
        }
        for resource in &pane.snapshot_resources {
            if let SnapshotTarget::Internal {
                part_name, data, ..
            } = &resource.target
            {
                if data.len() > limits.image_bytes {
                    return limit(
                        "web extension snapshot bytes",
                        limits.image_bytes,
                        data.len(),
                    );
                }
                if snapshot_names.insert(fold_part_name(part_name)) {
                    total_snapshot_bytes =
                        total_snapshot_bytes
                            .checked_add(data.len())
                            .ok_or(Error::Limit {
                                resource: "aggregate web extension snapshot bytes",
                                max: limits.total_image_bytes,
                                actual: usize::MAX,
                            })?;
                    if total_snapshot_bytes > limits.total_image_bytes {
                        return limit(
                            "aggregate web extension snapshot bytes",
                            limits.total_image_bytes,
                            total_snapshot_bytes,
                        );
                    }
                }
            }
        }
    }
    Ok(())
}

pub(in crate::web) fn charge_authored_metadata(
    task_panes: &Panes,
    budget: &mut OperationBudget,
    limits: &Limits,
) -> Result<()> {
    let generated_names = task_panes
        .panes
        .len()
        .checked_add(1)
        .and_then(|count| count.checked_mul(128))
        .ok_or(Error::Limit {
            resource: "authored web extension package metadata bytes",
            max: limits.total_string_bytes,
            actual: usize::MAX,
        })?;
    budget.charge_metadata(generated_names, 2, limits)?;
    for pane in &task_panes.panes {
        let pane_bytes = pane
            .relationship_id
            .len()
            .checked_add(ADD_IN_RELATIONSHIP.len())
            .ok_or(Error::Limit {
                resource: "authored web extension package metadata bytes",
                max: limits.total_string_bytes,
                actual: usize::MAX,
            })?;
        budget.charge_metadata(pane_bytes, 2, limits)?;
        for resource in &pane.snapshot_resources {
            let target_bytes = match &resource.target {
                SnapshotTarget::Internal {
                    part_name,
                    content_type,
                    ..
                } => part_name.as_str().len().checked_add(content_type.len()),
                SnapshotTarget::External { target } => Some(target.len()),
            }
            .and_then(|bytes| bytes.checked_add(resource.relationship_id.len()))
            .and_then(|bytes| bytes.checked_add(IMAGE_RELATIONSHIP_TYPE.len()))
            .ok_or(Error::Limit {
                resource: "authored web extension package metadata bytes",
                max: limits.total_string_bytes,
                actual: usize::MAX,
            })?;
            budget.charge_metadata(target_bytes, 3, limits)?;
        }
    }
    Ok(())
}

pub(in crate::web) fn add_xml_budget(
    total: &mut usize,
    bytes: usize,
    limits: &Limits,
) -> Result<()> {
    *total = total.checked_add(bytes).ok_or(Error::Limit {
        resource: "authored web extension XML bytes",
        max: limits.xml_bytes,
        actual: usize::MAX,
    })?;
    if *total > limits.xml_bytes {
        return limit("authored web extension XML bytes", limits.xml_bytes, *total);
    }
    Ok(())
}

pub(in crate::web) fn add_escaped_xml_budget(
    total: &mut usize,
    value: &str,
    limits: &Limits,
) -> Result<()> {
    let bytes = value.len().checked_mul(6).ok_or(Error::Limit {
        resource: "authored web extension XML bytes",
        max: limits.xml_bytes,
        actual: usize::MAX,
    })?;
    add_xml_budget(total, bytes, limits)
}

pub(in crate::web) fn add_reference_budget(
    total: &mut usize,
    reference: &Reference,
    limits: &Limits,
) -> Result<()> {
    add_xml_budget(total, 128, limits)?;
    add_escaped_xml_budget(total, &reference.id, limits)?;
    add_escaped_xml_budget(total, &reference.version, limits)?;
    if let Some(location) = &reference.location {
        add_escaped_xml_budget(total, location, limits)?;
    }
    if let Some(extension_list) = &reference.extension_list {
        add_xml_budget(total, extension_list.xml.len(), limits)?;
    }
    Ok(())
}

pub(in crate::web) fn validate_add_in_budget(extension: &AddIn, limits: &Limits) -> Result<()> {
    let mut total = 512usize;
    add_escaped_xml_budget(&mut total, &extension.id, limits)?;
    add_reference_budget(&mut total, &extension.reference, limits)?;
    for reference in &extension.alternate_references {
        add_reference_budget(&mut total, reference, limits)?;
    }
    for property in &extension.properties {
        add_xml_budget(&mut total, 64, limits)?;
        add_escaped_xml_budget(&mut total, &property.name, limits)?;
        add_escaped_xml_budget(&mut total, &property.value, limits)?;
    }
    for binding in &extension.bindings {
        add_xml_budget(&mut total, 96, limits)?;
        add_escaped_xml_budget(&mut total, &binding.id, limits)?;
        add_escaped_xml_budget(&mut total, binding.kind.as_str(), limits)?;
        add_escaped_xml_budget(&mut total, &binding.app_ref, limits)?;
        if let Some(extension_list) = &binding.extension_list {
            add_xml_budget(&mut total, extension_list.xml.len(), limits)?;
        }
    }
    if let Some(snapshot) = &extension.snapshot {
        add_xml_budget(&mut total, 160, limits)?;
        for effect in &snapshot.effects {
            add_xml_budget(&mut total, effect.xml.len(), limits)?;
        }
        if let Some(extension_list) = &snapshot.extension_list {
            add_xml_budget(&mut total, extension_list.xml.len(), limits)?;
        }
    }
    if let Some(extension_list) = &extension.extension_list {
        add_xml_budget(&mut total, extension_list.xml.len(), limits)?;
    }
    Ok(())
}

pub(in crate::web) fn validate_panes_budget(task_panes: &Panes, limits: &Limits) -> Result<()> {
    let mut total = 384usize;
    for pane in &task_panes.panes {
        add_xml_budget(&mut total, 192, limits)?;
        add_escaped_xml_budget(&mut total, pane.dock_state.as_str(), limits)?;
        add_escaped_xml_budget(&mut total, &pane.relationship_id, limits)?;
        if let Some(extension_list) = &pane.extension_list {
            add_xml_budget(&mut total, extension_list.xml.len(), limits)?;
        }
    }
    Ok(())
}

pub(in crate::web) fn validate_extension_list(
    extension_list: Option<&ExtList>,
    allowed: &[ExtKind],
) -> Result<()> {
    let Some(extension_list) = extension_list else {
        return Ok(());
    };
    if !allowed.contains(&extension_list.kind) {
        return invalid(format!(
            "extLst namespace '{}' is not valid at this location",
            extension_list.kind.namespace()
        ));
    }
    let reparsed = ExtList::from_xml(extension_list.as_xml())?;
    if reparsed != *extension_list {
        return invalid("extLst fragment is not a stable self-contained XML tree".into());
    }
    Ok(())
}

pub(in crate::web) fn validate_snapshot_resources_with(pane: &Pane, limits: &Limits) -> Result<()> {
    let mut expected = HashMap::new();
    if let Some(snapshot) = &pane.add_in.snapshot {
        if let Some(id) = snapshot.embedded_relationship_id.as_deref() {
            require_nonempty("embedded snapshot relationship ID", id)?;
            expected.insert(id, false);
        }
        if let Some(id) = snapshot.linked_relationship_id.as_deref() {
            require_nonempty("linked snapshot relationship ID", id)?;
            if expected.insert(id, true).is_some() {
                return invalid("snapshot embed and link IDs must differ".into());
            }
        }
    }
    if expected.len() != pane.snapshot_resources.len() {
        return invalid("snapshot relationship and resource counts differ".into());
    }
    let mut resource_ids = HashSet::new();
    for resource in &pane.snapshot_resources {
        require_nonempty(
            "snapshot resource relationship ID",
            &resource.relationship_id,
        )?;
        if !resource_ids.insert(resource.relationship_id.as_str()) {
            return invalid(format!(
                "duplicate snapshot resource relationship ID '{}'",
                resource.relationship_id
            ));
        }
        let Some(linked) = expected.get(resource.relationship_id.as_str()) else {
            return invalid(format!(
                "snapshot resource '{}' is not referenced by the web extension",
                resource.relationship_id
            ));
        };
        match &resource.target {
            SnapshotTarget::Internal {
                part_name,
                content_type,
                data,
            } => {
                if part_name.as_str() == "/" {
                    return invalid("snapshot image cannot target the package root".into());
                }
                validate_image_content_type(content_type)?;
                if data.len() > limits.image_bytes {
                    return limit(
                        "web extension snapshot bytes",
                        limits.image_bytes,
                        data.len(),
                    );
                }
            },
            SnapshotTarget::External { target } => {
                if !*linked {
                    return invalid(format!(
                        "embedded snapshot resource '{}' cannot be external",
                        resource.relationship_id
                    ));
                }
                validate_external_uri_reference(target)?;
            },
        }
    }
    Ok(())
}
pub(in crate::web) fn validate_image_content_type(value: &str) -> Result<()> {
    if value.len() > 255 || value.bytes().any(|byte| byte.is_ascii_control()) {
        return invalid(format!("invalid snapshot image content type '{value}'"));
    }
    let Some((top_level, subtype)) = value.split_once('/') else {
        return invalid(format!("invalid snapshot image content type '{value}'"));
    };
    if !top_level.eq_ignore_ascii_case("image")
        || subtype.is_empty()
        || subtype.contains('/')
        || !top_level.bytes().all(is_mime_token_byte)
        || !subtype.bytes().all(is_mime_token_byte)
    {
        return invalid(format!("invalid snapshot image content type '{value}'"));
    }
    Ok(())
}

pub(in crate::web) fn is_mime_token_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric()
        || matches!(
            byte,
            b'!' | b'#'
                | b'$'
                | b'%'
                | b'&'
                | b'\''
                | b'*'
                | b'+'
                | b'-'
                | b'.'
                | b'^'
                | b'_'
                | b'`'
                | b'|'
                | b'~'
        )
}

pub(in crate::web) fn validate_external_uri_reference(value: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 32 * 1024
        || value.chars().any(char::is_whitespace)
        || value.chars().any(char::is_control)
        || value.contains('\\')
    {
        return invalid("external snapshot target is not a valid URI-reference".into());
    }
    let bytes = value.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'%' {
            if bytes
                .get(index + 1..index + 3)
                .is_none_or(|encoded| !encoded.iter().all(u8::is_ascii_hexdigit))
            {
                return invalid("external snapshot target has invalid percent encoding".into());
            }
            index += 3;
        } else {
            index += 1;
        }
    }
    let base = url::Url::parse("https://litchi.invalid/")
        .map_err(|error| Error::Uri(error.to_string()))?;
    url::Url::options()
        .base_url(Some(&base))
        .parse(value)
        .map_err(|error| Error::Uri(format!("invalid external snapshot URI-reference: {error}")))?;
    Ok(())
}
