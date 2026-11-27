//! Worksheet-to-package loading for inert `ActiveX` graphs.

use super::super::codec::descriptor_relationship_ids;
use super::super::model::{Binary, Control, Descriptor, LoadedControl, PreviewImage};
use super::super::{
    BINARY_CONTENT_TYPE, BINARY_REL, CONTROL_REL, CONTROL_REL_STRICT, DESCRIPTOR_CONTENT_TYPE,
    IMAGE_REL, IMAGE_REL_STRICT, MAX_BINARY, MAX_TOTAL_BINARY, Result, content_type, limit, relerr,
};
use litchi_opc::{OpcPackage, Part};

/// Resolves one worksheet control into its descriptor, opaque binaries, and
/// optional preview image without activating or interpreting any payload.
pub(super) fn load_control(
    package: &OpcPackage,
    worksheet: &dyn Part,
    control: Control,
    total_binary: &mut usize,
) -> Result<LoadedControl> {
    let preview = if let Some(id) = control
        .properties
        .as_ref()
        .and_then(|properties| properties.preview_relationship_id.as_deref())
    {
        let relationship = worksheet
            .rels()
            .get(id)
            .ok_or_else(|| relerr("control preview relationship is missing"))?;
        if relationship.is_external()
            || !matches!(relationship.reltype(), IMAGE_REL | IMAGE_REL_STRICT)
        {
            return Err(relerr(
                "control preview must be an internal image relationship",
            ));
        }
        let part_uri = relationship.target_partname()?;
        let part = package.get_part(&part_uri)?;
        if !part.content_type().starts_with("image/") {
            return Err(content_type("image/*", part.content_type()));
        }
        account_bytes(
            total_binary,
            part.blob().len(),
            "ActiveX preview image bytes",
        )?;
        Some(PreviewImage {
            relationship_id: id.to_string(),
            part_uri,
            content_type: part.content_type().to_string(),
            bytes: part.blob().to_vec(),
        })
    } else {
        None
    };

    let relationship = worksheet
        .rels()
        .get(&control.relationship_id)
        .ok_or_else(|| relerr("control relationship is missing"))?;
    if relationship.is_external()
        || !matches!(relationship.reltype(), CONTROL_REL | CONTROL_REL_STRICT)
    {
        return Err(relerr("control must target an internal ActiveX descriptor"));
    }
    let descriptor_uri = relationship.target_partname()?;
    let part = package.get_part(&descriptor_uri)?;
    if part.content_type() != DESCRIPTOR_CONTENT_TYPE {
        return Err(content_type(DESCRIPTOR_CONTENT_TYPE, part.content_type()));
    }
    let descriptor = Descriptor::parse(part.blob())?;
    let ids = descriptor_relationship_ids(&descriptor)?;
    if part.rels().iter().count() != ids.len() {
        return Err(relerr(
            "ActiveX descriptor has unexpected or duplicate outgoing relationships",
        ));
    }

    let mut binaries = Vec::with_capacity(ids.len());
    for id in ids {
        let binary_relationship = part
            .rels()
            .get(&id)
            .ok_or_else(|| relerr("ActiveX binary relationship is missing"))?;
        if binary_relationship.is_external() || binary_relationship.reltype() != BINARY_REL {
            return Err(relerr(
                "ActiveX descriptor may relate only to internal ActiveX binaries",
            ));
        }
        let binary_uri = binary_relationship.target_partname()?;
        let binary = package.get_part(&binary_uri)?;
        if binary.content_type() != BINARY_CONTENT_TYPE {
            return Err(content_type(BINARY_CONTENT_TYPE, binary.content_type()));
        }
        if binary.rels().iter().next().is_some() {
            return Err(relerr("ActiveX binary part must not have relationships"));
        }
        account_bytes(total_binary, binary.blob().len(), "ActiveX binary bytes")?;
        binaries.push(Binary {
            relationship_id: id,
            part_uri: binary_uri,
            bytes: binary.blob().to_vec(),
        });
    }

    Ok(LoadedControl {
        control,
        descriptor_uri,
        descriptor,
        binaries,
        preview,
    })
}

fn account_bytes(total: &mut usize, amount: usize, what: &str) -> Result<()> {
    if amount > MAX_BINARY {
        return Err(limit(what));
    }
    *total = total
        .checked_add(amount)
        .ok_or_else(|| limit("aggregate ActiveX resource bytes"))?;
    if *total > MAX_TOTAL_BINARY {
        return Err(limit("aggregate ActiveX resource bytes"));
    }
    Ok(())
}
