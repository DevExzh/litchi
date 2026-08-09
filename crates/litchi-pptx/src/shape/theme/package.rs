//! OPC graph operations for `PresentationML` themes.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};

use crate::{Error, Result};

use litchi_drawingml::theme::{FontSet, Override, Palette, Theme, codec};

const STRICT_THEME_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";
const STRICT_OVERRIDE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/themeOverride";

/// Identity of a theme part added to a `PresentationML` package.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Authored {
    /// Package URI of the authored theme part.
    pub part_name: String,
}

/// Add a validated theme part at the next free `/ppt/theme/themeN.xml` URI.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn add(
    package: &mut OpcPackage,
    name: &str,
    colors: &Palette,
    fonts: &FontSet,
) -> Result<Authored> {
    let uri = next_part_uri(package, "/ppt/theme/theme", ".xml")?;
    let xml = codec::encode_part(name, colors, fonts)?;
    package.add_part(Box::new(BlobPart::new(
        uri.clone(),
        ct::OFC_THEME.to_owned(),
        xml,
    )));
    package.unsign();
    validate(package)?;
    Ok(Authored {
        part_name: uri.to_string(),
    })
}

/// Attach one theme part to a slide master, replacing any previous theme link.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn attach(package: &mut OpcPackage, master_name: &str, theme_name: &str) -> Result<String> {
    let master_uri = uri(master_name, "theme master")?;
    let theme_uri = uri(theme_name, "theme part")?;
    require_type(package, &master_uri, ct::PML_SLIDE_MASTER)?;
    require_type(package, &theme_uri, ct::OFC_THEME)?;
    let target = theme_uri.relative_ref(master_uri.base_uri());
    let master = package.get_part_mut(&master_uri)?;
    let old: Vec<String> = master
        .rels()
        .iter()
        .filter(|r| is_theme_rel(r.reltype()))
        .map(|r| r.r_id().to_owned())
        .collect();
    for id in old {
        master.rels_mut().remove(&id);
    }
    let relationship_id = master.relate_to(&target, rt::THEME);
    package.unsign();
    validate(package)?;
    Ok(relationship_id)
}

/// Read one theme part by URI.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load(package: &OpcPackage, theme_name: &str) -> Result<Theme> {
    let uri = uri(theme_name, "theme part")?;
    require_type(package, &uri, ct::OFC_THEME)?;
    Ok(codec::read(package.get_part(&uri)?.blob())?)
}

/// Read the permissive summary view used by producer-facing theme inspection.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_summary(package: &OpcPackage, theme_name: &str) -> Result<super::part::Summary> {
    let uri = uri(theme_name, "theme part")?;
    require_type(package, &uri, ct::OFC_THEME)?;
    super::part::Part::from_part(package.get_part(&uri)?)?.read()
}

/// Replace only the color palette of an existing theme part.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn put_colors(package: &mut OpcPackage, theme_name: &str, colors: &Palette) -> Result<()> {
    let uri = uri(theme_name, "theme part")?;
    require_type(package, &uri, ct::OFC_THEME)?;
    let original = package.get_part(&uri)?.blob().to_vec();
    let patched = codec::replace_scheme(
        &original,
        b"clrScheme",
        &codec::encode_palette_fragment(colors)?,
    )?;
    if codec::read(&patched)?.colors != *colors {
        return Err(invalid("color palette did not survive theme read-back"));
    }
    package.get_part_mut(&uri)?.set_blob(patched);
    package.unsign();
    validate(package)
}

/// Replace only the font set of an existing theme part.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn put_fonts(package: &mut OpcPackage, theme_name: &str, fonts: &FontSet) -> Result<()> {
    let uri = uri(theme_name, "theme part")?;
    require_type(package, &uri, ct::OFC_THEME)?;
    let original = package.get_part(&uri)?.blob().to_vec();
    let patched = codec::replace_scheme(
        &original,
        b"fontScheme",
        &codec::encode_fonts_fragment(fonts)?,
    )?;
    if codec::read(&patched)?.fonts != *fonts {
        return Err(invalid("font set did not survive theme read-back"));
    }
    package.get_part_mut(&uri)?.set_blob(patched);
    package.unsign();
    validate(package)
}

/// Create or replace a slide/layout theme override.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn put_override(
    package: &mut OpcPackage,
    parent_name: &str,
    value: &Override,
) -> Result<String> {
    let parent_uri = uri(parent_name, "theme override parent")?;
    let parent = package.get_part(&parent_uri)?;
    if !matches!(parent.content_type(), ct::PML_SLIDE | ct::PML_SLIDE_LAYOUT) {
        return Err(invalid("theme overrides attach only to slides and layouts"));
    }
    let xml = codec::encode_override(value)?;
    let existing = parent
        .rels()
        .iter()
        .find(|r| is_override_rel(r.reltype()) && !r.is_external());
    if let Some(relationship) = existing {
        let target = relationship.target_partname()?;
        require_type(package, &target, ct::OFC_THEME_OVERRIDE)?;
        package.get_part_mut(&target)?.set_blob(xml);
        if codec::read_override(package.get_part(&target)?.blob())? != *value {
            return Err(invalid("theme override did not survive read-back"));
        }
        package.unsign();
        return Ok(target.to_string());
    }
    let target = next_part_uri(package, "/ppt/theme/themeOverride", ".xml")?;
    package.add_part(Box::new(BlobPart::new(
        target.clone(),
        ct::OFC_THEME_OVERRIDE.to_owned(),
        xml,
    )));
    let reference = target.relative_ref(parent_uri.base_uri());
    package
        .get_part_mut(&parent_uri)?
        .relate_to(&reference, rt::THEME_OVERRIDE);
    if codec::read_override(package.get_part(&target)?.blob())? != *value {
        return Err(invalid("theme override did not survive read-back"));
    }
    package.unsign();
    Ok(target.to_string())
}

/// Read an optional override attached to a slide or layout.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn load_override(package: &OpcPackage, parent_name: &str) -> Result<Option<Override>> {
    let parent_uri = uri(parent_name, "theme override parent")?;
    let parent = package.get_part(&parent_uri)?;
    let Some(relationship) = parent
        .rels()
        .iter()
        .find(|r| is_override_rel(r.reltype()) && !r.is_external())
    else {
        return Ok(None);
    };
    let target = relationship.target_partname()?;
    require_type(package, &target, ct::OFC_THEME_OVERRIDE)?;
    Ok(Some(codec::read_override(
        package.get_part(&target)?.blob(),
    )?))
}

/// Remove an override and delete its part when no other parent references it.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn remove_override(package: &mut OpcPackage, parent_name: &str) -> Result<bool> {
    let parent_uri = uri(parent_name, "theme override parent")?;
    let parent = package.get_part(&parent_uri)?;
    let Some(relationship) = parent
        .rels()
        .iter()
        .find(|r| is_override_rel(r.reltype()) && !r.is_external())
    else {
        return Ok(false);
    };
    let target = relationship.target_partname()?;
    let id = relationship.r_id().to_owned();
    package.get_part_mut(&parent_uri)?.rels_mut().remove(&id);
    let still_used = package.iter_parts().any(|part| {
        part.rels().iter().any(|r| {
            is_override_rel(r.reltype())
                && !r.is_external()
                && r.target_partname().ok().as_ref() == Some(&target)
        })
    });
    if !still_used {
        package.remove_part(&target);
    }
    package.unsign();
    Ok(true)
}

/// Validate theme relationships, theme parts, and slide/layout overrides.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn validate(package: &OpcPackage) -> Result<()> {
    for part in package.iter_parts() {
        if part.content_type() == ct::PML_SLIDE_MASTER {
            let mut themes = part
                .rels()
                .iter()
                .filter(|r| is_theme_rel(r.reltype()) && !r.is_external());
            let Some(theme) = themes.next() else {
                return Err(invalid(format!(
                    "slide master '{}' has no theme relationship",
                    part.partname()
                )));
            };
            if themes.next().is_some() {
                return Err(invalid(format!(
                    "slide master '{}' has multiple theme relationships",
                    part.partname()
                )));
            }
            let target = theme.target_partname()?;
            require_type(package, &target, ct::OFC_THEME)?;
            let _ = codec::read(package.get_part(&target)?.blob())?;
        }
        if matches!(part.content_type(), ct::PML_SLIDE | ct::PML_SLIDE_LAYOUT) {
            for relationship in part.rels().iter().filter(|r| is_override_rel(r.reltype())) {
                if relationship.is_external() {
                    return Err(invalid("theme override relationships must be internal"));
                }
                let target = relationship.target_partname()?;
                require_type(package, &target, ct::OFC_THEME_OVERRIDE)?;
                let _ = codec::read_override(package.get_part(&target)?.blob())?;
            }
        }
    }
    Ok(())
}

/// Allocate a deterministic free part name for theme adapters.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn next_part_uri(package: &OpcPackage, prefix: &str, suffix: &str) -> Result<PackURI> {
    let mut index = 1u32;
    loop {
        let candidate = PackURI::new(format!("{prefix}{index}{suffix}")).map_err(Error::Uri)?;
        if !package.contains_part(&candidate) {
            return Ok(candidate);
        }
        index = index
            .checked_add(1)
            .ok_or_else(|| invalid("theme part index overflow"))?;
    }
}

fn require_type(package: &OpcPackage, uri: &PackURI, expected: &str) -> Result<()> {
    let part = package.get_part(uri)?;
    if part.content_type() != expected {
        return Err(Error::ContentType {
            expected: expected.to_owned(),
            actual: part.content_type().to_owned(),
        });
    }
    Ok(())
}

fn uri(value: &str, label: &str) -> Result<PackURI> {
    PackURI::new(value).map_err(|error| Error::Uri(format!("{label}: {error}")))
}
fn is_theme_rel(value: &str) -> bool {
    matches!(value, rt::THEME | STRICT_THEME_REL)
}
fn is_override_rel(value: &str) -> bool {
    matches!(value, rt::THEME_OVERRIDE | STRICT_OVERRIDE_REL)
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::shape::theme::{Color, Face, FontSet, Palette, Slot};

    fn palette() -> Palette {
        Slot::ALL
            .into_iter()
            .fold(Palette::new("Office"), |palette, slot| {
                palette.with(slot, Color::rgb("4F81BD").unwrap())
            })
    }

    #[test]
    fn add_and_load_theme_uses_direct_opc_graph() {
        let mut package = OpcPackage::new();
        let theme = add(
            &mut package,
            "Office",
            &palette(),
            &FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos")),
        )
        .unwrap();
        assert_eq!(load(&package, &theme.part_name).unwrap().name, "Office");
    }
}
