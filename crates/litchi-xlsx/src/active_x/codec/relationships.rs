//! Relationship-reference codec for ActiveX descriptor graphs.

use super::super::model::{Descriptor, Font, Property, PropertyObject};
use super::super::{Result, relerr};
use std::collections::BTreeSet;

/// Returns every binary relationship referenced by a descriptor in stable
/// order. Duplicate references are rejected because one relationship ID must
/// identify one package edge in the ActiveX graph.
pub(crate) fn descriptor_relationship_ids(value: &Descriptor) -> Result<Vec<String>> {
    let mut ids = BTreeSet::new();
    if let Some(id) = value.relationship_id.as_ref() {
        insert(&mut ids, id)?;
    }
    collect_properties(&value.properties, &mut ids)?;
    Ok(ids.into_iter().collect())
}

fn collect_properties(values: &[Property], ids: &mut BTreeSet<String>) -> Result<()> {
    for property in values {
        match property.object.as_ref() {
            Some(PropertyObject::Font(font)) => {
                collect_font(font, ids)?;
            },
            Some(PropertyObject::Picture(picture)) => {
                if let Some(id) = picture.relationship_id.as_ref() {
                    insert(ids, id)?;
                }
            },
            None => {},
        }
    }
    Ok(())
}

fn collect_font(font: &Font, ids: &mut BTreeSet<String>) -> Result<()> {
    if let Some(id) = font.relationship_id.as_ref() {
        insert(ids, id)?;
    }
    collect_properties(&font.properties, ids)
}

fn insert(ids: &mut BTreeSet<String>, id: &str) -> Result<()> {
    if ids.insert(id.to_string()) {
        Ok(())
    } else {
        Err(relerr(
            "ActiveX relationship ID is referenced more than once",
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::active_x::model::{Persistence, Property};

    #[test]
    fn emits_nested_relationships_in_stable_order() {
        let descriptor = Descriptor {
            class_id: "inert".into(),
            license: None,
            persistence: Persistence::PropertyBag,
            relationship_id: None,
            properties: vec![Property {
                name: "Picture".into(),
                value: None,
                object: Some(PropertyObject::Picture(crate::active_x::model::Picture {
                    relationship_id: Some("rId2".into()),
                })),
            }],
        };
        assert_eq!(descriptor_relationship_ids(&descriptor).unwrap(), ["rId2"]);
    }

    #[test]
    fn rejects_reused_relationship_ids() {
        let descriptor = Descriptor {
            class_id: "inert".into(),
            license: None,
            persistence: Persistence::PropertyBag,
            relationship_id: None,
            properties: vec![
                Property {
                    name: "one".into(),
                    value: None,
                    object: Some(PropertyObject::Picture(crate::active_x::model::Picture {
                        relationship_id: Some("rId1".into()),
                    })),
                },
                Property {
                    name: "two".into(),
                    value: None,
                    object: Some(PropertyObject::Picture(crate::active_x::model::Picture {
                        relationship_id: Some("rId1".into()),
                    })),
                },
            ],
        };
        assert!(descriptor_relationship_ids(&descriptor).is_err());
    }
}
