//! Target-driven, inert object discovery.

use super::codec::{self, Package};
use super::model::{Limits, Objects};
use super::target::Targets;
use litchi_cfb::{OleError, OleFile};
use std::io::{Read, Seek};

/// Captures exactly the host-selected CFB storages.
pub fn discover<R: Read + Seek>(
    ole: &mut OleFile<R>,
    targets: &Targets,
    limits: Limits,
) -> Result<Objects, OleError> {
    limits.validate()?;
    if targets.len() > limits.max_objects {
        return Err(OleError::InvalidFormat(format!(
            "object target count {} exceeds limit {}",
            targets.len(),
            limits.max_objects
        )));
    }
    codec::open(ole)?;
    let mut objects = Vec::new();
    let mut total = 0u64;
    for target in targets {
        let object = Package::capture_target(ole, target, limits)?;
        total = total
            .checked_add(object.compound().len() as u64)
            .ok_or_else(|| OleError::InvalidFormat("object total size overflow".into()))?;
        if total > limits.max_total_size {
            return Err(OleError::InvalidFormat(
                "object total size exceeds limit".into(),
            ));
        }
        objects.push(object);
    }
    Objects::new(objects)
}

pub(crate) fn from_package(
    package: &Package,
    targets: &Targets,
    limits: Limits,
) -> Result<Objects, OleError> {
    limits.validate()?;
    if targets.len() > limits.max_objects {
        return Err(OleError::InvalidFormat(format!(
            "object target count {} exceeds limit {}",
            targets.len(),
            limits.max_objects
        )));
    }
    let mut objects = Vec::new();
    let mut total = 0u64;
    for target in targets {
        let object = package.object(target.clone(), limits)?;
        total = total
            .checked_add(object.compound().len() as u64)
            .ok_or_else(|| OleError::InvalidFormat("object total size overflow".into()))?;
        if total > limits.max_total_size {
            return Err(OleError::InvalidFormat(
                "object total size exceeds limit".into(),
            ));
        }
        objects.push(object);
    }
    Objects::new(objects)
}
