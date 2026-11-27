//! Target-driven, inert object discovery.

use super::codec::{self, Package};
use super::model::{Limits, Objects};
use super::target::Targets;
use crate::{Error, LimitKind};
use litchi_cfb::{OleError, OleFile};
use std::io::{Read, Seek};

/// Captures exactly the host-selected CFB storages.
///
/// # Errors
///
/// Returns an error when the CFB is malformed or protected, a target is
/// missing/invalid, or a configured resource limit is exceeded.
pub fn discover<R: Read + Seek>(
    ole: &mut OleFile<R>,
    targets: &Targets,
    limits: Limits,
) -> Result<Objects, Error> {
    limits.validate()?;
    if targets.len() > limits.max_objects {
        return Err(Error::Limit {
            kind: LimitKind::Objects,
            observed: targets.len() as u64,
            maximum: limits.max_objects as u64,
        });
    }
    codec::open(ole, limits.max_package_directory_entries())?;
    let mut objects = Vec::new();
    let mut total = 0u64;
    for target in targets {
        let object = Package::capture_target(ole, target, limits)?;
        total = total
            .checked_add(object.compound().len() as u64)
            .ok_or_else(|| OleError::InvalidFormat("object total size overflow".into()))?;
        if total > limits.max_total_size {
            return Err(Error::Limit {
                kind: LimitKind::TotalBytes,
                observed: total,
                maximum: limits.max_total_size,
            });
        }
        objects.push(object);
    }
    Objects::new(objects)
}

pub(crate) fn from_package(
    package: &Package,
    targets: &Targets,
    limits: Limits,
) -> Result<Objects, Error> {
    limits.validate()?;
    if targets.len() > limits.max_objects {
        return Err(Error::Limit {
            kind: LimitKind::Objects,
            observed: targets.len() as u64,
            maximum: limits.max_objects as u64,
        });
    }
    let mut objects = Vec::new();
    let mut total = 0u64;
    for target in targets {
        let object = package.object(target.clone(), limits)?;
        total = total
            .checked_add(object.compound().len() as u64)
            .ok_or_else(|| OleError::InvalidFormat("object total size overflow".into()))?;
        if total > limits.max_total_size {
            return Err(Error::Limit {
                kind: LimitKind::TotalBytes,
                observed: total,
                maximum: limits.max_total_size,
            });
        }
        objects.push(object);
    }
    Objects::new(objects)
}
