//! Compare the decompressed IWA components in two iWork packages.
//!
//! This intentionally diagnostic example operates on the focused physical
//! archive model. It does not pull the Pages, Numbers, Keynote, or legacy
//! umbrella facades into low-level package comparison.

#![allow(
    clippy::print_stdout,
    reason = "This command-line diagnostic intentionally reports package differences."
)]

mod support;

use std::collections::BTreeSet;
use std::env;
use std::path::Path;

use litchi_iwa_archive::{ComponentCatalog, Limits, Result as ArchiveResult};

fn open(path: impl AsRef<Path>) -> ArchiveResult<ComponentCatalog> {
    let source = support::read_package(path.as_ref(), Limits::default())?;
    ComponentCatalog::from_bytes(source.as_ref())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let left_catalog = open(
        arguments
            .next()
            .ok_or("usage: compare_iwa_packages <left> <right>")?,
    )?;
    let right_catalog = open(
        arguments
            .next()
            .ok_or("usage: compare_iwa_packages <left> <right>")?,
    )?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let names = left_catalog
        .iter()
        .chain(right_catalog.iter())
        .map(|component| component.name().to_owned())
        .collect::<BTreeSet<_>>();
    let mut differences = 0usize;
    for name in names {
        let left_match = left_catalog.get(&name);
        let right_match = right_catalog.get(&name);
        let (Some(left_component), Some(right_component)) = (left_match, right_match) else {
            differences += 1;
            println!(
                "different member={name} left_present={} right_present={}",
                left_match.is_some(),
                right_match.is_some()
            );
            continue;
        };

        let left_archive = left_component.archive();
        let right_archive = right_component.archive();
        if left_archive.to_bytes()? == right_archive.to_bytes()? {
            continue;
        }
        differences += 1;
        println!("different member={name}");
        let identifiers = left_archive
            .objects
            .iter()
            .chain(&right_archive.objects)
            .filter_map(|object| object.archive_info.identifier)
            .collect::<BTreeSet<_>>();
        for identifier in identifiers {
            let left_object = left_archive.object(identifier);
            let right_object = right_archive.object(identifier);
            let equal = left_object
                .zip(right_object)
                .is_some_and(|(left_entry, right_entry)| {
                    left_entry.archive_info == right_entry.archive_info
                        && left_entry.messages == right_entry.messages
                });
            if !equal {
                println!(
                    "  object={identifier} left_types={:?} right_types={:?} payloads_equal={} metadata_equal={}",
                    left_object.map(|object| object
                        .messages
                        .iter()
                        .map(|message| message.type_)
                        .collect::<Vec<_>>()),
                    right_object.map(|object| object
                        .messages
                        .iter()
                        .map(|message| message.type_)
                        .collect::<Vec<_>>()),
                    left_object
                        .zip(right_object)
                        .is_some_and(|(left_entry, right_entry)| {
                            left_entry.messages == right_entry.messages
                        }),
                    left_object
                        .zip(right_object)
                        .is_some_and(|(left_entry, right_entry)| {
                            left_entry.archive_info == right_entry.archive_info
                        }),
                );
            }
        }
    }
    println!("different_decompressed_members={differences}");
    Ok(())
}
