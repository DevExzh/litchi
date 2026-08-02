use std::collections::BTreeSet;
use std::env;

use litchi_iwa::IWorkPackage;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let left = IWorkPackage::open(
        arguments
            .next()
            .ok_or("usage: compare_iwa_packages <left> <right>")?,
    )?;
    let right = IWorkPackage::open(
        arguments
            .next()
            .ok_or("usage: compare_iwa_packages <left> <right>")?,
    )?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let names = left
        .iwa_entry_names()
        .chain(right.iwa_entry_names())
        .map(str::to_owned)
        .collect::<BTreeSet<_>>();
    let mut differences = 0usize;
    for name in names {
        let left_present = left.contains_entry(&name);
        let right_present = right.contains_entry(&name);
        if !left_present || !right_present {
            differences += 1;
            println!(
                "different member={name} left_present={left_present} right_present={right_present}"
            );
            continue;
        }
        let left_archive = left.archive(&name)?;
        let right_archive = right.archive(&name)?;
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
            let equal = left_object.zip(right_object).is_some_and(|(left, right)| {
                left.archive_info == right.archive_info && left.messages == right.messages
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
                        .is_some_and(|(left, right)| left.messages == right.messages),
                    left_object
                        .zip(right_object)
                        .is_some_and(|(left, right)| left.archive_info == right.archive_info),
                );
            }
        }
    }
    println!("different_decompressed_members={differences}");
    Ok(())
}
