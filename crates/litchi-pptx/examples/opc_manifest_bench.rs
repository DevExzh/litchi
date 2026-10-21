//! Record 0356c: identical release probe for the raw OPC publication engine.
//! Build this source on the record's base and head, then run A-B-B-A pinned
//! to the same CPU. Every scenario prints median microseconds as JSON.

use std::hint::black_box;
use std::time::Instant;

use litchi_opc::{BlobPart, PackURI, TargetMode};
use litchi_pptx::{Package, Result};

fn measure(name: &str, mut run: impl FnMut() -> Result<()>) -> Result<()> {
    let mut samples = Vec::new();
    for index in 0..34 {
        let started = Instant::now();
        run()?;
        if index >= 3 {
            samples.push(started.elapsed().as_secs_f64() * 1e6);
        }
    }
    samples.sort_by(f64::total_cmp);
    println!(
        "{{\"scenario\":\"{name}\",\"median_us\":{}}}",
        samples[samples.len() / 2]
    );
    Ok(())
}

fn main() -> Result<()> {
    let bytes = include_bytes!("../../../test-data/poi/test-data/slideshow/45545_Comment.pptx");
    let package = Package::from_bytes(bytes)?;
    let name = PackURI::new("/ppt/presentation.xml").map_err(litchi_pptx::Error::Uri)?;
    let mut payload = package.opc()?.get_part(&name)?.blob().to_vec();
    payload.extend_from_slice(b"<!-- edited -->");
    measure("noop", || {
        let mut edit = package.edit_opc()?;
        edit.update(|_| Ok(()))?;
        drop(black_box(edit.commit()?));
        Ok(())
    })?;
    for (scenario, writes) in [
        ("payload_commit", 0),
        ("payload_save", 1),
        ("payload_save_twice", 2),
    ] {
        measure(scenario, || {
            let mut edit = package.edit_opc()?;
            edit.update(|opc| {
                opc.get_part_mut(&name)?.set_blob(payload.clone());
                Ok(())
            })?;
            let commit = edit.commit()?;
            for _ in 0..writes {
                black_box(commit.package().to_bytes()?);
            }
            drop(black_box(commit));
            Ok(())
        })?;
    }
    for (scenario, writes) in [("part_add_commit", 0), ("part_add_save", 1)] {
        measure(scenario, || {
            let mut edit = package.edit_opc()?;
            edit.update(|opc| {
                opc.try_add_part(Box::new(BlobPart::new(
                    PackURI::new("/custom/new.bin").map_err(litchi_pptx::Error::Uri)?,
                    "application/octet-stream".to_owned(),
                    vec![1, 2, 3],
                )))?;
                Ok(())
            })?;
            let commit = edit.commit()?;
            for _ in 0..writes {
                black_box(commit.package().to_bytes()?);
            }
            drop(black_box(commit));
            Ok(())
        })?;
    }
    measure("root_relationship_commit", || {
        let mut edit = package.edit_opc()?;
        edit.update(|opc| {
            opc.rels_mut().try_add_relationship(
                "urn:example:external".to_owned(),
                "https://example.invalid/".to_owned(),
                "external".to_owned(),
                TargetMode::External,
            )?;
            Ok(())
        })?;
        drop(black_box(edit.commit()?));
        Ok(())
    })?;
    let payload_commit = || {
        let mut edit = package.edit_opc()?;
        edit.update(|opc| {
            opc.get_part_mut(&name)?.set_blob(payload.clone());
            Ok(())
        })?;
        edit.commit()
    };
    let committed = payload_commit()?;
    measure("payload_replay_warm", || {
        drop(black_box(package.apply_opc(committed.patch())?));
        Ok(())
    })?;
    let inverse = committed.patch().inverse();
    measure("payload_inverse_warm", || {
        drop(black_box(committed.package().apply_opc(&inverse)?));
        Ok(())
    })?;
    measure("payload_commit_and_first_replay", || {
        let committed = payload_commit()?;
        drop(black_box(package.apply_opc(committed.patch())?));
        Ok(())
    })?;
    measure("payload_commit_and_first_inverse", || {
        let committed = payload_commit()?;
        drop(black_box(
            committed
                .package()
                .apply_opc(&committed.patch().inverse())?,
        ));
        Ok(())
    })?;
    measure("payload_patch_history_128", || {
        let mut patches = Vec::with_capacity(128);
        for _ in 0..128 {
            let committed = payload_commit()?;
            let (published, patch) = committed.into_parts();
            drop(published);
            patches.push(patch);
        }
        drop(black_box(patches));
        Ok(())
    })?;
    let add_commit = || {
        let mut edit = package.edit_opc()?;
        edit.update(|opc| {
            opc.try_add_part(Box::new(BlobPart::new(
                PackURI::new("/custom/new.bin").map_err(litchi_pptx::Error::Uri)?,
                "application/octet-stream".to_owned(),
                vec![1, 2, 3],
            )))?;
            Ok(())
        })?;
        edit.commit()
    };
    let added = add_commit()?;
    measure("part_add_replay_warm", || {
        drop(black_box(package.apply_opc(added.patch())?));
        Ok(())
    })?;
    let inverse = added.patch().inverse();
    measure("part_add_inverse_warm", || {
        drop(black_box(added.package().apply_opc(&inverse)?));
        Ok(())
    })?;
    measure("part_add_patch_history_128", || {
        let mut patches = Vec::with_capacity(128);
        for _ in 0..128 {
            let (published, patch) = add_commit()?.into_parts();
            drop(published);
            patches.push(patch);
        }
        drop(black_box(patches));
        Ok(())
    })?;
    measure("batch1000", || {
        let mut edit = package.edit_opc()?;
        for index in 0..1_000 {
            edit.update(|opc| {
                let mut changed = payload.clone();
                changed.extend_from_slice(format!("<!-- {index} -->").as_bytes());
                opc.get_part_mut(&name)?.set_blob(changed);
                Ok(())
            })?;
        }
        drop(black_box(edit.commit()?));
        Ok(())
    })
}
