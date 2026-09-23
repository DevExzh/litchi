//! Change 0742 legacy durable-patch fixture generator.
//!
//! Built against the unchanged base tree (009d515bef), it plans an owned
//! cross-presentation slide copy whose closure carries one PNG picture and
//! writes the resulting `LPCP0003` forward and inverse durable patches. The
//! authoring steps are deterministic and are repeated verbatim by the branch
//! test `legacy_lpcp0003_patches_are_refused_by_name`, which checks the
//! source and destination SHA-256 recorded here before relying on them.
//!
//! Usage: `change0742-legacy-fixture <output-directory>`

use std::error::Error;
use std::fs;
use std::path::PathBuf;

use litchi_pptx::media_parts::Resource;
use litchi_pptx::{CrossSlideCopyPatch, Package};
use sha2::{Digest, Sha256};

type BoxResult<T> = Result<T, Box<dyn Error>>;

fn image_payload() -> Vec<u8> {
    let mut bytes = Vec::with_capacity(64);
    let mut state = 0x0742_u32;
    bytes.extend_from_slice(b"\x89PNG\r\n\x1a\n");
    while bytes.len() < 64 {
        state = state.wrapping_mul(1_103_515_245).wrapping_add(12_345);
        bytes.push((state >> 16) as u8);
    }
    bytes
}

fn authored(prefix: &str, slides: usize) -> BoxResult<Vec<u8>> {
    let mut package = Package::new()?;
    let presentation = package.presentation_mut()?;
    for index in 0..slides {
        presentation
            .add_slide()?
            .set_title(&format!("change-0742-{prefix}-{index}"));
    }
    Ok(package.to_bytes()?)
}

fn source_archive() -> BoxResult<Vec<u8>> {
    let mut package = Package::from_bytes(&authored("source", 3)?)?;
    let mut edit = package.opened_presentation_transaction()?;
    edit.add_picture(
        2_usize,
        "change-0742-legacy-picture",
        &Resource::new("/ppt/media/change-0742-legacy.png", "image/png", image_payload()),
        (800, 800, 72, 72),
    )?;
    let commit = edit.commit()?;
    package.apply_opened_presentation_commit(commit)?;
    Ok(package.to_bytes()?)
}

fn hex(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn main() -> BoxResult<()> {
    let output = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .ok_or("usage: change0742-legacy-fixture <output-directory>")?;
    fs::create_dir_all(&output)?;
    let source = source_archive()?;
    let destination = authored("destination", 2)?;

    let source_package = Package::from_vec(source.clone())?;
    let destination_package = Package::from_vec(destination.clone())?;
    let plan = destination_package
        .opened_presentation()?
        .plan_cross_slide_copy(&source_package.opened_presentation()?, 2_usize, 1_usize, 1)?;
    if !plan
        .parts()
        .iter()
        .any(|part| part.content_type() == "image/png")
    {
        return Err("the legacy copy closure carries no PNG member".into());
    }
    let forward = plan.patch().to_bytes()?;
    let durable = CrossSlideCopyPatch::from_bytes(&forward)?;
    let inverse = durable.inverse().to_bytes()?;
    if &forward[..8] != b"LPCP0003" || &inverse[..8] != b"LPCP0003" {
        return Err("the base tree did not write LPCP0003".into());
    }

    let mut target = Package::from_vec(destination.clone())?;
    target.apply_cross_slide_copy_patch(&source_package, &durable)?;
    let target_bytes = target.to_bytes()?;
    let mut restored = Package::from_vec(target_bytes.clone())?;
    restored.apply_cross_slide_copy_patch(&source_package, &CrossSlideCopyPatch::from_bytes(&inverse)?)?;
    if restored.to_bytes()? != destination {
        return Err("the base inverse did not restore the destination".into());
    }

    fs::write(output.join("lpcp0003-forward.patch"), &forward)?;
    fs::write(output.join("lpcp0003-inverse.patch"), &inverse)?;
    let mut manifest = String::from("artifact\tbytes\tsha256\n");
    for (name, bytes) in [
        ("source.pptx", &source),
        ("destination.pptx", &destination),
        ("expected-target.pptx", &target_bytes),
        ("lpcp0003-forward.patch", &forward),
        ("lpcp0003-inverse.patch", &inverse),
    ] {
        manifest.push_str(&format!("{name}\t{}\t{}\n", bytes.len(), hex(bytes)));
    }
    fs::write(output.join("MANIFEST.tsv"), &manifest)?;
    print!("{manifest}");
    Ok(())
}
