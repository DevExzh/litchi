//! Outside-clock correctness companion for the 0357b scoped PPT workload.
//!
//! This intentionally is not the timed scoped_patch_probe and is not part
//! of its measured workload. It opens the same fixture with the same public
//! package/edit/apply APIs, makes the exact clip-0000.avi operation once,
//! and writes an external receipt proving the target bytes, parsed path, and
//! inverse source bytes. Checksums and all publication serialization happen
//! only in this correctness mode.

use std::{env, fmt::Write as _, fs};

use litchi_ppt::{Package, external_media::Object};
use sha2::{Digest, Sha256};

const MEDIA_ID: u32 = 1;
const EXPECTED_PATH: &str = "clip-0000.avi";

fn sha256(bytes: &[u8]) -> String {
    let mut digest = String::with_capacity(64);
    for byte in Sha256::digest(bytes) {
        write!(&mut digest, "{byte:02x}").expect("writing to a String cannot fail");
    }
    digest
}

fn external_media_path(
    package: &Package,
    id: u32,
) -> Result<Option<String>, Box<dyn std::error::Error>> {
    let presentation = package.presentation()?;
    let collection = presentation
        .external_media()?
        .ok_or("presentation has no external-media collection")?;
    let object = collection
        .get(id)
        .ok_or_else(|| format!("external media ID {id} was not found"))?;
    match object {
        Object::Movie(movie) => Ok(movie.video.path.clone()),
        Object::LinkedAudio(audio) => Ok(audio.path.clone()),
        Object::CdAudio(_) | Object::EmbeddedWav(_) => {
            Err("external media object does not own a linked path".into())
        },
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = env::args().skip(1);
    let fixture = args
        .next()
        .ok_or("usage: scoped_patch_correctness FIXTURE RECEIPT")?;
    let receipt_path = args
        .next()
        .ok_or("usage: scoped_patch_correctness FIXTURE RECEIPT")?;
    if args.next().is_some() {
        return Err("usage: scoped_patch_correctness FIXTURE RECEIPT".into());
    }

    let source_bytes = fs::read(&fixture)?;
    let source = Package::from_bytes(&source_bytes)?;
    let before_path = external_media_path(&source, MEDIA_ID)?;

    let mut edit = source.edit_external_media()?;
    edit.set_path(MEDIA_ID, Some(EXPECTED_PATH.to_owned()))?;
    let commit = edit.commit()?;
    let (target, patch) = commit.into_parts();
    let target_bytes = target.to_bytes()?;
    let target_media_path = external_media_path(&target, MEDIA_ID)?;
    if target_media_path.as_deref() != Some(EXPECTED_PATH) {
        return Err(format!(
            "external-media path mismatch: expected {EXPECTED_PATH:?}, got {target_media_path:?}"
        )
        .into());
    }
    if target_bytes == source_bytes {
        return Err("the correctness operation produced unchanged target bytes".into());
    }

    let replayed = source.apply_external_media(&patch)?;
    let replayed_bytes = replayed.package().to_bytes()?;
    let replay_matches_target = replayed_bytes == target_bytes;
    if !replay_matches_target {
        return Err("forward patch replay did not reproduce the committed target bytes".into());
    }
    let restored = replayed.package().apply_external_media(&patch.inverse())?;
    let restored_bytes = restored.package().to_bytes()?;
    if restored_bytes != source_bytes {
        return Err("the inverse did not restore the source bytes".into());
    }

    let receipt = serde_json::json!({
        "mode": "scoped_patch_correctness",
        "fixture": fixture,
        "media_id": MEDIA_ID,
        "requested_external_media_path": EXPECTED_PATH,
        "before_external_media_path": before_path,
        "external_media_path": target_media_path,
        "source_len": source_bytes.len(),
        "target_len": target_bytes.len(),
        "source_sha256": sha256(&source_bytes),
        "target_sha256": sha256(&target_bytes),
        "replay_target_sha256": sha256(&replayed_bytes),
        "replay_matches_target": replay_matches_target,
        "inverse_source_sha256": sha256(&restored_bytes),
        "inverse_matches_source": restored_bytes == source_bytes,
    });
    fs::write(receipt_path, serde_json::to_vec_pretty(&receipt)?)?;
    Ok(())
}
