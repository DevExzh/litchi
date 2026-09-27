//! Change 0741 pre-change durable-byte fixture.
//!
//! This is deliberately a packet-only generator.  It builds a tiny owned
//! cross-slide copy using the public PPTX/OPC APIs from the unchanged tree,
//! makes the source image's physical member Store-compressed, and records the
//! V3 durable forward/inverse patches plus the baseline target archive.
//!
//! Usage:
//!
//! ```text
//! cargo run --release --manifest-path \
//!   docs/performance/results/change-0741/legacy-fixture/Cargo.toml -- \
//!   docs/performance/results/change-0741/legacy-fixture/captured
//! ```

use std::error::Error;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::phys_pkg::{PhysPkgReader, PhysPkgWriter};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
use litchi_pptx::{CrossSlideCopyPatch, Package};
use sha2::{Digest, Sha256};

const SOURCE_MEDIA_MEMBER: &str = "ppt/media/change-0741-store-image.png";
const SOURCE_MEDIA_URI: &str = "/ppt/media/change-0741-store-image.png";
const SOURCE_SLIDE_URI: &str = "/ppt/slides/slide1.xml";
const SOURCE_IMAGE_RELATIONSHIP: &str = "rIdChange0741Image";
const SOURCE_IMAGE_PAYLOAD: &[u8] = &[
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0,
    0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 240, 31, 0,
    5, 0, 1, 255, 137, 153, 61, 29, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
];

type BoxResult<T> = Result<T, Box<dyn Error + Send + Sync>>;

fn main() -> BoxResult<()> {
    let output = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .ok_or("usage: change0741-legacy-fixture <output-directory>")?;
    fs::create_dir_all(&output)?;

    let source = source_fixture()?;
    let destination = destination_fixture()?;
    assert_source_image_is_stored(&source)?;

    let source_package = Package::from_vec(source.clone())?;
    let destination_for_plan = Package::from_vec(destination.clone())?;
    let source_snapshot = source_package.opened_presentation()?;
    let destination_snapshot = destination_for_plan.opened_presentation()?;
    let plan = destination_snapshot.plan_cross_slide_copy(&source_snapshot, 0_usize, 0_usize, 1)?;

    let forward = plan.patch().to_bytes()?;
    let forward_durable = CrossSlideCopyPatch::from_bytes(&forward)?;
    let inverse = forward_durable.inverse().to_bytes()?;
    let inverse_durable = CrossSlideCopyPatch::from_bytes(&inverse)?;

    let mut target = Package::from_vec(destination.clone())?;
    target.apply_cross_slide_copy_patch(&source_package, &forward_durable)?;
    let expected_target = target.to_bytes()?;

    // Prove that the recorded inverse is usable against the captured target
    // before publishing the packet.  This also catches accidental fixture
    // dependence on a plan-only retained candidate.
    let mut restored = Package::from_vec(expected_target.clone())?;
    restored.apply_cross_slide_copy_patch(&source_package, &inverse_durable)?;
    let restored_bytes = restored.to_bytes()?;
    if restored_bytes != destination {
        return Err("baseline inverse did not restore destination bytes".into());
    }

    write_artifact(&output, "source.pptx", &source)?;
    write_artifact(&output, "destination.pptx", &destination)?;
    write_artifact(&output, "forward.patch", &forward)?;
    write_artifact(&output, "inverse.patch", &inverse)?;
    write_artifact(&output, "expected-target.pptx", &expected_target)?;
    write_manifest(
        &output,
        &[
            ("source.pptx", &source),
            ("destination.pptx", &destination),
            ("forward.patch", &forward),
            ("inverse.patch", &inverse),
            ("expected-target.pptx", &expected_target),
        ],
    )?;

    println!("captured {}", output.display());
    println!("source_bytes\t{}", source.len());
    println!("destination_bytes\t{}", destination.len());
    println!("forward_patch_bytes\t{}", forward.len());
    println!("inverse_patch_bytes\t{}", inverse.len());
    println!("expected_target_bytes\t{}", expected_target.len());
    println!("source_media_member\t{SOURCE_MEDIA_MEMBER}");
    println!("source_media_compression\tstore");
    println!("forward_patch_magic\tLPCP0003");
    println!("inverse_patch_magic\tLPCP0003");
    Ok(())
}

fn source_fixture() -> BoxResult<Vec<u8>> {
    let bytes = authored_package(&["Change 0741 source"])?;
    let mut opc = OpcPackage::from_vec(bytes)?;
    opc.try_add_part(Box::new(BlobPart::new(
        pack_uri(SOURCE_MEDIA_URI)?,
        ct::PNG.to_owned(),
        SOURCE_IMAGE_PAYLOAD.to_vec(),
    )))?;
    opc.get_part_mut(&pack_uri(SOURCE_SLIDE_URI)?)?
        .rels_mut()
        .try_add_relationship(
            rt::IMAGE.to_owned(),
            "../media/change-0741-store-image.png".to_owned(),
            SOURCE_IMAGE_RELATIONSHIP.to_owned(),
            TargetMode::Internal,
        )?;
    append_picture(&mut opc)?;

    // PackageWriter writes ordinary generated members through its normal
    // baseline path.  Repack the complete owned archive through the public
    // physical writer so the source media member is unambiguously Store.
    let generated = PackageWriter::to_bytes(&opc)?;
    repack_all_stored(&generated)
}

fn destination_fixture() -> BoxResult<Vec<u8>> {
    authored_package(&[
        "Change 0741 destination first",
        "Change 0741 destination last",
    ])
}

fn authored_package(names: &[&str]) -> BoxResult<Vec<u8>> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        for name in names {
            let slide = presentation.add_slide()?;
            slide.set_title(name);
            slide.add_text_box(&format!("change-0741-body:{name}"), 10, 20, 300, 400);
        }
    }
    let bytes = package.to_bytes()?;
    let mut opc = OpcPackage::from_vec(bytes)?;
    for (index, name) in names.iter().enumerate() {
        let slide_uri = pack_uri(&format!("/ppt/slides/slide{}.xml", index + 1))?;
        let part = opc.get_part_mut(&slide_uri)?;
        let xml = std::str::from_utf8(part.blob())?;
        let marker = "<p:cSld name=\"";
        let marker_start = xml
            .find(marker)
            .ok_or("authored slide has no producer-visible cSld name")?;
        let value_start = marker_start + marker.len();
        let value_end = value_start
            + xml[value_start..]
                .find('"')
                .ok_or("producer-visible cSld name is unterminated")?;
        let updated = format!("{}{}{}", &xml[..value_start], name, &xml[value_end..]);
        part.set_blob(updated.into_bytes());
    }
    Ok(PackageWriter::to_bytes(&opc)?)
}

fn append_picture(opc: &mut OpcPackage) -> BoxResult<()> {
    let slide_uri = pack_uri(SOURCE_SLIDE_URI)?;
    let part = opc.get_part_mut(&slide_uri)?;
    let xml = std::str::from_utf8(part.blob())?;
    let marker = "</p:spTree>";
    let insertion = xml.rfind(marker).ok_or("source slide has no shape tree")?;
    let picture = format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="42" name="Change 0741 Store Image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{SOURCE_IMAGE_RELATIONSHIP}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
    );
    let updated = format!("{}{picture}{}", &xml[..insertion], &xml[insertion..]);
    part.set_blob(updated.into_bytes());
    Ok(())
}

fn repack_all_stored(bytes: &[u8]) -> BoxResult<Vec<u8>> {
    let reader = PhysPkgReader::new(bytes)?;
    let names = reader.member_names()?;
    let mut writer = PhysPkgWriter::new();
    for name in names {
        let uri = pack_uri(&format!("/{name}"))?;
        let payload = reader.read_member(&name)?;
        writer.write_stored(&uri, &payload)?;
    }
    Ok(writer.finish()?)
}

fn assert_source_image_is_stored(bytes: &[u8]) -> BoxResult<()> {
    let reader = PhysPkgReader::new(bytes)?;
    let uri = pack_uri(SOURCE_MEDIA_URI)?;
    if reader.blob_for_borrowed(&uri)?.is_none() {
        return Err("source image is not a borrowable ZIP Store member".into());
    }
    if reader.read_member(SOURCE_MEDIA_MEMBER)? != SOURCE_IMAGE_PAYLOAD {
        return Err("source image payload changed during Store repack".into());
    }
    Ok(())
}

fn pack_uri(value: &str) -> BoxResult<PackURI> {
    PackURI::new(value).map_err(|error| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            format!("invalid fixture PackURI {value:?}: {error}"),
        )
        .into()
    })
}

fn write_artifact(output: &Path, name: &str, bytes: &[u8]) -> BoxResult<()> {
    fs::write(output.join(name), bytes)?;
    Ok(())
}

fn write_manifest(output: &Path, artifacts: &[(&str, &[u8])]) -> BoxResult<()> {
    let mut manifest = String::from(
        "change=0741\n\
fixture=owned-cross-copy-store-image\n\
baseline=unchanged-source-tree\n\
source_media_member=ppt/media/change-0741-store-image.png\n\
source_media_compression=store\n\
patch_format=LPCP0003\n",
    );
    manifest.push_str("artifact\tbytes\tsha256\n");
    for (name, bytes) in artifacts {
        let digest = Sha256::digest(bytes);
        manifest.push_str(&format!(
            "{name}\t{}\t{}\n",
            bytes.len(),
            hex_digest(&digest),
        ));
    }
    fs::write(output.join("MANIFEST.tsv"), manifest)?;
    Ok(())
}

fn hex_digest(digest: &[u8]) -> String {
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        output.push_str(&format!("{byte:02x}"));
    }
    output
}
