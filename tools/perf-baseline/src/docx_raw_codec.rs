//! Allocator-only comparison of the DOCX raw codecs with their direct plan
//! lifecycles. The source and semantic gates are prepared outside each
//! measured region; the region covers the public operation and its result
//! drop. This is an opt-in diagnostic used by the allocator binary.

use crate::allocation_metrics::{self, Sample};
use litchi_core::{OwnedSource, ReadAt};
use litchi_docx::redact;
use litchi_docx::sanitize;
use litchi_docx::source_backed;
use litchi_docx::story_hyperlinks::{self, Mode};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, TargetMode};
use serde::Serialize;
use std::error::Error;
use std::ffi::OsString;
use std::fs;
use std::path::PathBuf;
use std::sync::Arc;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HYPERLINK: &str = "https://litchi-perf.invalid/raw-codec";
const WRAPPER: &[u8] = b"<w:hyperlink";

#[derive(Debug, Serialize)]
struct Receipt {
    harness: &'static str,
    lifecycle_baseline: &'static str,
    timing_scope: &'static str,
    samples: usize,
    warmup: usize,
    semantic_output_proof: bool,
    lifecycles: Vec<Lifecycle>,
}

#[derive(Debug, Serialize)]
struct Lifecycle {
    family: &'static str,
    operation: &'static str,
    measurements: Vec<Sample>,
}

struct Prepared {
    sanitize_snapshot: sanitize::Snapshot,
    redact_snapshot: redact::Snapshot,
    story_snapshot: story_hyperlinks::Snapshot,
    sanitize_limits: sanitize::Limits,
    story_limits: story_hyperlinks::Limits,
}

fn main_xml() -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:hyperlink r:id="rLink"><w:r><w:t>visible</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
    )
    .into_bytes()
}

fn relationships() -> Vec<sanitize::raw::Relationship> {
    vec![sanitize::raw::Relationship::new(
        "rLink",
        rt::HYPERLINK,
        HYPERLINK,
        true,
    )]
}

fn story_archive() -> Result<Vec<u8>, Box<dyn Error>> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml")?,
        ct::WML_DOCUMENT_MAIN.to_owned(),
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p>main</w:p></w:body></w:document>"#
        )
        .into_bytes(),
    );
    main.rels_mut().try_add_relationship(
        rt::HEADER.to_owned(),
        "header1.xml".to_owned(),
        "rHeader".to_owned(),
        TargetMode::Internal,
    )?;
    let mut header = BlobPart::new(
        PackURI::new("/word/header1.xml")?,
        ct::WML_HEADER.to_owned(),
        format!(
            r#"<w:hdr xmlns:w="{W}" xmlns:r="{R}"><w:p><w:hyperlink r:id="rStory"><w:r><w:t>story</w:t></w:r></w:hyperlink></w:p></w:hdr>"#
        )
        .into_bytes(),
    );
    header.rels_mut().try_add_relationship(
        rt::HYPERLINK.to_owned(),
        HYPERLINK.to_owned(),
        "rStory".to_owned(),
        TargetMode::External,
    )?;
    package.try_add_part(Box::new(header))?;
    package.try_add_part(Box::new(main))?;
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    Ok(PackageWriter::to_bytes(&package)?)
}

fn prepare() -> Result<Prepared, Box<dyn Error>> {
    let sanitize_limits = sanitize::Limits::default();
    let xml = main_xml();
    let sanitize_snapshot =
        sanitize::Snapshot::from_xml(xml.clone(), relationships(), sanitize_limits)?;
    let redact_snapshot = redact::Snapshot::from_xml(xml, relationships(), sanitize_limits)?;
    let archive = story_archive()?;
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(archive));
    let package = source_backed::Package::from_read_at(source)?;
    let story_limits = story_hyperlinks::Limits::default();
    let story_snapshot = package.story_hyperlinks_only_snapshot()?;
    Ok(Prepared {
        sanitize_snapshot,
        redact_snapshot,
        story_snapshot,
        sanitize_limits,
        story_limits,
    })
}

fn verify_semantics(prepared: &Prepared) -> Result<(), Box<dyn Error>> {
    let source_xml = prepared.sanitize_snapshot.xml_bytes().to_vec();
    let forward = prepared.sanitize_snapshot.plan().apply()?;
    let detached = forward.apply(&prepared.sanitize_snapshot)?;
    if detached
        .xml_bytes()
        .windows(WRAPPER.len())
        .any(|window| window == WRAPPER)
    {
        return Err("sanitizer semantic output retained the hyperlink wrapper".into());
    }
    let restored = forward.inverse().apply(&detached)?;
    if restored.xml_bytes() != source_xml.as_slice() {
        return Err("sanitizer inverse did not restore source XML".into());
    }

    let redaction = prepared
        .redact_snapshot
        .plan_target_urls(&[HYPERLINK])?
        .apply()?;
    let mut redaction_codec = redact::raw::Codec::from_snapshot(
        prepared.redact_snapshot.clone(),
        prepared.sanitize_limits,
    )?;
    redaction_codec.apply(&redaction)?;
    if redaction_codec
        .xml_bytes()
        .windows(WRAPPER.len())
        .any(|window| window == WRAPPER)
        || redaction_codec.removed_relationship_ids().is_empty()
    {
        return Err("redactor semantic output retained the hyperlink wrapper".into());
    }

    let story_patch = prepared
        .story_snapshot
        .plan_target_urls(&[HYPERLINK])?
        .apply()?;
    let mut story_codec = story_hyperlinks::raw::Codec::from_snapshot(
        prepared.story_snapshot.clone(),
        prepared.story_limits,
    )?;
    story_codec.apply(&story_patch)?;
    if story_codec.parts().iter().any(|part| {
        part.xml()
            .windows(WRAPPER.len())
            .any(|window| window == WRAPPER)
    }) || story_codec
        .parts()
        .iter()
        .all(|part| part.removed_relationship_ids().is_empty())
    {
        return Err("story redactor semantic output retained the hyperlink wrapper".into());
    }
    Ok(())
}

fn collect<F>(
    samples: usize,
    warmup: usize,
    mut operation: F,
) -> Result<Vec<Sample>, Box<dyn Error>>
where
    F: FnMut() -> Result<usize, Box<dyn Error>>,
{
    let total = warmup
        .checked_add(samples)
        .ok_or("warmup plus samples overflows usize")?;
    let mut measurements = Vec::with_capacity(samples);
    for iteration in 0..total {
        let region = allocation_metrics::begin();
        let result = operation();
        let measurement = region
            .finish()
            .ok_or("allocator region did not publish a measurement")?;
        let value = result?;
        std::hint::black_box(value);
        if iteration >= warmup {
            measurements.push(measurement);
        }
    }
    Ok(measurements)
}

fn parse(args: &[OsString]) -> Result<(usize, usize, PathBuf), Box<dyn Error>> {
    let mut samples = 5;
    let mut warmup = 1;
    let mut output = None;
    let mut args = args.iter();
    while let Some(argument) = args.next() {
        match argument.to_string_lossy().as_ref() {
            "--samples" => {
                samples = args
                    .next()
                    .ok_or("--samples requires a value")?
                    .to_string_lossy()
                    .parse()?;
            },
            "--warmup" => {
                warmup = args
                    .next()
                    .ok_or("--warmup requires a value")?
                    .to_string_lossy()
                    .parse()?;
            },
            "--output" => {
                output = Some(PathBuf::from(
                    args.next().ok_or("--output requires a path")?,
                ));
            },
            "--help" | "-h" => {
                println!("docx-raw-codec --samples N --warmup N --output PATH");
                return Err("help requested".into());
            },
            unknown => return Err(format!("unknown argument {unknown}").into()),
        }
    }
    if samples == 0 {
        return Err("--samples must be positive".into());
    }
    Ok((samples, warmup, output.ok_or("--output is required")?))
}

pub fn run_from_args(args: impl IntoIterator<Item = OsString>) -> Result<(), Box<dyn Error>> {
    let args = args.into_iter().collect::<Vec<_>>();
    let (samples, warmup, output) = parse(&args)?;
    let prepared = prepare()?;
    verify_semantics(&prepared)?;

    let mut lifecycles = Vec::new();
    lifecycles.push(Lifecycle {
        family: "sanitize",
        operation: "direct_plan_apply",
        measurements: collect(samples, warmup, || {
            let patch = prepared.sanitize_snapshot.plan().apply()?;
            let result = usize::from(patch.is_noop());
            drop(patch);
            Ok(result)
        })?,
    });
    lifecycles.push(Lifecycle {
        family: "sanitize",
        operation: "raw_codec_first_detach",
        measurements: collect(samples, warmup, || {
            let mut codec = sanitize::raw::Codec::from_snapshot(
                prepared.sanitize_snapshot.clone(),
                prepared.sanitize_limits,
            )?;
            let patch = codec.detach()?;
            let result = codec.xml_bytes().len();
            drop(patch);
            drop(codec);
            Ok(result)
        })?,
    });
    let forward = prepared.sanitize_snapshot.plan().apply()?;
    let detached = forward.apply(&prepared.sanitize_snapshot)?;
    let inverse = forward.inverse();
    lifecycles.push(Lifecycle {
        family: "sanitize",
        operation: "direct_inverse_apply",
        measurements: collect(samples, warmup, || {
            let restored = inverse.apply(&detached)?;
            let result = restored.xml_bytes().len();
            drop(restored);
            Ok(result)
        })?,
    });
    let mut after_codec = sanitize::raw::Codec::from_snapshot(
        prepared.sanitize_snapshot.clone(),
        prepared.sanitize_limits,
    )?;
    let raw_forward = after_codec.detach()?;
    lifecycles.push(Lifecycle {
        family: "sanitize",
        operation: "raw_codec_inverse_apply",
        measurements: collect(samples, warmup, || {
            let mut codec = after_codec.clone();
            codec.apply(&raw_forward.inverse())?;
            let result = codec.xml_bytes().len();
            drop(codec);
            Ok(result)
        })?,
    });
    let mut repeated_codec = after_codec.clone();
    lifecycles.push(Lifecycle {
        family: "sanitize",
        operation: "raw_codec_repeated_apply",
        measurements: collect(samples, warmup, || {
            if repeated_codec.apply(&raw_forward).is_ok() {
                return Err("sanitizer accepted a repeated forward patch".into());
            }
            Ok(repeated_codec.xml_bytes().len())
        })?,
    });
    lifecycles.push(Lifecycle {
        family: "redact",
        operation: "direct_plan_apply",
        measurements: collect(samples, warmup, || {
            let patch = prepared
                .redact_snapshot
                .plan_target_urls(&[HYPERLINK])?
                .apply()?;
            let result = usize::from(patch.is_noop());
            drop(patch);
            Ok(result)
        })?,
    });
    lifecycles.push(Lifecycle {
        family: "redact",
        operation: "raw_codec_first_redact",
        measurements: collect(samples, warmup, || {
            let mut codec = redact::raw::Codec::from_snapshot(
                prepared.redact_snapshot.clone(),
                prepared.sanitize_limits,
            )?;
            let patch = codec.redact(&[HYPERLINK])?;
            let result = codec.xml_bytes().len();
            drop(patch);
            drop(codec);
            Ok(result)
        })?,
    });
    let redaction_patch = prepared
        .redact_snapshot
        .plan_target_urls(&[HYPERLINK])?
        .apply()?;
    let mut redaction_after_codec = redact::raw::Codec::from_snapshot(
        prepared.redact_snapshot.clone(),
        prepared.sanitize_limits,
    )?;
    redaction_after_codec.apply(&redaction_patch)?;
    let mut repeated_redaction_codec = redaction_after_codec;
    lifecycles.push(Lifecycle {
        family: "redact",
        operation: "raw_codec_repeated_apply",
        measurements: collect(samples, warmup, || {
            if repeated_redaction_codec.apply(&redaction_patch).is_ok() {
                return Err("redactor accepted a repeated forward patch".into());
            }
            Ok(repeated_redaction_codec.xml_bytes().len())
        })?,
    });
    lifecycles.push(Lifecycle {
        family: "story_hyperlinks",
        operation: "direct_plan_apply",
        measurements: collect(samples, warmup, || {
            let patch = prepared
                .story_snapshot
                .plan_target_urls(&[HYPERLINK])?
                .apply()?;
            let result = usize::from(patch.is_noop());
            drop(patch);
            Ok(result)
        })?,
    });
    lifecycles.push(Lifecycle {
        family: "story_hyperlinks",
        operation: "raw_codec_first_redact",
        measurements: collect(samples, warmup, || {
            let mut codec = story_hyperlinks::raw::Codec::from_snapshot(
                prepared.story_snapshot.clone(),
                prepared.story_limits,
            )?;
            let patch = codec.redact(&[HYPERLINK], Mode::Strict)?;
            let result = codec.parts().len();
            drop(patch);
            drop(codec);
            Ok(result)
        })?,
    });
    let story_patch = prepared
        .story_snapshot
        .plan_target_urls(&[HYPERLINK])?
        .apply()?;
    let mut story_after_codec = story_hyperlinks::raw::Codec::from_snapshot(
        prepared.story_snapshot.clone(),
        prepared.story_limits,
    )?;
    story_after_codec.apply(&story_patch)?;
    let mut repeated_story_codec = story_after_codec;
    lifecycles.push(Lifecycle {
        family: "story_hyperlinks",
        operation: "raw_codec_repeated_apply",
        measurements: collect(samples, warmup, || {
            if repeated_story_codec.apply(&story_patch).is_ok() {
                return Err("story redactor accepted a repeated forward patch".into());
            }
            Ok(repeated_story_codec.parts().len())
        })?,
    });

    let receipt = Receipt {
        harness: "docx-raw-codec-allocator-v1",
        lifecycle_baseline: "direct immutable Snapshot::plan().apply()",
        timing_scope: "allocator region covers each operation and result drop; source ingress, snapshot construction, semantic output proof, and report serialization are outside",
        samples,
        warmup,
        semantic_output_proof: true,
        lifecycles,
    };
    fs::write(output, serde_json::to_vec_pretty(&receipt)?)?;
    Ok(())
}
