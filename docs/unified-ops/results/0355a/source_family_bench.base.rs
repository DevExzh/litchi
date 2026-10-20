//! Record 0355a, base variant: the same scenarios as
//! `crates/litchi-docx/examples/source_family_bench.rs`, written through the
//! `publish_*_to_stream` publishers of the record-0355 tree (3e686506f3).
//! Copy it over that example in a checkout of 3e686506f3 to build it.
//!
//! Each scenario opens a source-backed package over an in-memory archive
//! (1,000 paragraphs, a header, settings, one external hyperlink and a
//! 2 MiB stored image), makes one edit, commits it and writes the committed
//! package to a pre-sized `Vec`. The base variant of this file (in
//! `docs/unified-ops/results/0355a/`) makes the same edits through the
//! `publish_*_to_stream` publishers of the record-0355 tree.
//!
//! ```text
//! cargo build --release -p litchi-docx --example source_family_bench
//! taskset -c 27 target/release/examples/source_family_bench
//! ```
//!
//! Output: one `docx-family <scenario>: median N ns, min N ns, output N
//! bytes, sha256 H` line per scenario (the digest proves byte identity
//! between the two variants).

use std::error::Error;
use std::hint::black_box;
use std::sync::Arc;
use std::time::Instant;

use litchi_core::{OwnedSource, Position};
use litchi_docx::section::{Columns, Emu, Selector as SectionSelector};
use litchi_docx::source_backed::{self, StorySelector};
use litchi_docx::story_hyperlinks::Mode;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part as _, TargetMode};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PARAGRAPHS: usize = 1_000;
const URL: &str = "https://example.com/redacted";
const WARMUP: usize = 3;
const SAMPLES: usize = 15;

type BenchResult<T> = Result<T, Box<dyn Error>>;

/// The fixture; `link` adds the external hyperlink, the section
/// properties and the header reference (the plain-paragraph copy refuses
/// them, so the copy and the append run on the plain fixture).
fn fixture(link: bool) -> BenchResult<Vec<u8>> {
    let mut body = String::new();
    for index in 0..PARAGRAPHS {
        body.push_str(&format!(
            "<w:p><w:r><w:t>paragraph {index} of the source-family fixture</w:t></w:r></w:p>"
        ));
    }
    if link {
        body.push_str(
            r#"<w:p><w:hyperlink r:id="rLink"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>"#,
        );
    }
    // The plain-paragraph copy accepts only direct plain paragraphs in the
    // body: the plain fixture has no section properties or `r` namespace.
    if link {
        body.push_str(
            r#"<w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#,
        );
    }
    let namespaces = if link {
        format!(r#"xmlns:w="{W}" xmlns:r="{R}""#)
    } else {
        format!(r#"xmlns:w="{W}""#)
    };
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml")?,
        ct::WML_DOCUMENT_MAIN.to_owned(),
        format!(r#"<w:document {namespaces}><w:body>{body}</w:body></w:document>"#).into_bytes(),
    );
    for (kind, target, id, mode) in [
        (rt::HEADER, "header1.xml", "rHeader", TargetMode::Internal),
        (
            rt::SETTINGS,
            "settings.xml",
            "rSettings",
            TargetMode::Internal,
        ),
        (
            rt::IMAGE,
            "media/image1.png",
            "rImage",
            TargetMode::Internal,
        ),
        (rt::HYPERLINK, URL, "rLink", TargetMode::External),
    ] {
        if kind == rt::HYPERLINK && !link {
            continue;
        }
        main.rels_mut().try_add_relationship(
            kind.to_owned(),
            target.to_owned(),
            id.to_owned(),
            mode,
        )?;
    }
    package.try_add_part(Box::new(main))?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/word/header1.xml")?,
        ct::WML_HEADER.to_owned(),
        format!(r#"<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>header</w:t></w:r></w:p></w:hdr>"#)
            .into_bytes(),
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/word/settings.xml")?,
        ct::WML_SETTINGS.to_owned(),
        format!(
            r#"<w:settings xmlns:w="{W}"><w:docVars><w:docVar w:name="kept" w:val="1"/></w:docVars></w:settings>"#
        )
        .into_bytes(),
    )))?;
    // An incompressible 2 MiB payload the publication must raw-copy.
    let mut state = 0x9e37_79b9_7f4a_7c15_u64;
    let image = (0..2 << 20)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            state as u8
        })
        .collect::<Vec<u8>>();
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/word/media/image1.png")?,
        ct::PNG.to_owned(),
        image,
    )))?;
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    Ok(PackageWriter::to_bytes(&package)?)
}

fn open(source: &[u8]) -> BenchResult<source_backed::Package> {
    Ok(source_backed::Package::from_read_at(Arc::new(
        OwnedSource::new(source.to_vec()),
    ))?)
}

type Scenario = fn(&[u8], &mut Vec<u8>) -> BenchResult<()>;

fn document_text(source: &[u8], sink: &mut Vec<u8>) -> BenchResult<()> {
    let package = open(source)?;
    let mut edit = package.edit()?;
    edit.replace_paragraph_text(Position::new(PARAGRAPHS / 2), "changed")?;
    edit.commit()?.package().write_to(sink)?;
    Ok(())
}

fn document_detached(source: &[u8], sink: &mut Vec<u8>) -> BenchResult<()> {
    let package = open(source)?;
    let mut edit = package.document_snapshot()?.edit();
    edit.replace_paragraph_text(Position::new(PARAGRAPHS / 2), "changed")?;
    let commit = edit.commit()?;
    package.publish_document_commit_to_stream(&mut *sink, &commit)?;
    Ok(())
}

fn document_variables(source: &[u8], sink: &mut Vec<u8>) -> BenchResult<()> {
    let package = open(source)?;
    let mut edit = package.edit_document_variables()?;
    edit.set_variable("bench", "value")?;
    let commit = edit.commit()?;
    package.publish_document_variables_commit_to_stream(&mut *sink, &commit)?;
    Ok(())
}

fn section_layout(source: &[u8], sink: &mut Vec<u8>) -> BenchResult<()> {
    let package = open(source)?;
    let snapshot = package.section_layout_snapshot()?;
    let mut edit = snapshot.edit(SectionSelector::body_final())?;
    edit.set_columns(Some(Columns {
        equal_width: true,
        count: 2,
        space: Some(Emu::from_twips(240)),
        separator: false,
        columns: Vec::new(),
    }))?;
    let commit = edit.commit()?;
    package.publish_section_layout_commit_to_stream(&mut *sink, &commit)?;
    Ok(())
}

fn header_text(source: &[u8], sink: &mut Vec<u8>) -> BenchResult<()> {
    let package = open(source)?;
    let snapshot = package.story_text_snapshot(StorySelector::Header(0))?;
    let mut edit = snapshot.edit()?;
    edit.replace_paragraph_text(Position::new(0), "changed header")?;
    let commit = edit.commit()?;
    package.publish_story_text_commit_to_stream(&mut *sink, &commit)?;
    Ok(())
}

fn paragraph_copy(source: &[u8], sink: &mut Vec<u8>) -> BenchResult<()> {
    let package = open(source)?;
    let snapshot = package.plain_paragraph_copy_snapshot()?;
    let mut edit = snapshot.edit();
    edit.copy_plain_paragraph(Position::new(0), Position::new(PARAGRAPHS / 2))?;
    let commit = edit.commit();
    package.publish_plain_paragraph_copy_to_stream(&mut *sink, &commit)?;
    Ok(())
}

fn tail_append(source: &[u8], sink: &mut Vec<u8>) -> BenchResult<()> {
    let package = open(source)?;
    let plan = package.tail_append_plain_paragraph("appended").prepare()?;
    plan.write_to(&mut *sink)?;
    Ok(())
}

fn story_hyperlink_redaction(source: &[u8], sink: &mut Vec<u8>) -> BenchResult<()> {
    let package = open(source)?;
    let commit = package
        .plan_story_hyperlink_redaction(&[URL], Mode::Strict)?
        .apply()?;
    package.publish_story_hyperlink_redaction_to_stream(&mut *sink, &commit)?;
    Ok(())
}

fn main() -> BenchResult<()> {
    let linked = fixture(true)?;
    let plain = fixture(false)?;
    let scenarios: [(&str, bool, Scenario); 8] = [
        ("document paragraph text (edit)", true, document_text),
        (
            "document paragraph text (detached commit)",
            true,
            document_detached,
        ),
        ("document variables set", true, document_variables),
        ("section layout columns", true, section_layout),
        ("header story text", true, header_text),
        ("plain paragraph copy", false, paragraph_copy),
        ("tail append one paragraph", false, tail_append),
        ("story hyperlink redaction", true, story_hyperlink_redaction),
    ];
    for (name, link, scenario) in scenarios {
        let source = if link { &linked } else { &plain };
        let mut sink = Vec::with_capacity(source.len() * 2);
        for _ in 0..WARMUP {
            sink.clear();
            scenario(source, &mut sink)?;
        }
        let mut samples = Vec::with_capacity(SAMPLES);
        for _ in 0..SAMPLES {
            sink.clear();
            let started = Instant::now();
            scenario(source, &mut sink)?;
            samples.push(started.elapsed().as_nanos());
            black_box(&sink);
        }
        if sink.len() < source.len() / 2 {
            return Err(format!("{name}: output is truncated").into());
        }
        samples.sort_unstable();
        let digest = <sha2::Sha256 as sha2::Digest>::digest(&sink);
        let digest = digest
            .iter()
            .map(|byte| format!("{byte:02x}"))
            .collect::<String>();
        println!(
            "docx-family {name}: median {} ns, min {} ns, output {} bytes, sha256 {digest}",
            samples[SAMPLES / 2],
            samples[0],
            sink.len()
        );
    }
    Ok(())
}
