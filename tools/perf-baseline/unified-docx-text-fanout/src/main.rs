//! Paired prepared-owner paragraph fanout, with retained commits on both legs.
use litchi_core::{FileFormat, Position};
use litchi_docx::document::ParagraphTextReplacement;
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
use litchi_ops::{OpenOptions, Props, SourceMode};
use sha2::{Digest, Sha256};
use std::{error::Error, hint::black_box, time::Instant};
type Result<T> = std::result::Result<T, Box<dyn Error>>;

const CORPORA: [(usize, usize, &str); 5] = [
    (
        1,
        944,
        "22e3a46ca46acaa6a4ef44f6d5d1e0b1ad293a25440a9c0da7ccb929a191c41c",
    ),
    (
        10,
        952,
        "8f31b5239bcef2326b148c73ff398e2ebaba8cb3676192ca6319dd781db174db",
    ),
    (
        100,
        974,
        "dc87796e6ecaecf1d640031ac4faf9d5451dead80c2823b4443902b9d585be13",
    ),
    (
        1000,
        1132,
        "b31a23d2d24eccdbd1b010c7d67087a3559663586528ecb24c942aa88914d465",
    ),
    (
        10000,
        2582,
        "4fd82dfe9ceb190bd057ec797ca78cf1a88e3b27eb707a762612e2dd9e16c5ba",
    ),
];

fn corpus(count: usize) -> Result<Vec<u8>> {
    let xml = format!(
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{}<w:sectPr/></w:body></w:document>",
        "<w:p><w:r><w:t>seed</w:t></w:r></w:p>".repeat(count)
    );
    let mut package = OpcPackage::new();
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml")?,
        litchi_opc::constants::content_type::WML_DOCUMENT_MAIN.to_owned(),
        xml.into_bytes(),
    )))?;
    package.relate_to(
        "word/document.xml",
        litchi_opc::constants::relationship_type::OFFICE_DOCUMENT,
    );
    Ok(PackageWriter::to_bytes(&package)?)
}
fn leaf(
    base: &litchi_docx::Package,
    input: &[ParagraphTextReplacement],
) -> Result<litchi_docx::document::Commit> {
    let mut edit = base.edit()?;
    edit.replace_body_paragraph_texts(input)?;
    Ok(edit.commit()?)
}
fn source_leaf(
    base: &litchi_docx::source_backed::Package,
    input: &[ParagraphTextReplacement],
) -> Result<litchi_docx::document::Commit<litchi_docx::source_backed::Package>> {
    let mut edit = base.edit()?;
    edit.replace_body_paragraph_texts(input)?;
    Ok(edit.commit()?)
}
fn median(values: &[u128]) -> u128 {
    let mut sorted = values.to_vec();
    sorted.sort_unstable();
    sorted[sorted.len() / 2]
}
fn measure<T, H>(
    unified: &litchi_ops::File,
    native: &T,
    input: &[ParagraphTextReplacement],
    rounds: usize,
    leaf: fn(&T, &[ParagraphTextReplacement]) -> Result<H>,
    bytes: fn(&H, usize) -> Result<Vec<u8>>,
) -> Result<serde_json::Value> {
    let query = litchi_ops::query::Query::parse("/body/p")?;
    let plan = litchi_ops::query::Plan::compile(&query, FileFormat::Docx)?;
    let props = Props::new().with("text", "updated\ttext\nline");
    let oracle = leaf(native, input);
    let candidate = unified.edit(|edit| edit.set(&plan, &props));
    if input.len() > 4096 {
        let error = oracle.err().ok_or("native limit unexpectedly absent")?;
        if !matches!(
            error.downcast_ref::<litchi_docx::document::Error>(),
            Some(litchi_docx::document::Error::Limit {
                kind: litchi_docx::document::LimitKind::CompositeParagraphReplacements,
                ..
            })
        ) {
            return Err(error);
        }
        let error = candidate.err().ok_or("unified limit unexpectedly absent")?;
        if error.code() != litchi_ops::Code::LimitExceeded {
            return Err(error.into());
        }
        return Ok(
            serde_json::json!({"targets":input.len(),"status":"refused_native_composite_replacement_limit","timed":false}),
        );
    }
    let oracle = oracle?;
    let candidate = candidate?;
    if bytes(&oracle, input.len())? != candidate.file().to_bytes()?
        || candidate.diagnostics().leaf_operations() != input.len()
        || candidate.diagnostics().operations() != 1
    {
        return Err("output or operation count mismatch".into());
    }
    drop((oracle, candidate));
    let mut leaf_ns = Vec::new();
    let mut unified_ns = Vec::new();
    for _ in 0..rounds {
        for native_leg in [true, false, false, true] {
            if native_leg {
                let start = Instant::now();
                let output = leaf(black_box(native), black_box(input))?;
                leaf_ns.push(start.elapsed().as_nanos());
                black_box(&output);
                drop(output);
            } else {
                let start = Instant::now();
                let output = black_box(unified)
                    .edit(|edit| edit.set(black_box(&plan), black_box(&props)))?;
                unified_ns.push(start.elapsed().as_nanos());
                black_box(&output);
                drop(output);
            }
        }
    }
    Ok(
        serde_json::json!({"targets":input.len(),"status":"ok","byte_equivalence":true,"leaf_median_ns":median(&leaf_ns),"unified_median_ns":median(&unified_ns),"unified_over_leaf":median(&unified_ns) as f64/median(&leaf_ns) as f64,"leaf_ns":leaf_ns,"unified_ns":unified_ns}),
    )
}
fn main() -> Result<()> {
    let rounds = std::env::args()
        .nth(1)
        .map(|s| s.parse::<usize>())
        .transpose()?
        .unwrap_or(9);
    if rounds < 3 {
        return Err("at least three ABBA rounds required".into());
    }
    let mut results = Vec::new();
    for (count, expected_size, expected_hash) in CORPORA {
        let corpus = corpus(count)?;
        let hash = Sha256::digest(&corpus)
            .iter()
            .map(|byte| format!("{byte:02x}"))
            .collect::<String>();
        if corpus.len() != expected_size || hash != expected_hash {
            return Err("frozen corpus changed; define new routes before measuring".into());
        }
        let input: Vec<_> = (0..count)
            .map(|index| ParagraphTextReplacement::new(Position::new(index), "updated\ttext\nline"))
            .collect();
        let unified = litchi_ops::open_bytes_with(
            corpus.clone(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )?;
        let native = litchi_docx::Package::from_bytes(&corpus)?;
        let eager = measure(&unified, &native, &input, rounds, leaf, |commit, count| {
            if commit.diagnostics().operations() != count {
                return Err("native ledger count mismatch".into());
            }
            Ok(commit.package().clone().to_bytes()?)
        })?;
        let unified = litchi_ops::open_bytes(corpus.clone())?;
        let native = litchi_docx::source_backed::Package::from_bytes(&corpus)?;
        let source = measure(
            &unified,
            &native,
            &input,
            rounds,
            source_leaf,
            |commit, count| {
                if commit.diagnostics().operations() != count {
                    return Err("native ledger count mismatch".into());
                }
                let mut output = Vec::new();
                commit.package().write_to(&mut output)?;
                Ok(output)
            },
        )?;
        results.push(serde_json::json!({"corpus_bytes":corpus.len(),"corpus_sha256":hash,"targets":count,"eager":eager,"source":source}));
    }
    println!(
        "{}",
        serde_json::to_string_pretty(
            &serde_json::json!({"schema":1,"rounds":rounds,"order":"ABBA","scope":"prepared memory owners, compiled query and native target list; retained commits; excludes open/save/drop","leaf_routes":["docx_eager_leaf_text_fanout_v1","docx_source_leaf_text_fanout_v1"],"unified_routes":["docx_eager_unified_text_fanout_v1","docx_source_unified_text_fanout_v1"],"results":results})
        )?
    );
    Ok(())
}
