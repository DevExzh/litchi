//! Export the exact XML inputs used by the source-backed XLSX cell-value
//! publication benchmark.
//!
//! The benchmark normally creates its corpus in memory and retains only
//! digests in its report.  This opt-in exporter runs the same deterministic
//! builder and the same source-backed commit/publication path, then writes the
//! source and published archives together with the selected XML overlay members.
//! It is intended for offline XML attribute-cardinality analysis; it does not
//! add content to the timed benchmark.

use super::{
    Corpus, XlsxCellCrudShape, XlsxCoordinate, build_xlsx_cell_crud_corpus, sha256_hex,
    verify_xlsx_cell_crud_output, xlsx_value,
};
use litchi_core::OwnedSource;
use quick_xml::{Reader, events::Event};
use serde::Serialize;
use soapberry_zip::office::ArchiveReader;
use std::{collections::BTreeSet, error::Error, fs, path::Path, sync::Arc};
use xml_minifier::audit::{ReplacementError, ReplacementProof};

const SOURCE_XML_MEMBERS: [&str; 5] = [
    "xl/workbook.xml",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/sheet2.xml",
    "xl/worksheets/sheet3.xml",
    "xl/worksheets/sheet4.xml",
];

#[derive(Clone, Debug, Serialize)]
struct ByteDescriptor {
    path: String,
    bytes: usize,
    sha256: String,
}

#[derive(Clone, Debug, Serialize)]
struct XmlCensus {
    start_or_empty_tags: usize,
    zero_attribute_tags: usize,
    one_attribute_tags: usize,
    multiple_attribute_tags: usize,
    attributes: usize,
    xml_space_attributes: usize,
    xml_space_elements: usize,
}

#[derive(Clone, Debug, Serialize)]
#[serde(tag = "kind")]
enum ProofDescriptor {
    Identical,
    Complete,
    Window {
        original: [usize; 2],
        replacement: [usize; 2],
    },
}

#[derive(Clone, Debug, Serialize)]
struct XmlMemberDescriptor {
    member: String,
    original: ByteDescriptor,
    replacement: ByteDescriptor,
    original_census: XmlCensus,
    replacement_census: XmlCensus,
    proof: ProofDescriptor,
}

#[derive(Clone, Debug, Serialize)]
struct CaseManifest {
    id: String,
    shape: String,
    case: String,
    update_count: usize,
    source: ByteDescriptor,
    output: ByteDescriptor,
    parts: Vec<XmlMemberDescriptor>,
}

#[derive(Clone, Debug, Serialize)]
struct FixtureManifest {
    schema: &'static str,
    generator: &'static str,
    source_publication: &'static str,
    census_definition: &'static str,
    audited_member_definition: &'static str,
    cases: Vec<CaseManifest>,
}

/// Export source-backed XLSX XML fixtures for the medium and dense-sparse
/// one-percent workloads.  When `include_one_edit` is true, the same source
/// corpus is also published through the one-cell control workload.
pub(crate) fn export(directory: &Path, include_one_edit: bool) -> Result<(), Box<dyn Error>> {
    if directory.exists() {
        return Err(format!(
            "XLSX attribute fixture directory already exists: {}",
            directory.display()
        )
        .into());
    }
    fs::create_dir_all(directory)?;

    let mut cases = Vec::new();
    for shape in XlsxCellCrudShape::ALL {
        let corpus = build_xlsx_cell_crud_corpus(shape)?;
        let shape_dir = directory.join(shape.name());
        fs::create_dir_all(&shape_dir)?;
        let source_relative = format!("{}/source.zip", shape.name());
        write_bytes(directory.join(&source_relative), &corpus.archive)?;
        let source_archive = descriptor(&source_relative, &corpus.archive);

        cases.push(export_case(
            &corpus,
            &shape_dir,
            &source_archive,
            "one-percent",
            &one_percent_updates(&corpus)?,
        )?);
        if include_one_edit && shape == XlsxCellCrudShape::Medium {
            let inventory = super::xlsx_spec(&corpus)?
                .cell_inventory
                .as_ref()
                .ok_or("XLSX attribute fixture corpus has no cell inventory")?;
            let first = inventory
                .first()
                .and_then(|cells| cells.first())
                .copied()
                .ok_or("XLSX attribute fixture corpus has no first cell")?;
            cases.push(export_case(
                &corpus,
                &shape_dir,
                &source_archive,
                "one-edit",
                &[first],
            )?);
        }
    }

    let manifest = FixtureManifest {
        schema: "litchi.xlsx-attribute-fixtures.v1",
        generator: "litchi-xlsx-source-backed-cell-values-attribute-fixtures-v1",
        source_publication: "SourceBackedEditor::edit_sheets -> set -> commit -> publish_multi_commit_to_stream",
        census_definition: "Every quick-xml Start or Empty event; zero/one/multiple classify that event's attributes, including namespace declarations; xml:space counts matching attributes and containing elements.",
        audited_member_definition: "The source-backed topology plan's owner workbook plus each selected worksheet; one-percent cases have five parts and the one-edit control has the workbook plus its selected worksheet.",
        cases,
    };
    let manifest_bytes = serde_json::to_vec_pretty(&manifest)?;
    write_bytes(directory.join("manifest.json"), &manifest_bytes)?;
    Ok(())
}

fn one_percent_updates(corpus: &Corpus) -> Result<Vec<XlsxCoordinate>, Box<dyn Error>> {
    Ok(super::xlsx_spec(corpus)?.one_percent_updates.clone())
}

fn export_case(
    corpus: &Corpus,
    shape_dir: &Path,
    source_archive: &ByteDescriptor,
    case_name: &str,
    updates: &[XlsxCoordinate],
) -> Result<CaseManifest, Box<dyn Error>> {
    let output = publish_source_backed(corpus, updates)?;
    verify_xlsx_cell_crud_output(corpus, &output, updates)?;

    let case_dir = shape_dir.join(case_name);
    let original_dir = case_dir.join("original");
    let replacement_dir = case_dir.join("replacement");
    fs::create_dir_all(&original_dir)?;
    fs::create_dir_all(&replacement_dir)?;

    let output_relative = format!("{}/{}/output.zip", corpus.manifest.shape, case_name);
    write_bytes(
        shape_dir
            .parent()
            .ok_or("XLSX attribute fixture shape has no parent")?
            .join(&output_relative),
        &output,
    )?;
    let output_archive = descriptor(&output_relative, &output);
    let source_reader = ArchiveReader::new(&corpus.archive)?;
    let output_reader = ArchiveReader::new(&output)?;
    let members = audited_member_names(updates);
    let mut audited_xml_members = Vec::with_capacity(members.len());
    for member in members {
        let original = source_reader.read(&member)?;
        let replacement = output_reader.read(&member)?;
        let original_relative = format!(
            "{}/{}/original/{}",
            corpus.manifest.shape, case_name, member
        );
        let replacement_relative = format!(
            "{}/{}/replacement/{}",
            corpus.manifest.shape, case_name, member
        );
        write_bytes(
            shape_dir
                .parent()
                .ok_or("XLSX attribute fixture shape has no parent")?
                .join(&original_relative),
            &original,
        )?;
        write_bytes(
            shape_dir
                .parent()
                .ok_or("XLSX attribute fixture shape has no parent")?
                .join(&replacement_relative),
            &replacement,
        )?;
        let proof = replacement_proof(&member, &original, &replacement)?;
        audited_xml_members.push(XmlMemberDescriptor {
            member,
            original: descriptor(&original_relative, &original),
            replacement: descriptor(&replacement_relative, &replacement),
            original_census: census(&original)?,
            replacement_census: census(&replacement)?,
            proof,
        });
    }
    Ok(CaseManifest {
        id: format!("{}-{}", corpus.manifest.shape, case_name),
        shape: corpus.manifest.shape.to_owned(),
        case: case_name.to_owned(),
        update_count: updates.len(),
        source: source_archive.clone(),
        output: output_archive,
        parts: audited_xml_members,
    })
}

fn audited_member_names(updates: &[XlsxCoordinate]) -> Vec<String> {
    let mut members = vec![SOURCE_XML_MEMBERS[0].to_owned()];
    let mut sheets = BTreeSet::new();
    for coordinate in updates {
        sheets.insert(coordinate.sheet);
    }
    members.extend(
        sheets
            .into_iter()
            .map(|sheet| format!("xl/worksheets/sheet{}.xml", sheet + 1)),
    );
    members
}

fn publish_source_backed(
    corpus: &Corpus,
    updates: &[XlsxCoordinate],
) -> Result<Vec<u8>, Box<dyn Error>> {
    let source = Arc::new(OwnedSource::new(corpus.archive.clone()));
    let editor = litchi_xlsx::cell_values::SourceBackedEditor::from_read_at(source)?;
    let selectors = super::xlsx_update_sheet_selectors(updates);
    let mut edit = editor.edit_sheets(selectors)?;
    for coordinate in updates {
        edit.set(
            coordinate.sheet,
            litchi_xlsx::Address::at(
                u32::try_from(coordinate.row)?,
                u32::try_from(coordinate.column)?,
            )?,
            xlsx_value(*coordinate) + 1,
        )?;
    }
    let commit = edit.commit()?;
    if !commit.changed() || commit.diagnostics().changed_cells() != updates.len() {
        return Err("XLSX attribute fixture source commit changed an unexpected cell count".into());
    }
    let mut output = Vec::new();
    editor.publish_multi_commit_to_stream(&mut output, &commit)?;
    Ok(output)
}

fn replacement_proof(
    member: &str,
    original: &[u8],
    replacement: &[u8],
) -> Result<ProofDescriptor, Box<dyn Error>> {
    let proof = xml_minifier::audit::verify_source_replacement(
        original,
        replacement,
        xml_minifier::audit::Limits::default(),
    )
    .map_err(|error| match error {
        ReplacementError::Original(error) => {
            format!("{member}: original XML audit failed: {error}")
        },
        ReplacementError::Replacement(error) => {
            format!("{member}: replacement XML audit failed: {error}")
        },
        _ => format!("{member}: unsupported replacement XML audit failure"),
    })?;
    Ok(match proof {
        ReplacementProof::Identical => ProofDescriptor::Identical,
        ReplacementProof::Complete => ProofDescriptor::Complete,
        ReplacementProof::Window {
            original,
            replacement,
        } => ProofDescriptor::Window {
            original: [original.start, original.end],
            replacement: [replacement.start, replacement.end],
        },
        _ => return Err(format!("{member}: unsupported replacement proof variant").into()),
    })
}

fn census(bytes: &[u8]) -> Result<XmlCensus, Box<dyn Error>> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut result = XmlCensus {
        start_or_empty_tags: 0,
        zero_attribute_tags: 0,
        one_attribute_tags: 0,
        multiple_attribute_tags: 0,
        attributes: 0,
        xml_space_attributes: 0,
        xml_space_elements: 0,
    };
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Start(element) | Event::Empty(element) => {
                let mut count = 0usize;
                let mut has_xml_space = false;
                for attribute in element.attributes() {
                    let attribute = attribute?;
                    count = count
                        .checked_add(1)
                        .ok_or("XLSX XML attribute count overflowed")?;
                    if attribute.key.as_ref() == b"xml:space" {
                        result.xml_space_attributes = result
                            .xml_space_attributes
                            .checked_add(1)
                            .ok_or("XLSX XML xml:space count overflowed")?;
                        has_xml_space = true;
                    }
                }
                result.start_or_empty_tags = result
                    .start_or_empty_tags
                    .checked_add(1)
                    .ok_or("XLSX XML tag count overflowed")?;
                result.attributes = result
                    .attributes
                    .checked_add(count)
                    .ok_or("XLSX XML attribute total overflowed")?;
                match count {
                    0 => {
                        result.zero_attribute_tags = result
                            .zero_attribute_tags
                            .checked_add(1)
                            .ok_or("XLSX XML zero-attribute tag count overflowed")?;
                    },
                    1 => {
                        result.one_attribute_tags = result
                            .one_attribute_tags
                            .checked_add(1)
                            .ok_or("XLSX XML one-attribute tag count overflowed")?;
                    },
                    _ => {
                        result.multiple_attribute_tags = result
                            .multiple_attribute_tags
                            .checked_add(1)
                            .ok_or("XLSX XML multiple-attribute tag count overflowed")?;
                    },
                }
                if has_xml_space {
                    result.xml_space_elements = result
                        .xml_space_elements
                        .checked_add(1)
                        .ok_or("XLSX XML xml:space element count overflowed")?;
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(result)
}

fn descriptor(relative_path: &str, bytes: &[u8]) -> ByteDescriptor {
    ByteDescriptor {
        path: relative_path.to_owned(),
        bytes: bytes.len(),
        sha256: sha256_hex(bytes),
    }
}

fn write_bytes(path: impl AsRef<Path>, bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let path = path.as_ref();
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    fs::write(path, bytes)?;
    Ok(())
}
