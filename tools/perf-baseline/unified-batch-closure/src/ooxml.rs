//! DOCX, XLSX, PPTX and XLSB batches.

#[allow(unused_imports)]
use crate::{Case, Checked, Result, check_pair, remove, set, timed, timed_with, unified_leg};
#[allow(unused_imports)]
use litchi_core::Position;
#[allow(unused_imports)]
use litchi_ops::{BatchItem, File, OpenOptions, SourceMode};

pub const BUILDERS: &[(&str, fn(&[usize]) -> Result<Vec<Case<'static>>>)] = &[
    #[cfg(feature = "docx")]
    ("docx", docx),
    #[cfg(feature = "xlsx")]
    ("xlsx", xlsx),
    #[cfg(feature = "pptx")]
    ("pptx", pptx),
    #[cfg(feature = "xlsb")]
    ("xlsb", xlsb),
];

#[allow(dead_code)]
fn eager(bytes: Vec<u8>) -> Result<&'static File> {
    let file =
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))?;
    Ok(Box::leak(Box::new(file)))
}

#[allow(dead_code)]
fn leak<T>(value: T) -> &'static T {
    Box::leak(Box::new(value))
}

/// N one-run paragraphs, written as a bare OPC package (record 0078's corpus
/// shape, scaled).
#[cfg(feature = "docx")]
fn docx_corpus(n: usize) -> Result<Vec<u8>> {
    use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
    let xml = format!(
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{}<w:sectPr/></w:body></w:document>",
        "<w:p><w:r><w:t>seed</w:t></w:r></w:p>".repeat(n)
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

/// DOCX: N run-property sets (`bold`) on N distinct paragraphs; one
/// `Edit::format_run` per paragraph against `set /body/p[i]/r[1]`.
#[cfg(feature = "docx")]
fn docx(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    let max = counts.iter().copied().max().unwrap_or(0);
    let bytes = docx_corpus(max)?;
    let native = leak(litchi_docx::Package::from_bytes(&bytes)?);
    let file = eager(bytes)?;
    fn leaf(native: &litchi_docx::Package, n: usize) -> Result<Vec<litchi_docx::document::Commit>> {
        let mut commits = Vec::new();
        let mut edit = native.edit()?;
        for i in 0..n {
            if edit.remaining_operations() == 0 {
                commits.push(edit.commit()?);
                edit = commits.last().ok_or("commit")?.package().edit()?;
            }
            edit.format_run(Position::new(i), Position::new(0), |run| {
                run.bold = Some(true)
            })?;
        }
        commits.push(edit.commit()?);
        Ok(commits)
    }
    let mut cases = Vec::new();
    for &n in counts {
        let items = leak(
            (0..n)
                .map(|i| set(format!("/body/p[{}]/r[1]", i + 1), "bold", true))
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "docx",
            scenario: "run_bold_set",
            owner: "eager",
            count: n,
            check: check_pair(
                move || {
                    let commits = leaf(native, n)?;
                    let bytes = commits
                        .last()
                        .ok_or("commit")?
                        .package()
                        .clone()
                        .to_bytes()?;
                    Ok((bytes, commits.len()))
                },
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, n))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}

/// XLSX: N scalar cell sets (`A1..AN`) on a blank workbook.
#[cfg(feature = "xlsx")]
fn xlsx(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    let bytes = litchi_xlsx::Package::new()?.to_bytes()?;
    let native = leak(litchi_xlsx::Package::from_bytes(&bytes)?);
    let file = eager(bytes)?;
    fn leaf(
        native: &litchi_xlsx::Package,
        addresses: &[String],
    ) -> Result<litchi_xlsx::workbook::Commit> {
        let mut edit = native.edit()?;
        {
            let mut sheet = edit.sheet("Sheet1")?.ok_or("no Sheet1")?;
            for (i, address) in addresses.iter().enumerate() {
                sheet.set(address.as_str(), (i + 1) as i32)?;
            }
        }
        Ok(edit.commit()?)
    }
    let mut cases = Vec::new();
    for &n in counts {
        let addresses = leak((1..=n).map(|r| format!("A{r}")).collect::<Vec<_>>());
        let items = leak(
            (1..=n)
                .map(|r| set(format!("/Sheet1/A{r}"), "value", r as i64))
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "xlsx",
            scenario: "cell_value_set",
            owner: "eager",
            count: n,
            check: check_pair(
                move || Ok((leaf(native, addresses)?.package().to_bytes()?, 1)),
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, addresses))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}

/// PPTX: N shape text sets on one slide of N text boxes; the leaf stages
/// 256-replacement chunks in one `set_shape_text_chunks` call of one edit.
#[cfg(feature = "pptx")]
fn pptx(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_pptx::ShapeTextReplacement;
    const TEXT: &str = "batch closure text";
    let max = counts.iter().copied().max().unwrap_or(0);
    let mut builder = litchi_pptx::builder::Builder::new()?;
    let slide = builder.presentation()?.add_slide()?;
    for i in 0..max {
        slide.add_text_box(&format!("original-{i}"), 0, 0, 914_400, 914_400);
    }
    let bytes = builder.build()?.to_bytes()?;
    let native = leak(litchi_pptx::Package::from_bytes(&bytes)?);
    let file = eager(bytes)?;
    fn leaf(
        native: &litchi_pptx::Package,
        replacements: &[ShapeTextReplacement<'static>],
    ) -> Result<litchi_pptx::Package> {
        // The chunked call the unified run uses: one slide walk for every
        // chunk of at most `MAX_SHAPE_TEXT_REPLACEMENTS` (256).
        let chunks: Vec<&[ShapeTextReplacement<'static>]> = replacements.chunks(256).collect();
        let mut edit = native.edit()?;
        edit.set_shape_text_chunks(Position::new(0), &chunks)?;
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let replacements = leak(
            (0..n)
                .map(|i| ShapeTextReplacement::at(i, TEXT))
                .collect::<Vec<_>>(),
        );
        let items = leak(
            (0..n)
                .map(|i| set(format!("/slide[1]/shape[{}]", i + 1), "text", TEXT))
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "pptx",
            scenario: "shape_text_set",
            owner: "eager",
            count: n,
            check: check_pair(
                move || Ok((leaf(native, replacements)?.to_bytes()?, 1)),
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, replacements))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}

#[cfg(feature = "xlsb")]
fn column_name(mut col: u32) -> String {
    let mut out = Vec::new();
    loop {
        out.push(b'A' + (col % 26) as u8);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    out.reverse();
    String::from_utf8(out).expect("ascii")
}

/// XLSB: N scalar cell sets, and N cell removals, on a sheet of N cells
/// (32 columns wide). Both coalesce into one leaf edit.
#[cfg(feature = "xlsb")]
fn xlsb(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_xlsb::builder::{self, Builder};
    use litchi_xlsb::cell_values::{Reference, Scalar};
    let max = counts.iter().copied().max().unwrap_or(0);
    let mut writer = Builder::new();
    let mut sheet = builder::Worksheet::new("Sheet1");
    for i in 0..max {
        sheet.set_cell((i / 32) as u32, (i % 32) as u32, i as f64 + 0.5);
    }
    writer.add_worksheet(sheet);
    let mut out = std::io::Cursor::new(Vec::new());
    writer.build()?.write_to(&mut out)?;
    let bytes = out.into_inner();
    let native = leak(litchi_xlsb::Package::from_bytes(&bytes)?);
    let file = eager(bytes)?;
    let reference = |i: usize| ((i / 32) as u32, (i % 32) as u32);
    fn set_leaf(
        native: &litchi_xlsb::Package,
        refs: &[(u32, u32)],
    ) -> Result<litchi_xlsb::Package> {
        let mut edit = native
            .edit_cell_values(0usize)?
            .resolve()?
            .ok_or("no sheet")?;
        for (i, &(row, col)) in refs.iter().enumerate() {
            edit.set_scalar(Reference::new(row, col)?, Scalar::Number(-(i as f64) - 1.0))?;
        }
        Ok(edit.commit()?.into_package())
    }
    fn remove_leaf(
        native: &litchi_xlsb::Package,
        refs: &[(u32, u32)],
    ) -> Result<litchi_xlsb::Package> {
        let mut edit = native
            .edit_cell_values(0usize)?
            .resolve()?
            .ok_or("no sheet")?;
        for &(row, col) in refs {
            if !edit.remove(Reference::new(row, col)?)? {
                return Err("leaf removal staged nothing".into());
            }
        }
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let refs = leak((0..n).map(reference).collect::<Vec<_>>());
        let path = |&(row, col): &(u32, u32)| format!("/sheet[1]/{}{}", column_name(col), row + 1);
        let sets = leak(
            refs.iter()
                .enumerate()
                .map(|(i, r)| set(path(r), "value", -(i as f64) - 1.0))
                .collect::<Vec<_>>(),
        );
        let removes = leak(refs.iter().map(|r| remove(path(r))).collect::<Vec<_>>());
        cases.push(Case {
            format: "xlsb",
            scenario: "cell_value_set",
            owner: "eager",
            count: n,
            check: check_pair(
                move || Ok((set_leaf(native, refs)?.to_bytes()?, 1)),
                file,
                sets,
            ),
            leaf: Box::new(move || timed(|| set_leaf(native, refs))),
            unified: unified_leg(file, sets),
        });
        cases.push(Case {
            format: "xlsb",
            scenario: "cell_remove",
            owner: "eager",
            count: n,
            check: check_pair(
                move || Ok((remove_leaf(native, refs)?.to_bytes()?, 1)),
                file,
                removes,
            ),
            leaf: Box::new(move || timed(|| remove_leaf(native, refs))),
            unified: unified_leg(file, removes),
        });
    }
    Ok(cases)
}
