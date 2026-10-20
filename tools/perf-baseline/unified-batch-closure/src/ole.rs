//! DOC, XLS and PPT batches.

#[allow(unused_imports)]
use crate::{Case, Result, check_pair, remove, set, timed, unified_leg};
#[allow(unused_imports)]
use litchi_core::Position;
#[allow(unused_imports)]
use litchi_ops::{BatchItem, File, OpenOptions, SourceMode};

pub const BUILDERS: &[(&str, fn(&[usize]) -> Result<Vec<Case<'static>>>)] = &[
    #[cfg(feature = "doc")]
    ("doc", doc),
    #[cfg(feature = "xls")]
    ("xls", xls),
    #[cfg(feature = "ppt")]
    ("ppt", ppt),
];

fn eager(bytes: Vec<u8>) -> Result<&'static File> {
    let file =
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))?;
    Ok(Box::leak(Box::new(file)))
}

fn leak<T>(value: T) -> &'static T {
    Box::leak(Box::new(value))
}

/// DOC: N paragraph text sets on a document of N paragraphs; one leaf
/// `Edit::replace_paragraph` per paragraph, with the transaction limits the
/// unified layer passes the leaf (`format/doc.rs` `body_transaction_limits`).
#[cfg(feature = "doc")]
fn doc(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    let max = counts.iter().copied().max().unwrap_or(0);
    let mut builder = litchi_doc::Builder::new();
    for i in 0..max {
        builder.add_paragraph(&format!("seed paragraph {i:05}"))?;
    }
    let bytes = builder.build()?.to_bytes()?;
    let native = leak(litchi_doc::Package::from_bytes(&bytes)?);
    let file = eager(bytes)?;
    fn leaf(native: &litchi_doc::Package, texts: &[String]) -> Result<litchi_doc::Package> {
        let limits = crate::limits();
        let options = litchi_doc::EditOptions::default().transaction_limits(
            litchi_doc::body_text::TransactionLimits::new(
                limits.max_operations(),
                usize::MAX,
                usize::MAX,
            )
            .with_max_retained_render_bytes(limits.max_snapshot_bytes()),
        );
        let mut edit = native.edit_with_options(options)?;
        for (i, text) in texts.iter().enumerate() {
            edit.replace_paragraph(Position::new(i), text)?;
        }
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let texts = leak(
            (0..n)
                .map(|i| format!("changed paragraph {i:05}"))
                .collect::<Vec<_>>(),
        );
        let items = leak(
            texts
                .iter()
                .enumerate()
                .map(|(i, t)| set(format!("/body/p[{}]", i + 1), "text", t.clone()))
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "doc",
            scenario: "paragraph_text_set",
            owner: "eager",
            count: n,
            check: check_pair(
                move || Ok((leaf(native, texts)?.to_bytes()?, 1)),
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, texts))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}

#[cfg(feature = "xls")]
#[derive(Clone, Copy)]
enum XlsOp {
    Set,
    Remove,
}

/// One leaf edit per 4,096 staged changes (`MAX_STAGED_CHANGES`), as the
/// unified run flushes at the same bound; returns the package and commits.
#[cfg(feature = "xls")]
fn xls_leaf(
    native: &litchi_xls::Package,
    refs: &[litchi_xls::cell_values::Reference],
    op: XlsOp,
) -> Result<(litchi_xls::Package, usize)> {
    use litchi_xls::cell_values::{Selector, Value};
    let mut package: Option<litchi_xls::Package> = None;
    let mut commits = 0;
    let mut index = 0;
    loop {
        let mut edit = match &package {
            None => native.edit()?,
            Some(p) => p.edit()?,
        };
        while index < refs.len() {
            let selector = Selector::Position(Position::new(0));
            let staged = match op {
                XlsOp::Set => edit.set_value(selector, refs[index], Value::Number(-1.0)),
                XlsOp::Remove => edit.remove_cell(selector, refs[index]),
            };
            match staged {
                Ok(()) => index += 1,
                Err(e) if e.limit_kind() == Some(litchi_xls::LimitKind::Operations) => break,
                Err(e) => return Err(e.into()),
            }
        }
        package = Some(edit.commit()?.into_package());
        commits += 1;
        if index == refs.len() {
            return Ok((package.ok_or("package")?, commits));
        }
    }
}

/// XLS: N cell value sets and N cell removals on a column of N numbers, on
/// an eager owner and on a file-backed (source) owner — record 0209's
/// file-backed removal at 1,000 cells (2.093×) is the `file` `cell_remove`
/// case here.
#[cfg(feature = "xls")]
fn xls(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_xls::cell_values::Reference;
    let max = counts.iter().copied().max().unwrap_or(0);
    let mut builder = litchi_xls::writer::Builder::new();
    let sheet = builder.add_worksheet("Sheet1")?;
    for row in 0..max as u32 {
        builder.write_number(sheet, row, 0, f64::from(row + 1))?;
    }
    let bytes = builder.build()?.to_bytes()?;
    let eager_native = leak(litchi_xls::Package::from_bytes(&bytes)?);
    let eager_file = eager(bytes.clone())?;
    // The file-backed owners read one temporary file, kept for the run.
    let dir = leak(tempfile::tempdir()?);
    let path = dir.path().join("fixture.xls");
    std::fs::write(&path, &bytes)?;
    let source = litchi_core::FileSource::open(&path)?;
    let workbook = litchi_xls::SourceBackedWorkbook::from_read_at(std::sync::Arc::new(source))?;
    let file_native = leak(litchi_xls::Package::from_source(workbook)?);
    let file_unified = leak(litchi_ops::open_with(
        &path,
        &OpenOptions::new().source_mode(SourceMode::Auto),
    )?);
    let mut cases = Vec::new();
    for &n in counts {
        let refs = leak(
            (0..n as u32)
                .map(|r| Reference::new(r, 0))
                .collect::<std::result::Result<Vec<_>, _>>()?,
        );
        let paths: Vec<String> = (1..=n).map(|r| format!("/sheet[1]/A{r}")).collect();
        let sets = leak(
            paths
                .iter()
                .map(|p| set(p.clone(), "value", -1.0))
                .collect::<Vec<_>>(),
        );
        let removes = leak(paths.iter().map(|p| remove(p.clone())).collect::<Vec<_>>());
        for (owner, native, file) in [
            ("eager", eager_native, eager_file),
            ("file", file_native, file_unified),
        ] {
            for (scenario, op, items) in [
                ("cell_value_set", XlsOp::Set, sets),
                ("cell_remove", XlsOp::Remove, removes),
            ] {
                cases.push(Case {
                    format: "xls",
                    scenario,
                    owner,
                    count: n,
                    check: check_pair(
                        move || {
                            let (package, commits) = xls_leaf(native, refs, op)?;
                            Ok((package.to_bytes()?, commits))
                        },
                        file,
                        items,
                    ),
                    leaf: Box::new(move || timed(|| xls_leaf(native, refs, op))),
                    unified: unified_leg(file, items),
                });
            }
        }
    }
    Ok(cases)
}

/// PPT: N shape text sets on one slide of N text boxes; one leaf edit with
/// one `set_shape_text_chunks` call of 256-replacement chunks, as the
/// unified run stages them (`format/ppt/write/coalesce.rs`).
#[cfg(feature = "ppt")]
fn ppt(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_ppt::slide_order::{MAX_SHAPE_TEXT_REPLACEMENTS, ShapeTextReplacement};
    use litchi_ppt::text_edit::Target;
    let max = counts.iter().copied().max().unwrap_or(0);
    let mut builder = litchi_ppt::writer::Builder::new();
    let slide = builder.add_slide()?;
    builder.set_slide_timing(slide, litchi_ppt::writer::SlideTiming::default())?;
    for i in 0..max {
        builder.add_textbox(slide, 10 + (i % 100) as i32, 10, 240, 40, "before")?;
    }
    let bytes = builder.build()?.to_bytes()?;
    let native = leak(litchi_ppt::Package::from_bytes(&bytes)?);
    let file = eager(bytes)?;
    fn leaf(
        native: &litchi_ppt::Package,
        replacements: &[ShapeTextReplacement],
    ) -> Result<litchi_ppt::Package> {
        let chunks: Vec<&[ShapeTextReplacement]> =
            replacements.chunks(MAX_SHAPE_TEXT_REPLACEMENTS).collect();
        let mut edit = native.edit()?;
        edit.set_shape_text_chunks(&chunks)?;
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let replacements = leak(
            (0..n)
                .map(|i| {
                    ShapeTextReplacement::new(
                        Target::new(Position::new(0), Position::new(i)),
                        format!("after-{i}"),
                    )
                })
                .collect::<Vec<_>>(),
        );
        let items = leak(
            (0..n)
                .map(|i| {
                    set(
                        format!("/slide[1]/shape[{}]", i + 1),
                        "text",
                        format!("after-{i}"),
                    )
                })
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "ppt",
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
