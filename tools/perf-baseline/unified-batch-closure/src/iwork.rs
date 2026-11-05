//! Pages, Keynote and Numbers batches (leaf crates used, never edited).

#[allow(unused_imports)]
use crate::{Case, Result, check_pair, set, timed, unified_leg};
#[allow(unused_imports)]
use litchi_core::Position;

pub const BUILDERS: &[(&str, fn(&[usize]) -> Result<Vec<Case<'static>>>)] = &[
    #[cfg(feature = "pages")]
    ("pages", pages),
    #[cfg(feature = "keynote")]
    ("keynote", keynote),
    #[cfg(feature = "numbers")]
    ("numbers", numbers),
];

fn leak<T>(value: T) -> &'static T {
    Box::leak(Box::new(value))
}

/// Pages: N paragraph text sets on a body of N paragraphs (record 0314's
/// document); one leaf `edit_section_paragraphs` edit with N `set`s.
#[cfg(feature = "pages")]
fn pages(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_pages::builder::{Builder, Identity};
    let max = counts.iter().copied().max().unwrap_or(0);
    let text = (0..max)
        .map(|i| format!("Paragraph {i}"))
        .collect::<Vec<_>>()
        .join("\n");
    let bytes = Builder::new()
        .identity(Identity::from_bytes([0x29; 16]))
        .body_text(text)
        .build()?
        .to_bytes()?;
    let native = leak(litchi_pages::Package::from_bytes(&bytes)?);
    let file = leak(litchi_ops::open_bytes(bytes)?);
    fn leaf(native: &litchi_pages::Package, texts: &[String]) -> Result<litchi_pages::Package> {
        let mut edit = native.edit_section_paragraphs(0_usize)?;
        for (i, text) in texts.iter().enumerate() {
            edit.set(Position::new(i), text)?;
        }
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let texts = leak((0..n).map(|i| format!("Edited {i}")).collect::<Vec<_>>());
        let items = leak(
            texts
                .iter()
                .enumerate()
                .map(|(i, t)| set(format!("/body/p[{}]", i + 1), "text", t.clone()))
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "pages",
            scenario: "paragraph_text_set",
            owner: "default",
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

#[cfg(feature = "keynote")]
fn keynote_bytes(package: &litchi_keynote::Package) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

/// The span of `before` that differs from `after`, and its replacement
/// (record 0314's `span`).
#[cfg(feature = "keynote")]
fn span<'a>(before: &str, after: &'a str) -> Result<(litchi_keynote::TextSpan, &'a str)> {
    use litchi_keynote::{TextPosition, TextSpan};
    let prefix = before
        .char_indices()
        .zip(after.chars())
        .find(|((_, l), r)| l != r)
        .map_or_else(|| before.len().min(after.len()), |((i, _), _)| i);
    let (before_rest, after_rest) = (&before[prefix..], &after[prefix..]);
    let suffix = before_rest
        .chars()
        .rev()
        .zip(after_rest.chars().rev())
        .take_while(|(l, r)| l == r)
        .map(|(l, _)| l.len_utf8())
        .sum::<usize>();
    let removed = &before_rest[..before_rest.len() - suffix];
    let start = before[..prefix].encode_utf16().count();
    let end = start + removed.encode_utf16().count();
    let position = |units| TextPosition::from_utf16_index(units);
    Ok((
        TextSpan::new(position(start)?, position(end)?)?,
        &after_rest[..after_rest.len() - suffix],
    ))
}

/// Keynote: N paragraph text sets on a body placeholder of N paragraphs
/// (slide 2 of the tracked `multi-slide.key`, its body set to N lines outside
/// the clock). The leaf's slide-text edit stages one operation, so its leg is
/// the one replacement that yields the batch's final text.
#[cfg(feature = "keynote")]
fn keynote(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_keynote::slide::placeholder::Kind as Role;
    use litchi_keynote::{Package, SlideSelector};
    let max = counts.iter().copied().max().unwrap_or(0);
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../../test-data/iwork/keynote/multi-slide.key");
    let source = Package::from_bytes(&std::fs::read(path)?)?;
    let lines = (0..max)
        .map(|i| format!("line {i}"))
        .collect::<Vec<_>>()
        .join("\n");
    let bytes = {
        let mut edit = source.edit_slide_text(SlideSelector::index(1), Role::Body)?;
        edit.set(&lines)?;
        keynote_bytes(edit.commit()?.package())?
    };
    let native = leak(Package::from_bytes(&bytes)?);
    let file = leak(litchi_ops::open_bytes(bytes)?);
    fn leaf(native: &Package, final_body: &str) -> Result<Package> {
        let mut edit = native.edit_slide_text(SlideSelector::index(1), Role::Body)?;
        let (span, replacement) = span(edit.text()?, final_body)?;
        edit.replace(span, replacement)?;
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let texts: Vec<String> = (0..n).map(|i| format!("edited line {i}")).collect();
        let final_body = leak(
            (0..max)
                .map(|i| {
                    if i < n {
                        texts[i].clone()
                    } else {
                        format!("line {i}")
                    }
                })
                .collect::<Vec<_>>()
                .join("\n"),
        );
        let items = leak(
            texts
                .iter()
                .enumerate()
                .map(|(i, t)| {
                    set(
                        format!("/slide[2]/placeholder[body]/p[{}]", i + 1),
                        "text",
                        t.clone(),
                    )
                })
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "keynote",
            scenario: "paragraph_text_set",
            owner: "default",
            count: n,
            check: check_pair(
                move || Ok((keynote_bytes(&leaf(native, final_body)?)?, 1)),
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, final_body))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}

/// Numbers: N numeric cell sets, row-major over a table of 10 body columns
/// added to a new document; one leaf `edit_table_cells` edit.
#[cfg(feature = "numbers")]
fn numbers(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_numbers::table::At;
    use litchi_numbers::table::cells::{Change, Input};
    use litchi_numbers::{Package, SheetSelector, TableSelector};
    let max = counts.iter().copied().max().unwrap_or(0);
    let rows = u32::try_from(max.div_ceil(10) + 1)?;
    let package = Package::new()
        .edit_table_lifecycle(0_usize)?
        .add("Grid", rows, 11)?
        .commit()?
        .into_package();
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    let native = leak(Package::from_bytes(&bytes)?);
    let file = leak(litchi_ops::open_bytes(bytes)?);
    fn leaf(native: &Package, cells: &[(At, f64)]) -> Result<Package> {
        Ok(native
            .edit_table_cells(SheetSelector::index(0), TableSelector::index(1))?
            .extend(
                cells
                    .iter()
                    .map(|(p, v)| Change::set(*p, Input::number(*v).expect("finite"))),
            )?
            .commit()?
            .into_package())
    }
    fn bytes_of(package: &Package) -> Result<Vec<u8>> {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes)?;
        Ok(bytes)
    }
    let mut cases = Vec::new();
    for &n in counts {
        let cells: Vec<(At, String, f64)> = (0..n)
            .map(|i| {
                let row = (i / 10 + 1) as u32;
                let column = (i % 10 + 1) as u32;
                (
                    At::new(row, column),
                    format!("{}{}", char::from(b'A' + column as u8), row + 1),
                    i as f64,
                )
            })
            .collect();
        let items = leak(
            cells
                .iter()
                .map(|(_, a1, v)| set(format!("/sheet[1]/table[2]/{a1}"), "value", *v))
                .collect::<Vec<_>>(),
        );
        let cells = leak(
            cells
                .into_iter()
                .map(|(p, _, v)| (p, v))
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "numbers",
            scenario: "cell_value_set",
            owner: "default",
            count: n,
            check: check_pair(
                move || Ok((bytes_of(&leaf(native, cells)?)?, 1)),
                file,
                items,
            ),
            leaf: Box::new(move || timed(|| leaf(native, cells))),
            unified: unified_leg(file, items),
        });
    }
    Ok(cases)
}
