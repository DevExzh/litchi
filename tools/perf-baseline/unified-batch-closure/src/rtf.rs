//! RTF batches.

use crate::{Case, Result, check_pair, set, timed, unified_leg};
use litchi_core::Position;

pub const BUILDERS: &[(&str, fn(&[usize]) -> Result<Vec<Case<'static>>>)] = &[("rtf", rtf)];

fn leak<T>(value: T) -> &'static T {
    Box::leak(Box::new(value))
}

/// RTF: N paragraph text sets on a document of N paragraphs; one leaf
/// `replace_body_paragraph_texts` call with the limits the unified layer
/// passes the leaf (`format/rtf/write.rs`).
fn rtf(counts: &[usize]) -> Result<Vec<Case<'static>>> {
    use litchi_rtf::edit::{EditOptions, ParagraphTextReplacement};
    let max = counts.iter().copied().max().unwrap_or(0);
    let body = (0..max)
        .map(|i| format!("paragraph{i:05}"))
        .collect::<Vec<_>>()
        .join(r"\par ");
    let bytes = format!(r"{{\rtf1\ansi {body}}}").into_bytes();
    let native = leak(litchi_rtf::Package::from_bytes(&bytes)?);
    let file = leak(litchi_ops::open_bytes(bytes)?);
    fn leaf(
        native: &litchi_rtf::Package,
        replacements: &[ParagraphTextReplacement],
    ) -> Result<litchi_rtf::Package> {
        let limits = crate::limits();
        let native_limits = litchi_rtf::SemanticLimits::builder()
            .max_operations(limits.max_operations())?
            .max_replacement_bytes(limits.max_staged_bytes())?
            .build()?;
        let mut edit = native.edit_with_options(EditOptions::new().semantic_limits(native_limits));
        edit.replace_body_paragraph_texts(replacements)?;
        Ok(edit.commit()?.into_package())
    }
    let mut cases = Vec::new();
    for &n in counts {
        let replacements = leak(
            (0..n)
                .map(|i| ParagraphTextReplacement::new(Position::new(i), format!("changed{i:05}")))
                .collect::<Vec<_>>(),
        );
        let items = leak(
            (0..n)
                .map(|i| {
                    set(
                        format!("/body/p[{}]", i + 1),
                        "text",
                        format!("changed{i:05}"),
                    )
                })
                .collect::<Vec<_>>(),
        );
        cases.push(Case {
            format: "rtf",
            scenario: "paragraph_text_set",
            owner: "default",
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
