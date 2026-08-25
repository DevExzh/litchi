#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "transaction assertions intentionally panic on failure"
)]

use litchi_rtf::edit::{DurableComposition, DurableMergePlan, Limits, MergeResolution};
use litchi_rtf::{Alignment, Document, UnderlineStyle, edit::TextSpan};
use std::num::NonZeroU16;

fn limits(max_operations: usize) -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        1024 * 1024,
        max_operations,
        8,
        64 * 1024,
        256 * 1024,
    )
}

fn reversible_payload_limits() -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        300_000,
        8,
        8,
        256 * 1024,
        256 * 1024,
    )
}

fn replace(
    source: &Document,
    span: TextSpan,
    value: &str,
    max_operations: usize,
) -> litchi_core::patch::Patch<litchi_core::patch::Reversible> {
    let mut edit = source.edit();
    edit.replace_text(span, value).unwrap();
    edit.commit()
        .unwrap()
        .patch()
        .to_durable(limits(max_operations))
        .unwrap()
}

#[test]
fn disjoint_durable_join_replays_and_derives_combined_inverse() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    let first = replace(&source, TextSpan::new(0, 5).unwrap(), "Alpha", 8);
    let second = replace(&source, TextSpan::new(6, 12).unwrap(), "Bravo", 8);

    let mut composition = DurableComposition::new(&source, limits(8));
    composition.join(first).unwrap();
    composition.join(second).unwrap();
    let combined = composition.finish().unwrap();
    let target = source.apply_durable(&combined).unwrap();
    assert_eq!(target.text(), "Alpha Bravo");
    let restored = target.apply_durable(&combined.inverse()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
}

#[test]
fn durable_join_is_order_independent_and_rejects_unequal_character_overlap() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    let first = replace(&source, TextSpan::new(0, 5).unwrap(), "Alpha", 8);
    let second = replace(&source, TextSpan::new(6, 12).unwrap(), "Bravo", 8);

    let mut forward = DurableComposition::new(&source, limits(8));
    forward.join(first.clone()).unwrap();
    forward.join(second.clone()).unwrap();
    let mut reverse = DurableComposition::new(&source, limits(8));
    reverse.join(second).unwrap();
    reverse.join(first).unwrap();
    assert_eq!(
        forward.finish().unwrap().to_deterministic_json().unwrap(),
        reverse.finish().unwrap().to_deterministic_json().unwrap()
    );

    let mut bold = source.edit();
    bold.set_text_bold(TextSpan::new(0, 7).unwrap(), true)
        .unwrap();
    let bold = bold
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    let mut italic = source.edit();
    italic
        .set_text_italic(TextSpan::new(6, 12).unwrap(), true)
        .unwrap();
    let italic = italic
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    let mut character = DurableComposition::new(&source, limits(8));
    character.join(bold.clone()).unwrap();
    let formatting = character.finish().unwrap();
    let formatted = source.apply_durable(&formatting).unwrap();
    let inverse = formatting.inverse();
    let formatted_bytes = formatted.to_bytes().unwrap();
    let formatted_hash = litchi_core::patch::BlobId::of(&formatted_bytes).as_hex();
    assert!(inverse.operations().iter().all(|operation| {
        operation
            .preconditions
            .get("artifact_sha256")
            .and_then(serde_json::Value::as_str)
            == Some(formatted_hash.as_str())
    }));
    let restored = formatted.apply_durable(&inverse).unwrap();
    assert_eq!(restored.text(), source.text());
    assert!(restored.body().runs().all(|run| !run.format().bold()));

    let mut character = DurableComposition::new(&source, limits(8));
    character.join(bold).unwrap();
    assert!(character.join(italic).is_err());
}

#[test]
fn durable_composition_replays_and_inverts_exact_underline_styles() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    let mut edit = source.edit();
    edit.set_text_underline(TextSpan::new(6, 12).unwrap(), UnderlineStyle::DoubleWave)
        .unwrap();
    let branch = edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert_eq!(branch.operations()[0].op, "character-underline.set");
    assert_eq!(
        branch.operations()[0].value,
        serde_json::Value::String("double-wave".into())
    );

    let mut composition = DurableComposition::new(&source, limits(8));
    composition.join(branch).unwrap();
    let combined = composition.finish().unwrap();
    let formatted = source.apply_durable(&combined).unwrap();
    assert_eq!(
        formatted
            .body()
            .runs()
            .find(|run| run.text() == "Second")
            .unwrap()
            .format()
            .underline(),
        UnderlineStyle::DoubleWave
    );
    let restored = formatted.apply_durable(&combined.inverse()).unwrap();
    assert!(
        restored
            .body()
            .runs()
            .all(|run| { run.format().underline() == UnderlineStyle::None })
    );

    let mut invalid_edit = source.edit();
    invalid_edit
        .set_text_underline(TextSpan::new(7, 12).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let invalid_branch = invalid_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    let mut composition = DurableComposition::new(&source, limits(8));
    composition.join(combined).unwrap();
    assert!(composition.join(invalid_branch).is_err());
}

#[test]
fn durable_composition_replays_strike_and_composes_disjoint_or_rejects_overlap() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    let mut strike_edit = source.edit();
    strike_edit
        .set_text_strike(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let strike = strike_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert_eq!(strike.operations()[0].op, "character-strike.set");

    let mut underline_edit = source.edit();
    underline_edit
        .set_text_underline(TextSpan::new(6, 12).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let underline = underline_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();

    let mut disjoint = DurableComposition::new(&source, limits(8));
    disjoint.join(strike.clone()).unwrap();
    disjoint.join(underline.clone()).unwrap();
    let combined = disjoint.finish().unwrap();
    let formatted = source.apply_durable(&combined).unwrap();
    assert!(
        formatted
            .body()
            .runs()
            .find(|run| run.text() == "First")
            .unwrap()
            .format()
            .strike()
    );
    assert_eq!(
        formatted
            .body()
            .runs()
            .find(|run| run.text() == "Second")
            .unwrap()
            .format()
            .underline(),
        UnderlineStyle::Single
    );
    let restored = formatted.apply_durable(&combined.inverse()).unwrap();
    assert!(restored.body().runs().all(|run| !run.format().strike()));
    assert!(
        restored
            .body()
            .runs()
            .all(|run| run.format().underline() == UnderlineStyle::None)
    );

    let mut overlap = DurableComposition::new(&source, limits(8));
    overlap.join(strike).unwrap();
    let mut overlapping_underline = source.edit();
    overlapping_underline
        .set_text_underline(TextSpan::new(0, 5).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let overlapping_underline = overlapping_underline
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert!(overlap.join(overlapping_underline).is_err());
}

#[test]
fn durable_composition_replays_hidden_and_composes_disjoint_or_rejects_overlap() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    let mut hidden_edit = source.edit();
    hidden_edit
        .set_text_hidden(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let hidden = hidden_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert_eq!(hidden.operations()[0].op, "character-hidden.set");

    let mut underline_edit = source.edit();
    underline_edit
        .set_text_underline(TextSpan::new(6, 12).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let underline = underline_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();

    let mut disjoint = DurableComposition::new(&source, limits(8));
    disjoint.join(hidden.clone()).unwrap();
    disjoint.join(underline).unwrap();
    let combined = disjoint.finish().unwrap();
    let formatted = source.apply_durable(&combined).unwrap();
    assert!(
        formatted
            .body()
            .runs()
            .find(|run| run.text() == "First")
            .unwrap()
            .format()
            .hidden()
    );
    assert_eq!(
        formatted
            .body()
            .runs()
            .find(|run| run.text() == "Second")
            .unwrap()
            .format()
            .underline(),
        UnderlineStyle::Single
    );
    let restored = formatted.apply_durable(&combined.inverse()).unwrap();
    assert!(restored.body().runs().all(|run| !run.format().hidden()));
    assert!(
        restored
            .body()
            .runs()
            .all(|run| run.format().underline() == UnderlineStyle::None)
    );

    let mut overlap = DurableComposition::new(&source, limits(8));
    overlap.join(hidden).unwrap();
    let mut overlapping_underline_edit = source.edit();
    overlapping_underline_edit
        .set_text_underline(TextSpan::new(0, 5).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let overlapping_underline = overlapping_underline_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert!(overlap.join(overlapping_underline).is_err());
}

#[test]
fn durable_composition_replays_small_caps_and_composes_disjoint_or_rejects_overlap() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    let mut small_caps_edit = source.edit();
    small_caps_edit
        .set_text_small_caps(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let small_caps = small_caps_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert_eq!(small_caps.operations()[0].op, "character-small-caps.set");

    let mut underline_edit = source.edit();
    underline_edit
        .set_text_underline(TextSpan::new(6, 12).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let underline = underline_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();

    let mut disjoint = DurableComposition::new(&source, limits(8));
    disjoint.join(small_caps.clone()).unwrap();
    disjoint.join(underline).unwrap();
    let combined = disjoint.finish().unwrap();
    let formatted = source.apply_durable(&combined).unwrap();
    assert!(
        formatted
            .body()
            .runs()
            .find(|run| run.text() == "First")
            .unwrap()
            .format()
            .small_caps()
    );
    assert_eq!(
        formatted
            .body()
            .runs()
            .find(|run| run.text() == "Second")
            .unwrap()
            .format()
            .underline(),
        UnderlineStyle::Single
    );
    let restored = formatted.apply_durable(&combined.inverse()).unwrap();
    assert!(restored.body().runs().all(|run| !run.format().small_caps()));
    assert!(
        restored
            .body()
            .runs()
            .all(|run| run.format().underline() == UnderlineStyle::None)
    );

    let mut overlap = DurableComposition::new(&source, limits(8));
    overlap.join(small_caps).unwrap();
    let mut overlapping_edit = source.edit();
    overlapping_edit
        .set_text_underline(TextSpan::new(0, 5).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let overlapping = overlapping_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert!(overlap.join(overlapping).is_err());
}

#[test]
fn durable_composition_replays_all_caps_and_composes_disjoint_or_rejects_overlap() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    let mut all_caps_edit = source.edit();
    all_caps_edit
        .set_text_all_caps(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let all_caps = all_caps_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert_eq!(all_caps.operations()[0].op, "character-all-caps.set");

    let mut underline_edit = source.edit();
    underline_edit
        .set_text_underline(TextSpan::new(6, 12).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let underline = underline_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();

    let mut disjoint = DurableComposition::new(&source, limits(8));
    disjoint.join(all_caps.clone()).unwrap();
    disjoint.join(underline).unwrap();
    let combined = disjoint.finish().unwrap();
    let formatted = source.apply_durable(&combined).unwrap();
    let all_caps_text = formatted
        .body()
        .runs()
        .filter(|run| run.format().all_caps())
        .map(|run| run.text())
        .collect::<String>();
    let underlined_text = formatted
        .body()
        .runs()
        .filter(|run| run.format().underline() == UnderlineStyle::Single)
        .map(|run| run.text())
        .collect::<String>();
    assert!(all_caps_text.contains("First"));
    assert!(!all_caps_text.contains("Second"));
    assert!(underlined_text.contains("Second"));
    let reopened = Document::from_bytes(&formatted.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.text(), source.text());

    let restored = formatted.apply_durable(&combined.inverse()).unwrap();
    assert_eq!(restored.text(), source.text());
    assert!(restored.body().runs().all(|run| !run.format().all_caps()));
    assert!(
        restored
            .body()
            .runs()
            .all(|run| run.format().underline() == UnderlineStyle::None)
    );
    Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();

    let mut overlap = DurableComposition::new(&source, limits(8));
    overlap.join(all_caps).unwrap();
    let mut overlapping_underline_edit = source.edit();
    overlapping_underline_edit
        .set_text_underline(TextSpan::new(0, 5).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let overlapping_underline = overlapping_underline_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert!(overlap.join(overlapping_underline).is_err());
}

#[test]
fn durable_composition_replays_double_strike_and_composes_disjoint_or_rejects_overlap() {
    let source = Document::parse(r"{\rtf1\ansi\strike First Second}").unwrap();
    let mut double_strike_edit = source.edit();
    double_strike_edit
        .set_text_double_strike(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let double_strike = double_strike_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert_eq!(
        double_strike.operations()[0].op,
        "character-double-strike.set"
    );

    let mut underline_edit = source.edit();
    underline_edit
        .set_text_underline(TextSpan::new(6, 12).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let underline = underline_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();

    let mut disjoint = DurableComposition::new(&source, limits(8));
    disjoint.join(double_strike.clone()).unwrap();
    disjoint.join(underline).unwrap();
    let combined = disjoint.finish().unwrap();
    let formatted = source.apply_durable(&combined).unwrap();
    let double_strike_text = formatted
        .body()
        .runs()
        .filter(|run| run.format().double_strike())
        .map(|run| run.text())
        .collect::<String>();
    let single_strike_text = formatted
        .body()
        .runs()
        .filter(|run| run.format().strike())
        .map(|run| run.text())
        .collect::<String>();
    let underlined_text = formatted
        .body()
        .runs()
        .filter(|run| run.format().underline() == UnderlineStyle::Single)
        .map(|run| run.text())
        .collect::<String>();
    assert!(double_strike_text.contains("First"));
    assert!(!double_strike_text.contains("Second"));
    assert!(single_strike_text.contains("First"));
    assert!(single_strike_text.contains("Second"));
    assert!(underlined_text.contains("Second"));
    let reopened = Document::from_bytes(&formatted.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.text(), source.text());

    let restored = formatted.apply_durable(&combined.inverse()).unwrap();
    assert_eq!(restored.text(), source.text());
    assert!(restored.body().runs().all(|run| run.format().strike()));
    assert!(
        restored
            .body()
            .runs()
            .all(|run| !run.format().double_strike())
    );
    assert!(
        restored
            .body()
            .runs()
            .all(|run| run.format().underline() == UnderlineStyle::None)
    );
    Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();

    let mut overlap = DurableComposition::new(&source, limits(8));
    overlap.join(double_strike).unwrap();
    let mut overlapping_underline_edit = source.edit();
    overlapping_underline_edit
        .set_text_underline(TextSpan::new(0, 5).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let overlapping_underline = overlapping_underline_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert!(overlap.join(overlapping_underline).is_err());
}

#[test]
fn durable_join_preflights_combined_reversible_payload_before_mutation() {
    let left_text = "a".repeat(100_000);
    let right_text = "b".repeat(100_000);
    let input = format!(r"{{\rtf1\ansi {} {}}}", left_text, right_text);
    let source = Document::parse(&input).unwrap();
    let second_start = left_text.len().saturating_add(1);
    let first = {
        let mut edit = source.edit_with_limits(Limits::new(1));
        edit.replace_text(TextSpan::new(0, left_text.len()).unwrap(), "A")
            .unwrap();
        edit.commit()
            .unwrap()
            .patch()
            .to_durable(reversible_payload_limits())
            .unwrap()
    };
    let second = {
        let mut edit = source.edit_with_limits(Limits::new(1));
        edit.replace_text(
            TextSpan::new(second_start, second_start.saturating_add(right_text.len())).unwrap(),
            "B",
        )
        .unwrap();
        edit.commit()
            .unwrap()
            .patch()
            .to_durable(reversible_payload_limits())
            .unwrap()
    };
    let mut composition = DurableComposition::new(&source, reversible_payload_limits());
    composition.join(first).unwrap();
    assert!(composition.join(second).is_err());
    assert_eq!(composition.len(), 1);
}

#[test]
fn durable_inverse_projects_formatting_through_longer_and_shorter_replacements() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    for (replacement, expected_text, bold) in
        [("Longer", "Longer Second", true), ("X", "X Second", false)]
    {
        let mut edit = source.edit_with_limits(Limits::new(4));
        edit.replace_text(TextSpan::new(0, 5).unwrap(), replacement)
            .unwrap();
        if bold {
            edit.set_text_bold(TextSpan::new(6, 12).unwrap(), true)
                .unwrap();
        } else {
            edit.set_text_italic(TextSpan::new(6, 12).unwrap(), true)
                .unwrap();
        }
        let patch = edit
            .commit()
            .unwrap()
            .patch()
            .to_durable(limits(8))
            .unwrap();

        let mut composition = DurableComposition::new(&source, limits(8));
        composition.join(patch).unwrap();
        let combined = composition.finish().unwrap();
        let target = source.apply_durable(&combined).unwrap();
        assert_eq!(target.text(), expected_text);
        let restored = target.apply_durable(&combined.inverse()).unwrap();
        assert_eq!(restored.text(), source.text());
        if bold {
            assert!(restored.body().runs().all(|run| !run.format().bold()));
        } else {
            assert!(restored.body().runs().all(|run| !run.format().italic()));
        }
    }
}

#[test]
fn durable_three_way_requires_explicit_choice_and_refuses_stale_or_foreign() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    let left_patch = replace(&source, TextSpan::new(0, 5).unwrap(), "Left", 8);
    let right_patch = replace(&source, TextSpan::new(0, 5).unwrap(), "Right", 8);
    let mut left = DurableComposition::new(&source, limits(8));
    left.join(left_patch.clone()).unwrap();
    let mut right = DurableComposition::new(&source, limits(8));
    right.join(right_patch).unwrap();
    let mut plan = DurableMergePlan::new(left, right).unwrap();
    assert_eq!(plan.conflicts().len(), 1);
    plan = *plan.finish().unwrap_err();
    plan.resolve(MergeResolution::Right);
    let merged = plan.finish().unwrap().finish().unwrap();
    let result = source.apply_durable(&merged).unwrap();
    assert_eq!(result.text(), "Right Second");

    let foreign = Document::parse(r"{\rtf1\ansi Foreign}").unwrap();
    assert!(matches!(
        foreign.apply_durable(&merged),
        Err(litchi_rtf::edit::Error::PatchConflict)
    ));
}

#[test]
fn durable_join_respects_combined_operation_limit_without_publishing() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    let first = replace(&source, TextSpan::new(0, 5).unwrap(), "Alpha", 1);
    let second = replace(&source, TextSpan::new(6, 12).unwrap(), "Bravo", 1);
    let mut composition = DurableComposition::new(&source, limits(1));
    composition.join(first).unwrap();
    assert!(composition.join(second).is_err());
    assert_eq!(composition.len(), 1);
    assert!(source.to_bytes().is_ok());
}

#[test]
fn durable_replay_uses_the_exact_patch_operation_limit() {
    const OPERATION_COUNT: usize = 257;
    let input = format!(r"{{\rtf1\ansi {}}}", "a".repeat(OPERATION_COUNT));
    let source = Document::parse(&input).unwrap();
    let mut edit = source.edit_with_limits(Limits::new(OPERATION_COUNT));
    for position in 0..OPERATION_COUNT {
        edit.replace_text(TextSpan::new(position, position + 1).unwrap(), "b")
            .unwrap();
    }
    let patch = edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(OPERATION_COUNT))
        .unwrap();
    let target = source.apply_durable(&patch).unwrap();
    assert_eq!(target.text(), "b".repeat(OPERATION_COUNT));
    let restored = target.apply_durable(&patch.inverse()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
}

#[test]
fn empty_body_boundary_group_still_compares_alignment_conflicts() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    let mut left_edit = source.edit_with_limits(Limits::new(2));
    left_edit
        .replace_text(TextSpan::new(5, 5).unwrap(), "Inserted")
        .unwrap();
    left_edit
        .set_paragraph_alignment(0, Alignment::Right)
        .unwrap();
    let left_patch = left_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(4))
        .unwrap();

    let mut right_edit = source.edit_with_limits(Limits::new(2));
    right_edit
        .replace_text(TextSpan::new(0, 5).unwrap(), "Alpha")
        .unwrap();
    right_edit
        .set_paragraph_alignment(0, Alignment::Right)
        .unwrap();
    let right_patch = right_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(4))
        .unwrap();

    let mut left = DurableComposition::new(&source, limits(4));
    left.join(left_patch).unwrap();
    let mut right = DurableComposition::new(&source, limits(4));
    right.join(right_patch).unwrap();
    let mut plan = DurableMergePlan::new(left, right).unwrap();
    assert!(!plan.conflicts().is_empty());
    plan = *plan.finish().unwrap_err();
    plan.resolve(MergeResolution::Left);
    let merged = plan.finish().unwrap().finish().unwrap();
    let result = source.apply_durable(&merged).unwrap();
    assert_eq!(result.text(), "FirstInserted Second");
    assert_eq!(
        result
            .body()
            .paragraphs()
            .next()
            .unwrap()
            .format()
            .alignment(),
        Alignment::Right
    );
}

#[test]
fn durable_composition_replays_font_size_with_disjoint_facets_and_rejects_overlap() {
    let source = Document::parse(r"{\rtf1\ansi First Second}").unwrap();
    let mut size_edit = source.edit();
    size_edit
        .set_text_font_size(TextSpan::new(0, 5).unwrap(), NonZeroU16::new(23).unwrap())
        .unwrap();
    let size = size_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    assert_eq!(size.operations()[0].op, "character-font-size.set");
    assert_eq!(
        size.operations()[0].preconditions["font_size_half_points"],
        serde_json::Value::Number(serde_json::Number::from(24_u64))
    );

    let mut underline_edit = source.edit();
    underline_edit
        .set_text_underline(TextSpan::new(6, 12).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let underline = underline_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();

    let mut composition = DurableComposition::new(&source, limits(8));
    composition.join(size.clone()).unwrap();
    composition.join(underline.clone()).unwrap();
    let combined = composition.finish().unwrap();
    let formatted = source.apply_durable(&combined).unwrap();
    let size_text = formatted
        .body()
        .runs()
        .filter(|run| run.format().size().get() == 23)
        .map(|run| run.text())
        .collect::<String>();
    assert!(size_text.contains("First"));
    let underlined_text = formatted
        .body()
        .runs()
        .filter(|run| run.format().underline() == UnderlineStyle::Single)
        .map(|run| run.text())
        .collect::<String>();
    assert!(underlined_text.contains("Second"));
    let reopened = Document::from_bytes(&formatted.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.text(), source.text());

    let restored = formatted.apply_durable(&combined.inverse()).unwrap();
    assert!(
        restored
            .body()
            .runs()
            .all(|run| run.format().size().get() == 24)
    );
    assert!(
        restored
            .body()
            .runs()
            .all(|run| run.format().underline() == UnderlineStyle::None)
    );
    let reopened_restored = Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened_restored.text(), source.text());

    let mut overlapping_underline_edit = source.edit();
    overlapping_underline_edit
        .set_text_underline(TextSpan::new(0, 5).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let overlapping_underline = overlapping_underline_edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(limits(8))
        .unwrap();
    let mut overlap = DurableComposition::new(&source, limits(8));
    overlap.join(size).unwrap();
    assert!(overlap.join(overlapping_underline).is_err());
}
