use super::validation::part_name_conflicts;

#[test]
fn part_name_conflicts_handles_exact_parent_and_descendant_paths() {
    let existing = vec![
        "/ppt/slides/slide1.xml".to_owned(),
        "/ppt/tags".to_owned(),
        "/ppt/tags/tag2.xml".to_owned(),
    ];

    assert!(part_name_conflicts(&existing, "/ppt/tags"));
    assert!(part_name_conflicts(&existing, "/ppt/tags/tag1.xml"));
    assert!(!part_name_conflicts(&existing, "/ppt/tags2/tag1.xml"));
}
