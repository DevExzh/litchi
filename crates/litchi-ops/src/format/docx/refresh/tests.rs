use std::collections::HashMap;

use super::{heading_number, numbered, parse_range};
use crate::refresh::{Instruction, Status, sequence};

fn seq(instructions: &[&str]) -> Vec<Result<String, Status>> {
    let mut counters = HashMap::new();
    instructions
        .iter()
        .map(|text| {
            sequence(&Instruction::parse(text), &mut counters).map_err(|(status, _)| status)
        })
        .collect()
}

#[test]
fn sequences_count_in_document_order() {
    assert_eq!(
        seq(&[
            "SEQ Figure",
            r"SEQ Figure \* ROMAN",
            r"SEQ Figure \c",
            r"SEQ Table \* alphabetic",
            r"SEQ Figure \r 10",
            r"SEQ Figure \h",
            r"SEQ Figure \* MERGEFORMAT",
        ]),
        [
            Ok("1".to_owned()),
            Ok("II".to_owned()),
            Ok("2".to_owned()),
            Ok("a".to_owned()),
            Ok("10".to_owned()),
            Ok(String::new()),
            Ok("12".to_owned()),
        ]
    );
    // A sequence Litchi cannot follow stops being computed.
    assert_eq!(
        seq(&[
            r"SEQ A \s 1",
            "SEQ A",
            r"SEQ B \* Ordinal",
            "SEQ B",
            "SEQ C"
        ]),
        [
            Err(Status::Unsupported),
            Err(Status::Unsupported),
            Err(Status::Unsupported),
            Err(Status::Unsupported),
            Ok("1".to_owned()),
        ]
    );
}

#[test]
fn heading_and_range_names_are_parsed() {
    assert_eq!(heading_number("Heading3"), Some(3));
    assert_eq!(heading_number("heading9"), Some(9));
    assert_eq!(heading_number("Heading10"), None);
    assert_eq!(numbered("TOC2", "toc"), Some(2));
    assert_eq!(numbered("Verzeichnis1", "toc"), None);
    assert_eq!(parse_range("1-3"), Some((1, 3)));
    assert_eq!(parse_range(" 2 - 2 "), Some((2, 2)));
    assert_eq!(parse_range("3-1"), None);
    assert_eq!(parse_range("0-9"), None);
}
