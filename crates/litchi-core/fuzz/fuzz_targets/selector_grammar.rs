#![no_main]

//! Path and query grammar fuzz target.
//!
//! Any UTF-8 input either fails with a typed error whose offset is a char
//! boundary inside the input, or parses to a value whose canonical display
//! parses back to an equal value (parse -> display -> parse identity) and
//! whose display is a fixed point.

use libfuzzer_sys::fuzz_target;
use litchi_core::selector::{Path, Query, limits};

fuzz_target!(|data: &[u8]| {
    if data.len() > limits::MAX_INPUT_BYTES + 16 {
        return;
    }
    let Ok(input) = std::str::from_utf8(data) else {
        return;
    };

    for relative in [false, true] {
        let parsed = if relative {
            Path::parse_relative(input)
        } else {
            Path::parse(input)
        };
        match parsed {
            Ok(path) => {
                let shown = path.to_string();
                let again = if relative {
                    Path::parse_relative(&shown)
                } else {
                    Path::parse(&shown)
                }
                .expect("displayed path parses");
                assert_eq!(again, path, "{input:?} -> {shown:?}");
                assert_eq!(again.to_string(), shown);
                assert_eq!(path.clone().into_owned(), path);
            },
            Err(error) => {
                assert!(error.offset() <= input.len());
                assert!(input.is_char_boundary(error.offset()));
            },
        }
    }

    match Query::parse(input) {
        Ok(query) => {
            let shown = query.to_string();
            let again = Query::parse(&shown).expect("displayed query parses");
            assert_eq!(again, query, "{input:?} -> {shown:?}");
            assert_eq!(again.to_string(), shown);
            assert_eq!(query.clone().into_owned(), query);
        },
        Err(error) => {
            assert!(error.offset() <= input.len());
            assert!(input.is_char_boundary(error.offset()));
        },
    }
});
