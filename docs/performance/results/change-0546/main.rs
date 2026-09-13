#![forbid(unsafe_code)]
mod baseline;
mod counted;
use std::{env, fs, hint::black_box, time::Instant};

fn main() {
    let args: Vec<String> = env::args().collect();
    assert_eq!(args.len(), 4, "variant fixture output");
    let input = fs::read(&args[2]).expect("fixture");
    let function: fn(&[u8]) -> bool = match args[1].as_str() {
        "baseline" => baseline::baseline,
        "counted" => counted::counted,
        _ => panic!("variant"),
    };
    let expected = baseline::baseline(&input);
    assert_eq!(counted::counted(&input), expected);
    let full_bound = 1
        + usize::from(input.first().is_some_and(|b| !matches!(b, b'<' | b'&')))
        + input.iter().filter(|b| matches!(b, b'<' | b'&')).count()
        + input
            .windows(2)
            .filter(|w| matches!(w[0], b'>' | b';') && !matches!(w[1], b'<' | b'&'))
            .count();
    assert_eq!(expected, full_bound <= 131_072);
    let mut reader = quick_xml::NsReader::from_reader(input.as_slice());
    let config = reader.config_mut();
    config.allow_dangling_amp = false;
    config.allow_unmatched_ends = false;
    config.check_comments = false;
    config.check_end_names = true;
    config.expand_empty_elements = false;
    config.trim_markup_names_in_closing_tags = true;
    config.trim_text(false);
    let mut reader_events = 0usize;
    loop {
        let event = reader.read_event().expect("valid diagnostic XML");
        let eof = matches!(event, quick_xml::events::Event::Eof);
        let _ = reader.resolver().resolve_event(event);
        reader_events += 1;
        if eof {
            break;
        }
    }
    assert!(reader_events <= full_bound);
    for _ in 0..10 {
        black_box(function(black_box(&input)));
    }
    let mut samples = Vec::with_capacity(100);
    for _ in 0..100 {
        let clock = Instant::now();
        let result = function(black_box(&input));
        let elapsed = clock.elapsed().as_nanos();
        assert_eq!(black_box(result), expected);
        samples.push(elapsed);
    }
    let report = serde_json::json!({"variant":args[1],"fixture":args[2],"bytes":input.len(),"full_bound":full_bound,"reader_events":reader_events,"eligible":expected,"warmup":10,"samples":100,"duration_ns":samples,"scope":"isolated lexical admission only; no workflow or retained runtime claim"});
    fs::write(&args[3], serde_json::to_vec_pretty(&report).unwrap()).unwrap();
}

#[cfg(test)]
mod tests {
    use super::*;
    use quick_xml::{NsReader, events::Event};

    fn full_bound(input: &[u8]) -> usize {
        1 + usize::from(input.first().is_some_and(|b| !matches!(b, b'<' | b'&')))
            + input.iter().filter(|b| matches!(b, b'<' | b'&')).count()
            + input
                .windows(2)
                .filter(|w| matches!(w[0], b'>' | b';') && !matches!(w[1], b'<' | b'&'))
                .count()
    }
    fn check(input: &[u8]) {
        assert_eq!(baseline::baseline(input), full_bound(input) <= 131_072);
        assert_eq!(counted::counted(input), baseline::baseline(input));
        let mut reader = NsReader::from_reader(input);
        let config = reader.config_mut();
        config.allow_dangling_amp = false;
        config.allow_unmatched_ends = false;
        config.check_comments = false;
        config.check_end_names = true;
        config.expand_empty_elements = false;
        config.trim_markup_names_in_closing_tags = true;
        config.trim_text(false);
        let mut events = 0;
        while let Ok(event) = reader.read_event() {
            let eof = matches!(event, Event::Eof);
            let _ = reader.resolver().resolve_event(event);
            events += 1;
            if eof {
                break;
            }
        }
        assert!(events <= full_bound(input));
    }
    #[test]
    fn reader_edges_and_chunk_boundaries() {
        for atom in [
            b"".as_slice(),
            b" ",
            b"<x/>",
            b"<x>a&amp;b</x>",
            b"<!-- > & ; -->",
            b"<![CDATA[<&>]]>",
            b"<?xml version=\"1.0\"?><x/>",
            b"<!DOCTYPE x [<!ENTITY a \"<&>\">]><x/>",
            b"<x>&broken</x>",
            b"<x><",
            b"\xef\xbb\xbf<x/>",
        ] {
            check(atom);
            for n in [4094, 4095, 4096, 4097, 8191, 8192, 8193] {
                let mut input = vec![b' '; n];
                input.extend_from_slice(atom);
                check(&input);
            }
        }
        for n in [131_070, 131_071, 131_072, 131_073] {
            check(&b"<!--x-->".repeat(n));
        }
    }
    #[test]
    fn exact_text_cap_and_explicit_boundary_pairs() {
        let mut exact = b"<x/>a".repeat(65_535);
        exact.extend_from_slice(b"<x/>");
        assert_eq!(full_bound(&exact), 131_072);
        check(&exact);
        exact.push(b'a');
        assert_eq!(full_bound(&exact), 131_073);
        check(&exact);
        for offset in [4095, 4096, 8191, 8192, 65535, 65536] {
            for end in [b'>', b';'] {
                for next in [b'<', b'&', b'a'] {
                    let mut data = vec![b' '; offset];
                    data.extend_from_slice(&[end, next]);
                    let count = 131_072 - full_bound(&data);
                    data.extend_from_slice(&b"<!--x-->".repeat(count));
                    assert_eq!(full_bound(&data), 131_072);
                    check(&data);
                    data.extend_from_slice(b"<!--x-->");
                    check(&data);
                }
            }
        }
        for final_byte in [b'<', b'&', b'>', b';'] {
            check(&[final_byte]);
        }
    }
    #[test]
    fn deterministic_arbitrary_bytes() {
        let alphabet = b"<>&; abc/?!=\"\n\x00\xff";
        let mut state = 12345u64;
        for len in 0..512 {
            let mut input = Vec::with_capacity(len);
            for _ in 0..len {
                state = state.wrapping_mul(6364136223846793005).wrapping_add(1);
                input.push(alphabet[(state >> 32) as usize % alphabet.len()]);
            }
            check(&input);
        }
    }
}
