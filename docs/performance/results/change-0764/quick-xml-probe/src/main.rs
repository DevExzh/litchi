use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::time::Instant;

fn dump(label: &str, xml: &str) {
    let mut r = Reader::from_str(xml);
    loop {
        match r.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e) => {
                println!("{label}: checked");
                for a in e.attributes() {
                    match a {
                        Ok(a) => println!("   Ok {}={:?}", String::from_utf8_lossy(a.key.as_ref()), String::from_utf8_lossy(&a.value)),
                        Err(err) => println!("   Err {err:?}"),
                    }
                }
                println!("{label}: unchecked");
                for a in e.attributes().with_checks(false) {
                    match a {
                        Ok(a) => println!("   Ok {}={:?}", String::from_utf8_lossy(a.key.as_ref()), String::from_utf8_lossy(&a.value)),
                        Err(err) => println!("   Err {err:?}"),
                    }
                }
                break;
            }
            Event::Eof => break,
            _ => {}
        }
    }
}

fn tolerant_cost(distinct: usize, repeats: usize) {
    let mut tag = String::from("e");
    for i in 0..distinct { tag.push_str(&format!(" a{i}=\"\"")); }
    let last = format!("a{}", distinct - 1);
    for _ in 0..repeats { tag.push_str(&format!(" {last}=\"\"")); }
    let bytes = tag.len();
    let e = BytesStart::from_content(tag, 1);
    let t = Instant::now();
    let ok = e.attributes().flatten().count();
    let dt = t.elapsed();
    println!("tolerant distinct={distinct} repeats={repeats} tag_bytes={bytes} ok={ok} elapsed={dt:?}");
    let t = Instant::now();
    let failfast = e.attributes().take_while(|a| a.is_ok()).count();
    println!("   fail-fast ok_prefix={failfast} elapsed={:?}", t.elapsed());
}

fn main() {
    dump("dup-with-space", r#"<e a="1" a="x b='evil'"/>"#);
    dump("dup-unquoted-first", r#"<e a=x a="1"/>"#);
    dump("dup-plain", r#"<e a="1" b="2" a="3" c="4"/>"#);
    for &(d, r) in &[(1000usize, 1000usize), (4000, 4000), (16000, 16000)] {
        tolerant_cost(d, r);
    }
}
