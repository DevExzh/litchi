//! The shared text encoder must preserve every Unicode scalar through RTF's
//! signed UTF-16 transport, including adjacent surrogate pairs.
use litchi_rtf::{Package, write::Writer};

#[test]
fn text_writer_uses_signed_utf16_units_and_round_trips_scalars() {
    let text = "ASCII é 水 \u{ffff} 🦀😀\u{10ffff}";
    let mut bytes = br"{\rtf1\ansi\uc1 ".to_vec();
    Writer::new(&mut bytes).write_text(text).unwrap();
    bytes.push(b'}');
    let wire = std::str::from_utf8(&bytes).unwrap();
    assert!(wire.contains(r"\u-1?"));
    assert!(wire.contains(r"\u-10178?\u-8832?"));
    let reopened = Package::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.text(), text);
}
