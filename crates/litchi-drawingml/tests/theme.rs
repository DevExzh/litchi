use litchi_drawingml::theme::{self, Color, Face, FontSet, Palette, Slot, System};

fn palette() -> Palette {
    Slot::ALL
        .into_iter()
        .fold(Palette::new("Office"), |palette, slot| {
            let color = if slot == Slot::Dark1 {
                Color::system(System::WindowText, Some("000000")).unwrap()
            } else {
                Color::rgb("4F81BD").unwrap()
            };
            palette.with(slot, color)
        })
}

#[test]
fn theme_schema_model_round_trips_without_package_types() {
    let colors = palette();
    let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
    let xml = theme::codec::encode_part("Office", &colors, &fonts).unwrap();
    let parsed = theme::codec::read(&xml).unwrap();

    assert_eq!(parsed.name, "Office");
    assert_eq!(parsed.colors, colors);
    assert_eq!(parsed.fonts, fonts);
}

#[test]
fn strict_theme_namespace_is_accepted() {
    let xml = br#"<a:theme xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" name="Strict"><a:themeElements><a:clrScheme name="Strict"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="111111"/></a:accent1><a:accent2><a:srgbClr val="222222"/></a:accent2><a:accent3><a:srgbClr val="333333"/></a:accent3><a:accent4><a:srgbClr val="444444"/></a:accent4><a:accent5><a:srgbClr val="555555"/></a:accent5><a:accent6><a:srgbClr val="666666"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Strict"><a:majorFont><a:latin typeface="Aptos"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"#;
    assert_eq!(theme::codec::read(xml).unwrap().name, "Strict");
}

/// Change 0765: quick-xml skips a leading byte-order mark without counting
/// it, and the scheme scanner used to splice at its uncounted positions. In an
/// indented theme the splice then dropped three indentation bytes and kept the
/// replaced scheme's last three bytes (`me>`) as text inside
/// `a:themeElements`: well-formed output that the read-back accepted. The
/// replacement of a marked theme must equal the unmarked replacement behind
/// the same mark, compact or indented.
#[test]
fn scheme_replacement_of_a_byte_order_marked_theme_matches_the_unmarked_one() {
    let colors = palette();
    let fonts = FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos"));
    let compact = theme::codec::encode_part("Office", &colors, &fonts).unwrap();
    let mut indented = Vec::new();
    for (index, byte) in compact.iter().enumerate() {
        if *byte == b'<' && index > 0 && compact[index - 1] == b'>' && compact[index + 1] != b'/' {
            indented.extend_from_slice(b"\n    ");
        }
        indented.push(*byte);
    }
    let replacement_colors = colors
        .clone()
        .with(Slot::Accent1, Color::rgb("123456").unwrap());
    let fragment = theme::codec::encode_palette_fragment(&replacement_colors).unwrap();
    for plain in [compact, indented] {
        let marked = [b"\xEF\xBB\xBF".as_slice(), plain.as_slice()].concat();
        let expected = theme::codec::replace_scheme(&plain, b"clrScheme", &fragment).unwrap();
        let replaced = theme::codec::replace_scheme(&marked, b"clrScheme", &fragment).unwrap();
        assert_eq!(
            String::from_utf8_lossy(&replaced[3..]),
            String::from_utf8_lossy(&expected)
        );
        assert_eq!(&replaced[..3], b"\xEF\xBB\xBF");
        assert_eq!(
            theme::codec::read(&replaced).unwrap().colors,
            replacement_colors
        );
    }
}
