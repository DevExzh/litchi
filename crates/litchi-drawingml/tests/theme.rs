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
