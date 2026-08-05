#[test]
fn chart_reader_uses_the_fallback_branch() {
    let raw = br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><x:chart/></mc:Choice><mc:Fallback><c:chart/></mc:Fallback></mc:AlternateContent></c:chartSpace>"#;
    litchi_drawingml::chart::reader::read(raw.as_slice()).unwrap();
}
