use litchi_xlsx::drawing::{SourceDrawing, SvgOwnerState};
fn main() {
    let mut xml = String::from(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main""#,
    );
    for i in 0..128 {
        xml.push_str(&format!(" xmlns:p{i}=\"urn:opaque:{}\"", "x".repeat(1000)));
    }
    xml.push('>');
    for i in 0..32 {
        xml.push_str(&format!(r#"<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{}" name="image"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdRaster"><a:extLst><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}"><asvg:svgBlip r:embed="rIdSvg{i}"/></a:ext></a:extLst></a:blip></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>"#,i+1));
    }
    xml.push_str("</xdr:wsDr>");
    let scan = SourceDrawing::scan(xml.as_bytes()).unwrap();
    let mut retained = 0;
    for p in scan.pictures() {
        if let SvgOwnerState::Embedded(owner) = p.svg_owner() {
            retained += owner.value().source().unwrap().len();
        }
    }
    println!(
        "input_bytes={} pictures={} retained_svg_source_bytes={}",
        xml.len(),
        scan.pictures().len(),
        retained
    );
}
