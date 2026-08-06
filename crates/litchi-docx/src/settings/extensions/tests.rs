use super::*;
use litchi_opc::constants::content_type as ct;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::BlobPart;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14: &str = WORD_2010_NAMESPACE;
const W15: &str = WORD_2012_NAMESPACE;
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

#[test]
fn package_processing_understands_ignorable_word_settings_extensions() {
    let xml = format!(
        r#"<w:settings xmlns:w="{W}" xmlns:w14="{W14}" xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w14 w15"><w14:conflictMode/><w15:chartTrackingRefBased/></w:settings>"#
    );
    let part = BlobPart::new(
        PackURI::new("/word/settings.xml").unwrap(),
        ct::WML_SETTINGS.to_owned(),
        xml.into_bytes(),
    );
    let processed = process_part(&part).unwrap();
    let processed = std::str::from_utf8(processed.as_ref()).unwrap();
    assert!(processed.contains("conflictMode"));
    assert!(processed.contains("chartTrackingRefBased"));
}
