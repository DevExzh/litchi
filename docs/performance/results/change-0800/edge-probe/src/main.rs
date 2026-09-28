#![allow(dead_code)]
mod baseline;
mod corrected;
use baseline::BytesStartExt;
use quick_xml::events::BytesStart;
fn main() {
    for n in [1, 31, 32, 33, 34] {
        let mut input = String::from("e");
        for i in 0..n {
            input.push_str(&format!(" =n{i}=\"{i}\""));
        }
        input.push_str(" =n0=\"unterminated");
        let tag = BytesStart::from_content(input, 1);
        let baseline_error = tag.checked_attributes().find_map(Result::err);
        #[allow(clippy::disallowed_methods)]
        let reference_error = tag.attributes().find_map(Result::err);
        let corrected_error =
            corrected::BytesStartExt::checked_attributes(&tag).find_map(Result::err);
        assert_eq!(corrected_error, reference_error);
        assert_eq!(baseline_error == reference_error, n < 32);
        println!(
            "accepted_prefix={n} baseline={baseline_error:?} quick_xml={reference_error:?} equal={} corrected={corrected_error:?}",
            baseline_error == reference_error
        );
    }
}
