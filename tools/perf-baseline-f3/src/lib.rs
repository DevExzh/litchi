#![forbid(unsafe_code)]

#[path = "../../perf-baseline/src/xls_f3_corpus.rs"]
mod xls_f3_corpus;

pub(crate) fn xls_f3_number_fixture() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    xls_f3_corpus::build_xls_comments_fixture()
}

pub(crate) mod xls_numeric {
    pub(crate) fn build_rk_mulrk_fixture() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        super::xls_f3_corpus::build_rk_mulrk_fixture()
    }
}

#[path = "../../perf-baseline/src/xls_f3_qualification.rs"]
pub mod xls_f3_qualification;
