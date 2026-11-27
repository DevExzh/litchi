//! Package and flat-document settings dispatch.

use super::codec::{DocumentKind, parse};
use super::model::Settings;
use litchi_core::Result;

pub(crate) fn parse_flat(xml: &str) -> Result<Settings> {
    parse(xml, DocumentKind::Flat)
}

pub(crate) fn parse_package(xml: &str) -> Result<Settings> {
    parse(xml, DocumentKind::Package)
}
