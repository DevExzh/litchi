//! ODS package-boundary publication for annotation commits.

use litchi_core::Result;

/// Replace only the owned `content.xml` member after a validated annotation
/// transaction.  The package layer retains every other member unchanged.
pub(crate) fn replace_content(
    package: &crate::package::Package,
    content_xml: &str,
) -> Result<crate::package::Package> {
    package.replace_content_xml(content_xml)
}
