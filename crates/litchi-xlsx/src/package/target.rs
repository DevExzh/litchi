//! The publication target of a clone-staged scoped transaction.
//!
//! The scoped package transactions (slicers, timelines, Data Model, ...)
//! stage every change on a clone of the OPC graph and publish the complete
//! candidate once, at commit. They can run directly over a raw
//! [`OpcPackage`], or over a [`crate::Package`], whose commit revalidates the
//! candidate as a complete workbook before the owner adopts it.

use std::ops::Deref;

use litchi_opc::OpcPackage;

use crate::error::Result;

/// Where a scoped transaction publishes its staged OPC graph.
#[derive(Debug)]
pub(crate) enum Target<'a> {
    /// A raw OPC graph, replaced verbatim.
    Opc(&'a mut OpcPackage),
    /// A validated package owner, republished through workbook validation.
    Package(&'a mut crate::Package),
}

impl Target<'_> {
    /// Replace the target with `staged`. A package owner validates the
    /// candidate first and is unchanged when validation fails.
    pub(crate) fn publish(&mut self, staged: OpcPackage) -> Result<()> {
        match self {
            Self::Opc(package) => {
                **package = staged;
                Ok(())
            },
            Self::Package(package) => package.republish(staged),
        }
    }
}

impl Deref for Target<'_> {
    type Target = OpcPackage;

    fn deref(&self) -> &OpcPackage {
        match self {
            Self::Opc(package) => package,
            Self::Package(package) => package.opc(),
        }
    }
}

impl<'a> From<&'a mut OpcPackage> for Target<'a> {
    fn from(package: &'a mut OpcPackage) -> Self {
        Self::Opc(package)
    }
}
