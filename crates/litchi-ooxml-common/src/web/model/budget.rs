use super::super::codec::limit;
use super::super::*;
use super::*;
#[derive(Debug, Default)]
pub(in crate::web) struct OperationBudget {
    pub(in crate::web) xml_bytes: usize,
    pub(in crate::web) string_bytes: usize,
}

impl OperationBudget {
    pub(in crate::web) fn charge_xml(&mut self, bytes: usize, limits: &Limits) -> Result<()> {
        self.xml_bytes = self.xml_bytes.checked_add(bytes).ok_or(Error::Limit {
            resource: "aggregate web extension XML bytes",
            max: limits.total_xml_bytes,
            actual: usize::MAX,
        })?;
        if self.xml_bytes > limits.total_xml_bytes {
            return limit(
                "aggregate web extension XML bytes",
                limits.total_xml_bytes,
                self.xml_bytes,
            );
        }
        Ok(())
    }

    pub(in crate::web) fn charge_strings(&mut self, bytes: usize, limits: &Limits) -> Result<()> {
        self.string_bytes = self.string_bytes.checked_add(bytes).ok_or(Error::Limit {
            resource: "retained web extension string bytes",
            max: limits.total_string_bytes,
            actual: usize::MAX,
        })?;
        if self.string_bytes > limits.total_string_bytes {
            return limit(
                "retained web extension string bytes",
                limits.total_string_bytes,
                self.string_bytes,
            );
        }
        Ok(())
    }

    pub(in crate::web) fn charge_authored(&mut self, xml: &[u8], limits: &Limits) -> Result<()> {
        self.charge_xml(xml.len(), limits)?;
        self.charge_strings(xml.len(), limits)
    }

    pub(in crate::web) fn charge_metadata(
        &mut self,
        bytes: usize,
        copies: usize,
        limits: &Limits,
    ) -> Result<()> {
        let retained = bytes.checked_mul(copies).ok_or(Error::Limit {
            resource: "indexed web extension package metadata bytes",
            max: limits.total_string_bytes,
            actual: usize::MAX,
        })?;
        self.charge_strings(retained, limits)
    }
}
