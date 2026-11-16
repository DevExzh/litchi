//! A minimal record tree for the containers an equation edit opens.
//!
//! Only the containers on the path to an edited record are expanded; every
//! other record (atoms and unexpanded containers) keeps its exact bytes, so
//! encoding an unmodified tree reproduces its source bytes.
#![cfg_attr(
    not(feature = "equation"),
    allow(
        dead_code,
        reason = "the builders and the encoder serve the equation writer"
    )
)]

use crate::package::{Error, Result};

/// Deepest container nesting an equation edit opens.
const MAX_DEPTH: usize = 64;

/// One record: its eight-byte header and either its exact bytes or, for an
/// expanded container, its children.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct Node {
    pub(super) header: [u8; 8],
    pub(super) body: Body,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) enum Body {
    /// The payload bytes, kept as they are.
    Raw(Vec<u8>),
    /// The children of an expanded container, with any bytes after the last
    /// complete child (too few to hold a record header).
    Children(Vec<Node>, Vec<u8>),
}

impl Node {
    /// A new atom.
    pub(super) fn atom(version: u16, instance: u16, record_type: u16, payload: Vec<u8>) -> Self {
        Self {
            header: header(version, instance, record_type, 0),
            body: Body::Raw(payload),
        }
    }

    /// A new container.
    pub(super) fn container(record_type: u16, instance: u16, children: Vec<Self>) -> Self {
        Self {
            header: header(0xF, instance, record_type, 0),
            body: Body::Children(children, Vec::new()),
        }
    }

    pub(super) fn record_type(&self) -> u16 {
        u16::from_le_bytes([self.header[2], self.header[3]])
    }

    pub(super) fn instance(&self) -> u16 {
        u16::from_le_bytes([self.header[0], self.header[1]]) >> 4
    }

    pub(super) fn set_instance(&mut self, instance: u16) {
        let version = u16::from_le_bytes([self.header[0], self.header[1]]) & 0xF;
        self.header[..2].copy_from_slice(&((instance << 4) | version).to_le_bytes());
    }

    fn is_container(&self) -> bool {
        self.header[0] & 0xF == 0xF
    }

    /// The payload of an atom (or of an unexpanded container).
    pub(super) fn payload(&self) -> Option<&[u8]> {
        match &self.body {
            Body::Raw(bytes) => Some(bytes),
            Body::Children(..) => None,
        }
    }

    pub(super) fn payload_mut(&mut self) -> Option<&mut Vec<u8>> {
        match &mut self.body {
            Body::Raw(bytes) => Some(bytes),
            Body::Children(..) => None,
        }
    }

    /// The children of an expanded container.
    pub(super) fn children(&self) -> &[Self] {
        match &self.body {
            Body::Children(children, _) => children,
            Body::Raw(_) => &[],
        }
    }

    pub(super) fn children_mut(&mut self) -> Option<&mut Vec<Self>> {
        match &mut self.body {
            Body::Children(children, _) => Some(children),
            Body::Raw(_) => None,
        }
    }

    /// Expand this container's children one level (no-op when expanded).
    pub(super) fn expand(&mut self) -> Result<()> {
        if !self.is_container() {
            return Err(corrupted("an equation edit expected a record container"));
        }
        if let Body::Raw(bytes) = &self.body {
            let (children, tail) = parse_sequence(bytes)?;
            self.body = Body::Children(children, tail);
        }
        Ok(())
    }

    /// The first child of `record_type`, expanding this container.
    pub(super) fn child_mut(&mut self, record_type: u16) -> Result<Option<&mut Self>> {
        self.expand()?;
        Ok(self.children_mut().and_then(|children| {
            children
                .iter_mut()
                .find(|child| child.record_type() == record_type)
        }))
    }

    /// The first child of `record_type` of an expanded container.
    pub(super) fn child(&self, record_type: u16) -> Option<&Self> {
        self.children()
            .iter()
            .find(|child| child.record_type() == record_type)
    }

    /// Expand this container and every container below it (bounded depth),
    /// for reading.
    pub(super) fn expand_all(&mut self, depth: usize) -> Result<()> {
        if depth > MAX_DEPTH {
            return Err(corrupted(
                "PPT drawing nesting exceeds the equation-edit limit",
            ));
        }
        // Some atoms carry recVer 0xF (PowerPoint 2007's
        // `RoundTripCustomTableStyles12Atom`): only a payload that is a
        // complete record sequence is expanded.
        if !self.is_container() || self.payload().is_some_and(|payload| !is_sequence(payload)) {
            return Ok(());
        }
        self.expand()?;
        if let Some(children) = self.children_mut() {
            for child in children {
                child.expand_all(depth + 1)?;
            }
        }
        Ok(())
    }

    /// The encoded record.
    pub(super) fn encode(&self) -> Result<Vec<u8>> {
        let mut out = Vec::new();
        self.encode_into(&mut out)?;
        Ok(out)
    }

    fn encode_into(&self, out: &mut Vec<u8>) -> Result<()> {
        let start = out.len();
        out.extend_from_slice(&self.header);
        match &self.body {
            Body::Raw(bytes) => out.extend_from_slice(bytes),
            Body::Children(children, tail) => {
                for child in children {
                    child.encode_into(out)?;
                }
                out.extend_from_slice(tail);
            },
        }
        let len = u32::try_from(out.len() - start - 8)
            .map_err(|_error| Error::ResourceLimit("PPT record exceeds 4 GiB".into()))?;
        out[start + 4..start + 8].copy_from_slice(&len.to_le_bytes());
        Ok(())
    }
}

/// Whether `payload` is a complete sequence of records, each within it.
pub(super) fn is_sequence(payload: &[u8]) -> bool {
    let mut offset = 0usize;
    while offset < payload.len() {
        let Some(header) = payload.get(offset..offset + 8) else {
            return false;
        };
        let len = u32::from_le_bytes([header[4], header[5], header[6], header[7]]) as usize;
        match offset
            .checked_add(8)
            .and_then(|start| start.checked_add(len))
        {
            Some(end) if end <= payload.len() => offset = end,
            _ => return false,
        }
    }
    true
}

fn header(version: u16, instance: u16, record_type: u16, len: u32) -> [u8; 8] {
    let mut header = [0; 8];
    header[..2].copy_from_slice(&((instance << 4) | (version & 0xF)).to_le_bytes());
    header[2..4].copy_from_slice(&record_type.to_le_bytes());
    header[4..].copy_from_slice(&len.to_le_bytes());
    header
}

/// Parse one complete record; the whole of `bytes` must be that record.
pub(super) fn parse(bytes: &[u8]) -> Result<Node> {
    let (mut nodes, tail) = parse_sequence(bytes)?;
    if nodes.len() != 1 || !tail.is_empty() {
        return Err(corrupted(
            "a persisted PPT record is not exactly one record",
        ));
    }
    Ok(nodes.remove(0))
}

fn parse_sequence(bytes: &[u8]) -> Result<(Vec<Node>, Vec<u8>)> {
    let mut nodes = Vec::new();
    let mut offset = 0usize;
    while bytes.len() - offset >= 8 {
        let mut header = [0; 8];
        header.copy_from_slice(&bytes[offset..offset + 8]);
        let len = u32::from_le_bytes([header[4], header[5], header[6], header[7]]) as usize;
        let end = offset
            .checked_add(8)
            .and_then(|start| start.checked_add(len))
            .filter(|end| *end <= bytes.len())
            .ok_or_else(|| corrupted("a PPT record overruns its container"))?;
        nodes.push(Node {
            header,
            body: Body::Raw(bytes[offset + 8..end].to_vec()),
        });
        offset = end;
    }
    Ok((nodes, bytes[offset..].to_vec()))
}

fn corrupted(message: &str) -> Error {
    Error::Corrupted(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn encoding_an_expanded_tree_reproduces_its_bytes() {
        let leaf = Node::atom(0, 3, 0x0FBA, b"ab".to_vec());
        let inner = Node::container(0xF004, 0, vec![leaf.clone()]);
        let outer = Node::container(0xF003, 0, vec![inner, leaf]);
        let bytes = outer.encode().unwrap();
        let mut parsed = parse(&bytes).unwrap();
        parsed.expand_all(0).unwrap();
        assert_eq!(parsed.encode().unwrap(), bytes);
        assert_eq!(parsed.children().len(), 2);
        assert_eq!(parsed.children()[1].instance(), 3);
        assert!(parse(&bytes[..bytes.len() - 1]).is_err());
    }
}
