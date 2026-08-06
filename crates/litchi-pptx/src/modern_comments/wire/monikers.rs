use super::super::PC2;
use super::super::semantic::extensions::OpaqueXml;
use super::super::semantic::monikers::{Kind, List, Node};
use super::xml::{attr, attribute, close, no_attributes, open, scan, scan_with_context};
use crate::{Error, Result};

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn parse_monikers(xml: &[u8]) -> Result<List> {
    let scan = scan(xml, "comment moniker list")?;
    if scan.root.namespace != PC2 {
        return Err(invalid("comment moniker list has the wrong namespace"));
    }
    let kind = match scan.root.local.as_str() {
        "cmMkLst" => Kind::Comment,
        "cmRplyMkLst" => Kind::Reply,
        _ => return Err(invalid("unknown comment moniker list root")),
    };
    no_attributes(&scan.root.attributes, "comment moniker list")?;
    let root_scan = scan;
    let mut nodes = Vec::with_capacity(root_scan.children.len());
    for child in &root_scan.children {
        let child_scan = scan_with_context(&child.xml, "comment moniker", &root_scan.namespaces)?;
        let child = &child_scan.root;
        if child.namespace == PC2 && child.local == "cmMk" {
            no_attributes_except_id(&child.attributes)?;
            nodes.push(Node::Comment {
                id: attribute(&child.attributes, "id", true)?
                    .unwrap()
                    .to_owned(),
            });
        } else if child.namespace == PC2 && child.local == "cmRplyMk" {
            no_attributes_except_id(&child.attributes)?;
            nodes.push(Node::Reply {
                id: attribute(&child.attributes, "id", true)?
                    .unwrap()
                    .to_owned(),
            });
        } else {
            nodes.push(Node::Opaque(OpaqueXml::new(child.xml.clone())?));
        }
    }
    let value = List { kind, nodes };
    value.validate()?;
    Ok(value)
}

fn no_attributes_except_id(attributes: &[(String, String)]) -> Result<()> {
    for (key, _) in attributes {
        if key != "id" {
            return Err(invalid(format!(
                "unexpected comment moniker attribute '{key}'"
            )));
        }
    }
    Ok(())
}

pub(super) fn write_monikers(value: &List) -> Result<Vec<u8>> {
    value.validate()?;
    let mut out = Vec::new();
    let root = match value.kind {
        Kind::Comment => "cmMkLst",
        Kind::Reply => "cmRplyMkLst",
    };
    open(&mut out, "pc2", root);
    out.extend_from_slice(
        b" xmlns:pc2=\"http://schemas.microsoft.com/office/powerpoint/2019/9/main/command\"",
    );
    if value.nodes.is_empty() {
        out.extend_from_slice(b"/>");
        return Ok(out);
    }
    out.push(b'>');
    for node in &value.nodes {
        match node {
            Node::Opaque(value) => out.extend_from_slice(value.as_bytes()),
            Node::Comment { id } => {
                open(&mut out, "pc2", "cmMk");
                attr(&mut out, "id", id);
                out.extend_from_slice(b"/>");
            },
            Node::Reply { id } => {
                open(&mut out, "pc2", "cmRplyMk");
                attr(&mut out, "id", id);
                out.extend_from_slice(b"/>");
            },
        }
    }
    close(&mut out, "pc2", root);
    Ok(out)
}
