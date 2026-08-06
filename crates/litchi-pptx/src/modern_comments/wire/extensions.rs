use super::super::semantic::extensions::{Entry, List, OpaqueXml, Payload};
use super::super::{P, P188, P223, P228};
use super::reactions::{parse_reactions, write_reactions};
use super::tasks::{parse_task_details, write_task_details};
use super::xml::{
    attr, attribute, close, no_attributes, only_attributes, open, scan, scan_with_context,
};
use crate::{Error, Result};

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn parse_extensions(xml: Option<&[u8]>) -> Result<List> {
    let Some(xml) = xml else {
        return Ok(List::default());
    };
    let root_scan = scan(xml, "modern comment extension list")?;
    if root_scan.root.namespace != P188 || root_scan.root.local != "extLst" {
        return Err(invalid("modern comment extension root must be p188:extLst"));
    }
    no_attributes(&root_scan.root.attributes, "modern comment extension list")?;
    let mut entries = Vec::with_capacity(root_scan.children.len());
    for child in &root_scan.children {
        if child.namespace != P || child.local != "ext" {
            return Err(invalid("modern comment extLst permits only p:ext children"));
        }
        only_attributes(&child.attributes, &["uri"], "modern comment extension")?;
        let uri = attribute(&child.attributes, "uri", true)?
            .unwrap()
            .to_owned();
        let payloads = scan_with_context(
            &child.xml,
            "modern comment extension",
            &root_scan.namespaces,
        )?
        .children;
        let payload = if payloads.len() == 1 {
            let payload = &payloads[0];
            if payload.namespace == P228 && payload.local == "taskDetails" {
                let xml = scan_with_context(&payload.xml, "task details", &root_scan.namespaces)?;
                Payload::TaskDetails(parse_task_details(&xml.root.xml)?)
            } else if payload.namespace == P223 && payload.local == "reactions" {
                let xml = scan_with_context(&payload.xml, "reactions", &root_scan.namespaces)?;
                Payload::Reactions(parse_reactions(&xml.root.xml)?)
            } else {
                Payload::Opaque(OpaqueXml::new(child.xml.clone())?)
            }
        } else {
            Payload::Opaque(OpaqueXml::new(child.xml.clone())?)
        };
        entries.push(Entry { uri, payload });
    }
    let value = List {
        root_prefix: String::new(),
        namespace_declarations: root_scan.namespaces,
        entries,
    };
    value.validate()?;
    Ok(value)
}

pub(crate) fn write_extensions(value: &List) -> Result<Option<Vec<u8>>> {
    value.validate()?;
    if value.entries.is_empty() {
        return Ok(None);
    }
    let prefix = if value.root_prefix.is_empty() {
        "p188"
    } else {
        value.root_prefix.as_str()
    };
    let mut out = Vec::new();
    open(&mut out, prefix, "extLst");
    out.extend_from_slice(b" xmlns:");
    out.extend_from_slice(prefix.as_bytes());
    out.extend_from_slice(b"=\"http://schemas.microsoft.com/office/powerpoint/2018/8/main\"");
    out.extend_from_slice(
        b" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"",
    );
    for declaration in &value.namespace_declarations {
        if declaration.prefix == prefix || (declaration.prefix == "p" && declaration.uri == P) {
            continue;
        }
        super::xml::namespaces(&mut out, std::slice::from_ref(declaration));
    }
    out.push(b'>');
    for entry in &value.entries {
        match &entry.payload {
            Payload::Opaque(raw) => out.extend_from_slice(raw.as_bytes()),
            Payload::TaskDetails(task) => {
                open(&mut out, "p", "ext");
                attr(&mut out, "uri", &entry.uri);
                out.push(b'>');
                out.extend_from_slice(&write_task_details(task)?);
                close(&mut out, "p", "ext");
            },
            Payload::Reactions(reactions) => {
                open(&mut out, "p", "ext");
                attr(&mut out, "uri", &entry.uri);
                out.push(b'>');
                out.extend_from_slice(&write_reactions(reactions)?);
                close(&mut out, "p", "ext");
            },
        }
    }
    close(&mut out, prefix, "extLst");
    if out.len() > super::super::MAX_BYTES {
        return Err(invalid(
            "serialized modern comment extension list is too large",
        ));
    }
    Ok(Some(out))
}
