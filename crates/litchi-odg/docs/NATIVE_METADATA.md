# Headless native ODG metadata receipt

This receipt records one independent LibreOffice headless check of the ODG
transition projection and inert metadata inventory. It is a filter-level
consumer check; it is not GUI acceptance evidence.

The runtime was:

```text
LibreOffice 26.2.5.2 620(Build:2)
/usr/bin/libreoffice
```

The source was the checked-in Draw fixture
[`BlankDrawDocument.odg`](../../../test-data/libreoffice-core/desktop/qa/data/BlankDrawDocument.odg),
SHA-256
`1ce83012be111ae507666a739c8f25640db58fcad0c093a98a070b9ecae5987e`.
The ODG edit owner currently requires compact XML for a changed XML owner, so
the reproducible scratch preparation copied that fixture and applied the
following byte-only formatting transform to every XML member:

```python
import re
from zipfile import ZipFile

with ZipFile("BlankDrawDocument.odg") as source, ZipFile("compact-source.odg", "w") as target:
    for info in source.infolist():
        data = source.read(info.filename)
        if info.filename.endswith(".xml"):
            data = re.sub(rb">[\x09\x0a\x0d ]+<", b"><", data)
            data = re.sub(rb"\?>[\x09\x0a\x0d ]+<", b"?><", data, count=1)
        target.writestr(info, data)
```

The resulting compact scratch input was SHA-256
`6280f3301c3af0470337bd908ef1f97ace1165b6f0e2dda136a4bab447b930da`.
No checked-in fixture was modified.

The changed candidate was generated through the public API using the compact
copy. The complete reproducible generator is
[`native_metadata_receipt.rs`](../examples/native_metadata_receipt.rs); run it
from the workspace with:

```sh
cargo run -p litchi-odg --example native_metadata_receipt -- \
  /tmp/litchi-odg-native-validation-final/compact-source.odg \
  /tmp/litchi-odg-native-validation-final/litchi-changed.odg
```

The semantic core of that example is:

```rust
use litchi_odg::{Drawing, Transition};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let source = Drawing::from_bytes(std::fs::read("compact-source.odg")?)?;
    let mut transition = Transition::new();
    transition.set_transition_type(Some("automatic"))?;
    transition.set_style(Some("dissolve"))?;
    transition.set_speed(Some("medium"))?;
    transition.set_smil_type(Some("fade"))?;
    transition.set_smil_subtype(Some("crossfade"))?;
    transition.set_direction(Some("reverse"))?;
    transition.set_fade_color(Some("#AABBCC"))?;
    transition.set_duration(Some("PT3S"))?;
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition))?;
    let commit = edit.commit()?;
    std::fs::write("litchi-changed.odg", commit.snapshot().as_bytes())?;
    Ok(())
}
```

The candidate was then resaved with the repository helper and an isolated
LibreOffice profile:

```sh
python3 tools/native_odf_resave.py --probe
python3 tools/native_odf_resave.py \
  /tmp/litchi-odg-native-validation-final/litchi-changed.odg \
  /tmp/litchi-odg-native-validation-final/native-final
```

The Litchi candidate was 6,952 bytes with SHA-256
`96e95c8a9b85d9f7f9b75339376098466ec4ab9aee57dd09be26599311ceaa69`.
Reopening it with `litchi_odg::Drawing` returned:

```text
transition_type=automatic
style=dissolve
speed=medium
smil_type=fade
smil_subtype=crossfade
direction=reverse
fade_color=#AABBCC
duration=PT3S
sound=None
active_content=ActiveContentStatus { scripts: 1, events: 0, actions: 0, dde: 0, external_links: 0, embedded_objects: 0 }
```

LibreOffice produced the native output with SHA-256
`c717f1457879f0fb4aa56843a4dda874564dac70aa98bce0e5eb5173d8de29c0`.
The native ZIP contains LibreOffice-generated timestamps, so its whole-file
hash can vary between otherwise equivalent resaves; this is the hash from the
frozen-source run recorded here.
Reopening that output with `litchi_odg::Drawing` reported one page,
`transition=None`, and the same inert active-content inventory:

```text
ActiveContentStatus { scripts: 1, events: 0, actions: 0, dde: 0, external_links: 0, embedded_objects: 0 }
```

Direct XML extraction agrees: the Litchi candidate contains the three
`presentation:transition-*` attributes, four SMIL transition attributes, and
one `presentation:duration`; the LibreOffice `draw8` output contains none of
those transition attributes while retaining the script/inert metadata. This
receipt therefore records a headless Draw-filter limitation: LibreOffice
successfully loads and resaves the ODG and retains the inert inventory, but it
silently drops these transition attributes. It does not claim that native GUI
playback or native transition persistence accepts the Litchi projection.
