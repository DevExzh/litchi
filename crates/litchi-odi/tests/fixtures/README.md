# ODI fixture provenance

`odf-1.4-normative-synthetic.fodi` is a hand-authored, deterministic fixture
derived from the ODF 1.4 image-document grammar. It is normative synthetic
evidence only. It was not created or resaved by LibreOffice, Apache OpenOffice,
NeoOffice, or another native producer.

The repository and local fixture roots were searched for `.odi` and `.fodi`
files and contained no producer artifact. An official Apache OpenOffice 4.1.16
Linux distribution was then downloaded from `downloads.apache.org`; its SHA-256
digest matched the published value
`febd01695bbd9ff68d509dbb973bfd714dff0e0a99e50abb4ea32a37eb6aa2ce`.
The shipped filter registry contained no OpenDocument Image (`ODI`) or flat
OpenDocument Image (`FODI`) export filter. Consequently, this corpus still has
no genuine producer ODI/FODI evidence. A future producer fixture must include
the producer name/version and an unchanged original file.
