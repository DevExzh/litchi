# Retained native read fixture

`tdf169496_hidden_graphic.xlsx` is an unchanged copy of the local LibreOffice
test-corpus snapshot at
`3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx`.
The upstream revision of that snapshot was not retained; the exact input
identity is its 12,470-byte archive and SHA-256:

```text
0b647da300a085f39914fdfae961463ae9e54ffe772b2e0eb9860a841ab93f72
```

The archive is used only for native read/capture topology checks. It does not
establish acceptance of generated output by LibreOffice or Microsoft Office.
The adapter checks the hash before capture, and the source manifest binds the
fixture to the current committed harness snapshot.

The local source tree's unchanged `COPYING`, `COPYING.LGPL`, and `COPYING.MPL`
notices accompany it in `libreoffice-license/`.

The clean `4b81e22f9` checkout exposed that the original `3rdparty` path was
absent from Git. That run's missing-input refusal was correct; a temporary
external fixture symlink supported correctness checks only and was removed.
Retaining the archive here removes that dependency for subsequent sealed runs.
