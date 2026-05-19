# CATIA Visual BOM Exporter — Native BOM + Standalone Thumbnails

This branch adds a **safe alternative macro**:

`CATIA_VISUAL_BOM_EXPORTER_NATIVE_STANDALONE.CATScript`

The existing files are not modified on this branch. This was done deliberately so the current Codex/main workflow is not broken.

## Intended workflow

1. Open the main `CATProduct` in CATIA V5.
2. Configure BOM columns in CATIA:
   `Analyze -> Bill of Material -> Define formats`.
3. Run `CATIA_VISUAL_BOM_EXPORTER_NATIVE_STANDALONE.CATScript`.
4. The macro attempts to export the native CATIA BOM, create Excel, then add thumbnails.

## Core rules

- CATIA native Bill of Material is the BOM source.
- The macro does not use hide/show for image capture.
- The macro opens each `CATPart` / `CATProduct` standalone, captures image, creates thumbnail, inserts thumbnail into Excel, then closes the standalone document without saving.
- Standard fasteners remain in Excel but do not get images.
- If a source file cannot be found for a Part Number, the macro stops and saves what exists.

## Output

The macro creates a folder on Desktop:

`<ROOT_PARTNUMBER>_VISUAL_BOM_EXPORT_NATIVE`

with:

- `VISUAL_BOM_EXPORT.xlsx`
- `IMAGES/`
- `THUMBNAILS/`
- `DEBUG_PHASE_LOG.txt`

## Test mode

Default is:

```vb
Const TEST_MODE = True
Const TEST_MAX_ROWS = 20
```

This processes only the first 20 non-fastener rows that require images.

After testing successfully, switch to:

```vb
Const TEST_MODE = False
```

## Notes

This macro cannot be fully validated in GitHub because CATIA V5 Automation must be run on a Windows workstation with CATIA and Excel installed.
