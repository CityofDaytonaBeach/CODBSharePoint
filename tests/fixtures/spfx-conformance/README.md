# SPFx Conformance Fixtures

This directory is reserved for normalized Microsoft SPFx reference builds and CODBSharePoint comparison fixtures.

Required fixture set from `start.md`:

1. Vanilla Web Part
2. React Web Part
3. React + SCSS Web Part
4. Microsoft Graph Web Part
5. Multiple Web Parts
6. Application Customizer
7. Field Customizer
8. ListView Command Set
9. Adaptive Card Extension
10. Feature Framework
11. SharePoint List Provisioning
12. Multiple Components + Graph

Each fixture should contain:

- `source/` - source project used for both Microsoft and CODB builds.
- `microsoft/normalized-package.json` - normalized extracted Microsoft `.sppkg` structure.
- `codb/normalized-package.json` - normalized extracted CODB `.sppkg` structure.
- `comparison.json` - semantic comparison result.
- `sharepoint-test.json` - real SharePoint Online deployment result when required.

Do not mark `spfx122`, `sppkg`, or `productionBundling` as supported until the relevant fixture comparisons and SharePoint deployment records pass.
