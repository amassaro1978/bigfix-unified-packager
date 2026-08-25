Unified Packager – Recipe Feature Enhancement

The new Recipe feature allows us to capture application-specific packaging requirements once and reuse them for every future version of that application.

Instead of rebuilding custom installation and removal logic each time an application is updated, a packager can select a centrally maintained recipe. The Unified Packager then automatically applies the approved settings and scripting—including install/uninstall steps, prerequisite actions, process closures, shortcut cleanup, application details, and branding.

How this improves the packaging process:

• Faster recurring updates — proven application logic is reused rather than recreated for each new version.
• Greater consistency — packages for the same application follow the same approved process, regardless of who creates them.
• Fewer errors — recipes are validated before generation and can be previewed before a package is built.
• Reduced reliance on tribal knowledge — specialized packaging knowledge is stored in the shared recipe library instead of remaining with an individual packager.
• Easier maintenance — a process change can be made once in the recipe and used for future packages.
• Improved auditability — every generated package records the recipe name, version, source, and SHA-256 hash, and includes an exact snapshot of the recipe used.

Bottom line: The Recipe feature turns repeat application packaging from a largely manual, one-off effort into a faster, standardized, repeatable, and auditable process—while preserving the existing basic packaging workflow for applications that do not require a recipe.
