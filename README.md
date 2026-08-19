# vertical-flow

## Summary

Standalone SPFx solution containing the **Vertical Flow** web part — a vertical phase / section / step diagram for visualising a project delivery process.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Solution

| Solution | Author(s) |
| -------- | --------- |
| vertical-flow | Raghav Goswami |

## Version history

| Version | Date        | Comments        |
| ------- | ----------- | --------------- |
| 1.0     | August 2026 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- Set `initialPage` in `config/serve.json` to your own SharePoint site's workbench
- in the command-line run:
  - **npm install**
  - **gulp serve**

## Packaging

- **gulp clean**
- **gulp bundle --ship**
- **gulp package-solution --ship**

The package is written to `solution/Vertical Flow.sppkg`. Client-side assets are
included in the package, so no CDN configuration is required.

## Features

The **Vertical Flow** web part renders a delivery process as a vertical stack of phases. Each phase contains one or more sections, and each section holds a row of steps.

Steps come in four shapes, chosen when the step is added:

| Shape        | Used for                                                             |
| ------------ | -------------------------------------------------------------------- |
| Rectangle    | A standard process step                                              |
| Rounded pill | A client approval                                                    |
| Hexagon      | A SAP status, labelled `SAP Status` above the status code            |
| Diamond      | A gate                                                               |

Steps can also be wrapped in a dashed grouping container with its own title, for steps that apply only in particular cases.

Step reference numbers (`2.1.1`, `2.1.2`, …) are generated automatically from each step's position within its section, so inserting, removing, or reordering steps renumbers the rest. Gate diamonds are not numbered. Steps fill each row before wrapping to the next.

Phases, sections, and steps are authored directly on the page in edit mode: click any title to rename it, use **+ Add step** to insert a shape, and **×** to remove one.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
