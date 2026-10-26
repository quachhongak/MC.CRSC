# CRSC Prompts

CRSC Prompts is a SharePoint Framework web part for browsing approved prompts in a SharePoint list. Visitors can search and filter prompt cards, sort results, page through the list, and open a prompt's full details without leaving the page.

## Architecture

- `CrscPromptsWebPart` initializes PnPjs against the current SharePoint site and exposes the list and page-size settings.
- `SharePointService` reads visible lists and pages through approved prompt items using PnPjs v4 and SharePoint REST. `ISharePointService` keeps the data boundary mockable; `MockSharePointService` supplies local test data.
- The React component handles the interface and loading, empty, retry, and unconfigured states. Pure search, sorting, and date-formatting functions live in `utils/promptUtils.ts`.
- All data is read from the current site; no Microsoft Graph or additional API permissions are required.

## Web part properties

| Property | Type | Default | Effect |
| --- | --- | --- | --- |
| Prompt list | SharePoint list picker (property-pane dropdown) | Not selected | Selects the visible list to query. Only items with `Status` equal to `Approved` are displayed. |
| Items per page | Choice: 5, 10, 50, 100 | 50 | Sets the number of matching cards shown on each page. |
| Modal minimum width | Positive number (pixels) | 760 | Sets the minimum width of the full-prompt dialog. |
| Clear all button color and border | Hex color | `#d0d0d0` | Sets both the Clear all button background and border color. |
| Card border and deliverable background color | Hex color | `#0078d4` | Sets both the prompt card border and deliverable tag background color. |
| Card title color | Hex color | `#323130` | Sets prompt title text color. |
| View full prompt text color | Hex color | `#ffffff` | Sets the View full prompt button text color. |
| View full prompt background color | Hex color | `#004578` | Sets the View full prompt button background and border. |

The selected list must have these internal field names: `Title`, `Input`, `AITool`, `Category`, `Deliverable`, `RiskLevel`, `SourceInputsRequired`, `LimitationsandConstraints`, `AdditionalInformation`, `Status`, and the standard `Modified` field. See [.claude/schema.md](./.claude/schema.md) for the display-name mapping.

## Prerequisites

- Node.js 18.17.1 through 18.x (SPFx 1.18.2 toolchain compatibility)
- npm
- SharePoint Online site with the configured prompt list and read access for page visitors
- SharePoint Framework 1.18.2

## Local development

```powershell
npm install
npm test
npx gulp serve
```

Use the SharePoint workbench on a site where the prompt list exists. Add the web part, open its properties, select the list, and choose the page size. The browser session's SharePoint permissions are used to read the selected list.

## Testing

`npm test` runs the SPFx test task and Jest unit tests. Unit tests cover search fields, filtering, all sort orders, date formatting, and the mock data service; tests do not call SharePoint.

## Quality decisions

- Fluent UI React 8 is used for accessible form controls, messages, buttons, and the modal.
- English localization is provided through the SPFx `loc` resources. Additional languages, client-side caching, and Application Insights telemetry were excluded as requested.
- The user-facing component has explicit unconfigured, loading, error/retry, and no-results states, plus an error boundary.
- No API permission request is needed because data access is limited to the current site.
