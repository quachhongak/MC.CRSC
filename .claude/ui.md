The webpart would display

Section 1: gray background
Row1: search box with placeholder "Staryt typing to search" then clear all button on the right
The search will search for the following fields: Deliverable, prompt title, ai tool, risk level, category

Row2: Label called "Filters" will display, "Deliverables" dropdown (single select) and "Category" dropdown (single select), then on the right side show label"Sorty by" with a dropdown with options:

- Newest (Default)
- Oldest
- A to Z
- Z to A

Section 2: transparent background"
On the left "Show X to y of Y entries" and the right, show pagination controls

Section 3: List of cards
2 per rows
no border radius
border 1px solid blue
deliverable shows first in blue background and dark blue text
then prompt title (in bold)
then "Category: x" where x is the category of the prompt
then "AI Tool: x" where x is the AI tool used for the prompt
then "Risk Level: x" where x is the risk level of the prompt
then "Last modified: x" where x is the last modified date of the prompt. DD-MM-YY
then dark blue button "View full prompt >>"

When clicking on the view full prompt button, it will launch a modal (like the out of the box display view of the sharepoint list entry). It should only show the following:

- Prompt title
- Prompt input
- AI tool
- Category
- Deliverable
- Risk Level
- Last modified
- Source inputs
- Limitations and constraints
- Additional info
