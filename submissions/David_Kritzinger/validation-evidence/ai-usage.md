# AI Usage Disclosure

## Tool used

OpenAI ChatGPT was used interactively to assist with the workbook generator
and supporting documentation.

This disclosure was reconstructed after the original conversation was lost,
using the final generator script, validation report and the candidate's
account of the development process. It is not a verbatim conversation log.

## What AI assisted with

- Translating the wedding-planner requirements into a workbook structure
  covering the 21 planning areas.
- Developing and revising the Python/XlsxWriter generator (`build_planner.py`).
- Creating Excel formulas, structured tables, named ranges, dropdowns,
  conditional formatting, navigation and data-driven summaries.
- Designing the clean master and two fictional client demonstration files.
- Refining workbook formatting, dashboard presentation, date displays,
  instructions and print layouts through iterative feedback.
- Troubleshooting issues reported during Excel testing, including an Excel
  repair prompt associated with worksheet view/freeze-pane configuration.
- Preparing validation guidance and drafting `validation.md` from the
  candidate's reported results and supplied screenshots.
- Reconstructing this AI usage disclosure.

Python and XlsxWriter were used to generate the workbook files. They are
build tools, not AI tools, and are not required to use the resulting `.xlsx`
files.

## Candidate involvement and changes

The workbook was developed iteratively rather than submitted as an unchecked
AI output. I ran the generator locally, opened the generated files in Excel,
reviewed the results and requested revisions where functionality,
presentation or usability needed improvement.

I supplied feedback and screenshots from testing, applied the revised
generator, regenerated the workbooks and checked the resulting behaviour.
This included reviewing dashboard layout, whole-number counts, date
formatting, instructions and the correction of the Excel repair prompt.

The implementation was substantially AI-assisted; my contribution included
directing the revisions, executing the build, performing the manual testing
and reviewing the final deliverables.

## How the results were verified

I manually tested the generated workbooks in:

- Microsoft Excel for Microsoft 365 on Windows, 64-bit.
- Version 2610, Build 16.0.20527.20196.

Checks included:

- Budget, payment, outstanding-balance and remaining-budget calculations.
- Table growth and handling of blank optional inputs.
- RSVP changes, seating assignments and capacity warnings.
- Wedding-relative task dates, fixed-date overrides and calendar counts.
- Food-and-drink quantity × unit-cost calculations and totals.
- Saving, closing and reopening a modified workbook.
- Independence of the two client files and checking for external workbook links.
- Clean-master usability and budget/itinerary print or PDF output.
- Reopening and retesting after the repair-prompt correction.

Detailed inputs, expected results, actual results and limitations are recorded
in `validation.md`.

XlsxWriter writes formulas but does not calculate them. Verification therefore
relied on actual recalculation and interaction in Excel, not solely on AI
review or inspection of formula strings.

ChatGPT did not independently operate my local Excel installation. Statements
about manual test outcomes are based on my checks and the evidence I supplied.

## Boundaries and limitations

- AI assistance does not establish compatibility with untested Excel versions
  or platforms.
- Testing covered representative workflows, not every possible input.
- Fictional demonstration data was used; no real customer or client records
  were required.
- AI assistance does not confer ownership of, or commercial rights to,
  Johan Botha's supplied materials. The adaptation remains subject to the
  assessment licence.
- I remain responsible for reviewing the submission, understanding its
  implementation and accurately describing its capabilities and limitations.