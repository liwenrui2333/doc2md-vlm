# Contributing

Thanks for taking a look. This is a small tool with a narrow job, so the bar is simple: a change should make conversion more accurate or more predictable without changing the output format for existing users.

## Setup

```bash
git clone https://github.com/liwenrui2333/doc2md-vlm.git
cd doc2md-vlm
python -m venv .venv
```

Activate it, then:

```bash
python -m pip install -e ".[dev]"
python -m pytest -q
```

Windows PowerShell activation is `.venv\Scripts\Activate.ps1`; macOS and Linux use `source .venv/bin/activate`.

## Tests

The suite builds its own PDF and PPTX fixtures at runtime — there are no binary files in the repo. To inspect what the tests feed the converter:

```bash
python tests/fixtures.py out/     # writes out/report.pdf and out/deck.pptx
```

Write tests against behaviour, not exact glyphs. Chinese text spacing differs between PyMuPDF versions, so assert that a string is present rather than that a page matches byte for byte.

## Platform notes

Slide rendering for PPTX uses PowerPoint COM automation and only works on Windows with PowerPoint installed. Everywhere else, the tool exports text and chart CSVs and records a warning in `metadata.json`. If you touch that code, make sure the non-Windows path still passes — CI runs Ubuntu, macOS, and Windows.

## Output format is a contract

The generated Markdown, the chart CSVs, and `metadata.json` are consumed by downstream scripts. Adding a new key is fine. Renaming or removing one, or changing the folder layout, needs a note in the PR description.

## Commits and PRs

Keep commits focused; one concern per commit. In the PR description, say what you actually ran — which real files you converted and what changed in the output. "Tests pass" alone does not cover a change that alters conversion results.

## Reporting bugs

Use the bug template. The environment block it asks for answers most of the questions I would otherwise have to ask.

## License

By contributing you agree your work is released under the MIT license, matching the rest of the project.
