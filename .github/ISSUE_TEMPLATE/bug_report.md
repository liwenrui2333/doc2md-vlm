---
name: Bug report
about: Something broke or produced wrong output
title: ""
labels: bug
assignees: ""
---

## What happened

<!-- What did you run, and what came out? -->

```bash
# the exact command
```

## What you expected

## Environment

<!-- Run this and paste the output; it answers most questions up front. -->

```bash
python -c "import sys, platform; print(sys.version); print(platform.platform())"
python -c "import pymupdf, pptx, requests; print('pymupdf', pymupdf.__version__); print('python-pptx', pptx.__version__); print('requests', requests.__version__)"
```

- OS:
- Python version:
- Vision provider (and model, if used):

## Output

<!--
For a failed conversion, paste the ERROR: line from stderr.
For a bad vision reading, paste the `### Visual / chart reading` block.
Redact anything confidential — this tool touches business documents.
-->

## Anything else

<!-- Does it happen on every file, or one specific file? Does --charts off change it? -->
