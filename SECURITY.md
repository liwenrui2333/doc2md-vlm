# Data and privacy

Conversion is local by default. Running without `--local-llm` never opens a network connection: PDF text and page images are produced by PyMuPDF in-process, and PPTX text and charts by python-pptx and PowerPoint's own export. Nothing leaves the machine.

**A vision model changes that.** `--local-llm` sends each rendered page as an image to whichever endpoint you configured. Choose it with your data in mind:

| Provider | Where page images go |
|---|---|
| `ollama` on `127.0.0.1` | Stays on this machine |
| `ollama` on another host | That host |
| `openai` at `api.openai.com` | OpenAI, subject to their data policies |
| `openai` at a self-hosted endpoint | Wherever that server is |

There is no telemetry, no analytics, and no code that uploads anything except the page images described above. API keys are read from `DOC2MD_VLM_API_KEY` or `--vlm-api-key` and are sent only in the `Authorization` header of the request to the endpoint you named; they are never written to `metadata.json` or the Markdown.

`metadata.json` records the absolute path of the source file, so review it before sharing a converted folder.
