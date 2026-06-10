# doc2md-vlm

Convert PDF/PPTX files into Markdown plus visual evidence for charts.

## Usage

```powershell
python doc2md.py input.pdf --charts
python doc2md.py input.pptx --charts
python doc2md.py input.pdf --charts --local-llm --vlm-model gemma4-vision-thinking:12b
python doc2md.py E:\docs --batch --charts
```

## Output

```text
input_doc2md/
  input.md
  assets/
  charts/
  metadata.json
```

## Notes

- PDF pages are rendered to images when `--charts` is enabled, so visual charts are not lost.
- PPTX native chart objects are exported to CSV when readable.
- Local vision models are optional. The default endpoint is Ollama:
  `http://127.0.0.1:11434/api/chat`.
