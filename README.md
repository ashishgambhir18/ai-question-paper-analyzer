# AI Question Paper Analyzer — V8

## V8 fixes

### Mark reconciliation
The benchmark paper structure is:
- Section A: 40 compulsory
- Section B: six printed 10-mark questions = 60
- Complete printed paper: 100
- Student attempts four Section B questions: actual maximum 80

V8 adds a reconciliation diagnostic so a 101-mark extraction is flagged instead
of silently treated as correct.

### Figure cropping
V8 makes figure cropping more conservative. It prefers actual PDF vector and
embedded image objects, uses only objects inside the AI's candidate region,
and avoids aggressive contour grouping that can merge nearby question text.

For scanned PDFs without separable visual objects, V8 retains the AI candidate
rather than guessing a potentially incorrect crop.

## Secrets
```toml
OPENAI_API_KEY = "your_api_key"
OPENAI_MODEL = "gpt-5.6-luna"
```

Never commit API keys or `.streamlit/secrets.toml`.
