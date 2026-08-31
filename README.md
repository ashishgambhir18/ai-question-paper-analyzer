# AI Question Paper Analyzer — V7

V7 is designed for the benchmark Physics paper:
- Complete printed paper: 100 marks
- Actual examination maximum: 80 marks
- Chapter weightage pie chart: based on 100 printed marks

V7 also improves figure handling. The AI identifies the page and approximate
visual location; PDF vector/image objects are preferred for the actual crop,
and a scan fallback is used when the PDF is image-based. Nearby question prose
is deliberately excluded from the visual-object path.

## Deploy
Use `app.py` as the Streamlit entrypoint and add the API key in Streamlit Secrets:

```toml
OPENAI_API_KEY = "your_api_key"
OPENAI_MODEL = "gpt-5.6-luna"
```

Never commit API keys or `.streamlit/secrets.toml`.
