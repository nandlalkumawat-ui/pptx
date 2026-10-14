FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# officekit is not on PyPI under this name; install the same commit as the local venv.
RUN pip install \
        "officekit @ https://github.com/dream-zjk/officekit/archive/fe397bd28c5711c6372e939e661041c2354b8658.zip" \
        compono==0.5.0 \
        python-docx==1.2.0 \
        ollama==0.6.3 \
        pydantic==2.13.5 \
        python-dotenv==1.2.4 \
        requests==2.34.2 \
        ddgs==9.16.0 \
        python-pptx==1.0.2 \
        pillow==10.4.0 \
        arxiv==4.0.1 \
        pymupdf==1.28.2 \
        PyYAML==6.0.3

# Copy only the application code (no .venv, .env, output or sample files).
COPY main.py config.py ./
COPY generators/ generators/
COPY llm/ llm/
COPY office/ office/
COPY research/ research/
COPY schemas/ schemas/
COPY utils/ utils/

# Generated PPTX/JSON, caches and downloaded images.
VOLUME ["/app/output"]

ENTRYPOINT ["python", "main.py"]
