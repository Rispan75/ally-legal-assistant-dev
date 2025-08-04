import os
import uuid
import json
import re
import langcodes
from docx import Document
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
from io import BytesIO
from langdetect import detect
from azure.core.credentials import AzureKeyCredential
from azure.search.documents.indexes import SearchIndexClient
from azure.search.documents import SearchClient
from azure.search.documents.models import QueryType
from azure.search.documents.indexes.models import (
    SearchIndex, SimpleField, SearchableField, SearchField,
    SearchFieldDataType, VectorSearch, HnswAlgorithmConfiguration,
    HnswParameters, VectorSearchProfile, VectorSearchAlgorithmKind,
    VectorSearchAlgorithmMetric
)
from openai import AzureOpenAI


AZURE_SEARCH_ENDPOINT = os.getenv("AZURE_SEARCH_ENDPOINT")
AZURE_SEARCH_KEY = os.getenv("AZURE_SEARCH_KEY")
AZURE_OPENAI_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_ENDPOINT = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")
AZURE_EMBEDDING_DEPLOYMENT = os.getenv("AZURE_EMBEDDING_DEPLOYMENT")
INDEX_NAME = os.getenv("AZURE_SEARCH_INDEX")

SYSTEM_PROMPT = """
You are a legal policy extraction assistant. Your job is to extract structured information from legal clauses.
Important rules:
- Do not translate or paraphrase.
- Keep the language the same as input.
- Keep all values precise, short and enforceable.
- The summary must be 6-7 words max, capturing the essence of the clause.
Return valid JSON in this format:
{
  "title": "Title of the clause",
  "instruction": "Enforceable instruction",
  "summary": "6-7 word summary of clause",
  "tags": ["Tag1", "Tag2"],
  "severity": 2
}
"""

def detect_language(text):
    try:
        code = detect(text)
        return langcodes.get(code).language_name().title()
    except:
        return "Unknown"

def extract_text_from_docx_bytes(docx_bytes):
    doc = Document(BytesIO(docx_bytes))
    return "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())

def extract_text_from_pdf_bytes(pdf_bytes):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = "".join(page.get_text() for page in doc)
        return text.strip() or ocr_pdf_bytes(pdf_bytes)
    except:
        return ocr_pdf_bytes(pdf_bytes)

def ocr_pdf_bytes(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texts = []
    for page in doc:
        pix = page.get_pixmap(dpi=300)
        img = Image.open(BytesIO(pix.tobytes()))
        texts.append(pytesseract.image_to_string(img))
    return "\n".join(texts).strip()

def chunk_text_legal_policy(text):
    lines = text.splitlines()
    chunks = []
    current_chunk = []

    def is_heading(line):
        return bool(re.match(r"^[A-Z][A-Za-z\s\-]*:$", line.strip())) or \
               bool(re.match(r"^[A-Z][A-Za-z\s\-]*$", line.strip()))

    def is_definition_clause(line):
        return bool(re.match(r"^[A-Z][a-zA-Z\s\-]+:\s+", line.strip()))

    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        if is_heading(line):
            if current_chunk:
                chunks.append(" ".join(current_chunk).strip())
                current_chunk = []
            current_chunk.append(line)
        elif is_definition_clause(line):
            if current_chunk:
                chunks.append(" ".join(current_chunk).strip())
                current_chunk = []
            current_chunk.append(line)
        else:
            current_chunk.append(line)

    if current_chunk:
        chunks.append(" ".join(current_chunk).strip())

    return chunks

def analyze_text_with_openai(text):
    client = AzureOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        api_version="2025-01-01-preview",
        azure_endpoint=AZURE_OPENAI_ENDPOINT
    )
    resp = client.chat.completions.create(
        model=AZURE_OPENAI_DEPLOYMENT,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": text}
        ],
        temperature=0.2
    )
    content = resp.choices[0].message.content
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", content, re.DOTALL)
        if match:
            return json.loads(match.group())
        raise ValueError(f"Invalid JSON from OpenAI: {content}")

def get_embedding(text):
    client = AzureOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        api_version="2024-12-01-preview",
        azure_endpoint=AZURE_OPENAI_ENDPOINT
    )
    resp = client.embeddings.create(model=AZURE_EMBEDDING_DEPLOYMENT, input=[text])
    return resp.data[0].embedding

def is_file_indexed(filename):
    client = SearchClient(
        endpoint=AZURE_SEARCH_ENDPOINT,
        index_name=INDEX_NAME,
        credential=AzureKeyCredential(AZURE_SEARCH_KEY)
    )
    try:
        results = client.search("*", filter=f"filename eq '{filename}'", top=1)
        return any(True for _ in results)
    except:
        return False

def clause_exists(instruction_text):
    client = SearchClient(
        endpoint=AZURE_SEARCH_ENDPOINT,
        index_name=INDEX_NAME,
        credential=AzureKeyCredential(AZURE_SEARCH_KEY)
    )
    try:
        results = client.search(
            search_text=f'"{instruction_text}"',
            search_fields=["instruction"],
            top=1,
            query_type=QueryType.SIMPLE
        )
        return any(True for _ in results)
    except Exception as e:
        print(f"Failed duplicate check: {e}")
        return False

def upload_to_search(record):
    client = SearchClient(
        endpoint=AZURE_SEARCH_ENDPOINT,
        index_name=INDEX_NAME,
        credential=AzureKeyCredential(AZURE_SEARCH_KEY)
    )
    try:
        client.upload_documents(documents=[record])
        print(f"Uploaded: {record['title']} from file: {record['filename']}")
    except Exception as e:
        print(f"Upload failed for '{record['title']}' in file '{record['filename']}': {e}")

def create_index_if_not_exists():
    client = SearchIndexClient(
        endpoint=AZURE_SEARCH_ENDPOINT,
        credential=AzureKeyCredential(AZURE_SEARCH_KEY)
    )
    if INDEX_NAME in [i.name for i in client.list_indexes()]:
        return

    fields = [
        SimpleField(name="id", type=SearchFieldDataType.String, key=True),
        SimpleField(name="PolicyId", type=SearchFieldDataType.String, filterable=True),
        SimpleField(name="filename", type=SearchFieldDataType.String, filterable=True),
        SearchableField(name="title", type=SearchFieldDataType.String),
        SearchableField(name="instruction", type=SearchFieldDataType.String),
        SearchableField(name="summary", type=SearchFieldDataType.String),
        SearchField(name="embedding", type=SearchFieldDataType.Collection(SearchFieldDataType.Single),
                    searchable=True, vector_search_dimensions=1536, vector_search_profile_name="myHnswProfile"),
        SearchField(name="tags", type=SearchFieldDataType.Collection(SearchFieldDataType.String),
                    filterable=True, facetable=True),
        SimpleField(name="locked", type=SearchFieldDataType.Boolean, filterable=True),
        SearchField(name="groups", type=SearchFieldDataType.Collection(SearchFieldDataType.String),
                    filterable=True),
        SimpleField(name="severity", type=SearchFieldDataType.Int32, filterable=True),
        SimpleField(name="language", type=SearchFieldDataType.String, filterable=True),
        SearchableField(name="original_text", type=SearchFieldDataType.String)
    ]

    vector_search = VectorSearch(
        algorithms=[HnswAlgorithmConfiguration(
            name="myHnsw",
            kind=VectorSearchAlgorithmKind.HNSW,
            parameters=HnswParameters(m=5, ef_construction=300, ef_search=400,
                                      metric=VectorSearchAlgorithmMetric.COSINE)
        )],
        profiles=[VectorSearchProfile(name="myHnswProfile", algorithm_configuration_name="myHnsw")]
    )

    index = SearchIndex(name=INDEX_NAME, fields=fields, vector_search=vector_search)
    client.create_index(index)

def process_blob(blob_name, blob_bytes):
    create_index_if_not_exists()

    cleaned_blob_name = blob_name.strip().lower()
    if not cleaned_blob_name.endswith((".docx", ".pdf")):
        print(f"Unsupported file: {blob_name}")
        return

    if is_file_indexed(blob_name):
        print(f"Skipping file {blob_name} - already indexed")
        return

    base_name = os.path.splitext(os.path.basename(cleaned_blob_name))[0]

    try:
        full_text = extract_text_from_docx_bytes(blob_bytes) if cleaned_blob_name.endswith(".docx") else extract_text_from_pdf_bytes(blob_bytes)
        if not full_text:
            print(f"No text found in {blob_name}")
            return

        detected_language = detect_language(full_text)
        clauses = chunk_text_legal_policy(full_text)
        print(f"Extracted {len(clauses)} clauses from {blob_name}")

        for idx, clause in enumerate(clauses):
            try:
                structured = analyze_text_with_openai(clause)

                title = structured.get("title", "Untitled")
                instruction = structured.get("instruction", "").strip()
                summary = structured.get("summary", "").strip()
                tags = structured.get("tags", [])
                severity = structured.get("severity", 2)

                if not instruction:
                    print(f"Empty instruction in clause {idx}, skipping")
                    continue

                if severity not in (1, 2):
                    severity = 2

                if clause_exists(instruction):
                    print(f"Skipping duplicate clause in {blob_name}: {title}")
                    continue

                embedding = get_embedding(instruction)
                if len(embedding) != 1536:
                    print(f"Invalid embedding length for clause in {blob_name}")
                    continue

                record = {
                    "id": str(uuid.uuid4()),
                    "PolicyId": f"{base_name}-{idx}-{uuid.uuid4()}",
                    "filename":  os.path.basename(blob_name),
                    "title": title,
                    "instruction": instruction,
                    "summary": summary,
                    "embedding": embedding,
                    "tags": tags,
                    "locked": True,
                    "groups": [],
                    "severity": severity,
                    "language": detected_language,
                    "original_text": clause
                }

                upload_to_search(record)

            except Exception as e:
                print(f"Error processing clause {idx} in {blob_name}: {e}")

    except Exception as e:
        print(f"Failed to process file {blob_name}: {e}")




