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

# Azure config
AZURE_SEARCH_ENDPOINT = ""
AZURE_SEARCH_KEY = ""
INDEX_NAME = "legal-instructions"

AZURE_OPENAI_API_KEY = ""
AZURE_OPENAI_ENDPOINT = ""
AZURE_OPENAI_DEPLOYMENT = "gpt-4o"
AZURE_EMBEDDING_DEPLOYMENT = "text-embedding-ada-002"

# Critical: System prompt explicitly forbids translation and forces original language retention
SYSTEM_PROMPT = """
You are a legal policy extraction assistant. Your job is to extract structured information from policy or compliance text.

IMPORTANT:
- The input text is always in its ORIGINAL LANGUAGE (e.g., German, French, English).
- You MUST NOT translate, paraphrase, or change the language of any part of the text.
- Your output fields (title, instruction, tags) MUST BE IN THE SAME LANGUAGE as the input text.
- Keep the text precise and clear, but preserve the original wording and language.
- Return valid JSON with these exact fields:

{
  "title": "A concise, specific title of the policy clause in the original language",
  "instruction": "A clear and enforceable instruction extracted from the input text, in the original language",
  "tags": ["RelevantTag1", "RelevantTag2"], 
  "severity": 2
}
"""

# --- Helpers ---

def detect_language(text):
    try:
        code = detect(text)
        return langcodes.get(code).language_name().title()
    except:
        return "Unknown"

def extract_text_from_docx(path):
    doc = Document(path)
    return "\n".join(p.text.strip() for p in doc.paragraphs if p.text.strip())

def extract_text_from_pdf(path):
    try:
        doc = fitz.open(path)
        text = "".join(page.get_text() for page in doc)
        return text.strip() or ocr_pdf(path)
    except:
        return ocr_pdf(path)

def ocr_pdf(path):
    doc = fitz.open(path)
    texts = []
    for page in doc:
        pix = page.get_pixmap(dpi=300)
        img = Image.open(BytesIO(pix.tobytes()))
        texts.append(pytesseract.image_to_string(img))
    return "\n".join(texts).strip()

def smart_chunk_with_openai(text):
    prompt = """
You are a document processor. Given a long policy text in its original language, split it into coherent policy clauses that can stand alone.

IMPORTANT:
- Keep all output in the ORIGINAL LANGUAGE of the input text.
- Return only a JSON array of clauses.
- Do not translate or change the language.

Example output:
[
  "Clause 1: Mitarbeiter müssen die Datenschutzrichtlinien einhalten.",
  "Clause 2: Alle Finanzdaten müssen verschlüsselt gespeichert werden."
]
"""
    client = AzureOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        api_version="2023-05-15",
        azure_endpoint=AZURE_OPENAI_ENDPOINT
    )
    resp = client.chat.completions.create(
        model=AZURE_OPENAI_DEPLOYMENT,
        messages=[
            {"role": "system", "content": prompt},
            {"role": "user", "content": text}
        ],
        temperature=0.2
    )
    raw = resp.choices[0].message.content
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        match = re.search(r"\[.*\]", raw, re.DOTALL)
        if match:
            return json.loads(match.group())
        raise ValueError(f"Could not parse smart chunks: {raw}")

def analyze_text_with_openai(text):
    client = AzureOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        api_version="2023-05-15",
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
    if clause_exists(record["instruction"]):
        print(f"Skipped duplicate clause: {record['title']}")
        return

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
        print("Search index already exists.")
        return

    fields = [
        SimpleField(name="id", type=SearchFieldDataType.String, key=True, sortable=True),
        SimpleField(name="PolicyId", type=SearchFieldDataType.String, filterable=True),
        SimpleField(name="filename", type=SearchFieldDataType.String, filterable=True),
        SearchableField(name="title", type=SearchFieldDataType.String),
        SearchableField(name="instruction", type=SearchFieldDataType.String),
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
    print("Created Azure Cognitive Search index.")

def process_directory(folder_path):
    create_index_if_not_exists()
    for filename in os.listdir(folder_path):
        if filename.startswith("~$") or not filename.lower().endswith((".docx", ".pdf")):
            continue

        print(f"Processing file {filename}")
        path = os.path.join(folder_path, filename)
        base_name = os.path.splitext(filename)[0]
        policy_id = f"{base_name}-{uuid.uuid4()}"

        try:
            full_text = extract_text_from_docx(path) if filename.endswith(".docx") else extract_text_from_pdf(path)
            if not full_text:
                print(f"No text found in {filename}")
                continue

            detected_language = detect_language(full_text)
            
            # Smart chunking - keep original language (with explicit instructions)
            clauses = smart_chunk_with_openai(full_text)
            if not clauses or not isinstance(clauses, list):
                print(f"Chunking failed or no clauses in {filename}, fallback splitting.")
                clauses = re.split(r'\n(?=\d{1,2}\.|\w\.)', full_text)
                clauses = [c.strip() for c in clauses if len(c.strip()) > 30]

            for clause in clauses:
                try:
                    structured = analyze_text_with_openai(clause)

                    title = structured.get("title", "No Title")
                    instruction = structured.get("instruction", "").strip()
                    tags = structured.get("tags", [])
                    severity = structured.get("severity", 2)
                    if not instruction:
                        print(f"Empty instruction skipped in {filename}")
                        continue
                    if severity not in (1, 2):
                        severity = 2

                    embedding = get_embedding(instruction)
                    if len(embedding) != 1536:
                        print(f"Invalid embedding length for clause in {filename}")
                        continue

                    record = {
                        "id": str(uuid.uuid4()),
                        "PolicyId": policy_id,
                        "filename": filename,
                        "title": title,
                        "instruction": instruction,
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
                    print(f"Error processing clause in {filename}: {e}")

        except Exception as e:
            print(f"Failed to process file {filename}: {e}")

# Run
if __name__ == "__main__":
    directory_path = "/policy_documents"  # Update your folder path here
    process_directory(directory_path)
