import os, re, uuid, json, tempfile
from io import BytesIO
from datetime import datetime, timezone
from langdetect import detect
import fitz, pytesseract
from PIL import Image
from docx import Document

from openai import AzureOpenAI

from azure.core.credentials import AzureKeyCredential

from azure.search.documents import SearchClient
from azure.search.documents.indexes import SearchIndexClient

from azure.storage.blob import BlobServiceClient
from azure.core.credentials import AzureNamedKeyCredential

from azure.search.documents.indexes.models import (
    SearchIndex, SimpleField, SearchableField, SearchField, SearchFieldDataType,
    VectorSearch, HnswAlgorithmConfiguration, HnswParameters,
    VectorSearchProfile, VectorSearchAlgorithmKind, VectorSearchAlgorithmMetric
)

# --- CONFIGURATION ---
AZURE_SEARCH_ENDPOINT = "Write your AZURE_SEARCH_ENDPOINT here"
AZURE_SEARCH_KEY = "Write your AZURE_SEARCH_KEY here"
DOC_INDEX = "legal-documents"
POLICY_INDEX = "legal-instructions"
AZURE_OPENAI_API_KEY = "Write your AZURE_OPENAI_API_KEY here"
AZURE_OPENAI_ENDPOINT = "Write your AZURE_OPENAI_ENDPOINT here"
AZURE_OPENAI_DEPLOYMENT = "gpt-4o"
AZURE_EMBEDDING_DEPLOYMENT = "text-embedding-ada-002"

LANG_CODE_TO_NAME = {
    "en": "English", "de": "German", "fr": "French", "es": "Spanish",
    "it": "Italian", "pt": "Portuguese", "nl": "Dutch", "ru": "Russian",
    "zh-cn": "Chinese", "ja": "Japanese"
}

# --- PROMPTS ---
SPLIT_PROMPT = """
You are a smart legal document splitter. Your task is to divide a long legal document into individual paragraphs that are semantically meaningful. If there are sections or subsections, then divide them further. Each paragraph should be a self-contained paragraph or sentence that expresses a specific concept or clause. Preserve the original language of the input text (e.g., if the document is in German, French, or English, keep all output in that same language). Do not translate.

The document should be split as follows:
- Respect the structure of the document (e.g., titles, preambles, numbered sections, subsections like 1, 1.1, 2.2, etc.).
- If a section contains subpoints or bullet points, treat each as a separate paragraph if they represent different ideas.
- Do not merge unrelated ideas into one chunk; each chunk must cover only one key legal or conceptual idea.
- Retain headers and numbering (e.g., "2.1. Confidentiality Obligations") as part of the "text" field for clarity.

For each paragraph, generate:
- "id" (starting from 1)
- "title": You are a document summarizer. Provide a concise and descriptive title for paragraph.
- "text": Include the exact text of the paragraph from the original document.
- "keyphrases": List of 3 to 5 keyphrases that represent the core legal or conceptual ideas.
- "summary": A 2 to 4 sentence summary of the paragraph in plain language, in the same language as the input, explaining its meaning or intent.

Return a JSON array like:
{{"id": 1, "title": "...", "text": "...", "keyphrases": ["..."], "summary": "..."}}

Input document:
{text}
"""

COMPLIANCE_PROMPT = """
You are a legal compliance classifier. Classify the clause with respect to the policy instruction.

Use these categories:
- "compliant": The clause follows the policy or expected legal standard.
- "non-compliant": The clause violates, weakens, or contradicts the policy.
- "irrelevant": The clause has no relation to the policy.

Examples:
Compliant:
  • "Neither party shall be liable for failure to perform due to acts of God, war, or other natural disasters."
  • "Payment shall be made within 30 days of invoice receipt."
Non-Compliant:
  • "Seller is not liable for any damages whatsoever, regardless of the cause."
  • "Company A may terminate the contract at any time for any reason."
Irrelevant:
  • "Employee A enjoys gardening in their free time."
  • "Company C will provide catering services for the event." 

Output exactly one JSON object:
{{"category": "compliant"}}
{{"category": "non-compliant"}}
{{"category": "irrelevant"}}

Clause:
\"\"\"{clause}\"\"\"

Policy:
\"\"\"{policy}\"\"\"
"""

# --- HELPERS ---
def call_openai(prompt_text):
    client = AzureOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        api_version="2023-05-15"
    )
    resp = client.chat.completions.create(
        model=AZURE_OPENAI_DEPLOYMENT,
        messages=[{"role": "user", "content": prompt_text}],
        temperature=0
    )
    return resp.choices[0].message.content.strip()

def get_embedding(text):
    client = AzureOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        api_version="2023-05-15"
    )
    resp = client.embeddings.create(model=AZURE_EMBEDDING_DEPLOYMENT, input=[text])
    return resp.data[0].embedding

def extract_docx_text(path):
    return "\n".join(p.text.strip() for p in Document(path).paragraphs if p.text.strip())

def extract_pdf_text(path):
    doc = fitz.open(path)
    text = "".join(page.get_text() for page in doc).strip()
    return text or ocr_pdf_text(path)

def ocr_pdf_text(path):
    parts = []
    for page in fitz.open(path):
        pix = page.get_pixmap(dpi=300)
        img = Image.open(BytesIO(pix.tobytes()))
        parts.append(pytesseract.image_to_string(img))
    return "\n".join(parts).strip()

def detect_language(text):
    try:
        return LANG_CODE_TO_NAME.get(detect(text), detect(text))
    except:
        return "Unknown"

def smart_split_document(text):
    prompt = SPLIT_PROMPT.format(text=text[:7000])
    try:
        resp = call_openai(prompt)
        m = re.search(r"\[.*\]", resp, re.S)
        return json.loads(m.group()) if m else []
    except:
        return [{"id": i+1, "title": "", "text": p.strip(), "keyphrases": [], "summary": ""} for i, p in enumerate(text.split("\n\n")) if p.strip()]

def load_policies():
    client = SearchClient(endpoint=AZURE_SEARCH_ENDPOINT, index_name=POLICY_INDEX, credential=AzureKeyCredential(AZURE_SEARCH_KEY))
    return [{"id": doc["PolicyId"], "instruction": doc["instruction"], "language": doc.get("language", "Unknown")} for doc in client.search("*", top=1000)]

def check_compliance_with_gpt(clause, policies):
    compliant, nonc, irr = [], [], []
    for p in policies:
        prompt = COMPLIANCE_PROMPT.format(clause=clause, policy=p["instruction"])
        try:
            resp = call_openai(prompt)
            match = re.search(r"\{.*\}", resp)
            category = json.loads(match.group()).get("category", "").lower() if match else resp.strip().lower()
            if category == "compliant":
                compliant.append(p["id"])
            elif category == "non-compliant":
                nonc.append(p["id"])
            elif category == "irrelevant":
                irr.append(p["id"])
        except Exception as e:
            print(f"[WARN] Compliance check failed for policy {p['id']}: {e}")
    return compliant, nonc, irr

def upload_chunk(fname, pid, title, body, lang, emb, comp_ids, nonc_ids, irr_ids, keyphrases, summary):
    client = SearchClient(endpoint=AZURE_SEARCH_ENDPOINT, index_name=DOC_INDEX, credential=AzureKeyCredential(AZURE_SEARCH_KEY))
    is_compliant = (not comp_ids and not nonc_ids) or (bool(comp_ids) and not bool(nonc_ids))
    doc = {
        "id": str(uuid.uuid4()),
        "filename": fname,
        "ParagraphId": pid,
        "title": title,
        "paragraph": body,
        "embedding": emb,
        "language": lang,
        "isCompliant": is_compliant,
        "CompliantCollection": comp_ids,
        "NonCompliantCollection": nonc_ids,
        "IrrelevantCollection": irr_ids,
        "group": [],
        "keyphrases": keyphrases,
        "summary": summary,
        "department": "",
        "date": datetime.utcnow().replace(tzinfo=timezone.utc).isoformat()
    }
    client.upload_documents(documents=[doc])
    print(f"[Uploaded] {fname} clause {pid}, compliant={doc['isCompliant']}")

def create_index_if_not_exists():
    client = SearchIndexClient(endpoint=AZURE_SEARCH_ENDPOINT, credential=AzureKeyCredential(AZURE_SEARCH_KEY))
    if DOC_INDEX in [idx.name for idx in client.list_indexes()]:
        print(f"[INFO] Index '{DOC_INDEX}' already exists.")
        return

    fields = [
        SimpleField(name="id", type=SearchFieldDataType.String, key=True),
        SimpleField(name="ParagraphId", type=SearchFieldDataType.Int32, filterable=True, sortable=True),
        SearchableField(name="title", type=SearchFieldDataType.String, filterable=True),
        SearchableField(name="paragraph", type=SearchFieldDataType.String),
        SearchField(name="embedding", type=SearchFieldDataType.Collection(SearchFieldDataType.Single),
                    searchable=True, vector_search_dimensions=1536, vector_search_profile_name="vsProfile"),
        SimpleField(name="filename", type=SearchFieldDataType.String, filterable=True),
        SimpleField(name="language", type=SearchFieldDataType.String, filterable=True),
        SimpleField(name="isCompliant", type=SearchFieldDataType.Boolean, filterable=True),
        SearchField(name="CompliantCollection", type=SearchFieldDataType.Collection(SearchFieldDataType.String)),
        SearchField(name="NonCompliantCollection", type=SearchFieldDataType.Collection(SearchFieldDataType.String)),
        SearchField(name="IrrelevantCollection", type=SearchFieldDataType.Collection(SearchFieldDataType.String)),
        SearchField(name="group", type=SearchFieldDataType.Collection(SearchFieldDataType.String)),
        SearchField(name="keyphrases", type=SearchFieldDataType.Collection(SearchFieldDataType.String)),
        SearchableField(name="summary", type=SearchFieldDataType.String),
        SimpleField(name="department", type=SearchFieldDataType.String, filterable=True),
        SimpleField(name="date", type=SearchFieldDataType.String, filterable=True, sortable=True),
    ]

    vector_config = VectorSearch(
        algorithms=[HnswAlgorithmConfiguration(
            name="vsAlgo",
            kind=VectorSearchAlgorithmKind.HNSW,
            parameters=HnswParameters(metric=VectorSearchAlgorithmMetric.COSINE, m=4, ef_construction=200, ef_search=100)
        )],
        profiles=[VectorSearchProfile(name="vsProfile", algorithm_configuration_name="vsAlgo")]
    )

    index = SearchIndex(name=DOC_INDEX, fields=fields, vector_search=vector_config)
    client.create_index(index)
    print(f"[INFO] Index '{DOC_INDEX}' created.")


def process_document(path, filenamee):
    fname = filenamee
    txt = extract_docx_text(path) if fname.lower().endswith(".docx") else extract_pdf_text(path) 
    lang = detect_language(txt)
    print(f"[INFO] Detected language: {lang}")
    policies = [p for p in load_policies() if p["language"].lower() == lang.lower()]
    clauses = smart_split_document(txt)
    print(f"[INFO] Extracted {len(clauses)} clauses")

    for cl in clauses:
        body = cl["text"].strip()
        if len(body) < 30:
            continue
        title = cl.get("title") or "Untitled"
        keyphrases = cl.get("keyphrases", [])
        summary = cl.get("summary", "")
        emb = get_embedding(body)
        comp_ids, nonc_ids, irr_ids = check_compliance_with_gpt(body, policies)
        upload_chunk(fname, cl["id"], title, body, lang, emb, comp_ids, nonc_ids, irr_ids, keyphrases, summary)

def process_documents_from_blob():
    create_index_if_not_exists()
    
    try:
        # List blobs in container
        # Azure Storage configuration
        ACCOUNT_NAME = "dekradocuments"
        ACCOUNT_KEY = "Write your ACCOUNT_KEY here"
        CONTAINER_NAME = "contractdocuments"

        # Initialize Blob service client
        credential = AzureNamedKeyCredential(ACCOUNT_NAME, ACCOUNT_KEY)
        blob_service_client = BlobServiceClient(
            account_url=f"https://{ACCOUNT_NAME}.blob.core.windows.net",
            credential=credential
        )
        container_client = blob_service_client.get_container_client(CONTAINER_NAME)
        blob_list = container_client.list_blobs()
    except Exception as e3:
        print(f"Error connecting to Blob storage: {e3}")
        return_value = f"Error connecting to Blob storage."
        return return_value
    
    try:
        return_value = "good"
        for blob in blob_list:
            try:
                if blob.name.lower().endswith((".docx", ".pdf")):
                    print(f"[DOWNLOADING] {blob.name}")
                    blob_client = container_client.get_blob_client(blob.name)
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_file:
                        try:
                            download_stream = blob_client.download_blob()
                            tmp_file.write(download_stream.readall())
                            temp_file_path = tmp_file.name
                        except Exception as e:
                            print(f"Failed to download blob {blob.name}: {e}")
                            continue
                    try:
                        print(f"[PROCESSING] {blob.name}")
                        process_document(temp_file_path, blob.name)
                    except Exception as e:
                        print(f"[ERROR] Failed to process {blob.name}: {e}")
                    finally:
                        os.remove(temp_file_path)  # Clean up the disk
            except Exception as e:
                print(f"Failed to open Word document: {blob.name} | Error: {e}")
                return_value = f"Failed to open Word document."
                print("return_value = ",return_value)
                return return_value # Assuming if one and only one file fails to process, indexing operation will fail.
        return return_value
    except Exception as e2:
        print(f"Error in function app: {e2}")
        return_value = f"Failed to open Word document."

        return return_value


