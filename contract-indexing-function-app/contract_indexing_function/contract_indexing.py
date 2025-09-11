import os, re, uuid, json
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
from azure.search.documents.indexes.models import (
    SearchIndex, SimpleField, SearchableField, SearchField, SearchFieldDataType,
    VectorSearch, HnswAlgorithmConfiguration, HnswParameters,
    VectorSearchProfile, VectorSearchAlgorithmKind, VectorSearchAlgorithmMetric
)

# --- CONFIGURATION ---
DOC_INDEX = "legal-documents"
POLICY_INDEX = "legal-instructions"

AZURE_SEARCH_ENDPOINT = "XXXX"
AZURE_SEARCH_KEY = "XXXX"
 
AZURE_OPENAI_ENDPOINT = "XXXX"
AZURE_OPENAI_API_KEY = "XXXX"
 
AZURE_OPENAI_DEPLOYMENT = "gpt-4o"
AZURE_EMBEDDING_DEPLOYMENT = "text-embedding-ada-002"

LANG_CODE_TO_NAME = {
    "en": "English", "de": "German", "fr": "French", "es": "Spanish",
    "it": "Italian", "pt": "Portuguese", "nl": "Dutch", "ru": "Russian",
    "zh-cn": "Chinese", "ja": "Japanese"
}

# --- ENHANCED PROMPTS ---
ENHANCED_SPLIT_PROMPT = """
You are a legal document chunking expert. Your task is to divide a legal document into meaningful, section-based chunks.

INSTRUCTIONS:
1. Create chunks for major sections (1, 2, 3, etc.)
2. For subsections (4.1, 4.2, etc.), create separate chunks if they cover different legal concepts
3. Keep section headers with their content
4. Each chunk should be a complete legal concept

For each chunk, provide EXACTLY this JSON format:
{{"id": number, "title": "descriptive title", "text": "full text including headers", "keyphrases": ["term1", "term2", "term3"], "summary": "brief explanation"}}

Return ONLY a valid JSON array with no extra text:
[chunk1, chunk2, chunk3, ...]

Document to chunk:
{text}
"""

COMPLIANCE_PROMPT = """
You are a legal compliance classifier for contract clauses, comparing them to corporate policy requirements.

Your task is to classify the clause below as one of:
- "compliant": The clause aligns with the DEKRA Requirement and/or the Example Wording, even if not identical in wording.
- "non-compliant": The clause contradicts, exceeds, or fails to meet a key part of the DEKRA Requirement or contradicts DEKRA policy intent.
- "irrelevant": The clause is unrelated to the topic of the policy (even if it sounds legal), and therefore should not be judged under this policy.

🔍 HOW TO EVALUATE:

1. **Focus on the Policy Topic**  
   - Determine whether the clause addresses the same issue as the policy (e.g., confidentiality, governing law, penalties, etc.).  
   - If it doesn't address the same topic, classify as `"irrelevant"`.

2. **Check for Compliance**  
   - Compare the clause content with the **DEKRA Requirement** and **Example Wording**.  
   - Slight differences in language are acceptable as long as the clause respects the **intent, scope, and limits** of the requirement.  

3. **Identify Non-Compliance**  
   - Clauses that weaken or ignore a requirement, exceed allowed timeframes or legal standards, or omit essential protections are `"non-compliant"`.

4. **Ignore Negotiable Flags**  
   - The "Negotiable" or "Compromise Proposal" parts of the policy are for negotiation support and should NOT affect the classification.

✅ EXAMPLES:

✔️ Compliant:
- Clause: "This Agreement is governed by German law."
- Policy: Requires "Applicable law is German law, excluding conflict of laws provisions."

❌ Non-Compliant:
- Clause: "This Agreement is governed by Spanish law."
- Policy: Requires German law only.

🚫 Irrelevant:
- Clause: "Either party may assign this agreement to an affiliate."
- Policy: About governing law — this clause is unrelated.

📤 OUTPUT FORMAT:
Return **exactly one JSON object**, like:

{{"category": "compliant"}}
{{"category": "non-compliant"}}
{{"category": "irrelevant"}}

Clause:
{clause}

Policy:
{policy}
"""

# --- HELPER FUNCTIONS ---
def call_openai(prompt_text, max_retries=3):
    """Call OpenAI with retry logic and error handling"""
    client = AzureOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        api_version="2023-05-15"
    )
    
    for attempt in range(max_retries):
        try:
            resp = client.chat.completions.create(
                model=AZURE_OPENAI_DEPLOYMENT,
                messages=[{"role": "user", "content": prompt_text}],
                temperature=0,
                max_tokens=4000
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            print(f"[WARN] OpenAI call failed (attempt {attempt + 1}): {e}")
            if attempt == max_retries - 1:
                return None
    return None

def get_embedding(text):
    """Generate embedding for text with error handling"""
    try:
        client = AzureOpenAI(
            api_key=AZURE_OPENAI_API_KEY,
            azure_endpoint=AZURE_OPENAI_ENDPOINT,
            api_version="2023-05-15"
        )
        resp = client.embeddings.create(model=AZURE_EMBEDDING_DEPLOYMENT, input=[text])
        return resp.data[0].embedding
    except Exception as e:
        print(f"[WARN] Embedding generation failed: {e}")
        # Return a dummy embedding vector of the right size (1536 dimensions)
        return [0.0] * 1536

def extract_docx_text(path):
    """Extract text from DOCX file"""
    try:
        doc = Document(path)
        paragraphs = []
        for para in doc.paragraphs:
            if para.text.strip():
                paragraphs.append(para.text.strip())
        return "\n".join(paragraphs)
    except Exception as e:
        print(f"[ERROR] Failed to extract DOCX text: {e}")
        return ""

def extract_pdf_text(path):
    """Extract text from PDF file"""
    try:
        doc = fitz.open(path)
        text_parts = []
        for page in doc:
            page_text = page.get_text()
            if page_text.strip():
                text_parts.append(page_text)
        doc.close()
        
        full_text = "".join(text_parts).strip()
        if not full_text:
            return ocr_pdf_text(path)
        return full_text
    except Exception as e:
        print(f"[ERROR] Failed to extract PDF text: {e}")
        return ""

def ocr_pdf_text(path):
    """OCR text from PDF if regular extraction fails"""
    try:
        parts = []
        doc = fitz.open(path)
        for page in doc:
            pix = page.get_pixmap(dpi=300)
            img = Image.open(BytesIO(pix.tobytes()))
            text = pytesseract.image_to_string(img)
            if text.strip():
                parts.append(text)
        doc.close()
        return "\n".join(parts).strip()
    except Exception as e:
        print(f"[ERROR] OCR failed: {e}")
        return ""

def detect_language(text):
    """Detect language of the text"""
    try:
        lang_code = detect(text)
        return LANG_CODE_TO_NAME.get(lang_code, lang_code)
    except:
        return "English"  # Default to English if detection fails

def safe_json_parse(text):
    """Safely parse JSON from text with multiple fallback methods"""
    if not text:
        return None
    
    # Method 1: Try direct JSON parse
    try:
        return json.loads(text)
    except:
        pass
    
    # Method 2: Extract JSON array
    try:
        json_match = re.search(r'\[[\s\S]*\]', text)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    # Method 3: Extract JSON object
    try:
        json_match = re.search(r'\{[\s\S]*\}', text)
        if json_match:
            return json.loads(json_match.group())
    except:
        pass
    
    # Method 4: Fix common JSON issues and try again
    try:
        # Remove markdown code blocks
        cleaned = re.sub(r'```json\s*|\s*```', '', text)
        # Remove extra text before/after JSON
        cleaned = re.sub(r'^[^[\{]*', '', cleaned)
        cleaned = re.sub(r'[^\]\}]*$', '', cleaned)
        return json.loads(cleaned)
    except:
        pass
    
    return None

def enhanced_smart_split_document(text):
    """Enhanced document splitting with robust error handling"""
    print("[INFO] Starting enhanced document splitting...")
    
    # Limit text size for API call
    text_for_analysis = text[:10000] if len(text) > 10000 else text
    
    prompt = ENHANCED_SPLIT_PROMPT.format(text=text_for_analysis)
    
    # Try OpenAI first
    response = call_openai(prompt)
    if response:
        print(f"[DEBUG] OpenAI response received (length: {len(response)})")
        chunks = safe_json_parse(response)
        if chunks and isinstance(chunks, list) and len(chunks) > 0:
            print(f"[SUCCESS] Parsed {len(chunks)} chunks from OpenAI")
            return validate_chunks(chunks)
    
    print("[INFO] OpenAI parsing failed, using rule-based chunking...")
    return rule_based_document_split(text)

def validate_chunks(chunks):
    """Validate and clean chunk data"""
    validated_chunks = []
    for i, chunk in enumerate(chunks):
        try:
            if not isinstance(chunk, dict):
                continue
            
            validated_chunk = {
                "id": chunk.get("id", i + 1),
                "title": str(chunk.get("title", f"Section {i + 1}"))[:200],
                "text": str(chunk.get("text", ""))[:8000],
                "keyphrases": chunk.get("keyphrases", [])[:5] if isinstance(chunk.get("keyphrases"), list) else [],
                "summary": str(chunk.get("summary", ""))[:500]
            }
            
            if len(validated_chunk["text"]) > 30:  # Only keep chunks with substantial text
                validated_chunks.append(validated_chunk)
                
        except Exception as e:
            print(f"[WARN] Skipping invalid chunk {i}: {e}")
            continue
    
    return validated_chunks

def rule_based_document_split(text):
    """Rule-based document chunking as fallback"""
    print("[INFO] Using rule-based document splitting...")
    
    chunks = []
    chunk_id = 1
    
    # Split by major sections (numbered sections)
    section_patterns = [
        r'\n\s*(\d+\.?\s+[A-Z][^\n]*)\n',  # "1. Title" or "1 Title"
        r'\n\s*([A-Z][A-Z\s]+)\s*\n',      # "ALL CAPS TITLE"
        r'\n\s*([A-Z][^.\n]*)\s*\n'        # "Title Case"
    ]
    
    best_split = None
    best_pattern = None
    
    for pattern in section_patterns:
        sections = re.split(pattern, text, flags=re.IGNORECASE)
        if len(sections) > 3:  # Found good splits
            best_split = sections
            best_pattern = pattern
            break
    
    if best_split and len(best_split) > 1:
        print(f"[INFO] Found {len(best_split)//2} sections using pattern matching")
        
        current_title = ""
        current_text = ""
        
        for i, section in enumerate(best_split):
            if i == 0:  # First part (before first match)
                if section.strip():
                    chunks.append({
                        "id": chunk_id,
                        "title": "Document Header",
                        "text": section.strip(),
                        "keyphrases": extract_simple_keyphrases(section),
                        "summary": "Document introduction and header information"
                    })
                    chunk_id += 1
            elif i % 2 == 1:  # Odd indices are section titles
                current_title = section.strip()
            else:  # Even indices are section content
                current_text = section.strip()
                if current_text and len(current_text) > 30:
                    chunks.append({
                        "id": chunk_id,
                        "title": current_title or f"Section {chunk_id}",
                        "text": f"{current_title}\n\n{current_text}" if current_title else current_text,
                        "keyphrases": extract_simple_keyphrases(current_text),
                        "summary": f"Legal section covering {current_title.lower()}" if current_title else "Legal provisions"
                    })
                    chunk_id += 1
    
    else:
        # Last resort: paragraph-based splitting
        print("[INFO] Using paragraph-based splitting as final fallback")
        paragraphs = [p.strip() for p in text.split('\n\n') if p.strip() and len(p.strip()) > 50]
        
        for i, para in enumerate(paragraphs[:15]):  # Limit to 15 paragraphs
            chunks.append({
                "id": i + 1,
                "title": f"Paragraph {i + 1}",
                "text": para,
                "keyphrases": extract_simple_keyphrases(para),
                "summary": "Legal paragraph or clause"
            })
    
    print(f"[INFO] Created {len(chunks)} chunks using rule-based method")
    return chunks

def extract_simple_keyphrases(text):
    """Extract simple keyphrases using basic rules"""
    if not text:
        return []
    
    # Common legal terms
    legal_terms = [
        "confidential", "agreement", "party", "parties", "obligations", "disclosure",
        "governing law", "jurisdiction", "termination", "liability", "damages",
        "intellectual property", "breach", "compliance", "warranty", "contract"
    ]
    
    text_lower = text.lower()
    found_terms = [term for term in legal_terms if term in text_lower]
    
    # Add some key phrases (look for quoted terms)
    quoted_terms = re.findall(r'"([^"]*)"', text)
    found_terms.extend([term for term in quoted_terms if len(term) > 2 and len(term) < 30])
    
    return found_terms[:5]  # Return max 5 keyphrases

def load_policies():
    """Load policies from Azure Search with error handling"""
    try:
        client = SearchClient(
            endpoint=AZURE_SEARCH_ENDPOINT, 
            index_name=POLICY_INDEX, 
            credential=AzureKeyCredential(AZURE_SEARCH_KEY)
        )
        
        results = client.search("*", top=1000)
        policies = []
        
        for doc in results:
            policy = {
                "id": doc.get("PolicyId", f"policy_{len(policies)}"),
                "instruction": doc.get("instruction", ""),
                "language": doc.get("language", "English"),
                "original_text": doc.get("original_text", doc.get("instruction", ""))
            }
            
            if policy["original_text"]:  # Only add policies with text
                policies.append(policy)
        
        print(f"[INFO] Loaded {len(policies)} policies from Azure Search")
        return policies
        
    except Exception as e:
        print(f"[ERROR] Failed to load policies from Azure Search: {e}")
        # Return empty list - processing will continue without compliance checking
        return []

def check_compliance_with_gpt(clause_text, policies):
    """Check compliance of a clause against policies with robust error handling"""
    if not policies:
        print("[WARN] No policies available for compliance checking")
        return [], [], []
    
    compliant_ids, non_compliant_ids, irrelevant_ids = [], [], []
    
    print(f"[INFO] Checking compliance for clause against {len(policies)} policies...")
    
    for policy in policies:
        try:
            policy_text = policy["original_text"] if policy["original_text"] else policy["instruction"]
            
            if not policy_text:
                continue
            
            prompt = COMPLIANCE_PROMPT.format(clause=clause_text[:2000], policy=policy_text[:2000])
            
            response = call_openai(prompt)
            if not response:
                irrelevant_ids.append(policy["id"])
                continue
            
            # Parse compliance result
            result = safe_json_parse(response)
            if result and isinstance(result, dict):
                category = result.get("category", "irrelevant").lower()
            else:
                # Fallback text parsing
                response_lower = response.lower()
                if '"compliant"' in response_lower and '"non-compliant"' not in response_lower:
                    category = "compliant"
                elif '"non-compliant"' in response_lower:
                    category = "non-compliant"
                else:
                    category = "irrelevant"
            
            # Categorize
            if category == "compliant":
                compliant_ids.append(policy["id"])
            elif category == "non-compliant":
                non_compliant_ids.append(policy["id"])
            else:
                irrelevant_ids.append(policy["id"])
                
        except Exception as e:
            print(f"[WARN] Compliance check failed for policy {policy.get('id', 'unknown')}: {e}")
            irrelevant_ids.append(policy.get("id", "unknown"))
    
    print(f"[INFO] Compliance results - Compliant: {len(compliant_ids)}, Non-compliant: {len(non_compliant_ids)}, Irrelevant: {len(irrelevant_ids)}")
    return compliant_ids, non_compliant_ids, irrelevant_ids

def upload_chunk(filename, paragraph_id, title, text, language, embedding, 
                compliant_ids, non_compliant_ids, irrelevant_ids, keyphrases, summary):
    """Upload a single chunk to Azure Search with error handling"""
    try:
        client = SearchClient(
            endpoint=AZURE_SEARCH_ENDPOINT, 
            index_name=DOC_INDEX, 
            credential=AzureKeyCredential(AZURE_SEARCH_KEY)
        )
        
        # Determine overall compliance
        is_compliant = bool(compliant_ids) and not bool(non_compliant_ids)
        
        document = {
            "id": str(uuid.uuid4()),
            "filename": str(filename),
            "ParagraphId": int(paragraph_id),
            "title": str(title)[:200],
            "paragraph": str(text)[:8000],
            "embedding": embedding,
            "language": str(language),
            "isCompliant": bool(is_compliant),
            "CompliantCollection": list(compliant_ids),
            "NonCompliantCollection": list(non_compliant_ids),
            "IrrelevantCollection": list(irrelevant_ids),
            "group": [],
            "keyphrases": list(keyphrases)[:5],
            "summary": str(summary)[:500],
            "department": "",
            "date": datetime.utcnow().replace(tzinfo=timezone.utc).isoformat()
        }
        
        client.upload_documents(documents=[document])
        print(f"[SUCCESS] Uploaded chunk {paragraph_id} - '{title}' - Compliant: {is_compliant}")
        
    except Exception as e:
        print(f"[ERROR] Failed to upload chunk {paragraph_id}: {e}")


def create_index_if_not_exists():
    """Create the document index if it doesn't exist"""
    try:
        client = SearchIndexClient(
            endpoint=AZURE_SEARCH_ENDPOINT, 
            credential=AzureKeyCredential(AZURE_SEARCH_KEY)
        )
        
        existing_indexes = [idx.name for idx in client.list_indexes()]
        
        if DOC_INDEX in existing_indexes:
            print(f"[INFO] Index '{DOC_INDEX}' already exists")
            return
        
        # Define index fields
        fields = [
            SimpleField(name="id", type=SearchFieldDataType.String, key=True),
            SimpleField(name="ParagraphId", type=SearchFieldDataType.Int32, filterable=True, sortable=True),
            SearchableField(name="title", type=SearchFieldDataType.String, filterable=True),
            SearchableField(name="paragraph", type=SearchFieldDataType.String),
            SearchField(
                name="embedding", 
                type=SearchFieldDataType.Collection(SearchFieldDataType.Single),
                searchable=True, 
                vector_search_dimensions=1536, 
                vector_search_profile_name="vsProfile"
            ),
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
        
        # Vector search configuration
        vector_config = VectorSearch(
            algorithms=[
                HnswAlgorithmConfiguration(
                    name="vsAlgo",
                    kind=VectorSearchAlgorithmKind.HNSW,
                    parameters=HnswParameters(
                        metric=VectorSearchAlgorithmMetric.COSINE, 
                        m=4, 
                        ef_construction=200, 
                        ef_search=100
                    )
                )
            ],
            profiles=[
                VectorSearchProfile(name="vsProfile", algorithm_configuration_name="vsAlgo")
            ]
        )
        
        # Create index
        index = SearchIndex(name=DOC_INDEX, fields=fields, vector_search=vector_config)
        client.create_index(index)
        print(f"[SUCCESS] Created index '{DOC_INDEX}'")
        
    except Exception as e:
        print(f"[ERROR] Failed to create index: {e}")
 

def process_document(filename, filecontent):
    """Main document processing function with comprehensive error handling"""

    try:
        # Extract text based on file type
        document_text = filecontent
        
        if not document_text or len(document_text.strip()) < 100:
            print(f"[ERROR] Insufficient text extracted from {filename} (length: {len(document_text)})")
            return
        
        print(f"[INFO] Extracted {len(document_text)} characters from {filename}")
        
        # Detect language
        language = detect_language(document_text)
        print(f"[INFO] Detected language: {language}")
        
        # Load policies
        all_policies = load_policies()
        relevant_policies = [p for p in all_policies if p["language"].lower() == language.lower()]
        
        if not relevant_policies and all_policies:
            print(f"[WARN] No policies found for {language}, using all policies")
            relevant_policies = all_policies
        
        print(f"[INFO] Using {len(relevant_policies)} policies for compliance checking")
        
        # Split document into chunks
        chunks = enhanced_smart_split_document(document_text)
        
        if not chunks:
            print(f"[ERROR] No chunks created for {filename}")
            return
        
        print(f"[INFO] Created {len(chunks)} chunks from {filename}")
        
        # Process each chunk
        for i, chunk in enumerate(chunks):
            try:
                chunk_text = str(chunk.get("text", "")).strip()
                if len(chunk_text) < 30:
                    print(f"[SKIP] Chunk {i+1} too short ({len(chunk_text)} chars)")
                    continue
                
                chunk_id = chunk.get("id", i + 1)
                title = chunk.get("title", f"Section {chunk_id}")
                keyphrases = chunk.get("keyphrases", [])
                summary = chunk.get("summary", "Legal provision")
                
                print(f"[PROCESSING] Chunk {chunk_id}: {title[:50]}...")
                
                # Generate embedding
                embedding = get_embedding(chunk_text)
                
                # Check compliance
                compliant_ids, non_compliant_ids, irrelevant_ids = check_compliance_with_gpt(
                    chunk_text, relevant_policies
                )
                
                # Upload to search index
                upload_chunk(
                    filename, chunk_id, title, chunk_text, language, embedding,
                    compliant_ids, non_compliant_ids, irrelevant_ids, keyphrases, summary
                )
                
            except Exception as e:
                print(f"[ERROR] Failed to process chunk {i+1}: {e}")
                continue
        
        print(f"[COMPLETED] Successfully processed {filename}")
        
    except Exception as e:
        print(f"[ERROR] Document processing failed for {filename}: {e}")


def process_documents_main(filename, filecontent):
    create_index_if_not_exists()
    return_value = "good"
    try:
        print(f"[PROCESSING] {filename}")
        process_document(filename, filecontent)
    except Exception as e:
        print(f"[ERROR] Failed to process document - {filename}: {e}")
        return_value = "bad"
    return return_value
