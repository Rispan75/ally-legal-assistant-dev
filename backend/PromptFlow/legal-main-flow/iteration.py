import logging
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
from promptflow.core import tool
from promptflow.connections import CustomConnection
from summary_full_doc import get_policyinfo
 
 
# ------------------------------
# Configure minimal logging
# ------------------------------
# logging.basicConfig(level=logging.INFO)
# logger = logging.getLogger("policy_checker")
# logging.getLogger("azure.core.pipeline.policies.http_logging_policy").setLevel(logging.WARNING)
 
 
# ------------------------------
# Main function to find unused policies with details
# ------------------------------
@tool
def iterative_tool(searchconnection: CustomConnection, filename: str,language: str) -> object:
 
 
    credential = AzureKeyCredential(searchconnection.search_key)
    instruction_client = SearchClient(endpoint=searchconnection.search_endpoint, index_name=searchconnection.search_policy_index, credential=credential)
    document_client = SearchClient(endpoint=searchconnection.search_endpoint, index_name=searchconnection.search_document_index, credential=credential)
    # Step 1: Fetch all policy_ids from instruction index filtered by language
    #logger.info(f"Fetching policy IDs from instruction index filtered by language: {language}")
    instruction_filter = f"language eq '{language}'"
    instruction_results = list(instruction_client.search(search_text="*"))
    #instruction_results = list(instruction_client.search(search_text="*", filter=instruction_filter))
    print("instruction_results  ",  instruction_results)
    print("Policy Index:", searchconnection.search_policy_index)
    print("Endpoint:", searchconnection.search_endpoint)
 
 
    document_filter = f"filename eq '{filename}'"
    document_results = document_client.search(search_text="*", filter=document_filter)
    print("document_results  ",  document_results)
   
    all_policy_ids = set()
    for item in instruction_results:
        if 'PolicyId' in item:
            all_policy_ids.add(item['PolicyId'])
 
    #logger.info(f"Total unique policy IDs found for language '{language}': {len(all_policy_ids)}")
 
    # Step 2: Fetch document chunks filtered by filename
    #logger.info(f"Fetching document chunks from document index filtered by filename: {filename}")
    compliant_set = set()
    non_compliant_set = set()
    for chunk in document_results:
        compliant = chunk.get('CompliantCollection', [])
        non_compliant = chunk.get('NonCompliantCollection', [])
        compliant_set.update(compliant)
        non_compliant_set.update(non_compliant)
    used_policies = compliant_set | non_compliant_set
    unused_policies = all_policy_ids - used_policies
 
    # Step 3: Fetch title and instruction for each unused PolicyId
    #logger.info(f"Fetching detailed info for {len(unused_policies)} unused policy IDs...")
    result_list = []
 
    for policy_id in sorted(unused_policies):
        policy_info = get_policyinfo(policy_id, searchconnection)
        result_list.append({
            #"PolicyId": policy_id,
            #"Title": policy_info.get("title", "N/A") if policy_info else "N/A",
            #"Instruction": policy_info.get("instruction", "N/A") if policy_info else "N/A"
            "Title": policy_info.get("title", "N/A"),
            "summary":policy_info.get("summary","N/A") if policy_info else "N/A"
            #"Instruction": policy_info.get("instruction", "N/A"),
 
        })
 
    return result_list 
