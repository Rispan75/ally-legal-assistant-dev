import logging
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential
from promptflow.core import tool
from promptflow.connections import CustomConnection
from summary_full_doc import get_policyinfo

@tool
def iterative_tool(searchconnection: CustomConnection, filename: str, language: str) -> object:
    credential = AzureKeyCredential(searchconnection.search_key)

    instruction_client = SearchClient(
        endpoint=searchconnection.search_endpoint,
        index_name=searchconnection.search_policy_index,
        credential=credential
    )
    document_client = SearchClient(
        endpoint=searchconnection.search_endpoint,
        index_name=searchconnection.search_document_index,
        credential=credential
    )

    # Step 1: Fetch all policy_ids
    instruction_filter = f"language eq '{language}'"
    instruction_results = list(instruction_client.search(search_text="*"))
    # Optionally use the filter: instruction_client.search(search_text="*", filter=instruction_filter)

    document_filter = f"filename eq '{filename}'"
    document_results = document_client.search(search_text="*", filter=document_filter)

    all_policy_ids = set()
    for item in instruction_results:
        if 'PolicyId' in item:
            all_policy_ids.add(item['PolicyId'])

    compliant_set = set()
    non_compliant_set = set()
    for chunk in document_results:
        compliant = chunk.get('CompliantCollection', [])
        non_compliant = chunk.get('NonCompliantCollection', [])
        compliant_set.update(compliant)
        non_compliant_set.update(non_compliant)

    used_policies = compliant_set | non_compliant_set
    unused_policies = all_policy_ids - used_policies

    # Return early if no unused policies
    if not unused_policies:
        return [{
            "Title": "No unused Policies found",
            "summary": "There are no unused policies because all policies are either compliant or non-compliant."
        }]

    # Step 3: Fetch info for unused policies
    result_list = []
    for policy_id in sorted(unused_policies):
        policy_info = get_policyinfo(policy_id, searchconnection)
        result_list.append({
            "Title": policy_info.get("title", "N/A"),
            "summary": policy_info.get("summary", "N/A") if policy_info else "N/A"
        })

    return result_list
