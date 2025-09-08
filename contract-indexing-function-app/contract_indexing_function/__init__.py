import azure.functions as func
import logging
import json
from . import contract_indexing

def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Contract indexing is triggered.")

## fetch input - filename from http request's json body
    try:
        req_body = req.get_json()
        filename = req_body.get("filename")
        filecontent = req_body.get("filecontent")

        return_value = contract_indexing.process_documents_from_blob(filename, filecontent)
        
        if return_value=="good":
            print("return value is passed as good")
            logging.info("docs processed successfully.")
            return func.HttpResponse(f"Document prepared.", status_code=200)
        else:
            logging.info("Error in function - process_documents.")
            return func.HttpResponse(f"Bad response!", status_code=500)
    except Exception as e:
        logging.info("Error in executing function app.")
        return func.HttpResponse(f"Error in running Function app", status_code=500)
