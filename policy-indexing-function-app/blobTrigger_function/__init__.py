import azure.functions as func
import logging
from os.path import basename
from . import policy_indexing

def main(blob: func.InputStream):
    logging.info("Blob trigger function is triggered.")
    logging.info(f"Blob name: {blob.name}, Size: {blob.length} bytes")

    try:
        filename = basename(blob.name)
        content = blob.read()
        result = policy_indexing.process_blob(filename, content)
        logging.info("✅ Document processed successfully.")

    except Exception as e:
        logging.exception("❌ Unhandled exception while processing blob")
        raise  # re-raise to signal failure to Azure
