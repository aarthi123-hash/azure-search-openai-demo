from quart import Blueprint, jsonify, request
from azure.core.credentials import AzureKeyCredential
from azure.search.documents.indexes import SearchIndexClient

search_indexes_bp = Blueprint("search_indexes", __name__)

@search_indexes_bp.route("/api/search-indexes", methods=["GET"])
async def list_indexes():
    endpoint = "PUT ENDPOINT HERE"  # <-- Replace with your search service name
    admin_key = "PUT INDEX KEY HERE"  # <-- Replace with your admin key
    client = SearchIndexClient(endpoint, AzureKeyCredential(admin_key))
    indexes = client.list_index_names()
    return jsonify({"indexes": list(indexes)})

@search_indexes_bp.route("/api/search-index-details/<index_name>", methods=["GET"])
async def get_index_details(index_name):
    endpoint = "PUT ENDPOINT HERE"
    admin_key = "PUT INDEX KEY HERE"
    client = SearchIndexClient(endpoint, AzureKeyCredential(admin_key))
    index = client.get_index(index_name)
    fields = [f.name for f in index.fields]
    semantic_config = (
        index.semantic_search.default_configuration_name
        if getattr(index, "semantic_search", None)
        else None
    )
    # Find all vector fields (fields with vector_search_dimensions set)
    vector_fields = [
        f.name for f in index.fields if getattr(f, "vector_search_dimensions", None)
    ]
    return jsonify({
        "fields": fields,
        "semantic_config": semantic_config,
        "vector_fields": vector_fields
    })