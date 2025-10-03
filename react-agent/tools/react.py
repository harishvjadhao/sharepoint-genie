import os
from typing import Any, Dict, List

from dotenv import load_dotenv
from langchain.chains.retrieval import create_retrieval_chain
from langchain.prompts import PromptTemplate
from langchain_community.vectorstores import FAISS
from langchain_core.tools import tool
from langchain_openai import AzureChatOpenAI, AzureOpenAIEmbeddings
from langchain_text_splitters import CharacterTextSplitter
from langgraph.types import Command, interrupt

from context_store import get_all_context, get_context_value
from services.sharepoint_client import SharePointClient
from utils.output_parsers import Summary, summary_parser

load_dotenv()

template = """
SYSTEM
Answer any use questions based solely on the context below:

<context>
{context}
</context>

MESSAGES LIST
chat_history

HUMAN
{input}
\n{format_instructions}
"""


def _init_sharepoint_client() -> SharePointClient:
    """Initialize SharePointClient from request context, falling back to environment variables.
    override_site_url takes precedence if provided.
    """
    return SharePointClient(
        site_url=get_context_value("SITE_URL"),
        site_id=get_context_value("SITE_ID"),
        access_token=get_context_value("ACCESS_TOKEN"),
        obo_access_token=get_context_value("OBO_ACCESS_TOKEN"),
    )


@tool
def get_one_drive_id() -> str:
    """
    Returns the unique ID of the current user's OneDrive drive.

    Args:
        None

    Returns:
        str: The unique identifier (ID) of the user's OneDrive drive. If an error occurs, returns an error message.
    """
    try:
        client = _init_sharepoint_client()
        drive_id = client.get_one_drive_id()
        return drive_id
    except Exception as e:
        return f"Error: {str(e)}"


@tool
def get_drive_id(library_name: str) -> str:
    """
    Returns the unique ID of a document library (SharePoint drive) by its name.

    Args:
        library_name (str): The display name of the document library (e.g., 'Documents', 'Shared Documents', 'Policies').

    Returns:
        str: The unique identifier (ID) of the specified document library. If the library does not exist or an error occurs, returns an error message.
    """
    try:
        client = _init_sharepoint_client()
        drive_id = client.get_drive_id(library_name)
        return drive_id
    except Exception as e:
        return f"Error: {str(e)}"


@tool
def get_folder_id(drive_id: str, folder_path: str = "") -> str:
    """
    Returns the unique ID of a folder in a SharePoint document library or OneDrive.

    Args:
        drive_id (str): The ID of the document library or OneDrive drive.
        folder_path (str, optional): The path to the folder within the drive. Use an empty string to get the root folder.

    Returns:
        str: The unique identifier (ID) of the specified folder. If the folder does not exist or an error occurs, returns an error message.
    """
    try:
        client = _init_sharepoint_client()
        folder_id = client.get_folder_id(drive_id, folder_path)
        return folder_id
    except Exception as e:
        return f"Error: {str(e)}"


@tool
def recent_sharepoint_files(
    drive_id: str,
    file_name: str,
    top: int,
    file_download: bool = False,
) -> List[Dict[str, Any]]:
    """
    Returns a list of recent files from a SharePoint document library.

    Args:
        drive_id (str): The unique ID of the document library (SharePoint drive).
        file_name (str): The name or partial name of the file to search for. Use '.' to match all files.
        top (int): The maximum number of files to return.
        file_download (bool): If True, include direct download URLs for each file in the results.

    Returns:
        List[Dict[str, Any]]: Each dictionary contains metadata about a file, with the following keys:
            - name (str): The file's name.
            - modified (str): The last modified date and time of the file.
            - webUrl (str): The web URL to view the file in SharePoint.
            - id (str): The unique identifier of the file.
            - size (int): The size of the file in bytes.
            - drive_id (str): The ID of the drive (document library) where the file is stored.
            - folder_id (str): The ID of the folder containing the file.
            - file_type (str): The MIME type of the file.
            - download_url (str or None): The direct URL to download the file, if requested.
            - created_by (str or None): The display name of the user who created the file.
            - last_modified_by (str or None): The display name of the user who last modified the file.
    """
    try:
        client = _init_sharepoint_client()
        files = client.get_all_files_in_drive(drive_id, file_name, top)
        simplified = [
            {
                "name": f.get("name"),
                "modified": f.get("lastModifiedDateTime"),
                "webUrl": f.get("webUrl"),
                "id": f.get("id"),
                "size": f.get("size"),
                "drive_id": f.get("parentReference", {}).get("driveId"),
                "folder_id": f.get("parentReference", {}).get("id"),
                "file_type": f.get("file", {}).get("mimeType"),
                "download_url": (
                    client.get_file_download_url(drive_id, f.get("id"))
                    if file_download
                    else None
                ),
                "created_by": f.get("createdBy", {}).get("user", {}).get("displayName"),
                "last_modified_by": f.get("lastModifiedBy", {})
                .get("user", {})
                .get("displayName"),
            }
            for f in files
        ]
        return simplified
    except Exception as e:
        return f"Error: {str(e)}"


@tool
def recent_onedrive_files(
    file_name: str,
    top: int,
    file_download: bool = False,
) -> List[Dict[str, Any]]:
    """
    Returns a list of recent files from the user's OneDrive account.

    Args:
        file_name (str): The name or partial name of the file to search for. Use '.' to match all files.
        top (int): The maximum number of files to return.
        file_download (bool): If True, include direct download URLs for each file in the results.

    Returns:
        List[Dict[str, Any]]: Each dictionary in the list contains metadata about a file, with the following keys:
            - name (str): The file's name.
            - modified (str): The last modified date and time of the file.
            - webUrl (str): The web URL to view the file in OneDrive.
            - id (str): The unique identifier of the file.
            - size (int): The size of the file in bytes.
            - drive_id (str): The ID of the drive (OneDrive) where the file is stored.
            - folder_id (str): The ID of the folder containing the file.
            - file_type (str): The MIME type of the file.
            - download_url (str or None): The direct URL to download the file, if requested.
            - created_by (str or None): The display name of the user who created the file.
            - last_modified_by (str or None): The display name of the user who last modified the file.
    """
    try:
        client = _init_sharepoint_client()
        files = client.get_recent_onedrive_files(file_name, top)
        simplified = [
            {
                "name": f.get("name"),
                "modified": f.get("lastModifiedDateTime"),
                "webUrl": f.get("webUrl"),
                "id": f.get("id"),
                "size": f.get("size"),
                "drive_id": f.get("parentReference", {}).get("driveId"),
                "folder_id": f.get("parentReference", {}).get("id"),
                "file_type": f.get("file", {}).get("mimeType"),
                "download_url": (
                    client.get_file_download_url(
                        f.get("parentReference", {}).get("driveId"), f.get("id")
                    )
                    if file_download
                    else None
                ),
                "created_by": f.get("createdBy", {}).get("user", {}).get("displayName"),
                "last_modified_by": f.get("lastModifiedBy", {})
                .get("user", {})
                .get("displayName"),
            }
            for f in files
        ]
        return simplified
    except Exception as e:
        return f"Error: {str(e)}"


@tool
def copy_onedrive_file(file_id: str, drive_id: str, folder_id: str) -> Dict[str, Any]:
    """
    Copies a file from OneDrive to a specified SharePoint document library and folder.

    Args:
        file_id (str): The ID of the file in OneDrive to be copied.
        drive_id (str): The ID of the destination document library (SharePoint drive).
        folder_id (str): The ID of the destination folder within the document library.

    Returns:
        Dict[str, Any]: A dictionary describing the result of the copy operation:
            - success (bool): True if the copy was successful, False otherwise.
            - error (str, optional): Error message if the copy failed.
    """
    try:
        client = _init_sharepoint_client()
        client.copyfile(file_id, drive_id, folder_id)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


@tool
def get_site_analytics() -> Dict[str, Any]:
    """
    Retrieves analytics data for the current SharePoint site.

    Args:
        None

    Returns:
        Dict[str, Any]: A dictionary containing site analytics data. If the request is successful, the dictionary includes analytics properties such as usage statistics, activity counts, or other relevant metrics. If an error occurs, the dictionary contains:
            - success (bool): False if the request failed.
            - error (str): Error message describing the failure.
    """
    try:
        client = _init_sharepoint_client()
        analytics = client.get_site_analytics()
        return analytics
    except Exception as e:
        return {"success": False, "error": str(e)}


@tool
def summarize_file(drive_id: str, file_name: str) -> Dict[str, Any]:
    """
    Summarizes the content of a document from a SharePoint document library or OneDrive.

    Args:
        drive_id (str): The ID of the document library containing the file.
        file_name (str): The name of the document to summarize.

    Returns:
        Dict[str, Any]: A dictionary containing the summary and metadata of the file. If successful, the dictionary includes:
            - Summary (str): A short, clear title and a brief neutral summary of the document.
        If an error occurs, the dictionary includes:
            - error (str): Error message describing the failure.
    """
    try:
        client = _init_sharepoint_client()

        # return loader and file_id
        loader = client.load_sharepoint_document_by_name(drive_id, file_name)

        text_splitter = CharacterTextSplitter(
            separator="\n", chunk_size=1000, chunk_overlap=0
        )
        docs = loader.load_and_split(
            text_splitter=text_splitter
        )  # Uncomment this line if you want to use the specific text splitter.

        embeddings = AzureOpenAIEmbeddings()
        vectorstore = FAISS.from_documents(docs, embeddings)
        vectorstore.save_local("faiss_index")

        new_vectorstore = FAISS.load_local(
            "faiss_index", embeddings, allow_dangerous_deserialization=True
        )
        retrieval_qa_chat_prompt = PromptTemplate.from_template(
            template,
            partial_variables={
                "format_instructions": summary_parser.get_format_instructions()
            },
        )

        llmm = AzureChatOpenAI(deployment_name="gpt-4o")

        combine_docs_chain = retrieval_qa_chat_prompt | llmm | summary_parser

        retrieval_chain = create_retrieval_chain(
            new_vectorstore.as_retriever(), combine_docs_chain
        )
        query = "Please provide a brief, neutral summary of the document in 3–4 sentences, and suggest a short, clear title that reflects its main topic."
        response = retrieval_chain.invoke({"input": query})
        answer: Summary = response["answer"]

        metadata = {"Summary": answer.title + "|" + answer.summary}

        return metadata

    except Exception as e:
        return {"error": str(e)}


@tool
def update_file_metadata(
    drive_id: str,
    file_id: str,
    metadata: Dict[str, Any],
) -> Dict[str, Any]:
    """
    Updates the metadata of a file in SharePoint or OneDrive.

    Args:
        drive_id (str): The ID of the document library or OneDrive drive.
        file_id (str): The ID of the file to update.
        metadata (Dict[str, Any]): A dictionary of metadata fields to update (e.g., {"Summary": "New Summary"}).

    Returns:
        Dict[str, Any]: A dictionary describing the result of the update operation:
            - success (bool): True if the update was successful, False otherwise.
            - error (str, optional): Error message if the update failed.
    """
    try:
        client = _init_sharepoint_client()
        client.update_file_metadata(drive_id, file_id, metadata)
        return {"success": True}
    except Exception as e:
        return {"error": str(e)}


tools = [
    get_one_drive_id,
    get_drive_id,
    get_folder_id,
    recent_sharepoint_files,
    recent_onedrive_files,
    copy_onedrive_file,
    summarize_file,
    update_file_metadata,
    get_site_analytics,
]

llm = AzureChatOpenAI(deployment_name="gpt-4o", temperature=0).bind_tools(tools)
