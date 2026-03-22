from io import BytesIO
import json
import logging
import os

import googleapiclient.discovery
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account

import config


def get_logger():
    return logging.getLogger(__name__)


def load_credentials():
    logger = get_logger()

    # parameter_name = "/google/auth/service_account/excel-nielsen-automate-drive-service-account"
    # logger.info(
    #     f"Fetching credentials from AWS Parameter Store with key: {parameter_name}"
    # )

    try:
        # ssm = boto3.client("ssm", region_name="us-east-2")
        # response = ssm.get_parameter(Name=parameter_name, WithDecryption=True)
        # credentials_json = response["Parameter"]["Value"]
        credentials_json = config.get_gdrive_service_account_credentials()

        credentials = service_account.Credentials.from_service_account_info(
            json.loads(credentials_json),
            scopes=["https://www.googleapis.com/auth/drive"],
        )
        # logger.info("Successfully loaded credentials from AWS Parameter Store")
        logger.info("Successfully loaded credentials")
        return credentials
    # except ClientError as e:
    #     raise Exception(
    #         f"Error fetching credentials from AWS Parameter Store: {str(e)}"
    #     )
    except Exception as e:
        raise Exception(f"Error loading credentials: {str(e)}")


def build_drive_service(credentials):
    logger = get_logger()

    try:
        service = googleapiclient.discovery.build("drive", "v3", credentials=credentials)
        logger.info("Successfully built Drive API service")
        return service
    except Exception as e:
        raise Exception(f"Error building Drive API service: {str(e)}")


def list_files(service, folder_id, drive_id=None):
    logger = get_logger()

    query = f"'{folder_id}' in parents"
    query += " and trashed = false"
    query += " and mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
    fields = "nextPageToken, files(id, name)"

    logger.info(f"Querying folder ID: {folder_id}")
    try:
        results = (
            service.files()
            .list(
                q=query,
                pageSize=1000,
                fields=fields,
                supportsAllDrives=drive_id is not None,
                includeItemsFromAllDrives=drive_id is not None,
                corpora="drive" if drive_id is not None else "user",
                driveId=drive_id,
            )
            .execute()
        )
        logger.info("Successfully executed files list query")
        return results.get("files", [])
    except Exception as e:
        logger.exception("Error accessing Drive folder or listing files: %s", str(e))
        return None


def download_file(service, file_id):
    logger = get_logger()

    try:
        request = service.files().get_media(fileId=file_id)
        file_data = request.execute()
        return BytesIO(file_data)
    except Exception as e:
        logger.exception("Error downloading file with ID %s: %s", file_id, str(e))
        return None


def create_folder(service, parent_folder_id, name, only_if_not_exists=False, drive_id=None):
    logger = get_logger()

    if only_if_not_exists:
        query = f"'{parent_folder_id}' in parents"
        query += " and trashed = false"
        query += " and mimeType = 'application/vnd.google-apps.folder'"
        query += f" and name = '{name}'"
        fields = "nextPageToken, files(id, name)"

        logger.info("Querying folder ID: %s for '%s'", parent_folder_id, name)
        try:
            results = (
                service.files()
                .list(
                    q=query,
                    fields=fields,
                    supportsAllDrives=drive_id is not None,
                    includeItemsFromAllDrives=drive_id is not None,
                    corpora="drive" if drive_id is not None else "user",
                    driveId=drive_id,
                )
                .execute()
            )
            logger.info("Successfully executed files list query")
            existing = results.get("files", [])
            if existing:
                return existing[0]["id"]
        except Exception as e:
            logger.exception("Error accessing Drive folder or listing files: %s", str(e))
            return None

    try:
        result = (
            service.files()
            .create(
                body={
                    "name": name,
                    "mimeType": "application/vnd.google-apps.folder",
                    "parents": [parent_folder_id],
                },
                fields="id",
                supportsAllDrives=drive_id is not None,
            )
            .execute()
        )
    except Exception as e:
        logger.exception("Error creating folder ('%s') in google drive: %s", name, str(e))
        return None
    return result.get("id")


def upload_file(service, filepath, folder_id, upload_filename=None, drive_id=None):
    logger = get_logger()

    media = MediaFileUpload(filepath, resumable=True)
    try:
        result = (
            service.files()
            .create(
                body={
                    "name": upload_filename or os.path.basename(filepath),
                    "parents": [folder_id],
                },
                media_body=media,
                fields="id",
                supportsAllDrives=drive_id is not None,
            )
            .execute()
        )
    except Exception as e:
        logger.exception("Error uploading file '%s' to google drive: %s", filepath, str(e))
        return None
    return result.get("id")


def transfer_ownership(service, file_id, new_owner_email, drive_id=None):
    logger = get_logger()

    try:
        result = (
            service.permissions()
            .create(
                body={
                    "type": "user",
                    "role": "writer",
                    "emailAddress": new_owner_email,
                },
                fileId=file_id,
                fields="id",
                supportsAllDrives=drive_id is not None,
            )
            .execute()
        )
    except Exception:
        logger.exception("Error transferring ownership of '%s' to '%s'", file_id, new_owner_email)
        return None
    return result.get("id")
