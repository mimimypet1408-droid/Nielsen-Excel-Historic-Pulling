import dotenv

import os


def is_debug():
    debug = os.getenv("DEBUG")
    if debug is None:
        return False
    try:
        return int(debug)
    except ValueError:
        return False


def _get_required_env_var(name):
    value = os.environ[name]
    if not value:
        raise ValueError(f"{name} is not set")
    return value


def get_gdrive_service_account_credentials():
    return _get_required_env_var("GDRIVE_SERVICE_ACCOUNT_CREDENTIALS")


def get_gdrive_drive_id():
    return os.getenv("GOOGLE_DRIVE_DRIVE_ID")


def get_gdrive_folder_in_id():
    return _get_required_env_var("GOOGLE_DRIVE_IN_FOLDER_ID")


def get_gdrive_folder_out_id():
    return _get_required_env_var("GOOGLE_DRIVE_OUT_FOLDER_ID")


def get_nielsen_username():
    return _get_required_env_var("NIELSEN_USERNAME")


def get_nielsen_password():
    return _get_required_env_var("NIELSEN_PASSWORD")


dotenv.load_dotenv()
