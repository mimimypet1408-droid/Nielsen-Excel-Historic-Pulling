import argparse
import datetime
import logging
import math
import os.path
import queue
import signal
import subprocess
import sys
import tempfile
import time
import threading
from typing import Optional

from appium import webdriver
from appium.options.windows import WindowsOptions
from appium.webdriver.common.appiumby import AppiumBy
import excel_utils
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, NoSuchElementException, TimeoutException

import config
from errors import DataPlanLimitReachedError, FailureLimitExceededError, WeekEndingDateError
import gdrive
from constants import APPIUM_HOST, APPIUM_PORT, WAD_PORT, WAD_EXE, ROOT_DIR, AppiumServerState
import log

FILENAME_DATE_FORMAT = "%m%d%y"


def wait_until_visible(wait, locator: str, strategy: str = AppiumBy.XPATH):
    return wait.until(EC.visibility_of_element_located((strategy, locator)))


def _find_username_input(driver, raise_exception=True, wait=None):
    logger = get_logger()
    try:
        if wait:
            return wait_until_visible(wait, "input27", AppiumBy.ACCESSIBILITY_ID)
        return driver.find_element(AppiumBy.ACCESSIBILITY_ID, "input27")
    except (NoSuchElementException, TimeoutException):
        if raise_exception:
            raise
        logger.exception("Failed to find username input")
        return None


def _find_got_it_button(driver, raise_exception=True, wait=None):
    logger = get_logger()
    try:
        if wait:
            return wait_until_visible(wait, "Got it", AppiumBy.NAME)
        return driver.find_element(AppiumBy.NAME, "Got it")
    except (NoSuchElementException, TimeoutException):
        if raise_exception:
            raise
        logger.exception("Failed to find got it button")
        return None


def _find_data_plan_limit_reached_message(driver, raise_exception=True, wait=None):
    logger = get_logger()
    try:
        if wait:
            return wait_until_visible(wait, "Data plan limit reached", AppiumBy.NAME)
        return driver.find_element(AppiumBy.NAME, "Data plan limit reached")
    except (NoSuchElementException, TimeoutException):
        if raise_exception:
            raise
        logger.exception("Failed to find data plan limit reached message")
        return None


def _login_nielseniq(driver, wait):
    logger = get_logger()

    # username
    logger.info("find username")
    # xpath_LeftClickEditUsername_108_30 = f'{xpath_prefix}/Window[@ClassName="NUIDialog"][starts-with(@Name,"NielsenIQ Discover - https://login.identity.nielseniq.com/oauth2")]//Edit[@Name="Username "][starts-with(@AutomationId,"input")]'
    winElem_LeftClickEditUsername_108_30 = wait_until_visible(
        wait, "input27", AppiumBy.ACCESSIBILITY_ID
    )
    # KeyboardInput username
    logger.info("KeyboardInput username")
    time.sleep(0.1)
    username = config.get_nielsen_username()
    if username:
        winElem_LeftClickEditUsername_108_30.send_keys(username)

    # LeftClick on Button "Next" at (18,16)
    logger.info('LeftClick on Button "Next" at (18,16)')
    # xpath_LeftClickButtonNext_18_16 = f'{xpath_prefix}/Window[@ClassName="NUIDialog"][starts-with(@Name,"NielsenIQ Discover - https://login.identity.nielseniq.com/oauth2")]//Button[@ClassName="button button-primary"][@Name="Next"]'
    winElem_LeftClickButtonNext_18_16 = wait_until_visible(wait, "Next", AppiumBy.NAME)
    try:
        winElem_LeftClickButtonNext_18_16.click()
    except WebDriverException:
        # Sometimes the button is blocked by an element (autocomplete popup)
        # try clicking again
        winElem_LeftClickButtonNext_18_16.click()

    # password
    logger.info("find password input")
    # xpath_LeftClickEditPassword_64_22 = f'{xpath_prefix}/Window[@ClassName="NUIDialog"][starts-with(@Name,"NielsenIQ Discover - https://login.identity.nielseniq.com/oauth2")]//Edit[@Name="Password "][starts-with(@AutomationId,"input")]'
    winElem_LeftClickEditPassword_64_22 = wait_until_visible(
        wait, "input52", AppiumBy.ACCESSIBILITY_ID
    )
    # KeyboardInput password
    logger.info("KeyboardInput password")
    time.sleep(0.1)
    password = config.get_nielsen_password()
    if password:
        winElem_LeftClickEditPassword_64_22.send_keys(password)

    # LeftClick on Button "Verify" at (47,21)
    logger.info('LeftClick on Button "Verify" at (47,21)')
    # xpath_LeftClickButtonVerify_47_21 = f'{xpath_prefix}/Window[@ClassName="NUIDialog"][starts-with(@Name,"NielsenIQ Discover - https://login.identity.nielseniq.com/oauth2")]//Button[@ClassName="button button-primary"][@Name="Verify"]'
    wait_until_visible(wait, "Verify", AppiumBy.NAME).click()


def _refresh_sales(driver, wait, excel_filepath: str):
    logger = get_logger()

    # LeftClick on GotIt
    logger.info("LeftClick on GotIt")
    # xpath_LeftClickButtonGotit_39_16 = f'{xpath_prefix}//Button[@ClassName="ms-Button ms-Button--primary root-142"][@Name="Got it"]'
    try:
        wait_until_visible(wait, "Got it", AppiumBy.NAME).click()
    except TimeoutException as e:
        logger.warning("'Got It' button not found, check for data limit message.")
        try:
            wait_until_visible(wait, "Data plan limit reached", AppiumBy.NAME)
            raise DataPlanLimitReachedError([excel_filepath])
        except TimeoutException:
            logger.warning("Data plan limit message not found either, re-raise original exception.")
            raise e

    # # LeftClick on Button "Range Menu" at (14,16)
    # logger.info('LeftClick on Button "Range Menu" at (14,16)')
    # wait_until_visible(wait, "Range Menu", AppiumBy.NAME).click()

    # # LeftClick on Group "" at (189,207)
    # logger.info("LeftClick on Refresh Button")
    # # xpath_RefreshButton = (
    # #     f'{xpath_prefix}//Button[@Name="Refresh Refresh"][@AutomationId="menuButton"]'
    # # )
    # wait_until_visible(wait, "Refresh Refresh", AppiumBy.NAME).click()

    # Left click on "Refresh All" to refresh all ranges
    wait_until_visible(wait, "Refresh All", AppiumBy.NAME).click()

    # Wait for up to a 10 minutes and click OK
    wait_for_ok_after_refresh = WebDriverWait(driver, 600)
    logger.info("LeftClick on OK")
    # xpath_LeftClickButtonOK_31_16 = f'{xpath_prefix}//Button[@ClassName="ms-Button ms-Button--primary root-142"][@Name="OK"]'
    wait_until_visible(wait_for_ok_after_refresh, "OK", AppiumBy.NAME).click()


def create_driver(desired_caps, host=APPIUM_HOST, port=APPIUM_PORT):
    driver = webdriver.Remote(
        command_executor=f"{host}:{port}",
        options=WindowsOptions().load_capabilities(desired_caps),
    )
    driver.implicitly_wait(0)
    return driver


def wait_until_file_saved_to_this_pc_visible(driver, filename, timeout=60):
    wait_for_save_to_complete = WebDriverWait(driver, timeout)
    try:
        wait_until_visible(
            wait_for_save_to_complete,
            f"\u202a{filename}\u202c  -  Saved to this PC",
            AppiumBy.NAME,
        )
    except TimeoutException:
        # If the file save info is not found, try the filename without the extension
        wait = WebDriverWait(driver, 20)
        wait_until_visible(
            wait,
            f"\u202a{os.path.splitext(filename)[0]}\u202c  -  Saved to this PC",
            AppiumBy.NAME,
        )


def click_save_in_excel(driver, filename, logger):
    wait = WebDriverWait(driver, 20)
    # Click on the save button
    logger.info("Save")
    # xpath_ExcelWindow = f"{xpath_prefix}"
    wait_until_visible(wait, "FileSave", AppiumBy.ACCESSIBILITY_ID).click()
    # Wait for save to complete
    wait_until_file_saved_to_this_pc_visible(driver, filename)


def _update_sales(driver, excel_filepath):
    logger = get_logger()

    filename = os.path.basename(excel_filepath)
    # xpath_prefix = f'/Window[@ClassName="XLMAIN"][@Name="{filename} - Excel"]'

    # LeftClick on TabItem "Insert" at (27,17)
    logger.info('LeftClick on TabItem "Insert" at (27,17)')
    # xpath_LeftClickTabItemInsert_27_17 = (
    #     f'{xpath_prefix}//TabItem[@Name="Insert"][@AutomationId="TabInsert"]'
    # )
    wait_for_file_load_and_insert_tab = WebDriverWait(driver, 60)
    wait_until_visible(wait_for_file_load_and_insert_tab, "Insert", AppiumBy.NAME).click()

    # Maximize the window
    driver.maximize_window()

    # LeftClick on Button "NielsenIQ Discover" at (30,22)
    logger.info('LeftClick on Button "NielsenIQ Discover" at (30,22)')
    xpath_LeftClickButtonNielsenIQD_30_22 = (
        '//Button[@ClassName="NetUIRibbonButton"][@Name="NielsenIQ Discover"]'
    )
    wait = WebDriverWait(driver, 20)
    wait_until_visible(wait, xpath_LeftClickButtonNielsenIQD_30_22).click()

    # LeftClick on Text "Select your region here" at (141,9)
    logger.info('LeftClick on Text "Select your region here" at (141,9)')
    # xpath_LeftClickTextSelectyour_141_9 = (
    #     f'{xpath_prefix}//Text[@Name="Select your region here"]'
    # )
    wait_until_visible(wait, "Select your region here", AppiumBy.NAME).click()

    # LeftClick on Group "" at (18,14)
    logger.info('LeftClick on Group "" at (18,14)')
    # xpath_LeftClickGroup_18_14 = (
    #     f'{xpath_prefix}//Group[@ClassName="selectCustom-option"]'
    # )
    wait_until_visible(wait, "US", AppiumBy.NAME).click()

    # We need to check if we are logged in, or if we need to login again
    timeout = 20
    attempt_interval = 0.5
    need_to_login: Optional[bool] = None
    for _ in range(math.ceil(timeout / attempt_interval)):
        data_plan_limit_reached = _find_data_plan_limit_reached_message(
            driver, raise_exception=False
        )
        if data_plan_limit_reached:
            raise DataPlanLimitReachedError([excel_filepath])
        got_it_button = _find_got_it_button(driver, raise_exception=False)
        if got_it_button:
            need_to_login = False
            break
        username_input = _find_username_input(driver, raise_exception=False)
        if username_input:
            need_to_login = True
            break
        time.sleep(attempt_interval)

    if need_to_login is None:
        raise Exception("Could not determine if login is needed")
    if need_to_login is True:
        _login_nielseniq(driver, wait)
    _refresh_sales(driver, wait, excel_filepath)

    # Click on the save button
    click_save_in_excel(driver, filename, logger)


def start_appium_server(port=APPIUM_PORT):
    logger = get_logger()
    logger.info("Starting Appium server on port %d", port)
    start_time = time.time()
    appium_start_timeout = 60
    process = subprocess.Popen(
        [".\\node_modules\\.bin\\appium.cmd", "-p", str(port)],
        cwd=ROOT_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )

    # Start thread to capture and log stdout/stderr
    server_state_queue = queue.Queue()

    def log_stream(stream):
        server_state = AppiumServerState.STARTING
        appium_logger = logging.getLogger("appium_server")
        for line in iter(stream.readline, ""):
            if server_state == AppiumServerState.STARTING:
                if "Appium REST http interface listener started on" in line:
                    server_state = AppiumServerState.LISTENING
                    server_state_queue.put(server_state)
                elif time.time() - start_time > appium_start_timeout:
                    server_state = AppiumServerState.TIMED_OUT
                    server_state_queue.put(server_state)

            appium_logger.info("%s", line.strip())

    threading.Thread(target=log_stream, args=(process.stdout,), daemon=True).start()

    while True:
        server_state = server_state_queue.get()
        if server_state == AppiumServerState.TIMED_OUT:
            raise Exception("Appium server did not start in time")
        elif server_state == AppiumServerState.LISTENING:
            break

    return process


def stop_appium_server(process):
    logger = get_logger()
    logger.info("Stopping Appium server")
    # NOTE: process.terminate() does not work reliably in this case, so we use taskkill
    # process.terminate()
    process.stdout.flush()
    subprocess.call(["taskkill.exe", "/PID", str(process.pid), "/T", "/F"])
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        logger.error("Appium server did not stop in time, try CTRL + C")
        process.send_signal(signal.CTRL_C_EVENT)
    else:
        logger.info("Appium server stopped")


def start_wad_server(port=WAD_PORT):
    logger = get_logger()
    logger.info("Starting WinAppDriver on port %d", port)
    process = subprocess.Popen(
        [WAD_EXE, str(port)],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )
    # Give WAD a moment to start and check it hasn't exited immediately
    time.sleep(3)
    if process.poll() is not None:
        output = process.stdout.read() if process.stdout else ""
        raise Exception(f"WinAppDriver exited immediately with code {process.returncode}: {output}")
    logger.info("WinAppDriver started (pid=%d)", process.pid)
    return process


def stop_wad_server(process):
    logger = get_logger()
    logger.info("Stopping WinAppDriver")
    subprocess.call(["taskkill.exe", "/PID", str(process.pid), "/T", "/F"])
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        logger.error("WinAppDriver did not stop in time")
    else:
        logger.info("WinAppDriver stopped")


def get_argparser():
    parser = argparse.ArgumentParser()

    subcommands = parser.add_subparsers(dest="subcommand")
    subcommands.add_parser(
        "gdrive-weekly",
        help="Downloads and refreshes the sales from all Excel files in GOOGLE_DRIVE_IN_FOLDER_ID. Saves output to <YYYY-MM-DD> subdirectory in GOOGLE_DRIVE_OUT_FOLDER_ID.",
    )

    files = subcommands.add_parser("files", help="Refresh local Excel files")
    files.add_argument("excel_files", nargs="+", help="Excel files to update")

    return parser


def create_driver_for_file(excel_filepath):
    excel_filepath = os.path.normpath(os.path.join(ROOT_DIR, excel_filepath))
    desired_caps = {
        "app": "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
        "appArguments": f'"{excel_filepath}"',
        "newCommandTimeout": 240,
        # "ms:waitForAppLaunch": 120,
    }

    return create_driver(desired_caps)


def get_logger():
    return logging.getLogger(__name__)


def subcommand_files(args):
    logger = get_logger()

    excel_files = args.excel_files

    if not config.get_nielsen_username() or not config.get_nielsen_password():
        logger.error("The Nielsen credentials are not set. Do you have a .env file?")
        return 1

    today = datetime.date.today()
    nielsen_week_ending = get_latest_available_nielsen_week_ending(today)

    previous_driver = None
    driver = None
    appium_server = None
    wad_server = None
    failure_limit = 2
    failed_files = []
    try:
        wad_server = start_wad_server()
        appium_server = start_appium_server()
        for excel_filepath in excel_files:
            previous_driver = driver
            try:
                driver = create_driver_for_file(excel_filepath)
                wait_until_file_saved_to_this_pc_visible(
                    driver, os.path.basename(excel_filepath), timeout=120
                )
                _update_sales(driver, excel_filepath)
                # On success, detect the min / max periods from the file and confirm it aligns
                # with the expected week ending date
                try:
                    min_period, max_period = (
                        excel_utils.get_min_max_nielsen_periods_from_excel_file(excel_filepath)
                    )
                except excel_utils.MissingPeriodsColumnError as e:
                    logger.warning(
                        "The 'Periods' column is missing from the file '%s'. Falling back to using "
                        "the expected week ending date %s. Detail: %s",
                        excel_filepath,
                        nielsen_week_ending,
                        str(e),
                    )
                    min_period = nielsen_week_ending
                    max_period = nielsen_week_ending
                if max_period != nielsen_week_ending:
                    logger.error(
                        "The max period date %s in file '%s' does not match the expected week ending date %s",
                        max_period,
                        excel_filepath,
                        nielsen_week_ending,
                    )
                    raise WeekEndingDateError(
                        os.path.basename(excel_filepath),
                        nielsen_week_ending.strftime(FILENAME_DATE_FORMAT),
                        max_period.strftime(FILENAME_DATE_FORMAT),
                    )
            except Exception as e:
                logger.exception("Failed to refresh '%s': %s", excel_filepath, str(e))
                failed_files.append(excel_filepath)
                # Take a screenshot of the error
                if driver:
                    _take_screenshot(driver, os.path.basename(excel_filepath))
                    # Try to save the file so that we avoid the save prompt on exit
                    try:
                        click_save_in_excel(driver, os.path.basename(excel_filepath), logger)
                    except Exception as e2:
                        logger.exception(
                            "Failed to save file after error '%s': %s",
                            excel_filepath,
                            str(e2),
                        )
                # If the error was due to data plan limit, raise specific exception
                if isinstance(e, DataPlanLimitReachedError):
                    raise DataPlanLimitReachedError(failed_files) from e

                # If more than failure_limit failures, raise FailureLimitExceededError
                if len(failed_files) >= failure_limit:
                    raise FailureLimitExceededError(failed_files) from e
            else:
                # Rename the file to include the min/max periods
                excel_filepath_with_dates = get_filepath_with_dates(
                    excel_filepath, min_period, max_period
                )
                try:
                    os.rename(excel_filepath, excel_filepath_with_dates)
                    logger.info(
                        "Renamed '%s' to '%s'",
                        excel_filepath,
                        excel_filepath_with_dates,
                    )
                except Exception as e:
                    logger.exception(
                        "Failed to rename '%s' to '%s': %s",
                        excel_filepath,
                        excel_filepath_with_dates,
                        str(e),
                    )
            finally:
                if previous_driver:
                    try:
                        previous_driver.quit()
                    except Exception as e:
                        logger.exception("Failed to quit driver: %s", str(e))
    finally:
        if driver:
            try:
                driver.quit()
            except Exception as e:
                logger.exception("Failed to quit driver: %s", str(e))
            driver = None
        if appium_server:
            stop_appium_server(appium_server)
            appium_server = None
        if wad_server:
            stop_wad_server(wad_server)
            wad_server = None
    return 0


def _take_screenshot(driver, base_filename) -> str:
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filepath = os.path.join(
        log.get_logs_dir(),
        f"{os.path.splitext(base_filename)[0]}_{timestamp}.png",
    )
    driver.save_screenshot(filepath)
    return filepath


def get_nielsen_week_ending(date: datetime.date) -> datetime.date:
    """Returns the date of the most recent Nielsen week ending (Saturday) on or before the given date"""
    return date - datetime.timedelta(days=(date.weekday() + 2) % 7)


def get_latest_available_nielsen_week_ending(date: Optional[datetime.date] = None) -> datetime.date:
    """Returns the date of the latest available Nielsen week ending (Saturday) on or before the
    given date (or today if no date is given).
    The latest week ending is one week before the most recent Saturday relative to the given date.
    """
    if date is None:
        date = datetime.date.today()
    most_recent_week_ending = get_nielsen_week_ending(date)
    return most_recent_week_ending - datetime.timedelta(days=7)


def get_filepath_with_dates(
    file_path: str, min_date: datetime.date, max_date: datetime.date
) -> str:
    """Returns a new file path with the min and max dates appended to the filename
    (before the extension) in the format given by FILENAME_DATE_FORMAT."""
    dir_name = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)
    name, ext = os.path.splitext(base_name)
    if min_date == max_date:
        new_name = f"{name} - {min_date.strftime(FILENAME_DATE_FORMAT)}{ext}"
    else:
        new_name = f"{name} - {min_date.strftime(FILENAME_DATE_FORMAT)} to {max_date.strftime(FILENAME_DATE_FORMAT)}{ext}"
    new_file_path = os.path.join(dir_name, new_name)
    return new_file_path


def subcommand_gdrive_weekly(args):
    logger = get_logger()

    if not config.get_nielsen_username() or not config.get_nielsen_password():
        logger.error("The Nielsen credentials are not set. Do you have a .env file?")
        return 1

    service = gdrive.build_drive_service(gdrive.load_credentials())
    gdrive_drive_id = config.get_gdrive_drive_id()
    in_folder_id = config.get_gdrive_folder_in_id()
    out_folder_id = config.get_gdrive_folder_out_id()
    # Create the output folder
    today = datetime.date.today()
    nielsen_week_ending = get_latest_available_nielsen_week_ending(today)
    folder_name = nielsen_week_ending.strftime(FILENAME_DATE_FORMAT)
    date_folder_id = gdrive.create_folder(
        service, out_folder_id, folder_name, only_if_not_exists=True, drive_id=gdrive_drive_id
    )
    if not date_folder_id:
        logger.error("Could not create dated output folder: '%s'", folder_name)
        return 2
    logs_folder_name = "logs"
    logs_folder_id = gdrive.create_folder(
        service, date_folder_id, logs_folder_name, only_if_not_exists=True, drive_id=gdrive_drive_id
    )
    if not logs_folder_id:
        logger.error(
            "Could not create dated output logs folder: '%s/%s'", folder_name, logs_folder_name
        )
        return 2

    # Get list of files in the output folder
    out_date_folder_files = gdrive.list_files(service, date_folder_id, drive_id=gdrive_drive_id)
    if out_date_folder_files:
        out_date_folder_files = [f["name"].rsplit("-", 1)[0].strip() for f in out_date_folder_files]

    # Get a temporary directory to use
    tmpdir = tempfile.mkdtemp("excel_nielsen_automate")

    previous_driver = None
    previous_excel_filepath = None
    driver = None
    excel_filepath = None
    appium_server = None
    wad_server = None
    failure_limit = 2
    failed_files = []
    try:
        wad_server = start_wad_server()
        appium_server = start_appium_server()
        for excel_file_entry in gdrive.list_files(service, in_folder_id, drive_id=gdrive_drive_id):
            # Skip if the file is already in the output folder
            if (
                out_date_folder_files
                and os.path.splitext(excel_file_entry["name"])[0] in out_date_folder_files
            ):
                logger.info(
                    "Skipping file '%s' because it is already in the output folder",
                    excel_file_entry["name"],
                )
                continue

            previous_excel_filepath = excel_filepath
            excel_filepath = os.path.join(tmpdir, excel_file_entry["name"])
            try:
                with open(excel_filepath, "wb") as f:
                    f.write(gdrive.download_file(service, excel_file_entry["id"]).getbuffer())
            except Exception as e:
                logger.exception(
                    "Error while downloading file '(%s)' from gdrive to tmp local path: %s",
                    excel_file_entry,
                    str(e),
                )
                continue

            previous_driver = driver
            try:
                driver = create_driver_for_file(excel_filepath)
                wait_until_file_saved_to_this_pc_visible(
                    driver, os.path.basename(excel_filepath), timeout=120
                )
                _update_sales(driver, excel_filepath)
                # On success, detect the min / max periods from the file and confirm it aligns
                # with the expected week ending date
                try:
                    min_period, max_period = (
                        excel_utils.get_min_max_nielsen_periods_from_excel_file(excel_filepath)
                    )
                except excel_utils.MissingPeriodsColumnError as e:
                    logger.warning(
                        "The 'Periods' column is missing from the file '%s'. Falling back to using "
                        "the expected week ending date %s. Detail: %s",
                        excel_filepath,
                        nielsen_week_ending,
                        str(e),
                    )
                    min_period = nielsen_week_ending
                    max_period = nielsen_week_ending
                if max_period != nielsen_week_ending:
                    logger.error(
                        "The max period date %s in file '%s' does not match the expected week ending date %s",
                        max_period,
                        excel_filepath,
                        nielsen_week_ending,
                    )
                    raise WeekEndingDateError(
                        excel_file_entry["name"],
                        nielsen_week_ending.strftime(FILENAME_DATE_FORMAT),
                        max_period.strftime(FILENAME_DATE_FORMAT),
                    )
            except Exception as e:
                logger.exception("Failed to refresh '%s': %s", excel_filepath, str(e))
                failed_files.append(excel_file_entry)
                # Take a screenshot of the error
                if driver:
                    try:
                        screenshot = _take_screenshot(driver, excel_file_entry["name"])
                        if screenshot:
                            # Upload the screenshot to the logs folder
                            gdrive.upload_file(
                                service,
                                screenshot,
                                logs_folder_id,
                                drive_id=gdrive_drive_id,
                            )
                    except Exception as e2:
                        logger.exception(
                            "Failed to take/upload screenshot for file '%s': %s",
                            excel_filepath,
                            str(e2),
                        )
                # Try to save the file so that we avoid the save prompt on exit
                try:
                    click_save_in_excel(driver, os.path.basename(excel_filepath), logger)
                except Exception as e2:
                    logger.exception(
                        "Failed to save file after error '%s': %s",
                        excel_filepath,
                        str(e2),
                    )
                # If the error was due to data plan limit, raise specific exception
                if isinstance(e, DataPlanLimitReachedError):
                    raise DataPlanLimitReachedError(failed_files) from e

                # If more than failure_limit failures, raise FailureLimitExceededError
                if len(failed_files) >= failure_limit:
                    raise FailureLimitExceededError(failed_files) from e
            else:
                # Rename the file to include the min/max periods
                excel_filepath_with_dates = get_filepath_with_dates(
                    excel_filepath, min_period, max_period
                )
                gdrive.upload_file(
                    service,
                    excel_filepath,
                    date_folder_id,
                    upload_filename=os.path.basename(excel_filepath_with_dates),
                    drive_id=gdrive_drive_id,
                )
            finally:
                if previous_driver:
                    # Best effort to quit the driver
                    try:
                        previous_driver.quit()
                    except Exception as e:
                        logger.exception("Failed to quit driver: %s", str(e))
                if previous_excel_filepath and os.path.exists(previous_excel_filepath):
                    try:
                        os.remove(previous_excel_filepath)
                    except Exception as e:
                        logger.exception("Failed to remove file: %s", str(e))
    finally:
        if driver:
            try:
                driver.quit()
            except Exception as e:
                logger.exception("Failed to quit driver: %s", str(e))
            driver = None
        if excel_filepath and os.path.exists(excel_filepath):
            try:
                os.remove(excel_filepath)
            except Exception as e:
                logger.exception("Failed to remove file: %s", str(e))
        if appium_server:
            stop_appium_server(appium_server)
            appium_server = None
        if wad_server:
            stop_wad_server(wad_server)
            wad_server = None

    # Attempt to upload main log file to gdrive logs folder
    try:
        main_log_filepath = log.get_main_log_filepath()
        if main_log_filepath and os.path.exists(main_log_filepath):
            gdrive.upload_file(service, main_log_filepath, logs_folder_id, drive_id=gdrive_drive_id)
    except Exception as e:
        logger.exception("Failed to upload main log file to gdrive: %s", str(e))

    return 0


def main():
    log.setup_logging()

    logger = get_logger()

    parser = get_argparser()
    args = parser.parse_args()

    try:
        match args.subcommand:
            case "files":
                result = subcommand_files(args)
            case "gdrive-weekly":
                result = subcommand_gdrive_weekly(args)
            case _default:
                raise ValueError(f"Unsupported subcommand: '{args.subcommand}'")
    except Exception as e:
        logger.exception("Error: %s", str(e))
        raise

    return result


if __name__ == "__main__":
    sys.exit(main())
