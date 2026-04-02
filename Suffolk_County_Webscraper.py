import argparse
import asyncio
import csv
import logging
import os
import base64
import requests
import msal
from datetime import datetime, timedelta
from playwright.async_api import async_playwright
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
RECIPIENT_EMAIL = os.getenv("RECIPIENT_EMAIL")


def setup_logging() -> logging.Logger:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(script_dir, "logs")

    try:
        os.makedirs(log_dir, exist_ok=True)
    except OSError as e:
        print(f"[LOGGING ERROR] Could not create log directory '{log_dir}': {e}")
        raise

    log_filename = os.path.join(
        log_dir, f"suffolk_scraper_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )

    logger = logging.getLogger("SuffolkScraper")
    logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    try:
        file_handler = logging.FileHandler(log_filename, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except OSError as e:
        print(f"[LOGGING ERROR] Could not create log file '{log_filename}': {e}")
        raise

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    logger.info(f"Log file: {log_filename}")
    return logger



def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Suffolk County deed webscraper",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--days",
        type=int,
        default=1,
        help=(
            "How many calendar days back to search from today. "
            "Use 1 to get yesterday's deeds, "
            "7 to look back a full week, etc."
        ),
    )
    return parser.parse_args()


async def find_element(page, selectors: list, timeout: int = 10000):
    per_selector_timeout = max(1000, timeout // len(selectors))
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            await locator.wait_for(state="visible", timeout=per_selector_timeout)
            return locator
        except Exception:
            continue
    raise Exception(
        "[find_element] Could not find a visible element using any of these selectors:\n"
        + "\n".join(f"  • {s}" for s in selectors)
    )


async def find_all(page, selectors: list, timeout: int = 10000):
    per_selector_timeout = max(1000, timeout // len(selectors))
    for selector in selectors:
        try:
            locator = page.locator(selector)
            await locator.first.wait_for(state="visible", timeout=per_selector_timeout)
            return await locator.all()
        except Exception:
            continue
    raise Exception(
        "[find_all] Could not find elements using any of these selectors:\n"
        + "\n".join(f"  • {s}" for s in selectors)
    )


async def click_element(page, selectors: list, timeout: int = 10000):
    element = await find_element(page, selectors, timeout)
    await element.click()


async def select_option_resilient(
    page, selectors: list, label: str = None, value: str = None, timeout: int = 10000
):
    """Finds a <select> and selects by label or value."""
    element = await find_element(page, selectors, timeout)
    native = await element.element_handle()
    if label:
        await native.select_option(label=label)
    elif value:
        await native.select_option(value=value)


async def fill_field(page, selectors: list, text: str, timeout: int = 10000):
    element = await find_element(page, selectors, timeout)
    await element.click()
    await element.press("Control+A")
    await element.press("Backspace")
    await element.type(text)
    await element.press("Tab")

async def send_email_with_graph(file_path: str, logger: logging.Logger):
    logger.info(f"Attempting to send email from {SENDER_EMAIL} to {RECIPIENT_EMAIL}...")

    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=authority,
    )

    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" in result:
        with open(file_path, "rb") as f:
            content_bytes = f.read()
        encoded_content = base64.b64encode(content_bytes).decode("utf-8")

        email_data = {
            "message": {
                "subject": f"Suffolk County Deeds Export - {datetime.now().strftime('%m/%d/%Y')}",
                "body": {
                    "contentType": "Text",
                    "content": (
                        "Automated scrape complete. "
                        "Please find the attached CSV for Suffolk County deeds."
                    ),
                },
                "toRecipients": [{"emailAddress": {"address": RECIPIENT_EMAIL}}],
                "attachments": [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": os.path.basename(file_path),
                        "contentBytes": encoded_content,
                    }
                ],
            },
            "saveToSentItems": "true",
        }

        endpoint = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"
        response = requests.post(
            endpoint,
            headers={"Authorization": "Bearer " + result["access_token"]},
            json=email_data,
        )

        if response.status_code == 202:
            logger.info("Email sent successfully via Graph API.")
        else:
            logger.error(
                f"Failed to send email (HTTP {response.status_code}): {response.text}"
            )
    else:
        logger.error(
            f"Token acquisition failed: {result.get('error_description')}"
        )


async def run_scraper(days_back: int, logger: logging.Logger):
    today = datetime.today()
    date_to_dt = today - timedelta(days=1)        # yesterday (last completed day)
    date_from_dt = today - timedelta(days=days_back)

    date_from = date_from_dt.strftime("%m/%d/%Y")
    date_to = date_to_dt.strftime("%m/%d/%Y")

    logger.info(
        f"Date range: {date_from} → {date_to}  "
        f"(lookback = {days_back} day{'s' if days_back != 1 else ''})"
    )

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
        )
        page = await context.new_page()
        await page.goto("https://massrods.com/suffolk/")

        # Close pop-up
        try:
            await click_element(
                page,
                [
                    ".pum-close",
                    "button[aria-label='Close']",
                    "button:has-text('Close')",
                    "[class*='close']",
                ],
                timeout=5000,
            )
            logger.info("Pop-up closed.")
        except Exception:
            logger.debug("Pop-up not found, continuing...")

        # Open Document Search popup
        async with page.expect_popup() as popup_info:
            await click_element(
                page,
                [
                    "text=Document Search",
                    "a:has-text('Document Search')",
                    "button:has-text('Document Search')",
                ],
            )
        search_page = await popup_info.value
        await search_page.bring_to_front()

        # Office Name dropdown
        await select_option_resilient(
            search_page,
            [
                "#SearchCriteriaOffice1_DDL_OfficeName",
                "select[id*='DDL_OfficeName']",
                "select[id*='OfficeName']",
            ],
            label="Registered Land (Land Court)",
        )
        await search_page.wait_for_load_state("networkidle")
        await asyncio.sleep(1)

        # Search type dropdown
        await select_option_resilient(
            search_page,
            [
                "#SearchCriteriaName1_DDL_SearchName",
                "select[id*='DDL_SearchName']",
                "select[id*='SearchName']",
            ],
            label="Recorded Date Search",
        )
        await search_page.wait_for_load_state("networkidle")
        await asyncio.sleep(1)

        # Advanced button
        await click_element(
            search_page,
            [
                "#SearchFormEx1_BtnAdvanced",
                "input[id*='BtnAdvanced']",
                "button:has-text('Advanced')",
                "input[value*='Advanced']",
            ],
        )

        # Wait for date fields to appear
        await find_element(
            search_page,
            [
                "#SearchFormEx1_DRACSTextBox_DateFrom",
                "input[id*='DateFrom']",
                "//label[contains(text(),'From')]/following-sibling::input[1]",
            ],
        )
        await asyncio.sleep(1)

        # Fill Date From
        await fill_field(
            search_page,
            [
                "#SearchFormEx1_DRACSTextBox_DateFrom",
                "input[id*='DateFrom']",
                "//label[contains(text(),'From')]/following-sibling::input[1]",
            ],
            date_from,
        )

        # Fill Date To
        await fill_field(
            search_page,
            [
                "#SearchFormEx1_DRACSTextBox_DateTo",
                "input[id*='DateTo']",
                "//label[contains(text(),'To')]/following-sibling::input[1]",
            ],
            date_to,
        )

        # Document Type dropdown
        await select_option_resilient(
            search_page,
            [
                "#SearchFormEx1_ACSDropDownList_DocumentType",
                "select[id*='DocumentType']",
                "select[id*='DocType']",
            ],
            value="100056",
        )

        # Search button
        await click_element(
            search_page,
            [
                "#SearchFormEx1_btnSearch",
                "input[id*='btnSearch']",
                "button:has-text('Search')",
                "input[value='Search']",
            ],
        )
        await search_page.wait_for_load_state("networkidle")

       
        try:
            await search_page.locator(
                "#ProgressBar1_UpdateProgress2, [id*='UpdateProgress'], [id*='ProgressBar']"
            ).first.wait_for(state="hidden", timeout=15000)
        except Exception:
            pass  
        await asyncio.sleep(2) 

        try:
            await find_element(
                search_page,
                [
                    "a[id*='ButtonRow_Book/Page_']",
                    "//table[contains(@id,'DocList')]//a[contains(@id,'Book')]",
                    "table.SearchResultsGrid a",
                ],
                timeout=20000,  
            )
            logger.info("Results loaded successfully.")
        except Exception:
            logger.warning("No results found for this date range or timeout occurred.")
            await browser.close()
            return

        csv_filename = (
            f"suffolk_county_deeds_{date_from_dt.strftime('%Y%m%d')}"
            f"_to_{date_to_dt.strftime('%Y%m%d')}.csv"
        )
        csv_file = open(csv_filename, "w", newline="", encoding="utf-8")
        writer = csv.writer(csv_file)
        writer.writerow(["Consideration", "Grantor", "Grantee", "Street #", "Street Name"])

        last_address = ""
        total_saved = 0
        total_errors = 0

        while True:
            links = await find_all(
                search_page,
                [
                    "a[id*='ButtonRow_Book/Page_']",
                    "//table[contains(@id,'DocList')]//a[contains(@id,'Book')]",
                    "table.SearchResultsGrid a",
                ],
            )
            logger.info(f"Found {len(links)} entries on this page.")

            for i in range(len(links)):
                try:
                    target_link = search_page.locator(
                        "a[id*='ButtonRow_Book/Page_']"
                    ).nth(i)
                    target_id = await target_link.inner_text()

                    try:
                        await find_element(
                            search_page,
                            [
                                "#ProgressBar1_UpdateProgress2",
                                "[id*='UpdateProgress']",
                                "[id*='ProgressBar']",
                            ],
                            timeout=10000,
                        )
                        await search_page.locator(
                            "#ProgressBar1_UpdateProgress2, [id*='UpdateProgress']"
                        ).first.wait_for(state="hidden", timeout=10000)
                    except Exception:
                        pass

                    await target_link.click()
                    await search_page.wait_for_selector(
                        f"td:has-text('{target_id}')",
                        state="visible",
                        timeout=10000,
                    )

                    # Verify detail loaded
                    detail_id = await search_page.locator(
                        "//th[text()='Book/Page']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[1]"
                    ).first.inner_text()
                    if target_id not in detail_id:
                        await asyncio.sleep(1.5)

                    # Extract address
                    street_no = (
                        await search_page.locator(
                            "//th[text()='Street #']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[1]"
                        ).first.inner_text()
                    ).strip()
                    street_name = (
                        await search_page.locator(
                            "//th[text()='Street #']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[2]"
                        ).first.inner_text()
                    ).strip()

                    if f"{street_no} {street_name}" == last_address:
                        await asyncio.sleep(1)
                        street_no = (
                            await search_page.locator(
                                "//th[text()='Street #']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[1]"
                            ).first.inner_text()
                        ).strip()
                        street_name = (
                            await search_page.locator(
                                "//th[text()='Street #']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[2]"
                            ).first.inner_text()
                        ).strip()

                    # Extract consideration / parties
                    consideration = (
                        await search_page.locator(
                            "//th[text()='Consideration']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[7]"
                        ).first.inner_text()
                    ).strip()
                    grantors = await search_page.locator(
                        "//tr[td[2]='Grantor']/td[1]//a"
                    ).all_inner_texts()
                    grantees = await search_page.locator(
                        "//tr[td[2]='Grantee']/td[1]//a"
                    ).all_inner_texts()

                    writer.writerow(
                        [
                            consideration,
                            ", ".join(grantors),
                            ", ".join(grantees),
                            street_no,
                            street_name,
                        ]
                    )
                    csv_file.flush()
                    total_saved += 1

                    logger.info(
                        f"[{i + 1}/{len(links)}] Saved ID {target_id}: "
                        f"{street_no} {street_name}"
                    )
                    last_address = f"{street_no} {street_name}"

                except Exception as e:
                    total_errors += 1
                    logger.error(f"Error extracting entry {i + 1}: {e}")

                await search_page.go_back()
                await search_page.wait_for_selector(
                    "a[id*='ButtonRow_Book/Page_']", state="visible"
                )

                try:
                    await search_page.wait_for_selector(
                        "#ProgressBar1_UpdateProgress2",
                        state="hidden",
                        timeout=5000,
                    )
                except Exception:
                    pass

                await asyncio.sleep(1)

            try:
                next_btn = await find_element(
                    search_page,
                    [
                        "#DocList1_LinkButtonNext",
                        "a[id*='LinkButtonNext']",
                        "a:has-text('Next')",
                        "input[value='Next']",
                    ],
                    timeout=3000,
                )
                logger.info("Moving to next page...")
                await next_btn.click()
                await search_page.wait_for_load_state("networkidle")
            except Exception:
                logger.info("No more pages.")
                break

        csv_file.close()
        logger.info(
            f"Scrape complete — {total_saved} records saved, "
            f"{total_errors} errors. Output: {csv_filename}"
        )

        await send_email_with_graph(csv_filename, logger)
        await browser.close()

if __name__ == "__main__":
    args = parse_args()
    logger = setup_logging()
    logger.info(f"Starting Suffolk County scraper (--days {args.days})")
    asyncio.run(run_scraper(days_back=args.days, logger=logger))
