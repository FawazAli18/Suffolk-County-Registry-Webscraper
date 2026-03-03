import asyncio
import csv
from datetime import datetime, timedelta
from playwright.async_api import async_playwright
import msal
import requests
import base64
import os
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
RECIPIENT_EMAIL = os.getenv("RECIPIENT_EMAIL")


async def find_element(page, selectors: list, timeout: int = 10000):
    """
    Try multiple selectors in priority order.
    Returns the first visible locator found, or raises an
    Exception listing all selectors that were attempted.
    """
    per_selector_timeout = max(1000, timeout // len(selectors))

    for selector in selectors:
        try:
            locator = page.locator(selector).first
            await locator.wait_for(state="visible", timeout=per_selector_timeout)
            return locator
        except Exception:
            continue

    raise Exception(
        f"[find_element] Could not find a visible element using any of these selectors:\n"
        + "\n".join(f"  • {s}" for s in selectors)
    )


async def find_all(page, selectors: list, timeout: int = 10000):
    """
    Same idea as find_element but returns all matching elements
    for the first selector that produces at least one result.
    Useful for result-row links, etc.
    """
    per_selector_timeout = max(1000, timeout // len(selectors))

    for selector in selectors:
        try:
            locator = page.locator(selector)
            await locator.first.wait_for(state="visible", timeout=per_selector_timeout)
            return await locator.all()
        except Exception:
            continue

    raise Exception(
        f"[find_all] Could not find elements using any of these selectors:\n"
        + "\n".join(f"  • {s}" for s in selectors)
    )


async def click_element(page, selectors: list, timeout: int = 10000):
    """Convenience wrapper — finds then clicks."""
    element = await find_element(page, selectors, timeout)
    await element.click()


async def select_option_resilient(page, selectors: list, label: str = None, value: str = None, timeout: int = 10000):
    """Finds a <select> and selects by label or value."""
    element = await find_element(page, selectors, timeout)
    native = await element.element_handle()
    if label:
        await native.select_option(label=label)
    elif value:
        await native.select_option(value=value)


async def fill_field(page, selectors: list, text: str, timeout: int = 10000):
    """Finds an input, clears it, and types the given text."""
    element = await find_element(page, selectors, timeout)
    await element.click()
    await element.press("Control+A")
    await element.press("Backspace")
    await element.type(text)
    await element.press("Tab")


async def send_email_with_graph(file_path):
    print(f"Attempting to send email from {SENDER_EMAIL} to {RECIPIENT_EMAIL}...")

    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=authority,
    )

    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" in result:
        with open(file_path, "rb") as f:
            content_bytes = f.read()
        encoded_content = base64.b64encode(content_bytes).decode("utf-8")

        email_data = {
            "message": {
                "subject": f"Suffolk County Deeds Export - {datetime.now().strftime('%m/%d/%Y')}",
                "body": {
                    "contentType": "Text",
                    "content": "Automated single-day scrape complete. Please find the attached CSV for Suffolk County deeds.",
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
            print(" Email sent successfully via Graph API.")
        else:
            print(f" Failed to send email (HTTP {response.status_code}):")
            print(f"  Error Details: {response.text}")
    else:
        print(f"Token acquisition failed: {result.get('error_description')}")


async def run_scraper():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = await context.new_page()
        await page.goto("https://massrods.com/suffolk/")

        # Close pop-up
        try:
            await click_element(page, [
                ".pum-close",                        # original class
                "button[aria-label='Close']",        # aria label fallback
                "button:has-text('Close')",           # visible text fallback
                "[class*='close']",                   # any class containing "close"
            ], timeout=5000)
            print("Pop-up closed.")
        except Exception:
            print("Pop-up not found, continuing...")

        # Open Document Search popup
        async with page.expect_popup() as popup_info:
            await click_element(page, [
                "text=Document Search",              # visible text (primary)
                "a:has-text('Document Search')",     # link with that text
                "button:has-text('Document Search')",
            ])
        search_page = await popup_info.value
        await search_page.bring_to_front()

        # Office Name dropdown
        await select_option_resilient(search_page, [
            "#SearchCriteriaOffice1_DDL_OfficeName",   # original ID
            "select[id*='DDL_OfficeName']",             # partial ID
            "select[id*='OfficeName']",                 # even more partial
        ], label="Registered Land (Land Court)")
        await search_page.wait_for_load_state("networkidle")
        await asyncio.sleep(1)

        # Search type dropdown 
        await select_option_resilient(search_page, [
            "#SearchCriteriaName1_DDL_SearchName",     # original ID
            "select[id*='DDL_SearchName']",             # partial ID
            "select[id*='SearchName']",                 # even more partial
        ], label="Recorded Date Search")
        await search_page.wait_for_load_state("networkidle")
        await asyncio.sleep(1)

        # Advanced button 
        await click_element(search_page, [
            "#SearchFormEx1_BtnAdvanced",              # original ID
            "input[id*='BtnAdvanced']",                # partial ID
            "button:has-text('Advanced')",             # visible text
            "input[value*='Advanced']",                # input button by value
        ])

        # Wait for date fields to appear 
        await find_element(search_page, [
            "#SearchFormEx1_DRACSTextBox_DateFrom",    # original ID
            "input[id*='DateFrom']",                   # partial ID
            "//label[contains(text(),'From')]/following-sibling::input[1]",  # XPath by label
        ])
        await asyncio.sleep(1)

        # Build date strings
        target_date = datetime.today()
        date_from = target_date.strftime("%m/%d/%Y")
        date_to = target_date.strftime("%m/%d/%Y")
        print(f"Applying strict single-day filter: {date_from}")

        # Fill Date From
        await fill_field(search_page, [
            "#SearchFormEx1_DRACSTextBox_DateFrom",
            "input[id*='DateFrom']",
            "//label[contains(text(),'From')]/following-sibling::input[1]",
        ], date_from)

        # Fill Date To 
        await fill_field(search_page, [
            "#SearchFormEx1_DRACSTextBox_DateTo",
            "input[id*='DateTo']",
            "//label[contains(text(),'To')]/following-sibling::input[1]",
        ], date_to)

        # Document Type dropdown
        await select_option_resilient(search_page, [
            "#SearchFormEx1_ACSDropDownList_DocumentType",  # original ID
            "select[id*='DocumentType']",                    # partial ID
            "select[id*='DocType']",                         # alt naming
        ], value="100056")

        # Search button
        await click_element(search_page, [
            "#SearchFormEx1_btnSearch",                # original ID
            "input[id*='btnSearch']",                  # partial ID
            "button:has-text('Search')",               # visible text
            "input[value='Search']",                   # input button by value
        ])
        await search_page.wait_for_load_state("networkidle")

        # Confirm results loaded
        try:
            await find_element(search_page, [
                "a[id*='ButtonRow_Book/Page_']",         
                "//table[contains(@id,'DocList')]//a[contains(@id,'Book')]",
                "table.SearchResultsGrid a",
            ], timeout=10000)
            print("Results loaded successfully.")
        except Exception:
            print("No results found for this date or timeout occurred.")
            await browser.close()
            return

        csv_filename = "suffolk_county_deeds.csv"
        csv_file = open(csv_filename, "w", newline="", encoding="utf-8")
        writer = csv.writer(csv_file)
        writer.writerow(["Consideration", "Grantor", "Grantee", "Street #", "Street Name"])

        last_address = ""

        while True:
            # Re-query result links on every page using fallback selectors
            links = await find_all(search_page, [
                "a[id*='ButtonRow_Book/Page_']",            
                "//table[contains(@id,'DocList')]//a[contains(@id,'Book')]",
                "table.SearchResultsGrid a",
            ])
            print(f"Found {len(links)} entries on this page.")

            for i in range(len(links)):
                try:
                    target_link = search_page.locator("a[id*='ButtonRow_Book/Page_']").nth(i)
                    target_id = await target_link.inner_text()

                    try:
                        await find_element(search_page, [
                            "#ProgressBar1_UpdateProgress2",
                            "[id*='UpdateProgress']",
                            "[id*='ProgressBar']",
                        ], timeout=10000)
                        await search_page.locator(
                            "#ProgressBar1_UpdateProgress2, [id*='UpdateProgress']"
                        ).first.wait_for(state="hidden", timeout=10000)
                    except Exception:
                        pass

                    await target_link.click()
                    await search_page.wait_for_selector(
                        f"td:has-text('{target_id}')", state="visible", timeout=10000
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
                    grantors = await search_page.locator("//tr[td[2]='Grantor']/td[1]//a").all_inner_texts()
                    grantees = await search_page.locator("//tr[td[2]='Grantee']/td[1]//a").all_inner_texts()

                    writer.writerow([consideration, ", ".join(grantors), ", ".join(grantees), street_no, street_name])
                    csv_file.flush()

                    print(f"[{i+1}/{len(links)}] Saved ID {target_id}: {street_no} {street_name}")
                    last_address = f"{street_no} {street_name}"

                except Exception as e:
                    print(f"Error extracting entry {i+1}: {e}")

                await search_page.go_back()
                await search_page.wait_for_selector("a[id*='ButtonRow_Book/Page_']", state="visible")

                try:
                    await search_page.wait_for_selector(
                        "#ProgressBar1_UpdateProgress2", state="hidden", timeout=5000
                    )
                except Exception:
                    pass

                await asyncio.sleep(1)

            try:
                next_btn = await find_element(search_page, [
                    "#DocList1_LinkButtonNext",            # original ID
                    "a[id*='LinkButtonNext']",             # partial ID
                    "a:has-text('Next')",                  # visible text
                    "input[value='Next']",                 # input button
                ], timeout=3000)
                print("Moving to next page...")
                await next_btn.click()
                await search_page.wait_for_load_state("networkidle")
            except Exception:
                print("No more pages.")
                break

        csv_file.close()
        print(f"Scrape complete. Data saved to {csv_filename}")

        await send_email_with_graph(csv_filename)
        await browser.close()


if __name__ == "__main__":
    asyncio.run(run_scraper())
