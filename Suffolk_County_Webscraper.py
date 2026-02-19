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

async def send_email_with_graph(file_path):
    print(f"Attempting to send email from {SENDER_EMAIL} to {RECIPIENT_EMAIL}...")
    
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET, 
        authority=authority
    )
    
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    
    if "access_token" in result:
        with open(file_path, "rb") as f:
            content_bytes = f.read()
        encoded_content = base64.b64encode(content_bytes).decode('utf-8')

        email_data = {
            "message": {
                "subject": f"Suffolk County Deeds Export - {datetime.now().strftime('%m/%d/%Y')}",
                "body": {
                    "contentType": "Text", 
                    "content": "Automated single-day scrape complete. Please find the attached CSV for Suffolk County deeds."
                },
                "toRecipients": [{"emailAddress": {"address": RECIPIENT_EMAIL}}],
                "attachments": [{
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": os.path.basename(file_path),
                    "contentBytes": encoded_content
                }]
            },
            "saveToSentItems": "true" 
        }

        endpoint = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"
        
        response = requests.post(
            endpoint, 
            headers={'Authorization': 'Bearer ' + result['access_token']}, 
            json=email_data
        )
        
        if response.status_code == 202:
            print(" Email sent successfully via Graph API.")
        else:
            print(f" Failed to send email (HTTP {response.status_code}):")
            print(f"  Error Details: {response.text}")
    else:
        print(f" Token acquisition failed: {result.get('error_description')}")

async def run_scraper():
    async with async_playwright() as p:
        # Headless=False for easier monitoring during testing
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = await context.new_page()

        await page.goto('https://massrods.com/suffolk/')
        
        try:
            await page.click(".pum-close", timeout=5000)
            print("Pop-up closed.")
        except:
            print("Pop-up not found, continuing...")

        async with page.expect_popup() as popup_info:
            await page.click("text=Document Search")
        search_page = await popup_info.value
        await search_page.bring_to_front()

        await search_page.select_option("#SearchCriteriaOffice1_DDL_OfficeName", label="Registered Land (Land Court)")
        await search_page.wait_for_load_state("networkidle")
        await asyncio.sleep(1) 
        await search_page.select_option("#SearchCriteriaName1_DDL_SearchName", label="Recorded Date Search")
        await search_page.wait_for_load_state("networkidle")
        await asyncio.sleep(1)
        
        await search_page.click("#SearchFormEx1_BtnAdvanced")
        await search_page.wait_for_selector("#SearchFormEx1_DRACSTextBox_DateFrom", state="visible")
        
        await asyncio.sleep(1) 
        
        today = datetime.today()
        date_to = today.strftime('%m/%d/%Y')
        date_from = (today - timedelta(days=0)).strftime('%m/%d/%Y')
        
        print(f"Applying strict single-day filter: {date_from}")

        await search_page.click("#SearchFormEx1_DRACSTextBox_DateFrom")
        await search_page.keyboard.press("Control+A")
        await search_page.keyboard.press("Backspace")
        await search_page.type("#SearchFormEx1_DRACSTextBox_DateFrom", date_from)
        await search_page.keyboard.press("Tab")
        
        await search_page.click("#SearchFormEx1_DRACSTextBox_DateTo")
        await search_page.keyboard.press("Control+A")
        await search_page.keyboard.press("Backspace")
        await search_page.type("#SearchFormEx1_DRACSTextBox_DateTo", date_to)
        await search_page.keyboard.press("Tab")

        await search_page.select_option("#SearchFormEx1_ACSDropDownList_DocumentType", value="100056")
        
        await search_page.click("#SearchFormEx1_btnSearch")
        await search_page.wait_for_load_state("networkidle")

        try:
            await search_page.wait_for_selector("a[id*='ButtonRow_Book/Page_']", timeout=10000)
            print("Results loaded successfully.")
        except:
            print("No results found for this date or timeout occurred.")
            await browser.close()
            return

        csv_filename = 'suffolk_county_deeds.csv'
        csv_file = open(csv_filename, 'w', newline='', encoding='utf-8')
        writer = csv.writer(csv_file)
        writer.writerow(['Consideration', 'Grantor', 'Grantee', 'Street #', 'Street Name'])

        last_address = "" 
        while True:
            links = await search_page.query_selector_all("a[id*='ButtonRow_Book/Page_']")
            print(f"Found {len(links)} entries on this page.")
            
            for i in range(len(links)):
                try:
                   
                    target_link = search_page.locator("a[id*='ButtonRow_Book/Page_']").nth(i)
                    target_id = await target_link.inner_text()
                    
                    await search_page.wait_for_selector("#ProgressBar1_UpdateProgress2", state="hidden", timeout=10000)

                    await target_link.click()
                       
                    await search_page.wait_for_selector(f"td:has-text('{target_id}')", state="visible", timeout=10000)

                    detail_id = await search_page.locator("//th[text()='Book/Page']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[1]").first.inner_text()
                    if target_id not in detail_id:
                        await asyncio.sleep(1.5)
                    
                    street_no = (await search_page.locator("//th[text()='Street #']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[1]").first.inner_text()).strip()
                    street_name = (await search_page.locator("//th[text()='Street #']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[2]").first.inner_text()).strip()
                    
                    if f"{street_no} {street_name}" == last_address:
                        await asyncio.sleep(1) 
                        street_no = (await search_page.locator("//th[text()='Street #']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[1]").first.inner_text()).strip()
                        street_name = (await search_page.locator("//th[text()='Street #']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[2]").first.inner_text()).strip()

                    consideration = (await search_page.locator("//th[text()='Consideration']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[7]").first.inner_text()).strip()
                    grantors = await search_page.locator("//tr[td[2]='Grantor']/td[1]//a").all_inner_texts()
                    grantees = await search_page.locator("//tr[td[2]='Grantee']/td[1]//a").all_inner_texts()

                    writer.writerow([consideration, ', '.join(grantors), ', '.join(grantees), street_no, street_name])
                    csv_file.flush()
                    
                    print(f"[{i+1}/{len(links)}] Saved ID {target_id}: {street_no} {street_name}")
                    last_address = f"{street_no} {street_name}"
                    
                except Exception as e:
                    print(f"Error extracting entry {i+1}: {e}")
                
                await search_page.go_back()
                
                await search_page.wait_for_selector("a[id*='ButtonRow_Book/Page_']", state="visible")
                
                try:
                    await search_page.wait_for_selector("#ProgressBar1_UpdateProgress2", state="hidden", timeout=5000)
                except:
                    pass
                    
                await asyncio.sleep(1)

            next_btn = await search_page.query_selector("#DocList1_LinkButtonNext")
            if next_btn:
                print("Moving to next page...")
                await next_btn.click()
                await search_page.wait_for_load_state("networkidle")
            else:
                print("No more pages.")
                break

        csv_file.close()
        print(f"Scrape complete. Data saved to {csv_filename}")
        
        await send_email_with_graph(csv_filename)
        await browser.close()

if __name__ == "__main__":

    asyncio.run(run_scraper())
