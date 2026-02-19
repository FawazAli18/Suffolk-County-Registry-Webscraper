# Suffolk-County-Registry-Webscraper

# 1. Project Overview

This automated tool is designed to extract daily real estate deed records from the Suffolk County (MassRODS) Registry of Deeds. The script navigates the Land Court search portal, filters for specific document types, extracts transaction data, and transmits a CSV report via the Microsoft Graph API.

# 2. Tech Stack

Language: Python 3.10+
- Automation Tool: Playwright
- API Integration: Microsoft Graph API via msal and requests
- Configuration: python-dotenv for environment variable management.

# 3. Data Extraction Strategy

| Data Point | Locator Type | Selector / Path |
| :--- | :--- | :--- |
| **Result Links** | CSS (Partial ID) | `a[id*='ButtonRow_Book/Page_']` |
| **Street Number** | XPath (Relative) | `//th[text()='Street #']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[1]` |
| **Street Name** | XPath (Relative) | `//th[text()='Street #']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[2]` |
| **Consideration** | XPath (Relative) | `//th[text()='Consideration']/ancestor::table[1]//tr[contains(@class,'DataGridRow')]/td[7]` |
| **Grantor Name** | XPath (Functional) | `//tr[td[2]='Grantor']/td[1]//a` |
| **Grantee Name** | XPath (Functional) | `//tr[td[2]='Grantee']/td[1]//a` |
| **Loading Overlay**| ID | `#ProgressBar1_UpdateProgress2` |
| **Next Page** | ID | `#DocList1_LinkButtonNext` |

# 4. Key Features:

- The script implements a validation loop that compares the target_id (from the results list) with the detail_id (rendered in the detail view) before extraction. This ensures that the data being scraped has fully refreshed after an AJAX update.
- Uses index-based querying (.nth(i)) instead of element handles to prevent "Element not attached to DOM" errors after page navigation or "go_back" actions.
- Integrates with Microsoft Graph API to transmit the CSV output directly from a shared mailbox.

# 5. Set up and Config

Installation:
1. Clone the repository:
   ```
   git clone [repo-url]
   ```
2. Install Dependencies:
   ```
   pip install -r requirements.txt
   playwright install chromium
   ```
3. Environment Variables create a .env file in the root directory with the following credentials:
   ```
   TENANT_ID=your_tenant_id
   CLIENT_ID=your_client_id
   CLIENT_SECRET=your_client_secret
   SENDER_EMAIL=sender email
   RECIPIENT_EMAIL= recipient email
   ```
