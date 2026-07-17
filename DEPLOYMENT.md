# Deployment guide — Wan Boutique Order Manager

## Part 1: Prepare Google Sheets

1. In Google Drive, create a new blank Google Sheet named `Wan Boutique CRM`.
2. Open **Extensions → Apps Script**.
3. Replace the default `Code.gs` content with [`google-appscript/Code.gs`](google-appscript/Code.gs).
4. In Apps Script, open **Project Settings** and set the timezone to `Asia/Kuala_Lumpur` if it is not already selected.
5. Save the project.
6. In the function dropdown, select `setupDatabase` and click **Run**.
7. Approve the Google permission prompts. The script will create `Orders`, `OrderItems`, `Customers`, `Payments`, `Settings`, `AuditLog`, and `LegacyImport` with correct headers.

Do not rename the created sheet tabs or their row-one headers.

### Optional: import the old Excel history

Keep the original workbook as a historical record. If you import its `PS` data into the `LegacyImport` tab, map these columns manually:

```text
No → Legacy Row
ID → Legacy ID
Date → Date
Name → Name
Contact No → Contact No
Brand → Brand
Item → Item
Size → Size
Remark → Remark
```

The historical sheet has one item per row, so do not import it directly into `Orders`.

## Part 2: Deploy the Apps Script API

1. In Apps Script, click **Deploy → New deployment**.
2. Select **Web app**.
3. Set **Execute as** to **Me**.
4. Set access to the narrowest option suitable for the owner. For a single owner, choose the owner’s Google account access where available. If the frontend is hosted publicly on GitHub Pages and needs direct access, select the appropriate web-app access option and protect the URL operationally.
5. Click **Deploy**, approve permissions, and copy the URL ending with `/exec`.
6. Do not use the `/dev` URL in the public frontend; it is only for development editors.

## Part 3: Configure and publish GitHub Pages

1. Create a new GitHub repository, for example `wan-boutique-order-manager`.
2. Upload all project files except private notes and the legacy workbook if you do not want customer data in GitHub.
3. Commit and push to the `main` branch.
4. In GitHub, open **Settings → Pages**.
5. Under **Build and deployment**, choose **Deploy from a branch**.
6. Select branch `main`, folder `/ (root)`, then click **Save**.
7. Wait for GitHub Pages to publish. GitHub gives you a URL such as `https://your-name.github.io/wan-boutique-order-manager/`.

## Part 4: Connect the frontend

1. Open the GitHub Pages URL.
2. Open **Settings** in the app.
3. Paste the Apps Script `/exec` URL into **Google Apps Script Web App URL**.
4. Set the business name and invoice prefix, then click **Save settings**.
5. Create a test manual order, save it, and confirm that a row appears in `Orders` and one or more rows appear in `OrderItems`.

The API URL is stored in that browser. Each owner device must enter it once. If preferred, place the value in `CONFIG.API_URL` in `app.js` before uploading to GitHub.

## Part 5: Daily use

1. Ask customers to use the template in **Settings → Customer copy template**.
2. Paste WhatsApp text into **New order → Paste WhatsApp**.
3. Review the parser output.
4. Enter the manual item cost and service fee for every item.
5. Save the order.
6. Record payments in **Payments**.
7. Open an invoice from **Orders** and use the browser print dialog to save or print it.

## Pricing rule

For the example `Qty 3`, `Item Price RM100`, `Service Fee RM50`, and total `RM150`, choose **Total item cost** and enter RM100. Choose **Unit price × qty** only when RM100 is the price of each item.

## Security note

Do not put a password, spreadsheet ID, or administrative secret in the GitHub Pages source. A public static site cannot keep a frontend secret. Use restricted Apps Script deployment access where practical, and only share the deployed app with the owner/staff who should access customer information.
