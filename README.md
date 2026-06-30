# PI / Quotation Tracker — Team Web App

A multi-user web app for tracking Proforma Invoices (PI) and Quotations,
with customer/item bulk import, GST, freight charges, and follow-up reminders.

---

## What's inside

- Login system — each team member has their own username & password
- **Roles:**
  - `admin` — sees ALL customers, items, and PI/Quotes from every user, manages team accounts
  - `staff` — only sees customers and PI/Quotes they personally created
- Quotation numbers auto-start at **260001** and auto-increment
- Date is always today's date automatically
- Dispatch From dropdown: UNIT 1 / DABASPETE / MAHIMAPURA / DADRA
- UOM dropdown: Nos / KG / Pieces / BAG / TON / BOX
- GST toggle (18%) calculated automatically
- Freight charges field added on top of GST, before the grand total
- Bulk import customers and items via CSV or Excel (.xlsx)
- WhatsApp reminder links per customer

---

## Step 1 — Run it locally first (to test on your own computer)

You need Python 3.10+ installed.

```bash
cd pi-tracker
pip install -r requirements.txt
python app.py
```

Open **http://localhost:5000** in your browser.

Default login (created automatically):
- **Username:** `admin`
- **Password:** `admin123`

⚠️ Log in and change this password immediately by creating new admin
credentials via "Manage Users" once you're live, or update the database directly.

---

## Step 2 — Deploy online so your team can access it from anywhere

The easiest free option is **Render.com**. Steps:

### A. Push this project to GitHub
1. Create a free GitHub account if you don't have one: https://github.com
2. Create a new repository (e.g. `pi-tracker`)
3. Upload all the files in this folder to that repository
   (you can drag-and-drop the files into GitHub's web interface — no command line needed)

### B. Create the database on Render
1. Go to https://render.com and sign up free (no credit card needed for free tier)
2. Click **New +** → **PostgreSQL** → choose the Free plan → Create Database
3. Once created, copy the **Internal Database URL** shown on the database page

### C. Deploy the web app
1. Click **New +** → **Web Service**
2. Connect your GitHub account and select the `pi-tracker` repository
3. Fill in:
   - **Name:** pi-tracker (or anything you like)
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
4. Under **Environment Variables**, add:
   - `DATABASE_URL` = (paste the Internal Database URL from step B)
   - `SECRET_KEY` = (type any long random text, e.g. `myCompany2024SecretKey987`)
5. Click **Create Web Service**

Render will build and deploy automatically. After a few minutes you'll get a
public link like:

```
https://pi-tracker.onrender.com
```

This is the link your whole team will use to log in — works on phone,
laptop, any browser, anywhere.

### D. Create your team's logins
1. Open the live link and log in as `admin` / `admin123`
2. **First, change the admin password** — go to Manage Users → Reset PW on the
   `admin` account
3. Go to **Manage Users → + Add user** for each team member
   - Give each person a username + password
   - Choose role: `admin` (sees everything) or `staff` (sees only their own work)
4. Share the link + their individual username/password with each team member

That's it — everyone logs in with their own account from the same shared link.

---

## Notes

- Free Render web services "sleep" after 15 minutes of no traffic and take
  ~30 seconds to wake up on the next visit. If that's a problem, upgrade to
  Render's paid Starter plan (~$7/month) for an always-on app.
- The free Render PostgreSQL database expires after 90 days, after which
  you'll need to create a new one (or upgrade to a paid database) to keep
  your data — back this up periodically by exporting your customers/items.
- Bulk import templates are downloadable from the Customers and Items pages
  inside the app — use those exact column headers when preparing your own
  spreadsheet.
