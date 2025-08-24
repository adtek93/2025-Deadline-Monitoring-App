# 🚀 Google Apps Script + Google Sheets Web App

## 1. Introduction
Google Apps Script (GAS) is a JavaScript-based cloud development platform by Google.  
You can use GAS to **automate, connect, and extend** Google Sheets, Docs, Gmail, and many other Google services.  

This repo demonstrates how to build a **web application from Google Sheets** and **schedule alerts to Telegram**.

---

## 2. Create Apps Script from Google Sheets
- Open the Google Sheets file where you want to attach the script.  
- From the menu, select:  
Extensions → Apps Script
- A new window (**Apps Script Editor**) will open.  
- Default file: `Code.gs`  
- You can also add `index.html` for a custom web interface.  

---

## 3. Write Code
- Place your main logic in **Code.gs**.  
- Add your UI (if any) in **index.html**.  
- Manage code directly in the Apps Script editor or sync with GitHub.  

---

## 4. Publish Web App
1. In Apps Script editor:  

Deploy → New deployment

2. Select **Web app**.  
3. Configure:  
- **Execute as**: `Me` (runs under your account)  
- **Who has access**: `Anyone` (anyone with the link) or restricted to your domain  
4. Click **Deploy** → copy the generated URL.  

👉 You now have a **web app powered by Google Sheets** 🎉  

---

## 5. Set Up Triggers (Scheduled Execution)
1. In Apps Script editor:  

Triggers → Add Trigger

2. Configure:  
- **Function**: `checkDeadlines_3` (or any function you want to schedule)  
- **Deployment**: Head  
- **Event source**: Time-driven  
- **Type of time-based trigger**: choose frequency (e.g., every day at 8am)  

👉 The script will automatically run on schedule, e.g., send alerts via Telegram.  

---

## 6. Result
- A web app that displays or interacts with Google Sheets data.  
- Automated scheduled alerts (e.g., reminders sent to Telegram).

<img width="1526" height="536" alt="image" src="https://github.com/user-attachments/assets/05f37c94-1384-46fe-863e-dbd94c3983b7" />
<img width="1516" height="810" alt="image" src="https://github.com/user-attachments/assets/68aa70cb-a71f-4a07-84e5-930bfe628c26" />
