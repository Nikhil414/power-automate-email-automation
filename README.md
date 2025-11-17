# 📧 Step-by-Step: Send Emails with Custom Interval per Row (Power Automate + Excel + Outlook Web)

Automate personalized email sending with Power Automate, Excel, and Outlook while controlling the interval between each email.

---

## 1. Prepare Your Excel File

- Fill in all important data (e.g., Name, Country, Email, etc.)
- **Format as Table:**
  - Select your entire data range
  - Go to **Insert** → **Table**
- Save the file to **OneDrive for Business**

---

## 2. Start Your Power Automate Flow

- Open [Power Automate](https://flow.microsoft.com)
- **Create a New Flow:**
  - Choose **Instant cloud flow**
  - Select **Manual trigger a flow**
  - Click **Create**

---

## 3. Configure the Excel Step

- Click **+ New Step**
- Search for `"Excel"`
- Select **List rows present in a table [Excel Online (Business)]**
- Set parameters:
  - **Location:** OneDrive for Business
  - **Document Library:** OneDrive
  - **File:** Select your Excel file
  - **Table:** Choose your table (e.g., Table1)

---

## 4. Process Each Row with a Loop

- Click **+ New Step**
- Search for `"Apply to each"`, and select it
- In the "Select an output..." box, choose **Body/Value** from previous step

---

## 5. Inside the Loop: Send Email & Add Delay

- **Add "Send an Email (V2)" Action:**
  - In **To**, select `Email Id` from Excel
  - In **Subject**, select `Subject` or write your own
  - In **Body**, compose your message using Excel fields

- **Add "Delay" Action (inside same loop):**
  - Count: Enter number (e.g. 1, 2, 3)
  - Unit: Set to "Minute"

---

## 6. Save and Test the Flow

- Click **Save**
- Choose **Test > Manually**
- **Run the Flow.**
  - The first email sends immediately
  - Flow waits for your chosen delay, then sends the next email
  - Repeats for all rows

---

### 📊 Example Flow Structure

```text
Manually trigger a flow
↓
List rows present in a table
↓
Apply to each (body/value)
↓
Send an email (V2)
↓
Delay (X minutes)
```

---

## 🎯 Use Cases

- Sending personalized job applications
- Newsletter distribution with rate limiting
- Follow-up emails with timing control
- Event invitations with scheduled delivery

---

## 📝 Tips

- Test with a small dataset first
- Adjust delay based on email provider limits
- Use dynamic content from Excel for personalization
- Monitor flow runs in Power Automate dashboard

---

## 🤝 Contributing

Feel free to fork this repository and submit pull requests with improvements or additional examples!

---

## 📄 License

This project is open source and available under the MIT License.

---

**Created by Nikhil Airsang** | [GitHub](https://github.com/Nikhil414) | [LinkedIn](https://www.linkedin.com/in/nikhilairsang)

*⭐ If you find this helpful, please star the repository!*
