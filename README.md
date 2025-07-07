# Google Apps Script: File Linker and Booking System
This Google Apps Script project provides functionalities to automate the linking of PDF/Excel files from a Google Drive folder to a Google Sheet and to manage a booking system directly within a Google Sheet.

## Table of Contents
- Features

- How It Works

- Setup and Installation

- Usage

- Customization

- Important Notes

- Contributing

- License

## Features
- **Automated PDF/File Linking:** Scans a specified Google Drive folder for files, extracts a numerical identifier from their names, and automatically updates corresponding rows in a Google Sheet with the file name, URL, and a "Attached" status.

- **Booking System:** Manages a booking process for "numbers" or "items" within a Google Sheet, marking them as "In Use" and assigning them to a PIC (Person In Charge).

- **Automatic File Generation (for "Other" bookings):** For certain booking types, it can copy a template Excel/Sheet file from Drive, rename it, update a specific cell with the booked number, and link it back to the main Google Sheet.

- **User Interface (UI) for Booking:** Provides a simple modal dialog within Google Sheets to facilitate the booking process.

- **Concurrency Control:** Uses LockService to prevent multiple users from simultaneously booking, ensuring data integrity.

- **Last Update Timestamp:** Automatically records the last time the PDF linking process was run.

## How It Works
This script consists of several functions designed to interact with Google Sheets and Google Drive.

`updateLinkPDF()`
This function iterates through files in a specified Google Drive folder. It attempts to extract a single number from the beginning of each file name (e.g., "1234 Document.pdf"). If a number is found, it looks for that number in column B of your Google Sheet. If a match is found and the status in column I is not "✅ Sudah Kembali", it updates columns E (File Name), H (Link), and I (Status) for that row.

`lastUpdate()`
A simple helper function that updates a specific cell (H2) in the sheet named '2025' with the current date and time, indicating when the `updateLinkPDF` function was last executed.

`showBookingForm()`
This function creates and displays a modal dialog in Google Sheets. It loads content from an HTML file named `BookingUI.html` (which you'll need to create) to provide a user interface for the booking process.

`processBooking(jenis, jumlah, pic)`
This is the main entry point for handling booking requests, typically called from the BookingUI.html form.
- It uses LockService to ensure only one booking process runs at a time, preventing conflicts.
- Based on the jenis (type) of booking, it calls either bookNumbersWithRange or bookOthersWithRange.
- It displays a success or error message to the user.

`bookNumbersWithRange(quantity, jenis, pic)`
This function is used for "Automatic" type bookings. It scans the sheet named '2025' for rows with a "READY" status in column I. It then marks the specified `quantity` of consecutive "READY" numbers as "🔒 Dipakai" (In Use), records the `pic`, `jenis`, and `Tanggal Booking` (Booking Date). It returns the range of booked numbers (e.g., "0001 - 0005").

`bookOthersWithRange(jumlah, jenis, pic)`
This function is used for "Other" type bookings.
- It finds jumlah (quantity) of "READY" numbers in the sheet.
- It then copies a specified template file (YOUR_FOLDER_ID) to a target folder (File Excel Generate RI).
- The copied file is renamed to include the booked number (e.g., "Form serah Terima 0001").
- A specific cell ("C4") within the copied spreadsheet is updated with the booked number.
- Finally, it updates the main sheet with the pic, file name, date, jenis, link to the new file, and sets the status to "🔒 Dipakai".
- It returns the range of booked numbers.

## Setup and Installation
To use this script, you'll need a Google Account and access to Google Sheets and Google Drive.
1. **Create a New Google Sheet:** Go to Google Sheets and create a new spreadsheet.
2. **Open Apps Script:** In your new Google Sheet, go to Extensions > Apps Script. This will open the Google Apps Script editor.
3. **Copy and Paste Code:** Delete any existing code in Code.gs and paste the entire code provided into the editor.
4. **Create BookingUI.html:**
- In the Apps Script editor, go to File > New > HTML file.
- Name the file BookingUI.
- You will need to create the HTML content for this file. It should contain a form with inputs for jenis (type, e.g., "Automatic", "Other"), jumlah (quantity), and pic (person in charge), and a button to submit the data. The form should call the processBooking function using google.script.run.

Example `BookingUI.html` structure (you'll need to fill in the actual form elements):
```
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    /* Add your CSS styling here */
    body { font-family: 'Inter', sans-serif; margin: 20px; }
    label { display: block; margin-bottom: 5px; font-weight: bold; }
    input[type="text"], input[type="number"], select {
      width: calc(100% - 22px);
      padding: 10px;
      margin-bottom: 15px;
      border: 1px solid #ccc;
      border-radius: 8px;
      box-sizing: border-box;
    }
    button {
      background-color: #4CAF50;
      color: white;
      padding: 12px 20px;
      border: none;
      border-radius: 8px;
      cursor: pointer;
      font-size: 16px;
      width: 100%;
      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
      transition: background-color 0.3s ease;
    }
    button:hover {
      background-color: #45a049;
    }
    .message-box {
      margin-top: 15px;
      padding: 10px;
      border-radius: 8px;
      background-color: #f0f0f0;
      border: 1px solid #ddd;
      display: none; /* Hidden by default */
    }
    .message-box.show {
      display: block;
    }
  </style>
</head>
<body>
  <h2>Form Pengambilan Nomor</h2>
  <form>
    <label for="jenis">Jenis Booking:</label>
    <select id="jenis" name="jenis">
      <option value="Automatic">Automatic</option>
      <option value="Other">Other</option>
    </select>

    <label for="jumlah">Jumlah:</label>
    <input type="number" id="jumlah" name="jumlah" min="1" value="1">

    <label for="pic">PIC:</label>
    <input type="text" id="pic" name="pic">

    <button type="button" onclick="handleBooking()">Book Now</button>
  </form>

  <div id="messageBox" class="message-box"></div>

  <script>
    function handleBooking() {
      const jenis = document.getElementById('jenis').value;
      const jumlah = document.getElementById('jumlah').value;
      const pic = document.getElementById('pic').value;
      const messageBox = document.getElementById('messageBox');

      messageBox.textContent = 'Processing...';
      messageBox.classList.add('show');

      google.script.run
        .withSuccessHandler(function(message) {
          messageBox.textContent = 'Booking berhasil:\n' + message;
          messageBox.style.backgroundColor = '#d4edda';
          messageBox.style.borderColor = '#28a745';
          messageBox.style.color = '#155724';
        })
        .withFailureHandler(function(error) {
          messageBox.textContent = 'Error: ' + error.message;
`          messageBox.style.backgroundColor = '#f8d7da';
          messageBox.style.borderColor = '#dc3545';
          messageBox.style.color = '#721c24';
        })
        .processBooking(jenis, parseInt(jumlah), pic);
    }
  </script>
</body>
</html>
```
5. **Update Placeholders:** In the Code.gs file, you MUST replace the following placeholders:

- `"Your_Folder_Name"`: Replace with the exact name of your Google Drive folder containing the files you want to link (in `updateLinkPDF`).

- `"Your_Sheet_Name"`: Replace with the exact name of the sheet where your data resides (in `updateLinkPDF`).

- `"YOUR_FOLDER_ID"`: Replace with the actual ID of your template file in Google Drive (in `bookOthersWithRange`). You can find this in the URL of the file: `https://docs.google.com/spreadsheets/d/YOUR_FOLDER_ID/edit...`

6. **Set up `onOpen()` function (Recommended):** To easily access the booking form from your Google Sheet, add the following function to your Code.gs file. This will create a custom menu in your spreadsheet.
```
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
      .addItem('Show Booking Form', 'showBookingForm')
      .addItem('Update PDF Links', 'updateLinkPDF')
      .addToUi();
}
```
After adding `onOpen()`, save the script and refresh your Google Sheet. A new "Custom Tools" menu will appear.

7. **Grant Permissions:** The first time you run any function that interacts with Google Drive or modifies the spreadsheet, Google will ask you to authorize the script. Follow the prompts to grant the necessary permissions.

## Usage
Updating PDF Links
- Manually: Go to Extensions > Apps Script in your Google Sheet, select the updateLinkPDF function from the dropdown, and click the "Run" button (play icon).

- Via Custom Menu: If you implemented the onOpen() function, simply go to Custom Tools > Update PDF Links in your Google Sheet.

- Via Time-driven Trigger (Optional): You can set up a time-driven trigger to run updateLinkPDF automatically at a set interval (e.g., hourly, daily).
  - In the Apps Script editor, click the Triggers icon (looks like a clock).
  - Click Add Trigger.
  - Choose updateLinkPDF for the function to run.
  - Select Time-driven as the event source.
  - Choose your desired time interval.

Booking Numbers
- Via Custom Menu: Go to Custom Tools > Show Booking Form in your Google Sheet. This will open the modal dialog.

- Fill in the Jenis Booking, Jumlah, and PIC fields, then click "Book Now".

## Customization
- Sheet Names: Adjust sheet names ("2025", "Your_Sheet_Name") in the script to match your actual sheet names.
- Folder Names/IDs: Update Your_Folder_Name and YOUR_FOLDER_ID to point to your specific Google Drive resources.
- Column Indices: The script uses fixed column indices (e.g., data[i][1] for column B, data[i][8] for column I). If your sheet layout changes, you'll need to adjust these indices.
  - Column B (index 1): nomor (number)
  - Column E (index 4): fileName (file name)
  - Column H (index 7): link (file URL)
  - Column I (index 8): status (e.g., "READY", "✅ Sudah Kembali", "🔒 Dipakai", "📎 Terlampir")
  - Column C (index 2): pic
  - Column F (index 5): Tanggal Booking
  - Column G (index 6): jenis

- File Naming Convention: The updateLinkPDF function currently expects file names to start with a number (e.g., "1234 Document.pdf"). If your file naming convention is different, you'll need to modify the regular expressions within updateLinkPDF.

- BookingUI.html: Customize the BookingUI.html file to fit your desired form layout and styling.

## Important Notes
- Permissions: Ensure the script has the necessary permissions to access your Google Sheets and Google Drive.

- Folder/File IDs: Double-check that the folder names and file IDs you provide in the script are correct and accessible by the script.

- Error Handling: The script includes basic error handling and uses SpreadsheetApp.getUi().alert() for user notifications. In a production environment, you might want more robust logging or notification mechanisms.

- Concurrency: The LockService is crucial for multi-user environments to prevent race conditions during booking.

- Contributing
Feel free to fork this repository, make improvements, and submit pull requests.

## License
This project is licensed under the MIT License - see the LICENSE.md file for details (if you create one).
