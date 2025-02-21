# TestLodge Connect

TestLodge Connect is a Google Apps Script project that integrates Google Sheets with the TestLodge API. It allows users to get data about test cases, and test runs.

## Features
- 📊 Fetch TestLodge test cases and results into a Google Spreadsheet.
- 📂 Organize test case data into structured sheets for better tracking and metrics.

## Installation
1. Open your Google Spreadsheet.
2. Navigate to **Extensions > Apps Script**.
3. Copy and paste the contents of this repository into the Apps Script editor.
4. Save and deploy the script.

## Configuration
1. 🔑 Set up your TestLodge API credentials in the script.
2. 🌍 Define your TestLodge project ID and test suite details.
3. 📑 Ensure your Google Spreadsheet has the necessary structure to receive and update TestLodge data.

## Main Functions
### `manageTestLodgeData()`
- 🚀 The core function that interacts with the TestLodge API.
- 📥 Fetches, processes, and updates data in Google Sheets.

### `write*` Functions
- ✍️ Any function starting with `write` is responsible for writing data to the spreadsheet.
- 📌 These functions handle different aspects of test case management and distribute them across designated sheets.

## Usage
1. ▶️ Run `manageTestLodgeData()` to fetch and update TestLodge test case information.
2. 📝 Use `write*` functions to modify or add test execution results in the spreadsheet.
3. ⏰ Automate script execution using Google Apps Script triggers for periodic updates.

## Contributing
1. 🍴 Fork this repository.
2. 🌱 Create a new branch (`feature-branch-name`).
3. 💾 Commit your changes.
4. 📤 Push the branch and create a pull request.

## License
📝 This project is licensed under the MIT License.

## Contact
📬 For any issues or feature requests, open an issue on GitHub or reach out to the repository owner.

