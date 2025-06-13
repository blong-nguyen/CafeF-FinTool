# CafeF FinTool Excel Add-in Readme

A Custom Excel Add-in for Retrieving Historical Stock Price Data from CafeF

*Last Updated: June 13, 2025*

## Project Overview

CafeF FinTool is a custom Excel Add-in designed to automate the retrieval of historical stock price data from the CafeF website, a prominent financial data source in Vietnam. The add-in addresses a significant limitation in CafeF's data export functionality, which only downloads the first page of historical stock price data (typically 20 entries), even when the complete dataset spans multiple pages. This inefficiency often required extensive manual intervention to collect comprehensive data for financial analysis.

By leveraging JavaScript, HTML, CSS, and the Office.js API, CafeF FinTool streamlines the data collection process, enabling users to retrieve complete historical datasets with minimal effort. The add-in navigates through all available pages of stock price data on the CafeF website and consolidates them into a single, downloadable dataset, significantly enhancing productivity for financial analysts and investors.

## Installation

Since CafeF FinTool is a personal project shared with friends, it is not distributed through Microsoft's AppSource. Instead, installation is performed manually using the command prompt (Windows) or terminal (macOS). Follow the steps below to set up and run the add-in:

1. **Download and Install Node.js**
   - Visit [https://nodejs.org/en/download](https://nodejs.org/en/download) to download and install Node.js.
   - Follow the installation instructions for your operating system (Windows or macOS).
   - Verify installation by opening a command prompt (Windows) or terminal (macOS) and typing `node -v`. If installed correctly, you should see the Node.js version (e.g., `v18.x.x`).

2. **Download and Extract the Add-in**
   - Download the CafeF FinTool repository from GitHub: [https://github.com/blong-nguyen/CafeF-FinTool](https://github.com/blong-nguyen/CafeF-FinTool).
   - Extract the downloaded ZIP file to a folder of your choice (e.g., `C:\Users\YourName\CafeF-FinTool` on Windows or `/Users/YourName/CafeF-FinTool` on macOS).

3. **Navigate to the Project Directory**
   - Open a command prompt (Windows) or terminal (macOS).
   - Backtrack to the main disk by typing:
     - On Windows: `cd /`
     - On macOS: `cd ~`
   - Copy the full path of the extracted folder (e.g., right-click the folder in File Explorer on Windows or Finder on macOS, and select "Copy as path" or similar).
   - Use the `cd` command with the copied path enclosed in quotes. For example:
     - On Windows: `cd "C:\Users\YourName\CafeF-FinTool"`
     - On macOS: `cd "/Users/YourName/CafeF-FinTool"`
   - Verify you are in the correct directory by typing `dir` (Windows) or `ls` (macOS). You should see files like `package.json` and folders like `src`.
   - *Note*: The command has worked correctly if you see the project files listed.

4. **Install Required Packages**
   - In the command prompt (Windows) or terminal (macOS), type:
     ```bash
     npm install
     ```
   - This command installs all necessary dependencies listed in `package.json`.
   - The command is complete when you see a `node_modules` folder in the project directory and no error messages in the command prompt or terminal.

5. **Start the Add-in**
   - In the command prompt (Windows) or terminal (macOS), type:
     ```bash
     npm start
     ```
   - This command launches the add-in, making it available in Excel.
   - Open Excel, and you should see the CafeF FinTool add-in in the ribbon or add-ins menu.