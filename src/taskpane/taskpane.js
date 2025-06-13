/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { fetchStockPrice, storePriceData } from "../utils/utils.js";

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("fetch-data-button")
      .addEventListener("click", () => {
        fetchData();
      })
  }
});

async function fetchData() {
  console.log("===============start===================");
  const ticker = document.getElementById("ticker").value.trim();
  const startDate = document.getElementById("start-date").value;
  const endDate = document.getElementById("end-date").value;
  const status = document.getElementById("status");
  const button = document.getElementById("fetch-data-button");

  status.innerText = "";

  if (!ticker || !startDate || !endDate) {
    status.innerText = "Please fill in ticker and both dates.";
    return;
  } else if (new Date(startDate) > new Date(endDate)) {
    status.innerText = "Start Date must be before End Date!";
    return;
  }

  button.disabled = true;
  status.innerText = "Fetching Data...";

  try {
    const rawData = await fetchStockPrice(ticker, startDate, endDate);
    const storedData = storePriceData(rawData);

    insertDataToExcel(storedData, ticker, startDate, endDate, status);
  } catch (error) {
    status.innerText = `Error: ${error.message}`;
    console.error("Failed to Fetch Data.");
  } finally {
    status.innerText = "All Done!";
    button.disabled = false;
  }
}

async function insertDataToExcel (data, ticker, startDate, endDate, status) {
  Excel.run(async (context) => {
    try {
      const sheets = context.workbook.worksheets;
      let dataSheet = sheets.getItemOrNullObject(ticker);
      await context.sync();
      
      // Create new worksheets for results
      if (dataSheet.isNullObject) {
        dataSheet = sheets.add(ticker.toUpperCase());
      } else {
        dataSheet.delete();
        dataSheet = sheets.add(ticker);
      }      
      dataSheet.activate();
      
      data = convertDate(data);

      // Insert into Excel
      const range = dataSheet.getRange("A2").getResizedRange(data.length - 1, data[0].length - 1);
      dataSheet.getRange("A1").values = ticker.toUpperCase() + ": " + startDate + " - " + endDate;
      range.values = data;
      await context.sync();

      status.innerText = "All Done!";
      console.log("Data inserted into", ticker);
    } catch (error) {
      console.error("Excel insertion failed:", error);
      throw error;
    }
  })
}

function convertDate (data) {
  // Convert "Ngay" into Date objects
  if (!data.length || !data[0]?.length) {
    throw new Error("Data array is empty or invalid");
  }

  // Find the index of the "Ngay" column (date column)
  const headers = data[0];
  const ngayIndex = headers.indexOf("Ngay");
  if (ngayIndex === -1) {
    throw new Error("Date column 'Ngay' not found in data");
  }

  // Create a new 2D array with "Ngay" converted to Date objects
  const correctedData = data.map((row, index) => {
    if (index === 0) return row; // Keep the header row unchanged
    const dateStr = row[ngayIndex]; // Get the date string
    if (dateStr) {
      const [day, month, year] = dateStr.split('/');
      // Create a Date object with month-1
      const dateObj = new Date(year, month - 1, day);
      row[ngayIndex] = dateObj; // Replace with Date object
    }
    return row;
  });

  return correctedData;
}
