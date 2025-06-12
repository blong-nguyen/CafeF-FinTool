import axios from 'axios';

// Fetch the raw HTML from the website
export async function fetchStockPrice(ticker, startDate, endDate) {
    const url = `https://cafef.vn/du-lieu/Ajax/PageNew/DataHistory/PriceHistory.ashx`;

    const startDateAsDate = new Date(startDate);
    const endDateAsDate = new Date(endDate);

    let numQueries = Math.floor((endDateAsDate - startDateAsDate) / (1000 * 60 * 60 * 24)) + 1;
    console.log(numQueries);
    const params = {
        Symbol: ticker.toUpperCase(),
        StartDate: startDate,
        EndDate: endDate,
        PageSize: numQueries
    }
    
    try {
        const response = await axios.get(url, {
            params,
            headers: {
                'Content-Type': 'text/plain'
            }
        });
        return response.data;
    } catch (error) {
        throw new Error(`Failed to fetch page: ${error.message}`);
    }
}

export function storePriceData(jsonData) {
    const dataArray = jsonData.Data.Data;
    if (!dataArray || dataArray.length === 0) {
        throw new Error("No data available!");
    }

    const headers = Object.keys(dataArray[0]);
    const rows = dataArray.map(row => Object.values(row))

    return [headers, ...rows]
}

