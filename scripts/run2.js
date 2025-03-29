// Step 6: Add reconciliation summary (mirroring the right side of "Final" sheet)
    // Calculate card brand totals
    
    function calculateCardTotals(data, cardBrandColIndex, amountColIndex) {
        const totals = {};
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row.length > Math.max(cardBrandColIndex, amountColIndex) && row[cardBrandColIndex] && row[cardBrandColIndex] !== null) {
                const cardBrand = row[cardBrandColIndex].toString().toLowerCase();
                const amount = parseFloat(row[amountColIndex]) || 0;
                totals[cardBrand] = (totals[cardBrand] || 0) + amount;
            }
        }
        return totals;
    }

    const paymentsHubTotals = calculateCardTotals(paymentsHubWithCount, cardBrandColIndex, krColIndex);
    // Use salesCardBrandColIndex and salesAmountColIndex here
    const salesTotals = calculateCardTotals(processedSalesData, salesCardBrandColIndex, salesAmountColIndex);

    // FIXED: Updated header and calculated totals
    const summaryData = [
        ["Hub Report", "Total", "", "Sales Report", "Total", "", "Difference"],
        ["Visa", paymentsHubTotals['visa'] || 0, "", "Visa", salesTotals['visa'] || 0, "", (paymentsHubTotals['visa'] || 0) - (salesTotals['visa'] || 0)],
        ["Mastercard", paymentsHubTotals['mastercard'] || 0, "", "Mastercard", salesTotals['mastercard'] || 0, "", (paymentsHubTotals['mastercard'] || 0) - (salesTotals['mastercard'] || 0)],
        ["American Express", paymentsHubTotals['american express'] || 0, "", "American Express", salesTotals['american express'] || 0, "", (paymentsHubTotals['american express'] || 0) - (salesTotals['american express'] || 0)],
        ["Discover", paymentsHubTotals['discover'] || 0, "", "Discover", salesTotals['discover'] || 0, "", (paymentsHubTotals['discover'] || 0) - (salesTotals['discover'] || 0)]
    ];