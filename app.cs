#:package MiniExcel@1.31.3

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using MiniExcelLibs;
using MiniExcelLibs.Csv;

var storesFile = "PipStores.xlsx";
var movementsFile = "PipMovimientosGiftcards.xlsx";
var outputFile = "PipMovimientosWithStores.xlsx";

// 1. Log the first 5 rows
// We use useHeaderRow: true so columns are identified by name (Id, StoreName) instead of A, B
PrintRows(storesFile);
PrintRows(movementsFile);

// 2. Generate the new file
Console.WriteLine($"--- Generating {outputFile} ---");

Console.WriteLine("Loading stores into memory...");
// Fix: Pass 'useHeaderRow: true' to ensure we can access row.Id and row.StoreName
var storesMap = MiniExcel.Query(storesFile, useHeaderRow: true)
    .Where(row => row.Id != null)
    .ToDictionary(row => (string)row.Id, row => (string)row.StoreName);

Console.WriteLine($"Loaded {storesMap.Count} stores.");

Console.WriteLine("Processing movements...");
var movements = MiniExcel.Query(movementsFile, useHeaderRow: true);

var result = movements.Select(m => new
{
    m.movementId,
    m.amount,
    m.newBalance,
    m.originType,
    m.originId,
    m.timestamp,
    m.owner,
    // Join logic: if Store Purchase, look up the name using the ID
    OriginName = (m.originType == "STORE_PURCHASE" && m.originId != null && storesMap.ContainsKey((string)m.originId))
                 ? storesMap[(string)m.originId]
                 : null
});

Console.WriteLine($"Saving to {outputFile}...");
MiniExcel.SaveAs(outputFile, result);
Console.WriteLine("Process completed successfully!");


// --- Helper Function ---

void PrintRows(string filePath)
{
    Console.WriteLine($"--- Reading File: {filePath} ---");
    if (!File.Exists(filePath))
    {
        Console.WriteLine($"[Error] File not found: {filePath}\n");
        return;
    }

    try
    {
        // Fix: useHeaderRow: true allows us to see the actual column names in the log
        var rows = MiniExcel.Query(filePath, useHeaderRow: true).Take(5);

        foreach (var row in rows)
        {
            var data = (IDictionary<string, object>)row;
            var rowLog = string.Join(" | ", data.Select(x => $"{x.Key}: {x.Value}"));
            Console.WriteLine(rowLog);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[Error] Failed to read file: {ex.Message}");
    }
    Console.WriteLine();
}
