#include <xlnt/xlnt.hpp>
#include <iostream>
#include <vector>
using namespace std;

struct Trade {
    string type;
    double outcome;
};

static bool loadWorkbook(xlnt::workbook& workbook, const string& filename) {
    try {
        workbook.load(filename);
        return true;
    }
    catch (const std::exception& e) {
        std::cerr << "Failed to load MT5 backtest file!\n" << e.what() << "\n";
        return false;
    }
}

int main() {
    xlnt::workbook wb;
    string fileName = "C:\\Users\\jedi\\Desktop\\BacktestReport\\ReportTester-153306864.xlsx";

    if (!loadWorkbook(wb, fileName)) {
        return 1;
    }

    // get active sheet
    auto sheet = wb.active_sheet();

    double initialBalance = 0.0;
    vector<Trade> trades;

    // loop through rows and find the deals section
    bool dealsSectionFound = false;
    bool columnNamesRowFound = false;

    // iterator for rows
    auto rowIter = sheet.rows().begin();

    for (; rowIter != sheet.rows().end(); ++rowIter) {
        auto& row = *rowIter;
        string cellValue = row[0].to_string(); // Column 0 (Header)
        if (cellValue == "Deals") {
            dealsSectionFound = true;
            continue;
        }

        // if the deals section is found then process the rows
        if (dealsSectionFound) {
            // check if the next row is the column names row
            if (!columnNamesRowFound) {
                columnNamesRowFound = true;
                continue;
            }

            // extract the initial balance from the next row
            if (initialBalance == 0.0) {
                initialBalance = row[11].value<double>(); // Column 11 (Balance)
                continue; 
            }

            // process the trade data rows in pairs (in and out)
            if (row[4].to_string() == "in") { // Column 4 (Direction in or out)
                // get "in" row details
                string type = row[3].to_string(); // Column 3 Type (buy or sell)

                // move to next row (out)
                if (++rowIter != sheet.rows().end()) {
                    auto& nextRow = *rowIter;
                    double outcome = nextRow[10].value<double>(); // Column 10 (Profit)

                    trades.push_back({ type, outcome });
                }
            }
        }
    }

    cout << "Initial Balance: " << initialBalance << "\n";
    cout << "Extracted Trades:\n";
    int tradeSequnce = 1;
    for (const auto& trade : trades) {
        cout << tradeSequnce << ": Type: " << trade.type << ", Outcome: " << trade.outcome << "\n";
        tradeSequnce++;
    }

    return 0;
}
