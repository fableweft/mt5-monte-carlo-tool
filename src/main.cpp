#include <xlnt/xlnt.hpp>
#include <iostream>
#include <vector>
#include <random>
#include <algorithm>
#include <cmath>
#include <numeric>
#include <iomanip>

using namespace std;
struct Trade
{
    string type;
    double outcome;
};

struct SimulationMetrics
{
    double finalBalance;
    double maxDrawdown;
    double maxDrawdownPercent;
    double profitFactor;
    int totalTrades;
    double winRate;
    double sharpeRatio;
    double maxConsecutiveLosses;
    double averageWin;
    double averageLoss;
    double riskRewardRatio;
};

static bool loadWorkbook(xlnt::workbook &workbook, const string &filename)
{
    try
    {
        workbook.load(filename);
        return true;
    }
    catch (const std::exception &e)
    {
        std::cerr << "Failed to load file: " << filename << "\n"
                  << e.what() << "\n";
        return false;
    }
}

vector<SimulationMetrics> runMonteCarloSimulations(const vector<Trade> &trades, double initialBalance, int numSimulations, string &filename)
{
    if (trades.empty())
    {
        cerr << "Error: No trades provided for simulation.\n";
        return {};
    }
    if (initialBalance <= 0)
    {
        cerr << "Error: Initial balance must be positive.\n";
        return {};
    }

    vector<SimulationMetrics> simulationResults;
    random_device random;
    mt19937 generator(random());

    // create a normal distribution based on the historical trades
    double meanOutcome = accumulate(trades.begin(), trades.end(), 0.0, [](double sum, const Trade &trade)
                                    { return sum + trade.outcome; }) /
                         trades.size();
    double stdDev = 0;

    for (const auto &trade : trades)
    {
        stdDev += pow(trade.outcome - meanOutcome, 2);
    }

    stdDev = sqrt(stdDev / (trades.size() - 1));
    normal_distribution<double> distribution(meanOutcome, stdDev);

    for (int i = 0; i < numSimulations; ++i)
    {
        SimulationMetrics metrics{};
        double balance = initialBalance;
        double peakBalance = initialBalance;
        double currentDrawdown = 0;
        double maxDrawdown = 0;

        vector<double> tradeOutcomes;
        int consecutiveLosses = 0;
        int maxConsecutiveLosses = 0;
        double grossProfit = 0;
        double grossLoss = 0;

        // simulate trades
        for (size_t j = 0; j < trades.size(); ++j)
        {
            double outcome = distribution(generator);
            tradeOutcomes.push_back(outcome);

            balance += outcome;

            if (balance > peakBalance)
            {
                peakBalance = balance;
                currentDrawdown = 0;
            }
            else
            {
                currentDrawdown = peakBalance - balance;
                maxDrawdown = max(maxDrawdown, currentDrawdown);
            }

            // track the consecutive losses
            if (outcome < 0)
            {
                consecutiveLosses++;
                grossLoss += abs(outcome);
            }
            else
            {
                maxConsecutiveLosses = max(maxConsecutiveLosses, consecutiveLosses);
                consecutiveLosses = 0; // reset consecutive losses
                grossProfit += outcome;
            }
        }

        // calc metrics
        metrics.finalBalance = balance;
        metrics.maxDrawdown = maxDrawdown;
        metrics.maxDrawdownPercent = (peakBalance != 0) ? (maxDrawdown / peakBalance) * 100 : 0;
        metrics.profitFactor = (grossLoss != 0) ? grossProfit / grossLoss : 0;
        metrics.totalTrades = trades.size();

        // verify total trades
        if (metrics.totalTrades != trades.size())
        {
            cerr << "Error: Mismatch in total trades for simulation " << i + 1 << ".\n";
        }

        
        int winningTrades = count_if(tradeOutcomes.begin(), tradeOutcomes.end(), [](double outcome)
                                     { return outcome > 0; });
        metrics.winRate = (!tradeOutcomes.empty()) ? (static_cast<double>(winningTrades) / tradeOutcomes.size()) * 100 : 0;

        
        double returns_mean = accumulate(tradeOutcomes.begin(), tradeOutcomes.end(), 0.0) / tradeOutcomes.size();
        double returns_stddev = 0;
        for (double outcome : tradeOutcomes)
        {
            returns_stddev += pow(outcome - returns_mean, 2);
        }
        returns_stddev = sqrt(returns_stddev / (tradeOutcomes.size() - 1));
        metrics.sharpeRatio = (returns_stddev != 0) ? (returns_mean / returns_stddev) * sqrt(252) : 0; // annualized
        metrics.maxConsecutiveLosses = maxConsecutiveLosses;

        
        vector<double> wins, losses;
        for (double outcome : tradeOutcomes)
        {
            if (outcome > 0)
                wins.push_back(outcome);
            if (outcome < 0)
                losses.push_back(abs(outcome));
        }

        metrics.averageWin = !wins.empty() ? accumulate(wins.begin(), wins.end(), 0.0) / wins.size() : 0;
        metrics.averageLoss = !losses.empty() ? accumulate(losses.begin(), losses.end(), 0.0) / losses.size() : 0;
        metrics.riskRewardRatio = (metrics.averageLoss != 0) ? metrics.averageWin / metrics.averageLoss : 0;

        simulationResults.push_back(metrics);
    }

    
    cout << "Total trades parsed from " << filename << ": " << trades.size() << "\n";
    cout << "Total trades per simulation: " << simulationResults[0].totalTrades << "\n";

    return simulationResults;
}

int main()
{
    xlnt::workbook wb;
    string fileName = "C:\\Users\\jedi\\Desktop\\BacktestReport\\ReportTester-153306864.xlsx";

    if (!loadWorkbook(wb, fileName))
    {
        return 1;
    }

    // get active sheet
    auto sheet = wb.active_sheet();

    double initialBalance = 0.0;
    vector<Trade> trades;

    // loop through rows and find the deals section
    bool dealsSectionFound = false;
    bool columnNamesRowFound = false;


    auto rows = sheet.rows(); 
    for (auto rowIter = rows.begin(); rowIter != rows.end(); ++rowIter)
    {
        auto row = *rowIter;
        string cellValue = row[0].to_string(); 
        if (cellValue == "Deals")
        {
            dealsSectionFound = true;
            continue;
        }

       if (dealsSectionFound)
        {
            if (!columnNamesRowFound)
            {
                columnNamesRowFound = true;
                continue;
            }

            if (initialBalance == 0.0)
            {
                initialBalance = row[11].value<double>(); 
                continue;
            }

            // process the trade data rows in pairs (in and out)
            if (row[4].to_string() == "in")
            { 
                string type = row[3].to_string(); 

                if (++rowIter != rows.end())
                {
                    
                    auto nextRow = *rowIter;
                    double outcome = nextRow[10].value<double>(); 

                    trades.push_back({type, outcome});
                }
            }
        }
    }
    cout << "Initial Balance: " << initialBalance << "\n";
    cout << "Extracted Trades:\n";
    int tradeSequnce = 1;

    for (const auto &trade : trades)
    {
        cout << tradeSequnce << ": Type: " << trade.type << ", Outcome: " << trade.outcome << "\n";
        tradeSequnce++;
    }

    cout << "\n";

    // run monte carlo sims

    int numSimulations = 1000;
    vector<SimulationMetrics> results = runMonteCarloSimulations(trades, initialBalance, numSimulations, fileName);

    if (results.empty())
    {
        cerr << "Error: No results from simulations.\n";
        return 1;
    }

    cout << "\nMonte Carlo Simulation Results:\n";
    cout << "--------------------------------\n";

    // calculate percentiles
    vector<double> finalBalances;
    vector<double> maxDrawdowns;
    for (const auto &result : results)
    {
        finalBalances.push_back(result.finalBalance);
        maxDrawdowns.push_back(result.maxDrawdownPercent);
    }

    sort(finalBalances.begin(), finalBalances.end());
    sort(maxDrawdowns.begin(), maxDrawdowns.end());

    size_t index5th = size_t(numSimulations * 0.05);
    size_t index50th = size_t(numSimulations * 0.5);
    size_t index95th = size_t(numSimulations * 0.95);

    if (index5th >= finalBalances.size() || index50th >= finalBalances.size() || index95th >= finalBalances.size())
    {
        cerr << "Error: Not enough simulations to calculate percentiles.\n";
        return 1;
    }

    cout << fixed << setprecision(2);
    cout << "5th Percentile Balance: $" << finalBalances[index5th] << "\n";
    cout << "50th Percentile Balance: $" << finalBalances[index50th] << "\n";
    cout << "95th Percentile Balance: $" << finalBalances[index95th] << "\n\n";

    cout << "Maximum Drawdown (95th percentile): " << maxDrawdowns[index95th] << "%\n";

    if (numSimulations == 0)
    {
        cerr << "Error: Number of simulations must be greater than zero.\n";
        return 1;
    }

    SimulationMetrics avgMetrics = {
        0.0, // finalBalance
        0.0, // maxDrawdown
        0.0, // maxDrawdownPercent
        0.0, // profitFactor
        0,   // totalTrades
        0.0, // winRate
        0.0, // sharpeRatio
        0.0, // maxConsecutiveLosses
        0.0, // averageWin
        0.0, // averageLoss
        0.0  // riskRewardRatio
    };

    for (const auto &result : results)
    {
        avgMetrics.winRate += result.winRate;
        avgMetrics.profitFactor += result.profitFactor;
        avgMetrics.sharpeRatio += result.sharpeRatio;
        avgMetrics.maxConsecutiveLosses += result.maxConsecutiveLosses;
        avgMetrics.riskRewardRatio += result.riskRewardRatio;
    }

    cout << "\nAverage Metrics Across All Simulations:\n";
    cout << "Win Rate: " << avgMetrics.winRate / numSimulations << "%\n";
    cout << "Profit Factor: " << avgMetrics.profitFactor / numSimulations << "\n";
    cout << "Sharpe Ratio: " << avgMetrics.sharpeRatio / numSimulations << "\n";
    cout << "Average Max Consecutive Losses: " << avgMetrics.maxConsecutiveLosses / numSimulations << "\n";
    cout << "Risk/Reward Ratio: " << avgMetrics.riskRewardRatio / numSimulations << "\n";

    return 0;
}
