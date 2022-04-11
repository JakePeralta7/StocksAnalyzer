# Imports
import yfinance as yf
from datetime import datetime
from os import path
from docx import Document


class Stock:
    def __init__(self, symbol):
        self.data = yf.Ticker(symbol)
    
    def extract_info(self):
        """
        Extracts intresting info about the stock
        """

        self.name = self.extract_value("longName")
        self.symbol = self.extract_value("symbol")
        self.business_summary = self.extract_value("longBusinessSummary")
        self.industry = self.extract_value("industry")
        self.sector = self.extract_value("sector")
        self.country = self.extract_value("country")
        self.current_price = self.extract_value("currentPrice")
        self.market_cap = self.extract_value("marketCap")
        self.volume = self.extract_value("volume")
        self.average_volume = self.extract_value("averageVolume")
        self.trailing_pe = self.extract_value("trailingPE")
        self.forward_pe = self.extract_value("forwardPE")
        self.beta = self.extract_value("beta")
        self.dividend_yield = self.extract_value("yield")
        self.fifty_two_week_high = self.extract_value("fiftyTwoWeekHigh")
        self.fifty_two_week_low = self.extract_value("fiftyTwoWeekLow")
        self.two_hundred_day_average = self.extract_value("twoHundredDayAverage")
        self.fifty_two_week_change = self.extract_value("52WeekChange")
        self.s_and_p_fifty_two_week_change = self.extract_value("SandP52WeekChange")
    
    def extract_value(self, key):
        try:
            return self.data.info[key]
        except KeyError:
            return None
    
    def analyze_earnings(self):

        # Extracts stock's last four years earnings
        self.earnings = self.data.get_earnings().to_dict()["Earnings"]
        current_earnings = list(self.earnings.values())[-1]
        for year, earnings in list(self.earnings.items())[:-1]:
            if (earnings > 0):
                growth = (current_earnings - earnings) / earnings
            else:
                growth = (current_earnings - earnings) / earnings * -1
            print(f"Earning's growth since {year}: {growth:%}")
    
    def write_report(self):
        
        # Example for report name: aapl(11-04-22).txt
        self.report_path = f"{self.symbol}({datetime.now().strftime('%d-%m-%y')}).txt"
        
        # Opens the file for editing
        self.report = open(self.report_path, "w")


def main():
    #stock_symbol = input("Enter stock symbol: ")
    stock_symbol = "kbwy"
    my_stock = Stock(stock_symbol)
    my_stock.extract_info()
    print(f"{my_stock.dividend_yield:.2%}")


if __name__ == "__main__":
    main()
