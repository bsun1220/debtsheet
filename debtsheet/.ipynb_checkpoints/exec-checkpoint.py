from debtsheet import *
import sys
import warnings
warnings.filterwarnings('ignore')

if __name__ == "__main__":
    if len(sys.argv) <= 1:
        print("Error: add company ticker like 'exec.py TICKER1 TICKER2'")
    else:
        for i in sys.argv[1:]:
            try:
                DebtSheet(i)
                print("Completed document for " + i)
            except:
                print("Could not find " + i + " ticker")