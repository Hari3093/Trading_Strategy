"""
Trading Journal - CLI Trade Logger
Add, view, edit, delete trades from terminal
Run: python scripts/log_trade.py
"""

import csv
import os
from datetime import datetime

DATA_FILE = "data/trades.csv"
FIELDS = ['Date','Time','Asset','Trade_Type','Entry_Price','Exit_Price',
          'Stop_Loss','Take_Profit','Position_Size','Risk_Pct','Profit_Loss','Notes','Strategy_Tag']

STRATEGIES = ['Breakout','Reversal','Trend Follow','Momentum','Support/Resistance','Scalping','Other']

def load_trades():
    if not os.path.exists(DATA_FILE):
        return []
    with open(DATA_FILE, newline='') as f:
        return list(csv.DictReader(f))

def save_trades(trades):
    with open(DATA_FILE, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(trades)

def get_input(prompt, default='', validator=None):
    while True:
        val = input(f"  {prompt}" + (f" [{default}]" if default else "") + ": ").strip()
        if not val and default:
            return default
        if validator:
            try:
                return validator(val)
            except:
                print(f"  ❌ Invalid input. Try again.")
        else:
            return val

def add_trade():
    print("\n➕ ADD NEW TRADE")
    print("─" * 40)
    trade = {}
    trade['Date']          = get_input("Date (YYYY-MM-DD)", datetime.today().strftime('%Y-%m-%d'))
    trade['Time']          = get_input("Time (HH:MM)", datetime.now().strftime('%H:%M'))
    trade['Asset']         = get_input("Asset (e.g. NIFTY, RELIANCE)").upper()
    trade['Trade_Type']    = get_input("Trade Type (Buy/Sell)").capitalize()
    trade['Entry_Price']   = get_input("Entry Price", validator=float)
    trade['Exit_Price']    = get_input("Exit Price", validator=float)
    trade['Stop_Loss']     = get_input("Stop Loss", validator=float)
    trade['Take_Profit']   = get_input("Take Profit", validator=float)
    trade['Position_Size'] = get_input("Position Size (qty)", validator=int)
    trade['Risk_Pct']      = get_input("Risk % (e.g. 1.0)", validator=float)

    # Auto-calculate P&L
    ep = float(trade['Entry_Price'])
    xp = float(trade['Exit_Price'])
    qty = int(trade['Position_Size'])
    pnl = (xp - ep) * qty if trade['Trade_Type'] == 'Buy' else (ep - xp) * qty
    print(f"  💰 Calculated P&L: ₹{pnl:,.2f}")
    trade['Profit_Loss'] = get_input("Confirm P&L or enter manually", default=str(round(pnl, 2)), validator=float)

    trade['Notes'] = get_input("Notes (optional)", default='')

    print("\n  Strategies:")
    for i, s in enumerate(STRATEGIES, 1):
        print(f"  {i}. {s}")
    strat_idx = get_input("Select strategy (1-7)", default='7', validator=int)
    trade['Strategy_Tag'] = STRATEGIES[min(int(strat_idx)-1, 6)]

    trades = load_trades()
    trades.append({k: str(v) for k, v in trade.items()})
    save_trades(trades)
    print(f"\n  ✅ Trade logged! Total trades: {len(trades)}")

def view_trades():
    trades = load_trades()
    if not trades:
        print("\n  📭 No trades yet.")
        return

    print(f"\n📋 TRADE LOG ({len(trades)} trades)")
    print("─" * 90)
    print(f"{'#':<4} {'Date':<12} {'Asset':<12} {'Type':<6} {'Entry':<10} {'Exit':<10} {'P&L':<12} {'Strategy'}")
    print("─" * 90)
    for i, t in enumerate(trades, 1):
        pnl = float(t['Profit_Loss'])
        sign = "+" if pnl > 0 else ""
        print(f"{i:<4} {t['Date']:<12} {t['Asset']:<12} {t['Trade_Type']:<6} "
              f"₹{float(t['Entry_Price']):<9,.0f} ₹{float(t['Exit_Price']):<9,.0f} "
              f"{sign}₹{pnl:<10,.0f} {t['Strategy_Tag']}")
    print("─" * 90)
    total = sum(float(t['Profit_Loss']) for t in trades)
    print(f"{'TOTAL P&L:':<60} {'+'if total>=0 else ''}₹{total:,.0f}")

def delete_trade():
    view_trades()
    trades = load_trades()
    if not trades:
        return
    idx = get_input(f"\nEnter trade # to delete (1-{len(trades)})", validator=int)
    idx = int(idx) - 1
    if 0 <= idx < len(trades):
        removed = trades.pop(idx)
        save_trades(trades)
        print(f"  ✅ Deleted trade: {removed['Date']} {removed['Asset']}")
    else:
        print("  ❌ Invalid trade number.")

def summary():
    trades = load_trades()
    if not trades:
        print("\n  📭 No trades yet.")
        return
    pnls = [float(t['Profit_Loss']) for t in trades]
    wins = sum(1 for p in pnls if p > 0)
    losses = sum(1 for p in pnls if p < 0)
    total = len(pnls)
    print(f"\n📊 QUICK SUMMARY")
    print("─" * 35)
    print(f"  Total Trades : {total}")
    print(f"  Wins         : {wins} ({wins/total*100:.1f}%)")
    print(f"  Losses       : {losses} ({losses/total*100:.1f}%)")
    print(f"  Total P&L    : ₹{sum(pnls):,.0f}")
    print(f"  Avg Trade    : ₹{sum(pnls)/total:,.0f}")
    print(f"  Best Trade   : ₹{max(pnls):,.0f}")
    print(f"  Worst Trade  : ₹{min(pnls):,.0f}")

def main():
    print("\n🗂️  TRADING JOURNAL CLI")
    print("=" * 35)
    while True:
        print("\n  1. ➕ Add Trade")
        print("  2. 📋 View All Trades")
        print("  3. 🗑️  Delete Trade")
        print("  4. 📊 Quick Summary")
        print("  5. 🚪 Exit")
        choice = input("\n  Select option (1-5): ").strip()
        if choice == '1':   add_trade()
        elif choice == '2': view_trades()
        elif choice == '3': delete_trade()
        elif choice == '4': summary()
        elif choice == '5': print("\n  👋 Goodbye!\n"); break
        else:               print("  ❌ Invalid option.")

if __name__ == "__main__":
    main()
