# pfpeek

<img alt="pfpeek" src="https://github.com/user-attachments/assets/983393df-78d7-4da8-bf93-f706ebf5b67c" />

> [!NOTE]
> 📌 **Please read the `#comments` in the topmost part of the script to find detailed and updated information.**

# ⚡ Features
- 📊 **Graphical display** of your portfolio performance, both today and since you began investing.
- 📂 **Flexible input:** Assets can be read automatically from a CSV file or defined via the `$manualAssets` variable, complete with quantity and average price.
- 📈 **Overall metrics:** Profit/Loss, Day change, Portfolio cost, and Current value.
- 📋 **Directa Integration:** Place a "Movimenti" CSV export file from Directa (a sample file is included) in the script directory to unlock:
  - 💰 *Money-Weighted Return Rate (MWRR)* & *Time-Weighted Rate of Return (TWRR)*
  - 💸 Total expenses & number of trades
  - 📊 Open and closed positions summary
  - 🔄 Total deposited and withdrawn funds to date
  - 💵 Realized gains from closed and partial positions
- 🔍 **Asset breakdown:** For each portfolio asset, view its type (*Equity, Bonds, Commodity, Money Market*), ticker, full name, quantity, average cost, current price, current value, weight %, and individual P/L (overall & daily).
- 🖥️ **Adaptive console layout:** Dynamic text wrapping designed for terminal windows. *Tall is better to fit all output without scrolling; fullscreen is ideal.* Try adjusting font sizes (<kbd>Ctrl</kbd> + mousewheel).
- 🌐 **Real-time data:** Fetches live market quotes directly from the web.
- 💻 **Cross-platform PowerShell support:** Runs smoothly on modern Windows 10/11 environments, fully compatible with both **PowerShell 5.1** and **PowerShell 7+**.



# ❓ Why?
This project exists because I love [Ticker](https://github.com/achannarasappa/ticker) and wanted even more detailed data right on my screen.

It was also a fun test to see if all the hype around **"vibe coding"** held up...  
*Spoiler: It actually did!* 🚀

It all started as an experiment: dropping a Ticker screenshot into Claude and asking it to write a PowerShell script to replicate the view (*"let's see what happens..."*).  
⏱️ **2 minutes later**, I had a working clone!

From there, I decided to build in all the missing features I enjoy in Yahoo Finance web interface, bringing them directly to my beloved terminal:
- 📉 Performance charts
- 📄 Directa CSV parsing and historical metrics
- 🎨 Fine-tuned cosmetic tweaks *(because I'm a f'ing perfectionist sometimes)*

...And **38 iterations later** (hence starting with v1.38), built during small pockets of free time, it’s ready to share!



# 💡 Suggestions Welcome
Feel free to open an **Issue** to share ideas for new features, improvements, or to report any bugs you spot.



# ⚠️ Disclaimers

### 1. Built for my own workflow
I invest in Italy, holding long-term ETFs only.  
The metrics displayed reflect the specific data points I care about daily. It *might* work with crypto, individual stocks, or bonds, but it hasn't been tested for those asset classes yet — give it a try!

### 2. Vibe coding is not bulletproof
About 95% of this codebase was generated using free tiers of *Claude, Gemini, DeepSeek*, and a locally hosted *Qwen* model.  
While I am not a professional developer, I write small scripts for my own tools and am committed to keeping this project functional and well-polished.

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/I3I5MBHBZ)
