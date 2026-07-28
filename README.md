# rbPfPeek

<img alt="rbPfPeek" src="https://github.com/user-attachments/assets/cd7514f4-560b-49a6-bc2c-75651200fd0e" />

Please read the `#comments` in the topmost part of the script to find detailed and updated information on configuration.

# Features
- Graphical display of your portfolio performance today, and since a custom date.
- The portfolio assets can be read from a CSV exported from Directa (italian broker), or manually inputed in the `$manualAssets` splatted variable, with quantity and average price.
- Overall metrics: Profit/Loss, Money-weighted Return Rate, Day change, Portfolio cost and value.
  - If you provide the Directa CSV with all the portfolio transactions, then you also get info on expenses, trades, positions, and the total deposited and total withdrawn funds up to date.
- For each asset you own, you get: type (equity, bonds, commodity, money market), ticker, full name, qty, average cost, current price, current value and portfolio share (useful for periodical re-balancing), overall and daily profit/loss.
- The output partially adapts to the console window size. Tall is better; fullscreen is even better. Try also with different font sizes (usually <kbd>Ctrl</kbd>+mousewheel works).
- Current assets market data is retrieved from the web.
- Runs on any Powershell version (tested in Windows Powershell v5.1 and Powershell v7.6.4).

# Why?
This exists because I love [Ticker](https://github.com/achannarasappa/ticker), but i wanted some more information on screen.

And also because I wanted to verify if all the hype on "vibe coding" meant something...  
Oh well actually it did!  
It all started playfully: I dropped a screenshot of Ticker in Claude, asking to build a Powershell script which could give me the same output, "let's see if something happens"...  
2 minutes later I had the same output I have in Ticker!  
So I decided to add stuff I missed, to get the same information i like to have in Yahoo Finance, but in my beloved terminal:  
- the graphs  
- the CSV parsing feature  
- cosmetic changes done manually, because I am a f'ing perfectionist sometimes

... and after 38 iterations/releases done in small free time spans, I decided to release it here.

# Disclaimers
### 1. This was built for my own needs.
I invest in Italy, buying only ETFs for the long run.  
So I picked the data I needed to show on screen, and ignored some other data.  
It could fail with crypto, stocks, bonds, and other asset classes which are not ETFs.

### 2. Vibe coding is not fail-proof
This was 98% built by the free version of these AI models: Claude, Gemini, Deepseek, and a bit of Qwen run locally.  
So please expect errors.  
Feel free to open a new Issue if you can.
I am not a developer, but I am willing to have my published stuff well working.  
Thank you!
