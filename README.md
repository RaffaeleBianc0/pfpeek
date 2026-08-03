# pfpeek

<img alt="pfpeek" src="https://github.com/user-attachments/assets/cd7514f4-560b-49a6-bc2c-75651200fd0e" />

Please read the `#comments` in the topmost part of the script to find detailed and updated information.

# Features
- Graphical display of your portfolio performance today, and since a custom date.
- The portfolio assets can be read from a CSV or manually inputed in the `$manualAssets` splatted variable, each with its quantity and average price.
- Overall metrics: Profit/Loss, Day change, Portfolio cost and value.
- If you put a CSV file in the same folder as the script, containing all the portfolio transactions in the same format as Directa's "Movimenti" CSV (a sample file is included), then you also get: Money-Weighted Return Rate, expenses, # of trades, positions, and the total deposited / total withdrawn funds up to date.
- For each asset in the portfolio, you get: type (equity / bonds / commodity / money market), ticker, full name, qty, average cost, current price, current value and portfolio share (useful for periodical re-balancing), overall and daily profit/loss.
- The output adapts to the console window size. Tall is better, to get all the output without scrolling; fullscreen is even better. Try also with different font sizes (usually <kbd>Ctrl</kbd>+mousewheel works).
- Current assets market data is retrieved from the web.
- Runs on Windows Powershell 5.1 and Powershell 7.

# Why?
This exists because I really love many TUI apps, including [Ticker](https://github.com/achannarasappa/ticker), but i wanted some more information on screen.

And also because I wanted to verify if all the hype on "vibe coding" meant something...  
Oh well actually it did!  
It all started playfully: I dropped a screenshot of Ticker in Claude, asking to build a Powershell script which could give me the same output, "let's see if something happens"...  
2 minutes later I had the same output I have in Ticker!  
So I decided to add small enhancements that I missed from Ticker, to get the same information i like to see in my Yahoo Finance dashboard, but in my beloved terminal:  
- the graphs  
- the CSV parsing feature  
- cosmetic changes done manually, because I am a f'ing perfectionist sometimes

... and after 38 iterations/releases (this is why I'm starting with v1.38 here), all done in small free time spans, I feel confident enough to release it here.  

# Suggestions are welcome
Feel free to open an Issue here to share some ideas on any enhancement you may find useful.

# Disclaimers
### 1. This was built for my own needs
I invest in Italy, buying only ETFs for the long run.  
So I picked the data I wanted to see on my screen, and ignored some other data.  
It could fail with crypto, stocks, bonds, and other asset classes which are not ETFs... but I didn't test this. Give it a try!

### 2. Vibe coding is not fail-proof
This was 95% built by the free version of Claude, Gemini, Deepseek, and a bit of Qwen run locally.  
So there might be some bugs here and there: feel free to open a new Issue if you spot any.  
I am not an experienced developer, I only build small scripts to fit my needs, but I am willing to have my published stuff working good and nicely polished too.
