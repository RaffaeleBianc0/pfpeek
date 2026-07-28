# rbPfPeek
Windows Powershell script to show portfolio performance in the console.

<img alt="rbPfPeek" src="https://github.com/user-attachments/assets/cd7514f4-560b-49a6-bc2c-75651200fd0e" />


The portfolio assets can be read from a CSV exported from Directa (my current broker), or manually inputed in the `$manualAssets` splatted variable.

Please read the comments in the topmost part of the script to find detailed and updated information on configuration.

# Features
- Compatible with the CSV format of your portfolio transactions as exported by Directa (my current italian broker).
  - Alternatively, you can specify your own tickers with quantity and average price.
- The UI can adapt to the console window size.
- Gets the current market data from the web (Yahoo Finance).
- Runs on any Powershell version (tested in Windows Powershell v5.1 and Powershell v7.6.4). 

# Why?
This exists because I love [Ticker](https://github.com/achannarasappa/ticker), but i wanted some more information on screen.

And also because I wanted to verify if all the hype on "vibe coding" meant something... Well actually it did: i dropped a screenshot of Ticker in Claude, asking to build a Powershell script which could retrieve the same data from the web, then added the graphs, then added the CSV parsing feature, then some cosmetic changes done manually... and after 38 iterations/releases here it is.

# Disclaimer
This was 98% built by AI models: Claude, Gemini, Deepseek, and a bit of Qwen run locally.

So please expect errors.

Please report them if you can. I am not a developer, but I am willing to have my published stuff well working. 
Thank you!
